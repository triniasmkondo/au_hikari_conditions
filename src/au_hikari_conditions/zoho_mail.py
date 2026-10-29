from __future__ import annotations

import json
import os
import ctypes
import binascii
import re
from ctypes import wintypes
from pathlib import Path
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import quote
from typing import Callable, Protocol
from urllib.error import HTTPError
from urllib.parse import urlencode
from urllib.request import HTTPRedirectHandler, Request, build_opener


_REGION_ENDPOINTS = {
    "us": ("https://accounts.zoho.com", "https://mail.zoho.com"),
    "eu": ("https://accounts.zoho.eu", "https://mail.zoho.eu"),
    "in": ("https://accounts.zoho.in", "https://mail.zoho.in"),
    "au": ("https://accounts.zoho.com.au", "https://mail.zoho.com.au"),
    "cn": ("https://accounts.zoho.com.cn", "https://mail.zoho.com.cn"),
    "jp": ("https://accounts.zoho.jp", "https://mail.zoho.jp"),
    "ae": ("https://accounts.zoho.ae", "https://mail.zoho.ae"),
    "ca": ("https://accounts.zohocloud.ca", "https://mail.zohocloud.ca"),
    "sa": ("https://accounts.zoho.sa", "https://mail.zoho.sa"),
    "uk": ("https://accounts.zoho.uk", "https://mail.zoho.uk"),
}
_ACCESS_TOKEN_REFRESH_SKEW_SECONDS = 60
_DEFAULT_TIMEOUT_SECONDS = 15
_REPOSITORY_ROOT = Path(__file__).resolve().parents[2]


class SecretProvider(Protocol):
    def get_secret(self, name: str) -> str | None: ...


class EnvironmentSecretProvider:
    """Reads credentials from environment variables without retaining them."""

    def get_secret(self, name: str) -> str | None:
        return os.environ.get(name)


class DpapiSecretProvider:
    """Reads PowerShell CurrentUser DPAPI refresh tokens and a restricted client JSON."""

    _CLIENT_JSON_KEYS = {
        "ZOHO_CLIENT_ID": "client_id",
        "ZOHO_CLIENT_SECRET": "client_secret",
    }

    def __init__(
        self,
        *,
        token_file: Path | None = None,
        self_client_file: Path | None = None,
        fallback: SecretProvider | None = None,
    ) -> None:
        local_app_data = os.environ.get("LOCALAPPDATA")
        configured_token_path = os.environ.get("ZOHO_REFRESH_TOKEN_DPAPI_PATH")
        if token_file is None and configured_token_path:
            expanded_token_path = Path(
                os.path.expandvars(configured_token_path)
            ).expanduser()
            if not expanded_token_path.is_absolute():
                raise ZohoConfigurationError(
                    "ZOHO_REFRESH_TOKEN_DPAPI_PATH must be an absolute path."
                )
            try:
                expanded_token_path.resolve().relative_to(_REPOSITORY_ROOT)
            except ValueError:
                pass
            else:
                raise ZohoConfigurationError(
                    "ZOHO_REFRESH_TOKEN_DPAPI_PATH must point outside the Git repository."
                )
            token_file = expanded_token_path
        if token_file is None or self_client_file is None:
            if not local_app_data:
                raise ZohoConfigurationError(
                    "LOCALAPPDATA is required for the DPAPI secret provider."
                )
            root = Path(local_app_data) / "ZohoMailApiTest"
            token_file = token_file or root / "refresh_token.dpapi"
            self_client_file = self_client_file or root / "self_client.json"
        self._token_file = token_file
        self._self_client_file = self_client_file
        self._fallback = (
            fallback if fallback is not None else EnvironmentSecretProvider()
        )

    def get_secret(self, name: str) -> str | None:
        if name == "ZOHO_REFRESH_TOKEN":
            protected_text = _read_secret_text(self._token_file)
            return _unprotect_powershell_dpapi(protected_text)
        json_key = self._CLIENT_JSON_KEYS.get(name)
        if json_key is not None:
            return self._read_client_json_value(json_key)
        return self._fallback.get_secret(name)

    def _read_client_json_value(self, key: str) -> str:
        content = _read_secret_text(self._self_client_file)
        try:
            values = json.loads(content)
        except (json.JSONDecodeError, UnicodeDecodeError):
            raise ZohoConfigurationError(
                "Self Client configuration file is not valid JSON."
            ) from None
        if not isinstance(values, dict):
            raise ZohoConfigurationError(
                "Self Client configuration must be a JSON object."
            )
        value = values.get(key)
        if not isinstance(value, str) or not value.strip():
            raise ZohoConfigurationError(
                f"Self Client configuration is missing required field: {key}."
            )
        return value


class ZohoMailError(RuntimeError):
    """Base class for sanitized Zoho Mail connection errors."""


class ZohoConfigurationError(ZohoMailError):
    pass


class ZohoAuthenticationError(ZohoMailError):
    pass


class ZohoAuthorizationError(ZohoMailError):
    pass


class ZohoRegionError(ZohoMailError):
    pass


class ZohoRateLimitError(ZohoMailError):
    pass


class ZohoServiceError(ZohoMailError):
    pass


class ZohoConnectionError(ZohoMailError):
    pass


class ZohoProtocolError(ZohoMailError):
    pass


def _read_secret_text(path: Path) -> str:
    try:
        raw = path.read_bytes()
    except OSError:
        raise ZohoConfigurationError(
            "A required Zoho secret file is missing or unreadable."
        ) from None
    try:
        if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
            text = raw.decode("utf-16")
        else:
            text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        raise ZohoConfigurationError(
            "A required Zoho secret file has an unsupported text encoding."
        ) from None
    return text.strip()


class _DataBlob(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_ubyte))]


def _unprotect_powershell_dpapi(protected_text: str) -> str:
    if os.name != "nt":
        raise ZohoConfigurationError(
            "PowerShell DPAPI secrets can only be read on Windows."
        )
    normalized = protected_text.strip()
    if (
        len(normalized) < 64
        or len(normalized) % 2
        or not normalized.lower().startswith(
            "01000000d08c9ddf0115d1118c7a00c04fc297eb"
        )
    ):
        raise ZohoConfigurationError(
            "The protected token file is not in the supported PowerShell DPAPI format."
        )
    try:
        protected_bytes = binascii.unhexlify(normalized)
    except (binascii.Error, ValueError):
        raise ZohoConfigurationError(
            "The protected token file is not in the supported PowerShell DPAPI format."
        ) from None

    input_buffer = (ctypes.c_ubyte * len(protected_bytes)).from_buffer_copy(
        protected_bytes
    )
    input_blob = _DataBlob(len(protected_bytes), input_buffer)
    output_blob = _DataBlob()
    try:
        crypt32 = ctypes.WinDLL("crypt32", use_last_error=True)
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    except OSError:
        raise ZohoConfigurationError(
            "Windows DPAPI is unavailable in this process."
        ) from None
    unprotect = crypt32.CryptUnprotectData
    unprotect.argtypes = (
        ctypes.POINTER(_DataBlob),
        ctypes.POINTER(wintypes.LPWSTR),
        ctypes.POINTER(_DataBlob),
        ctypes.c_void_p,
        ctypes.c_void_p,
        wintypes.DWORD,
        ctypes.POINTER(_DataBlob),
    )
    unprotect.restype = wintypes.BOOL
    kernel32.LocalFree.argtypes = (ctypes.c_void_p,)
    kernel32.LocalFree.restype = ctypes.c_void_p

    if not unprotect(
        ctypes.byref(input_blob),
        None,
        None,
        None,
        None,
        0,
        ctypes.byref(output_blob),
    ):
        raise ZohoConfigurationError(
            "DPAPI could not decrypt the token for the current Windows user."
        )
    try:
        if not output_blob.pbData or output_blob.cbData < 2 or output_blob.cbData % 2:
            raise ZohoConfigurationError(
                "DPAPI returned an invalid protected token payload."
            )
        plaintext_buffer = (ctypes.c_ubyte * output_blob.cbData).from_address(
            ctypes.addressof(output_blob.pbData.contents)
        )
        try:
            plaintext = bytes(plaintext_buffer).decode("utf-16-le")
        except UnicodeDecodeError:
            raise ZohoConfigurationError(
                "DPAPI returned an invalid protected token payload."
            ) from None
        if plaintext.endswith("\x00"):
            plaintext = plaintext[:-1]
        if not plaintext or "\x00" in plaintext:
            raise ZohoConfigurationError(
                "DPAPI returned an invalid protected token payload."
            )
        return plaintext
    finally:
        if output_blob.pbData:
            ctypes.memset(output_blob.pbData, 0, output_blob.cbData)
            kernel32.LocalFree(output_blob.pbData)


@dataclass(frozen=True)
class ZohoMailConfig:
    client_id: str = field(repr=False)
    client_secret: str = field(repr=False)
    refresh_token: str = field(repr=False)
    account_id: str = field(repr=False)
    region: str

    def __post_init__(self) -> None:
        if not all(
            value.strip()
            for value in (
                self.client_id,
                self.client_secret,
                self.refresh_token,
                self.account_id,
            )
        ):
            raise ZohoConfigurationError(
                "Zoho client ID, client secret, refresh token, and account ID are required."
            )
        normalized_region = self.region.strip().lower()
        if normalized_region not in _REGION_ENDPOINTS:
            supported = ", ".join(sorted(_REGION_ENDPOINTS))
            raise ZohoConfigurationError(
                f"Unsupported Zoho region. Select one of: {supported}."
            )
        object.__setattr__(self, "region", normalized_region)

    @classmethod
    def from_secret_provider(
        cls, provider: SecretProvider, *, prefix: str = "ZOHO_"
    ) -> ZohoMailConfig:
        names = {
            "client_id": f"{prefix}CLIENT_ID",
            "client_secret": f"{prefix}CLIENT_SECRET",
            "refresh_token": f"{prefix}REFRESH_TOKEN",
            "account_id": f"{prefix}ACCOUNT_ID",
            "region": f"{prefix}REGION",
        }
        values = {key: provider.get_secret(name) for key, name in names.items()}
        missing = [names[key] for key, value in values.items() if not value]
        if missing:
            raise ZohoConfigurationError(
                "Missing required Zoho configuration: " + ", ".join(missing)
            )
        return cls(
            client_id=values["client_id"] or "",
            client_secret=values["client_secret"] or "",
            refresh_token=values["refresh_token"] or "",
            account_id=values["account_id"] or "",
            region=values["region"] or "",
        )

    @property
    def accounts_base_url(self) -> str:
        return _REGION_ENDPOINTS[self.region][0]

    @property
    def mail_base_url(self) -> str:
        return _REGION_ENDPOINTS[self.region][1]


@dataclass(frozen=True)
class ZohoAccount:
    account_id: str = field(repr=False)
    account_name: str = field(repr=False)
    primary_email: str = field(repr=False)


class HttpTransport(Protocol):
    def request(
        self,
        *,
        url: str,
        method: str,
        headers: dict[str, str],
        body: bytes | None,
        timeout: float,
    ) -> tuple[int, bytes]: ...


class UrllibTransport:
    def request(
        self,
        *,
        url: str,
        method: str,
        headers: dict[str, str],
        body: bytes | None,
        timeout: float,
    ) -> tuple[int, bytes]:
        request = Request(url, data=body, headers=headers, method=method)
        try:
            opener = build_opener(_NoRedirectHandler())
            with opener.open(request, timeout=timeout) as response:
                return response.status, response.read()
        except HTTPError as error:
            return error.code, error.read()


class _NoRedirectHandler(HTTPRedirectHandler):
    def redirect_request(
        self,
        _request: Request,
        _response: object,
        _code: int,
        _message: str,
        _headers: object,
        _new_url: str,
    ) -> None:
        return None


class ZohoMailClient:
    """Zoho Mail account reader with draft creation only; sending is not implemented."""

    def __init__(
        self,
        config: ZohoMailConfig,
        *,
        transport: HttpTransport | None = None,
        timeout: float = _DEFAULT_TIMEOUT_SECONDS,
        clock: Callable[[], float] = time.monotonic,
    ) -> None:
        if timeout <= 0:
            raise ZohoConfigurationError("HTTP timeout must be greater than zero.")
        self._config = config
        self._transport = transport or UrllibTransport()
        self._timeout = timeout
        self._clock = clock
        self._token: str | None = None
        self._token_expires_at = 0.0
        self._token_lock = threading.Lock()

    def get_accounts(self) -> tuple[ZohoAccount, ...]:
        token = self._get_access_token()
        status, payload = self._request_accounts(token)
        if status == 401:
            self._discard_access_token(token)
            token = self._get_access_token()
            status, payload = self._request_accounts(token)
        self._raise_for_status(status, is_token_request=False)
        response = self._decode_json(payload, "Zoho Mail account response")
        return self._parse_accounts(response)

    def get_configured_account(self) -> ZohoAccount:
        for account in self.get_accounts():
            if account.account_id == self._config.account_id:
                return account
        raise ZohoConfigurationError(
            "Configured Zoho account ID was not present in the accessible accounts."
        )

    def upload_draft_attachment(
        self, *, file_name: str, content: bytes, content_type: str
    ) -> dict[str, str]:
        if not file_name.strip() or Path(file_name).name != file_name:
            raise ZohoConfigurationError("Attachment file name is invalid.")
        if not re.fullmatch(
            r"[A-Za-z0-9!#$&^_.+-]+/[A-Za-z0-9!#$&^_.+-]+", content_type
        ):
            raise ZohoConfigurationError("Attachment content type is invalid.")
        if not content:
            raise ZohoConfigurationError("Attachment content is empty.")
        account_id = quote(self._config.account_id, safe="")
        query = urlencode({"fileName": file_name, "isInline": "false"})
        status, payload = self._request(
            url=(
                f"{self._config.mail_base_url}/api/accounts/{account_id}"
                f"/messages/attachments?{query}"
            ),
            method="POST",
            headers={
                "Authorization": f"Zoho-oauthtoken {self._get_access_token()}",
                "Content-Type": content_type,
                "Accept": "application/json",
            },
            body=content,
        )
        self._raise_for_status(status, is_token_request=False)
        response = self._decode_json(payload, "Zoho Mail attachment response")
        return self._parse_uploaded_attachment(response)

    def save_draft(self, payload: dict[str, object]) -> dict[str, object]:
        if payload.get("mode") != "draft":
            raise ZohoConfigurationError(
                'Zoho Mail draft request must set mode="draft".'
            )
        required_fields = ("fromAddress", "toAddress", "subject", "content")
        if any(
            not isinstance(payload.get(field_name), str)
            or not str(payload[field_name]).strip()
            for field_name in required_fields
        ):
            raise ZohoConfigurationError("Zoho Mail draft request is incomplete.")
        account_id = quote(self._config.account_id, safe="")
        status, response_bytes = self._request(
            url=f"{self._config.mail_base_url}/api/accounts/{account_id}/messages",
            method="POST",
            headers={
                "Authorization": f"Zoho-oauthtoken {self._get_access_token()}",
                "Content-Type": "application/json",
                "Accept": "application/json",
            },
            body=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        )
        self._raise_for_status(status, is_token_request=False)
        response = self._decode_json(response_bytes, "Zoho Mail draft response")
        api_status = response.get("status")
        if not isinstance(api_status, dict) or str(api_status.get("code")) not in {
            "200",
            "201",
        }:
            raise ZohoServiceError("Zoho Mail did not confirm draft creation.")
        return response

    def _get_access_token(self) -> str:
        with self._token_lock:
            if (
                self._token
                and self._token_expires_at
                > self._clock() + _ACCESS_TOKEN_REFRESH_SKEW_SECONDS
            ):
                return self._token
            self._refresh_access_token()
            if not self._token:
                raise ZohoProtocolError("Zoho did not return an access token.")
            return self._token

    def _refresh_access_token(self) -> None:
        body = urlencode(
            {
                "client_id": self._config.client_id,
                "client_secret": self._config.client_secret,
                "refresh_token": self._config.refresh_token,
                "grant_type": "refresh_token",
            }
        ).encode("ascii")
        status, payload = self._request(
            url=f"{self._config.accounts_base_url}/oauth/v2/token",
            method="POST",
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            body=body,
        )
        self._raise_for_status(status, is_token_request=True)
        response = self._decode_json(payload, "Zoho OAuth token response")
        if response.get("error"):
            raise ZohoAuthenticationError(
                "Zoho rejected the OAuth credentials or refresh token."
            )
        access_token = response.get("access_token")
        expires_in = response.get("expires_in")
        if not isinstance(access_token, str) or not access_token.strip():
            raise ZohoProtocolError("Zoho did not return a valid access token.")
        try:
            lifetime = float(expires_in)
        except (TypeError, ValueError):
            raise ZohoProtocolError("Zoho returned an invalid token lifetime.") from None
        if lifetime <= 0:
            raise ZohoProtocolError("Zoho returned an invalid token lifetime.")
        self._token = access_token
        self._token_expires_at = self._clock() + lifetime

    def _request_accounts(self, token: str) -> tuple[int, bytes]:
        return self._request(
            url=f"{self._config.mail_base_url}/api/accounts",
            method="GET",
            headers={"Authorization": f"Zoho-oauthtoken {token}"},
            body=None,
        )

    def _request(
        self,
        *,
        url: str,
        method: str,
        headers: dict[str, str],
        body: bytes | None,
    ) -> tuple[int, bytes]:
        try:
            return self._transport.request(
                url=url,
                method=method,
                headers=headers,
                body=body,
                timeout=self._timeout,
            )
        except OSError:
            raise ZohoConnectionError(
                "Could not connect to the configured Zoho service."
            ) from None

    def _discard_access_token(self, rejected_token: str) -> None:
        with self._token_lock:
            if self._token == rejected_token:
                self._token = None
                self._token_expires_at = 0.0

    @staticmethod
    def _decode_json(payload: bytes, description: str) -> dict[str, object]:
        try:
            response = json.loads(payload)
        except (UnicodeDecodeError, json.JSONDecodeError):
            raise ZohoProtocolError(f"{description} was not valid JSON.") from None
        if not isinstance(response, dict):
            raise ZohoProtocolError(f"{description} had an unexpected format.")
        return response

    @staticmethod
    def _raise_for_status(status: int, *, is_token_request: bool) -> None:
        if 200 <= status < 300:
            return
        if status in {400, 401}:
            if is_token_request:
                raise ZohoAuthenticationError(
                    "Zoho rejected the OAuth credentials or refresh token."
                )
            raise ZohoAuthenticationError(
                "Zoho Mail rejected authentication after token refresh."
            )
        if status == 403:
            raise ZohoAuthorizationError(
                "Zoho denied access; verify the account and OAuth scope."
            )
        if status == 404:
            raise ZohoRegionError(
                "Zoho endpoint was not found; verify the configured data center."
            )
        if status == 429:
            raise ZohoRateLimitError("Zoho rate limit reached; retry later.")
        if status >= 500:
            raise ZohoServiceError("Zoho reported a temporary service error.")
        raise ZohoServiceError(f"Zoho returned an unexpected HTTP status ({status}).")

    @staticmethod
    def _parse_accounts(response: dict[str, object]) -> tuple[ZohoAccount, ...]:
        data = response.get("data")
        if isinstance(data, dict):
            accounts = data.get("accounts")
            if accounts is None and _account_id(data):
                accounts = [data]
        else:
            accounts = data
        if not isinstance(accounts, list):
            raise ZohoProtocolError("Zoho account response contained no account list.")

        parsed: list[ZohoAccount] = []
        for item in accounts:
            if not isinstance(item, dict):
                raise ZohoProtocolError("Zoho returned an invalid account record.")
            account_id = _account_id(item)
            if not account_id:
                raise ZohoProtocolError("Zoho account record had no account ID.")
            parsed.append(
                ZohoAccount(
                    account_id=account_id,
                    account_name=_first_text(item, "accountName", "account_name"),
                    primary_email=_first_text(
                        item,
                        "primaryEmailAddress",
                        "primary_email_address",
                        "emailAddress",
                    ),
                )
            )
        return tuple(parsed)

    @staticmethod
    def _parse_uploaded_attachment(response: dict[str, object]) -> dict[str, str]:
        data = response.get("data")
        if not isinstance(data, dict):
            raise ZohoProtocolError("Zoho attachment response was incomplete.")
        fields = {
            "storeName": data.get("storeName"),
            "attachmentName": data.get("attachmentName"),
            "attachmentPath": data.get("attachmentPath"),
        }
        if any(not isinstance(value, str) or not value.strip() for value in fields.values()):
            raise ZohoProtocolError("Zoho attachment response was incomplete.")
        return {key: value for key, value in fields.items() if isinstance(value, str)}


def _account_id(account: dict[str, object]) -> str:
    return _first_text(account, "accountId", "account_id")


def _first_text(account: dict[str, object], *keys: str) -> str:
    for key in keys:
        value = account.get(key)
        if isinstance(value, (str, int)) and str(value).strip():
            return str(value).strip()
    return ""
