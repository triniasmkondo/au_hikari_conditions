from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from collections import deque
from pathlib import Path
from unittest.mock import patch
from urllib.error import URLError
from urllib.parse import parse_qs
from urllib.request import Request

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from au_hikari_conditions.zoho_mail import (
    EnvironmentSecretProvider,
    _NoRedirectHandler,
    DpapiSecretProvider,
    ZohoAuthenticationError,
    ZohoAuthorizationError,
    ZohoConfigurationError,
    ZohoConnectionError,
    ZohoMailClient,
    ZohoMailConfig,
    ZohoProtocolError,
    ZohoRateLimitError,
    ZohoRegionError,
    ZohoServiceError,
)


class MappingSecretProvider:
    def __init__(self, values: dict[str, str]) -> None:
        self.values = values

    def get_secret(self, name: str) -> str | None:
        return self.values.get(name)


class FakeTransport:
    def __init__(self, responses: list[tuple[int, bytes]] | None = None) -> None:
        self.responses = deque(responses or [])
        self.requests: list[dict[str, object]] = []
        self.error: Exception | None = None

    def request(
        self,
        *,
        url: str,
        method: str,
        headers: dict[str, str],
        body: bytes | None,
        timeout: float,
    ) -> tuple[int, bytes]:
        self.requests.append(
            {
                "url": url,
                "method": method,
                "headers": headers,
                "body": body,
                "timeout": timeout,
            }
        )
        if self.error is not None:
            raise self.error
        if not self.responses:
            raise AssertionError("Unexpected HTTP request in test.")
        return self.responses.popleft()


class MutableClock:
    def __init__(self, value: float = 100.0) -> None:
        self.value = value

    def __call__(self) -> float:
        return self.value


def _json_response(status: int, value: object) -> tuple[int, bytes]:
    return status, json.dumps(value).encode("utf-8")


def _config(region: str = "us") -> ZohoMailConfig:
    return ZohoMailConfig(
        client_id="synthetic-client-id",
        client_secret="synthetic-client-secret",
        refresh_token="synthetic-refresh-token",
        account_id="synthetic-account-id",
        region=region,
    )


def _token(access_token: str, expires_in: int = 3600) -> tuple[int, bytes]:
    return _json_response(
        200, {"access_token": access_token, "expires_in": expires_in}
    )


def _accounts(email: str = "anonymous@example.invalid") -> tuple[int, bytes]:
    return _json_response(
        200,
        {
            "data": [
                {
                    "accountId": "synthetic-account-id",
                    "accountName": "Anonymous Mailbox",
                    "primaryEmailAddress": email,
                }
            ]
        },
    )


def _write_powershell_dpapi_fixture(destination: Path) -> bytes:
    if os.name != "nt":
        raise unittest.SkipTest("PowerShell DPAPI tests require Windows.")
    powershell = shutil.which("powershell.exe")
    if powershell is None:
        raise unittest.SkipTest("Windows PowerShell is unavailable.")
    script = (
        "$ErrorActionPreference = 'Stop'\n"
        "$secure = ConvertTo-SecureString 'synthetic-dpapi-refresh-token' "
        "-AsPlainText -Force\n"
        "ConvertFrom-SecureString -SecureString $secure | "
        "Set-Content -LiteralPath $env:ZOHO_DPAPI_TEST_FILE\n"
    )
    environment = os.environ.copy()
    environment["ZOHO_DPAPI_TEST_FILE"] = str(destination)
    completed = subprocess.run(
        [
            powershell,
            "-NoLogo",
            "-NoProfile",
            "-NonInteractive",
            "-Command",
            "-",
        ],
        input=script,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=environment,
        timeout=30,
        check=False,
    )
    if completed.returncode != 0:
        raise AssertionError("PowerShell could not create the synthetic DPAPI fixture.")
    protected_file = destination.read_bytes()
    if protected_file.startswith((b"\xff\xfe", b"\xfe\xff", b"\xef\xbb\xbf")):
        raise AssertionError("PowerShell Set-Content unexpectedly wrote a BOM.")
    if not protected_file.endswith(b"\r\n"):
        raise AssertionError("PowerShell Set-Content did not write a CRLF newline.")
    try:
        protected = protected_file.decode("ascii").strip()
    except UnicodeDecodeError:
        raise AssertionError("PowerShell Set-Content output was not ASCII-compatible.") from None
    if not protected.startswith("01000000d08c9ddf"):
        raise AssertionError("PowerShell returned an unexpected DPAPI fixture format.")
    return protected_file


class ZohoMailClientTest(unittest.TestCase):
    def test_reads_credentials_from_provider_without_exposing_secret_repr(self) -> None:
        config = ZohoMailConfig.from_secret_provider(
            MappingSecretProvider(
                {
                    "ZOHO_CLIENT_ID": "synthetic-client-id",
                    "ZOHO_CLIENT_SECRET": "synthetic-client-secret",
                    "ZOHO_REFRESH_TOKEN": "synthetic-refresh-token",
                    "ZOHO_ACCOUNT_ID": "synthetic-account-id",
                    "ZOHO_REGION": "JP",
                }
            )
        )

        self.assertEqual(config.region, "jp")
        self.assertEqual(config.accounts_base_url, "https://accounts.zoho.jp")
        self.assertEqual(config.mail_base_url, "https://mail.zoho.jp")
        self.assertNotIn("synthetic-client-secret", repr(config))
        self.assertNotIn("synthetic-refresh-token", repr(config))
        self.assertNotIn("synthetic-account-id", repr(config))

    def test_environment_provider_reads_values_at_call_time(self) -> None:
        provider = EnvironmentSecretProvider()
        with patch.dict("os.environ", {"ZOHO_REGION": "eu"}, clear=False):
            self.assertEqual(provider.get_secret("ZOHO_REGION"), "eu")

    def test_dpapi_provider_reads_synthetic_token_and_self_client_configuration(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            token_file = root / "refresh_token.dpapi"
            client_file = root / "self_client.json"
            protected_file = _write_powershell_dpapi_fixture(token_file)
            self.assertEqual(token_file.read_bytes(), protected_file)
            client_file.write_text(
                json.dumps(
                    {
                        "client_id": "synthetic-client-id",
                        "client_secret": "synthetic-client-secret",
                    }
                ),
                encoding="utf-16",
            )
            provider = DpapiSecretProvider(
                token_file=token_file,
                self_client_file=client_file,
                fallback=MappingSecretProvider(
                    {
                        "ZOHO_ACCOUNT_ID": "synthetic-account-id",
                        "ZOHO_REGION": "us",
                    }
                ),
            )

            config = ZohoMailConfig.from_secret_provider(provider)

            self.assertEqual(config.client_id, "synthetic-client-id")
            self.assertEqual(config.client_secret, "synthetic-client-secret")
            self.assertEqual(config.refresh_token, "synthetic-dpapi-refresh-token")
            self.assertEqual(config.account_id, "synthetic-account-id")
            self.assertEqual(config.region, "us")
            self.assertNotIn("synthetic-client-secret", repr(config))
            self.assertNotIn("synthetic-dpapi-refresh-token", repr(config))

    def test_dpapi_provider_uses_default_local_app_data_paths_and_environment_fallback(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            local_app_data = Path(temporary_directory)
            secret_directory = local_app_data / "ZohoMailApiTest"
            secret_directory.mkdir()
            _write_powershell_dpapi_fixture(
                secret_directory / "refresh_token.dpapi"
            )
            (secret_directory / "self_client.json").write_text(
                '{"client_id":"synthetic-client-id",'
                '"client_secret":"synthetic-client-secret"}',
                encoding="utf-8",
            )
            with patch.dict(
                os.environ,
                {
                    "LOCALAPPDATA": str(local_app_data),
                    "ZOHO_ACCOUNT_ID": "synthetic-account-id",
                    "ZOHO_REGION": "us",
                },
                clear=False,
            ):
                config = ZohoMailConfig.from_secret_provider(DpapiSecretProvider())

            self.assertEqual(config.refresh_token, "synthetic-dpapi-refresh-token")
            self.assertEqual(config.account_id, "synthetic-account-id")
            self.assertEqual(config.region, "us")

    def test_dpapi_provider_can_select_external_v2_token_without_changing_legacy_file(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            legacy = root / "ZohoMailApiTest" / "refresh_token.dpapi"
            v2_token = root / "secure" / "refresh_token_v2.dpapi"
            legacy.parent.mkdir()
            v2_token.parent.mkdir()
            legacy.write_text("legacy-protected-dummy", encoding="ascii")
            v2_token.write_text("v2-protected-dummy", encoding="ascii")
            client_file = root / "self_client.json"
            client_file.write_text("{}", encoding="utf-8")
            with patch.dict(
                os.environ,
                {
                    "LOCALAPPDATA": str(root),
                    "ZOHO_REFRESH_TOKEN_DPAPI_PATH": str(v2_token),
                },
                clear=False,
            ):
                provider = DpapiSecretProvider(self_client_file=client_file)

            self.assertEqual(provider._token_file, v2_token)
            self.assertEqual(legacy.read_text(encoding="ascii"), "legacy-protected-dummy")
            self.assertEqual(v2_token.read_text(encoding="ascii"), "v2-protected-dummy")

    def test_dpapi_v2_path_must_be_absolute_and_outside_repository(self) -> None:
        with patch.dict(
            os.environ,
            {
                "LOCALAPPDATA": tempfile.gettempdir(),
                "ZOHO_REFRESH_TOKEN_DPAPI_PATH": "relative-token.dpapi",
            },
            clear=False,
        ):
            with self.assertRaises(ZohoConfigurationError):
                DpapiSecretProvider()

        with patch.dict(
            os.environ,
            {
                "LOCALAPPDATA": tempfile.gettempdir(),
                "ZOHO_REFRESH_TOKEN_DPAPI_PATH": str(
                    Path(__file__).resolve().parents[1]
                    / "synthetic-token.dpapi"
                ),
            },
            clear=False,
        ):
            with self.assertRaises(ZohoConfigurationError):
                DpapiSecretProvider()

    def test_dpapi_provider_fails_safely_for_missing_files(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            provider = DpapiSecretProvider(
                token_file=root / "missing.dpapi",
                self_client_file=root / "missing.json",
            )
            with self.assertRaises(ZohoConfigurationError) as raised:
                provider.get_secret("ZOHO_REFRESH_TOKEN")
            self.assertNotIn(str(root), str(raised.exception))

    def test_dpapi_provider_rejects_malformed_and_wrong_user_ciphertext(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            token_file = Path(temporary_directory) / "refresh_token.dpapi"
            client_file = Path(temporary_directory) / "self_client.json"
            client_file.write_text(
                '{"client_id":"synthetic-id","client_secret":"synthetic-secret"}',
                encoding="utf-8",
            )
            provider = DpapiSecretProvider(
                token_file=token_file,
                self_client_file=client_file,
            )
            token_file.write_text("not-a-dpapi-value", encoding="utf-8")
            with self.assertRaises(ZohoConfigurationError) as malformed:
                provider.get_secret("ZOHO_REFRESH_TOKEN")
            self.assertNotIn("not-a-dpapi-value", str(malformed.exception))

            _write_powershell_dpapi_fixture(token_file)
            protected = token_file.read_bytes().decode("ascii").strip()
            damaged = protected[:-2] + ("00" if protected[-2:] != "00" else "ff")
            token_file.write_text(damaged, encoding="utf-16")
            with self.assertRaises(ZohoConfigurationError) as decryption:
                provider.get_secret("ZOHO_REFRESH_TOKEN")
            self.assertIn("current Windows user", str(decryption.exception))
            self.assertNotIn("synthetic-dpapi-refresh-token", str(decryption.exception))

    def test_dpapi_provider_rejects_missing_or_invalid_client_json_fields(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            token_file = root / "token.dpapi"
            client_file = root / "client.json"
            token_file.write_text("unused", encoding="utf-8")
            provider = DpapiSecretProvider(
                token_file=token_file,
                self_client_file=client_file,
            )
            with self.assertRaises(ZohoConfigurationError):
                provider.get_secret("ZOHO_CLIENT_ID")

            client_file.write_text('{"client_id":12}', encoding="utf-8")
            with self.assertRaises(ZohoConfigurationError) as missing_field:
                provider.get_secret("ZOHO_CLIENT_SECRET")
            self.assertIn("client_secret", str(missing_field.exception))
            self.assertNotIn("12", str(missing_field.exception))

    def test_missing_secret_names_are_reported_without_values(self) -> None:
        with self.assertRaises(ZohoConfigurationError) as raised:
            ZohoMailConfig.from_secret_provider(
                MappingSecretProvider(
                    {
                        "ZOHO_CLIENT_ID": "synthetic-id",
                        "ZOHO_REGION": "us",
                    }
                )
            )

        self.assertIn("ZOHO_CLIENT_SECRET", str(raised.exception))
        self.assertIn("ZOHO_REFRESH_TOKEN", str(raised.exception))
        self.assertIn("ZOHO_ACCOUNT_ID", str(raised.exception))
        self.assertNotIn("synthetic-id", str(raised.exception))

    def test_unsupported_region_is_rejected_without_building_a_url(self) -> None:
        with self.assertRaises(ZohoConfigurationError):
            _config("https://attacker.example")

    def test_oauth_http_redirects_are_not_followed(self) -> None:
        redirect = _NoRedirectHandler().redirect_request(
            Request(
                "https://accounts.zoho.com/oauth/v2/token",
                data=b"client_secret=synthetic-secret",
                method="POST",
            ),
            object(),
            307,
            "Temporary Redirect",
            {},
            "https://attacker.example/collect",
        )

        self.assertIsNone(redirect)

    def test_refreshes_token_and_reads_accounts_with_zoho_mail_auth_header(self) -> None:
        transport = FakeTransport([_token("synthetic-access-token"), _accounts()])
        client = ZohoMailClient(_config(), transport=transport)

        accounts = client.get_accounts()

        self.assertEqual(len(accounts), 1)
        self.assertEqual(accounts[0].account_id, "synthetic-account-id")
        self.assertEqual(accounts[0].primary_email, "anonymous@example.invalid")
        self.assertNotIn("anonymous@example.invalid", repr(accounts[0]))
        token_request, account_request = transport.requests
        self.assertEqual(
            token_request["url"], "https://accounts.zoho.com/oauth/v2/token"
        )
        self.assertEqual(token_request["method"], "POST")
        self.assertEqual(
            parse_qs(token_request["body"].decode("ascii")),
            {
                "client_id": ["synthetic-client-id"],
                "client_secret": ["synthetic-client-secret"],
                "refresh_token": ["synthetic-refresh-token"],
                "grant_type": ["refresh_token"],
            },
        )
        self.assertEqual(account_request["url"], "https://mail.zoho.com/api/accounts")
        self.assertEqual(
            account_request["headers"]["Authorization"],
            "Zoho-oauthtoken synthetic-access-token",
        )

    def test_selects_configured_account_from_multiple_accounts(self) -> None:
        accounts_response = _json_response(
            200,
            {
                "data": [
                    {
                        "accountId": "synthetic-other-account",
                        "accountName": "Other Mailbox",
                        "primaryEmailAddress": "other@example.invalid",
                    },
                    {
                        "accountId": "synthetic-account-id",
                        "accountName": "Configured Mailbox",
                        "primaryEmailAddress": "sender@example.invalid",
                    },
                ]
            },
        )
        transport = FakeTransport(
            [_token("synthetic-access-token"), accounts_response]
        )

        account = ZohoMailClient(_config(), transport=transport).get_configured_account()

        self.assertEqual(account.account_id, "synthetic-account-id")
        self.assertEqual(account.primary_email, "sender@example.invalid")

    def test_configured_account_id_must_match_an_accessible_account(self) -> None:
        transport = FakeTransport([_token("synthetic-access-token"), _accounts()])
        config = ZohoMailConfig(
            client_id="synthetic-client-id",
            client_secret="synthetic-client-secret",
            refresh_token="synthetic-refresh-token",
            account_id="synthetic-not-configured",
            region="us",
        )

        with self.assertRaises(ZohoConfigurationError):
            ZohoMailClient(config, transport=transport).get_configured_account()

    def test_reuses_token_until_refresh_skew_window(self) -> None:
        clock = MutableClock()
        transport = FakeTransport(
            [_token("synthetic-access-token", expires_in=1600), _accounts(), _accounts()]
        )
        client = ZohoMailClient(_config(), transport=transport, clock=clock)

        client.get_accounts()
        clock.value = 1639
        client.get_accounts()
        self.assertEqual(len(transport.requests), 3)

        clock.value = 1640
        transport.responses.extend([_token("synthetic-next-token"), _accounts()])
        client.get_accounts()
        self.assertEqual(len(transport.requests), 5)

    def test_retries_once_with_refreshed_token_after_mail_api_401(self) -> None:
        transport = FakeTransport(
            [
                _token("synthetic-expired-token"),
                (401, b'{"error":"INVALID_OAUTHTOKEN"}'),
                _token("synthetic-fresh-token"),
                _accounts(),
            ]
        )
        client = ZohoMailClient(_config(), transport=transport)

        accounts = client.get_accounts()

        self.assertEqual(len(accounts), 1)
        self.assertEqual(len(transport.requests), 4)
        self.assertEqual(
            transport.requests[1]["headers"]["Authorization"],
            "Zoho-oauthtoken synthetic-expired-token",
        )
        self.assertEqual(
            transport.requests[3]["headers"]["Authorization"],
            "Zoho-oauthtoken synthetic-fresh-token",
        )

    def test_maps_token_refresh_to_configured_region(self) -> None:
        transport = FakeTransport([_token("synthetic-token"), _accounts()])
        client = ZohoMailClient(_config("ca"), transport=transport)

        client.get_accounts()

        self.assertEqual(
            transport.requests[0]["url"],
            "https://accounts.zohocloud.ca/oauth/v2/token",
        )
        self.assertEqual(
            transport.requests[1]["url"], "https://mail.zohocloud.ca/api/accounts"
        )

    def test_authentication_and_scope_errors_are_sanitized(self) -> None:
        cases = (
            (400, True, ZohoAuthenticationError),
            (401, False, ZohoAuthenticationError),
            (403, False, ZohoAuthorizationError),
            (404, False, ZohoRegionError),
            (429, False, ZohoRateLimitError),
            (503, False, ZohoServiceError),
        )
        for status, token_request, error_type in cases:
            with self.subTest(status=status, token_request=token_request):
                error_response = (
                    status,
                    b'{"error":"contains-sensitive-response"}',
                )
                if token_request:
                    responses = [error_response]
                elif status == 401:
                    responses = [
                        _token("synthetic-access-token"),
                        error_response,
                        _token("synthetic-refreshed-token"),
                        error_response,
                    ]
                else:
                    responses = [_token("synthetic-access-token"), error_response]
                transport = FakeTransport(responses)
                client = ZohoMailClient(_config(), transport=transport)
                with self.assertRaises(error_type) as raised:
                    client.get_accounts()
                self.assertNotIn("contains-sensitive-response", str(raised.exception))
                self.assertNotIn("synthetic-client-secret", str(raised.exception))

    def test_oauth_error_in_success_status_is_still_authentication_failure(self) -> None:
        transport = FakeTransport(
            [_json_response(200, {"error": "invalid_grant"})]
        )

        with self.assertRaises(ZohoAuthenticationError):
            ZohoMailClient(_config(), transport=transport).get_accounts()

    def test_second_401_after_refresh_stops_with_authentication_error(self) -> None:
        transport = FakeTransport(
            [
                _token("synthetic-expired-token"),
                (401, b"{}"),
                _token("synthetic-next-token"),
                (401, b"{}"),
            ]
        )
        with self.assertRaises(ZohoAuthenticationError):
            ZohoMailClient(_config(), transport=transport).get_accounts()
        self.assertEqual(len(transport.requests), 4)

    def test_network_failure_does_not_echo_underlying_exception(self) -> None:
        transport = FakeTransport()
        transport.error = URLError("synthetic-refresh-token")

        with self.assertRaises(ZohoConnectionError) as raised:
            ZohoMailClient(_config(), transport=transport).get_accounts()

        self.assertNotIn("synthetic-refresh-token", str(raised.exception))

    def test_malformed_token_and_account_responses_fail_explicitly(self) -> None:
        with self.subTest(response="token"):
            client = ZohoMailClient(_config(), transport=FakeTransport([(200, b"{}")]))
            with self.assertRaises(ZohoProtocolError):
                client.get_accounts()
        with self.subTest(response="account"):
            transport = FakeTransport(
                [_token("synthetic-access-token"), _json_response(200, {"data": {}})]
            )
            with self.assertRaises(ZohoProtocolError):
                ZohoMailClient(_config(), transport=transport).get_accounts()

    def test_attachment_upload_uses_mail_api_and_parses_uploaded_file(self) -> None:
        transport = FakeTransport(
            [
                _token("synthetic-access-token"),
                _json_response(
                    200,
                    {
                        "data": {
                            "storeName": "synthetic-store",
                            "attachmentName": "conditions.xlsx",
                            "attachmentPath": "/synthetic/path",
                        }
                    },
                ),
            ]
        )
        client = ZohoMailClient(_config(), transport=transport)

        result = client.upload_draft_attachment(
            file_name="conditions.xlsx",
            content=b"synthetic attachment",
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        self.assertEqual(result["attachmentName"], "conditions.xlsx")
        request = transport.requests[1]
        self.assertEqual(request["method"], "POST")
        self.assertTrue(
            str(request["url"]).startswith(
                "https://mail.zoho.com/api/accounts/synthetic-account-id/messages/attachments?"
            )
        )
        self.assertEqual(request["body"], b"synthetic attachment")

    def test_attachment_upload_rejects_invalid_content_type_before_network(self) -> None:
        transport = FakeTransport()
        client = ZohoMailClient(_config(), transport=transport)

        with self.assertRaises(ZohoConfigurationError):
            client.upload_draft_attachment(
                file_name="conditions.xlsx",
                content=b"synthetic attachment",
                content_type="application/octet-stream\r\nInjected: value",
            )

        self.assertEqual(transport.requests, [])

    def test_draft_save_requires_mode_before_network_and_sends_mode_draft(self) -> None:
        transport = FakeTransport()
        client = ZohoMailClient(_config(), transport=transport)
        with self.assertRaises(ZohoConfigurationError):
            client.save_draft(
                {
                    "fromAddress": "sender@example.invalid",
                    "toAddress": "recipient@example.invalid",
                    "subject": "Synthetic subject",
                    "content": "Synthetic body",
                }
            )
        self.assertEqual(transport.requests, [])

        transport = FakeTransport(
            [
                _token("synthetic-access-token"),
                _json_response(200, {"status": {"code": 200}, "data": {}}),
            ]
        )
        client = ZohoMailClient(_config(), transport=transport)
        client.save_draft(
            {
                "fromAddress": "sender@example.invalid",
                "toAddress": "recipient@example.invalid",
                "subject": "Synthetic subject",
                "content": "Synthetic body",
                "mode": "draft",
                "attachments": [
                    {
                        "storeName": "synthetic-store",
                        "attachmentName": "conditions.xlsx",
                        "attachmentPath": "/synthetic/path",
                    }
                ],
            }
        )

        request = transport.requests[1]
        self.assertEqual(request["method"], "POST")
        self.assertEqual(
            request["url"],
            "https://mail.zoho.com/api/accounts/synthetic-account-id/messages",
        )
        payload = json.loads(request["body"])
        self.assertEqual(payload["mode"], "draft")
        self.assertEqual(len(payload["attachments"]), 1)

    def test_draft_save_rejects_success_response_without_confirmation(self) -> None:
        transport = FakeTransport(
            [_token("synthetic-access-token"), _json_response(200, {"data": {}})]
        )
        client = ZohoMailClient(_config(), transport=transport)
        with self.assertRaises(ZohoServiceError):
            client.save_draft(
                {
                    "fromAddress": "sender@example.invalid",
                    "toAddress": "recipient@example.invalid",
                    "subject": "Synthetic subject",
                    "content": "Synthetic body",
                    "mode": "draft",
                }
            )


if __name__ == "__main__":
    unittest.main()
