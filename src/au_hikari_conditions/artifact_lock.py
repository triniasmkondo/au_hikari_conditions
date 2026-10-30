"""Freeze user-approved Word notices without recalculating prices or approving mail."""
from __future__ import annotations

import hashlib
import json
import re
import shutil
import tempfile
import unicodedata
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
SCHEMA = 'approved-condition-artifacts-v1'
ROLES = {'conditions', 'campaign'}


def file_sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _write_json(path: Path, data: object) -> None:
    with path.open('x', encoding='utf-8') as stream:
        json.dump(data, stream, ensure_ascii=False, indent=2)
        stream.write('\n')


def _safe_path(root: Path, relative: str) -> Path:
    value = Path(relative)
    if value.is_absolute() or '..' in value.parts:
        raise ValueError('Absolute paths and traversal are forbidden')
    path = (root / value).resolve()
    if not path.is_relative_to(root.resolve()):
        raise ValueError('Path escapes root')
    return path


def monetary_cells(path: Path) -> list[dict]:
    """Keep source text and numeric values, including Word text-box tables and dashes."""
    with ZipFile(path) as archive:
        root = ET.fromstring(archive.read('word/document.xml'))
    result = []
    for ti, table in enumerate(root.findall('.//w:tbl', NS), 1):
        for ri, row in enumerate(table.findall('w:tr', NS), 1):
            for ci, cell in enumerate(row.findall('w:tc', NS), 1):
                text = ''.join(n.text or '' for n in cell.findall('.//w:t', NS))
                normalized = unicodedata.normalize('NFKC', text).strip()
                # Preserve "-" as a state, never turn it into zero. Exclude count thresholds.
                pure = re.fullmatch(r'[¥￥]?\s*([−－-]?\d[\d,]*)\s*(?:円(?:\s*/\s*件)?)?', normalized)
                if pure:
                    value = int(pure.group(1).replace(',', '').replace('−', '-').replace('－', '-'))
                    result.append(dict(table=ti, row=ri, cell=ci, text=text, state='VALUE', amount=value))
                elif normalized in ('-', '－', '—', '―'):
                    result.append(dict(table=ti, row=ri, cell=ci, text=text, state='DASH', amount=None))
    return result


def create_lock(root: Path, plan_path: Path, output: Path) -> dict:
    root = root.resolve()
    plan = json.loads(plan_path.read_text(encoding='utf-8'))
    if plan.get('approval', {}).get('status') != 'APPROVED_BY_USER':
        raise ValueError('Explicit user approval is required')
    if not plan['approval'].get('statement') or not plan['approval'].get('source'):
        raise ValueError('Approval statement and source are required')
    agencies = plan['agencies']
    if not agencies or len({a['agency'] for a in agencies}) != len(agencies):
        raise ValueError('Empty or duplicate agencies')
    excluded = set(plan.get('excluded_agencies', []))
    if excluded & {a['agency'] for a in agencies}:
        raise ValueError('Excluded agency in active lock')
    documents = []
    seen = set()
    for agency in agencies:
        if set(agency['documents']) != ROLES:
            raise ValueError('Each agency requires conditions and campaign notices')
        if Path(agency['agency']).name != agency['agency']:
            raise ValueError('Agency name must be a single path component')
        for role, relative in agency['documents'].items():
            path = _safe_path(root, relative)
            if path.suffix.lower() != '.docx' or not path.is_file() or str(path) in seen:
                raise ValueError('Missing, duplicate or non-Word artifact')
            seen.add(str(path))
            cells = monetary_cells(path)
            if not cells:
                raise ValueError('Word artifact contains no monetary cells')
            documents.append(dict(agency=agency['agency'], role=role, source_path=relative,
                                  path=f"documents/{agency['agency']}/{path.name}",
                                  sha256=file_sha256(path), monetary_cells=cells))
    if output.exists():
        raise FileExistsError('Existing LOCK cannot be overwritten')
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='.condition-lock-', dir=output.parent) as temporary:
        stage = Path(temporary) / 'release'
        stage.mkdir()
        files = []
        prices = []
        for document in documents:
            source = _safe_path(root, document['source_path'])
            target = _safe_path(stage, document['path'])
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(source, target)
            if file_sha256(target) != document['sha256'] or file_sha256(source) != document['sha256']:
                raise ValueError('Source changed during lock creation')
            prices.append(dict(agency=document['agency'], role=document['role'],
                               document_sha256=document['sha256'], cells=document.pop('monetary_cells')))
            files.append(dict(path=document['path'], sha256=document['sha256']))
        approval = dict(plan['approval'], recorded_at=datetime.now(timezone.utc).isoformat(),
                        scope='condition documents and prices only', email_send_approved=False)
        _write_json(stage / 'approval.json', approval)
        _write_json(stage / 'amounts.json', prices)
        for name in ('approval.json', 'amounts.json'):
            files.append(dict(path=name, sha256=file_sha256(stage / name)))
        manifest = dict(schema=SCHEMA, period=plan['period'], state='LOCKED', agency_count=len(agencies),
                        document_count=len(documents), excluded_agencies=sorted(excluded),
                        documents=documents, files=files,
                        approval_sha256=file_sha256(stage / 'approval.json'),
                        amounts_sha256=file_sha256(stage / 'amounts.json'))
        _write_json(stage / 'manifest.json', manifest)
        verify_lock(root, stage, file_sha256(stage / 'manifest.json'))
        stage.rename(output)
    return dict(state='LOCKED', agency_count=len(agencies), document_count=len(documents),
                manifest_sha256=file_sha256(output / 'manifest.json'))


def verify_lock(root: Path, output: Path, expected_manifest_sha256: str) -> dict:
    if file_sha256(output / 'manifest.json') != expected_manifest_sha256:
        raise ValueError('LOCK manifest fingerprint mismatch')
    manifest = json.loads((output / 'manifest.json').read_text(encoding='utf-8'))
    if manifest.get('schema') != SCHEMA or manifest.get('state') != 'LOCKED':
        raise ValueError('Unknown LOCK schema or state')
    paths = [f['path'] for f in manifest['files']]
    if len(paths) != len(set(paths)):
        raise ValueError('Duplicate LOCK file paths')
    for record in manifest['files']:
        if file_sha256(_safe_path(output, record['path'])) != record['sha256']:
            raise ValueError('Locked file changed: ' + record['path'])
    expected_files = set(paths) | {'manifest.json'}
    actual_files = {p.relative_to(output).as_posix() for p in output.rglob('*') if p.is_file()}
    if actual_files != expected_files:
        raise ValueError('Missing or unlisted files in LOCK')
    approval = json.loads((output / 'approval.json').read_text(encoding='utf-8'))
    if approval['status'] != 'APPROVED_BY_USER' or approval.get('email_send_approved') is not False:
        raise ValueError('Invalid artifact-only approval')
    prices = json.loads((output / 'amounts.json').read_text(encoding='utf-8'))
    price_index = {(p['agency'], p['role']): p for p in prices}
    if len(price_index) != len(prices):
        raise ValueError('Duplicate monetary snapshot')
    roles = {}
    for document in manifest['documents']:
        key = (document['agency'], document['role'])
        if document['agency'] in manifest['excluded_agencies']:
            raise ValueError('Excluded agency in LOCK')
        if document['role'] not in ROLES or key in roles:
            raise ValueError('Duplicate or unknown document role')
        roles[key] = True
        if file_sha256(_safe_path(root, document['source_path'])) != document['sha256']:
            raise ValueError('Approved original changed: ' + document['source_path'])
        if file_sha256(_safe_path(output, document['path'])) != document['sha256']:
            raise ValueError('Approved copy changed')
        saved = price_index.get(key)
        if not saved or saved['document_sha256'] != document['sha256'] or saved['cells'] != monetary_cells(_safe_path(output, document['path'])):
            raise ValueError('Locked monetary snapshot mismatch')
    agency_names = {a for a, _ in roles}
    if len(agency_names) != manifest['agency_count'] or len(roles) != manifest['document_count']:
        raise ValueError('LOCK counts mismatch')
    if set(roles) != {(a, r) for a in agency_names for r in ROLES} or set(price_index) != set(roles):
        raise ValueError('Incomplete agency document pairs')
    return dict(state='LOCKED', agency_count=len(agency_names), document_count=len(roles),
                manifest_sha256=expected_manifest_sha256)
