"""Prepare 10 locked PDF pairs or save exactly one authorized Zoho draft; no sending."""
import argparse
import importlib.util
import json
from pathlib import Path
import sys

ROOT=Path(__file__).resolve().parents[1]
spec=importlib.util.spec_from_file_location('locked_delivery_runtime',ROOT/'src/au_hikari_conditions/locked_delivery.py')
runtime=importlib.util.module_from_spec(spec);sys.modules[spec.name]=runtime;spec.loader.exec_module(runtime)


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('operation',choices=['prepare','probe','upload-probe','draft','remaining-nine'])
    parser.add_argument('--master',type=Path)
    parser.add_argument('--template',type=Path)
    parser.add_argument('--period',default='2026-11')
    parser.add_argument('--pdf-root',type=Path)
    parser.add_argument('--prepared',type=Path)
    parser.add_argument('--expected-prepared-sha256')
    parser.add_argument('--agency-id')
    parser.add_argument('--history',type=Path)
    parser.add_argument('--result',type=Path)
    parser.add_argument('--token-file',type=Path)
    parser.add_argument('--sender')
    parser.add_argument('--authorize-disabled-draft',action='store_true')
    parser.add_argument('--retry-failed-upload',action='store_true')
    args=parser.parse_args()
    try:
        if args.operation=='probe':
            _,account=runtime.load_zoho_client(args.token_file,args.sender)
            result=dict(status='ACCOUNT_VERIFIED',sender=account.sender_address,draft_created=False)
        elif args.operation=='remaining-nine':
            if not all([args.prepared,args.expected_prepared_sha256,args.history,args.result,args.token_file,args.sender]):
                parser.error('remaining-nine requires prepared, fingerprint, history, result directory, token file and sender')
            result=runtime.save_remaining_nine(ROOT,args.prepared,args.expected_prepared_sha256,args.history,args.result,
                token_file=args.token_file,requested_sender=args.sender,retry_failed_upload=args.retry_failed_upload)
        elif args.operation=='upload-probe':
            if not all([args.prepared,args.expected_prepared_sha256,args.agency_id,args.sender]):parser.error('upload-probe requires prepared, fingerprint, agency and sender')
            batch=runtime.verify_prepared(ROOT,args.prepared,args.expected_prepared_sha256)
            matches=[m for m in batch['messages'] if m['agency_id']==args.agency_id]
            if len(matches)!=1:raise ValueError('Exactly one agency required')
            client,_=runtime.load_zoho_client(args.token_file,args.sender)
            original=client._transport.request
            def traced(**kwargs):
                status,body=original(**kwargs)
                if '/messages/attachments?' in kwargs['url']:
                    print(json.dumps(dict(upload_http_status=status,content_type=kwargs['headers'].get('Content-Type'))))
                return status,body
            from types import SimpleNamespace
            client._transport=SimpleNamespace(request=traced)
            a=matches[0]['attachments'][0]
            client.upload_draft_attachment(file_name=a['file_name'],content=runtime.relative_path(ROOT,a['path']).read_bytes(),content_type='application/octet-stream')
            result=dict(status='UPLOAD_VERIFIED',draft_created=False)
        elif args.operation=='prepare':
            if not all([args.master,args.template,args.prepared]):parser.error('prepare requires master, template and prepared')
            result=runtime.prepare_batch(ROOT,args.master,args.template,args.pdf_root or ROOT/'04_成果物'/f'{args.period}_PDF',args.prepared,args.period)
        else:
            if not all([args.prepared,args.expected_prepared_sha256,args.agency_id,args.history,args.result,args.sender]):
                parser.error('draft requires prepared, fingerprint, agency, history and result')
            result=runtime.save_one(ROOT,args.prepared,args.expected_prepared_sha256,args.agency_id,args.history,args.result,
                token_file=args.token_file,authorize_disabled_draft=args.authorize_disabled_draft,requested_sender=args.sender,retry_failed_upload=args.retry_failed_upload)
        print(json.dumps(result,ensure_ascii=False));return 0
    except Exception as error:
        print(json.dumps(dict(status='BLOCKED',error_category=type(error).__name__,
                              details=str(error) if (isinstance(error,(ValueError,FileExistsError)) or type(error).__name__.startswith('Zoho')) else 'See sanitized error category. Never retry an uncertain draft POST.',sent=False),ensure_ascii=False),file=sys.stderr)
        return 2


if __name__=='__main__':raise SystemExit(main())
