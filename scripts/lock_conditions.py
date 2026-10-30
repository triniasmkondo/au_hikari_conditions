"""Create or verify a local approved-document LOCK. Never sends email."""
import argparse
import json
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from au_hikari_conditions.artifact_lock import create_lock, verify_lock


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('operation', choices=['create', 'verify'])
    parser.add_argument('--root', type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--plan', type=Path)
    parser.add_argument('--expected-manifest-sha256')
    args = parser.parse_args()
    if args.operation == 'create':
        if not args.plan:
            parser.error('create requires --plan with explicit user approval')
        result = create_lock(args.root, args.plan, args.output)
    else:
        if not args.expected_manifest_sha256:
            parser.error('verify requires --expected-manifest-sha256 from the committed receipt')
        result = verify_lock(args.root, args.output, args.expected_manifest_sha256)
    print(json.dumps(result, ensure_ascii=False))


if __name__ == '__main__':
    main()
