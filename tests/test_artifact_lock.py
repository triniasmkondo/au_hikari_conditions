import json
from pathlib import Path
import tempfile
import unittest
from zipfile import ZipFile
from au_hikari_conditions.artifact_lock import create_lock, verify_lock, monetary_cells


def word(path):
    text = '''<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tr><w:tc><w:p><w:r><w:t>２，０００円／件</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>-</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>0</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>'''
    with ZipFile(path, 'w') as z:
        z.writestr('word/document.xml', text)


class ArtifactLockTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        word(self.root / 'conditions.docx'); word(self.root / 'campaign.docx')
        self.output = self.root / 'LOCK'
        self.plan = self.root / 'plan.json'
        self.data = dict(period='2026-11', excluded_agencies=['Excluded'],
                         approval=dict(status='APPROVED_BY_USER', source='user conversation', statement='承認します'),
                         agencies=[dict(agency='Example', documents=dict(conditions='conditions.docx', campaign='campaign.docx'))])
        self.save()

    def save(self):
        self.plan.write_text(json.dumps(self.data), encoding='utf-8')

    def create(self):
        return create_lock(self.root, self.plan, self.output)['manifest_sha256']

    def test_exact_snapshot_and_no_send_approval(self):
        digest = self.create()
        result = verify_lock(self.root, self.output, digest)
        self.assertEqual(result['document_count'], 2)
        self.assertEqual((self.output / 'documents/Example/conditions.docx').read_bytes(), (self.root / 'conditions.docx').read_bytes())
        cells = monetary_cells(self.root / 'conditions.docx')
        self.assertEqual([c['amount'] for c in cells], [2000, None, 0])
        approval = json.loads((self.output / 'approval.json').read_text())
        self.assertFalse(approval['email_send_approved'])
        with self.assertRaises(FileExistsError):self.create()

    def test_source_copy_manifest_and_price_tampering_fail(self):
        digest = self.create()
        for path in [self.root/'conditions.docx', self.output/'documents/Example/conditions.docx',
                     self.output/'amounts.json', self.output/'manifest.json']:
            original = path.read_bytes();path.write_bytes(original+b'changed')
            with self.assertRaises(ValueError):verify_lock(self.root, self.output, digest)
            path.write_bytes(original)
        self.assertEqual(verify_lock(self.root, self.output, digest)['state'], 'LOCKED')

    def test_exclusion_missing_pair_and_unapproved_plan_fail(self):
        self.data['excluded_agencies']=['Example'];self.save()
        with self.assertRaises(ValueError):self.create()
        self.data['excluded_agencies']=[]
        self.data['agencies'][0]['documents'].pop('campaign');self.save()
        with self.assertRaises(ValueError):self.create()
        self.data['approval']['status']='PENDING';self.save()
        with self.assertRaises(ValueError):self.create()
        self.assertFalse(self.output.exists())

    def test_path_escape_and_unlisted_file_fail(self):
        self.data['agencies'][0]['documents']['conditions']='../conditions.docx';self.save()
        with self.assertRaises(ValueError):self.create()
        self.data['agencies'][0]['documents']['conditions']='conditions.docx';self.save()
        digest=self.create();(self.output/'extra.docx').write_bytes(b'extra')
        with self.assertRaises(ValueError):verify_lock(self.root,self.output,digest)
