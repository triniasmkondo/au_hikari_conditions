from pathlib import Path
from types import SimpleNamespace
import tempfile
import unittest
from unittest.mock import patch
from email.message import EmailMessage
from au_hikari_conditions import locked_delivery as delivery


class FakeClient:
    def __init__(self,fail=False):self.posts=[];self.uploads=[];self.fail=fail
    def upload_draft_attachment(self,**kwargs):
        self.uploads.append(kwargs)
        return dict(attachmentName=kwargs['file_name'],storeName='store',attachmentPath='path')
    def save_draft(self,payload):
        self.posts.append(payload)
        if self.fail:raise TimeoutError('uncertain result')
        return dict(data=dict(messageId='123'),status=dict(code=200))


class LockedDeliveryTest(unittest.TestCase):
    def test_body_contact_and_linebreaks_and_legacy_rejection(self):
        body='{mail_company_name}\n{mail_contact_name}\n\n{target_year_month_ja}度のご案内'
        self.assertEqual(delivery.render_body(body,None,'高橋様','2026-11'),'高橋様\n\n2026年11月度のご案内')
        self.assertEqual(delivery.render_body(body,'会社','後道様\n上田様','2026-11'),'会社\n後道様\n上田様\n\n2026年11月度のご案内')
        with self.assertRaises(ValueError):delivery.render_body(body,'会社','','2026-11')
        with self.assertRaises(ValueError):delivery.render_body('{mail_addressee}',None,'担当様','2026-11')

    def test_address_display_names_and_injection_rejection(self):
        self.assertEqual(delivery.parse_addresses('"担当者" <one@example.com>; two@example.com'),['one@example.com','two@example.com'])
        for value in ['bad;one@example.com','one@example.com\nBcc:other@example.com','one@example.com;one@example.com']:
            with self.assertRaises(ValueError):delivery.parse_addresses(value)

    def fixture(self,root):
        attachments=[]
        for role in ('conditions','campaign'):
            p=root/(role+'.pdf');p.write_bytes(b'%PDF-test-'+role.encode())
            attachments.append(dict(path=p.name,file_name=p.name,sha256=delivery.digest(p),role=role))
        message=dict(agency_id='example',agency_name='会社',enabled='NO',mail_contact_name='担当様',
                     to=['one@example.com'],cc=[],bcc=[],body='会社\n担当様\n\n本文',
                     subject=delivery.render_company_subject('2026-11','会社'),attachments=attachments,message_hash='hash')
        return dict(period='2026-11',messages=[message]),message

    def run_fake(self,root,prepared,client,result='result.json',authorized=True):
        with patch.object(delivery,'verify_prepared',return_value=prepared),patch.object(delivery,'_read_original',side_effect=ValueError('readback unavailable')):
            return delivery.save_one(root,root/'prepared.json','hash','example',root/'history.sqlite3',root/result,
                authorize_disabled_draft=authorized,client=client,account=SimpleNamespace(primary_email='sender@example.com'))

    def test_one_draft_only_and_duplicate_is_blocked(self):
        with tempfile.TemporaryDirectory() as d:
            root=Path(d);prepared,message=self.fixture(root);client=FakeClient()
            result=self.run_fake(root,prepared,client)
            self.assertEqual(result['status'],'CREATED');self.assertFalse(result['sent'])
            self.assertEqual(len(client.posts),1);self.assertEqual(client.posts[0]['mode'],'draft')
            self.assertEqual(client.posts[0]['subject'],message['subject']);self.assertEqual(client.posts[0]['content'],message['body'])
            self.assertEqual(len(client.uploads),2)
            with self.assertRaises(ValueError):self.run_fake(root,prepared,client,result='retry.json')
            self.assertEqual(len(client.posts),1)

    def test_disabled_and_changed_pdf_never_post(self):
        with tempfile.TemporaryDirectory() as d:
            root=Path(d);prepared,_=self.fixture(root);client=FakeClient()
            with self.assertRaises(ValueError):self.run_fake(root,prepared,client,authorized=False)
            (root/'campaign.pdf').write_bytes(b'changed')
            with self.assertRaises(ValueError):self.run_fake(root,prepared,client)
            self.assertEqual(client.posts,[])

    def test_uncertain_post_is_never_retried(self):
        with tempfile.TemporaryDirectory() as d:
            root=Path(d);prepared,_=self.fixture(root);client=FakeClient(fail=True)
            with self.assertRaises(TimeoutError):self.run_fake(root,prepared,client)
            with self.assertRaises(ValueError):self.run_fake(root,prepared,client)
            self.assertEqual(len(client.posts),1)

    def test_mime_subject_recipients_body_and_exact_pdf_bytes(self):
        with tempfile.TemporaryDirectory() as d:
            root=Path(d);_,message=self.fixture(root)
            mime=EmailMessage();mime['Subject']=message['subject'];mime['From']='sender@example.com';mime['To']='one@example.com'
            mime.set_content(message['body'])
            for a in message['attachments']:
                mime.add_attachment((root/a['path']).read_bytes(),maintype='application',subtype='pdf',filename=a['file_name'])
            delivery.verify_mime(mime.as_string(),message,'sender@example.com')
            mime.replace_header('Subject','wrong')
            with self.assertRaises(ValueError):delivery.verify_mime(mime.as_string(),message,'sender@example.com')

    def test_explicit_registered_alias_is_used_without_primary_fallback(self):
        records=[dict(type='ZOHO_ACCOUNT',enabled=True,accountId='example-account',primaryEmailAddress='primary@example.com',
            emailAddress=[dict(mailId='alias@example.com',isConfirmed=True)],
            sendMailDetails=[dict(fromAddress='alias@example.com',status=True)])]
        account=delivery.choose_sending_account(records,'alias@example.com')
        self.assertEqual(account.sender_address,'alias@example.com')
        self.assertEqual(account.primary_email,'primary@example.com')
        with self.assertRaises(ValueError):delivery.choose_sending_account(records,'other@example.com')
        records[0]['emailAddress'][0]['isConfirmed']=False
        with self.assertRaises(ValueError):delivery.choose_sending_account(records,'alias@example.com')

    def test_from_uses_requested_alias_and_not_account_primary(self):
        with tempfile.TemporaryDirectory() as d:
            root=Path(d);prepared,message=self.fixture(root);client=FakeClient()
            account=SimpleNamespace(primary_email='primary@example.com',sender_address='alias@example.com')
            with patch.object(delivery,'verify_prepared',return_value=prepared),patch.object(delivery,'_read_original',side_effect=ValueError('unavailable')):
                delivery.save_one(root,root/'prepared.json','hash','example',root/'history.sqlite3',root/'result.json',
                    authorize_disabled_draft=True,requested_sender='alias@example.com',client=client,account=account)
            self.assertEqual(client.posts[0]['fromAddress'],'alias@example.com')
            self.assertEqual(client.posts[0]['content'],message['body'])
            self.assertTrue(all(a['content_type']=='application/octet-stream' for a in client.uploads))

    def test_only_confirmed_pre_draft_upload_failure_can_be_retried(self):
        class FailedUpload(FakeClient):
            def upload_draft_attachment(self,**kwargs):raise ValueError('upload rejected')
        with tempfile.TemporaryDirectory() as d:
            root=Path(d);prepared,_=self.fixture(root)
            with self.assertRaises(ValueError):self.run_fake(root,prepared,FailedUpload())
            client=FakeClient()
            with self.assertRaises(ValueError):self.run_fake(root,prepared,client)
            with patch.object(delivery,'verify_prepared',return_value=prepared),patch.object(delivery,'_read_original',side_effect=ValueError('unavailable')):
                delivery.save_one(root,root/'prepared.json','hash','example',root/'history.sqlite3',root/'result.json',
                    authorize_disabled_draft=True,retry_failed_upload=True,client=client,account=SimpleNamespace(primary_email='sender@example.com'))
            self.assertEqual(len(client.posts),1)
            with self.assertRaises(ValueError):
                with patch.object(delivery,'verify_prepared',return_value=prepared):
                    delivery.save_one(root,root/'prepared.json','hash','example',root/'history.sqlite3',root/'retry.json',
                        authorize_disabled_draft=True,retry_failed_upload=True,client=client,account=SimpleNamespace(primary_email='sender@example.com'))

    def test_deferred_readback_does_not_query_or_retry_draft(self):
        with tempfile.TemporaryDirectory() as d:
            root=Path(d);prepared,_=self.fixture(root);client=FakeClient()
            with patch.object(delivery,'verify_prepared',return_value=prepared),patch.object(delivery,'_read_original') as read:
                result=delivery.save_one(root,root/'prepared.json','hash','example',root/'history.sqlite3',root/'result.json',
                    authorize_disabled_draft=True,skip_readback=True,client=client,account=SimpleNamespace(primary_email='sender@example.com'))
            read.assert_not_called();self.assertEqual(result['readback_status'],'DEFERRED_BY_USER')
            self.assertEqual(len(client.posts),1);self.assertEqual(client.posts[0]['mode'],'draft')

    def test_remaining_nine_creates_only_nine_and_preserves_keioh(self):
        import sqlite3
        with tempfile.TemporaryDirectory() as d:
            root=Path(d);_,message=self.fixture(root)
            ids=['netconnect','an_ocean','centrix','j_and_c','duo_communications','value_ecology','takahashi_daichi','ichimaru','hualan']
            batch=dict(period='2026-11',messages=[dict(message,agency_id=i) for i in ids])
            history=root/'history.sqlite3'
            with delivery.database(history) as db:
                db.execute('CREATE TABLE locked_pdf_drafts (agency_id TEXT,period TEXT,request_hash TEXT,status TEXT,message_id TEXT,recorded_at TEXT,PRIMARY KEY(agency_id,period))')
                db.execute('INSERT INTO locked_pdf_drafts VALUES (?,?,?,?,?,?)',('keioh','2026-11','prior','CREATED','prior-id','prior-time'))
            client=FakeClient();account=SimpleNamespace(primary_email='sender@example.com')
            with patch.object(delivery,'verify_prepared',return_value=batch),patch.object(delivery,'load_zoho_client',return_value=(client,account)),patch.object(delivery,'_read_original') as read,patch('builtins.print'):
                result=delivery.save_remaining_nine(root,root/'prepared.json','hash',history,root,
                    token_file=root/'unused',requested_sender='sender@example.com')
            self.assertEqual(result['new_draft_count'],9);self.assertEqual(len(client.posts),9)
            self.assertTrue(all(p['mode']=='draft' for p in client.posts));read.assert_not_called()
            with delivery.database(history) as db:
                self.assertEqual(db.execute('SELECT count(*) FROM locked_pdf_drafts WHERE status="CREATED"').fetchone()[0],10)
                self.assertEqual(db.execute('SELECT message_id FROM locked_pdf_drafts WHERE agency_id="keioh"').fetchone()[0],'prior-id')

    def test_html_escaped_filename_metadata_preserves_original_attachment_name(self):
        class EscapingClient(FakeClient):
            def upload_draft_attachment(self,**kwargs):
                value=super().upload_draft_attachment(**kwargs)
                value['attachmentName']=value['attachmentName'].replace('&','&amp;')
                return value
        with tempfile.TemporaryDirectory() as d:
            root=Path(d);prepared,message=self.fixture(root)
            p=root/'J&C.pdf';p.write_bytes(b'%PDF-conditions')
            message['attachments'][0].update(path=p.name,file_name=p.name,sha256=delivery.digest(p))
            client=EscapingClient();self.run_fake(root,prepared,client)
            self.assertEqual(client.posts[0]['attachments'][0]['attachmentName'],'J&amp;C.pdf')

    def test_stripped_filename_is_saved_and_flagged_for_manual_review(self):
        class StrippingClient(FakeClient):
            def upload_draft_attachment(self,**kwargs):
                value=super().upload_draft_attachment(**kwargs)
                value['attachmentName']=value['attachmentName'].replace('&','')
                return value
        with tempfile.TemporaryDirectory() as d:
            root=Path(d);prepared,message=self.fixture(root)
            p=root/'J&C.pdf';p.write_bytes(b'%PDF-conditions')
            message['attachments'][0].update(path=p.name,file_name=p.name,sha256=delivery.digest(p))
            client=StrippingClient()
            result=self.run_fake(root,prepared,client)
            self.assertEqual(len(client.posts),1)
            self.assertTrue(result['attachment_name_changed'])
            self.assertTrue(result['requires_manual_review'])
            self.assertEqual(result['attachment_name_mapping'][0]['original'],'J&C.pdf')
            self.assertEqual(result['attachment_name_mapping'][0]['zoho'],'JC.pdf')

    def test_december_company_subject(self):
        self.assertEqual(delivery.render_company_subject('2026-12','keioh株式会社'),
            '【202612】【keioh株式会社様】auひかり(AL商流)/auひかりプラス_条件表通知表')
