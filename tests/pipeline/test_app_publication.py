import io
import json
import sys
import unittest
from pathlib import Path
from unittest.mock import patch
from types import SimpleNamespace
from urllib.error import HTTPError, URLError

from botocore.exceptions import ClientError

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / 'pipeline'))
from process_job import complete_app_publication, PublicationCompletionError, JobCancelled, NoPublicationRedirect
import process_job

JOB = '11111111-1111-4111-8111-111111111111'
OWNER = '22222222-2222-4222-8222-222222222222'


class Storage:
    def __init__(self):
        self.body = dict(operationId=JOB, ownerId=OWNER, token='a'*64)
        self.cancelled = False

    def get_object(self, **kw):
        assert kw['Key'] == f'jobs/{JOB}/app-publication.json'
        return {'Body': io.BytesIO(json.dumps(self.body).encode())}

    def head_object(self, **kw):
        if self.cancelled: return {}
        raise ClientError({'Error': {'Code': '404'}}, 'HeadObject')


class Response(io.BytesIO):
    status = 200


class Client:
    def __init__(self, failures=()):
        self.failures = list(failures)
        self.calls = []
        self.result = dict(published=True, snapshotId=JOB, bookId=JOB)

    def open(self, request, timeout):
        self.calls.append(request)
        assert timeout == 30
        if self.failures: raise self.failures.pop(0)
        return Response(json.dumps(self.result).encode())


class PublicationTests(unittest.TestCase):
    def setUp(self):
        self.env = patch.dict('os.environ', PLAYBOOK_PUBLIC_ORIGIN='https://playbook.example')
        self.env.start(); self.addCleanup(self.env.stop)
        self.storage, self.client, self.sleeps = Storage(), Client(), []

    def run_callback(self):
        return complete_app_publication(self.storage, 'jobs-only', JOB, opener=self.client, sleep=self.sleeps.append)

    def test_callback_contains_only_job_capability_and_matches_ack(self):
        self.assertTrue(self.run_callback()['published'])
        request = self.client.calls[0]
        self.assertEqual(request.full_url, 'https://playbook.example/api/app-publication/complete')
        self.assertEqual(json.loads(request.data), self.storage.body)
        self.assertNotIn('Authorization', request.headers)
        self.assertEqual(request.get_header('User-agent'), 'GSS-Playbook-Renderer/1.0')
        self.assertEqual(request.get_header('Accept'), 'application/json')

    def test_transient_and_lost_ack_retries_are_bounded_and_same_identity(self):
        self.client.failures = [URLError('synthetic'), TimeoutError()]
        self.assertTrue(self.run_callback()['published'])
        self.assertEqual(self.sleeps, [1, 2])
        self.assertEqual(len({r.data for r in self.client.calls}), 1)

    def test_edge_admits_identified_renderer_without_weakening_capability_auth(self):
        original_open = self.client.open
        def edge(request, timeout):
            if request.get_header('User-agent') != 'GSS-Playbook-Renderer/1.0':
                raise HTTPError(request.full_url, 403, 'Forbidden', {}, None)
            self.assertEqual(json.loads(request.data), self.storage.body)
            return original_open(request, timeout)
        self.client.open = edge
        self.assertTrue(self.run_callback()['published'])
        self.assertEqual(len(self.client.calls), 1)

    def test_three_failures_stop_without_exposing_credentials(self):
        self.client.failures = [URLError('sensitive-detail')]*3
        with self.assertRaisesRegex(PublicationCompletionError, 'could not be confirmed'):
            self.run_callback()
        self.assertEqual(len(self.client.calls), 3)

    def test_forbidden_is_not_retried_or_redirected(self):
        self.client.failures = [HTTPError('https://playbook.example', 403, 'private', {}, None)]
        with self.assertRaises(PublicationCompletionError): self.run_callback()
        self.assertEqual(self.sleeps, [])
        self.assertIsNone(NoPublicationRedirect().redirect_request(None, None, 302, '', {}, 'https://elsewhere.example'))

    def test_wrong_job_and_cancelled_job_never_send_callback(self):
        self.storage.body['operationId'] = OWNER
        with self.assertRaises(PublicationCompletionError): self.run_callback()
        self.storage.body['operationId'] = JOB; self.storage.cancelled = True
        with self.assertRaises(JobCancelled): self.run_callback()
        self.assertEqual(self.client.calls, [])

    def test_wrong_response_cannot_claim_publication(self):
        self.client.result['snapshotId'] = OWNER
        with self.assertRaises(PublicationCompletionError): self.run_callback()

    def test_job_cannot_choose_callback_origin(self):
        self.storage.body['url'] = 'https://elsewhere.example'
        with self.assertRaises(PublicationCompletionError): self.run_callback()
        self.assertEqual(self.client.calls, [])

    def test_invalid_callback_configuration_never_sends_capability(self):
        for origin in ['', 'http://playbook.example', 'https://user:secret@playbook.example',
                       'https://playbook.example/path', 'https://playbook.example?token=secret',
                       'https://playbook.example:bad', 'https://[broken',
                       'https://playbook.example\n']:
            with self.subTest(origin=origin), patch.dict('os.environ', PLAYBOOK_PUBLIC_ORIGIN=origin):
                with self.assertRaisesRegex(PublicationCompletionError, '^Publication callback is not configured$'):
                    self.run_callback()
        self.assertEqual(self.client.calls, [])

    def test_missing_callback_fails_job_before_download_or_render(self):
        status = {'mode': 'images', 'publicationId': JOB, 'ownerId': OWNER}
        self.storage.get_object = lambda **kw: {'Body': io.BytesIO(json.dumps(status).encode())}
        def put(**kw):
            status.clear(); status.update(json.loads(kw['Body']))
        self.storage.put_object = put
        with patch.dict('os.environ', R2_BUCKET='jobs-only', PLAYBOOK_PUBLIC_ORIGIN=''), \
             patch.object(sys, 'argv', ['process_job.py', JOB]), \
             patch.object(process_job, 'get_r2_client', return_value=self.storage), \
             patch.object(process_job, 'download_play_images') as download, \
             patch.object(process_job, 'complete_app_publication') as callback:
            with self.assertRaises(SystemExit) as result:
                process_job.main()
            self.assertEqual(result.exception.code, 1)
            download.assert_not_called(); callback.assert_not_called()
        self.assertEqual(status['status'], 'error')
        self.assertEqual(status['message'], 'Publication callback is not configured')
        self.assertEqual(status['publicationId'], JOB)

    def test_orchestrator_calls_publication_after_upload_without_browser(self):
        self.assert_pipeline_completion(False)

    def test_orchestrator_keeps_rendered_files_when_publication_ack_is_lost(self):
        self.assert_pipeline_completion(True)

    def assert_pipeline_completion(self, fail):
        storage = self.storage
        status = {'mode':'images', 'publicationId':JOB, 'options':{
            'offense_coach_card':True, 'offense_wristband':False,
            'defense_coach_card':False, 'defense_wristband':False}}
        uploads, callbacks = [], []
        storage.get_object = lambda **kw: {'Body':io.BytesIO(json.dumps(status).encode())}
        def put(**kw):
            status.clear(); status.update(json.loads(kw['Body']))
        storage.put_object = put
        storage.upload_file = lambda path, bucket, key, **kw: uploads.append(key)
        class Generator:
            def __init__(self, source, output): self.output = Path(output)
            def generate_all(self, **kwargs): (self.output/'offense_coach_card.pdf').write_bytes(b'%PDF-1.4\n%%EOF')
        def callback(s3, bucket, job):
            self.assertEqual(status['status'], 'complete')
            self.assertEqual(status['publicationId'], JOB)
            self.assertEqual(uploads, [f'jobs/{JOB}/offense_coach_card.pdf'])
            callbacks.append(job)
            if fail: raise PublicationCompletionError('unconfirmed')
            return dict(published=True, bookId=JOB, snapshotId=JOB)
        with patch.dict('os.environ', R2_BUCKET='jobs-only'), patch.object(sys, 'argv', ['process_job.py', JOB]), \
             patch.object(process_job, 'get_r2_client', return_value=storage), \
             patch.object(process_job, 'download_play_images', return_value=1), \
             patch.object(process_job, 'complete_app_publication', side_effect=callback), \
             patch.dict(sys.modules, playbook_pipeline=SimpleNamespace(PlaybookGenerator=Generator)):
            if fail:
                with self.assertRaises(SystemExit) as result: process_job.main()
                self.assertEqual(result.exception.code, 1)
            else:
                process_job.main()
        self.assertEqual(callbacks, [JOB])
        self.assertEqual(status['status'], 'complete')
        self.assertEqual(status['publication']['published'], not fail)
        self.assertEqual(status['files'], ['offense_coach_card.pdf'])
