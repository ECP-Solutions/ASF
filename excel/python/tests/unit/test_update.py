"""Tests of the add-in's updater (xll/studio/update.py) and of the studio's calls for it, on any OS,
against a stand-in for GitHub's releases API on 127.0.0.1: what a release offers, the checks a
download must pass before it replaces the add-in, the swap, the state shared by every Excel,
and the Pythons an update leaves behind."""
import hashlib
import http.client
import http.server
import importlib.util
import json
import os
import re
import shutil
import struct
import sys
import tempfile
import threading
import time
import unittest
from unittest import mock

XLL = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'xll')
STUDIO = os.path.join(XLL, 'studio')
spec = importlib.util.spec_from_file_location('asfpy_studio', os.path.join(STUDIO, '__init__.py'),
                                              submodule_search_locations=[STUDIO])
package = importlib.util.module_from_spec(spec)
sys.modules.setdefault('asfpy_studio', package)
if sys.modules['asfpy_studio'] is package:
    spec.loader.exec_module(package)
from asfpy_studio import server, update  # noqa: E402


def addin(version, bits=64, size=5000):
    """An add-in file as far as the updater looks: a DLL of a machine, with a version mark."""
    head = bytearray(512)
    head[:2] = b'MZ'
    struct.pack_into('<I', head, 0x3C, 0x80)
    head[0x80:0x84] = b'PE\0\0'
    struct.pack_into('<H', head, 0x84, 0x8664 if bits == 64 else 0x14C)
    struct.pack_into('<H', head, 0x80 + 22, 0x2022)
    body = bytes(head) + b'\0' * size + (b'ASFPY-VERSION:%s;' % version.encode() if version else b'') + b'\0' * 64
    return body


class Releases(http.server.BaseHTTPRequestHandler):
    """GET /latest: the release; GET /files/NAME: its assets. The test sets release and files."""
    release, files = {}, {}

    def log_message(self, *args):
        pass

    def do_GET(self):
        if self.path == '/latest':
            body, kind = json.dumps(self.release).encode(), 'application/json'
        elif self.path.startswith('/files/') and self.path[7:] in self.files:
            body, kind = self.files[self.path[7:]], 'application/octet-stream'
        else:
            self.send_response(404)
            self.end_headers()
            return
        self.send_response(200)
        self.send_header('Content-Type', kind)
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)


class UpdateTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.http = http.server.ThreadingHTTPServer(('127.0.0.1', 0), Releases)
        cls.base = 'http://127.0.0.1:%d' % cls.http.server_address[1]
        threading.Thread(target=cls.http.serve_forever, daemon=True).start()

    @classmethod
    def tearDownClass(cls):
        cls.http.shutdown()

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix='asfpy-update-')
        self.env = mock.patch.dict(os.environ, {'LOCALAPPDATA': os.path.join(self.tmp, 'local')})
        self.env.start()
        self.folder = os.path.join(self.tmp, 'addins')
        os.makedirs(self.folder)
        self.xll = os.path.join(self.folder, 'asfpy64.xll')
        with open(self.xll, 'wb') as f:
            f.write(addin('4.1.1'))
        self.publish('4.2.0')
        self.source = self.base + '/latest'

    def tearDown(self):
        self.env.stop()
        shutil.rmtree(self.tmp, ignore_errors=True)

    def publish(self, version, data=None, digest=True, sums=False, names=('asfpy64.xll', 'asfpy32.xll')):
        Releases.files = {}
        assets = []
        for name in names:
            body = data if data is not None else addin(version, 64 if '64' in name else 32)
            Releases.files[name] = body
            asset = {'name': name, 'size': len(body), 'browser_download_url': self.base + '/files/' + name}
            if digest:
                asset['digest'] = 'sha256:' + hashlib.sha256(body).hexdigest()
            assets.append(asset)
        if sums:
            text = ''.join('%s  %s\n' % (hashlib.sha256(Releases.files[n]).hexdigest(), n) for n in names)
            Releases.files['SHA256SUMS'] = text.encode()
            assets.append({'name': 'SHA256SUMS', 'size': len(text), 'browser_download_url': self.base + '/files/SHA256SUMS'})
        Releases.release = {'tag_name': 'v' + version, 'html_url': 'https://example.invalid/v' + version, 'assets': assets}

    def left(self):
        return sorted(os.listdir(self.folder))


class Versions(UpdateTest):
    def test_versions_compare_as_numbers(self):
        self.assertTrue(update.newer('v4.10.0', '4.9.3'))
        self.assertTrue(update.newer('4.1.1', '4.1'))
        self.assertFalse(update.newer('4.1.0', '4.1'))
        self.assertFalse(update.newer('', '4.1.1'))
        self.assertTrue(update.newer('4.1.1', ''))

    def test_the_file_says_its_version_and_bitness(self):
        self.assertEqual((update.file_version(self.xll), update.dll_bits(self.xll)), ('4.1.1', 64))
        path = os.path.join(self.tmp, 'old.xll')
        with open(path, 'wb') as f:
            f.write(addin(None, 32))
        self.assertEqual((update.file_version(path), update.dll_bits(path)), (None, 32))
        with open(path, 'wb') as f:
            f.write(b'not a DLL' * 100)
        self.assertIsNone(update.dll_bits(path))

    def test_the_built_addins_carry_asf_version(self):
        with open(os.path.join(XLL, '..', '..', '..', 'src', 'ASF.cls'), encoding='latin-1') as f:
            want = re.search(r'ASF_VERSION As String = "([\d.]+)"', f.read()).group(1)
        for bits in (64, 32):
            path = os.path.join(XLL, 'build', 'asfpy%d.xll' % bits)
            if not os.path.isfile(path):
                self.skipTest('build.py has not run')
            self.assertEqual((update.file_version(path), update.dll_bits(path)), (want, bits))


class Check(UpdateTest):
    def test_a_newer_release_is_available(self):
        status = update.check(self.xll, '4.1.1', self.source)
        self.assertEqual((status['installed'], status['latest'], status['available'], status['restart']),
                         ('4.1.1', '4.2.0', True, False))
        self.assertEqual(status['release']['url'], self.base + '/files/asfpy64.xll')

    def test_the_same_release_is_not(self):
        self.publish('4.1.1')
        self.assertFalse(update.check(self.xll, '4.1.1', self.source)['available'])

    def test_an_installed_update_waits_for_a_restart(self):
        status = update.check(self.xll, '4.0.0', self.source)
        self.assertTrue(status['restart'])

    def test_the_32_bit_addin_gets_the_32_bit_asset(self):
        with open(self.xll, 'wb') as f:
            f.write(addin('4.1.1', 32))
        self.assertEqual(update.check(self.xll, '4.1.1', self.source)['release']['name'], 'asfpy32.xll')

    def test_a_release_without_the_addin_or_its_digest_offers_nothing(self):
        self.publish('4.2.0', names=('ASF.xlam',))
        with self.assertRaisesRegex(update.UpdateError, 'has no asfpy64.xll'):
            update.check(self.xll, '4.1.1', self.source)
        self.publish('4.2.0', digest=False)
        with self.assertRaisesRegex(update.UpdateError, 'no SHA-256'):
            update.check(self.xll, '4.1.1', self.source)

    def test_sha256sums_stands_in_for_the_digest(self):
        self.publish('4.2.0', digest=False, sums=True)
        info = update.check(self.xll, '4.1.1', self.source)['release']
        self.assertEqual(info['sha256'], hashlib.sha256(Releases.files['asfpy64.xll']).hexdigest())

    def test_only_https_or_this_computer(self):
        with self.assertRaisesRegex(update.UpdateError, 'https only'):
            update.latest('http://example.com/latest')

    def test_no_network_is_an_answer(self):
        with self.assertRaisesRegex(update.UpdateError, "can't be read"):
            update.latest(self.base + '/nothing-here')


class Install(UpdateTest):
    def info(self):
        return update.check(self.xll, '4.1.1', self.source)['release']

    def test_the_new_addin_takes_the_place_of_the_old(self):
        result = update.install(self.xll, self.info())
        self.assertEqual(result, {'installed': self.xll, 'version': '4.2.0'})
        self.assertEqual(update.file_version(self.xll), '4.2.0')
        self.assertEqual(update.file_version(self.xll + '.old'), '4.1.1')
        self.assertEqual(self.left(), ['asfpy64.xll', 'asfpy64.xll.old'])

    def test_an_old_file_in_use_gets_a_name_of_its_own(self):
        with open(self.xll + '.old', 'wb') as f:
            f.write(b'in use')
        real_remove = os.remove

        def busy(path):
            if path.endswith('.old'):
                raise PermissionError(path)
            real_remove(path)
        with mock.patch.object(update.os, 'remove', busy):
            update.install(self.xll, self.info())
        self.assertEqual(update.file_version(self.xll), '4.2.0')
        self.assertTrue(any(re.fullmatch(r'asfpy64\.xll\.\d+\.old', n) for n in self.left()))

    def test_a_download_that_is_not_the_release_changes_nothing(self):
        info = self.info()
        for data, message in ((addin('4.2.0')[:-10], 'bytes of'), (addin('4.2.1'), 'SHA-256')):
            Releases.files['asfpy64.xll'] = data
            with self.assertRaisesRegex(update.UpdateError, message):
                update.install(self.xll, info)
            self.assertEqual((update.file_version(self.xll), self.left()), ('4.1.1', ['asfpy64.xll']))

    def test_the_wrong_bitness_or_version_is_refused(self):
        for data, message in ((addin('4.2.0', 32), 'not a 64-bit'), (addin('4.3.0'), 'says it is version 4.3.0'),
                              (b'x' * 600, 'not a 64-bit')):
            self.publish('4.2.0', data=data)
            with self.assertRaisesRegex(update.UpdateError, message):
                update.install(self.xll, self.info())
            self.assertEqual(self.left(), ['asfpy64.xll'])

    def test_a_folder_that_cannot_be_written_keeps_the_download_elsewhere(self):
        with mock.patch.object(update, 'swap', side_effect=PermissionError('read-only')):
            result = update.install(self.xll, self.info())
        self.assertEqual(update.file_version(result['saved']), '4.2.0')
        self.assertTrue(result['saved'].startswith(os.path.join(self.tmp, 'local', 'ASF', 'updates')))
        self.assertEqual(self.left(), ['asfpy64.xll'])
        self.assertIn("can't be written to", update.outcome(result, self.xll))


class OnStart(UpdateTest):
    def test_auto_installs_and_says_so(self):
        with mock.patch.object(update, 'tell', return_value=False) as tell:
            self.assertEqual(update.on_start(self.xll, '4.1.1', 'auto', self.source), 'installed 4.2.0')
        self.assertIn('Restart Excel', tell.call_args[0][0])
        self.assertEqual(update.file_version(self.xll), '4.2.0')

    def test_at_most_a_check_a_day(self):
        with mock.patch.object(update, 'tell', return_value=False):
            update.on_start(self.xll, '4.1.1', 'check', self.source)
            self.assertIn('hours ago', update.on_start(self.xll, '4.1.1', 'check', self.source))
            later = time.time() + update.INTERVAL + 60
            self.assertEqual(update.on_start(self.xll, '4.1.1', 'check', self.source, now=later), '4.2.0 was declined')

    def test_check_asks_and_no_waits_a_week(self):
        with mock.patch.object(update, 'tell', return_value=False) as tell:
            self.assertEqual(update.on_start(self.xll, '4.1.1', 'check', self.source), '4.2.0 declined')
        self.assertTrue(tell.call_args[1]['question'])
        self.assertIn('ASF Python 4.2.0 is available; this add-in is 4.1.1', tell.call_args[0][0])
        self.assertEqual(update.file_version(self.xll), '4.1.1')
        week = time.time() + update.REMIND + 60
        with mock.patch.object(update, 'tell', side_effect=[True, False]):
            self.assertEqual(update.on_start(self.xll, '4.1.1', 'check', self.source, now=week), 'installed 4.2.0')

    def test_up_to_date_says_nothing(self):
        self.publish('4.1.1')
        with mock.patch.object(update, 'tell') as tell:
            self.assertEqual(update.on_start(self.xll, '4.1.1', 'check', self.source), '4.1.1 is the latest')
        tell.assert_not_called()

    def test_one_updater_at_a_time(self):
        with update.Lock() as first, update.Lock() as second:
            self.assertEqual((first.held, second.held), (True, False))
            self.assertEqual(update.on_start(self.xll, '4.1.1', 'auto', self.source), 'another updater runs')
        with update.Lock() as again:
            self.assertTrue(again.held)

    def test_ask_now_says_what_came_of_it(self):
        self.publish('4.1.1')
        with mock.patch.object(update, 'tell') as tell:
            update.ask_now(self.xll, '4.1.1', self.source)
        self.assertEqual(tell.call_args[0][0], 'ASF Python 4.1.1 is the latest version.')
        with mock.patch.object(update, 'tell') as tell:
            update.ask_now(self.xll, '4.1.1', self.base + '/nothing-here')
        self.assertIn('could not check for updates', tell.call_args[0][0])


class Studio(UpdateTest):
    def setUp(self):
        super().setUp()
        self.studio = server.Studio('', xll=self.xll, running='4.1.1', update_source=self.source)
        self.http2 = server.serve(self.studio)
        self.port = self.http2.server_address[1]
        threading.Thread(target=self.http2.serve_forever, daemon=True).start()

    def tearDown(self):
        self.http2.shutdown()
        super().tearDown()

    def call(self, method, name, body=None):
        c = http.client.HTTPConnection('127.0.0.1', self.port, timeout=60)
        headers = {'Host': '127.0.0.1:%d' % self.port, 'Cookie': 'asfstudio=' + self.studio.token}
        if body is not None:
            headers.update({'Content-Type': 'application/json', 'X-Studio': '1'})
        c.request(method, '/api/' + name, json.dumps(body) if body is not None else None, headers)
        r = c.getresponse()
        data = json.loads(r.read())
        c.close()
        return r.status, data

    def test_the_last_check_then_a_live_one(self):
        status, s = self.call('GET', 'update')
        self.assertEqual((status, s['installed'], s['latest'], s['available']), (200, '4.1.1', '', False))
        status, s = self.call('GET', 'update?live=1')
        self.assertEqual((s['latest'], s['available']), ('4.2.0', True))
        self.assertEqual(self.call('GET', 'update')[1]['latest'], '4.2.0')

    def test_install_as_a_job(self):
        status, job = self.call('POST', 'update', {})
        self.assertEqual((status, job['kind']), (200, 'update'))
        for _ in range(100):
            jobs = self.call('GET', 'jobs')[1]['jobs']
            if jobs[-1]['state'] not in ('queued', 'running'):
                break
            time.sleep(0.05)
        self.assertEqual(jobs[-1]['state'], 'done', jobs[-1]['output'])
        self.assertIn('ASF Python 4.2.0 is installed', jobs[-1]['output'])
        s = self.call('GET', 'update')[1]
        self.assertEqual((s['installed'], s['restart']), ('4.2.0', True))

    def test_no_releases_is_an_answer(self):
        self.studio.update_source = self.base + '/nothing-here'
        status, r = self.call('GET', 'update?live=1')
        self.assertEqual(status, 502)
        self.assertIn("can't be read", r['error'])

    def test_without_the_addin_nothing_to_update(self):
        self.studio.xll = ''
        self.assertEqual(self.call('GET', 'update')[0], 400)


class Previous(unittest.TestCase):
    """The Pythons of earlier add-ins beside this one, with the packages added to them."""
    def setUp(self):
        self.root = tempfile.mkdtemp(prefix='asfpy-runtime-')

    def tearDown(self):
        shutil.rmtree(self.root, ignore_errors=True)

    def python(self, name, packages):
        site = os.path.join(self.root, name, 'Lib', 'site-packages')
        os.makedirs(site)
        for p, v in packages:
            info = os.path.join(site, '%s-%s.dist-info' % (p.replace('-', '_'), v))
            os.makedirs(info)
            with open(os.path.join(info, 'METADATA'), 'w') as f:
                f.write('Metadata-Version: 2.1\nName: %s\nVersion: %s\n' % (p, v))
        return os.path.join(self.root, name)

    def test_what_the_previous_python_has_that_this_one_lacks(self):
        now = self.python('python-3.14.11-64-bbbbbbbbbbbb', [('numpy', '2.6.0'), ('pip', '26.0')])
        self.python('python-3.14.10-64-aaaaaaaaaaaa', [('numpy', '2.5.5'), ('SciPy', '1.16.0'), ('pip', '25.3'),
                                                       ('statsmodels', '0.15')])
        self.python('python-3.14.10-32-cccccccccccc', [('requests', '2.33')])
        os.makedirs(os.path.join(self.root, 'python-3.14.9-64-dddddddddddd-removing'))
        found = server.previous(now)
        self.assertEqual([p['folder'] for p in found], ['python-3.14.10-64-aaaaaaaaaaaa'])
        self.assertEqual([p['name'] for p in found[0]['packages']], ['SciPy', 'statsmodels'])
        self.assertFalse(os.path.exists(os.path.join(self.root, 'python-3.14.9-64-dddddddddddd-removing')))

        studio = server.Studio('')
        studio.prefix = now
        with mock.patch.object(studio, 'pip') as pip:
            studio.previous_install('python-3.14.10-64-aaaaaaaaaaaa')
        pip.assert_called_once_with('install SciPy statsmodels')
        self.assertEqual(studio.previous_remove('python-3.14.10-64-aaaaaaaaaaaa'),
                         {'removed': 'python-3.14.10-64-aaaaaaaaaaaa', 'left': False})
        self.assertEqual(server.previous(now), [])
        with self.assertRaises(LookupError):
            studio.previous_remove('..')

    def test_a_python_not_of_the_addin_has_none(self):
        self.assertEqual(server.previous(sys.prefix if not sys.prefix.endswith(os.sep) else sys.prefix[:-1]), [])


if __name__ == '__main__':
    unittest.main()
