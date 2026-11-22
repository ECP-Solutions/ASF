"""Tests of the studio's server (xll/studio, the package asfpy_studio of the add-in's Python), on
any OS: its security, the files of the Python folder, the checks, completions, the outline, the
console and the packages. Excel's calls fail cleanly here, as Excel is out of reach."""
import http.client
import importlib.util
import json
import os
import shutil
import sys
import tempfile
import textwrap
import threading
import unittest

STUDIO = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'xll', 'studio')
spec = importlib.util.spec_from_file_location('asfpy_studio', os.path.join(STUDIO, '__init__.py'),
                                              submodule_search_locations=[STUDIO])
package = importlib.util.module_from_spec(spec)
sys.modules['asfpy_studio'] = package
spec.loader.exec_module(package)
from asfpy_studio import excel, server  # noqa: E402
from unittest import mock  # noqa: E402

try:
    import jedi  # noqa: F401
except ImportError:
    jedi = None
try:
    import pyflakes  # noqa: F401
except ImportError:
    pyflakes = None

MODULE = '''
from asfpy import export, xl_func, xl_macro


@xl_func(name='PY.DOUBLE')
def double(x: float):
    """Twice x."""
    return 2 * x


@xl_macro
def tidy():
    return 1


@export
def half(x):
    return x / 2
'''


class Studio(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp(prefix='asfpy-studio-')
        cls.folder = os.path.join(cls.tmp, 'python')
        os.makedirs(cls.folder)
        with open(os.path.join(cls.folder, 'stats.py'), 'w', encoding='utf-8') as f:
            f.write(textwrap.dedent(MODULE))
        with open(os.path.join(cls.tmp, 'secret.py'), 'w') as f:
            f.write('SECRET = 1\n')
        cls.studio = server.Studio(cls.folder)
        cls.http = server.serve(cls.studio)
        cls.port = cls.http.server_address[1]
        threading.Thread(target=cls.http.serve_forever, daemon=True).start()
        cls.cookie = 'asfstudio=' + cls.studio.token

    @classmethod
    def tearDownClass(cls):
        cls.http.shutdown()
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def request(self, method, path, body=None, cookie=True, host=None, studio_header=True, origin=None):
        c = http.client.HTTPConnection('127.0.0.1', self.port, timeout=60)
        headers = {'Host': host or '127.0.0.1:%d' % self.port}
        if cookie:
            headers['Cookie'] = self.cookie
        if body is not None:
            headers['Content-Type'] = 'application/json'
            if studio_header:
                headers['X-Studio'] = '1'
        if origin:
            headers['Origin'] = origin
        c.request(method, path, json.dumps(body) if body is not None else None, headers)
        r = c.getresponse()
        data = r.read()
        c.close()
        try:
            return r.status, json.loads(data), r
        except ValueError:
            return r.status, data.decode('utf-8', 'replace'), r

    def get(self, path, **kw):
        return self.request('GET', path, **kw)

    def post(self, name, body, **kw):
        return self.request('POST', '/api/' + name, body, **kw)

    # --- security
    def test_the_token_opens_the_page_with_a_cookie(self):
        status, _, r = self.get('/?t=' + self.studio.token, cookie=False)
        self.assertEqual(status, 302)
        self.assertIn('asfstudio=%s; HttpOnly; SameSite=Strict' % self.studio.token, r.getheader('Set-Cookie'))
        self.assertEqual(self.get('/?t=wrong', cookie=False)[0], 403)

    def test_without_the_cookie_nothing(self):
        self.assertEqual(self.get('/api/files', cookie=False)[0], 403)
        self.assertEqual(self.get('/', cookie=False)[0], 403)
        self.assertEqual(self.post('save', {'path': 'x.py', 'text': ''}, cookie=False)[0], 403)

    def test_another_host_or_origin_or_no_studio_header_is_refused(self):
        self.assertEqual(self.get('/api/files', host='evil.example:%d' % self.port)[0], 403)
        self.assertEqual(self.post('lint', {'text': ''}, studio_header=False)[0], 403)
        self.assertEqual(self.post('lint', {'text': ''}, origin='http://evil.example')[0], 403)
        self.assertEqual(self.post('lint', {'text': ''}, origin='http://127.0.0.1:%d' % self.port)[0], 200)

    def test_only_py_files_of_the_folder(self):
        self.assertEqual(self.get('/api/file?path=../secret.py')[0], 400)
        self.assertEqual(self.get('/api/file?path=' + os.path.join(self.tmp, 'secret.py'))[0], 400)
        self.assertEqual(self.post('save', {'path': 'notes.txt', 'text': 'x'})[0], 400)
        self.assertEqual(self.get('/../server.py')[0], 404)

    def test_the_page_and_its_files(self):
        status, html, r = self.get('/')
        self.assertEqual(status, 200)
        self.assertIn('Python Studio', html)
        self.assertIn("default-src 'self'", r.getheader('Content-Security-Policy'))
        self.assertEqual(self.get('/app.js')[0], 200)

    # --- files, checks, outline
    def test_files_read_save_and_new(self):
        status, files, _ = self.get('/api/files')
        self.assertIn('stats.py', files)
        self.assertIn('def double', self.get('/api/file?path=stats.py')[1]['text'])
        status, r, _ = self.post('save', {'path': 'sub/extra.py', 'text': 'VALUE = 1\n'})
        self.assertEqual((status, r['problems']), (200, []))
        self.assertIn('sub/extra.py', self.get('/api/files')[1])
        status, r, _ = self.post('new', {'path': 'fresh'})
        self.assertEqual(r['path'], 'fresh.py')
        self.assertIn('@xl_func', self.get('/api/file?path=fresh.py')[1]['text'])
        self.assertEqual(self.post('new', {'path': 'fresh.py'})[0], 400)

    def test_syntax_errors_and_warnings(self):
        r = self.post('lint', {'text': 'def f(:\n    pass\n'})[1]
        self.assertEqual((r[0]['line'], r[0]['severity']), (1, 'error'))
        if pyflakes:
            r = self.post('lint', {'text': 'import os\nprint(undefined_name)\n'})[1]
            self.assertEqual([p['line'] for p in r], [1, 2])
            self.assertIn("undefined name 'undefined_name'", r[1]['message'])

    def test_the_addin_page_and_its_settings(self):
        status, a, _ = self.get('/api/addin')
        self.assertEqual((status, a['PythonUpdates'], a['PythonTips']), (200, 'check', 'on'))
        self.assertTrue(a['log'].endswith('asfpy.log') and a['updateLog'].endswith('asfpy-update.log'))
        with mock.patch.dict(server._no_registry, clear=True):
            self.assertEqual(self.post('setting', {'name': 'PythonUpdates', 'value': 'Auto'})[1], {'PythonUpdates': 'auto'})
            self.assertEqual(self.post('setting', {'name': 'PythonTips', 'value': 'off'})[0], 200)
            a = self.get('/api/addin')[1]
            self.assertEqual((a['PythonUpdates'], a['PythonTips']), ('auto', 'off'))
            self.assertEqual(self.post('setting', {'name': 'PythonUpdates', 'value': 'sometimes'})[0], 400)
            self.assertEqual(self.post('setting', {'name': 'PythonVenv', 'value': 'C:\\x'})[0], 400)
        if os.name != 'nt':
            self.assertEqual(self.post('open', {'what': 'log'})[0], 400)

    def test_the_token_can_open_a_page_of_the_studio(self):
        _, _, r = self.get('/?t=%s&view=addin' % self.studio.token, cookie=False)
        self.assertEqual(r.getheader('Location'), '/#addin')
        _, _, r = self.get('/?t=%s&view=javascript:x' % self.studio.token, cookie=False)
        self.assertEqual(r.getheader('Location'), '/')

    def test_explorer_is_for_windows(self):
        if os.name == 'nt':
            self.skipTest('it would open a window')
        status, r, _ = self.post('reveal', {})
        self.assertEqual(status, 400)
        self.assertIn('Windows', r['error'])

    def test_the_outline_reads_the_decorators(self):
        items = {(i['kind'], i['name']) for i in self.get('/api/outline')[1]}
        self.assertTrue({('xl_func', 'PY.DOUBLE'), ('xl_macro', 'tidy'), ('export', 'stats.half')} <= items)

    @unittest.skipIf(jedi is None, 'jedi is not installed')
    def test_completions_and_signatures(self):
        text = 'import stats\nstats.dou'
        r = self.post('complete', {'text': text, 'line': 2, 'column': 9, 'path': 'other.py'})[1]
        self.assertIn('double', [c['text'] for c in r['completions']])
        text = 'from asfpy import xl_func\nxl_func('
        r = self.post('complete', {'text': text, 'line': 2, 'column': 8})[1]
        self.assertTrue(r['signatures'] and 'asynchronous' in r['signatures'][0]['label'])

    # --- console
    def test_the_console_keeps_its_names_and_imports_the_folder(self):
        self.assertEqual(self.post('console', {'code': 'x = 21'})[1]['output'], '')
        self.assertEqual(self.post('console', {'code': 'x * 2'})[1]['output'], '42\n')
        self.assertEqual(self.post('console', {'code': 'import stats\nstats.double(4)'})[1]['output'], '8\n')
        out = self.post('console', {'code': '1 / 0'})[1]['output']
        self.assertIn('ZeroDivisionError', out)
        self.assertNotIn('server.py', out)
        self.post('console', {'restart': True, 'code': ''})
        self.assertIn('NameError', self.post('console', {'code': 'x'})[1]['output'])

    # --- packages and Excel
    def test_packages(self):
        names = [p['name'].lower() for p in self.get('/api/packages')[1]]
        self.assertIn('pip', names)
        self.assertEqual(self.post('pip', {'command': ''})[0], 400)

    def test_a_macro_runs_with_its_arguments_and_says_what_came_of_it(self):
        class App:
            Hwnd = 0
            calls = []

            def Run(self, name, *args):
                self.calls.append((name, args))
                if name == 'check_macro':
                    return float(sum(args))
                if name == 'table':
                    return ((1.0, 'a'), (2.5, None))
                if name == 'AsfPyLastError':
                    return 'Traceback (most recent call last):\nTypeError: missing b'
                return -2146826273                     # #VALUE!: the macro failed
        app = App()
        with mock.patch.object(excel, 'call', lambda fn, timeout=600: fn(app)):
            status, r, _ = self.post('excel/macro', {'name': 'check_macro', 'args': [2, 3, None], 'front': True})
            self.assertEqual((status, r), (200, {'result': '5', 'error': False, 'front': False}))   # no Windows here
            self.assertEqual(app.calls[-1], ('check_macro', (2, 3)))     # a trailing blank: the default
            self.assertEqual(self.post('excel/macro', {'name': 'table', 'args': []})[1]['result'], '1\ta\n2.5\t')
            r = self.post('excel/macro', {'name': 'broken', 'args': [None, 1]})[1]
            self.assertTrue(r['error'])
            self.assertTrue(r['result'].startswith('#VALUE!\nTraceback'))
            self.assertEqual(app.calls[-2], ('broken', (None, 1)))     # a blank between: Empty
        self.assertEqual(self.post('excel/macro', {'name': 'x', 'args': [{'a': 1}]})[0], 400)

    def test_a_macro_run_with_excels_own_dialogs(self):
        answers = {'check_macro': '5', 'cancelled': '', 'broken': 'Traceback (most recent call last):\nTypeError: x',
                   'unknown': -2146826259}

        class App:
            Hwnd = 0
            calls = []

            def Run(self, name, *args):
                self.calls.append((name, args))
                return answers[args[0]]
        app = App()
        with mock.patch.object(excel, 'call', lambda fn, timeout=600: fn(app)):
            r = self.post('excel/macro-in-excel', {'name': 'check_macro'})[1]
            self.assertEqual(r, {'result': '5', 'error': False, 'cancelled': False, 'front': False})
            self.assertEqual(app.calls[-1], ('AsfPyRunMacro', ('check_macro',)))
            self.assertTrue(self.post('excel/macro-in-excel', {'name': 'cancelled'})[1]['cancelled'])
            self.assertTrue(self.post('excel/macro-in-excel', {'name': 'broken'})[1]['error'])
            r = self.post('excel/macro-in-excel', {'name': 'unknown'})[1]
            self.assertEqual((r['error'], r['result'][:6]), (True, '#NAME?'))

    def test_excel_out_of_reach_is_an_answer_not_a_crash(self):
        if os.name == 'nt':
            self.skipTest('Excel may run here')
        status, r, _ = self.post('excel/reload', {})
        self.assertEqual(status, 503)
        self.assertIn('Excel', r['error'])
        self.assertIsNone(self.get('/api/info')[1]['excel'])


if __name__ == '__main__':
    unittest.main()
