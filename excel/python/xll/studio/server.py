"""The studio's server: the files of the Python folder, completions and checks, a console, the
packages of this Python, and Excel's commands, as a JSON API for the page in static/.

It listens on 127.0.0.1 only. Every request must come for that host and port, with the session's
cookie, set when the window opens the page with the session's token; a POST must also carry the
X-Studio header, which another site's page can't send. The server stops when no window has
called for IDLE seconds and no pip command or update runs.

    python -m asfpy_studio --folder PYTHON_FOLDER [--xll ADDIN --running VERSION] [--update-source URL]
                           [--no-window] [--port N] [--log FILE]
"""
import argparse
import ast
import http.server
import importlib.metadata
import importlib.util
import io
import json
import os
import queue
import re
import secrets
import shlex
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import urllib.parse
import urllib.request

from . import excel, update

HERE = os.path.dirname(os.path.abspath(__file__))
STATIC = os.path.join(HERE, 'static')
STUBS = os.path.join(HERE, 'stubs')
IDLE = 120
NAME = re.compile(r'^[A-Za-z0-9][A-Za-z0-9._-]*$')
VIEWS = ('files', 'functions', 'packages', 'addin')
SETTINGS_KEY = r'Software\VB and VBA Program Settings\ASF\AddIn'      # where the add-in reads them
SETTINGS = ('PythonUpdates', 'PythonTips', 'PythonUpdateSource', 'PythonVenv', 'PythonFolder')
EDITABLE = {'PythonUpdates': ('check', 'auto', 'off'), 'PythonTips': ('on', 'off')}
_no_registry = {}                    # the settings where there is no registry: the tests
RUNTIME = re.compile(r'^python-[0-9.]+-(?P<bits>32|64)-[0-9a-f]{6,}$')   # a Python of the add-in, unpacked
TYPES = {'.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
         '.css': 'text/css; charset=utf-8', '.svg': 'image/svg+xml', '.png': 'image/png'}
TEMPLATE = '''"""{module}: functions for Excel."""
from asfpy import export, xl_func, xl_macro


@xl_func
def {function}(x: float, factor: float = 2):
    """What the function gives: Insert Function and the tooltip in Excel show this line.

    Args:
        x: what x is; Excel shows this while the caret is in x
        factor: how much to multiply x by; 2 by default
    """
    return x * factor
'''


def _local_app_data():
    return os.environ.get('LOCALAPPDATA') or os.path.join(os.path.expanduser('~'), '.cache')


class Studio:
    def __init__(self, folder, idle=IDLE, xll='', running='', update_source=None):
        self.folder = os.path.abspath(folder) if folder else ''
        self.xll, self.running, self.update_source = xll, running, update_source or None
        self.prefix = sys.prefix                # the Python whose earlier versions previous() lists
        self.token = secrets.token_urlsafe(32)
        self.idle = idle
        self.seen = time.time()
        self.lock = threading.Lock()            # the jobs
        self.console_lock = threading.Lock()
        self.jobs = []
        self.pending = queue.Queue()
        self.revision = 0
        self.namespace = None
        threading.Thread(target=self._pip_worker, name='pip', daemon=True).start()

    # --- files of the Python folder ----------------------------------------------------------

    def path(self, rel):
        """The full path of a .py file of the Python folder, or PermissionError."""
        if not self.folder:
            raise PermissionError('no Python folder')
        root = os.path.realpath(self.folder)
        full = os.path.realpath(os.path.join(root, str(rel)))
        if os.path.commonpath([full, root]) != root or not full.lower().endswith('.py'):
            raise PermissionError('only .py files of the Python folder')
        return full

    def files(self):
        out = []
        if self.folder and os.path.isdir(self.folder):
            for root, dirs, names in os.walk(self.folder):
                dirs[:] = sorted(d for d in dirs if not d.startswith(('.', '__')))
                for n in sorted(names):
                    if n.lower().endswith('.py'):
                        out.append(os.path.relpath(os.path.join(root, n), self.folder).replace(os.sep, '/'))
        return out

    def read(self, rel):
        with open(self.path(rel), encoding='utf-8') as f:
            return f.read()

    def save(self, rel, text):
        full = self.path(rel)
        os.makedirs(os.path.dirname(full), exist_ok=True)
        with open(full, 'w', encoding='utf-8', newline='') as f:
            f.write(text)
        return {'path': rel, 'problems': lint(text, full)}

    def new(self, rel):
        rel = rel if rel.lower().endswith('.py') else rel + '.py'
        full = self.path(rel)
        if os.path.exists(full):
            raise FileExistsError('%s is there already' % rel)
        stem = os.path.splitext(os.path.basename(rel))[0]
        function = re.sub(r'\W', '_', stem).lower() or 'example'
        return self.save(rel, TEMPLATE.format(module=stem, function='%s_example' % function))

    def reveal(self):
        """The Python folder in Windows Explorer, made first when it is missing."""
        if not self.folder:
            raise PermissionError('no Python folder')
        if os.name != 'nt':
            raise OSError('Explorer is on Windows only')
        os.makedirs(self.folder, exist_ok=True)
        os.startfile(self.folder)
        return {'folder': self.folder}

    def outline(self):
        """The functions of the Python folder that Excel gets: worksheet functions, macros and
        exported ones, read from the source, never imported."""
        out = []
        for rel in self.files():
            try:
                tree = ast.parse(self.read(rel))
            except (SyntaxError, ValueError, OSError):
                continue
            stem = rel[:-3].replace('/', '.')
            for node in tree.body:
                if not isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                    continue
                for d in node.decorator_list:
                    target = d.func if isinstance(d, ast.Call) else d
                    kind = getattr(target, 'id', None) or getattr(target, 'attr', None)
                    if kind not in ('xl_func', 'xl_macro', 'export'):
                        continue
                    name = node.name
                    if isinstance(d, ast.Call):
                        for k in d.keywords:
                            if k.arg == 'name' and isinstance(k.value, ast.Constant):
                                name = k.value.value
                    out.append({'kind': kind, 'name': name if kind != 'export' else '%s.%s' % (stem, node.name),
                                'python': node.name, 'file': rel, 'line': node.lineno,
                                'args': [a.arg for a in node.args.args],
                                'doc': (ast.get_docstring(node) or '').strip().split('\n')[0]})
        return out

    # --- the console: Python apart from Excel, the Python folder importable ---------------------

    def console(self, source, restart=False):
        with self.console_lock:
            if restart or self.namespace is None:
                self.namespace = {'__name__': '__console__'}
                _stand_in_asfpy()
                if self.folder and self.folder not in sys.path:
                    sys.path.append(self.folder)
                if restart:
                    return {'output': 'A new session.\n'}
            out = io.StringIO()
            saved = sys.stdout, sys.stderr
            sys.stdout = sys.stderr = out
            try:
                tree = ast.parse(source, '<console>', 'exec')
                last = tree.body.pop() if tree.body and isinstance(tree.body[-1], ast.Expr) else None
                exec(compile(tree, '<console>', 'exec'), self.namespace)
                if last is not None:
                    value = eval(compile(ast.Expression(last.value), '<console>', 'eval'), self.namespace)
                    if value is not None:
                        print(repr(value))
                        self.namespace['_'] = value
            except SystemExit:
                print('exit() closes nothing here: close the window instead')
            except BaseException:
                tb = traceback.format_exc().split('\n')
                print('\n'.join([tb[0]] + [l for l in tb[3:] if l]))   # without the studio's own frame
            finally:
                sys.stdout, sys.stderr = saved
            return {'output': out.getvalue()}

    # --- packages ----------------------------------------------------------------------------

    def packages(self):
        seen, out = set(), []
        for d in importlib.metadata.distributions():
            name = d.metadata.get('Name') or ''
            if name.lower() in seen or not name:
                continue
            seen.add(name.lower())
            out.append({'name': name, 'version': d.version, 'summary': d.metadata.get('Summary') or ''})
        return sorted(out, key=lambda p: p['name'].lower())

    def pip(self, command):
        args = shlex.split(command.replace('\\', '\\\\'))
        if args and args[0].lower() == 'pip':
            args = args[1:]
        if not args:
            raise ValueError('a pip command, such as: install numpy')
        with self.lock:
            job = {'id': len(self.jobs) + 1, 'command': ' '.join(args), 'state': 'queued', 'code': None, 'output': ''}
            self.jobs.append(job)
            self.revision += 1
        self.pending.put((job, args))
        return job

    def _pip_worker(self):
        while True:
            job, args = self.pending.get()
            self._update(job, state='running')
            flags = 0x08000000 if os.name == 'nt' else 0          # CREATE_NO_WINDOW
            try:
                p = subprocess.Popen([_python(), '-I', '-m', 'pip', '--disable-pip-version-check', '--no-input'] + args,
                                     stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding='utf-8',
                                     errors='replace', creationflags=flags)
                for line in p.stdout:
                    self._update(job, output=line)
                code = p.wait()
            except OSError as exc:
                self._update(job, output='%s\n' % exc)
                code = -1
            self._update(job, state='done' if code == 0 else 'failed', code=code)

    def _update(self, job, state=None, code=None, output=None):
        with self.lock:
            if state:
                job['state'] = state
            if code is not None:
                job['code'] = code
            if output:
                job['output'] = (job['output'] + output)[-200000:]
            self.revision += 1

    def busy(self):
        return any(j['state'] in ('queued', 'running') for j in self.jobs)

    # --- updates of the add-in, and the Python an update left -----------------------------------

    def update_status(self, live=False):
        """The add-in's version and the latest release: from GitHub when live, else as the last
        check (Excel's start, or this) left it."""
        if not self.xll:
            raise ValueError('the studio was opened without the add-in to update')
        installed = update.file_version(self.xll) or self.running
        state = update.load_state()
        if live:
            status = update.check(self.xll, self.running, self.update_source)
            del status['release']
            state.update(checked=time.time(), latest=status['latest'], page=status['page'], size=status['size'])
            update.save_state(state)
        else:
            latest = state.get('latest') or ''
            status = {'running': self.running or installed, 'installed': installed, 'latest': latest,
                      'page': state.get('page') or update.RELEASES, 'available': update.newer(latest, installed),
                      'restart': bool(self.running) and update.newer(installed, self.running), 'size': state.get('size', 0)}
        status['checked'] = state.get('checked')
        return status

    def update(self):
        """The latest release installed in the add-in's place, as a job of the Activity list."""
        if not self.xll:
            raise ValueError('the studio was opened without the add-in to update')
        with self.lock:
            if any(j.get('kind') == 'update' and j['state'] in ('queued', 'running') for j in self.jobs):
                raise ValueError('the add-in is being updated')
            job = {'id': len(self.jobs) + 1, 'kind': 'update', 'command': 'update ASF Python', 'state': 'queued',
                   'code': None, 'output': ''}
            self.jobs.append(job)
            self.revision += 1
        threading.Thread(target=self._update_worker, args=(job,), name='update', daemon=True).start()
        return job

    def _update_worker(self, job):
        self._update(job, state='running')
        code = 1
        try:
            with update.Lock() as lock:
                if not lock.held:
                    raise update.UpdateError('another Excel is updating the add-in')
                info = update.latest(self.update_source, update.dll_bits(self.xll) or update.bits())
                current = update.file_version(self.xll) or self.running
                if not update.newer(info['version'], current):
                    self._update(job, output='ASF Python %s is the latest version.\n' % current)
                else:
                    self._update(job, output='Downloading ASF Python %s (%.1f MB): %s\n' % (info['version'], info['size'] / 1e6, info['url']))
                    shown = [0]

                    def progress(done, total):
                        percent = done * 100 // total if total else 0
                        if percent >= shown[0] + 10:
                            shown[0] = percent - percent % 10
                            self._update(job, output='%d%%\n' % shown[0])
                    result = update.install(self.xll, info, progress)
                    self._update(job, output='SHA-256 and version checked.\n' + update.outcome(result, self.xll) + '\n')
                    job['result'] = result
                code = 0
        except (update.UpdateError, OSError) as exc:
            self._update(job, output='Not updated: %s\n' % exc)
        self._update(job, state='done' if code == 0 else 'failed', code=code)

    def addin(self):
        """The add-in's file, its Python, its logs and the settings the studio can change."""
        values = read_settings()
        updates = values.get('PythonUpdates', '').strip().lower()
        return {'xll': self.xll, 'running': self.running, 'python': sys.prefix, 'folder': self.folder,
                'log': os.path.join(tempfile.gettempdir(), 'asfpy.log'),
                'updateLog': os.path.join(tempfile.gettempdir(), 'asfpy-update.log'),
                'releases': self.update_source or update.RELEASES, 'venv': values.get('PythonVenv', ''),
                'PythonUpdates': updates if updates in EDITABLE['PythonUpdates'] else 'check',
                'PythonTips': 'off' if values.get('PythonTips', '').strip().lower() == 'off' else 'on'}

    def open_item(self, what):
        """The add-in's folder (its file selected), a log, or the Python, in Explorer or Notepad."""
        info = self.addin()
        if os.name != 'nt':
            raise OSError('Explorer is on Windows only')
        if what == 'addin' and info['xll']:
            subprocess.Popen(['explorer', '/select,', info['xll']])
        elif what in ('log', 'updateLog'):
            if not os.path.isfile(info[what]):
                raise FileNotFoundError('%s is not there yet' % info[what])
            subprocess.Popen(['notepad.exe', info[what]])
        elif what == 'python':
            os.startfile(info['python'])
        else:
            raise ValueError('nothing to open: %s' % what)
        return {'opened': what}

    def previous(self):
        return previous(self.prefix)

    def previous_install(self, name):
        packages = [p['name'] for p in self._previous(name)['packages']]
        if not packages:
            raise ValueError('that Python has no package this one lacks')
        return self.pip('install ' + ' '.join(packages))

    def previous_remove(self, name):
        """A Python an update left, removed: renamed first, which fails while an Excel still uses it."""
        folder = os.path.join(os.path.dirname(os.path.abspath(self.prefix)), self._previous(name)['folder'])
        gone = folder + '-removing'
        try:
            os.rename(folder, gone)
        except OSError:
            raise PermissionError('%s is in use: close every Excel that runs the older add-in, then try again' % name)
        shutil.rmtree(gone, ignore_errors=True)
        return {'removed': name, 'left': os.path.exists(gone)}

    def _previous(self, name):
        for p in previous(self.prefix):
            if p['folder'] == name:
                return p
        raise LookupError('no earlier Python named %s' % name)


def read_settings():
    """The add-in's settings, as Excel's add-ins read them from HKEY_CURRENT_USER."""
    try:
        import winreg
    except ImportError:
        return dict(_no_registry)
    out = {}
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, SETTINGS_KEY) as k:
            for name in SETTINGS:
                try:
                    out[name] = str(winreg.QueryValueEx(k, name)[0])
                except OSError:
                    pass
    except OSError:
        pass
    return out


def write_setting(name, value):
    """One of the settings the studio changes (EDITABLE), where the add-in reads it when Excel
    loads it; ASF.xlam's Python settings changes the same."""
    value = str(value).strip().lower()
    if name not in EDITABLE or value not in EDITABLE[name]:
        raise ValueError('%s takes %s' % (name, ', '.join(EDITABLE.get(name, ()))) if name in EDITABLE
                         else 'the studio does not change %s' % name)
    try:
        import winreg
    except ImportError:
        _no_registry[name] = value
        return {name: value}
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, SETTINGS_KEY) as k:
        winreg.SetValueEx(k, name, 0, winreg.REG_SZ, value)
    return {name: value}


def _key(name):
    return re.sub(r'[-_.]+', '-', name).lower()


def _site(prefix):
    for site in (os.path.join(prefix, 'Lib', 'site-packages'), os.path.join(prefix, 'lib', 'site-packages')):
        if os.path.isdir(site):
            return site
    return os.path.join(prefix, 'Lib', 'site-packages')


def previous(prefix):
    """The other Pythons of the add-in, of this bitness, beside the one at prefix (an update's new
    add-in unpacks a new one): their size and the packages they have that this one lacks."""
    prefix = os.path.abspath(prefix)
    root, me = os.path.split(prefix)
    mine = RUNTIME.match(me)
    if not mine or not os.path.isdir(root):
        return []
    have = {_key(d.metadata.get('Name') or '') for d in importlib.metadata.distributions(path=[_site(prefix)])}
    out = []
    for name in sorted(os.listdir(root)):
        folder = os.path.join(root, name)
        if name.endswith('-removing'):              # what a removal couldn't delete
            shutil.rmtree(folder, ignore_errors=True)
            continue
        m = RUNTIME.match(name)
        if not m or name == me or m.group('bits') != mine.group('bits') or not os.path.isdir(folder):
            continue
        extra = {}
        for d in importlib.metadata.distributions(path=[_site(folder)]):
            n = d.metadata.get('Name') or ''
            if n and _key(n) not in have:
                extra[_key(n)] = {'name': n, 'version': d.version}
        size = 0
        for path, _, files in os.walk(folder):
            for f in files:
                try:
                    size += os.path.getsize(os.path.join(path, f))
                except OSError:
                    pass
        out.append({'folder': name, 'mb': round(size / 1e6), 'packages': sorted(extra.values(), key=lambda p: p['name'].lower())})
    return out


def _python():
    """python.exe beside this one: pip's output reaches the page either way, and the windowless
    pythonw.exe has no console to give pip."""
    exe = sys.executable
    if os.path.basename(exe).lower() == 'pythonw.exe':
        alt = os.path.join(os.path.dirname(exe), 'python.exe')
        if os.path.isfile(alt):
            return alt
    return exe


def _stand_in_asfpy():
    if 'asfpy' not in sys.modules:
        spec = importlib.util.spec_from_file_location('asfpy', os.path.join(STUBS, 'asfpy.py'))
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        sys.modules['asfpy'] = module


def lint(source, filename='<file>'):
    """Syntax errors, then pyflakes's warnings (undefined names, unused imports...)."""
    try:
        tree = ast.parse(source, filename)
    except SyntaxError as e:
        return [{'line': e.lineno or 1, 'col': max((e.offset or 1) - 1, 0), 'message': e.msg, 'severity': 'error'}]
    try:
        from pyflakes import checker
    except ImportError:
        return []
    w = checker.Checker(tree, filename=filename)
    return sorted(({'line': m.lineno, 'col': getattr(m, 'col', 0) or 0, 'message': m.message % m.message_args,
                    'severity': 'warning'} for m in w.messages), key=lambda p: (p['line'], p['col']))


def complete(studio, source, line, column, rel=None):
    """Completions and the signature at a place of the source, by jedi."""
    try:
        import jedi
    except ImportError:
        return {'completions': [], 'signatures': []}
    project = jedi.Project(studio.folder or HERE, added_sys_path=[STUBS] + ([studio.folder] if studio.folder else []))
    path = studio.path(rel) if rel else None
    script = jedi.Script(source, path=path, project=project, environment=jedi.InterpreterEnvironment())
    try:
        found = script.complete(line, column)[:150]
        signatures = script.get_signatures(line, column)
    except Exception:
        return {'completions': [], 'signatures': []}
    return {'completions': [{'text': c.name, 'type': c.type} for c in found],
            'signatures': [{'label': s.to_string(), 'index': s.index} for s in signatures]}


def pypi(name):
    if not NAME.match(name or ''):
        raise ValueError('a package name: letters, digits, . _ -')
    try:
        with urllib.request.urlopen('https://pypi.org/pypi/%s/json' % urllib.parse.quote(name), timeout=15) as r:
            info = json.load(r)['info']
    except Exception as exc:
        code = getattr(exc, 'code', None)
        raise LookupError('%s isn\'t on PyPI' % name if code == 404 else 'PyPI didn\'t answer: %s' % exc)
    return {'name': info['name'], 'version': info['version'], 'summary': info.get('summary') or '',
            'home': info.get('home_page') or (info.get('project_urls') or {}).get('Homepage') or ''}


# --- HTTP ----------------------------------------------------------------------------------------

class Handler(http.server.BaseHTTPRequestHandler):
    studio = None
    server_version = 'ASFStudio'

    def log_message(self, fmt, *args):
        pass

    def _send(self, code, body, kind='application/json', headers=()):
        data = body if isinstance(body, bytes) else json.dumps(body).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', kind)
        self.send_header('Content-Length', str(len(data)))
        self.send_header('Cache-Control', 'no-store')
        self.send_header('X-Content-Type-Options', 'nosniff')
        self.send_header('Content-Security-Policy', "default-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data:")
        for k, v in headers:
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(data)

    def _host_ok(self):
        port = self.server.server_address[1]
        return self.headers.get('Host') in ('127.0.0.1:%d' % port, 'localhost:%d' % port)

    def _cookie_ok(self):
        want = 'asfstudio=%s' % self.studio.token
        return any(c.strip() == want for c in (self.headers.get('Cookie') or '').split(';'))

    def do_GET(self):
        url = urllib.parse.urlsplit(self.path)
        query = urllib.parse.parse_qs(url.query)
        if not self._host_ok():
            return self._send(403, {'error': 'forbidden'})
        if url.path == '/api/alive':
            ok = secrets.compare_digest(self.headers.get('X-Studio-Token', ''), self.studio.token)
            return self._send(200 if ok else 403, {'folder': self.studio.folder} if ok else {'error': 'forbidden'})
        if url.path == '/' and 't' in query:
            if not secrets.compare_digest(query['t'][0], self.studio.token):
                return self._send(403, {'error': 'forbidden'})
            view = query.get('view', [''])[0]
            return self._send(302, b'', 'text/plain', [('Location', '/#' + view if view in VIEWS else '/'), (
                'Set-Cookie', 'asfstudio=%s; HttpOnly; SameSite=Strict; Path=/' % self.studio.token)])
        if not self._cookie_ok():
            return self._send(403, b'Open the studio from Excel: ASF tab, Python studio.', 'text/plain; charset=utf-8')
        self.studio.seen = time.time()
        if url.path.startswith('/api/'):
            return self._api('GET', url.path[5:], {k: v[0] for k, v in query.items()})
        rel = 'index.html' if url.path == '/' else url.path.lstrip('/')
        full = os.path.realpath(os.path.join(STATIC, rel))
        if os.path.commonpath([full, os.path.realpath(STATIC)]) != os.path.realpath(STATIC) or not os.path.isfile(full):
            return self._send(404, {'error': 'not found'})
        with open(full, 'rb') as f:
            self._send(200, f.read(), TYPES.get(os.path.splitext(full)[1], 'application/octet-stream'))

    def do_POST(self):
        port = self.server.server_address[1]
        origin = self.headers.get('Origin')
        if (not self._host_ok() or not self._cookie_ok() or self.headers.get('X-Studio') != '1'
                or (origin and origin not in ('http://127.0.0.1:%d' % port, 'http://localhost:%d' % port))):
            return self._send(403, {'error': 'forbidden'})
        self.studio.seen = time.time()
        try:
            body = json.loads(self.rfile.read(int(self.headers.get('Content-Length') or 0)) or b'{}')
        except ValueError:
            return self._send(400, {'error': 'not JSON'})
        return self._api('POST', urllib.parse.urlsplit(self.path).path[5:], body)

    def _api(self, method, name, args):
        s = self.studio
        try:
            if method == 'GET' and name == 'info':
                result = {'python': sys.version.split()[0], 'bits': 64 if sys.maxsize > 2 ** 32 else 32,
                          'prefix': sys.prefix, 'folder': s.folder, 'excel': excel.connected()}
            elif method == 'GET' and name == 'files':
                result = s.files()
            elif method == 'GET' and name == 'file':
                result = {'path': args['path'], 'text': s.read(args['path'])}
            elif method == 'GET' and name == 'outline':
                result = s.outline()
            elif method == 'GET' and name == 'packages':
                result = s.packages()
            elif method == 'GET' and name == 'jobs':
                with s.lock:
                    result = {'revision': s.revision, 'jobs': [dict(j) for j in s.jobs]}
            elif method == 'GET' and name == 'pypi':
                result = pypi(args.get('name', ''))
            elif method == 'GET' and name == 'update':
                result = s.update_status(args.get('live') in ('1', 'true'))
            elif method == 'GET' and name == 'previous':
                result = s.previous()
            elif method == 'GET' and name == 'addin':
                result = s.addin()
            elif method == 'POST' and name == 'setting':
                result = write_setting(args.get('name', ''), args.get('value', ''))
            elif method == 'POST' and name == 'open':
                result = s.open_item(args.get('what', ''))
            elif method == 'POST' and name == 'ping':
                result = {'busy': s.busy()}
            elif method == 'POST' and name == 'save':
                result = s.save(args['path'], args['text'])
            elif method == 'POST' and name == 'new':
                result = s.new(args['path'])
            elif method == 'POST' and name == 'reveal':
                result = s.reveal()
            elif method == 'POST' and name == 'lint':
                result = lint(args['text'])
            elif method == 'POST' and name == 'complete':
                result = complete(s, args['text'], int(args['line']), int(args['column']), args.get('path'))
            elif method == 'POST' and name == 'console':
                result = s.console(args.get('code', ''), bool(args.get('restart')))
            elif method == 'POST' and name == 'pip':
                result = s.pip(args.get('command', ''))
            elif method == 'POST' and name == 'update':
                result = s.update()
            elif method == 'POST' and name == 'previous/install':
                result = s.previous_install(args.get('folder', ''))
            elif method == 'POST' and name == 'previous/remove':
                result = s.previous_remove(args.get('folder', ''))
            elif method == 'POST' and name == 'excel/reload':
                result = {'done': excel.reload()}
            elif method == 'POST' and name == 'excel/run':
                out = excel.run('AsfPyRunScript', s.path(args['path']))
                result = {'output': out if isinstance(out, str) else 'The Python add-in has no Python: About Python on the ASF tab says why.'}
            elif method == 'POST' and name == 'excel/macro-in-excel':
                result = excel.macro_in_excel(args['name'])
            elif method == 'POST' and name == 'excel/macro':
                values = args.get('args') or []
                if not isinstance(values, list) or len(values) > 30 or any(
                        v is not None and not isinstance(v, (str, int, float, bool)) for v in values):
                    raise ValueError('the arguments of a macro are numbers, texts, TRUE or FALSE, or blank')
                result = excel.macro(args['name'], values, bool(args.get('front')))
            else:
                return self._send(404, {'error': 'no such call'})
        except excel.NotReachable as exc:
            return self._send(503, {'error': str(exc)})
        except update.UpdateError as exc:
            return self._send(502, {'error': str(exc)})
        except (PermissionError, ValueError, LookupError, KeyError, OSError, FileExistsError) as exc:
            return self._send(400, {'error': str(exc) or type(exc).__name__})
        return self._send(200, result)


# --- starting -----------------------------------------------------------------------------------

def _state_file():
    return os.path.join(_local_app_data(), 'ASF', 'studio.json')


def _running(folder):
    """The URL of a studio that runs already for this folder, or None."""
    try:
        with open(_state_file(), encoding='utf-8') as f:
            state = json.load(f)
        req = urllib.request.Request('http://127.0.0.1:%d/api/alive' % state['port'],
                                     headers={'X-Studio-Token': state['token']})
        with urllib.request.urlopen(req, timeout=2) as r:
            if os.path.normcase(json.load(r).get('folder', '')) == os.path.normcase(os.path.abspath(folder)):
                return 'http://127.0.0.1:%d/?t=%s' % (state['port'], state['token'])
    except Exception:
        pass
    return None


def open_window(url):
    """An app window of Microsoft Edge, with a profile of its own; the default browser otherwise."""
    edge = None
    if os.name == 'nt':
        try:
            import winreg
            for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
                try:
                    with winreg.OpenKey(hive, r'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe') as k:
                        edge = winreg.QueryValue(k, None)
                        break
                except OSError:
                    pass
        except ImportError:
            pass
        for base in (os.environ.get('ProgramFiles(x86)'), os.environ.get('ProgramFiles')):
            candidate = base and os.path.join(base, 'Microsoft', 'Edge', 'Application', 'msedge.exe')
            if not edge and candidate and os.path.isfile(candidate):
                edge = candidate
    if edge and os.path.isfile(edge):
        profile = os.path.join(_local_app_data(), 'ASF', 'studio-window')
        subprocess.Popen([edge, '--app=' + url, '--window-size=1320,880', '--no-first-run',
                          '--no-default-browser-check', '--user-data-dir=' + profile])
    else:
        import webbrowser
        webbrowser.open(url)


def serve(studio, port=0):
    Handler.studio = studio
    server = http.server.ThreadingHTTPServer(('127.0.0.1', port), Handler)
    server.daemon_threads = True

    def watch():
        while True:
            time.sleep(5)
            if time.time() - studio.seen > studio.idle and not studio.busy():
                server.shutdown()
                return
    threading.Thread(target=watch, name='idle', daemon=True).start()
    return server


def main(argv=None):
    a = argparse.ArgumentParser(prog='asfpy_studio')
    a.add_argument('--folder', default='')
    a.add_argument('--port', type=int, default=0)
    a.add_argument('--idle', type=int, default=IDLE)
    a.add_argument('--log', default='')
    a.add_argument('--xll', default='')
    a.add_argument('--running', default='')
    a.add_argument('--update-source', default='')
    a.add_argument('--view', default='', help='the view to open: ' + ', '.join(VIEWS))
    a.add_argument('--no-window', action='store_true')
    o = a.parse_args(argv)
    folder = o.folder.strip('"')
    if o.log or sys.stdout is None:           # pythonw: no console to print to
        log = open(o.log or os.path.join(os.environ.get('TEMP', '.'), 'asfpy-studio.log'), 'a', encoding='utf-8')
        sys.stdout = sys.stderr = log
    view = '&view=' + o.view if o.view in VIEWS else ''
    url = None if o.no_window else _running(folder)
    if url:
        open_window(url + view)
        return
    studio = Studio(folder, o.idle, o.xll.strip('"'), o.running, o.update_source)
    server = serve(studio, o.port)
    port = server.server_address[1]
    url = 'http://127.0.0.1:%d/?t=%s' % (port, studio.token) + view
    try:
        os.makedirs(os.path.dirname(_state_file()), exist_ok=True)
        with open(_state_file(), 'w', encoding='utf-8') as f:
            json.dump({'port': port, 'token': studio.token, 'pid': os.getpid(), 'folder': studio.folder}, f)
    except OSError:
        pass
    print('studio on port %d for %s' % (port, studio.folder or 'no folder'), flush=True)
    if not o.no_window:
        open_window(url)
    else:
        print(url, flush=True)
    server.serve_forever()
