"""Updates of the Python add-in (asfpy64.xll, asfpy32.xll) from the project's releases on GitHub.

The add-in runs this file with its Python: when Excel loads it (the setting PythonUpdates: check,
the default, asks before installing; auto installs; off never looks) and for its command
AsfPyUpdate. The studio imports it. Only the standard library: it also runs with PythonVenv's Python.

An update is the release's asset of the add-in's bitness. Before it replaces anything, its SHA-256
must match the digest GitHub gives for the asset (or the release's SHA256SUMS), and the file must
be a DLL of that bitness whose version mark names the release. Windows lets a loaded DLL be
renamed, not overwritten: the add-in becomes NAME.old and the new file takes its name, so the
next Excel loads it, and the add-in deletes NAME.old then.

    python update.py --xll PATH [--running VERSION] [--source URL] [--ask | --mode check|auto]
"""
import argparse
import ctypes
import hashlib
import json
import mmap
import os
import re
import struct
import sys
import tempfile
import time
import traceback
import urllib.parse
import urllib.request

SOURCE = 'https://api.github.com/repos/ECP-Solutions/ASF/releases/latest'
RELEASES = 'https://github.com/ECP-Solutions/ASF/releases'
MARK = re.compile(rb'ASFPY-VERSION:(\d+(?:\.\d+)*);')    # in the add-in's file, from version.h
INTERVAL = 24 * 3600        # between two checks when Excel starts
REMIND = 7 * 24 * 3600      # after No, until a version is offered again
TITLE = 'ASF Python'
LOOPBACK = ('127.0.0.1', 'localhost', '::1')


class UpdateError(Exception):
    pass


def version_tuple(text):
    m = re.match(r'v?(\d+(?:\.\d+)*)', (text or '').strip())
    return tuple(int(p) for p in m.group(1).split('.')) if m else None


def newer(a, b):
    """a newer than b; a version that can't be read is never newer."""
    ta, tb = version_tuple(a), version_tuple(b)
    if ta is None:
        return False
    if tb is None:
        return True
    n = max(len(ta), len(tb))
    return ta + (0,) * (n - len(ta)) > tb + (0,) * (n - len(tb))


def file_version(path):
    """The version an add-in file carries (its mark ASFPY-VERSION:x.y.z;), or None."""
    try:
        with open(path, 'rb') as f, mmap.mmap(f.fileno(), 0, access=mmap.ACCESS_READ) as m:
            found = MARK.search(m)
            return found.group(1).decode('ascii') if found else None
    except (OSError, ValueError):
        return None


def dll_bits(path):
    """64 or 32 for a Windows DLL of that bitness, else None."""
    try:
        with open(path, 'rb') as f:
            head = f.read(4096)
        if head[:2] != b'MZ':
            return None
        pe = struct.unpack_from('<I', head, 0x3C)[0]
        if head[pe:pe + 4] != b'PE\0\0':
            return None
        machine, = struct.unpack_from('<H', head, pe + 4)
        characteristics, = struct.unpack_from('<H', head, pe + 22)
        if not characteristics & 0x2000:            # IMAGE_FILE_DLL
            return None
        return {0x8664: 64, 0x14C: 32}.get(machine)
    except (OSError, struct.error):
        return None


def bits():
    return 64 if sys.maxsize > 2 ** 32 else 32


def _open(url, timeout=20):
    """https, or http to this computer (a mirror for tests); GitHub's API wants a User-Agent."""
    parts = urllib.parse.urlsplit(url)
    if parts.scheme != 'https' and not (parts.scheme == 'http' and parts.hostname in LOOPBACK):
        raise UpdateError('updates come over https only: %s' % url)
    req = urllib.request.Request(url, headers={'User-Agent': 'asfpy-update', 'Accept': 'application/vnd.github+json'})
    return urllib.request.urlopen(req, timeout=timeout)


def latest(source=None, want_bits=None):
    """The latest release and its add-in of this bitness: version, tag, page, url, size, sha256."""
    want_bits = want_bits or bits()
    try:
        with _open(source or SOURCE) as r:
            release = json.load(r)
    except UpdateError:
        raise
    except Exception as exc:
        raise UpdateError("the releases can't be read (%s)" % exc)
    name = 'asfpy%d.xll' % want_bits
    assets = {a.get('name', '').lower(): a for a in release.get('assets', [])}
    asset = assets.get(name)
    if not asset:
        raise UpdateError('release %s has no %s' % (release.get('tag_name', '?'), name))
    digest = asset.get('digest') or ''
    sha = digest[7:].lower() if digest.lower().startswith('sha256:') else ''
    if not sha and 'sha256sums' in assets:
        with _open(assets['sha256sums']['browser_download_url']) as r:
            for line in r.read().decode('utf-8', 'replace').splitlines():
                parts = line.split()
                if len(parts) == 2 and parts[1].lstrip('*').lower() == name:
                    sha = parts[0].lower()
    if not re.fullmatch(r'[0-9a-f]{64}', sha):
        raise UpdateError('release %s gives no SHA-256 for %s' % (release.get('tag_name', '?'), name))
    tag = release.get('tag_name', '')
    return {'version': tag.lstrip('v'), 'tag': tag, 'name': name, 'page': release.get('html_url') or RELEASES,
            'url': asset['browser_download_url'], 'size': int(asset.get('size') or 0), 'sha256': sha, 'bits': want_bits}


def check(xll, running=None, source=None):
    """What the add-in file at xll is, and what the latest release offers."""
    on_disk = file_version(xll) or running or ''
    info = latest(source, dll_bits(xll) or bits())
    return {'running': running or on_disk, 'installed': on_disk, 'latest': info['version'], 'page': info['page'],
            'size': info['size'], 'available': newer(info['version'], on_disk),
            'restart': bool(running) and newer(on_disk, running), 'release': info}


# --- downloading and installing ------------------------------------------------------------------

def download(info, folder, progress=None):
    """The add-in of the release in folder (NAME.download), checked; its path."""
    path = os.path.join(folder, info['name'] + '.download')
    digest, done = hashlib.sha256(), 0
    try:
        with _open(info['url'], timeout=60) as r, open(path, 'wb') as f:
            while True:
                chunk = r.read(1 << 18)
                if not chunk:
                    break
                f.write(chunk)
                digest.update(chunk)
                done += len(chunk)
                if progress:
                    progress(done, info['size'])
        problem = None
        if info['size'] and done != info['size']:
            problem = 'the download has %d bytes of %d' % (done, info['size'])
        elif digest.hexdigest() != info['sha256']:
            problem = "the download's SHA-256 is not the release's"
        elif dll_bits(path) != info['bits']:
            problem = 'the download is not a %d-bit add-in' % info['bits']
        elif file_version(path) != info['version']:
            problem = 'the download says it is version %s, not %s' % (file_version(path), info['version'])
        if problem:
            raise UpdateError(problem)
    except BaseException:
        try:
            os.remove(path)
        except OSError:
            pass
        raise
    return path


def swap(xll, new):
    """new in the place of xll; xll kept as xll.old (or xll.N.old while an Excel still uses that)."""
    old = xll + '.old'
    if os.path.exists(old):
        try:
            os.remove(old)
        except OSError:
            old = '%s.%d.old' % (xll, int(time.time()))
    os.replace(xll, old)
    try:
        os.replace(new, xll)
    except OSError:
        os.replace(old, xll)
        raise
    return old


def updates_folder():
    return os.path.join(os.environ.get('LOCALAPPDATA') or tempfile.gettempdir(), 'ASF', 'updates')


def install(xll, info, progress=None):
    """The release's add-in in xll's place: {'installed': xll}; when the add-in's folder can't be
    written to, the download is kept elsewhere: {'saved': path}."""
    try:
        new = download(info, os.path.dirname(os.path.abspath(xll)), progress)
    except PermissionError:
        folder = updates_folder()
        os.makedirs(folder, exist_ok=True)
        new = download(info, folder, progress)
        saved = os.path.join(folder, 'asfpy%d-%s.xll' % (info['bits'], info['version']))
        os.replace(new, saved)
        return {'saved': saved, 'version': info['version']}
    try:
        swap(xll, new)
    except PermissionError:
        folder = updates_folder()
        os.makedirs(folder, exist_ok=True)
        saved = os.path.join(folder, 'asfpy%d-%s.xll' % (info['bits'], info['version']))
        os.replace(new, saved)
        return {'saved': saved, 'version': info['version']}
    return {'installed': xll, 'version': info['version']}


def outcome(result, xll):
    if 'installed' in result:
        return ('ASF Python %s is installed. Restart Excel to use it.\n\nPackages you added with pip stay with the '
                "previous Python: Packages in Python studio installs them again." % result['version'])
    return ("ASF Python %s is downloaded to\n%s\nbut the add-in's folder can't be written to. Close Excel and "
            'copy it over\n%s' % (result['version'], result['saved'], xll))


# --- state, shared by every Excel ----------------------------------------------------------------

def state_file():
    return os.path.join(os.environ.get('LOCALAPPDATA') or tempfile.gettempdir(), 'ASF', 'update.json')


def load_state():
    try:
        with open(state_file(), encoding='utf-8') as f:
            state = json.load(f)
        return state if isinstance(state, dict) else {}
    except (OSError, ValueError):
        return {}


def save_state(state):
    try:
        os.makedirs(os.path.dirname(state_file()), exist_ok=True)
        with open(state_file(), 'w', encoding='utf-8') as f:
            json.dump(state, f)
    except OSError:
        pass


class Lock:
    """One updater at a time, whichever Excel started it; a lock older than an hour is stale."""
    def __init__(self):
        self.path = state_file() + '.lock'
        self.held = False

    def __enter__(self):
        os.makedirs(os.path.dirname(self.path), exist_ok=True)
        for _ in range(2):
            try:
                os.close(os.open(self.path, os.O_CREAT | os.O_EXCL | os.O_WRONLY))
                self.held = True
                break
            except FileExistsError:
                try:
                    if time.time() - os.path.getmtime(self.path) > 3600:
                        os.remove(self.path)
                        continue
                except OSError:
                    pass
                break
        return self

    def __exit__(self, *exc):
        if self.held:
            try:
                os.remove(self.path)
            except OSError:
                pass


# --- talking to the user -------------------------------------------------------------------------

def tell(text, question=False):
    """A message box (Yes/No for a question: True for Yes); printed where there is none."""
    if os.name == 'nt':
        flags = 0x00010000 | 0x00040000 | (0x04 | 0x20 if question else 0x40)   # foreground, topmost
        return ctypes.windll.user32.MessageBoxW(None, text, TITLE, flags) == 6
    print(text)
    return False


def offer(status):
    size = status['size'] / 1e6
    return ('ASF Python %s is available; this add-in is %s.\n\nDownload it (%.0f MB) and install it now? Excel uses it '
            'the next time it starts.\n\nWhat changed: %s' % (status['latest'], status['installed'] or '?', size, status['page']))


def ask_now(xll, running, source):
    """AsfPyUpdate: check now, and say what came of it."""
    try:
        status = check(xll, running, source)
    except UpdateError as exc:
        tell('ASF Python could not check for updates: %s.\n\nThe releases: %s' % (exc, RELEASES))
        return
    state = load_state()
    state.update(checked=time.time(), latest=status['latest'], page=status['page'], size=status['size'])
    save_state(state)
    if status['restart'] and not status['available']:
        tell('ASF Python %s is installed. Restart Excel to use it.' % status['installed'])
    elif not status['available']:
        tell('ASF Python %s is the latest version.' % status['installed'])
    elif tell(offer(status), question=True):
        with Lock() as lock:
            if not lock.held:
                tell('ASF Python is being updated by another Excel.')
                return
            try:
                tell(outcome(install(xll, status['release']), xll))
            except (UpdateError, OSError) as exc:
                tell('ASF Python was not updated: %s.\n\nThe releases: %s' % (exc, RELEASES))


def on_start(xll, running, mode, source, now=None):
    """When Excel loads the add-in: at most a check a day; check asks before installing, auto doesn't.
    What it did, for the log."""
    now = now or time.time()
    state = load_state()
    if now - state.get('checked', 0) < INTERVAL:
        return 'checked %.0f hours ago' % ((now - state['checked']) / 3600)
    with Lock() as lock:
        if not lock.held:
            return 'another updater runs'
        status = check(xll, running, source)
        state.update(checked=now, latest=status['latest'], page=status['page'], size=status['size'])
        save_state(state)
        if not status['available']:
            return '%s is the latest' % status['installed']
        if state.get('saved') == status['latest']:
            return '%s is downloaded already, waiting to be copied' % status['latest']
        declined = state.get('declined', {})
        if mode != 'auto':
            if now - declined.get(status['latest'], 0) < REMIND:
                return '%s was declined' % status['latest']
            if not tell(offer(status) + '\n\n(No: asked again in a week. PythonUpdates in Python settings turns '
                        'this off.)', question=True):
                state['declined'] = {status['latest']: now}
                save_state(state)
                return '%s declined' % status['latest']
        result = install(xll, status['release'])
        if 'saved' in result:
            state['saved'] = result['version']
            save_state(state)
        tell(outcome(result, xll))
        return '%s %s' % ('installed' if 'installed' in result else 'saved', result['version'])


def main(argv=None):
    a = argparse.ArgumentParser(prog='update.py')
    a.add_argument('--xll', required=True)
    a.add_argument('--running', default='')
    a.add_argument('--source', default='')
    a.add_argument('--mode', default='check', choices=['check', 'auto'])
    a.add_argument('--ask', action='store_true', help='check now and say what came of it (AsfPyUpdate)')
    a.add_argument('--log', default=os.path.join(tempfile.gettempdir(), 'asfpy-update.log'))
    o = a.parse_args(argv)
    xll = o.xll.strip('"')
    if sys.stdout is None or not o.ask:        # pythonw, or Excel's start: the log only
        sys.stdout = sys.stderr = open(o.log, 'a', encoding='utf-8')
    stamp = time.strftime('%Y-%m-%d %H:%M:%S')
    try:
        if o.ask:
            ask_now(xll, o.running, o.source or None)
        else:
            print('%s %s: %s' % (stamp, xll, on_start(xll, o.running, o.mode, o.source or None)))
    except Exception:
        print('%s %s:\n%s' % (stamp, xll, traceback.format_exc()))
        if o.ask:
            tell('ASF Python could not check for updates. Details: %s' % o.log)


if __name__ == '__main__':
    main()
