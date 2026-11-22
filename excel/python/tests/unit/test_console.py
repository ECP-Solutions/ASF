"""Tests of xll/console.py, the Python console of the ASF tab, without Excel: run as asfpy.xll runs it
(python -I -i console.py FOLDER), with what a user types on its standard input."""
import os
import shutil
import subprocess
import sys
import tempfile
import textwrap
import unittest

CONSOLE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'xll', 'console.py')

STATS = '''
from asfpy import export, xl_func, xl_macro, xl_app

@xl_func
def double(x):
    return 2 * x

@xl_func(volatile=True)
def triple(x):
    return 3 * x

@export
def half(x):
    return x / 2

@xl_macro(name='Tidy')
def tidy():
    return xl_app()
'''


class Console(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix='asfpy-console-')
        self.folder = os.path.join(self.tmp, 'python')
        os.makedirs(self.folder)
        with open(os.path.join(self.folder, 'stats.py'), 'w', encoding='utf-8') as f:
            f.write(textwrap.dedent(STATS))

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def run_console(self, typed, folder=None):
        r = subprocess.run([sys.executable, '-I', '-i', CONSOLE, self.folder if folder is None else folder],
                           input=textwrap.dedent(typed), capture_output=True, text=True, timeout=120)
        return r.stdout, r.stderr

    def test_the_banner(self):
        out, err = self.run_console('')
        self.assertIn('ASF Python console: Python %s' % sys.version.split()[0], out)
        self.assertIn('Python folder (importable here): %s' % self.folder, out)
        self.assertIn('pip("install numpy")', out)

    def test_the_python_folder_imports_with_the_stand_in_asfpy(self):
        out, err = self.run_console('''
            import stats
            print(stats.double(21), stats.triple(2), stats.half(5))
            try:
                stats.tidy()
            except RuntimeError as e:
                print(e)
            ''')
        self.assertIn('42 6 2.5', out)
        self.assertIn('xl_app() works inside Excel only', out)
        self.assertNotIn('Traceback', err)

    def test_pip_runs_for_this_python(self):
        out, err = self.run_console('pip("--version")\n')
        self.assertIn('pip ', out + err)
        self.assertNotIn('NameError', err)

    def test_no_folder(self):
        out, err = self.run_console('import sys\nprint(len(sys.argv))\n', folder='')
        self.assertIn('Python folder (importable here): none', out)
        self.assertIn('\n1\n', '\n' + out)     # the folder argument is gone from sys.argv


if __name__ == '__main__':
    unittest.main()
