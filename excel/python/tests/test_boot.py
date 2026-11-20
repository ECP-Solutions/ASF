"""Tests of xll/boot.py, the Python side of asfpy.xll, without Excel: on any OS, with any Python
from 3.11. boot.py is loaded from its file and given a temporary Python folder.

    python -m unittest discover -s excel/python/tests -p "test_*.py"
"""
import datetime
import importlib.util
import os
import shutil
import sys
import tempfile
import textwrap
import unittest

BOOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'xll', 'boot.py')

STATS = '''
import datetime
from asfpy import export, xl_func


@xl_func
def double(x):
    """Twice x.

    More text, which Insert Function doesn't show."""
    return x * 2


@xl_func(volatile=True, thread_safe=True)
def flags_vt():
    return 1


@xl_func(thread_safe=True, macro=True)
def refused():
    return 1


@xl_func(name='STATS.HINTS')
def hints(n: int, when: datetime.date, text: str, values: list, f: float, b: bool, dt: datetime.datetime):
    return [[repr(n), repr(when), repr(text), repr(values), repr(f), repr(b), repr(dt)]]


@xl_func
def nextday(when: 'datetime.date'):
    return when + datetime.timedelta(days=1)


@xl_func
def grid():
    return [[1, 'a'], [2]]


@xl_func
def bad():
    return float('nan')


@export
def scale(values: list, factor: float):
    return [v * factor for v in values]


@export
def fails():
    raise ValueError('on purpose')


def plain():
    return 'not marked'
'''

USES = '''
import z_helper
from asfpy import export


@export
def uses_helper():
    return z_helper.VALUE
'''

HELPER = '''
import asfpy
asfpy.helper_loads = getattr(asfpy, 'helper_loads', 0) + 1
VALUE = 7
'''


class BootTest(unittest.TestCase):
    """A fresh boot.py and Python folder for each test; sys is put back afterwards."""

    def setUp(self):
        self.saved = (sys.path[:], sys.stdout, sys.stderr, set(sys.modules))
        self.tmp = tempfile.mkdtemp(prefix='asfpy-boot-')
        self.folder = os.path.join(self.tmp, 'python')
        os.makedirs(self.folder)
        self.write('stats.py', STATS)
        self.write('a_uses.py', USES)
        self.write('z_helper.py', HELPER)
        self.write('broken.py', 'raise RuntimeError("broken on purpose")\n')
        spec = importlib.util.spec_from_file_location('asfpy_boot_under_test', BOOT)
        self.boot = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(self.boot)
        self.functions = self.boot.start(self.folder, '', '', os.path.join(self.tmp, 'asfpy.log'))

    def tearDown(self):
        path, out, err, modules = self.saved
        sys.path[:] = path
        sys.stdout, sys.stderr = out, err
        for name in set(sys.modules) - modules:
            del sys.modules[name]
        shutil.rmtree(self.tmp, ignore_errors=True)

    def write(self, name, text):
        with open(os.path.join(self.folder, name), 'w', encoding='utf-8') as f:
            f.write(textwrap.dedent(text))

    def slot(self, name):
        return [f[0] for f in self.functions].index(name)

    def call(self, name, *args):
        return self.boot.xl_call(self.slot(name), list(args))

    def info(self):
        return self.boot.root(1).startInfo


class Start(BootTest):
    def test_functions_their_arguments_descriptions_and_suffixes(self):
        by_name = {f[0]: f for f in self.functions}
        self.assertEqual(by_name['double'], ('double', ['x'], 'Twice x.', ''))
        self.assertEqual(by_name['flags_vt'][3], '!$')
        self.assertEqual(by_name['STATS.HINTS'][1], ['n', 'when', 'text', 'values', 'f', 'b', 'dt'])
        self.assertNotIn('refused', by_name)
        self.assertNotIn('plain', by_name)

    def test_start_info_says_what_was_refused_and_what_did_not_import(self):
        info = self.info()
        self.assertIn("stats.refused: thread_safe and macro can't go together", info)
        self.assertIn('broken.py did not import: RuntimeError: broken on purpose', info)
        self.assertNotIn('__pycache__', info)

    def test_a_module_imported_by_another_runs_once(self):
        self.assertEqual(sys.modules['asfpy'].helper_loads, 1)
        self.assertEqual(self.boot.root(1).call('a_uses.uses_helper'), 7)

    def test_no_folder(self):
        self.boot.root(1)  # the root works without functions
        self.boot._python_folder = os.path.join(self.tmp, 'missing')
        self.assertEqual(self.boot.load_folder(), [])
        self.assertIn('does not exist', self.info())

    def test_the_current_folder_leaves_sys_path(self):
        here = os.path.normcase(os.path.abspath(os.getcwd()))
        self.assertNotIn(here, [os.path.normcase(os.path.abspath(p)) for p in sys.path])
        self.assertIn(self.folder, sys.path)


class Calls(BootTest):
    def test_a_number(self):
        self.assertEqual(self.call('double', 2.0), (self.boot.TAG_FLOAT, 4.0))

    def test_type_hints(self):
        tag, rows = self.call('STATS.HINTS', 2.0, 45000.0, 12.0, 7.0, 3.0, 1.0, 45000.5)
        self.assertEqual(tag, self.boot.TAG_GRID)
        self.assertEqual([v for _, v in rows[0]],
                         ['2', 'datetime.date(2023, 3, 15)', "'12'", '[7.0]', '3.0', 'True',
                          'datetime.datetime(2023, 3, 15, 12, 0)'])

    def test_an_int_hint_refuses_a_fraction(self):
        tag, message = self.call('STATS.HINTS', 2.5, 1.0, '', None, 1.0, False, 1.0)
        self.assertEqual(tag, self.boot.TAG_FAIL)
        self.assertIn('expected a whole number', message)

    def test_empty_cells_stay_none_and_a_list_hint_gives_an_empty_list(self):
        tag, rows = self.call('STATS.HINTS', None, None, None, None, None, None, None)
        self.assertEqual([v for _, v in rows[0]], ['None', 'None', 'None', '[]', 'None', 'None', 'None'])

    def test_a_date_back_is_a_serial_number(self):
        self.assertEqual(self.call('nextday', 45000.0), (self.boot.TAG_FLOAT, 45001.0))

    def test_a_grid_is_padded(self):
        T = self.boot
        self.assertEqual(self.call('grid'),
                         (T.TAG_GRID, [[(T.TAG_FLOAT, 1.0), (T.TAG_STR, 'a')], [(T.TAG_FLOAT, 2.0), (T.TAG_STR, '')]]))

    def test_nan_is_num_error(self):
        self.assertEqual(self.call('bad'), (self.boot.TAG_XLERR, self.boot.XLERR_NUM))


class PyCall(BootTest):
    def test_an_exported_function_with_a_column(self):
        T = self.boot
        self.assertEqual(T.xl_call_export('stats.scale', [[[1.0], [2.0], [3.0]], 10.0]),
                         (T.TAG_GRID, [[(T.TAG_FLOAT, 10.0)], [(T.TAG_FLOAT, 20.0)], [(T.TAG_FLOAT, 30.0)]]))

    def test_worksheet_functions_are_exported_too(self):
        self.assertEqual(self.boot.xl_call_export('stats.double', [4.0]), (self.boot.TAG_FLOAT, 8.0))

    def test_an_unknown_name_is_name_error_and_never_imported(self):
        T = self.boot
        self.assertEqual(T.xl_call_export('os.system', ['calc']), (T.TAG_XLERR, T.XLERR_NAME))
        self.assertEqual(T.xl_call_export(3.0, []), (T.TAG_XLERR, T.XLERR_NAME))
        self.assertEqual(T.xl_call_export('stats.plain', []), (T.TAG_XLERR, T.XLERR_NAME))
        self.assertEqual(T.root(1).lastError, "NameError: no exported function 'stats.plain' in the Python folder")

    def test_an_exception_is_value_error_with_its_traceback(self):
        T = self.boot
        self.assertEqual(T.xl_call_export('stats.fails', []), (T.TAG_FAIL, 'ValueError: on purpose'))
        py = T.root(1)
        self.assertEqual(py.lastError, 'ValueError: on purpose')
        self.assertIn('Traceback (most recent call last)', py.lastTraceback)
        self.assertIn('in fails', py.lastTraceback)
        py.clearError()
        self.assertEqual((py.lastError, py.lastTraceback), ('', ''))


class Reload(BootTest):
    def test_changed_added_and_removed_modules(self):
        T = self.boot
        # the same size as before, in the same second: no stale bytecode
        self.write('stats.py', STATS.replace('return x * 2', 'return x * 3').replace('@xl_func\ndef bad', 'def bad'))
        self.write('extra.py', 'from asfpy import export\n\n@export\ndef answer():\n    return 42\n')
        os.remove(os.path.join(self.folder, 'a_uses.py'))
        self.functions = T.reload()
        names = [f[0] for f in self.functions]
        self.assertEqual(self.call('double', 2.0), (T.TAG_FLOAT, 6.0))
        self.assertNotIn('bad', names)
        self.assertEqual(T.xl_call_export('extra.answer', []), (T.TAG_FLOAT, 42.0))
        self.assertEqual(T.xl_call_export('a_uses.uses_helper', []), (T.TAG_XLERR, T.XLERR_NAME))
        self.assertNotIn('a_uses', sys.modules)
        self.assertIn('imported extra.py', self.info())

    def test_python_and_its_packages_stay(self):
        before = sys.modules['asfpy']
        self.boot.reload()
        self.assertIs(sys.modules['asfpy'], before)


class Proxies(BootTest):
    def test_levels(self):
        T = self.boot
        tag, message = T.com_invoke(T.root(1), 'module', 1, True, ['math'])
        self.assertEqual((tag, message), (T.TAG_FAIL, "AttributeError: '_Root' has no member 'module'"))
        tag, mod = T.com_invoke(T.root(2), 'module', 2, True, ['math'])
        self.assertEqual(tag, T.TAG_OBJECT)
        self.assertEqual(T.com_invoke(T.root(2), 'module', 2, True, ['_thread'])[0], T.TAG_FAIL)
        self.assertEqual(T.com_invoke(T.root(1), '__class__', 1, False, [])[0], T.TAG_FAIL)

    def test_a_case_insensitive_member_and_the_default_value(self):
        T = self.boot
        import collections
        counter = collections.Counter('abcab')
        self.assertEqual(T.com_invoke(counter, 'MOST_COMMON', 2, True, [1]), (T.TAG_LIST, [(T.TAG_LIST, [(T.TAG_STR, 'a'), (T.TAG_INT, 2)])]))
        self.assertEqual(T.com_invoke(counter, '', 2, False, []), (T.TAG_STR, "Counter({'a': 2, 'b': 2, 'c': 1})"))

    def test_for_each(self):
        T = self.boot
        tag, it = T.com_iter(range(3))
        self.assertEqual([T.com_next(it, 1) for _ in range(4)],
                         [(T.TAG_INT, 0), (T.TAG_INT, 1), (T.TAG_INT, 2), (T.TAG_STOP, None)])
        self.assertEqual(T.com_iter(5)[0], T.TAG_FAIL)

    def test_dates_to_vba_are_serial_numbers(self):
        self.assertEqual(self.boot._com(datetime.date(2023, 3, 15), 1), (self.boot.TAG_FLOAT, 45000.0))


if __name__ == '__main__':
    unittest.main()
