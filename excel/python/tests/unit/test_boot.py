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

# imported once, before any test: numpy can't be imported again after a test drops it from sys.modules
try:
    import numpy
except ImportError:
    numpy = None
try:
    import pandas
except ImportError:
    pandas = None

BOOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'xll', 'boot.py')

STATS = '''
import datetime
from asfpy import export, xl_func


@xl_func
def double(x):
    """Twice x.

    More text, which Insert Function doesn't show."""
    return x * 2


@xl_func(name='STATS.ZSCORE')
def zscore(values: list, ddof: int = 0, label=''):
    """How far each value is from the mean, in standard deviations.

    Args:
        values: the numbers, a row or a column
        ddof: delta degrees of freedom,
            0 for a population
    """
    return values


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

OBJECTS = '''
import asyncio
import collections
import time
from asfpy import export, xl_func, xl_macro, xl_app


@xl_func
def counter(text):
    return collections.Counter(text)


@xl_func
def most(c, n: int):
    return c.most_common(n)[0][0]


@xl_func
def as_text(c: str):
    return c


@xl_func
def boxed(x) -> object:
    return [x, x]


@xl_func
def mixed():
    return [[1, collections.OrderedDict()], [2, 'b']]


@xl_func
async def later(x):
    await asyncio.sleep(0.01)
    return x * 2


@xl_func(asynchronous=True)
def slow(x):
    time.sleep(0.01)
    return x + 1


@xl_func(asynchronous=True)
def slow_fails():
    raise KeyError('async on purpose')


@xl_func(asynchronous=True, macro=True)
def async_macro():
    return 1


@xl_macro
def command(a, b):
    return a + b


@xl_macro(name='TryApp')
def try_app():
    return repr(xl_app())


@xl_func
def app_in_function():
    return repr(xl_app())


@export
async def awaited(x):
    await asyncio.sleep(0)
    return x * 3
'''

FRAMES = '''
import numpy
from asfpy import xl_func


@xl_func
def arr_info(a: numpy.ndarray):
    return '%s %s %s' % (a.shape, a.dtype, a.tolist())


@xl_func
def arr_back(n: int):
    return numpy.arange(n * 2, dtype=float).reshape(n, 2)


@xl_func
def arr_column(n: int):
    return numpy.arange(n)


@xl_func
def arr_nan():
    return numpy.array([1.0, float('nan')])


@xl_func
def arr_scalar():
    return numpy.float64(2.5), numpy.int64(3), numpy.bool_(True)


@xl_func
def arr_3d():
    return numpy.zeros((2, 2, 2))
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
        self.write('objects.py', OBJECTS)
        if numpy is not None:
            self.write('frames.py', FRAMES)
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

    def in_cell(self, caller, name, *args):
        # a call from a cell: _native_caller says which
        self.boot._native_caller = lambda: caller
        try:
            return self.call(name, *args)
        finally:
            self.boot._native_caller = None

    def info(self):
        return self.boot.root(1).startInfo


class Tips(BootTest):
    """The tooltip asfpy.c shows as a formula is typed: boot.tip(formula up to the caret, list separator)."""
    def test_the_argument_the_caret_is_in(self):
        self.assertEqual(self.boot.tip('=stats.zscore('),
                         ('STATS.ZSCORE(', 'values', ', [ddof], [label])',
                          'How far each value is from the mean, in standard deviations.',
                          'values: the numbers, a row or a column'))
        tip = self.boot.tip('=SUM(1, STATS.ZSCORE(A1:A9, ')
        self.assertEqual(tip[:3], ('STATS.ZSCORE(values, ', '[ddof]', ', [label])'))
        self.assertEqual(tip[4], 'ddof: delta degrees of freedom, 0 for a population')
        self.assertEqual(self.boot.tip('=STATS.ZSCORE(A1:A9;', ';')[:2], ('STATS.ZSCORE(values; ', '[ddof]'))
        self.assertEqual(self.boot.tip('=double(')[1:], ('x', ')', 'Twice x.', ''))

    def test_texts_sheets_arrays_and_nesting_dont_count_as_arguments(self):
        for formula in ('=STATS.ZSCORE("a, (b", ', "=STATS.ZSCORE('My, sheet'!A1:A3, ", '=STATS.ZSCORE({1,2,3}, ',
                        '=STATS.ZSCORE(SUM(1, 2), ', '=STATS.ZSCORE((A1, A2), ', '=STATS.ZSCORE(Table1[[a],[b]], '):
            self.assertEqual(self.boot.tip(formula)[1], '[ddof]', formula)
        self.assertIsNone(self.boot.tip('=STATS.ZSCORE(1, 2, MAX('))        # Excel's own function: its own help

    def test_a_name_being_typed(self):
        self.assertEqual(self.boot.tip('=stats.z'), ('STATS.ZSCORE(values, [ddof], [label])', '', '',
                                                     'How far each value is from the mean, in standard deviations.', ''))
        self.assertIsNone(self.boot.tip('=stats.'))       # several functions start so: Excel's own list shows them
        self.assertIsNone(self.boot.tip('=s'))

    def test_py_call_shows_the_exported_function(self):
        self.assertEqual(self.boot.tip('=PY_CALL(')[1], 'name')
        self.assertEqual(self.boot.tip('=PY_CALL("stats.scale", A1:A3, ')[:3], ('PY_CALL("stats.scale", values, ', 'factor', ')'))
        self.assertEqual(self.boot.tip('=PY_CALL("nothing.here", ')[1], '[arg1]')

    def test_nothing_outside_the_python_folder(self):
        for formula in ('=SUM(', '=', 'text', '=SUM(1)+', '=STATS.ZSCORE(1, 2, 3, ', '=STATS.ZSCORE(1)'):
            self.assertIsNone(self.boot.tip(formula), formula)


class Start(BootTest):
    def test_functions_their_arguments_descriptions_and_suffixes(self):
        by_name = {f[0]: f for f in self.functions}
        self.assertEqual(by_name['double'], ('double', ['x'], 'Twice x.', '', ['']))
        self.assertEqual(by_name['STATS.ZSCORE'][4], ['the numbers, a row or a column', 'delta degrees of freedom, 0 for a population', ''])
        self.assertEqual(by_name['STATS.HINTS'][4][:3], ['a whole number', 'a date', 'a text'])
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


class Handles(BootTest):
    A1, B1, OTHER = '[Book1.xlsx]Sheet1!R1C1', '[Book1.xlsx]Sheet1!R1C2', '[Book2.xlsx]Sheet1!R1C1'

    def test_an_object_is_a_handle_another_function_takes(self):
        T = self.boot
        tag, text = self.in_cell(self.A1, 'counter', 'abcab')
        self.assertEqual(tag, T.TAG_STR)
        self.assertRegex(text, r'^<Counter#\d+>$')
        self.assertEqual(self.in_cell(self.B1, 'most', text, 1.0), (T.TAG_STR, 'a'))

    def test_a_str_hint_keeps_the_text(self):
        _, text = self.in_cell(self.A1, 'counter', 'ab')
        self.assertEqual(self.in_cell(self.B1, 'as_text', text), (self.boot.TAG_STR, text))

    def test_the_cell_owns_its_handle_and_drops_the_old_one(self):
        T = self.boot
        _, first = self.in_cell(self.A1, 'counter', 'ab')
        _, second = self.in_cell(self.A1, 'counter', 'abb')
        self.assertNotEqual(first, second)
        self.assertEqual(len(T._handles), 1)
        tag, message = self.in_cell(self.B1, 'most', first, 1.0)
        self.assertEqual(tag, T.TAG_FAIL)
        self.assertIn('names no object of this workbook', message)
        self.assertEqual(self.in_cell(self.B1, 'most', second, 1.0), (T.TAG_STR, 'b'))
        self.in_cell(self.A1, 'as_text', 'no object now')
        self.assertEqual(T._handles, {})

    def test_handles_do_not_cross_workbooks(self):
        _, text = self.in_cell(self.A1, 'counter', 'ab')
        tag, message = self.in_cell(self.OTHER, 'most', text, 1.0)
        self.assertEqual(tag, self.boot.TAG_FAIL)
        self.assertIn('LookupError', message)

    def test_an_object_return_hint_and_objects_in_a_grid(self):
        T = self.boot
        tag, text = self.in_cell(self.A1, 'boxed', 2.0)
        self.assertRegex(text, r'^<list#\d+>$')
        tag, rows = self.in_cell(self.B1, 'mixed')
        self.assertEqual(tag, T.TAG_GRID)
        self.assertRegex(rows[0][1][1], r'^<OrderedDict#\d+>$')
        self.assertEqual(len(T._owned[self.B1]), 1)

    def test_no_handles_no_caller(self):
        # a call that makes no handle, while none exists, never asks Excel for its cell
        self.boot._native_caller = lambda: self.fail('the caller was read')
        try:
            self.assertEqual(self.call('double', 2.0), (self.boot.TAG_FLOAT, 4.0))
        finally:
            self.boot._native_caller = None


class Async(BootTest):
    def wait(self, token):
        import time
        for _ in range(500):
            if token in self.boot._async_results:
                return self.boot._async_results.pop(token)
            time.sleep(0.01)
        self.fail('no result for %r' % token)

    def test_flags_and_the_refused_one(self):
        by_name = {f[0]: f for f in self.functions}
        self.assertEqual(by_name['later'][3], '&')
        self.assertEqual(by_name['slow'][3], '&')
        self.assertNotIn('async_macro', by_name)
        self.assertIn("objects.async_macro: asynchronous and macro can't go together", self.info())

    def test_async_def_and_a_thread(self):
        T = self.boot
        T.xl_call_async(self.slot('later'), [4.0], 'one')
        T.xl_call_async(self.slot('slow'), [4.0], 'two')
        self.assertEqual(self.wait('one'), (T.TAG_FLOAT, 8.0))
        self.assertEqual(self.wait('two'), (T.TAG_FLOAT, 5.0))

    def test_an_exception_comes_back_too(self):
        T = self.boot
        T.xl_call_async(self.slot('slow_fails'), [], 'three')
        self.assertEqual(self.wait('three'), (T.TAG_FAIL, "KeyError: 'async on purpose'"))
        self.assertEqual(T.root(1).lastError, "KeyError: 'async on purpose'")

    def test_py_call_and_pycall_wait_for_a_coroutine(self):
        T = self.boot
        self.assertEqual(T.root(1).call('objects.awaited', 2), 6)
        self.assertEqual(T.xl_call_export('objects.awaited', [2.0]), (T.TAG_FLOAT, 6.0))


class Commands(BootTest):
    def test_a_command_is_registered_and_runs(self):
        by_name = {f[0]: f for f in self.functions}
        self.assertEqual(by_name['command'], ('command', ['a', 'b'], '', '%', ['', '']))
        self.assertIsNone(self.boot.tip('=command('))          # a command isn't typed in a formula
        self.assertEqual(by_name['TryApp'][3], '%')
        self.assertEqual(self.call('command', 2.0, 3.0), (self.boot.TAG_FLOAT, 5.0))

    def test_xl_app_in_a_command_only(self):
        T = self.boot
        T._native_app = lambda: 0
        T._native_release = lambda address: None
        try:
            tag, message = self.call('app_in_function')
            self.assertEqual(tag, T.TAG_FAIL)
            self.assertIn('works in a command (@xl_macro) only', message)
            tag, message = self.call('TryApp')
            try:
                import win32com  # noqa: F401
            except ImportError:
                try:
                    import comtypes  # noqa: F401
                except ImportError:
                    self.assertIn('needs pywin32', message)
            self.assertEqual(T._macro_depth, 0)
        finally:
            T._native_app = T._native_release = None

    def test_outside_excel(self):
        tag, message = self.call('TryApp')
        self.assertIn('works inside Excel only', message)

    def test_the_list_of_macros(self):
        tag, text = self.boot.macro_list()
        self.assertEqual(tag, self.boot.TAG_STR)
        lines = text.split('\n')
        self.assertIn('command\t2\t', lines)
        self.assertIn('TryApp\t0\t', lines)
        self.assertFalse(any(line.startswith('add\t') for line in lines))   # worksheet functions aren't listed


class RunScript(BootTest):
    def script(self, name, text):
        folder = os.path.join(self.tmp, 'scripts')
        os.makedirs(folder, exist_ok=True)
        path = os.path.join(folder, name)
        with open(path, 'w', encoding='utf-8') as f:
            f.write(textwrap.dedent(text))
        return path

    def test_runs_as_main_with_its_folder_on_the_path(self):
        self.script('neighbour.py', 'WORD = "next door"\n')
        path = self.script('report.py', '''
            import sys
            import neighbour
            print(__name__, neighbour.WORD, len(sys.argv))
            sys.stderr.write('a warning\\n')
            ''')
        out, err, before = sys.stdout, sys.stderr, sys.path[:]
        tag, text = self.boot.run_script(path)
        self.assertEqual(tag, self.boot.TAG_STR)
        self.assertEqual(text, '__main__ next door 1\na warning\n')
        self.assertEqual((sys.stdout, sys.stderr, sys.path), (out, err, before))
        self.assertEqual(self.boot._macro_depth, 0)

    def test_a_command_so_xl_app_is_allowed(self):
        path = self.script('app.py', '''
            import asfpy
            try:
                asfpy.xl_app()
            except Exception as exc:
                print(exc)
            ''')
        tag, text = self.boot.run_script(path)
        self.assertIn('works inside Excel only', text)     # not 'works in a command only'

    def test_an_exception_gives_its_traceback(self):
        path = self.script('fails.py', 'print("before")\nraise ValueError("on purpose")\n')
        tag, text = self.boot.run_script(path)
        self.assertTrue(text.startswith('before\n'))
        self.assertIn('ValueError: on purpose', text)
        self.assertEqual(self.boot.root(1).lastError, 'ValueError: on purpose')

    def test_exit_codes_and_not_a_script(self):
        self.assertEqual(self.boot.run_script(self.script('quiet.py', 'raise SystemExit\n'))[1], '')
        self.assertEqual(self.boot.run_script(self.script('three.py', 'raise SystemExit(3)\n'))[1], '\nexit code 3')
        self.assertIn('not a .py file', self.boot.run_script(os.path.join(self.tmp, 'missing.py'))[1])
        self.assertIn('not a .py file', self.boot.run_script(42)[1])


class FakeFrame:
    """Stands in for pandas in _kind: only the module and the name are read."""


class Frames(BootTest):
    def setUp(self):
        if numpy is None:
            self.skipTest('numpy is not installed')
        super().setUp()

    def test_a_range_as_an_array(self):
        T = self.boot
        self.assertEqual(self.call('arr_info', [[1.0, 2.0], [None, 4.0]]), (T.TAG_STR, '(2, 2) float64 [[1.0, 2.0], [nan, 4.0]]'))
        self.assertEqual(self.call('arr_info', [[1.0], [2.0]]), (T.TAG_STR, '(2,) float64 [1.0, 2.0]'))
        self.assertEqual(self.call('arr_info', [['a', 1.0]]), (T.TAG_STR, "(2,) object ['a', 1.0]"))
        self.assertEqual(self.call('arr_info', 5.0), (T.TAG_STR, '(1,) float64 [5.0]'))

    def test_arrays_back(self):
        T = self.boot
        self.assertEqual(self.call('arr_back', 2.0), (T.TAG_GRID, [[(T.TAG_FLOAT, 0.0), (T.TAG_FLOAT, 1.0)],
                                                                   [(T.TAG_FLOAT, 2.0), (T.TAG_FLOAT, 3.0)]]))
        self.assertEqual(self.call('arr_column', 2.0), (T.TAG_GRID, [[(T.TAG_FLOAT, 0.0)], [(T.TAG_FLOAT, 1.0)]]))
        self.assertEqual(self.call('arr_nan')[1][1], [(T.TAG_XLERR, T.XLERR_NUM)])
        self.assertEqual(self.call('arr_scalar'), (T.TAG_GRID, [[(T.TAG_FLOAT, 2.5)], [(T.TAG_FLOAT, 3.0)], [(T.TAG_BOOL, True)]]))
        tag, message = self.call('arr_3d')
        self.assertIn('3-dimensional', message)

    def test_hints_name_the_classes_without_importing_them(self):
        T = self.boot
        FakeFrame.__module__ = 'pandas.core.frame'
        FakeFrame.__name__ = 'DataFrame'
        self.assertEqual(T._hint(FakeFrame), 'DataFrame')
        self.assertEqual(T._hint('pd.Series'), 'Series')
        self.assertEqual(T._hint('numpy.ndarray'), 'ndarray')
        self.assertEqual(T._hint('list[float]'), list)
        import numpy.typing
        self.assertEqual(T._hint(numpy.typing.NDArray[numpy.float64]), 'ndarray')


class Pandas(BootTest):
    def setUp(self):
        if pandas is None:
            self.skipTest('pandas is not installed')
        super().setUp()

    def test_data_frames_both_ways(self):
        T = self.boot
        frame = T._data_frame([['x', 'y'], [1.0, 'a'], [2.0, None]])
        self.assertEqual(list(frame.columns), ['x', 'y'])
        self.assertEqual(len(frame), 2)
        self.assertEqual(T._result(frame), (T.TAG_GRID, [[(T.TAG_STR, 'x'), (T.TAG_STR, 'y')],
                                                         [(T.TAG_FLOAT, 1.0), (T.TAG_STR, 'a')],
                                                         [(T.TAG_FLOAT, 2.0), (T.TAG_EMPTY, None)]]))
        named = pandas.DataFrame({'v': [1.5, float('nan')]}, index=pandas.Index(['a', 'b'], name='k'))
        self.assertEqual(T._result(named), (T.TAG_GRID, [[(T.TAG_STR, 'k'), (T.TAG_STR, 'v')],
                                                         [(T.TAG_STR, 'a'), (T.TAG_FLOAT, 1.5)],
                                                         [(T.TAG_STR, 'b'), (T.TAG_EMPTY, None)]]))
        self.assertEqual(T._result(pandas.Series([1, 2])), (T.TAG_GRID, [[(T.TAG_FLOAT, 1.0)], [(T.TAG_FLOAT, 2.0)]]))
        stamp = T._result(pandas.Series([pandas.Timestamp('2023-03-15'), pandas.NaT]))
        self.assertEqual(stamp, (T.TAG_GRID, [[(T.TAG_FLOAT, 45000.0)], [(T.TAG_EMPTY, None)]]))


if __name__ == '__main__':
    unittest.main()
