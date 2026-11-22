# Tests of the Python add-in

| Folder | What it checks | Runs |
|---|---|---|
| [`unit`](unit) | `xll/boot.py`, the Python side of the add-in, alone: the decorators, the Python folder, conversions, handles, `PY_CALL`'s lookup, reload; the console (`console.py`) and the studio's server (`xll/studio`): its security, files, checks, completions, console and packages | anywhere, with any Python from 3.11; in CI |
| [`fakehost`](fakehost/README.md) | `asfpy64.xll` and `asfpy32.xll` themselves, driven by a stand-in for Excel: registration, every slot, calls, threads, memory, `py` through `IDispatch`, commands, unticking | Windows natively, Linux under Wine; in CI |
| [`excel`](excel/README.md) | The add-in in Excel: a checklist in VBA, and the Python module it calls | by hand, in each Excel |

By default the fake host and the checklist run each add-in with the Python it carries; both take another Python too (a venv or an install).

```
python -m unittest discover -s excel/python/tests/unit -p "test_*.py"   # boot.py alone
python excel/python/xll/build.py                                      # both add-ins
python excel/python/tests/fakehost/run_fakehost.py                    # the fake host, the add-ins' own Python
```

The GitHub workflow `.github/workflows/python-addin.yml` runs the first two on every change to `excel/python`: MinGW-w64 under Wine on Linux, MSVC natively on Windows, with the add-ins' own Python and with installs of 3.11 to 3.14 (64-bit, and 32-bit 3.12).
