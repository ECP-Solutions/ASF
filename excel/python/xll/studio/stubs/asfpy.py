"""The asfpy module, as the modules of the Python folder see it in Excel. Outside Excel (the
studio's console and completions) the decorators only mark nothing, and xl_app() says where it
works."""


def xl_func(fn=None, *, name=None, volatile=False, thread_safe=False, macro=False, asynchronous=False):
    """A worksheet function. name: its name in Excel (the function's own by default); volatile:
    calculated on every recalculation; thread_safe: may run on Excel's calculation threads; macro:
    a macro sheet equivalent; asynchronous: runs on a thread while Excel calculates other cells (an
    async def function is asynchronous by itself)."""
    return fn if fn is not None else (lambda f: f)


def xl_macro(fn=None, *, name=None):
    """A command: VBA runs it with Application.Run "name", and a button or the ASF tab too. It runs
    on Excel's main thread and may use xl_app()."""
    return fn if fn is not None else (lambda f: f)


def export(fn):
    """A function that PY_CALL and ASF scripts may call, by its module and name."""
    return fn


def xl_app():
    """Excel's Application object, in a command (@xl_macro) only. Needs pywin32 (or comtypes)."""
    raise RuntimeError('xl_app() works inside Excel only, in a command (@xl_macro)')
