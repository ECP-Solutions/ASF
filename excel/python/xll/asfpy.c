/* asfpy.xll, phase 1a spike: loads the Python of a venv into Excel's process, registers the
   worksheet functions of the Python folder through a pool of exported slots, and gives VBA a
   root object (py) through AsfPy_Root. See excel/python/DESIGN.md and spike/README.md.

   Python is reached only through GetProcAddress: the add-in imports no Python DLL, so one
   build works with every Python from 3.11 of its bitness. Python values cross to C as the
   tagged pairs of boot.py. */
#define _WIN32_WINNT 0x0601
#include <windows.h>
#include <oleauto.h>
#include <stdio.h>
#include <stdarg.h>
#include <stdlib.h>
#include <stddef.h>
#include <string.h>
#include <wchar.h>
#include "xloper.h"
#include "boot_py.h"

#define NSLOTS    256
#define MAX_ARGS  32
#define SETTINGS_KEY L"Software\\VB and VBA Program Settings\\ASF\\AddIn"
#define ADDIN_NAME L"ASF Python (spike)"
#define EXPORT __declspec(dllexport)
#define ASFPY_SCODE 0x800403E9   /* VBA's Err.Number: vbObjectError + 1001 */

enum { TAG_EMPTY, TAG_BOOL, TAG_FLOAT, TAG_INT, TAG_STR, TAG_LIST, TAG_OBJECT, TAG_FAIL, TAG_XLERR, TAG_GRID };

/* --- Python, as far as the add-in uses it ---------------------------------------------------- */

typedef struct _object PyObject;
typedef intptr_t Py_ssize_t;
typedef int PyGILState_STATE;
typedef struct PyInitConfig PyInitConfig;

static struct {
    void (*Py_InitializeEx)(int);
    void (*Py_SetPythonHome)(const wchar_t *);
    PyInitConfig *(*PyInitConfig_Create)(void);
    void (*PyInitConfig_Free)(PyInitConfig *);
    int (*PyInitConfig_GetError)(PyInitConfig *, const char **);
    int (*PyInitConfig_SetInt)(PyInitConfig *, const char *, long long);
    int (*PyInitConfig_SetStr)(PyInitConfig *, const char *, const char *);
    int (*Py_InitializeFromInitConfig)(PyInitConfig *);
    void *(*PyEval_SaveThread)(void);
    PyGILState_STATE (*PyGILState_Ensure)(void);
    void (*PyGILState_Release)(PyGILState_STATE);
    PyObject *(*PyImport_ImportModule)(const char *);
    PyObject *(*PyModule_New)(const char *);
    PyObject *(*PyModule_GetDict)(PyObject *);
    PyObject *(*PyObject_GetAttrString)(PyObject *, const char *);
    PyObject *(*PyObject_CallObject)(PyObject *, PyObject *);
    PyObject *(*PyTuple_New)(Py_ssize_t);
    int (*PyTuple_SetItem)(PyObject *, Py_ssize_t, PyObject *);
    PyObject *(*PyTuple_GetItem)(PyObject *, Py_ssize_t);
    PyObject *(*PyList_New)(Py_ssize_t);
    int (*PyList_SetItem)(PyObject *, Py_ssize_t, PyObject *);
    PyObject *(*PyList_GetItem)(PyObject *, Py_ssize_t);
    Py_ssize_t (*PyList_Size)(PyObject *);
    PyObject *(*PyLong_FromLong)(long);
    PyObject *(*PyLong_FromLongLong)(long long);
    long (*PyLong_AsLong)(PyObject *);
    PyObject *(*PyFloat_FromDouble)(double);
    double (*PyFloat_AsDouble)(PyObject *);
    PyObject *(*PyBool_FromLong)(long);
    PyObject *(*PyUnicode_FromWideChar)(const wchar_t *, Py_ssize_t);
    PyObject *(*PyUnicode_FromString)(const char *);
    wchar_t *(*PyUnicode_AsWideCharString)(PyObject *, Py_ssize_t *);
    void (*PyMem_Free)(void *);
    int (*PyObject_IsTrue)(PyObject *);
    void (*Py_IncRef)(PyObject *);
    void (*Py_DecRef)(PyObject *);
    PyObject *(*PyErr_Occurred)(void);
    void (*PyErr_Print)(void);
    void (*PyErr_Clear)(void);
    PyObject *None;
} P;

static HMODULE g_python;
static PyObject *g_boot;          /* the module of boot.py */
static void *g_mainstate;         /* the thread state saved after start-up */
static int g_started;             /* 1: Python runs; -1: it failed to start */
static wchar_t g_start_error[1024];
static wchar_t g_home[MAX_PATH];  /* kept: Py_SetPythonHome holds the pointer */
static wchar_t g_log[MAX_PATH];

/* --- the log, %TEMP%\asfpy.log ---------------------------------------------------------------- */

static void note(const wchar_t *fmt, ...)
{
    wchar_t text[2048];
    char utf8[6144];
    va_list ap;
    SYSTEMTIME t;
    HANDLE f;
    int n;
    DWORD written;
    if (!g_log[0]) {
        DWORD k = GetTempPathW(MAX_PATH - 16, g_log);
        if (!k) return;
        wcscat(g_log, L"asfpy.log");
    }
    GetLocalTime(&t);
    n = swprintf(text, 64, L"%02d:%02d:%02d ", t.wHour, t.wMinute, t.wSecond);
    va_start(ap, fmt);
    vswprintf(text + n, 2048 - n - 2, fmt, ap);
    va_end(ap);
    wcscat(text, L"\n");
    n = WideCharToMultiByte(CP_UTF8, 0, text, -1, utf8, sizeof utf8, NULL, NULL);
    f = CreateFileW(g_log, FILE_APPEND_DATA, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    if (f == INVALID_HANDLE_VALUE) return;
    if (n > 1) WriteFile(f, utf8, n - 1, &written, NULL);
    CloseHandle(f);
}

static double ms_since(LARGE_INTEGER t0)
{
    LARGE_INTEGER t, f;
    QueryPerformanceCounter(&t);
    QueryPerformanceFrequency(&f);
    return (double)(t.QuadPart - t0.QuadPart) * 1000.0 / (double)f.QuadPart;
}

static void start_failed(const wchar_t *fmt, ...)
{
    va_list ap;
    va_start(ap, fmt);
    vswprintf(g_start_error, 1024, fmt, ap);
    va_end(ap);
    g_started = -1;
    note(L"start failed: %ls", g_start_error);
}

/* --- settings and finding Python ---------------------------------------------------------------- */

static int setting(const wchar_t *name, wchar_t *out, DWORD chars)
{
    DWORD size = chars * sizeof(wchar_t);
    out[0] = 0;
    if (RegGetValueW(HKEY_CURRENT_USER, SETTINGS_KEY, name, RRF_RT_REG_SZ, NULL, out, &size) != ERROR_SUCCESS) {
        out[0] = 0;
        return 0;
    }
    return out[0] != 0;
}

static void trim(wchar_t *s)
{
    size_t n = wcslen(s), i = 0;
    while (n && (s[n - 1] == L' ' || s[n - 1] == L'\t' || s[n - 1] == L'\r' || s[n - 1] == L'\n')) s[--n] = 0;
    while (s[i] == L' ' || s[i] == L'\t') i++;
    if (i) memmove(s, s + i, (n - i + 1) * sizeof(wchar_t));
}

static int file_exists(const wchar_t *path)
{
    DWORD a = GetFileAttributesW(path);
    return a != INVALID_FILE_ATTRIBUTES && !(a & FILE_ATTRIBUTE_DIRECTORY);
}

/* home = the base install named by venv\pyvenv.cfg, or the folder itself when it is an install */
static int python_home(const wchar_t *venv, wchar_t *home, wchar_t *site, wchar_t *exe)
{
    wchar_t cfg[MAX_PATH];
    char line[1024];
    FILE *f;
    swprintf(cfg, MAX_PATH, L"%ls\\pyvenv.cfg", venv);
    if (!file_exists(cfg)) {
        wcsncpy(home, venv, MAX_PATH);
        swprintf(site, MAX_PATH, L"%ls\\Lib\\site-packages", venv);
        swprintf(exe, MAX_PATH, L"%ls\\python.exe", venv);
        return 1;
    }
    home[0] = 0;
    f = _wfopen(cfg, L"rb");
    if (!f) return 0;
    while (fgets(line, sizeof line, f)) {
        char *eq = strchr(line, '=');
        wchar_t key[64], value[MAX_PATH];
        if (!eq) continue;
        *eq = 0;
        MultiByteToWideChar(CP_UTF8, 0, line, -1, key, 64);
        MultiByteToWideChar(CP_UTF8, 0, eq + 1, -1, value, MAX_PATH);
        trim(key);
        trim(value);
        if (_wcsicmp(key, L"home") == 0) wcsncpy(home, value, MAX_PATH);
    }
    fclose(f);
    swprintf(site, MAX_PATH, L"%ls\\Lib\\site-packages", venv);
    swprintf(exe, MAX_PATH, L"%ls\\Scripts\\python.exe", venv);
    return home[0] != 0;
}

/* the versioned DLL of an install: python312.dll, not python3.dll */
static int python_dll(const wchar_t *home, wchar_t *dll)
{
    wchar_t pattern[MAX_PATH];
    WIN32_FIND_DATAW fd;
    HANDLE h;
    int found = 0;
    swprintf(pattern, MAX_PATH, L"%ls\\python3*.dll", home);
    h = FindFirstFileW(pattern, &fd);
    if (h == INVALID_HANDLE_VALUE) return 0;
    do {
        /* python3NN.dll only: not python3.dll, nor a free-threaded build (python3t.dll, python313t.dll) */
        const wchar_t *d = fd.cFileName + 7;
        int digits = 0;
        while (*d >= L'0' && *d <= L'9') { d++; digits++; }
        if (digits < 2 || _wcsicmp(d, L".dll") != 0) continue;
        if (!found || _wcsicmp(fd.cFileName, wcsrchr(dll, L'\\') + 1) > 0) {
            swprintf(dll, MAX_PATH, L"%ls\\%ls", home, fd.cFileName);
            found = 1;
        }
    } while (FindNextFileW(h, &fd));
    FindClose(h);
    return found;
}

/* the machine a DLL is built for, from its PE header */
static WORD dll_machine(const wchar_t *path)
{
    HANDLE f = CreateFileW(path, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, 0, NULL);
    IMAGE_DOS_HEADER dos;
    DWORD sig, got;
    IMAGE_FILE_HEADER fh;
    WORD machine = 0;
    if (f == INVALID_HANDLE_VALUE) return 0;
    if (ReadFile(f, &dos, sizeof dos, &got, NULL) && got == sizeof dos && dos.e_magic == IMAGE_DOS_SIGNATURE &&
        SetFilePointer(f, dos.e_lfanew, NULL, FILE_BEGIN) != INVALID_SET_FILE_POINTER &&
        ReadFile(f, &sig, 4, &got, NULL) && sig == IMAGE_NT_SIGNATURE &&
        ReadFile(f, &fh, sizeof fh, &got, NULL) && got == sizeof fh)
        machine = fh.Machine;
    CloseHandle(f);
    return machine;
}

static const wchar_t *python_in_process(void)
{
    static wchar_t name[32];
    int v;
    if (GetModuleHandleW(L"python3.dll")) return L"python3.dll";
    for (v = 6; v <= 40; v++) {
        swprintf(name, 32, L"python3%d.dll", v);
        if (GetModuleHandleW(name)) return name;
    }
    return NULL;
}

static int sym(void *slot, const char *name)
{
    FARPROC f = GetProcAddress(g_python, name);
    memcpy(slot, &f, sizeof f);
    return f != NULL;
}

#define NEED(f) if (!sym(&P.f, #f)) { start_failed(L"%ls has no %hs", dll, #f); return 0; }
#define MAYBE(f) sym(&P.f, #f)

static int load_api(const wchar_t *dll)
{
    NEED(Py_InitializeEx) NEED(PyEval_SaveThread) NEED(PyGILState_Ensure) NEED(PyGILState_Release)
    NEED(PyImport_ImportModule) NEED(PyModule_New) NEED(PyModule_GetDict) NEED(PyObject_GetAttrString)
    NEED(PyObject_CallObject) NEED(PyTuple_New) NEED(PyTuple_SetItem) NEED(PyTuple_GetItem)
    NEED(PyList_New) NEED(PyList_SetItem) NEED(PyList_GetItem) NEED(PyList_Size)
    NEED(PyLong_FromLong) NEED(PyLong_FromLongLong) NEED(PyLong_AsLong) NEED(PyFloat_FromDouble)
    NEED(PyFloat_AsDouble) NEED(PyBool_FromLong) NEED(PyUnicode_FromWideChar) NEED(PyUnicode_FromString)
    NEED(PyUnicode_AsWideCharString) NEED(PyMem_Free) NEED(PyObject_IsTrue) NEED(Py_IncRef) NEED(Py_DecRef)
    NEED(PyErr_Occurred) NEED(PyErr_Print) NEED(PyErr_Clear)
    MAYBE(Py_SetPythonHome);
    MAYBE(PyInitConfig_Create); MAYBE(PyInitConfig_Free); MAYBE(PyInitConfig_GetError);
    MAYBE(PyInitConfig_SetInt); MAYBE(PyInitConfig_SetStr); MAYBE(Py_InitializeFromInitConfig);
    if (!(P.None = (PyObject *)GetProcAddress(g_python, "_Py_NoneStruct"))) { start_failed(L"%ls has no _Py_NoneStruct", dll); return 0; }
    return 1;
}

static void set_flag(const char *name)
{
    int *flag = (int *)GetProcAddress(g_python, name);
    if (flag) *flag = 1;
}

/* Python 3.14 and later: the configuration API of PEP 741. Before: the global flags, which
   those versions export for good. Either way Python is isolated from the environment and
   doesn't import site: boot.py adds the venv's site-packages itself. */
static int initialize(const char **how)
{
    if (P.PyInitConfig_Create && P.PyInitConfig_SetInt && P.PyInitConfig_SetStr && P.Py_InitializeFromInitConfig) {
        PyInitConfig *cfg = P.PyInitConfig_Create();
        char home[MAX_PATH * 3];
        const char *msg = NULL;
        int bad = !cfg;
        WideCharToMultiByte(CP_UTF8, 0, g_home, -1, home, sizeof home, NULL, NULL);
        if (!bad) bad = P.PyInitConfig_SetStr(cfg, "home", home) < 0
                     || P.PyInitConfig_SetInt(cfg, "site_import", 0) < 0
                     || P.PyInitConfig_SetInt(cfg, "user_site_directory", 0) < 0
                     || P.PyInitConfig_SetInt(cfg, "use_environment", 0) < 0
                     || P.PyInitConfig_SetInt(cfg, "safe_path", 1) < 0
                     || P.PyInitConfig_SetInt(cfg, "install_signal_handlers", 0) < 0
                     || P.Py_InitializeFromInitConfig(cfg) < 0;
        if (bad) {
            if (cfg && P.PyInitConfig_GetError) P.PyInitConfig_GetError(cfg, &msg);
            start_failed(L"Python didn't start: %hs", msg ? msg : "no configuration");
            if (cfg && P.PyInitConfig_Free) P.PyInitConfig_Free(cfg);
            return 0;
        }
        if (P.PyInitConfig_Free) P.PyInitConfig_Free(cfg);
        *how = "PyInitConfig (PEP 741)";
        return 1;
    }
    set_flag("Py_IsolatedFlag");
    set_flag("Py_IgnoreEnvironmentFlag");
    set_flag("Py_NoSiteFlag");
    set_flag("Py_NoUserSiteDirectory");
    if (P.Py_SetPythonHome) P.Py_SetPythonHome(g_home);
    P.Py_InitializeEx(0);
    *how = "global flags and Py_InitializeEx";
    return 1;
}

/* --- calling boot.py ------------------------------------------------------------------------- */

static PyObject *wstr(const wchar_t *s) { return P.PyUnicode_FromWideChar(s, (Py_ssize_t)wcslen(s)); }

/* boot.<name>(args...), args stolen; a new reference or NULL (the error is logged) */
static PyObject *boot_call(const char *name, int n, ...)
{
    PyObject *fn = P.PyObject_GetAttrString(g_boot, name), *args = P.PyTuple_New(n), *r;
    va_list ap;
    int i;
    va_start(ap, n);
    for (i = 0; i < n; i++) P.PyTuple_SetItem(args, i, va_arg(ap, PyObject *));
    va_end(ap);
    r = fn ? P.PyObject_CallObject(fn, args) : NULL;
    if (fn) P.Py_DecRef(fn);
    P.Py_DecRef(args);
    if (!r) {
        note(L"boot.%hs raised", name);
        if (P.PyErr_Occurred()) P.PyErr_Print();
    }
    return r;
}

/* a str as a malloc'd wide string */
static wchar_t *pytext(PyObject *s, Py_ssize_t *len)
{
    Py_ssize_t n = 0;
    wchar_t *w = s ? P.PyUnicode_AsWideCharString(s, &n) : NULL, *copy;
    if (!w) {
        P.PyErr_Clear();
        copy = calloc(1, sizeof(wchar_t));
        if (len) *len = 0;
        return copy;
    }
    copy = malloc((n + 1) * sizeof(wchar_t));
    memcpy(copy, w, (n + 1) * sizeof(wchar_t));
    P.PyMem_Free(w);
    if (len) *len = n;
    return copy;
}

/* --- worksheet functions: registration ----------------------------------------------------- */

typedef struct { wchar_t *name, *args, *help; wchar_t *arghelp[MAX_ARGS]; int nargs; } Func;
static Func g_funcs[NSLOTS];
static int g_nfuncs;
int asfpy_nargs[NSLOTS];          /* read by the 32-bit slots (slots_x86.S) */

static EXCEL12PROC g_excel12;

EXPORT void WINAPI SetExcel12EntryPt(EXCEL12PROC p) { g_excel12 = p; }

static int excel12(int fn, LPXLOPER12 res, int n, LPXLOPER12 *opers)
{
    if (!g_excel12) g_excel12 = (EXCEL12PROC)GetProcAddress(GetModuleHandleW(NULL), "MdCallBack12");
    if (!g_excel12) return -1;
    return g_excel12(fn, n, opers, res);
}

/* an XLOPER12 string of at most 255 characters (the limit of xlfRegister's texts) */
static void xltext(LPXLOPER12 x, const wchar_t *s)
{
    size_t n = wcslen(s);
    if (n > 255) n = 255;
    x->xltype = xltypeStr;
    x->val.str = malloc((n + 1) * sizeof(XCHAR));
    x->val.str[0] = (XCHAR)n;
    memcpy(x->val.str + 1, s, n * sizeof(XCHAR));
}

static void alert(const wchar_t *msg)
{
    XLOPER12 text, res;
    LPXLOPER12 op[1] = { &text };
    xltext(&text, msg);
    excel12(118 | 0x8000 /* xlcAlert */, &res, 1, op);
    free(text.val.str);
}

static void register_functions(void)
{
    XLOPER12 dll, res, x[10 + MAX_ARGS + 1];
    LPXLOPER12 op[10 + MAX_ARGS + 1];
    int i, k, n;
    if (excel12(xlGetName, &dll, 0, NULL) != xlretSuccess) {
        note(L"xlGetName failed");
        return;
    }
    op[0] = &dll;
    for (i = 0; i < g_nfuncs; i++) {
        Func *f = &g_funcs[i];
        wchar_t proc[32], type[MAX_ARGS + 2];
        swprintf(proc, 32, L"asfpy_f%04d", i);
        for (k = 0; k <= f->nargs; k++) type[k] = L'Q';
        type[f->nargs + 1] = 0;
        xltext(&x[1], proc);
        xltext(&x[2], type);
        xltext(&x[3], f->name);
        xltext(&x[4], f->args);
        x[5].xltype = xltypeNum; x[5].val.num = 1;
        xltext(&x[6], L"Python");
        xltext(&x[7], L"");
        xltext(&x[8], L"");
        xltext(&x[9], f->help);
        n = 10;
        for (k = 0; k < f->nargs; k++) xltext(&x[n++], f->arghelp[k]);
        xltext(&x[n++], L"");    /* Insert Function cuts the last description: one more, empty */
        for (k = 1; k < n; k++) op[k] = &x[k];
        asfpy_nargs[i] = f->nargs;
        if (excel12(xlfRegister, &res, n, op) != xlretSuccess || res.xltype != xltypeNum)
            note(L"registering %ls failed", f->name);
        else
            note(L"registered %ls as %ls (%d arguments)", f->name, proc, f->nargs);
        for (k = 1; k < n; k++)
            if (k != 5) free(x[k].val.str);    /* x[5] is the macro type, a number */
    }
    excel12(xlFree, NULL, 1, op);
}

/* boot.start's list of (name, [argument names], description) into g_funcs */
static void read_functions(PyObject *list)
{
    Py_ssize_t i, k, n = P.PyList_Size(list);
    for (i = 0; i < n && g_nfuncs < NSLOTS; i++) {
        PyObject *t = P.PyList_GetItem(list, i), *args = P.PyTuple_GetItem(t, 1);
        Func *f = &g_funcs[g_nfuncs++];
        Py_ssize_t na = P.PyList_Size(args);
        size_t len = 0;
        f->name = pytext(P.PyTuple_GetItem(t, 0), NULL);
        f->help = pytext(P.PyTuple_GetItem(t, 2), NULL);
        f->nargs = (int)na;
        for (k = 0; k < na; k++) {
            f->arghelp[k] = pytext(P.PyList_GetItem(args, k), NULL);
            len += wcslen(f->arghelp[k]) + 1;
        }
        f->args = calloc(len + 1, sizeof(wchar_t));
        for (k = 0; k < na; k++) {
            if (k) wcscat(f->args, L",");
            wcscat(f->args, f->arghelp[k]);
        }
    }
}

/* --- starting Python --------------------------------------------------------------------------- */

static void start_python(void)
{
    wchar_t venv[MAX_PATH], folder[MAX_PATH], site[MAX_PATH], exe[MAX_PATH], dll[MAX_PATH];
    const wchar_t *other;
    const char *how = "";
    WORD want, got;
    LARGE_INTEGER t0, t1;
    PyObject *code, *dict, *builtins, *exec, *args, *r, *list;

    QueryPerformanceCounter(&t0);
    note(L"--- %ls, %d-bit", ADDIN_NAME, (int)sizeof(void *) * 8);
    if (!setting(L"PythonVenv", venv, MAX_PATH)) {
        start_failed(L"no Python is set: run AsfPySetup in VBA, or set PythonVenv under HKCU\\%ls", SETTINGS_KEY);
        return;
    }
    if (!setting(L"PythonFolder", folder, MAX_PATH)) {
        DWORD k = GetEnvironmentVariableW(L"APPDATA", folder, MAX_PATH);
        if (!k || k > MAX_PATH - 20) folder[0] = 0; else wcscat(folder, L"\\ASF\\python");
    }
    if (!python_home(venv, g_home, site, exe)) {
        start_failed(L"%ls\\pyvenv.cfg names no home", venv);
        return;
    }
    if (!python_dll(g_home, dll)) {
        start_failed(L"no python3XX.dll in %ls", g_home);
        return;
    }
    if ((other = python_in_process()) != NULL) {
        start_failed(L"another add-in has already loaded %ls into Excel; one Python per process", other);
        return;
    }
    want = sizeof(void *) == 8 ? IMAGE_FILE_MACHINE_AMD64 : IMAGE_FILE_MACHINE_I386;
    got = dll_machine(dll);
    if (got != want) {
        start_failed(L"%ls is not a %d-bit DLL: Excel and Python must have the same bitness", dll, (int)sizeof(void *) * 8);
        return;
    }
    g_python = LoadLibraryExW(dll, NULL, LOAD_WITH_ALTERED_SEARCH_PATH);
    if (!g_python) {
        start_failed(L"%ls didn't load (error %lu)", dll, GetLastError());
        return;
    }
    if (!load_api(dll)) return;
    note(L"loaded %ls in %.0f ms", dll, ms_since(t0));
    QueryPerformanceCounter(&t1);
    if (!initialize(&how)) return;
    note(L"Python started with %hs in %.0f ms", how, ms_since(t1));

    QueryPerformanceCounter(&t1);
    g_boot = P.PyModule_New("_asfpy_boot");
    dict = g_boot ? P.PyModule_GetDict(g_boot) : NULL;
    builtins = P.PyImport_ImportModule("builtins");
    exec = builtins ? P.PyObject_GetAttrString(builtins, "exec") : NULL;
    code = P.PyUnicode_FromString(BOOT_PY);
    r = NULL;
    if (dict && exec && code) {
        args = P.PyTuple_New(2);
        P.Py_IncRef(dict);
        P.PyTuple_SetItem(args, 0, code);
        P.PyTuple_SetItem(args, 1, dict);
        r = P.PyObject_CallObject(exec, args);
        P.Py_DecRef(args);
    }
    if (!r) {
        if (P.PyErr_Occurred()) P.PyErr_Print();
        start_failed(L"the bootstrap of the add-in failed");
        g_mainstate = P.PyEval_SaveThread();
        return;
    }
    P.Py_DecRef(r);
    list = boot_call("start", 4, wstr(folder), wstr(site), wstr(exe), wstr(g_log));
    if (!list) {
        start_failed(L"the start of the Python folder failed");
        g_mainstate = P.PyEval_SaveThread();
        return;
    }
    read_functions(list);
    P.Py_DecRef(list);
    g_mainstate = P.PyEval_SaveThread();
    g_started = 1;
    note(L"bootstrap and the Python folder (%ls) in %.0f ms; %d functions; start-up %.0f ms",
         folder, ms_since(t1), g_nfuncs, ms_since(t0));
}

/* --- worksheet functions: calls ---------------------------------------------------------------- */

static XLOPER12 g_err[64];

static LPXLOPER12 xlerror(int code)
{
    g_err[code].xltype = xltypeErr;
    g_err[code].val.err = code;
    return &g_err[code];
}

static const wchar_t *error_text(int code)
{
    switch (code) {
    case xlerrNull: return L"#NULL!";
    case xlerrDiv0: return L"#DIV/0!";
    case xlerrValue: return L"#VALUE!";
    case xlerrRef: return L"#REF!";
    case xlerrName: return L"#NAME?";
    case xlerrNum: return L"#NUM!";
    default: return L"#N/A";
    }
}

/* a cell value; an error is its text inside a range */
static PyObject *cell_to_py(LPXLOPER12 x)
{
    switch (x->xltype & 0xFFF) {
    case xltypeNum: return P.PyFloat_FromDouble(x->val.num);
    case xltypeInt: return P.PyLong_FromLong(x->val.w);
    case xltypeStr: return P.PyUnicode_FromWideChar(x->val.str + 1, x->val.str[0]);
    case xltypeBool: return P.PyBool_FromLong(x->val.xbool);
    case xltypeErr: return wstr(error_text(x->val.err));
    default: P.Py_IncRef(P.None); return P.None;
    }
}

static PyObject *multi_to_py(LPXLOPER12 x)
{
    RW r, rows = x->val.array.rows;
    COL c, cols = x->val.array.columns;
    PyObject *list = P.PyList_New(rows);
    for (r = 0; r < rows; r++) {
        PyObject *row = P.PyList_New(cols);
        for (c = 0; c < cols; c++) P.PyList_SetItem(row, c, cell_to_py(&x->val.array.lparray[r * cols + c]));
        P.PyList_SetItem(list, r, row);
    }
    return list;
}

static void set_xlstr(LPXLOPER12 x, PyObject *s)
{
    Py_ssize_t n;
    wchar_t *w = pytext(s, &n);
    if (n > 32767) n = 32767;
    x->xltype = xltypeStr;
    x->val.str = malloc((n + 1) * sizeof(XCHAR));
    x->val.str[0] = (XCHAR)n;
    memcpy(x->val.str + 1, w, n * sizeof(XCHAR));
    free(w);
}

/* a tagged scalar of boot.py into x; strings are malloc'd */
static void tag_to_cell(PyObject *t, LPXLOPER12 x)
{
    long tag = P.PyLong_AsLong(P.PyTuple_GetItem(t, 0));
    PyObject *v = P.PyTuple_GetItem(t, 1);
    switch (tag) {
    case TAG_BOOL: x->xltype = xltypeBool; x->val.xbool = P.PyObject_IsTrue(v); break;
    case TAG_FLOAT: case TAG_INT: x->xltype = xltypeNum; x->val.num = P.PyFloat_AsDouble(v); break;
    case TAG_STR: set_xlstr(x, v); break;
    case TAG_XLERR: x->xltype = xltypeErr; x->val.err = (int)P.PyLong_AsLong(v); break;
    case TAG_FAIL: x->xltype = xltypeErr; x->val.err = xlerrValue; break;
    default: x->xltype = xltypeStr; x->val.str = calloc(1, sizeof(XCHAR)); break;
    }
}

static LPXLOPER12 tag_to_result(PyObject *t)
{
    long tag = P.PyLong_AsLong(P.PyTuple_GetItem(t, 0));
    LPXLOPER12 res = calloc(1, sizeof(XLOPER12));
    if (tag == TAG_GRID) {
        PyObject *rows = P.PyTuple_GetItem(t, 1), *row0 = P.PyList_GetItem(rows, 0);
        RW r, nr = (RW)P.PyList_Size(rows);
        COL c, nc = (COL)P.PyList_Size(row0);
        res->xltype = xltypeMulti;
        res->val.array.rows = nr;
        res->val.array.columns = nc;
        res->val.array.lparray = calloc((size_t)nr * nc, sizeof(XLOPER12));
        for (r = 0; r < nr; r++) {
            PyObject *row = P.PyList_GetItem(rows, r);
            for (c = 0; c < nc; c++) tag_to_cell(P.PyList_GetItem(row, c), &res->val.array.lparray[r * nc + c]);
        }
    } else {
        tag_to_cell(t, res);
    }
    res->xltype |= xlbitDLLFree;
    return res;
}

EXPORT void WINAPI xlAutoFree12(LPXLOPER12 x)
{
    DWORD type = x->xltype & ~xlbitDLLFree;
    if (type == xltypeStr) {
        free(x->val.str);
    } else if (type == xltypeMulti) {
        RW i, n = x->val.array.rows * x->val.array.columns;
        for (i = 0; i < n; i++)
            if (x->val.array.lparray[i].xltype == xltypeStr) free(x->val.array.lparray[i].val.str);
        free(x->val.array.lparray);
    }
    free(x);
}

/* every slot comes here: from C on 64-bit, from the stubs of slots_x86.S on 32-bit */
LPXLOPER12 __cdecl asfpy_dispatch(int slot, LPXLOPER12 *argv)
{
    PyGILState_STATE gil;
    PyObject *args, *r;
    LPXLOPER12 res;
    int i, n;
    if (g_started != 1 || slot >= g_nfuncs) return xlerror(xlerrName);
    n = g_funcs[slot].nargs;
    for (i = 0; i < n; i++)
        if ((argv[i]->xltype & 0xFFF) == xltypeErr) return xlerror(argv[i]->val.err);
    gil = P.PyGILState_Ensure();
    args = P.PyList_New(n);
    for (i = 0; i < n; i++)
        P.PyList_SetItem(args, i, (argv[i]->xltype & 0xFFF) == xltypeMulti ? multi_to_py(argv[i]) : cell_to_py(argv[i]));
    r = boot_call("xl_call", 2, P.PyLong_FromLong(slot), args);
    res = r ? tag_to_result(r) : xlerror(xlerrValue);
    if (r) P.Py_DecRef(r);
    P.PyGILState_Release(gil);
    return res;
}

#ifdef _WIN64
#include "slots_x64.h"
#endif

/* --- the add-in's entry points ----------------------------------------------------------------- */

EXPORT int WINAPI xlAutoOpen(void)
{
    HMODULE self;
    /* pinned: Excel can't unload code that Python and VBA's proxies still use */
    GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_PIN, (LPCWSTR)(void *)xlAutoOpen, &self);
    if (!g_started) start_python();
    if (g_started == 1) {
        register_functions();
        if (g_nfuncs == 0) {
            wchar_t msg[1200];
            swprintf(msg, 1200, L"ASF Python: Python started, but the Python folder gave no worksheet functions. Details: %ls", g_log);
            alert(msg);
        }
    } else {
        wchar_t msg[1200];
        swprintf(msg, 1200, L"ASF Python: %ls. Details: %ls", g_start_error, g_log);
        alert(msg);
    }
    return 1;
}

/* Python is never finalized (numpy and pandas can't start again in one process) */
EXPORT int WINAPI xlAutoClose(void) { return 1; }

EXPORT LPXLOPER12 WINAPI xlAddInManagerInfo12(LPXLOPER12 action)
{
    static XLOPER12 name;
    static XCHAR text[64];
    if ((action->xltype & 0xFFF) == xltypeNum && action->val.num == 1) {
        size_t n = wcslen(ADDIN_NAME);
        text[0] = (XCHAR)n;
        memcpy(text + 1, ADDIN_NAME, n * sizeof(XCHAR));
        name.xltype = xltypeStr;
        name.val.str = text;
        return &name;
    }
    return xlerror(xlerrValue);
}

/* --- the proxy: a Python object as IDispatch ------------------------------------------------ */

/* a proxy is an IDispatch, and also an ISupportErrorInfo: VBA's CallByName takes the message of
   an error only from objects that say they set the thread's error info */
typedef struct { IDispatchVtbl *lpVtbl; ISupportErrorInfoVtbl *lpSeiVtbl; LONG refs; PyObject *obj; int level; } Proxy;
static IDispatchVtbl g_proxy_vtbl;
static ISupportErrorInfoVtbl g_sei_vtbl;
static const GUID IID_IUnknown_ = { 0x00000000, 0, 0, { 0xC0, 0, 0, 0, 0, 0, 0, 0x46 } };
static const GUID IID_IDispatch_ = { 0x00020400, 0, 0, { 0xC0, 0, 0, 0, 0, 0, 0, 0x46 } };
static const GUID IID_ISupportErrorInfo_ = { 0xDF0B3D60, 0x548F, 0x101B, { 0x8E, 0x65, 0x08, 0x00, 0x2B, 0x2B, 0xD1, 0x19 } };

static wchar_t **g_names;         /* DISPID n is g_names[n - 1] */
static int g_nnames, g_cap;
static CRITICAL_SECTION g_names_lock;

static IDispatch *new_proxy(PyObject *obj, int level)  /* obj: a new reference, kept */
{
    Proxy *p = calloc(1, sizeof(Proxy));
    p->lpVtbl = &g_proxy_vtbl;
    p->lpSeiVtbl = &g_sei_vtbl;
    p->refs = 1;
    p->obj = obj;
    p->level = level;
    return (IDispatch *)p;
}

static HRESULT STDMETHODCALLTYPE px_QueryInterface(IDispatch *self, REFIID riid, void **out)
{
    if (IsEqualGUID(riid, &IID_IUnknown_) || IsEqualGUID(riid, &IID_IDispatch_)) {
        *out = self;
        self->lpVtbl->AddRef(self);
        return S_OK;
    }
    if (IsEqualGUID(riid, &IID_ISupportErrorInfo_)) {
        *out = &((Proxy *)self)->lpSeiVtbl;
        self->lpVtbl->AddRef(self);
        return S_OK;
    }
    *out = NULL;
    return E_NOINTERFACE;
}

#define SEI_PROXY(self) ((IDispatch *)((char *)(self) - offsetof(Proxy, lpSeiVtbl)))

static HRESULT STDMETHODCALLTYPE sei_QueryInterface(ISupportErrorInfo *self, REFIID riid, void **out)
{
    IDispatch *d = SEI_PROXY(self);
    return d->lpVtbl->QueryInterface(d, riid, out);
}

static ULONG STDMETHODCALLTYPE sei_AddRef(ISupportErrorInfo *self) { IDispatch *d = SEI_PROXY(self); return d->lpVtbl->AddRef(d); }
static ULONG STDMETHODCALLTYPE sei_Release(ISupportErrorInfo *self) { IDispatch *d = SEI_PROXY(self); return d->lpVtbl->Release(d); }

static HRESULT STDMETHODCALLTYPE sei_InterfaceSupportsErrorInfo(ISupportErrorInfo *self, REFIID riid)
{
    (void)self; (void)riid;
    return S_OK;
}

static ISupportErrorInfoVtbl g_sei_vtbl = { sei_QueryInterface, sei_AddRef, sei_Release, sei_InterfaceSupportsErrorInfo };

static ULONG STDMETHODCALLTYPE px_AddRef(IDispatch *self) { return InterlockedIncrement(&((Proxy *)self)->refs); }

static ULONG STDMETHODCALLTYPE px_Release(IDispatch *self)
{
    Proxy *p = (Proxy *)self;
    LONG n = InterlockedDecrement(&p->refs);
    if (n == 0) {
        PyGILState_STATE gil = P.PyGILState_Ensure();
        P.Py_DecRef(p->obj);
        P.PyGILState_Release(gil);
        free(p);
    }
    return n;
}

static HRESULT STDMETHODCALLTYPE px_GetTypeInfoCount(IDispatch *self, UINT *n) { (void)self; *n = 0; return S_OK; }

static HRESULT STDMETHODCALLTYPE px_GetTypeInfo(IDispatch *self, UINT i, LCID lcid, ITypeInfo **ti)
{
    (void)self; (void)i; (void)lcid;
    *ti = NULL;
    return DISP_E_BADINDEX;
}

/* any name gets an id: it is looked up in Python when it is invoked */
static HRESULT STDMETHODCALLTYPE px_GetIDsOfNames(IDispatch *self, REFIID riid, LPOLESTR *names, UINT n, LCID lcid, DISPID *ids)
{
    int i;
    UINT k;
    (void)self; (void)riid; (void)lcid;
    EnterCriticalSection(&g_names_lock);
    for (i = 0; i < g_nnames; i++)
        if (wcscmp(g_names[i], names[0]) == 0) break;
    if (i == g_nnames) {
        if (g_nnames == g_cap) {
            g_cap = g_cap ? g_cap * 2 : 64;
            g_names = realloc(g_names, g_cap * sizeof(wchar_t *));
        }
        g_names[g_nnames++] = _wcsdup(names[0]);
    }
    ids[0] = i + 1;
    LeaveCriticalSection(&g_names_lock);
    for (k = 1; k < n; k++) ids[k] = DISPID_UNKNOWN;
    return n > 1 ? DISP_E_UNKNOWNNAME : S_OK;
}

static wchar_t g_convert_error[256];

/* a VARIANT as a new Python reference; NULL with g_convert_error set when it can't cross */
static PyObject *variant_to_py(VARIANT *v);

static PyObject *array_to_py(SAFEARRAY *sa)
{
    VARTYPE vt;
    UINT dims = SafeArrayGetDim(sa);
    LONG lo1, hi1, lo2 = 0, hi2 = 0, i, j, idx[2];
    PyObject *out;
    if (dims < 1 || dims > 2 || FAILED(SafeArrayGetVartype(sa, &vt))) {
        wcscpy(g_convert_error, L"only arrays of one or two dimensions can be passed to Python");
        return NULL;
    }
    SafeArrayGetLBound(sa, 1, &lo1);
    SafeArrayGetUBound(sa, 1, &hi1);
    if (dims == 2) {
        SafeArrayGetLBound(sa, 2, &lo2);
        SafeArrayGetUBound(sa, 2, &hi2);
    }
    out = P.PyList_New(hi1 - lo1 + 1);
    for (i = lo1; i <= hi1; i++) {
        PyObject *item;
        if (dims == 1) {
            VARIANT e;
            VariantInit(&e);
            idx[0] = i;
            if (vt == VT_VARIANT) SafeArrayGetElement(sa, idx, &e);
            else if (vt == VT_DECIMAL) { SafeArrayGetElement(sa, idx, &e.decVal); e.vt = VT_DECIMAL; }
            else { SafeArrayGetElement(sa, idx, &e.llVal); e.vt = vt; }
            item = variant_to_py(&e);
            VariantClear(&e);
        } else {
            item = P.PyList_New(hi2 - lo2 + 1);
            for (j = lo2; j <= hi2 && item; j++) {
                VARIANT e;
                PyObject *x;
                VariantInit(&e);
                idx[0] = j;     /* the right-most dimension first */
                idx[1] = i;
                if (vt == VT_VARIANT) SafeArrayGetElement(sa, idx, &e);
                else if (vt == VT_DECIMAL) { SafeArrayGetElement(sa, idx, &e.decVal); e.vt = VT_DECIMAL; }
                else { SafeArrayGetElement(sa, idx, &e.llVal); e.vt = vt; }
                x = variant_to_py(&e);
                VariantClear(&e);
                if (!x) { P.Py_DecRef(item); item = NULL; break; }
                P.PyList_SetItem(item, j - lo2, x);
            }
        }
        if (!item) { P.Py_DecRef(out); return NULL; }
        P.PyList_SetItem(out, i - lo1, item);
    }
    return out;
}

static PyObject *variant_to_py(VARIANT *v)
{
    VARIANT t;
    PyObject *r = NULL;
    if (V_VT(v) & VT_BYREF) {
        VariantInit(&t);
        if (FAILED(VariantCopyInd(&t, v))) { wcscpy(g_convert_error, L"a ByRef argument can't be read"); return NULL; }
        r = variant_to_py(&t);
        VariantClear(&t);
        return r;
    }
    if (V_VT(v) & VT_ARRAY) return array_to_py(V_ARRAY(v));
    switch (V_VT(v)) {
    case VT_EMPTY: case VT_NULL: case VT_ERROR:
        P.Py_IncRef(P.None);
        return P.None;
    case VT_BOOL: return P.PyBool_FromLong(V_BOOL(v) != VARIANT_FALSE);
    case VT_BSTR: return P.PyUnicode_FromWideChar(V_BSTR(v) ? V_BSTR(v) : L"", V_BSTR(v) ? SysStringLen(V_BSTR(v)) : 0);
    case VT_I1: case VT_I2: case VT_I4: case VT_INT: case VT_UI1: case VT_UI2:
        VariantInit(&t);
        VariantChangeType(&t, v, 0, VT_I4);
        return P.PyLong_FromLong(V_I4(&t));
    case VT_I8: return P.PyLong_FromLongLong(V_I8(v));
    case VT_DISPATCH:
        if (V_DISPATCH(v) && V_DISPATCH(v)->lpVtbl == &g_proxy_vtbl) {
            r = ((Proxy *)V_DISPATCH(v))->obj;
            P.Py_IncRef(r);
            return r;
        }
        wcscpy(g_convert_error, L"objects other than Python's can't be passed to Python");
        return NULL;
    default:
        VariantInit(&t);
        if (SUCCEEDED(VariantChangeType(&t, v, 0, VT_R8))) return P.PyFloat_FromDouble(V_R8(&t));
        swprintf(g_convert_error, 256, L"a value of type %d can't be passed to Python", (int)V_VT(v));
        return NULL;
    }
}

/* a tagged value of boot.py into *out; 0 when it is TAG_FAIL (msg: the message, malloc'd) */
static int tag_to_variant(PyObject *t, VARIANT *out, int level, wchar_t **msg)
{
    long tag = P.PyLong_AsLong(P.PyTuple_GetItem(t, 0));
    PyObject *v = P.PyTuple_GetItem(t, 1);
    Py_ssize_t i, n;
    VariantInit(out);
    switch (tag) {
    case TAG_BOOL: V_VT(out) = VT_BOOL; V_BOOL(out) = P.PyObject_IsTrue(v) ? VARIANT_TRUE : VARIANT_FALSE; break;
    case TAG_INT: V_VT(out) = VT_I4; V_I4(out) = P.PyLong_AsLong(v); break;
    case TAG_FLOAT: V_VT(out) = VT_R8; V_R8(out) = P.PyFloat_AsDouble(v); break;
    case TAG_STR: {
        wchar_t *w = pytext(v, &n);
        V_VT(out) = VT_BSTR;
        V_BSTR(out) = SysAllocStringLen(w, (UINT)n);
        free(w);
        break;
    }
    case TAG_LIST: {
        SAFEARRAY *sa;
        VARIANT *data;
        n = P.PyList_Size(v);
        sa = SafeArrayCreateVector(VT_VARIANT, 0, (ULONG)n);
        SafeArrayAccessData(sa, (void **)&data);
        for (i = 0; i < n; i++) tag_to_variant(P.PyList_GetItem(v, i), &data[i], level, NULL);
        SafeArrayUnaccessData(sa);
        V_VT(out) = VT_ARRAY | VT_VARIANT;
        V_ARRAY(out) = sa;
        break;
    }
    case TAG_OBJECT:
        P.Py_IncRef(v);
        V_VT(out) = VT_DISPATCH;
        V_DISPATCH(out) = new_proxy(v, level);
        break;
    case TAG_FAIL:
        if (msg) *msg = pytext(v, NULL);
        return 0;
    default: break;
    }
    return 1;
}

static const GUID IID_IErrorInfo_ = { 0x1CF2B120, 0x547D, 0x101B, { 0x8E, 0x65, 0x08, 0x00, 0x2B, 0x2B, 0xD1, 0x19 } };

/* the message goes to VBA twice: in EXCEPINFO and as the thread's error info, which VBA's
   Err.Description may read instead; and to the log */
static HRESULT fail(EXCEPINFO *ei, const wchar_t *msg)
{
    ICreateErrorInfo *cei;
    note(L"py error: %ls", msg);
    if (SUCCEEDED(CreateErrorInfo(&cei))) {
        IErrorInfo *info;
        cei->lpVtbl->SetSource(cei, (LPOLESTR)L"Python");
        cei->lpVtbl->SetDescription(cei, (LPOLESTR)msg);
        if (SUCCEEDED(cei->lpVtbl->QueryInterface(cei, &IID_IErrorInfo_, (void **)&info))) {
            SetErrorInfo(0, info);
            info->lpVtbl->Release(info);
        }
        cei->lpVtbl->Release(cei);
    }
    if (ei) {
        memset(ei, 0, sizeof *ei);
        ei->bstrSource = SysAllocString(L"Python");
        ei->bstrDescription = SysAllocString(msg);
        ei->scode = ASFPY_SCODE;
    }
    return DISP_E_EXCEPTION;
}

static HRESULT STDMETHODCALLTYPE px_Invoke(IDispatch *self, DISPID id, REFIID riid, LCID lcid, WORD flags,
                                           DISPPARAMS *dp, VARIANT *result, EXCEPINFO *ei, UINT *argerr)
{
    Proxy *p = (Proxy *)self;
    const wchar_t *name = L"";
    PyGILState_STATE gil;
    PyObject *args, *r;
    HRESULT hr = S_OK;
    UINT i, n = dp ? dp->cArgs : 0;
    wchar_t *msg = NULL;
    (void)riid; (void)lcid; (void)argerr;
    SetErrorInfo(0, NULL);    /* no stale message from an earlier call */
    if (flags & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF)) return DISP_E_MEMBERNOTFOUND;
    if (dp && dp->cNamedArgs) return DISP_E_NONAMEDARGS;
    if (id > 0) {
        EnterCriticalSection(&g_names_lock);
        if (id <= g_nnames) name = g_names[id - 1];
        LeaveCriticalSection(&g_names_lock);
    } else if (id != DISPID_VALUE) {
        return DISP_E_MEMBERNOTFOUND;
    }
    gil = P.PyGILState_Ensure();
    args = P.PyList_New(n);
    for (i = 0; i < n; i++) {   /* rgvarg holds the arguments last first */
        PyObject *a = variant_to_py(&dp->rgvarg[n - 1 - i]);
        if (!a) {
            P.Py_DecRef(args);
            P.PyGILState_Release(gil);
            return fail(ei, g_convert_error);
        }
        P.PyList_SetItem(args, i, a);
    }
    P.Py_IncRef(p->obj);
    r = boot_call("com_invoke", 5, p->obj, wstr(name), P.PyLong_FromLong(p->level),
                  P.PyBool_FromLong((flags & DISPATCH_METHOD) != 0), args);
    if (!r) {
        hr = fail(ei, L"the add-in's bootstrap failed: see the log");
    } else {
        VARIANT v;
        if (!tag_to_variant(r, &v, p->level, &msg)) {
            hr = fail(ei, msg);
            free(msg);
        } else if (result) {
            *result = v;
        } else {
            VariantClear(&v);
        }
        P.Py_DecRef(r);
    }
    P.PyGILState_Release(gil);
    return hr;
}

static IDispatchVtbl g_proxy_vtbl = {
    px_QueryInterface, px_AddRef, px_Release, px_GetTypeInfoCount, px_GetTypeInfo, px_GetIDsOfNames, px_Invoke
};

/* VBA: Declare PtrSafe Function AsfPy_Root Lib "asfpy64.xll" (ByVal level As Long, ByRef out As Variant) As Long
   0 and out = py (level 1: py.call; level 2: py.module too); else out = the reason, as text */
EXPORT long WINAPI AsfPy_Root(long level, VARIANT *out)
{
    PyGILState_STATE gil;
    PyObject *r;
    if (!out) return 2;
    VariantClear(out);
    if (g_started != 1) {
        V_VT(out) = VT_BSTR;
        V_BSTR(out) = SysAllocString(g_started ? g_start_error : L"the Python add-in hasn't started");
        return 1;
    }
    gil = P.PyGILState_Ensure();
    r = boot_call("root", 1, P.PyLong_FromLong(level >= 2 ? 2 : 1));
    if (r) {
        V_VT(out) = VT_DISPATCH;
        V_DISPATCH(out) = new_proxy(r, level >= 2 ? 2 : 1);
    }
    P.PyGILState_Release(gil);
    if (!r) {
        V_VT(out) = VT_BSTR;
        V_BSTR(out) = SysAllocString(L"the add-in's bootstrap failed: see the log");
        return 1;
    }
    return 0;
}

BOOL WINAPI DllMain(HINSTANCE h, DWORD reason, LPVOID reserved)
{
    (void)h; (void)reserved;
    if (reason == DLL_PROCESS_ATTACH) InitializeCriticalSection(&g_names_lock);
    return TRUE;
}
