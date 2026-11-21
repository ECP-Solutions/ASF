/* asfpy.xll: loads Python into Excel's process, registers the worksheet functions of the Python
   folder through a pool of exported slots, and gives VBA a root object (py) through AsfPy_Root.
   See excel/python/DESIGN.md and spike/README.md.

   The Python is the one the add-in carries, unpacked once under %LOCALAPPDATA%\ASF\runtime,
   unless the setting PythonVenv names another.

   Python is reached only through GetProcAddress: the add-in imports no Python DLL, so one
   build works with every Python from 3.11 of its bitness. Python values cross to C as the
   tagged pairs of boot.py. */
#define _WIN32_WINNT 0x0601
#include <windows.h>
#include <oleauto.h>
#include <tlhelp32.h>
#include <stdio.h>
#include <stdarg.h>
#include <stdlib.h>
#include <stddef.h>
#include <string.h>
#include <wchar.h>
#include "xloper.h"
#include "boot_py.h"
#ifdef ASFPY_RUNTIME          /* build.py: the add-in carries a Python, in the resource ASFPY_RUNTIME */
#include <fdi.h>
#include "runtime.h"
#endif

#define NSLOTS    1000       /* asfpy_f0000 to asfpy_f0999, as build.py generates them */
#define MAX_ARGS  32
#define SETTINGS_KEY L"Software\\VB and VBA Program Settings\\ASF\\AddIn"
#define ADDIN_NAME L"ASF Python"
#ifdef ASFPY_DEF             /* MSVC: a .def file exports the plain names */
#define EXPORT
#else
#define EXPORT __declspec(dllexport)
#endif
#define ASFPY_SCODE 0x800403E9   /* VBA's Err.Number: vbObjectError + 1001 */

enum { TAG_EMPTY, TAG_BOOL, TAG_FLOAT, TAG_INT, TAG_STR, TAG_LIST, TAG_OBJECT, TAG_FAIL, TAG_XLERR, TAG_GRID, TAG_STOP };

/* --- Python, as far as the add-in uses it ---------------------------------------------------- */

typedef struct _object PyObject;
typedef intptr_t Py_ssize_t;
typedef int PyGILState_STATE;
typedef struct PyInitConfig PyInitConfig;
typedef PyObject *(*PyCFunction)(PyObject *, PyObject *);
typedef struct { const char *ml_name; PyCFunction ml_meth; int ml_flags; const char *ml_doc; } PyMethodDef;
#define METH_VARARGS 0x0001
#define METH_NOARGS  0x0004

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
    void (*PyEval_RestoreThread)(void *);
    PyObject *(*PyCFunction_NewEx)(PyMethodDef *, PyObject *, PyObject *);
    int (*PyObject_SetAttrString)(PyObject *, const char *, PyObject *);
    long long (*PyLong_AsLongLong)(PyObject *);
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
static wchar_t g_exe[MAX_PATH];   /* the python.exe of the Python, for pip */

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

/* --- the Python the add-in carries --------------------------------------------------------------- */

#ifdef ASFPY_RUNTIME
/* FDI (cabinet.dll) reads the cabinet from the resource and writes the files. A stream is the
   cabinet in memory, or a file being written. */
typedef struct { const BYTE *mem; DWORD size, pos; HANDLE file; } Stream;
static const BYTE *g_cab;
static DWORD g_cabsize;
static wchar_t g_unpack[MAX_PATH];

static FNALLOC(fdi_alloc) { return malloc(cb); }
static FNFREE(fdi_free) { free(pv); }

static FNOPEN(fdi_open)
{
    Stream *s = calloc(1, sizeof *s);
    (void)pszFile; (void)oflag; (void)pmode;   /* the only file FDI opens is the cabinet */
    s->mem = g_cab;
    s->size = g_cabsize;
    return (INT_PTR)s;
}

static FNREAD(fdi_read)
{
    Stream *s = (Stream *)hf;
    DWORD got = 0;
    if (s->file) return ReadFile(s->file, pv, cb, &got, NULL) ? got : (UINT)-1;
    if (cb > s->size - s->pos) cb = s->size - s->pos;
    memcpy(pv, s->mem + s->pos, cb);
    s->pos += cb;
    return cb;
}

static FNWRITE(fdi_write)
{
    Stream *s = (Stream *)hf;
    DWORD put = 0;
    return s->file && WriteFile(s->file, pv, cb, &put, NULL) ? put : (UINT)-1;
}

static FNCLOSE(fdi_close)
{
    Stream *s = (Stream *)hf;
    if (s->file) CloseHandle(s->file);
    free(s);
    return 0;
}

static FNSEEK(fdi_seek)
{
    Stream *s = (Stream *)hf;
    LONG base = seektype == SEEK_CUR ? (LONG)s->pos : seektype == SEEK_END ? (LONG)s->size : 0;
    if (s->file) return SetFilePointer(s->file, dist, NULL, seektype);
    s->pos = (DWORD)(base + dist);
    return s->pos;
}

static FNFDINOTIFY(fdi_notify)
{
    wchar_t path[MAX_PATH], *p;
    size_t n;
    Stream *s;
    switch (fdint) {
    case fdintCOPY_FILE:
        n = wcslen(g_unpack);
        swprintf(path, MAX_PATH, L"%ls\\", g_unpack);
        if (!MultiByteToWideChar(CP_UTF8, 0, pfdin->psz1, -1, path + n + 1, MAX_PATH - (int)n - 1)) return -1;
        for (p = path + n + 1; *p; p++)
            if (*p == L'\\') { *p = 0; CreateDirectoryW(path, NULL); *p = L'\\'; }
        s = calloc(1, sizeof *s);
        s->file = CreateFileW(path, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
        if (s->file == INVALID_HANDLE_VALUE) { free(s); return -1; }
        return (INT_PTR)s;
    case fdintCLOSE_FILE_INFO:
        fdi_close(pfdin->hf);
        return TRUE;
    case fdintNEXT_CABINET:
        return -1;
    default:
        return 0;
    }
}

static void remove_tree(const wchar_t *dir)
{
    wchar_t pattern[MAX_PATH], path[MAX_PATH];
    WIN32_FIND_DATAW fd;
    HANDLE h;
    swprintf(pattern, MAX_PATH, L"%ls\\*", dir);
    h = FindFirstFileW(pattern, &fd);
    if (h != INVALID_HANDLE_VALUE) {
        do {
            if (!wcscmp(fd.cFileName, L".") || !wcscmp(fd.cFileName, L"..")) continue;
            swprintf(path, MAX_PATH, L"%ls\\%ls", dir, fd.cFileName);
            if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) remove_tree(path); else DeleteFileW(path);
        } while (FindNextFileW(h, &fd));
        FindClose(h);
    }
    RemoveDirectoryW(dir);
}

/* %LOCALAPPDATA%\ASF\runtime\RUNTIME_FOLDER, unpacked the first time: into a temporary folder,
   renamed when complete, so a folder of that name is always whole. Pip installs packages there. */
static int unpack_runtime(wchar_t *home)
{
    HMODULE self = NULL;
    HRSRC res;
    HANDLE mutex;
    HFDI fdi;
    ERF erf;
    LARGE_INTEGER t0;
    wchar_t base[MAX_PATH];
    char cab[] = "runtime.cab", path[] = "";
    int ok;
    DWORD k = GetEnvironmentVariableW(L"LOCALAPPDATA", base, MAX_PATH);
    memset(&erf, 0, sizeof erf);
    if (!k || k > MAX_PATH - 80) {
        start_failed(L"%%LOCALAPPDATA%% isn't set");
        return 0;
    }
    wcscat(base, L"\\ASF");
    CreateDirectoryW(base, NULL);
    wcscat(base, L"\\runtime");
    CreateDirectoryW(base, NULL);
    swprintf(home, MAX_PATH, L"%ls\\%ls", base, RUNTIME_FOLDER);
    if (GetFileAttributesW(home) != INVALID_FILE_ATTRIBUTES) return 1;

    mutex = CreateMutexW(NULL, FALSE, L"Local\\ASF.Python.Runtime");   /* two Excels starting at once */
    if (mutex) WaitForSingleObject(mutex, INFINITE);
    ok = GetFileAttributesW(home) != INVALID_FILE_ATTRIBUTES;
    if (!ok) {
        QueryPerformanceCounter(&t0);
        GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                           (LPCWSTR)(void *)unpack_runtime, &self);
        res = FindResourceW(self, L"ASFPY_RUNTIME", (LPCWSTR)RT_RCDATA);
        g_cab = res ? LockResource(LoadResource(self, res)) : NULL;
        g_cabsize = res ? SizeofResource(self, res) : 0;
        swprintf(g_unpack, MAX_PATH, L"%ls.part%lu", home, GetCurrentProcessId());
        remove_tree(g_unpack);
        CreateDirectoryW(g_unpack, NULL);
        fdi = g_cab ? FDICreate(fdi_alloc, fdi_free, fdi_open, fdi_read, fdi_write, fdi_close, fdi_seek, cpuUNKNOWN, &erf) : NULL;
        ok = fdi && FDICopy(fdi, cab, path, 0, fdi_notify, NULL, NULL);
        if (fdi) FDIDestroy(fdi);
        if (ok && !MoveFileExW(g_unpack, home, 0)) ok = GetFileAttributesW(home) != INVALID_FILE_ATTRIBUTES;
        if (ok) note(L"unpacked the Python of the add-in into %ls in %.0f ms", home, ms_since(t0));
        else start_failed(L"the Python of the add-in didn't unpack into %ls (%ls, FDI error %d)", g_unpack,
                          g_cab ? L"cabinet" : L"no resource", erf.erfOper);
        remove_tree(g_unpack);
    }
    if (mutex) { ReleaseMutex(mutex); CloseHandle(mutex); }
    return ok;
}
#endif

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

/* the path of another copy of this add-in in Excel's process (one that exports AsfPy_Root), or
   NULL. Unticking the add-in can't unload it, so browsing to a second copy finds the first. */
static const wchar_t *other_copy(void)
{
    static wchar_t path[MAX_PATH];
    HMODULE self = NULL;
    MODULEENTRY32W me;
    HANDLE snap;
    BOOL more;
    int found = 0;
    GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                       (LPCWSTR)(void *)other_copy, &self);
    snap = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, 0);
    if (snap == INVALID_HANDLE_VALUE) return NULL;
    me.dwSize = sizeof me;
    for (more = Module32FirstW(snap, &me); more && !found; more = Module32NextW(snap, &me))
        if (me.hModule != self && GetProcAddress(me.hModule, "AsfPy_Root")) {
            wcsncpy(path, me.szExePath, MAX_PATH - 1);
            found = 1;
        }
    CloseHandle(snap);
    return found ? path : NULL;
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
    NEED(Py_InitializeEx) NEED(PyEval_SaveThread) NEED(PyEval_RestoreThread) NEED(PyCFunction_NewEx)
    NEED(PyObject_SetAttrString) NEED(PyLong_AsLongLong) NEED(PyGILState_Ensure) NEED(PyGILState_Release)
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

typedef struct {
    wchar_t *name, *args, *help;
    wchar_t *arghelp[MAX_ARGS];
    int nargs;
    wchar_t flags[4];             /* the type's suffixes: ! volatile, # macro equivalent, $ thread-safe */
    int async;                    /* asynchronous: Excel passes a handle after the arguments (X) */
    int command;                  /* a command (@xl_macro), registered with macro type 2 */
    double id;                    /* xlfRegister's result while registered, else 0 */
} Func;
static Func g_funcs[NSLOTS];
static int g_nfuncs;              /* functions in slots */
static int g_total;               /* worksheet functions the Python folder gave, in slots or not */
int asfpy_nargs[NSLOTS];          /* read by the 32-bit slots */

static EXCEL12PROC g_excel12;
static const GUID IID_IDispatch_ = { 0x00020400, 0, 0, { 0xC0, 0, 0, 0, 0, 0, 0, 0x46 } };
static const GUID IID_NULL_ = { 0, 0, 0, { 0, 0, 0, 0, 0, 0, 0, 0 } };

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

/* xlfRegister of one procedure of the add-in (macro type 1: a worksheet function; 2: a command);
   its register ID, or 0 */
static double register_proc(LPXLOPER12 dll, const wchar_t *proc, const wchar_t *type, const wchar_t *name,
                            const wchar_t *args, int macro_type, const wchar_t *help, int nhelp, wchar_t **arghelp)
{
    XLOPER12 res, x[10 + MAX_ARGS + 1];
    LPXLOPER12 op[10 + MAX_ARGS + 1];
    double id = 0;
    int k, n;
    op[0] = dll;
    xltext(&x[1], proc);
    xltext(&x[2], type);
    xltext(&x[3], name);
    xltext(&x[4], args);
    x[5].xltype = xltypeNum;
    x[5].val.num = macro_type;
    xltext(&x[6], L"Python");
    xltext(&x[7], L"");
    xltext(&x[8], L"");
    xltext(&x[9], help);
    n = 10;
    for (k = 0; k < nhelp && k < MAX_ARGS; k++) xltext(&x[n++], arghelp[k]);
    xltext(&x[n++], L"");    /* Insert Function cuts the last description: one more, empty */
    for (k = 1; k < n; k++) op[k] = &x[k];
    if (excel12(xlfRegister, &res, n, op) == xlretSuccess && res.xltype == xltypeNum) {
        id = res.val.num;
        note(L"registered %ls as %ls, type %ls", name, proc, type);
    } else {
        note(L"registering %ls failed", name);
    }
    for (k = 1; k < n; k++)
        if (k != 5) free(x[k].val.str);    /* x[5] is the macro type, a number */
    return id;
}

/* the functions of the Python folder, one in each slot */
static void register_functions(void)
{
    XLOPER12 dll;
    LPXLOPER12 op[1] = { &dll };
    int i, k;
    if (excel12(xlGetName, &dll, 0, NULL) != xlretSuccess) {
        note(L"xlGetName failed");
        return;
    }
    for (i = 0; i < g_nfuncs; i++) {
        Func *f = &g_funcs[i];
        wchar_t proc[32], type[MAX_ARGS + 8];
        swprintf(proc, 32, L"asfpy_f%04d", i);
        /* QQ...: the result and each argument; an asynchronous function returns nothing (>) and
           gets Excel's handle last (X); a command has no suffixes */
        type[0] = f->async ? L'>' : L'Q';
        for (k = 1; k <= f->nargs; k++) type[k] = L'Q';
        if (f->async) type[k++] = L'X';
        wcscpy(type + k, f->command ? L"" : f->flags);
        asfpy_nargs[i] = f->nargs + f->async;     /* what the 32-bit slot removes from the stack */
        f->id = register_proc(&dll, proc, type, f->name, f->args, f->command ? 2 : 1, f->help, f->nargs, f->arghelp);
    }
    excel12(xlFree, NULL, 1, op);
}

/* PY_CALL and the commands AsfPyReload and AsfPyPackages: the add-in's own, registered once
   Python runs */
#define PYCALL_ARGS 29          /* PY_CALL(name, 29 arguments): 30 in all, below Excel's 255 */
static double g_pycall_id, g_reload_id, g_packages_id;

static void register_builtins(void)
{
    XLOPER12 dll;
    LPXLOPER12 op[1] = { &dll };
    wchar_t type[PYCALL_ARGS + 3], args[256] = L"name";
    wchar_t *help[2] = { L"module.function: a function of the Python folder marked with @export",
                         L"an argument of the function; a range arrives as a list" };
    int k;
    if (excel12(xlGetName, &dll, 0, NULL) != xlretSuccess) return;
    for (k = 0; k < PYCALL_ARGS + 2; k++) type[k] = L'Q';
    type[PYCALL_ARGS + 2] = 0;
    for (k = 1; k <= PYCALL_ARGS; k++) swprintf(args + wcslen(args), 256 - wcslen(args), L",arg%d", k);
    g_pycall_id = register_proc(&dll, L"asfpy_call", type, L"PY_CALL", args, 1,
                                L"Calls a function of the Python folder exported with @export, by its module and name.", 2, help);
    g_reload_id = register_proc(&dll, L"asfpy_reload", L"J", L"AsfPyReload", L"", 2,
                                L"Imports the modules of the Python folder again and registers their functions.", 0, NULL);
    g_packages_id = register_proc(&dll, L"asfpy_packages", L"J", L"AsfPyPackages", L"", 2,
                                  L"Opens a console window with pip for the Python of the add-in.", 0, NULL);
    excel12(xlFree, NULL, 1, op);
}

static int unregister_id(double *id)
{
    XLOPER12 x, res;
    LPXLOPER12 op[1] = { &x };
    int ok;
    if (!*id) return 0;
    x.xltype = xltypeNum;
    x.val.num = *id;
    ok = excel12(xlfUnregister, &res, 1, op) == xlretSuccess;
    *id = 0;
    return ok;
}

/* the functions of the slots; Excel no longer calls them, and their cells show #VALUE! */
static int unregister_functions(void)
{
    int i, n = 0;
    for (i = 0; i < g_nfuncs; i++) n += unregister_id(&g_funcs[i].id);
    return n;
}

/* g_funcs empty again, for the functions of a reload */
static void free_functions(void)
{
    int i, k;
    for (i = 0; i < g_nfuncs; i++) {
        Func *f = &g_funcs[i];
        free(f->name);
        free(f->args);
        free(f->help);
        for (k = 0; k < f->nargs; k++) free(f->arghelp[k]);
        memset(f, 0, sizeof *f);
    }
    g_nfuncs = g_total = 0;
}

/* boot.start's list of (name, [argument names], description, suffixes) into g_funcs */
static void read_functions(PyObject *list)
{
    Py_ssize_t i, k, n = P.PyList_Size(list);
    g_total = (int)n;
    for (i = 0; i < n && g_nfuncs < NSLOTS; i++) {
        PyObject *t = P.PyList_GetItem(list, i), *args = P.PyTuple_GetItem(t, 1);
        Func *f = &g_funcs[g_nfuncs++];
        Py_ssize_t na = P.PyList_Size(args);
        wchar_t *flags = pytext(P.PyTuple_GetItem(t, 3), NULL), *c;
        size_t len = 0;
        f->name = pytext(P.PyTuple_GetItem(t, 0), NULL);
        f->help = pytext(P.PyTuple_GetItem(t, 2), NULL);
        f->nargs = (int)na;
        for (c = flags, k = 0; *c && k < 3; c++)
            if (wcschr(L"!#$", *c)) f->flags[k++] = *c;
        f->flags[k] = 0;
        f->async = wcschr(flags, L'&') != NULL;      /* boot.py's marks, not Excel's suffixes */
        f->command = wcschr(flags, L'%') != NULL;
        free(flags);
        for (k = 0; k < na; k++) {
            f->arghelp[k] = pytext(P.PyList_GetItem(args, k), NULL);
            len += wcslen(f->arghelp[k]) + 1;
        }
        /* Excel takes at most 255 characters of argument names: whole names only */
        f->args = calloc(len + 1, sizeof(wchar_t));
        for (k = 0; k < na; k++) {
            if (wcslen(f->args) + (k ? 1 : 0) + wcslen(f->arghelp[k]) > 255) {
                note(L"%ls: the argument names are longer than 255 characters; Excel shows the first %d", f->name, (int)k);
                break;
            }
            if (k) wcscat(f->args, L",");
            wcscat(f->args, f->arghelp[k]);
        }
    }
    if (g_total > g_nfuncs) {
        wchar_t *first = pytext(P.PyTuple_GetItem(P.PyList_GetItem(list, g_nfuncs), 0), NULL);
        note(L"%d worksheet functions for %d slots: %d not registered, from %ls on", g_total, NSLOTS, g_total - g_nfuncs, first);
        free(first);
    }
}

/* --- starting Python --------------------------------------------------------------------------- */

static void set_natives(void);

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
    if (setting(L"PythonVenv", venv, MAX_PATH)) {
        note(L"PythonVenv: %ls", venv);
    } else {
#ifdef ASFPY_RUNTIME
        if (!unpack_runtime(venv)) return;
#else
        start_failed(L"no Python is set: run AsfPySetup in VBA, or set PythonVenv under HKCU\\%ls", SETTINGS_KEY);
        return;
#endif
    }
    if (!setting(L"PythonFolder", folder, MAX_PATH)) {
        DWORD k = GetEnvironmentVariableW(L"APPDATA", folder, MAX_PATH);
        if (!k || k > MAX_PATH - 20) folder[0] = 0; else wcscat(folder, L"\\ASF\\python");
    }
    if (!python_home(venv, g_home, site, exe)) {
        start_failed(L"%ls\\pyvenv.cfg names no home", venv);
        return;
    }
    wcscpy(g_exe, exe);
    if (!python_dll(g_home, dll)) {
        start_failed(L"no python3XX.dll in %ls", g_home);
        return;
    }
    if ((other = python_in_process()) != NULL) {
        const wchar_t *copy = other_copy();
        if (copy)
            start_failed(L"another copy of this add-in (%ls) already runs Python in this Excel. Close Excel and start it again to use this copy", copy);
        else
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
    set_natives();
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

/* error results, filled once in DllMain and only read afterwards: any calculation thread may
   return them, and Excel doesn't free them (no xlbitDLLFree) */
#define NERRORS 64
static XLOPER12 g_err[NERRORS];

static LPXLOPER12 xlerror(int code)
{
    return &g_err[code >= 0 && code < NERRORS ? code : xlerrValue];
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

/* an argument of a worksheet function: a range as a list of rows, else a cell value */
static PyObject *arg_to_py(LPXLOPER12 x)
{
    return (x->xltype & 0xFFF) == xltypeMulti ? multi_to_py(x) : cell_to_py(x);
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

/* a result of tag_to_result, with what it holds */
static void free_result(LPXLOPER12 x)
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

EXPORT void WINAPI xlAutoFree12(LPXLOPER12 x)
{
    free_result(x);
}

/* --- asynchronous functions: the result goes back later, with xlAsyncReturn ------------------ */

typedef struct { XLOPER12 handle; } AsyncJob;   /* Excel's handle, copied: its XLOPER12 is Excel's */

/* res to Excel for the job, which is then freed; owned: res came from tag_to_result. Any thread
   may call it: xlAsyncReturn is the one callback Excel takes from a thread of the add-in's own. */
static int async_finish(AsyncJob *job, LPXLOPER12 res, int owned)
{
    XLOPER12 ret;
    LPXLOPER12 op[2] = { &job->handle, res };
    int rc;
    ret.xltype = xltypeNil;
    if (owned) res->xltype &= ~xlbitDLLFree;    /* xlAsyncReturn copies the value: the add-in frees it */
    rc = excel12(xlAsyncReturn, &ret, 2, op);
    if (rc != xlretSuccess) note(L"xlAsyncReturn failed (%d): the calculation was over or cancelled", rc);
    if (owned) free_result(res);
    free(job);
    return rc == xlretSuccess;
}

/* boot._native_async_return(token, tagged): called by Python, which holds the GIL; it is let go
   while Excel takes the result, since Excel's threads may be waiting for it */
static PyObject *native_async_return(PyObject *self, PyObject *args)
{
    AsyncJob *job = (AsyncJob *)(intptr_t)P.PyLong_AsLongLong(P.PyTuple_GetItem(args, 0));
    LPXLOPER12 res = tag_to_result(P.PyTuple_GetItem(args, 1));
    void *state;
    int ok;
    (void)self;
    state = P.PyEval_SaveThread();
    ok = async_finish(job, res, 1);
    P.PyEval_RestoreThread(state);
    return P.PyBool_FromLong(ok);
}

/* an asynchronous slot: Python starts the function and returns; its result comes later */
static void call_async(int slot, LPXLOPER12 *argv, int n)
{
    AsyncJob *job = malloc(sizeof *job);
    PyGILState_STATE gil;
    PyObject *args, *r;
    int i;
    job->handle = *argv[n];
    for (i = 0; i < n; i++)
        if ((argv[i]->xltype & 0xFFF) == xltypeErr) { async_finish(job, xlerror(argv[i]->val.err), 0); return; }
    gil = P.PyGILState_Ensure();
    args = P.PyList_New(n);
    for (i = 0; i < n; i++) P.PyList_SetItem(args, i, arg_to_py(argv[i]));
    r = boot_call("xl_call_async", 3, P.PyLong_FromLong(slot), args, P.PyLong_FromLongLong((long long)(intptr_t)job));
    if (r) P.Py_DecRef(r);
    P.PyGILState_Release(gil);
    if (!r) async_finish(job, xlerror(xlerrValue), 0);
}

/* --- the calling cell, for handles --------------------------------------------------------------- */

/* boot._native_caller(): '[Book1.xlsx]Sheet1!R2C3', the cell that called the function now
   running, or '' (a call from VBA, or a calculation without a cell). xlfCaller and xlSheetNm are
   thread-safe, so a thread-safe function may ask too. */
static PyObject *native_caller(PyObject *self, PyObject *unused)
{
    XLOPER12 ref, name;
    LPXLOPER12 op[1] = { &ref };
    wchar_t text[600] = L"";
    (void)self; (void)unused;
    if (excel12(xlfCaller, &ref, 0, NULL) != xlretSuccess) return wstr(L"");
    if ((ref.xltype & 0xFFF) == xltypeSRef || (ref.xltype & 0xFFF) == xltypeRef) {
        const XLREF12 *at = (ref.xltype & 0xFFF) == xltypeSRef ? &ref.val.sref.ref
                          : ref.val.mref.lpmref && ref.val.mref.lpmref->count ? &ref.val.mref.lpmref->reftbl[0] : NULL;
        if (at && excel12(xlSheetNm, &name, 1, op) == xlretSuccess) {
            if ((name.xltype & 0xFFF) == xltypeStr) {
                int k = name.val.str[0] < 500 ? name.val.str[0] : 500;
                memcpy(text, name.val.str + 1, k * sizeof(wchar_t));
                swprintf(text + k, 100, L"!R%dC%d", (int)at->rwFirst + 1, (int)at->colFirst + 1);
            }
            op[0] = &name;
            excel12(xlFree, NULL, 1, op);
            op[0] = &ref;
        }
    }
    excel12(xlFree, NULL, 1, op);
    return wstr(text);
}

/* --- Excel's Application, for xl_app() in a command ------------------------------------------- */

static BOOL CALLBACK find_xlmain(HWND w, LPARAM out)
{
    DWORD pid = 0;
    wchar_t cls[16];
    GetWindowThreadProcessId(w, &pid);
    if (pid == GetCurrentProcessId() && GetClassNameW(w, cls, 16) && wcscmp(cls, L"XLMAIN") == 0) {
        *(HWND *)out = w;
        return FALSE;
    }
    return TRUE;
}

/* boot._native_app(): the address of Excel's Application (an IDispatch with one reference for
   the caller), or a message. Through a workbook window's object (AccessibleObjectFromWindow),
   which is in this process: not the running object table, which may give another Excel. */
static PyObject *native_app(PyObject *self, PyObject *unused)
{
    typedef HRESULT (WINAPI *AOFW)(HWND, DWORD, REFIID, void **);
    static AOFW aofw;
    HWND main = NULL, desk, book;
    IDispatch *win = NULL;
    DISPID id;
    LPOLESTR name = (LPOLESTR)L"Application";
    DISPPARAMS none = { NULL, NULL, 0, 0 };
    VARIANT app;
    HRESULT hr;
    (void)self; (void)unused;
    if (!aofw) {
        HMODULE oleacc = LoadLibraryW(L"oleacc.dll");
        aofw = oleacc ? (AOFW)(void *)GetProcAddress(oleacc, "AccessibleObjectFromWindow") : NULL;
        if (!aofw) return wstr(L"oleacc.dll has no AccessibleObjectFromWindow");
    }
    EnumWindows(find_xlmain, (LPARAM)&main);
    desk = main ? FindWindowExW(main, NULL, L"XLDESK", NULL) : NULL;
    book = desk ? FindWindowExW(desk, NULL, L"EXCEL7", NULL) : NULL;
    if (!book) return wstr(L"Excel's window wasn't found: xl_app() needs a workbook open in this Excel");
    hr = aofw(book, (DWORD)0xFFFFFFF0 /* OBJID_NATIVEOM */, &IID_IDispatch_, (void **)&win);
    if (FAILED(hr) || !win) return wstr(L"the workbook window gave no object");
    VariantInit(&app);
    hr = win->lpVtbl->GetIDsOfNames(win, &IID_NULL_, &name, 1, LOCALE_USER_DEFAULT, &id);
    if (SUCCEEDED(hr))
        hr = win->lpVtbl->Invoke(win, id, &IID_NULL_, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &none, &app, NULL, NULL);
    win->lpVtbl->Release(win);
    if (FAILED(hr) || V_VT(&app) != VT_DISPATCH || !V_DISPATCH(&app)) {
        VariantClear(&app);
        return wstr(L"the workbook window gave no Application");
    }
    return P.PyLong_FromLongLong((long long)(intptr_t)V_DISPATCH(&app));   /* its reference goes to Python */
}

/* boot._native_release(address): IUnknown::Release of an address native_app gave */
static PyObject *native_release(PyObject *self, PyObject *args)
{
    IUnknown *u = (IUnknown *)(intptr_t)P.PyLong_AsLongLong(P.PyTuple_GetItem(args, 0));
    (void)self;
    if (u) u->lpVtbl->Release(u);
    P.Py_IncRef(P.None);
    return P.None;
}

static PyMethodDef g_native[] = {
    { "_native_caller", native_caller, METH_NOARGS, NULL },
    { "_native_async_return", native_async_return, METH_VARARGS, NULL },
    { "_native_app", native_app, METH_NOARGS, NULL },
    { "_native_release", native_release, METH_VARARGS, NULL },
};

/* the functions above as attributes of boot.py's module */
static void set_natives(void)
{
    size_t i;
    for (i = 0; i < sizeof g_native / sizeof g_native[0]; i++) {
        PyObject *f = P.PyCFunction_NewEx(&g_native[i], NULL, NULL);
        if (!f || P.PyObject_SetAttrString(g_boot, g_native[i].ml_name, f) < 0) note(L"no %hs for boot.py", g_native[i].ml_name);
        if (f) P.Py_DecRef(f);
    }
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
    if (g_funcs[slot].async) {
        call_async(slot, argv, n);
        return NULL;      /* registered with >: Excel reads no result */
    }
    for (i = 0; i < n; i++)
        if ((argv[i]->xltype & 0xFFF) == xltypeErr) return xlerror(argv[i]->val.err);
    gil = P.PyGILState_Ensure();
    args = P.PyList_New(n);
    for (i = 0; i < n; i++) P.PyList_SetItem(args, i, arg_to_py(argv[i]));
    r = boot_call("xl_call", 2, P.PyLong_FromLong(slot), args);
    res = r ? tag_to_result(r) : xlerror(xlerrValue);
    if (r) P.Py_DecRef(r);
    P.PyGILState_Release(gil);
    return res;
}

#ifdef _WIN64
#include "slots_x64.h"
#elif defined(_MSC_VER)
#include "slots_x86_msvc.h"    /* MinGW assembles slots_x86.S instead */
#endif

/* PY_CALL(name, arguments...): an exported function of the Python folder, looked up by name
   (boot.xl_call_export). Excel passes all 30 arguments, those not given as xltypeMissing, so
   one __stdcall signature serves in both bitnesses; the missing ones at the end are dropped. */
EXPORT LPXLOPER12 WINAPI asfpy_call(LPXLOPER12 name,
    LPXLOPER12 a0, LPXLOPER12 a1, LPXLOPER12 a2, LPXLOPER12 a3, LPXLOPER12 a4, LPXLOPER12 a5,
    LPXLOPER12 a6, LPXLOPER12 a7, LPXLOPER12 a8, LPXLOPER12 a9, LPXLOPER12 a10, LPXLOPER12 a11,
    LPXLOPER12 a12, LPXLOPER12 a13, LPXLOPER12 a14, LPXLOPER12 a15, LPXLOPER12 a16, LPXLOPER12 a17,
    LPXLOPER12 a18, LPXLOPER12 a19, LPXLOPER12 a20, LPXLOPER12 a21, LPXLOPER12 a22, LPXLOPER12 a23,
    LPXLOPER12 a24, LPXLOPER12 a25, LPXLOPER12 a26, LPXLOPER12 a27, LPXLOPER12 a28)
{
    LPXLOPER12 argv[PYCALL_ARGS] = { a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14,
                                     a15, a16, a17, a18, a19, a20, a21, a22, a23, a24, a25, a26, a27, a28 };
    PyGILState_STATE gil;
    PyObject *args, *r;
    LPXLOPER12 res;
    int i, n = PYCALL_ARGS;
    if (g_started != 1) return xlerror(xlerrName);
    if ((name->xltype & 0xFFF) == xltypeErr) return xlerror(name->val.err);
    while (n > 0 && (argv[n - 1]->xltype & 0xFFF) == xltypeMissing) n--;
    for (i = 0; i < n; i++)
        if ((argv[i]->xltype & 0xFFF) == xltypeErr) return xlerror(argv[i]->val.err);
    gil = P.PyGILState_Ensure();
    args = P.PyList_New(n);
    for (i = 0; i < n; i++) P.PyList_SetItem(args, i, arg_to_py(argv[i]));
    r = boot_call("xl_call_export", 2, cell_to_py(name), args);
    res = r ? tag_to_result(r) : xlerror(xlerrValue);
    if (r) P.Py_DecRef(r);
    P.PyGILState_Release(gil);
    return res;
}

/* the alerts after the functions were registered: none, or more than the slots */
static void report_functions(void)
{
    wchar_t msg[1200];
    if (g_nfuncs == 0) {
        swprintf(msg, 1200, L"ASF Python: Python started, but the Python folder gave no worksheet functions. Details: %ls", g_log);
        alert(msg);
    } else if (g_total > g_nfuncs) {
        swprintf(msg, 1200, L"ASF Python: the Python folder gave %d worksheet functions, and the add-in has %d slots: %d functions are not registered. Details: %ls",
                 g_total, NSLOTS, g_total - g_nfuncs, g_log);
        alert(msg);
    }
}

/* The command AsfPyReload (VBA: Application.Run "AsfPyReload"): the Python folder's modules are
   imported again and their functions registered again, without restarting Excel. Python and
   the packages it imported stay. A command, so that Excel takes the C API calls of registration. */
EXPORT int WINAPI asfpy_reload(void)
{
    PyGILState_STATE gil;
    PyObject *list;
    LARGE_INTEGER t0;
    int gone;
    if (g_started != 1) return 0;
    QueryPerformanceCounter(&t0);
    gone = unregister_functions();
    free_functions();
    gil = P.PyGILState_Ensure();
    list = boot_call("reload", 0);
    if (list) {
        read_functions(list);
        P.Py_DecRef(list);
    }
    P.PyGILState_Release(gil);
    register_functions();
    note(L"reload: %d functions unregistered, %d registered, in %.0f ms", gone, g_nfuncs, ms_since(t0));
    report_functions();
    return 1;
}

/* The command AsfPyPackages (VBA: Application.Run "AsfPyPackages"; ASF's ribbon: Python packages):
   pip for the Python of the add-in, in a console window of its own (packages.py), so downloads
   never block Excel. */
EXPORT int WINAPI asfpy_packages(void)
{
    wchar_t script[MAX_PATH], cmd[3 * MAX_PATH], msg[1200];
    STARTUPINFOW si;
    PROCESS_INFORMATION pi;
    HANDLE f;
    DWORD put, k = GetTempPathW(MAX_PATH - 24, script);
    memset(&si, 0, sizeof si);
    si.cb = sizeof si;
    if (!g_exe[0] || GetFileAttributesW(g_exe) == INVALID_FILE_ATTRIBUTES) {
        swprintf(msg, 1200, L"ASF Python: no python.exe for pip (%ls). Details: %ls", g_exe[0] ? g_exe : g_start_error, g_log);
        alert(msg);
        return 0;
    }
    if (!k) return 0;
    wcscat(script, L"asfpy-packages.py");
    f = CreateFileW(script, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    if (f == INVALID_HANDLE_VALUE) return 0;
    WriteFile(f, PACKAGES_PY, (DWORD)strlen(PACKAGES_PY), &put, NULL);
    CloseHandle(f);
    swprintf(cmd, 3 * MAX_PATH, L"\"%ls\" -I \"%ls\"", g_exe, script);
    if (!CreateProcessW(g_exe, cmd, NULL, NULL, FALSE, CREATE_NEW_CONSOLE, NULL, NULL, &si, &pi)) {
        note(L"%ls didn't start (error %lu)", cmd, GetLastError());
        return 0;
    }
    CloseHandle(pi.hThread);
    CloseHandle(pi.hProcess);
    return 1;
}

/* --- the add-in's entry points ----------------------------------------------------------------- */

/* Each time Excel loads the add-in: Python starts the first time only; the functions are
   registered each time, since xlAutoClose unregisters them. */
EXPORT int WINAPI xlAutoOpen(void)
{
    HMODULE self;
    wchar_t msg[1200];
    /* pinned: Excel can't unload code that Python and VBA's proxies still use */
    GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_PIN, (LPCWSTR)(void *)xlAutoOpen, &self);
    if (!g_started) start_python();
    if (g_started == 1) {
        register_builtins();
        register_functions();
        report_functions();
    } else {
        swprintf(msg, 1200, L"ASF Python: %ls. Details: %ls", g_start_error, g_log);
        alert(msg);
    }
    return 1;
}

/* When the add-in is unticked or Excel closes: the functions are unregistered, so Excel no longer
   calls them, and their cells show #VALUE! while the add-in is off. Excel keeps a function's name
   for the session: deleting it, even after registering the function again as hidden as
   Excel-DNA does, still gave #VALUE! (phase 1b's checklist). Python stays: it is never
   finalized, since numpy and pandas can't start again in one process. */
EXPORT int WINAPI xlAutoClose(void)
{
    int n = unregister_functions();
    unregister_id(&g_pycall_id);
    unregister_id(&g_reload_id);
    unregister_id(&g_packages_id);
    if (n) note(L"unregistered %d functions", n);
    return 1;
}

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

/* --- For Each: a Python iterator as IEnumVARIANT ------------------------------------------------ */

typedef struct { IEnumVARIANTVtbl *lpVtbl; LONG refs; PyObject *it; int level; } Enum;
static IEnumVARIANTVtbl g_enum_vtbl;
static const GUID IID_IEnumVARIANT_ = { 0x00020404, 0, 0, { 0xC0, 0, 0, 0, 0, 0, 0, 0x46 } };

static HRESULT STDMETHODCALLTYPE en_QueryInterface(IEnumVARIANT *self, REFIID riid, void **out)
{
    if (IsEqualGUID(riid, &IID_IUnknown_) || IsEqualGUID(riid, &IID_IEnumVARIANT_)) {
        *out = self;
        self->lpVtbl->AddRef(self);
        return S_OK;
    }
    *out = NULL;
    return E_NOINTERFACE;
}

static ULONG STDMETHODCALLTYPE en_AddRef(IEnumVARIANT *self) { return InterlockedIncrement(&((Enum *)self)->refs); }

static ULONG STDMETHODCALLTYPE en_Release(IEnumVARIANT *self)
{
    Enum *e = (Enum *)self;
    LONG n = InterlockedDecrement(&e->refs);
    if (n == 0) {
        PyGILState_STATE gil = P.PyGILState_Ensure();
        P.Py_DecRef(e->it);
        P.PyGILState_Release(gil);
        free(e);
    }
    return n;
}

static HRESULT STDMETHODCALLTYPE en_Next(IEnumVARIANT *self, ULONG celt, VARIANT *out, ULONG *fetched)
{
    Enum *e = (Enum *)self;
    PyGILState_STATE gil;
    HRESULT hr = S_OK;
    ULONG i;
    if (fetched) *fetched = 0;
    gil = P.PyGILState_Ensure();
    for (i = 0; i < celt; i++) {
        wchar_t *msg = NULL;
        PyObject *r;
        P.Py_IncRef(e->it);
        r = boot_call("com_next", 2, e->it, P.PyLong_FromLong(e->level));
        if (!r) {
            fail(NULL, L"the add-in's bootstrap failed: see the log");
            hr = E_FAIL;
            break;
        }
        if (P.PyLong_AsLong(P.PyTuple_GetItem(r, 0)) == TAG_STOP) {
            P.Py_DecRef(r);
            hr = S_FALSE;
            break;
        }
        if (!tag_to_variant(r, &out[i], e->level, &msg)) {
            fail(NULL, msg);
            free(msg);
            P.Py_DecRef(r);
            hr = E_FAIL;
            break;
        }
        P.Py_DecRef(r);
    }
    P.PyGILState_Release(gil);
    if (fetched) *fetched = i;
    return hr;
}

static HRESULT STDMETHODCALLTYPE en_Skip(IEnumVARIANT *self, ULONG celt)
{
    HRESULT hr = S_OK;
    while (celt-- && hr == S_OK) {
        VARIANT v;
        VariantInit(&v);
        hr = en_Next(self, 1, &v, NULL);
        VariantClear(&v);
    }
    return hr;
}

/* a Python iterator runs once: it can't start again, nor be copied */
static HRESULT STDMETHODCALLTYPE en_Reset(IEnumVARIANT *self) { (void)self; return E_NOTIMPL; }
static HRESULT STDMETHODCALLTYPE en_Clone(IEnumVARIANT *self, IEnumVARIANT **out) { (void)self; *out = NULL; return E_NOTIMPL; }

static IEnumVARIANTVtbl g_enum_vtbl = { en_QueryInterface, en_AddRef, en_Release, en_Next, en_Skip, en_Reset, en_Clone };

/* For Each over a proxy: Invoke(DISPID_NEWENUM) gives an IEnumVARIANT of iter(obj) */
static HRESULT new_enum(Proxy *p, VARIANT *result, EXCEPINFO *ei)
{
    PyGILState_STATE gil = P.PyGILState_Ensure();
    HRESULT hr = S_OK;
    PyObject *r;
    P.Py_IncRef(p->obj);
    r = boot_call("com_iter", 1, p->obj);
    if (!r) {
        hr = fail(ei, L"the add-in's bootstrap failed: see the log");
    } else if (P.PyLong_AsLong(P.PyTuple_GetItem(r, 0)) == TAG_FAIL) {
        wchar_t *msg = pytext(P.PyTuple_GetItem(r, 1), NULL);
        hr = fail(ei, msg);
        free(msg);
    } else if (result) {
        Enum *e = calloc(1, sizeof(Enum));
        e->lpVtbl = &g_enum_vtbl;
        e->refs = 1;
        e->level = p->level;
        e->it = P.PyTuple_GetItem(r, 1);
        P.Py_IncRef(e->it);
        VariantInit(result);
        V_VT(result) = VT_UNKNOWN;
        V_UNKNOWN(result) = (IUnknown *)e;
    }
    if (r) P.Py_DecRef(r);
    P.PyGILState_Release(gil);
    return hr;
}

/* A failed call fills EXCEPINFO: a late-bound call in VBA then gets the message, as error
   vbObjectError + 1001. VBA's CallByName gets it too but raises error 440 "Automation error"
   without it, so ASF reads py.lastError (phase 1b's checklist). */
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
    if (id == DISPID_NEWENUM) return new_enum(p, result, ei);
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
    int i;
    (void)h; (void)reserved;
    if (reason == DLL_PROCESS_ATTACH) {
        InitializeCriticalSection(&g_names_lock);
        for (i = 0; i < NERRORS; i++) {
            g_err[i].xltype = xltypeErr;
            g_err[i].val.err = i;
        }
    }
    return TRUE;
}
