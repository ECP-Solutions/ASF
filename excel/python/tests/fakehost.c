/* A stand-in for Excel: loads asfpy64.xll (or asfpy32.xll), answers the C API calls the add-in
   makes, calls the registered worksheet functions through their slots and drives the py object
   through IDispatch, as VBA does. Run by run_wine.py; prints one line per check and ends with
   "N checks, M failed". The add-in finds MdCallBack12 in this program, as it does in Excel. */
#define _WIN32_WINNT 0x0601
#include <windows.h>
#include <oleauto.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <wchar.h>
#include <io.h>
#include <fcntl.h>
#include "../xll/xloper.h"

typedef struct { wchar_t proc[32], name[64]; int nargs; } Reg;
static Reg regs[256];
static int nregs, checks, failed;
static wchar_t xll_path[MAX_PATH];
static HMODULE xll;

static XCHAR *xstr(const wchar_t *s)
{
    size_t n = wcslen(s);
    XCHAR *x = malloc((n + 1) * sizeof(XCHAR));
    x[0] = (XCHAR)n;
    memcpy(x + 1, s, n * sizeof(XCHAR));
    return x;
}

static void copy_text(wchar_t *dst, LPXLOPER12 x, int max)
{
    int n = x->val.str[0] < max - 1 ? x->val.str[0] : max - 1;
    memcpy(dst, x->val.str + 1, n * sizeof(wchar_t));
    dst[n] = 0;
}

__declspec(dllexport) int WINAPI MdCallBack12(int fn, int n, LPXLOPER12 *op, LPXLOPER12 res)
{
    switch (fn) {
    case xlGetName:
        res->xltype = xltypeStr | xlbitXLFree;
        res->val.str = xstr(xll_path);
        return xlretSuccess;
    case xlFree:
        if (n > 0 && (op[0]->xltype & ~xlbitXLFree) == xltypeStr) free(op[0]->val.str);
        return xlretSuccess;
    case xlfRegister: {
        Reg *r = &regs[nregs++];
        wchar_t type[64];
        copy_text(r->proc, op[1], 32);
        copy_text(type, op[2], 64);
        copy_text(r->name, op[3], 64);
        r->nargs = (int)wcslen(type) - 1;
        res->xltype = xltypeNum;
        res->val.num = nregs;
        return xlretSuccess;
    }
    case 118 | 0x8000: {   /* xlcAlert */
        wchar_t text[1024];
        copy_text(text, op[0], 1024);
        wprintf(L"ALERT: %ls\n", text);
        return xlretSuccess;
    }
    default:
        return 2;   /* xlretInvXlfn */
    }
}

static void check(int ok, const wchar_t *what, const wchar_t *got)
{
    checks++;
    if (!ok) failed++;
    wprintf(L"%ls %ls: %ls\n", ok ? L"ok  " : L"FAIL", what, got);
}

/* a cell value as text: numbers with %g, errors as #n */
static void cell_text(LPXLOPER12 x, wchar_t *out, int max)
{
    switch (x->xltype & 0xFFF) {
    case xltypeNum: swprintf(out, max, L"%g", x->val.num); break;
    case xltypeStr: copy_text(out, x, max); break;
    case xltypeBool: wcscpy(out, x->val.xbool ? L"TRUE" : L"FALSE"); break;
    case xltypeErr: swprintf(out, max, L"#%d", x->val.err); break;
    case xltypeMulti: {
        int r, c;
        out[0] = 0;
        for (r = 0; r < x->val.array.rows; r++) {
            for (c = 0; c < x->val.array.columns; c++) {
                wchar_t cell[256];
                cell_text(&x->val.array.lparray[r * x->val.array.columns + c], cell, 256);
                wcscat(out, c ? L"," : (r ? L";" : L""));
                wcscat(out, cell);
            }
        }
        break;
    }
    default: wcscpy(out, L"?"); break;
    }
}

static XLOPER12 num(double v) { XLOPER12 x; x.xltype = xltypeNum; x.val.num = v; return x; }
static XLOPER12 str(const wchar_t *s) { XLOPER12 x; x.xltype = xltypeStr; x.val.str = xstr(s); return x; }
static XLOPER12 nil(void) { XLOPER12 x; x.xltype = xltypeNil; return x; }
static XLOPER12 err(int e) { XLOPER12 x; x.xltype = xltypeErr; x.val.err = e; return x; }

typedef LPXLOPER12 (WINAPI *F0)(void);
typedef LPXLOPER12 (WINAPI *F1)(LPXLOPER12);
typedef LPXLOPER12 (WINAPI *F2)(LPXLOPER12, LPXLOPER12);
typedef LPXLOPER12 (WINAPI *F3)(LPXLOPER12, LPXLOPER12, LPXLOPER12);
typedef LPXLOPER12 (WINAPI *F7)(LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12);

static FARPROC slot_of(const wchar_t *name)
{
    int i;
    char proc[32];
    for (i = 0; i < nregs; i++)
        if (_wcsicmp(regs[i].name, name) == 0) {
            WideCharToMultiByte(CP_ACP, 0, regs[i].proc, -1, proc, 32, NULL, NULL);
            return GetProcAddress(xll, proc);
        }
    return NULL;
}

/* calls name with n arguments (n: 0, 1, 2, 3 or 7) and compares the result as text */
static void call_check(const wchar_t *name, int n, LPXLOPER12 *a, const wchar_t *want)
{
    FARPROC f = slot_of(name);
    LPXLOPER12 r = NULL;
    wchar_t got[1024], what[128];
    void (WINAPI *free12)(LPXLOPER12) = (void (WINAPI *)(LPXLOPER12))GetProcAddress(xll, "xlAutoFree12");
    swprintf(what, 128, L"%ls", name);
    if (!f) { check(0, what, L"not registered"); return; }
    switch (n) {
    case 0: r = ((F0)f)(); break;
    case 1: r = ((F1)f)(a[0]); break;
    case 2: r = ((F2)f)(a[0], a[1]); break;
    case 3: r = ((F3)f)(a[0], a[1], a[2]); break;
    case 7: r = ((F7)f)(a[0], a[1], a[2], a[3], a[4], a[5], a[6]); break;
    }
    cell_text(r, got, 1024);
    check(wcsncmp(got, want, wcslen(want)) == 0, what, got);
    if (r->xltype & xlbitDLLFree) free12(r);
}

/* --- py through IDispatch, as VBA's late binding does -------------------------------------- */

static HRESULT invoke(IDispatch *d, const wchar_t *name, int n, VARIANT *args, VARIANT *res, wchar_t *errtext)
{
    DISPID id;
    LPOLESTR nm = (LPOLESTR)name;
    DISPPARAMS dp;
    VARIANT rev[8];
    EXCEPINFO ei;
    HRESULT hr;
    int i;
    for (i = 0; i < n; i++) rev[i] = args[n - 1 - i];
    dp.rgvarg = rev;
    dp.cArgs = n;
    dp.rgdispidNamedArgs = NULL;
    dp.cNamedArgs = 0;
    VariantInit(res);
    memset(&ei, 0, sizeof ei);
    hr = d->lpVtbl->GetIDsOfNames(d, &IID_NULL, &nm, 1, 0, &id);
    if (FAILED(hr)) return hr;
    hr = d->lpVtbl->Invoke(d, id, &IID_NULL, 0, DISPATCH_METHOD | DISPATCH_PROPERTYGET, &dp, res, &ei, NULL);
    if (hr == DISP_E_EXCEPTION && errtext) {
        wcsncpy(errtext, ei.bstrDescription ? ei.bstrDescription : L"", 255);
        errtext[255] = 0;
        SysFreeString(ei.bstrSource);
        SysFreeString(ei.bstrDescription);
    }
    return hr;
}

/* as VBA's CallByName with VbMethod calls it: each argument ByRef, the method flag only */
static HRESULT invoke_vba(IDispatch *d, const wchar_t *name, int n, VARIANT *args, VARIANT *res, wchar_t *errtext)
{
    DISPID id;
    LPOLESTR nm = (LPOLESTR)name;
    DISPPARAMS dp;
    VARIANT rev[8];
    EXCEPINFO ei;
    HRESULT hr;
    int i;
    for (i = 0; i < n; i++) {
        VariantInit(&rev[i]);
        V_VT(&rev[i]) = VT_BYREF | VT_VARIANT;
        V_VARIANTREF(&rev[i]) = &args[n - 1 - i];
    }
    dp.rgvarg = rev;
    dp.cArgs = n;
    dp.rgdispidNamedArgs = NULL;
    dp.cNamedArgs = 0;
    VariantInit(res);
    memset(&ei, 0, sizeof ei);
    hr = d->lpVtbl->GetIDsOfNames(d, &IID_NULL, &nm, 1, 0x409, &id);
    if (FAILED(hr)) return hr;
    hr = d->lpVtbl->Invoke(d, id, &IID_NULL, 0x409, DISPATCH_METHOD, &dp, res, &ei, NULL);
    if (hr == DISP_E_EXCEPTION && errtext) {
        wcsncpy(errtext, ei.bstrDescription ? ei.bstrDescription : L"", 255);
        errtext[255] = 0;
    }
    return hr;
}

static void variant_text(VARIANT *v, wchar_t *out, int max)
{
    VARIANT t;
    if (V_VT(v) == (VT_ARRAY | VT_VARIANT)) {
        LONG lo, hi, i;
        SafeArrayGetLBound(V_ARRAY(v), 1, &lo);
        SafeArrayGetUBound(V_ARRAY(v), 1, &hi);
        wcscpy(out, L"[");
        for (i = lo; i <= hi; i++) {
            VARIANT e;
            wchar_t s[256];
            VariantInit(&e);
            SafeArrayGetElement(V_ARRAY(v), &i, &e);
            variant_text(&e, s, 256);
            VariantClear(&e);
            if (i > lo) wcscat(out, L",");
            wcscat(out, s);
        }
        wcscat(out, L"]");
        return;
    }
    if (V_VT(v) == VT_DISPATCH) { wcscpy(out, L"<object>"); return; }
    if (V_VT(v) == VT_EMPTY) { wcscpy(out, L"<empty>"); return; }
    VariantInit(&t);
    if (SUCCEEDED(VariantChangeType(&t, v, 0, VT_BSTR))) {
        swprintf(out, max, L"%ls:%d", V_BSTR(&t), (int)V_VT(v));
        VariantClear(&t);
    } else {
        wcscpy(out, L"?");
    }
}

static VARIANT vnum(double d) { VARIANT v; VariantInit(&v); V_VT(&v) = VT_R8; V_R8(&v) = d; return v; }
static VARIANT vlong(long d) { VARIANT v; VariantInit(&v); V_VT(&v) = VT_I4; V_I4(&v) = d; return v; }
static VARIANT vstr(const wchar_t *s) { VARIANT v; VariantInit(&v); V_VT(&v) = VT_BSTR; V_BSTR(&v) = SysAllocString(s); return v; }

static void com_check(IDispatch *d, const wchar_t *name, int n, VARIANT *args, const wchar_t *want, VARIANT *keep)
{
    VARIANT r;
    wchar_t got[1024], errtext[256] = L"", what[128];
    HRESULT hr = invoke(d, name, n, args, &r, errtext);
    swprintf(what, 128, L"py %ls", name);
    if (hr == DISP_E_EXCEPTION) swprintf(got, 1024, L"error: %ls", errtext);
    else if (FAILED(hr)) swprintf(got, 1024, L"HRESULT 0x%08lx", (unsigned long)hr);
    else variant_text(&r, got, 1024);
    check(wcsncmp(got, want, wcslen(want)) == 0, what, got);
    if (keep && SUCCEEDED(hr)) *keep = r; else VariantClear(&r);
}

int wmain(int argc, wchar_t **argv)
{
    int (WINAPI *open)(void);
    long (WINAPI *root)(long, VARIANT *);
    LARGE_INTEGER t0, t1, fq;
    XLOPER12 a[8];
    LPXLOPER12 p[8];
    XLOPER12 multi, cells[6];
    VARIANT py, args[4], obj, mod;
    int i;
    if (argc < 2) { wprintf(L"usage: fakehost path\\to\\asfpy64.xll or asfpy32.xll\n"); return 2; }
    GetFullPathNameW(argv[1], MAX_PATH, xll_path, NULL);
    for (i = 0; i < 8; i++) p[i] = &a[i];
    CoInitialize(NULL);
    _setmode(_fileno(stdout), _O_U8TEXT);
    QueryPerformanceFrequency(&fq);

    xll = LoadLibraryW(xll_path);
    if (!xll) { wprintf(L"FAIL load %ls: %lu\n", xll_path, GetLastError()); return 1; }
    open = (int (WINAPI *)(void))GetProcAddress(xll, "xlAutoOpen");
    QueryPerformanceCounter(&t0);
    check(open() == 1, L"xlAutoOpen", L"1");
    QueryPerformanceCounter(&t1);
    wprintf(L"     start-up: %.0f ms, %d functions registered\n", (t1.QuadPart - t0.QuadPart) * 1000.0 / fq.QuadPart, nregs);
    check(nregs >= 10, L"registered", nregs >= 10 ? L"10 or more" : L"too few");
    if (nregs < 10) { wprintf(L"%d checks, %d failed\n", checks, failed); return 1; }

    call_check(L"spike_version", 0, p, L"Python 3.");
    a[0] = num(2); a[1] = num(3.5);
    call_check(L"spike_add", 2, p, L"5.5");
    a[0] = num(1); a[1] = str(L"x"); a[2] = nil();
    call_check(L"spike_args3", 3, p, L"1.0 | 'x' | None");
    for (i = 0; i < 7; i++) a[i] = num(i + 1);
    a[3] = nil();
    call_check(L"spike_sum7", 7, p, L"24");
    a[0] = num(2); a[1] = num(3);
    call_check(L"spike_grid", 2, p, L"11,12,13;21,22,23");
    for (i = 0; i < 6; i++) cells[i] = num(i + 1);
    cells[4] = str(L"t");
    cells[5] = err(xlerrNA);
    multi.xltype = xltypeMulti; multi.val.array.lparray = cells; multi.val.array.rows = 3; multi.val.array.columns = 2;
    a[0] = multi;
    call_check(L"spike_shape", 1, p, L"3 rows x 2 columns");
    multi.val.array.rows = 1; multi.val.array.columns = 4; a[0] = multi;
    call_check(L"spike_shape", 1, p, L"list of 4");
    multi.val.array.rows = 3; multi.val.array.columns = 1; a[0] = multi;
    call_check(L"spike_column", 1, p, L"10;20;30");
    call_check(L"spike_fail", 0, p, L"#15");
    a[0] = num(-1);
    call_check(L"spike_sqrt", 1, p, L"#15");
    a[0] = err(xlerrDiv0);
    call_check(L"spike_sqrt", 1, p, L"#7");
    a[0] = num(16);
    call_check(L"spike_sqrt", 1, p, L"4");
    a[0] = str(L"Ωmega");
    call_check(L"spike_text", 1, p, L"Ωmega · ü · 中文 · 5 chars");
    if (argc > 2 && wcscmp(argv[2], L"numpy") == 0) {
        XLOPER12 col[10];
        for (i = 0; i < 10; i++) col[i] = num(i + 1);
        multi.val.array.lparray = col; multi.val.array.rows = 10; multi.val.array.columns = 1; a[0] = multi;
        call_check(L"spike_numpy", 1, p, L"9.1");
    }

    /* the cost of a call: a function of two arguments, 20,000 times */
    {
        F2 f = (F2)slot_of(L"spike_add");
        void (WINAPI *free12)(LPXLOPER12) = (void (WINAPI *)(LPXLOPER12))GetProcAddress(xll, "xlAutoFree12");
        a[0] = num(1); a[1] = num(2);
        QueryPerformanceCounter(&t0);
        for (i = 0; i < 20000; i++) { LPXLOPER12 r = f(p[0], p[1]); if (r->xltype & xlbitDLLFree) free12(r); }
        QueryPerformanceCounter(&t1);
        wprintf(L"     worksheet call: %.1f us\n", (t1.QuadPart - t0.QuadPart) * 1e6 / fq.QuadPart / 20000);
    }

    root = (long (WINAPI *)(long, VARIANT *))GetProcAddress(xll, "AsfPy_Root");
    VariantInit(&py);
    i = root(1, &py);
    check(i == 0 && V_VT(&py) == VT_DISPATCH, L"AsfPy_Root(1)", V_VT(&py) == VT_DISPATCH ? L"object" : L"no object");
    args[0] = vstr(L"spike.add"); args[1] = vnum(2); args[2] = vlong(40);
    com_check(V_DISPATCH(&py), L"call", 3, args, L"42:5", NULL);
    {
        VARIANT r;
        wchar_t got[256], errtext[256] = L"";
        HRESULT hr;
        args[0] = vstr(L"spike.add");
        VariantInit(&args[1]); V_VT(&args[1]) = VT_I2; V_I2(&args[1]) = 2;
        VariantInit(&args[2]); V_VT(&args[2]) = VT_I2; V_I2(&args[2]) = 40;
        hr = invoke_vba(V_DISPATCH(&py), L"call", 3, args, &r, errtext);
        if (hr == DISP_E_EXCEPTION) swprintf(got, 256, L"error: %ls", errtext);
        else if (FAILED(hr)) swprintf(got, 256, L"HRESULT 0x%08lx", (unsigned long)hr);
        else variant_text(&r, got, 256);
        check(wcscmp(got, L"42:3") == 0, L"py call, as VBA's CallByName makes it", got);
    }
    args[0] = vstr(L"spike.echo"); args[1] = vstr(L"Ωmega"); args[2] = vlong(7);
    com_check(V_DISPATCH(&py), L"call", 3, args, L"[Ωmega:8,7:3]", NULL);
    {
        SAFEARRAYBOUND b[2] = { { 2, 1 }, { 2, 1 } };   /* VBA's a(1 To 2, 1 To 2) */
        SAFEARRAY *sa = SafeArrayCreate(VT_VARIANT, 2, b);
        LONG ix[2];
        VARIANT e;
        int r, c;
        for (r = 1; r <= 2; r++)
            for (c = 1; c <= 2; c++) {
                ix[0] = c; ix[1] = r;
                e = vnum(r * 10 + c);
                SafeArrayPutElement(sa, ix, &e);
            }
        args[0] = vstr(L"spike.mean");
        VariantInit(&args[1]); V_VT(&args[1]) = VT_ARRAY | VT_VARIANT; V_ARRAY(&args[1]) = sa;
        com_check(V_DISPATCH(&py), L"call", 2, args, L"16.5:5", NULL);
    }
    args[0] = vstr(L"os.system"); args[1] = vstr(L"calc");
    com_check(V_DISPATCH(&py), L"call", 2, args, L"error: NameError: no exported function 'os.system'", NULL);
    com_check(V_DISPATCH(&py), L"lastError", 0, args, L"NameError", NULL);
    args[0] = vstr(L"math");
    com_check(V_DISPATCH(&py), L"module", 1, args, L"error: AttributeError", NULL);
    args[0] = vstr(L"spike.echo");
    com_check(V_DISPATCH(&py), L"__class__", 0, args, L"error: AttributeError", NULL);

    VariantClear(&py);
    i = root(2, &py);
    check(i == 0 && V_VT(&py) == VT_DISPATCH, L"AsfPy_Root(2)", V_VT(&py) == VT_DISPATCH ? L"object" : L"no object");
    args[0] = vstr(L"math");
    com_check(V_DISPATCH(&py), L"module", 1, args, L"<object>", &mod);
    args[0] = vnum(2);
    com_check(V_DISPATCH(&mod), L"sqrt", 1, args, L"1.41421", NULL);
    com_check(V_DISPATCH(&mod), L"PI", 0, args, L"3.14159", NULL);      /* case: pi */
    args[0] = vstr(L"collections");
    com_check(V_DISPATCH(&py), L"module", 1, args, L"<object>", &obj);
    args[0] = vstr(L"abcab");
    com_check(V_DISPATCH(&obj), L"counter", 1, args, L"<object>", &mod); /* Counter, by case */
    args[0] = vlong(1);
    com_check(V_DISPATCH(&mod), L"most_common", 1, args, L"[[a:8,2:3]]", NULL);
    VariantClear(&mod);
    VariantClear(&obj);
    if (argc > 2 && wcscmp(argv[2], L"numpy") == 0) {
        args[0] = vstr(L"numpy");
        com_check(V_DISPATCH(&py), L"module", 1, args, L"<object>", &mod);
        args[0] = vstr(L"spike.echo"); args[1] = vlong(1); args[2] = vlong(2);
        com_check(V_DISPATCH(&py), L"call", 3, args, L"[1:3,2:3]", &obj);
        args[0] = obj; args[1] = vlong(50);
        com_check(V_DISPATCH(&mod), L"percentile", 2, args, L"1.5:5", NULL);
        args[0] = obj;
        com_check(V_DISPATCH(&mod), L"array", 1, args, L"<object>", &obj);
        com_check(V_DISPATCH(&obj), L"tolist", 0, args, L"[1:3,2:3]", NULL);
        VariantClear(&mod);
        VariantClear(&obj);
    }

    /* the cost of a call through IDispatch, as VBA makes it */
    {
        VARIANT r;
        args[0] = vstr(L"spike.add"); args[1] = vnum(1); args[2] = vnum(2);
        QueryPerformanceCounter(&t0);
        for (i = 0; i < 20000; i++) { invoke(V_DISPATCH(&py), L"call", 3, args, &r, NULL); VariantClear(&r); }
        QueryPerformanceCounter(&t1);
        wprintf(L"     py.call: %.1f us\n", (t1.QuadPart - t0.QuadPart) * 1e6 / fq.QuadPart / 20000);
    }
    VariantClear(&py);
    wprintf(L"%d checks, %d failed\n", checks, failed);
    return failed ? 1 : 0;
}
