/* A stand-in for Excel: loads asfpy64.xll (or asfpy32.xll), answers the C API calls the add-in
   makes, calls the registered worksheet functions through their slots and drives the py object
   through IDispatch, as VBA does. Run by run_fakehost.py; prints one line per check and ends with
   "N checks, M failed". The add-in finds MdCallBack12 in this program, as it does in Excel. */
#define _WIN32_WINNT 0x0601
#include <windows.h>
#include <oleauto.h>
#include <psapi.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <wchar.h>
#include <io.h>
#include <fcntl.h>
#include "../../xll/xloper.h"

#define NSLOTS 2000   /* as in asfpy.c */
#define NBUILTIN 9    /* PY_CALL and the commands AsfPyReload, AsfPyPackages, AsfPyConsole, AsfPyStudio, AsfPyRunScript, AsfPyMacros, AsfPyLastError, AsfPyUpdate */
#define SETTINGS_KEY L"Software\\VB and VBA Program Settings\\ASF\\AddIn"

#if defined(_MSC_VER) && !defined(_WIN64)
#pragma comment(linker, "/EXPORT:MdCallBack12=_MdCallBack12@16")   /* undecorated, as Excel exports it */
#endif

typedef struct { wchar_t proc[32], name[64], type[48]; int nargs, macro_type; } Reg;
static Reg regs[8 * NSLOTS];
static int nregs, nunregs, checks, failed;
static wchar_t xll_path[MAX_PATH], last_alert[1024];
static HMODULE xll;

/* the cell xlfCaller gives, and its sheet's name; row 0: no cell (a call from VBA) */
static int caller_row, caller_col;
static wchar_t caller_sheet[64] = L"[Book1.xlsx]Sheet1";

/* the results of asynchronous functions, as xlAsyncReturn gives them, by handle */
#define NASYNC 64
static struct { INT_PTR handle; wchar_t text[256]; } async_results[NASYNC];
static int nasync;
static CRITICAL_SECTION async_lock;
static void cell_text(LPXLOPER12 x, wchar_t *out, int max);

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
        Reg *r;
        if (nregs == 8 * NSLOTS) return 2;
        r = &regs[nregs++];
        copy_text(r->proc, op[1], 32);
        copy_text(r->type, op[2], 48);
        copy_text(r->name, op[3], 64);
        r->nargs = r->type[0] == L'>' ? (int)wcsspn(r->type + 1, L"QX") : (int)wcsspn(r->type, L"Q") - 1;
        r->macro_type = (int)op[5]->val.num;
        res->xltype = xltypeNum;
        res->val.num = nregs;
        return xlretSuccess;
    }
    case xlfUnregister:
        nunregs++;
        res->xltype = xltypeBool;
        res->val.xbool = TRUE;
        return xlretSuccess;
    case xlfCaller:
        if (!caller_row) { res->xltype = xltypeErr; res->val.err = xlerrRef; return xlretSuccess; }
        res->xltype = xltypeSRef;
        res->val.sref.count = 1;
        res->val.sref.ref.rwFirst = res->val.sref.ref.rwLast = caller_row - 1;
        res->val.sref.ref.colFirst = res->val.sref.ref.colLast = caller_col - 1;
        return xlretSuccess;
    case xlSheetNm:
        res->xltype = xltypeStr | xlbitXLFree;
        res->val.str = xstr(caller_sheet);
        return xlretSuccess;
    case xlAsyncReturn:     /* from any thread, as in Excel */
        EnterCriticalSection(&async_lock);
        if (nasync < NASYNC && (op[0]->xltype & 0xFFF) == xltypeBigData) {
            async_results[nasync].handle = (INT_PTR)op[0]->val.bigdata.h.hdata;
            cell_text(op[1], async_results[nasync].text, 256);
            nasync++;
        }
        LeaveCriticalSection(&async_lock);
        res->xltype = xltypeBool;
        res->val.xbool = TRUE;
        return xlretSuccess;
    case 118 | 0x8000:     /* xlcAlert */
        copy_text(last_alert, op[0], 1024);
        wprintf(L"ALERT: %ls\n", last_alert);
        return xlretSuccess;
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
typedef LPXLOPER12 (WINAPI *F4)(LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12);
typedef LPXLOPER12 (WINAPI *F7)(LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12);

static void (WINAPI *free12)(LPXLOPER12);

/* a result back to the add-in, as Excel does: only when it says the add-in frees it */
static void release(LPXLOPER12 r)
{
    if (r->xltype & xlbitDLLFree) free12(r);
}

/* the latest registration of a name: registering it again replaces it, as in Excel */
static Reg *reg_of(const wchar_t *name)
{
    int i;
    for (i = nregs - 1; i >= 0; i--)
        if (_wcsicmp(regs[i].name, name) == 0) return &regs[i];
    return NULL;
}

static FARPROC slot_of(const wchar_t *name)
{
    Reg *r = reg_of(name);
    char proc[32];
    if (!r) return NULL;
    WideCharToMultiByte(CP_ACP, 0, r->proc, -1, proc, 32, NULL, NULL);
    return GetProcAddress(xll, proc);
}

static Reg *reg_of_proc(const wchar_t *proc)
{
    int i;
    for (i = nregs - 1; i >= 0; i--)
        if (wcscmp(regs[i].proc, proc) == 0) return &regs[i];
    return NULL;
}

/* PY_CALL as Excel calls it: always 30 arguments, those not given as xltypeMissing */
typedef LPXLOPER12 (WINAPI *F30)(LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12,
                                 LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12,
                                 LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12,
                                 LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12,
                                 LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12);

static void pycall_check(const wchar_t *what, int n, LPXLOPER12 *given, const wchar_t *want)
{
    F30 f = (F30)slot_of(L"PY_CALL");
    XLOPER12 missing;
    LPXLOPER12 a[30], r;
    wchar_t got[1024];
    int i;
    if (!f) { check(0, what, L"PY_CALL not registered"); return; }
    missing.xltype = xltypeMissing;
    for (i = 0; i < 30; i++) a[i] = i < n ? given[i] : &missing;
    r = f(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8], a[9], a[10], a[11], a[12], a[13], a[14],
          a[15], a[16], a[17], a[18], a[19], a[20], a[21], a[22], a[23], a[24], a[25], a[26], a[27], a[28], a[29]);
    cell_text(r, got, 1024);
    check(wcscmp(got, want) == 0, what, got);
    release(r);
}

/* an asynchronous function's call: its arguments and a handle, then the result xlAsyncReturn
   gives for that handle, waited for up to 10 seconds */
static XLOPER12 async_handle(INT_PTR h) { XLOPER12 x; x.xltype = xltypeBigData; x.val.bigdata.h.hdata = (HANDLE)h; x.val.bigdata.cbData = 0; return x; }

static int async_result(INT_PTR h, wchar_t *got, int max)
{
    int i, t;
    for (t = 0; t < 1000; t++) {
        EnterCriticalSection(&async_lock);
        for (i = 0; i < nasync; i++)
            if (async_results[i].handle == h) {
                wcsncpy(got, async_results[i].text, max);
                LeaveCriticalSection(&async_lock);
                return 1;
            }
        LeaveCriticalSection(&async_lock);
        Sleep(10);
    }
    wcscpy(got, L"no result in 10 s");
    return 0;
}

static void async_check(const wchar_t *name, XLOPER12 arg, INT_PTR h, const wchar_t *want)
{
    F2 f = (F2)slot_of(name);
    XLOPER12 handle = async_handle(h);
    wchar_t got[256], what[128];
    swprintf(what, 128, L"%ls, asynchronous", name);
    if (!f) { check(0, what, L"not registered"); return; }
    f(&arg, &handle);       /* registered with >: no result now */
    async_result(h, got, 256);
    check(wcscmp(got, want) == 0, what, got);
}

/* a file in the add-in's Python folder (the setting run_fakehost.py made); NULL text deletes it */
static int folder_file(const wchar_t *name, const char *text)
{
    wchar_t folder[MAX_PATH], path[MAX_PATH];
    DWORD size = sizeof folder;
    FILE *f;
    if (RegGetValueW(HKEY_CURRENT_USER, SETTINGS_KEY, L"PythonFolder", RRF_RT_REG_SZ, NULL, folder, &size) != ERROR_SUCCESS)
        return 0;
    swprintf(path, MAX_PATH, L"%ls\\%ls", folder, name);
    if (!text) return DeleteFileW(path) != 0;
    f = _wfopen(path, L"wb");
    if (!f) return 0;
    fputs(text, f);
    fclose(f);
    return 1;
}

static void type_check(const wchar_t *name, const wchar_t *want)
{
    Reg *r = reg_of(name);
    wchar_t what[128];
    swprintf(what, 128, L"%ls registered with type %ls", name, want);
    check(r && wcscmp(r->type, want) == 0, what, r ? r->type : L"not registered");
}

/* calls name with n arguments (n: 0, 1, 2, 3 or 7) and compares the result as text */
static void call_check(const wchar_t *name, int n, LPXLOPER12 *a, const wchar_t *want)
{
    FARPROC f = slot_of(name);
    LPXLOPER12 r = NULL;
    wchar_t got[1024], what[128];
    swprintf(what, 128, L"%ls", name);
    if (!f) { check(0, what, L"not registered"); return; }
    switch (n) {
    case 0: r = ((F0)f)(); break;
    case 1: r = ((F1)f)(a[0]); break;
    case 2: r = ((F2)f)(a[0], a[1]); break;
    case 3: r = ((F3)f)(a[0], a[1], a[2]); break;
    case 4: r = ((F4)f)(a[0], a[1], a[2], a[3]); break;
    case 7: r = ((F7)f)(a[0], a[1], a[2], a[3], a[4], a[5], a[6]); break;
    default: check(0, what, L"the fake host can't call this many arguments"); return;
    }
    cell_text(r, got, 1024);
    check(wcsncmp(got, want, wcslen(want)) == 0, what, got);
    release(r);
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

/* VBA's For Each: Invoke(DISPID_NEWENUM), then IEnumVARIANT::Next, chunk items at a time */
static void enum_check(IDispatch *d, ULONG chunk, const wchar_t *what, const wchar_t *want)
{
    VARIANT r, items[8];
    DISPPARAMS dp = { NULL, NULL, 0, 0 };
    EXCEPINFO ei;
    IEnumVARIANT *en = NULL;
    wchar_t got[1024] = L"", item[256];
    HRESULT hr;
    ULONG n, i;
    VariantInit(&r);
    memset(&ei, 0, sizeof ei);
    hr = d->lpVtbl->Invoke(d, DISPID_NEWENUM, &IID_NULL, 0, DISPATCH_METHOD | DISPATCH_PROPERTYGET, &dp, &r, &ei, NULL);
    if (hr == DISP_E_EXCEPTION) {
        swprintf(got, 1024, L"error: %ls", ei.bstrDescription ? ei.bstrDescription : L"");
        SysFreeString(ei.bstrSource);
        SysFreeString(ei.bstrDescription);
    } else if (FAILED(hr) || V_VT(&r) != VT_UNKNOWN ||
               FAILED(V_UNKNOWN(&r)->lpVtbl->QueryInterface(V_UNKNOWN(&r), &IID_IEnumVARIANT, (void **)&en))) {
        swprintf(got, 1024, L"HRESULT 0x%08lx, type %d", (unsigned long)hr, (int)V_VT(&r));
    } else {
        do {
            hr = en->lpVtbl->Next(en, chunk, items, &n);
            for (i = 0; i < n; i++) {
                variant_text(&items[i], item, 256);
                if (got[0]) wcscat(got, L",");
                wcscat(got, item);
                VariantClear(&items[i]);
            }
        } while (hr == S_OK);
        if (FAILED(hr)) swprintf(got + wcslen(got), 64, L" then HRESULT 0x%08lx", (unsigned long)hr);
        en->lpVtbl->Release(en);
    }
    VariantClear(&r);
    check(wcscmp(got, want) == 0, what, got);
}

/* VBA's x = obj: the default value, DISPID_VALUE read as a property */
static void value_check(IDispatch *d, const wchar_t *what, const wchar_t *want)
{
    VARIANT r;
    DISPPARAMS dp = { NULL, NULL, 0, 0 };
    EXCEPINFO ei;
    wchar_t got[1024];
    HRESULT hr;
    VariantInit(&r);
    memset(&ei, 0, sizeof ei);
    hr = d->lpVtbl->Invoke(d, DISPID_VALUE, &IID_NULL, 0, DISPATCH_PROPERTYGET, &dp, &r, &ei, NULL);
    if (FAILED(hr)) swprintf(got, 1024, L"HRESULT 0x%08lx", (unsigned long)hr);
    else variant_text(&r, got, 1024);
    VariantClear(&r);
    check(wcscmp(got, want) == 0, what, got);
}

/* --- threads: a thread-safe function called from several threads at once ------------------- */

typedef struct { F1 f; int base, bad; } Job;

/* Excel calls worksheet functions from deep in its own frames. A 64-bit slot reads all 32 of
   its argument places from the caller's stack (asfpy.c), whatever the function takes, so a thread
   that calls one needs that much stack above the call, as Excel's threads have: room is it. A
   fresh thread under Wine starts too near the top of its stack without it. */
static DWORD WINAPI worker(LPVOID arg)
{
    Job *j = (Job *)arg;
    volatile char room[1024];
    int i;
    room[0] = 0;
    for (i = 0; i < 2000; i++) {
        XLOPER12 x = num(j->base + i);
        LPXLOPER12 r = j->f(&x);
        if ((r->xltype & 0xFFF) != xltypeNum || r->val.num != 2.0 * (j->base + i)) j->bad++;
        release(r);
    }
    return room[0];
}

static SIZE_T private_bytes(void)
{
    PROCESS_MEMORY_COUNTERS_EX pm;
    pm.cb = sizeof pm;
    GetProcessMemoryInfo(GetCurrentProcess(), (PROCESS_MEMORY_COUNTERS *)&pm, sizeof pm);
    return pm.PrivateUsage;
}

/* n rounds of calls that allocate results: text, a grid, a range in, an array back to VBA */
static void churn(int n, IDispatch *py)
{
    F1 text = (F1)slot_of(L"check_text"), column = (F1)slot_of(L"check_column");
    F2 grid = (F2)slot_of(L"check_grid");
    XLOPER12 s = str(L"Ωmega"), rows = num(2), cols = num(3), multi, cells[3];
    VARIANT args[3], r;
    int i;
    for (i = 0; i < 3; i++) cells[i] = num(i + 1);
    multi.xltype = xltypeMulti; multi.val.array.lparray = cells; multi.val.array.rows = 3; multi.val.array.columns = 1;
    args[0] = vstr(L"checks.echo"); args[1] = vstr(L"Ωmega"); args[2] = vlong(7);
    for (i = 0; i < n; i++) {
        release(text(&s));
        release(grid(&rows, &cols));
        release(column(&multi));
        invoke(py, L"call", 3, args, &r, NULL);
        VariantClear(&r);
    }
    free(s.val.str);
    for (i = 0; i < 3; i++) VariantClear(&args[i]);
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
    wchar_t got[256];
    int i;
    if (argc < 2) { wprintf(L"usage: fakehost path\\to\\asfpy64.xll or asfpy32.xll\n"); return 2; }
    GetFullPathNameW(argv[1], MAX_PATH, xll_path, NULL);
    for (i = 0; i < 8; i++) p[i] = &a[i];
    CoInitialize(NULL);
    InitializeCriticalSection(&async_lock);
    _setmode(_fileno(stdout), _O_U8TEXT);
    QueryPerformanceFrequency(&fq);

    {   /* what an update leaves next to the add-in: the add-in deletes it when it loads */
        wchar_t old[MAX_PATH + 16];
        FILE *f;
        swprintf(old, MAX_PATH + 16, L"%ls.old", xll_path);
        if ((f = _wfopen(old, L"w")) != NULL) fclose(f);
        swprintf(old, MAX_PATH + 16, L"%ls.7.old", xll_path);
        if ((f = _wfopen(old, L"w")) != NULL) fclose(f);
    }
    xll = LoadLibraryW(xll_path);
    if (!xll) { wprintf(L"FAIL load %ls: %lu\n", xll_path, GetLastError()); return 1; }
    open = (int (WINAPI *)(void))GetProcAddress(xll, "xlAutoOpen");
    free12 = (void (WINAPI *)(LPXLOPER12))GetProcAddress(xll, "xlAutoFree12");
    QueryPerformanceCounter(&t0);
    check(open() == 1, L"xlAutoOpen", L"1");
    QueryPerformanceCounter(&t1);
    wprintf(L"     start-up: %.0f ms, %d functions registered\n", (t1.QuadPart - t0.QuadPart) * 1000.0 / fq.QuadPart, nregs);
    swprintf(got, 256, L"%d", nregs);
    check(nregs == NSLOTS + NBUILTIN, L"registered: one function in each slot, PY_CALL and the 8 commands", got);
    {
        wchar_t old[MAX_PATH + 16];
        int left;
        swprintf(old, MAX_PATH + 16, L"%ls.old", xll_path);
        left = GetFileAttributesW(old) != INVALID_FILE_ATTRIBUTES;
        swprintf(old, MAX_PATH + 16, L"%ls.7.old", xll_path);
        left += GetFileAttributesW(old) != INVALID_FILE_ATTRIBUTES;
        swprintf(got, 256, L"%d left", left);
        check(left == 0, L"xlAutoOpen deletes the .old files an update left", got);
    }
    if (nregs < 20) { wprintf(L"%d checks, %d failed\n", checks, failed); return 1; }
    check(wcsstr(last_alert, L"are not registered") != NULL, L"alert: functions beyond the slots", last_alert);

    /* --- registration: types and their suffixes */
    type_check(L"PY_CALL", L"QQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQ");
    type_check(L"AsfPyReload", L"J");
    type_check(L"AsfPyPackages", L"J");
    {
        Reg *r = reg_of(L"AsfPyPackages");
        check(r && r->macro_type == 2, L"AsfPyPackages registered as a command (macro type 2)", r && r->macro_type == 2 ? L"2" : L"not 2");
    }
    type_check(L"AsfPyConsole", L"J");    /* not run: it opens a console window */
    type_check(L"AsfPyStudio", L"J");     /* not run: it opens a window */
    type_check(L"AsfPyUpdate", L"J");     /* not run: it asks in a message box */
    {
        Reg *r = reg_of(L"AsfPyConsole");
        check(r && r->macro_type == 2, L"AsfPyConsole registered as a command (macro type 2)", r && r->macro_type == 2 ? L"2" : L"not 2");
    }
    type_check(L"AsfPyRunScript", L"QQ");
    type_check(L"AsfPyMacros", L"Q");
    {   /* AsfPyMacros: a line per command; AsfPyRunScript: a .py file's output */
        Reg *r1 = reg_of(L"AsfPyRunScript"), *r2 = reg_of(L"AsfPyMacros");
        wchar_t path[MAX_PATH];
        FILE *f;
        check(r1 && r2 && r1->macro_type == 2 && r2->macro_type == 2, L"AsfPyRunScript and AsfPyMacros are commands",
              r1 && r2 && r1->macro_type == 2 && r2->macro_type == 2 ? L"2, 2" : L"not commands");
        call_check(L"AsfPyMacros", 0, p, L"check_app\t0\t");
        GetTempPathW(MAX_PATH - 40, path);
        wcscat(path, L"asfpy-fakehost-script.py");
        f = _wfopen(path, L"w");
        if (f) {
            fputs("import sys\nprint('ran', __name__, sys.argv[0].endswith('.py'))\n", f);
            fclose(f);
        }
        a[0] = str(path);
        call_check(L"AsfPyRunScript", 1, p, L"ran __main__ True");
        a[0] = str(L"no-such-script.py");
        call_check(L"AsfPyRunScript", 1, p, L"not a .py file");
        DeleteFileW(path);
    }
    type_check(L"check_add", L"QQQ");
    type_check(L"check_version", L"Q");
    type_check(L"check_volatile", L"Q!");
    {   /* a volatile function counts its calls: two calls, two numbers, the second larger */
        F0 f = (F0)slot_of(L"check_volatile");
        LPXLOPER12 r1 = f ? f() : NULL, r2;
        double n1 = r1 && (r1->xltype & 0xFFF) == xltypeNum ? r1->val.num : -1, n2;
        if (r1) release(r1);
        r2 = f ? f() : NULL;
        n2 = r2 && (r2->xltype & 0xFFF) == xltypeNum ? r2->val.num : -1;
        if (r2) release(r2);
        swprintf(got, 256, L"%g then %g", n1, n2);
        check(n1 >= 1 && n2 > n1, L"check_volatile counts its calls", got);
    }
    type_check(L"check_threadsafe", L"QQ$");
    check(reg_of(L"check_refused") == NULL, L"check_refused ($ with #) not registered",
          reg_of(L"check_refused") ? reg_of(L"check_refused")->type : L"not registered");

    /* --- the pool: the last slot, and slots of 0 to 3 arguments (the 32-bit stubs pop them) */
    {
        Reg *last = reg_of_proc(L"asfpy_f1999");
        int k = last ? _wtoi(last->name + 5) : 0, n = k % 4;
        wchar_t want[32];
        check(last && wcsncmp(last->name, L"slot_", 5) == 0, L"the last slot, asfpy_f1999, holds a slot_ function",
              last ? last->name : L"nothing");
        for (i = 0; i < n; i++) a[i] = num(1);
        swprintf(want, 32, L"%d", k * 100 + n);   /* below 10^6: %g writes it whole */
        if (last) call_check(last->name, n, p, want);
    }
    call_check(L"slot_0000", 0, p, L"0");
    a[0] = num(5);
    call_check(L"slot_0001", 1, p, L"105");
    a[0] = num(1); a[1] = num(2);
    call_check(L"slot_0502", 2, p, L"50203");
    a[0] = num(1); a[1] = num(2); a[2] = num(3);
    call_check(L"slot_0951", 3, p, L"95106");
    a[0] = num(1); a[1] = num(2); a[2] = num(3);
    call_check(L"slot_1503", 3, p, L"150306");
    check(reg_of(L"slot_1999") == NULL, L"slot_1999, beyond the pool, not registered", reg_of(L"slot_1999") ? L"registered" : L"not registered");

    call_check(L"check_version", 0, p, L"Python 3.");
    a[0] = num(2); a[1] = num(3.5);
    call_check(L"check_add", 2, p, L"5.5");
    a[0] = num(1); a[1] = str(L"x"); a[2] = nil();
    call_check(L"check_args3", 3, p, L"1.0 | 'x' | None");
    for (i = 0; i < 7; i++) a[i] = num(i + 1);
    a[3] = nil();
    call_check(L"check_sum7", 7, p, L"24");
    a[0] = num(2); a[1] = num(3);
    call_check(L"check_grid", 2, p, L"11,12,13;21,22,23");
    for (i = 0; i < 6; i++) cells[i] = num(i + 1);
    cells[4] = str(L"t");
    cells[5] = err(xlerrNA);
    multi.xltype = xltypeMulti; multi.val.array.lparray = cells; multi.val.array.rows = 3; multi.val.array.columns = 2;
    a[0] = multi;
    call_check(L"check_shape", 1, p, L"3 rows x 2 columns");
    multi.val.array.rows = 1; multi.val.array.columns = 4; a[0] = multi;
    call_check(L"check_shape", 1, p, L"list of 4");
    multi.val.array.rows = 3; multi.val.array.columns = 1; a[0] = multi;
    call_check(L"check_column", 1, p, L"10;20;30");
    call_check(L"check_fail", 0, p, L"#15");
    type_check(L"AsfPyLastError", L"Q");
    call_check(L"AsfPyLastError", 0, p, L"Traceback (most recent call last)");
    a[0] = num(-1);
    call_check(L"check_sqrt", 1, p, L"#15");
    a[0] = err(xlerrDiv0);
    call_check(L"check_sqrt", 1, p, L"#7");
    a[0] = num(16);
    call_check(L"check_sqrt", 1, p, L"4");
    a[0] = str(L"Ωmega");
    call_check(L"check_text", 1, p, L"Ωmega · ü · 中文 · 5 chars");

    /* --- type hints: int, date, str and list from cell values; a date back as a serial number */
    a[0] = num(2); a[1] = num(45000); a[2] = num(12); a[3] = num(7);
    call_check(L"check_hints", 4, p, L"2 | datetime.date(2023, 3, 15) | '12' | [7.0]");
    a[0] = num(45000);
    call_check(L"check_nextday", 1, p, L"45001");

    /* --- PY_CALL: exported functions by name, never imported; missing arguments at the end dropped */
    {
        XLOPER12 g[4], col[3], missing;
        LPXLOPER12 q[4] = { &g[0], &g[1], &g[2], &g[3] };
        missing.xltype = xltypeMissing;
        g[0] = str(L"checks.add"); g[1] = num(2); g[2] = num(40);
        pycall_check(L"PY_CALL(\"checks.add\", 2, 40)", 3, q, L"42");
        for (i = 0; i < 3; i++) col[i] = num(i + 1);
        g[0] = str(L"checks.scale");
        g[1].xltype = xltypeMulti; g[1].val.array.lparray = col; g[1].val.array.rows = 3; g[1].val.array.columns = 1;
        g[2] = num(10);
        pycall_check(L"PY_CALL(\"checks.scale\", A1:A3, 10)", 3, q, L"10;20;30");
        g[0] = str(L"checks.check_args3"); g[1] = num(1); g[2] = missing; g[3] = str(L"x");
        pycall_check(L"PY_CALL with a missing argument between two", 4, q, L"1.0 | None | 'x'");
        g[0] = str(L"os.system"); g[1] = str(L"calc");
        pycall_check(L"PY_CALL(\"os.system\", \"calc\"): #NAME?", 2, q, L"#29");
        g[0] = str(L"checks.add"); g[1] = str(L"x"); g[2] = num(1);
        pycall_check(L"PY_CALL raising TypeError: #VALUE!", 3, q, L"#15");
        g[0] = str(L"checks.add"); g[1] = num(1); g[2] = err(xlerrNA);
        pycall_check(L"PY_CALL with #N/A as an argument: #N/A", 3, q, L"#42");
        g[0] = err(xlerrDiv0);
        pycall_check(L"PY_CALL with #DIV/0! as the name: #DIV/0!", 1, q, L"#7");
    }
    if (argc > 2 && wcscmp(argv[2], L"numpy") == 0) {
        XLOPER12 col[10];
        for (i = 0; i < 10; i++) col[i] = num(i + 1);
        multi.val.array.lparray = col; multi.val.array.rows = 10; multi.val.array.columns = 1; a[0] = multi;
        call_check(L"check_numpy", 1, p, L"9.1");
    }

    /* --- handles, asynchronous functions, commands */
    type_check(L"check_async", L">QX");
    type_check(L"check_async_thread", L">QX");
    type_check(L"check_macro", L"QQQ");
    {
        Reg *r = reg_of(L"check_macro");
        check(r && r->macro_type == 2, L"check_macro registered as a command (macro type 2)", r && r->macro_type == 2 ? L"2" : L"not 2");
    }
    {
        wchar_t handle[64];
        LPXLOPER12 r;
        caller_row = 1; caller_col = 1;         /* A1 */
        a[0] = str(L"abcab");
        r = ((F1)slot_of(L"check_counter"))(p[0]);
        cell_text(r, handle, 64);
        check(wcsncmp(handle, L"<Counter#", 9) == 0, L"check_counter: a handle", handle);
        release(r);
        caller_col = 2;                         /* B1 */
        a[0] = str(handle); a[1] = num(1);
        call_check(L"check_most", 2, p, L"a");
        wcscpy(caller_sheet, L"[Book2.xlsx]Sheet1");
        call_check(L"check_most", 2, p, L"#15");   /* the handle names nothing in another workbook */
        wcscpy(caller_sheet, L"[Book1.xlsx]Sheet1");
        caller_col = 1;                         /* A1 again: its old handle goes */
        a[0] = str(L"xyy");
        release(((F1)slot_of(L"check_counter"))(p[0]));
        caller_col = 2;
        a[0] = str(handle); a[1] = num(1);
        call_check(L"check_most", 2, p, L"#15");
        caller_row = caller_col = 0;
    }
    async_check(L"check_async", num(21), 1001, L"42");
    async_check(L"check_async_thread", num(1), 1002, L"2");
    async_check(L"check_async", err(xlerrNA), 1003, L"#42");
    {
        /* 30 at once: each result comes back to its handle */
        F2 f = (F2)slot_of(L"check_async");
        int bad = 0;
        for (i = 0; f && i < 30; i++) {
            XLOPER12 x = num(i), h = async_handle(2000 + i);
            f(&x, &h);
        }
        for (i = 0; f && i < 30; i++) {
            wchar_t want[16];
            async_result(2000 + i, got, 256);
            swprintf(want, 16, L"%d", i * 2);
            if (wcscmp(got, want) != 0) bad++;
        }
        swprintf(got, 256, L"%d wrong of 30", bad);
        check(f && bad == 0, L"check_async, 30 calls at once", got);
    }
    a[0] = num(2); a[1] = num(3);
    call_check(L"check_macro", 2, p, L"5");
    call_check(L"check_app", 0, p, L"#15");    /* no pywin32 or no Excel window: an error, not a crash */
    if (argc > 2 && wcscmp(argv[2], L"numpy") == 0) {
        XLOPER12 g4[4] = { num(1), num(2), num(3), num(4) };
        multi.xltype = xltypeMulti;
        multi.val.array.lparray = g4; multi.val.array.rows = 2; multi.val.array.columns = 2; a[0] = multi;
        call_check(L"check_matrix", 1, p, L"4;6");
    }

    /* the cost of a call: a function of two arguments, 20,000 times */
    {
        F2 f = (F2)slot_of(L"check_add");
        a[0] = num(1); a[1] = num(2);
        QueryPerformanceCounter(&t0);
        for (i = 0; i < 20000; i++) release(f(p[0], p[1]));
        QueryPerformanceCounter(&t1);
        wprintf(L"     worksheet call: %.1f us\n", (t1.QuadPart - t0.QuadPart) * 1e6 / fq.QuadPart / 20000);
    }

    /* --- threads: functions that aren't thread-safe run on the main thread; a thread-safe one
       from 8 threads at once, as Excel's multi-threaded recalculation calls it */
    call_check(L"check_thread", 0, p, L"MainThread");
    if (!slot_of(L"check_threadsafe")) {
        check(0, L"check_threadsafe from 8 threads", L"not registered");
    } else {
        Job jobs[8];
        HANDLE th[8];
        int bad = 0;
        for (i = 0; i < 8; i++) {
            jobs[i].f = (F1)slot_of(L"check_threadsafe");
            jobs[i].base = i * 10000;
            jobs[i].bad = 0;
            th[i] = CreateThread(NULL, 0, worker, &jobs[i], 0, NULL);
        }
        WaitForMultipleObjects(8, th, TRUE, INFINITE);
        for (i = 0; i < 8; i++) {
            bad += jobs[i].bad;
            CloseHandle(th[i]);
        }
        swprintf(got, 256, L"%d wrong results of 16,000", bad);
        check(bad == 0, L"check_threadsafe from 8 threads", got);
    }

    root = (long (WINAPI *)(long, VARIANT *))GetProcAddress(xll, "AsfPy_Root");
    VariantInit(&py);
    i = root(1, &py);
    check(i == 0 && V_VT(&py) == VT_DISPATCH, L"AsfPy_Root(1)", V_VT(&py) == VT_DISPATCH ? L"object" : L"no object");
    args[0] = vstr(L"checks.add"); args[1] = vnum(2); args[2] = vlong(40);
    com_check(V_DISPATCH(&py), L"call", 3, args, L"42:5", NULL);
    {
        VARIANT r;
        wchar_t got[256], errtext[256] = L"";
        HRESULT hr;
        args[0] = vstr(L"checks.add");
        VariantInit(&args[1]); V_VT(&args[1]) = VT_I2; V_I2(&args[1]) = 2;
        VariantInit(&args[2]); V_VT(&args[2]) = VT_I2; V_I2(&args[2]) = 40;
        hr = invoke_vba(V_DISPATCH(&py), L"call", 3, args, &r, errtext);
        if (hr == DISP_E_EXCEPTION) swprintf(got, 256, L"error: %ls", errtext);
        else if (FAILED(hr)) swprintf(got, 256, L"HRESULT 0x%08lx", (unsigned long)hr);
        else variant_text(&r, got, 256);
        check(wcscmp(got, L"42:3") == 0, L"py call, as VBA's CallByName makes it", got);
    }
    args[0] = vstr(L"checks.echo"); args[1] = vstr(L"Ωmega"); args[2] = vlong(7);
    com_check(V_DISPATCH(&py), L"call", 3, args, L"[Ωmega:8,7:3]", NULL);
    args[0] = vstr(L"checks.squares"); args[1] = vlong(4);
    com_check(V_DISPATCH(&py), L"call", 2, args, L"<object>", &obj);
    if (V_VT(&obj) == VT_DISPATCH) enum_check(V_DISPATCH(&obj), 1, L"For Each over a generator, one at a time", L"0:3,1:3,4:3,9:3");
    VariantClear(&obj);
    args[0] = vstr(L"checks.squares"); args[1] = vlong(3);
    com_check(V_DISPATCH(&py), L"call", 2, args, L"<object>", &obj);
    if (V_VT(&obj) == VT_DISPATCH) enum_check(V_DISPATCH(&obj), 8, L"For Each over a generator, 8 at a time", L"0:3,1:3,4:3");
    VariantClear(&obj);
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
        args[0] = vstr(L"checks.mean");
        VariantInit(&args[1]); V_VT(&args[1]) = VT_ARRAY | VT_VARIANT; V_ARRAY(&args[1]) = sa;
        com_check(V_DISPATCH(&py), L"call", 2, args, L"16.5:5", NULL);
    }
    args[0] = vstr(L"os.system"); args[1] = vstr(L"calc");
    com_check(V_DISPATCH(&py), L"call", 2, args, L"error: NameError: no exported function 'os.system'", NULL);
    {
        /* what VBA's CallByName reads: ISupportErrorInfo, then the thread's error info */
        ISupportErrorInfo *sei = NULL;
        IErrorInfo *info = NULL;
        BSTR desc = NULL;
        HRESULT hr = V_DISPATCH(&py)->lpVtbl->QueryInterface(V_DISPATCH(&py), &IID_ISupportErrorInfo, (void **)&sei);
        int ok = SUCCEEDED(hr) && sei->lpVtbl->InterfaceSupportsErrorInfo(sei, &IID_IDispatch) == S_OK;
        if (sei) sei->lpVtbl->Release(sei);
        args[0] = vstr(L"os.system"); args[1] = vstr(L"calc");
        invoke_vba(V_DISPATCH(&py), L"call", 2, args, &obj, NULL);
        if (ok && GetErrorInfo(0, &info) == S_OK && info) {
            info->lpVtbl->GetDescription(info, &desc);
            info->lpVtbl->Release(info);
        }
        ok = ok && desc && wcsncmp(desc, L"NameError", 9) == 0;
        check(ok, L"py error info (ISupportErrorInfo, GetErrorInfo)", desc ? desc : L"none");
        SysFreeString(desc);
    }
    com_check(V_DISPATCH(&py), L"lastError", 0, args, L"NameError", NULL);
    com_check(V_DISPATCH(&py), L"lastTraceback", 0, args, L"Traceback (most recent call last)", NULL);
    com_check(V_DISPATCH(&py), L"clearError", 0, args, L"<empty>", NULL);
    com_check(V_DISPATCH(&py), L"lastError", 0, args, L":8", NULL);       /* empty text after clearError */
    args[0] = vstr(L"math");
    com_check(V_DISPATCH(&py), L"module", 1, args, L"error: AttributeError", NULL);
    args[0] = vstr(L"checks.echo");
    com_check(V_DISPATCH(&py), L"__class__", 0, args, L"error: AttributeError", NULL);

    VariantClear(&py);
    i = root(2, &py);
    check(i == 0 && V_VT(&py) == VT_DISPATCH, L"AsfPy_Root(2)", V_VT(&py) == VT_DISPATCH ? L"object" : L"no object");
    args[0] = vstr(L"math");
    com_check(V_DISPATCH(&py), L"module", 1, args, L"<object>", &mod);
    args[0] = vnum(2);
    com_check(V_DISPATCH(&mod), L"sqrt", 1, args, L"1.41421", NULL);
    com_check(V_DISPATCH(&mod), L"PI", 0, args, L"3.14159", NULL);      /* case: pi */
    enum_check(V_DISPATCH(&mod), 1, L"For Each over a module", L"error: TypeError: 'module' object is not iterable");
    VariantClear(&mod);
    args[0] = vstr(L"collections");
    com_check(V_DISPATCH(&py), L"module", 1, args, L"<object>", &obj);
    args[0] = vstr(L"abcab");
    com_check(V_DISPATCH(&obj), L"counter", 1, args, L"<object>", &mod); /* Counter, by case */
    args[0] = vlong(1);
    com_check(V_DISPATCH(&mod), L"most_common", 1, args, L"[[a:8,2:3]]", NULL);
    value_check(V_DISPATCH(&mod), L"the default value of a Counter: its text", L"Counter({'a': 2, 'b': 2, 'c': 1}):8");
    enum_check(V_DISPATCH(&mod), 1, L"For Each over a Counter", L"a:8,b:8,c:8");
    VariantClear(&mod);
    VariantClear(&obj);
    if (argc > 2 && wcscmp(argv[2], L"numpy") == 0) {
        args[0] = vstr(L"numpy");
        com_check(V_DISPATCH(&py), L"module", 1, args, L"<object>", &mod);
        args[0] = vstr(L"checks.echo"); args[1] = vlong(1); args[2] = vlong(2);
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
        args[0] = vstr(L"checks.add"); args[1] = vnum(1); args[2] = vnum(2);
        QueryPerformanceCounter(&t0);
        for (i = 0; i < 20000; i++) { invoke(V_DISPATCH(&py), L"call", 3, args, &r, NULL); VariantClear(&r); }
        QueryPerformanceCounter(&t1);
        wprintf(L"     py.call: %.1f us\n", (t1.QuadPart - t0.QuadPart) * 1e6 / fq.QuadPart / 20000);
    }

    /* --- memory: 100,000 rounds of calls whose results the add-in allocates and frees. The
       limit catches a leak of about 10 bytes a round: one result string not freed is 5 MB. */
    {
        SIZE_T before;
        churn(5000, V_DISPATCH(&py));
        before = private_bytes();
        churn(100000, V_DISPATCH(&py));
        swprintf(got, 256, L"%+.2f MB", ((double)private_bytes() - (double)before) / 1048576.0);
        check(private_bytes() < before + 1048576, L"memory after 100,000 rounds of calls", got);
    }
    VariantClear(&py);

    /* --- the add-in unticked and ticked again: its functions go and come back */
    {
        int (WINAPI *close)(void) = (int (WINAPI *)(void))GetProcAddress(xll, "xlAutoClose");
        int before = nregs;
        nunregs = 0;
        close();
        swprintf(got, 256, L"%d", nunregs);
        check(nunregs == NSLOTS + NBUILTIN, L"xlAutoClose unregisters every function, PY_CALL and the commands", got);
        open();
        swprintf(got, 256, L"%d", nregs - before);
        check(nregs - before == NSLOTS + NBUILTIN, L"xlAutoOpen again registers them again", got);
        a[0] = num(2); a[1] = num(3.5);
        call_check(L"check_add", 2, p, L"5.5");
    }

    /* --- AsfPyReload: a module added, changed (same size, same second) and removed */
    {
        int (WINAPI *reload)(void) = (int (WINAPI *)(void))GetProcAddress(xll, "asfpy_reload");
        XLOPER12 g[1];
        LPXLOPER12 q[1] = { &g[0] };
        const char *v1 = "from asfpy import export, xl_func\n\n@xl_func\ndef reload_check():\n    return 1\n\n"
                         "@export\ndef version():\n    return 1\n";
        const char *v2 = "from asfpy import export, xl_func\n\n@xl_func\ndef reload_check():\n    return 2\n\n"
                         "@export\ndef version():\n    return 2\n";
        int before;
        g[0] = str(L"aa_reload.version");
        check(folder_file(L"aa_reload.py", v1), L"aa_reload.py written in the Python folder", L"");
        before = nunregs;
        check(reload && reload() == 1, L"AsfPyReload", L"1");
        swprintf(got, 256, L"%d", nunregs - before);
        check(nunregs - before == NSLOTS, L"AsfPyReload unregisters the functions before", got);
        call_check(L"reload_check", 0, p, L"1");
        pycall_check(L"PY_CALL(\"aa_reload.version\") after a reload", 1, q, L"1");
        folder_file(L"aa_reload.py", v2);
        if (reload) reload();
        call_check(L"reload_check", 0, p, L"2");
        pycall_check(L"PY_CALL(\"aa_reload.version\") after the module changed", 1, q, L"2");
        folder_file(L"aa_reload.py", NULL);
        if (reload) reload();
        pycall_check(L"PY_CALL(\"aa_reload.version\") after the module went: #NAME?", 1, q, L"#29");
        a[0] = num(2); a[1] = num(3.5);
        call_check(L"check_add", 2, p, L"5.5");
    }

    /* --- a second copy of the add-in, loaded while the first runs Python */
    {
        wchar_t tmp[MAX_PATH], copy[MAX_PATH];
        HMODULE second;
        GetTempPathW(MAX_PATH, tmp);
        swprintf(copy, MAX_PATH, L"%lsasfpy_copy%d.xll", tmp, (int)sizeof(void *) * 8);
        last_alert[0] = 0;
        if (CopyFileW(xll_path, copy, FALSE) && (second = LoadLibraryW(copy)) != NULL) {
            ((int (WINAPI *)(void))GetProcAddress(second, "xlAutoOpen"))();
            check(wcsstr(last_alert, L"another copy of this add-in") != NULL, L"a second copy says that the first runs Python", last_alert);
        } else {
            check(0, L"a second copy", L"can't copy or load it");
        }
    }
    wprintf(L"%d checks, %d failed\n", checks, failed);
    return failed ? 1 : 0;
}
