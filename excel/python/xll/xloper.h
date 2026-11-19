/* The parts of the Excel C API that asfpy.xll uses, declared from Microsoft's documentation
   ("Excel XLL SDK: XLOPER12", "Excel12/Excel12v", "xlfRegister"). No SDK file is needed. */
#ifndef ASFPY_XLOPER_H
#define ASFPY_XLOPER_H
#include <windows.h>

typedef INT32 RW;
typedef INT32 COL;
typedef WCHAR XCHAR;
typedef DWORD_PTR IDSHEET;

typedef struct xlref12 { RW rwFirst; RW rwLast; COL colFirst; COL colLast; } XLREF12;
typedef struct xlmref12 { WORD count; XLREF12 reftbl[1]; } XLMREF12;

typedef struct xloper12 {
    union {
        double num;
        XCHAR *str;            /* length-prefixed: str[0] is the length */
        BOOL xbool;
        int err;
        int w;
        struct { WORD count; XLREF12 ref; } sref;
        struct { XLMREF12 *lpmref; IDSHEET idSheet; } mref;
        struct { struct xloper12 *lparray; RW rows; COL columns; } array;
        struct {
            union { int level; int tbctrl; IDSHEET idSheet; } valflow;
            RW rw; COL col; BYTE xlflow;
        } flow;
        struct { union { BYTE *lpbData; HANDLE hdata; } h; long cbData; } bigdata;
    } val;
    DWORD xltype;
} XLOPER12, *LPXLOPER12;

#define xltypeNum     0x0001
#define xltypeStr     0x0002
#define xltypeBool    0x0004
#define xltypeRef     0x0008
#define xltypeErr     0x0010
#define xltypeFlow    0x0020
#define xltypeMulti   0x0040
#define xltypeMissing 0x0080
#define xltypeNil     0x0100
#define xltypeSRef    0x0400
#define xltypeInt     0x0800
#define xlbitXLFree   0x1000
#define xlbitDLLFree  0x4000

#define xlerrNull  0
#define xlerrDiv0  7
#define xlerrValue 15
#define xlerrRef   23
#define xlerrName  29
#define xlerrNum   36
#define xlerrNA    42

#define xlSpecial  0x4000
#define xlFree     (0 | xlSpecial)
#define xlGetName  (9 | xlSpecial)
#define xlfRegister 149

#define xlretSuccess 0

/* MdCallBack12, exported by Excel; __stdcall on 32-bit Windows */
typedef int (WINAPI *EXCEL12PROC)(int xlfn, int coper, LPXLOPER12 *rgpxloper12, LPXLOPER12 xloper12Res);

#endif
