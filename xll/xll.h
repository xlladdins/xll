// xll.h - Excel add-in interface
#pragma once
#include <Windows.h>
#include "XLCALL.H"

// 64-bit uses different name decoration
#ifdef _M_X64 
#define XLL_DECORATE(s,n) s
#define XLL_X64(x) x
#define XLL_X32(x)
#else
#define XLL_DECORATE(s,n) "_" s "@" #n
#define XLL_X64(x)	
#define XLL_X32(x) x
#endif 

#pragma comment(linker, "/include:" XLL_DECORATE("DllMain", 12))
//#pragma comment(linker, "/export:" XLL_DECORATE("XLCallVer", 0))
#pragma comment(linker, "/export:xlAutoOpen" XLL_X32("@0=xlAutoOpen"))
#pragma comment(linker, "/export:xlAutoClose" XLL_X32("@0=xlAutoClose"))
#pragma comment(linker, "/export:xlAutoAdd" XLL_X32("@0=xlAutoAdd"))
#pragma comment(linker, "/export:xlAutoRemove" XLL_X32("@0=xlAutoRemove"))
#pragma comment(linker, "/export:xlAutoFree12" XLL_X32("@4=xlAutoFree12"))
//#pragma comment(linker, "/export:xlAutoRegister12" XLL_X32("@4=xlAutoRegister12"))
//#pragma comment(linker, "/export:xlAddInManagerInfo12@4=xlAddInManagerInfo12")
#ifdef _DEBUG
#pragma comment(linker, "/export:?crtDbg@@3UCrtDbg@@A")
#endif

