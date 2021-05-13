// exports.h - Functions to export to xlls
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include "defines.h"

// Don't export for xll.t project
#if !defined(_CONSOLE)

// Used to export undecorated function name from a dll.
// Put '#pragma XLLEXPORT' in every add-in function body.
#define XLLEXPORT comment(linker, "/export:" __FUNCDNAME__ "=" __FUNCTION__)

#pragma comment(lib, "Shlwapi.lib")
// xlAuto functions required by Excel
#pragma comment(linker, "/include:" XLL_DECORATE("DllMain", 12))
#pragma comment(linker, "/export:" XLL_DECORATE("XLCallVer", 0))
#pragma comment(linker, "/export:xlAutoOpen" XLL_X32("@0=xlAutoOpen"))
#pragma comment(linker, "/export:xlAutoClose" XLL_X32("@0=xlAutoClose"))
#pragma comment(linker, "/export:xlAutoAdd" XLL_X32("@0=xlAutoAdd"))
#pragma comment(linker, "/export:xlAutoRemove" XLL_X32("@0=xlAutoRemove"))
#pragma comment(linker, "/export:xlAutoFree" XLL_X32("@4=xlAutoFree"))
#pragma comment(linker, "/export:xlAutoFree12" XLL_X32("@4=xlAutoFree12"))
//#pragma comment(linker, "/export:xlAutoRegister" XLL_X32("@4=xlAutoRegister"))
//#pragma comment(linker, "/export:xlAutoRegister12" XLL_X32("@4=xlAutoRegister12"))
//#pragma comment(linker, "/export:xlAddInManagerInfo" XLL_X32("@4=xlAddInManagerInfo"))
//#pragma comment(linker, "/export:xlAddInManagerInfo12" XLL_X32("@4=xlAddInManagerInfo12"))
//#pragma comment(linker, "/export:" XLL_DECORATE("xll_replace_eq_by_eq", 0))
//#pragma comment(linker, "/export:" XLL_DECORATE("xll_paste_basic", 0))
//#pragma comment(linker, "/export:" XLL_DECORATE("xll_paste_create", 0))
//#pragma comment(linker, "/export:" XLL_DECORATE("xll_spreadsheet_doc", 0))

#ifndef _LIB
//#pragma comment(linker, "/include:" XLL_DECORATE("xll_replace_eq_by_eq", 0))
//#pragma comment(linker, "/include:" XLL_DECORATE("xll_paste_basic", 0))
//#pragma comment(linker, "/include:" XLL_DECORATE("xll_paste_create", 0))
//#pragma comment(linker, "/include:" XLL_DECORATE("xll_spreadsheet_doc", 0))
#endif // _LIB

#endif // CONSOLE

#ifdef _DEBUG
#pragma comment(linker, "/export:?crtDbg@@3UCrtDbg@@A")
#endif // _DEBUG

