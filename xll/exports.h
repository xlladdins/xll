// exports.h - Functions to export to xlls
#pragma once
#include "defines.h"

#pragma comment(linker, "/include:" XLL_DECORATE("DllMain", 12))
//#pragma comment(linker, "/export:" XLL_DECORATE("XLCallVer", 0))
#pragma comment(linker, "/export:xlAutoOpen" XLL_X32("@0=xlAutoOpen"))
#pragma comment(linker, "/export:xlAutoClose" XLL_X32("@0=xlAutoClose"))
#pragma comment(linker, "/export:xlAutoAdd" XLL_X32("@0=xlAutoAdd"))
#pragma comment(linker, "/export:xlAutoRemove" XLL_X32("@0=xlAutoRemove"))
#pragma comment(linker, "/export:xlAutoFree" XLL_X32("@4=xlAutoFree"))
#pragma comment(linker, "/export:xlAutoFree12" XLL_X32("@4=xlAutoFree12"))
//#pragma comment(linker, "/export:xlAutoRegister12" XLL_X32("@4=xlAutoRegister12"))
//#pragma comment(linker, "/export:xlAddInManagerInfo12@4=xlAddInManagerInfo12")
#ifdef _DEBUG
#pragma comment(linker, "/export:?crtDbg@@3UCrtDbg@@A")
#endif


