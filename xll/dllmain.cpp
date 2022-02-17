// dllmain.cpp
// Copyright (c) KALX, LLC. All rights reserved. No warranty is made.
#include <Windows.h>

//HINSTANCE xll_hModule;

#pragma warning(disable: 4100)
extern "C"
BOOL WINAPI
DllMain(HINSTANCE hDLL, ULONG reason, LPVOID lpReserved)
{
	switch (reason) {
	case DLL_PROCESS_ATTACH:
		//xll_hModule = hDLL;
		DisableThreadLibraryCalls(hDLL);
		break;
	case DLL_THREAD_ATTACH:
		break;
	case DLL_THREAD_DETACH:
		break;
	case DLL_PROCESS_DETACH:
		break;
	}

	return TRUE;
}
