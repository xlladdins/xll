// auto.cpp - Implement well known Excel interfaces.
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#include "xll.h"

// Called by Excel when the xll is opened.
extern "C" 
int __declspec(dllexport) WINAPI
xlAutoOpen(void)
{
	for (auto& [key, arg] : xll::AddIn::Map) {
		arg.Register();
	}
	for (auto& [key, arg] : xll::AddIn12::Map) {
		arg.Register();
	}

	return TRUE;
}

// Called by Microsoft Excel when the xll is deactivated.
extern "C"
int __declspec(dllexport) WINAPI
xlAutoClose(void)
{
	// No need to call xlfUnregister, just delete names.
	for (const auto& [key, args] : xll::AddIn12::Map) {
		Excel<XLOPER12>(xlfSetName, key);
	}
	for (const auto& [key, args] : xll::AddIn::Map) {
		Excel<XLOPER>(xlfSetName, key);
	}
	
	return TRUE;
}

// Called when the user activates the xll using the Add-In Manager. 
// This function is *not* called when Excel starts up.
extern "C"
int __declspec(dllexport) WINAPI
xlAutoAdd(void)
{
	return TRUE;
}

// Called when user deactivates the xll using the Add-In Manager. 
// This function is *not* called when an Excel session closes.
extern "C"
int __declspec(dllexport) WINAPI
xlAutoRemove(void)
{
	return TRUE;
}

// Called just after a worksheet function returns an XLOPER/XLOPER12 with xltype&xlbitDLLFree set.
extern "C"
void __declspec(dllexport) WINAPI
xlAutoFree(LPXLOPER px)
{
	if (px->xltype & xlbitDLLFree) {
		if (px->xltype & xltypeStr) {
			free(px->val.str);
		}
		else if (px->xltype & xltypeMulti) {
			free(px->val.array.lparray);
		}
	}
}
extern "C"
void __declspec(dllexport) WINAPI
xlAutoFree12(LPXLOPER12 px)
{
	if (px->xltype & xlbitDLLFree) {
		if (px->xltype & xltypeStr) {
			free(px->val.str);
		}
		else if (px->xltype & xltypeMulti) {
			free(px->val.array.lparray);
		}
	}
}

// Look up name and register.
extern "C"
LPXLOPER __declspec(dllexport) WINAPI
xlAutoRegister(LPXLOPER pxName)
{
	return xll::AddIn::Map[*pxName].Register();
}

// Look up name and register.
extern "C"
LPXLOPER12 __declspec(dllexport) WINAPI 
xlAutoRegister12(LPXLOPER12 pxName)
{
	return xll::AddIn12::Map[*pxName].Register();
}

// Name of dll with suffix stripped off.
template<class X>
xll::XOPER<X> DllName()
{
	xll::XOPER<X> xName = Excel<X>(xlfGetName);

	return xName;
}


// Called by Microsoft Excel when the Add-in Manager is invoked for the first time.
// This function is used to provide the Add-In Manager with information about your add-in.
LPXLOPER12 WINAPI xlAddInManagerInfo12(LPXLOPER12 pxAction)
{
	static XLOPER12 o;

	// coerce to int and check if action is 1
	// return string indicating xll name
	XLOPER12 xResult;
	XLOPER12 xInt = { .val = { .w = xltypeInt }, .xltype = xltypeInt };
	int ret = Excel12(xlCoerce, &xResult, 2, pxAction, &xInt);
	if (ret == xlretSuccess && xResult.xltype == xltypeInt && xResult.val.w == 1) {
		o.xltype = xltypeStr;
		o.val.str = (XCHAR*)L"\06Add-In"; // use xlGetName!!!
	}
	else {
		o = xll::ErrValue12;
	}

	return &o;
}