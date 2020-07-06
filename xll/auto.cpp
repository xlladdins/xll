// auto.cpp - Implement well known Excel interfaces.
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#include "xll.h"

// Called by Excel when the xll is opened.
extern "C" 
int __declspec(dllexport) WINAPI
xlAutoOpen(void)
{
	/*
	for (const auto& [key, args] : xll::AddIns::Map) {
		args.Register();
	}
	*/
    return TRUE;
}

// Called by Microsoft Excel whenever the xll is deactivated.
extern "C"
int __declspec(dllexport) WINAPI
xlAutoClose(void)
{
	return TRUE;
}

// Called when the user activates the xll using the Add-In Manager. 
// This function is not called when Excel starts up and loads a pre-installed add-in.
extern "C"
int __declspec(dllexport) WINAPI
xlAutoAdd(void)
{
	return TRUE;
}

// Called when user deactivates the xll using the Add-In Manager. 
// This function is not called when an Excel session closes, normally or abnormally, with the add-in installed.
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
/*
// Look up name and register.
extern "C"
LPXLOPER12 __declspec(dllexport) WINAPI 
xlAutoRegister12(LPXLOPER12 pxName)
{
	static XLOPER12 xResult;

	return &xResult;
}
*/

static const XLOPER12 xErrValue = XLOPER12{ .val = { .err = xlerrValue }, .xltype = xltypeErr };

LPXLOPER12 GetName()
{
	static XLOPER12 xName;

	Excel12(xlGetName, &xName, 0);

	return &xName; // needs to be xlFree'd
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
		o = xErrValue;
	}

	return &o;
}