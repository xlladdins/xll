// auto.cpp - Implement well known Excel interfaces.
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#include "xll.h"

// Called by Excel when the xll is opened.
extern "C" 
int __declspec(dllexport) WINAPI
xlAutoOpen(void)
{
	try {
		if (!xll::Auto<xll::Open>::Call()) {
			return FALSE;
		}

		for (auto& [key, arg] : xll::AddIn4::Map) {
			arg.Register();
		}
		for (auto& [key, arg] : xll::AddIn12::Map) {
			arg.Register();
		}

		if (!xll::Auto<xll::OpenAfter>::Call()) {
			return FALSE;
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}
	catch (...) {
		XLL_ERROR("Unknown exception in xlAutoOpen");

		return FALSE;
	}

	return TRUE;
}

// Called by Microsoft Excel when the xll is deactivated.
extern "C"
int __declspec(dllexport) WINAPI
xlAutoClose(void)
{
	try {
		if (!xll::Auto<xll::CloseBefore>::Call()) {
			return FALSE;
		}

		// No need to call xlfUnregister, just delete names.
		for (const auto& [key, args] : xll::AddIn12::Map) {
			::Excel12(xlfSetName, 0, 1, &key);
		}
		for (const auto& [key, args] : xll::AddIn4::Map) {
			::Excel4(xlfSetName, 0, 1, &key);
		}

		if (!xll::Auto<xll::Close>::Call()) {
			return FALSE;
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}
	catch (...) {
		XLL_ERROR("Unknown exception in xlAutoClose");

		return FALSE;
	}

	return TRUE;
}

// Called when the user activates the xll using the Add-In Manager. 
// This function is *not* called when Excel starts up.
extern "C"
int __declspec(dllexport) WINAPI
xlAutoAdd(void)
{
	try {
		if (!xll::Auto<xll::Add>::Call()) {
			return FALSE;
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}
	catch (...) {
		XLL_ERROR("Unknown exception in xlAutoAdd");

		return FALSE;
	}

	return TRUE;
}

// Called when user deactivates the xll using the Add-In Manager. 
// This function is *not* called when an Excel session closes.
extern "C"
int __declspec(dllexport) WINAPI
xlAutoRemove(void)
{
	try {
		if (!xll::Auto<xll::Remove>::Call()) {
			return FALSE;
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}
	catch (...) {
		XLL_ERROR("Unknown exception in xlAutoRemove");

		return FALSE;
	}

	return TRUE;
}

// Called just after a worksheet function returns an XLOPER/XLOPER12 with xltype&xlbitDLLFree set.
extern "C"
void __declspec(dllexport) WINAPI
xlAutoFree(LPXLOPER px)
{
	try {
		if (px->xltype & xlbitDLLFree) {
			if (px->xltype & xltypeStr) {
				free(px->val.str);
			}
			else if (px->xltype & xltypeMulti) {
				free(px->val.array.lparray);
			}
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR("Unknown exception in xlAutoFree");
	}
}
extern "C"
void __declspec(dllexport) WINAPI
xlAutoFree12(LPXLOPER12 px)
{
	try {
		if (px->xltype & xlbitDLLFree) {
			if (px->xltype & xltypeStr) {
				free(px->val.str);
			}
			else if (px->xltype & xltypeMulti) {
				free(px->val.array.lparray);
			}
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR("Unknown exception in xlAutoFree12");
	}
}

// Look up name and register.
extern "C"
LPXLOPER __declspec(dllexport) WINAPI
xlAutoRegister(LPXLOPER pxName)
{
	static XLOPER o;

	try {
		o = *xll::AddIn4::Map[*pxName].Register();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = { .val = { .err = xlerrValue }, .xltype = xltypeErr };
	}
	catch (...) {
		XLL_ERROR("Unknown exception in xlAutoRegister");

		o = { .val = { .err = xlerrValue }, .xltype = xltypeErr };
	}

	return &o;
}

// Look up name and register.
extern "C"
LPXLOPER12 __declspec(dllexport) WINAPI 
xlAutoRegister12(LPXLOPER12 pxName)
{
	static XLOPER12 o;
	
	try {
		o = *xll::AddIn12::Map[*pxName].Register();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = { .val = { .err = xlerrValue }, .xltype = xltypeErr };
	}
	catch (...) {
		XLL_ERROR("Unknown exception in xlAutoRegister12");

		o = { .val = { .err = xlerrValue }, .xltype = xltypeErr };
	}

	return &o;
}


// Called by Microsoft Excel when the Add-in Manager is invoked for the first time.
// This function is used to provide the Add-In Manager with information about your add-in.
LPXLOPER12 WINAPI xlAddInManagerInfo12(LPXLOPER12 pxAction)
{
	static XLOPER12 o;

	try {
		// coerce to int and check if action is 1
		// return string indicating xll name
		XLOPER12 xResult;
		XLOPER12 xInt = { .val = { .w = xltypeInt }, .xltype = xltypeInt };
		int ret = ::Excel12(xlCoerce, &xResult, 2, pxAction, &xInt);
		if (ret == xlretSuccess && xResult.xltype == xltypeInt && xResult.val.w == 1) {
			o.xltype = xltypeStr;
			o.val.str = (XCHAR*)L"\06Add-In"; // use xlGetName!!!
		}
		else {
			o = { .val = { .err = xlerrValue }, .xltype = xltypeErr };
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = { .val = { .err = xlerrValue }, .xltype = xltypeErr };
	}
	catch (...) {
		XLL_ERROR("Unknown exception in xlAddInManagerInfo12");

		o = { .val = { .err = xlerrValue }, .xltype = xltypeErr };
	}

	return &o;
}