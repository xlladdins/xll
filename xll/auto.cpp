// auto.cpp - Implement well known Excel interfaces.
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
// https://docs.microsoft.com/en-us/office/client-developer/excel/creating-xlls
#include "xll.h"

using namespace xll;

// Called by Excel when the xll is opened.
extern "C" int __declspec(dllexport) WINAPI
xlAutoOpen(void)
{
	try {
		if (!Auto<Open>::Call()) {
			return FALSE;
		}

		for (auto& [key, arg] : AddIn4::Map) {
			arg.Register();
		}
		for (auto& [key, arg] : AddIn12::Map) {
			arg.Register();
		}

		if (!Auto<OpenAfter>::Call()) {
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
extern "C" int __declspec(dllexport) WINAPI
xlAutoClose(void)
{
	try {
		if (!Auto<CloseBefore>::Call()) {
			return FALSE;
		}

		// No need to call xlfUnregister, just delete names.
		for (const auto& [key, args] : AddIn12::Map) {
			Excel12(xlfSetName, key);
		}
		for (const auto& [key, args] : AddIn4::Map) {
			Excel4(xlfSetName, key);
		}

		if (!Auto<Close>::Call()) {
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
extern "C" int __declspec(dllexport) WINAPI
xlAutoAdd(void)
{
	try {
		if (!Auto<Add>::Call()) {
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
extern "C" int __declspec(dllexport) WINAPI
xlAutoRemove(void)
{
	try {
		if (!Auto<Remove>::Call()) {
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

// Called just after a worksheet function returns an XLOPER with xltype&xlbitDLLFree set.
extern "C" void __declspec(dllexport) WINAPI
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
// Called just after a worksheet function returns an XLOPER12 with xltype&xlbitDLLFree set.
extern "C" void __declspec(dllexport) WINAPI
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
extern "C" LPXLOPER __declspec(dllexport) WINAPI
xlAutoRegister(LPXLOPER pxName)
{
	static XLOPER o;

	try {
		auto px = AddIn4::Map.find(*pxName);
		ensure(px != AddIn4::Map.end());
		// returns an xltypeNum or xltypeErr
		o = *px->second.Register();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = ErrNA4;
	}
	catch (...) {
		XLL_ERROR("Unknown exception in xlAutoRegister");

		o = ErrNA4;
	}

	return &o;
}

// Look up name and register.
extern "C" LPXLOPER12 __declspec(dllexport) WINAPI 
xlAutoRegister12(LPXLOPER12 pxName)
{
	static XLOPER12 o;
	
	try {
		auto px = AddIn12::Map.find(*pxName);
		ensure(px != AddIn12::Map.end());
		// returns an xltypeNum or xltypeErr
		o = *px->second.Register();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = ErrNA12;
	}
	catch (...) {
		XLL_ERROR("Unknown exception in xlAutoRegister12");

		o = ErrNA12;
	}

	return &o;
}


// Called by Microsoft Excel when the Add-in Manager is invoked for the first time.
// This function is used to provide the Add-In Manager with information about your add-in.
extern "C" LPXLOPER12 WINAPI xlAddInManagerInfo12(LPXLOPER12 pxAction)
{
	static XLOPER12 o;

	try {
		// coerce to int and check if action is 1
		OPER12 xResult_ = Excel12(xlCoerce, *pxAction, OPER12(xltypeInt));
		if (xResult_ == 1) {
			o = Excel12(xlGetName);
			// return string to use
			int b = 1;
			int e = 1;
			for (int i = 1; i <= o.val.str[0]; ++i) {
				if (o.val.str[i] == '\\') {
					b = i + 1;
				}
				else if (o.val.str[i] == '.') {
					e = i;
				}
			}
			ensure(b < e);
			o = Excel12(xlfMid, o, OPER12(b), OPER12(e - b));
		}
		else {
			o = ErrNA12;
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = ErrNA12;
	}
	catch (...) {
		XLL_ERROR("Unknown exception in xlAddInManagerInfo12");

		o = ErrNA12;
	}

	return &o;
}
