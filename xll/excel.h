// excel.h - Wrapper functions for Excel* functions.
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include "oper.h"

namespace xll {

	template<typename... Args>
	inline XLOPER Excel(int xlf, const Args&... args)
	{
		XLOPER o;

		::Excel(xlf, &o, sizeof...(args), &OPER(args)...);

		return o;
	}

	template<typename... Args>
	inline OPER12 Excel(int xlf, const Args&... args)
	{
		OPER12 o;

        try {
            XLOPER12 o_;
		    int ret = ::Excel12(xlf, &o_, sizeof...(args), &args...);
		    ensure (ret == xlretSuccess);
            o = o_;
		    if (!o.isScalar()) {
                ::Excel12(xlFree, 0, 1, &o_);
            }
        }
        catch (const std::exception& ex) {
            MessageBoxA(GetForegroundWindow(), ex.what(), "Excel failed", MB_OKCANCEL| MB_ICONERROR);
        }
		return o;
	}

} // namespace xll

// Just like Excel
/*
inline xll::OPER12 operator&(const xll::OPER12& a, const xll::OPER12& b)
{
	return xll::Excel(xlfConcatenate, a, b);
}
*/
