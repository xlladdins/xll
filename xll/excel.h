// excel.h - Wrapper functions for Excelv
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <utility>
#include "oper.h"

namespace xll {

	template<typename X, typename... Args>
	inline XOPER<X> XExcel(int xlf, const Args&... args)
	{
		XOPER<X> o;
		std::vector<const X*> xargs { &args... };

		int ret = traits<X>::Excelv(xlf, &o, sizeof...(args), (X**)xargs.data());
		ensure(ret == xlretSuccess); // !!!indicate ref???
		if (!o.is_scalar()) {
			o.xltype |= xlbitXLFree;
		}

		return o;
	}

	template<typename... Args>
	inline OPER4 Excel4(int fn, Args&&... args)
	{
		return XExcel<XLOPER, Args...>(fn, args...);
	}
	template<typename... Args>
	inline OPER12 Excel12(int fn, Args&&... args)
	{
		return XExcel<XLOPER12, Args...>(fn, args...);
	}
	template<typename... Args>
	inline OPER Excel(int fn, Args&&... args)
	{
		return XExcel<XLOPERX, Args...>(fn, args...);
	}

} // namespace xll
