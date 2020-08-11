// excel.h - Wrapper functions for Excel
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <utility>
#include "oper.h"

namespace xll {

	template<typename X, typename... Args>
	inline XOPER<X> XExcel(int xlf, const Args&... args)
	{
		XOPER<X> o;
		X x;
		std::vector<const X*> xargs { &args... };

		int ret = traits<X>::Excelv(xlf, &x, sizeof...(args), (X**)xargs.data());
		if (ret != xlretSuccess) {
			throw ret;
		}
		o = x;
		X* px[1] = { &x };
		traits<X>::Excelv(xlFree, NULL, 1, px);

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
