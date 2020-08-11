// excel.h - Wrapper functions for Excel
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
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
	inline OPER4 Excel4(Args&&... args)
	{
		return XExcel<XLOPER>(std::forward<Args...>(args...));
	}
	template<typename... Args>
	inline OPER12 Excel4(Args&&... args)
	{
		return XExcel<XLOPER12>(std::forward<Args...>(args...));
	}
	// Avoid collision with global Excel function.
	template<typename... Args>
	inline OPER ExcelX(Args&&... args)
	{
		return XExcel<XLOPERX>(std::forward<Args...>(args...));
	}

} // namespace xll
