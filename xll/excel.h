// excel.h - Wrapper functions for Excel
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include "oper.h"

#define ExcelX Excel<XLOPERX>

namespace xll {

	template<typename X, typename... Args>
	inline XOPER<X> Excel(int xlf, const Args&... args)
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

} // namespace xll
