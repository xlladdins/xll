// excel.h - Wrapper functions for Excel
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include "oper.h"

namespace xll {

	template<typename X, typename... Args>
	inline XOPER<X> Excel(int xlf, const Args&... args)
	{
		X x[1];
		std::vector<const X*> xargs { &args... };

		int ret = traits<X>::Excelv(xlf, &x[0], sizeof...(args), (X**)xargs.data());
		ensure(ret == xlretSuccess);

		XOPER<X> o = x[0];
		traits<X>::Excelv(xlFree, 0, 1, (X**)&x);

		return o;
	}

} // namespace xll
