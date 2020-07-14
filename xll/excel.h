// excel.h - Wrapper functions for Excel
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include "oper.h"

namespace xll {

	template<typename X, typename... Args>
	inline XOPER<X> Excel(int xlf, const Args&... args)
	{
		X x;
		std::vector<const X*> xargs { &args... };

		int ret = traits<X>::Excelv(xlf, &x, sizeof...(args), (X**)xargs.data());
		if (ret != xlretSuccess) {
			return XErrValue<X>;
		}

		XOPER<X> o = x;
		traits<X>::Excelv(xlFree, NULL, 0, NULL);

		return o;
	}

} // namespace xll
