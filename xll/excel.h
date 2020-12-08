// excel.h - Wrapper functions for Excelv
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <array>
#include <utility>
#include "oper.h"

namespace xll {

	template<typename X>
	inline XOPER<X> XExcelv(int xlfn, size_t n, X* opers[])
	{
		XOPER<X> o;

		int ret = traits<X>::Excelv(xlfn, &o, static_cast<int>(n), &opers[0]);
		ensure(ret == xlretSuccess); // !!!indicate ret???
		if (!o.is_scalar()) {
			o.xltype |= xlbitXLFree;
		}

		return o;
	}

	template<typename X, typename... Args>
	inline XOPER<X> XExcel(int xlfn, const Args&... args)
	{
		std::array<const X*,sizeof...(Args)> xargs = { &args... };

		return XExcelv<X>(xlfn, sizeof...(args), (X**)xargs.data());
	}

	inline OPER4 Excelv(int xlfn, size_t n, XLOPER* opers[])
	{
		return XExcelv<XLOPER>(xlfn, n, opers);
	}
	inline OPER12 Excelv(int xlfn, size_t n, XLOPER12* opers[])
	{
		return XExcelv<XLOPER12>(xlfn, n, opers);
	}
	/*
	inline OPER Excelv(int xlfn, size_t n, XLOPERX* opers[])
	{
		return XExcelv<XLOPERX>(xlfn, n, opers);
	}
	*/

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
