// xloper.h - XLOPER related code
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <concepts>
#include "defines.h"

namespace xll {

// Missing and Nil
#define X(a, b)                                        \
	template<class T>                                  \
	inline constexpr T X##a = { .xltype = xltype##a }; \
	inline constexpr XLOPER a = X##a<XLOPER>;          \
	inline constexpr XLOPER12 a##12 = X##a<XLOPER12>;  \
	inline constexpr XLOPERX a##X = X##a<XLOPERX>;     \

	XLL_NULL_TYPE(X)
#undef X

// Error types
#define X(a, b) \
	template<class T>                                         \
	inline constexpr T XErr##a =                              \
		{ .val = { .err = xlerr##a }, .xltype = xltypeErr };  \
	inline constexpr XLOPER Err##a = XErr##a<XLOPER>;         \
	inline constexpr XLOPER12 Err##a##12 = XErr##a<XLOPER12>; \
	inline constexpr XLOPERX Err##a##X = XErr##a<XLOPERX>;    \

	XLL_ERROR_TYPE(X)
#undef X

}

namespace { // doesn't hide xloper_cmp!!!
	// use std::strong_ordering!!!
	inline int xloper_cmp(const XLOPER& x, const XLOPER& y)
	{
		if (x.xltype != y.xltype) {
			return x.xltype - y.xltype;
		}

		switch (x.xltype) {
		case xltypeNum:
			return x.val.num < y.val.num ? -1 : x.val.num > y.val.num ? 1 : 0;
		case xltypeStr:
			return x.val.str[0] != y.val.str[0]
				? x.val.str[0] - y.val.str[0]
				: xll::traits<XLOPER>::cmp(x.val.str + 1, y.val.str + 1, x.val.str[0]);
		case xltypeErr:
			return INT_MIN; // never equal, like NaN
		//case xltypeMulti:
		//case xltypeSRef:
		//case xltypeRef:
		case xltypeInt:
			return x.val.w - y.val.w;
		case xltypeBool:
			return x.val.xbool - y.val.xbool;
		default:
			return 0;
		}
	}
	inline int xloper_cmp(const XLOPER12& x, const XLOPER12& y)
	{
		if (x.xltype != y.xltype) {
			return x.xltype - y.xltype;
		}

		switch (x.xltype) {
		case xltypeNum:
			return x.val.num < y.val.num ? -1 : x.val.num > y.val.num ? 1 : 0;
		case xltypeStr:
			return x.val.str[0] != y.val.str[0]
				? x.val.str[0] - y.val.str[0]
				: xll::traits<XLOPER12>::cmp(x.val.str + 1, y.val.str + 1, x.val.str[0]);
		case xltypeErr:
			return INT_MIN; // never equal, like NaN
		//case xltypeMulti:
		//case xltypeSRef:
		//case xltypeRef:
		case xltypeInt:
			return x.val.w - y.val.w;
		case xltypeBool:
			return x.val.xbool - y.val.xbool;
		default:
			return INT_MAX;
		}
	}
}

// works for any compination of XLOPER and OPER.
template<typename X, typename Y>
requires (std::is_base_of_v<XLOPER, X>&& std::is_base_of_v<XLOPER, Y>)
|| (std::is_base_of_v<XLOPER12, X> && std::is_base_of_v<XLOPER12, Y>)
inline bool operator==(const X& x, const Y& y)
{
	return xloper_cmp(x, y) == 0;
}
// /* does not compile!!!
template<typename X, typename Y>
requires (std::is_base_of_v<XLOPER,X> && std::is_base_of_v<XLOPER,Y>)
      || (std::is_base_of_v<XLOPER12, X>&& std::is_base_of_v<XLOPER12, Y>)
inline bool operator!=(const X& x, const Y& y)
{
	return xloper_cmp(x, y) != 0;
}
// */
template<typename X, typename Y>
inline bool operator<(const X& x, const Y& y)
{
	return xloper_cmp(x, y) < 0;
}
template<typename X, typename Y>
inline bool operator>=(const X& x, const Y& y)
{
	return xloper_cmp(x, y) >= 0;
}
template<typename X, typename Y>
inline bool operator>(const X& x, const Y& y)
{
	return xloper_cmp(x, y) > 0;
}
template<typename X, typename Y>
inline bool operator<=(const X& x, const Y& y)
{
	return xloper_cmp(x, y) <= 0;
}


