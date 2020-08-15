// xloper.h - XLOPER related code
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <compare>
#include <concepts>
#include "traits.h"

namespace xll {

	template<class X>
		requires std::is_base_of_v<XLOPER, X> || std::is_base_of_v<XLOPER12, X>
	inline size_t rows(const X& x)
	{
		// ref type???
		return x.xltype == xltypeMulti ? x.val.array.rows 
			 : x.xltype == xltypeNil ? 0
			 : 1;
	}

	template<class X>
		requires std::is_base_of_v<XLOPER, X> || std::is_base_of_v<XLOPER12, X>
	inline size_t columns(const X& x)
	{
		return x.xltype == xltypeMulti ? x.val.array.columns 
			 : x.xltype == xltypeNil ? 0
			 : 1;
	}
	template<class X>
		requires std::is_base_of_v<XLOPER, X> || std::is_base_of_v<XLOPER12, X>
	inline size_t size(const X& x)
	{
		return rows(x) * columns(x);
	}
	template<class X>
		requires std::is_base_of_v<XLOPER, X> || std::is_base_of_v<XLOPER12, X>
	inline X* begin(X& x)
	{
		return x.xltype == xltypeMulti ? static_cast<X*>(x.val.array.lparray) : &x;
	}
	template<class X>
		requires std::is_base_of_v<XLOPER, X> || std::is_base_of_v<XLOPER12, X>
	inline const X* begin(const X& x)
	{
		return x.xltype == xltypeMulti ? static_cast<const X*>(x.val.array.lparray) : &x;
	}
	template<class X>
		requires std::is_base_of_v<XLOPER, X> || std::is_base_of_v<XLOPER12, X>
	inline X * end(X & x)
	{
		return x.xltype == xltypeMulti ? static_cast<X*>(x.val.array.lparray + size(x)) : &x + 1;
	}
	template<class X>
		requires std::is_base_of_v<XLOPER, X> || std::is_base_of_v<XLOPER12, X>
	inline const X* end(const X & x)
	{
		return x.xltype == xltypeMulti ? static_cast<const X*>(x.val.array.lparray + size(x)) : &x + 1;
	}
	// one-dimensional index
	template<class X>
		requires std::is_base_of_v<XLOPER, X> || std::is_base_of_v<XLOPER12, X>
	X& index(X& x, size_t i)
	{
		ensure(i < size(x));

		if (x.xltype == xltypeMulti) {
			return static_cast<X&>(x.val.array.lparray[i]);
		}
		else {
			ensure(i == 0);
			return x;
		}
	}
	template<class X>
		requires std::is_base_of_v<XLOPER, X> || std::is_base_of_v<XLOPER12, X>
	const X& index(const X& x, size_t i)
	{
		return index(x,i);
	}
	// two-dimensional index
	template<class X>
		requires std::is_base_of_v<XLOPER, X> || std::is_base_of_v<XLOPER12, X>
	X& index(X& x, size_t rw, size_t col)
	{
		return index(x, columns(x)* rw + col);
	}
	template<class X>
		requires std::is_base_of_v<XLOPER, X> || std::is_base_of_v<XLOPER12, X>
	const X& index(const X& x, size_t rw, size_t col)
	{
		return index(x, rw, col);
	}

	// Missing and Nil
#define X(a, b)                                        \
	template<class T>                                  \
	requires (std::is_same_v<T, XLOPER> || std::is_same_v<T, XLOPER12>) \
	inline constexpr T X##a = { .xltype = xltype##a }; \
	inline constexpr XLOPER a##4 = X##a<XLOPER>;       \
	inline constexpr XLOPER12 a##12 = X##a<XLOPER12>;  \
	inline constexpr XLOPERX a = X##a<XLOPERX>;        \

	XLL_NULL_TYPE(X)
#undef X

// Error types xll::ErrNAX, ...
#define X(a, b, c)                                            \
	template<class T>                                         \
	requires (std::is_same_v<T, XLOPER> || std::is_same_v<T, XLOPER12>) \
	inline constexpr T XErr##a =                              \
		{ .val = { .err = xlerr##a }, .xltype = xltypeErr };  \
	inline constexpr XLOPER Err##a##4 = XErr##a<XLOPER>;      \
	inline constexpr XLOPER12 Err##a##12 = XErr##a<XLOPER12>; \
	inline constexpr XLOPERX Err##a = XErr##a<XLOPERX>;       \

	XLL_ERR_TYPE(X)
#undef X
}
namespace { // doesn't hide xloper_cmp!!!
	// use std::strong_ordering!!!
	template<typename X, typename Y>
	requires (std::is_same_v<X, XLOPER> && std::is_same_v<Y, XLOPER>) 
		  || (std::is_same_v<X, XLOPER12> && std::is_same_v<Y, XLOPER12>)
	inline std::strong_ordering xloper_cmpx(const X& x, const Y& y)
	{
		if (x.xltype != y.xltype) {
			return x.xltype <=> y.xltype;
		}

		switch (x.xltype) {
		case xltypeNum:
			return x.val.num == y.val.num ? std::strong_ordering::equal
				: x.val.num < y.val.num ? std::strong_ordering::less
				: std::strong_ordering::greater;
		case xltypeStr:
			return x.val.str[0] != y.val.str[0]
				? x.val.str[0] <=> y.val.str[0]
				: xll::traits<X>::cmp(x.val.str + 1, y.val.str + 1, x.val.str[0]) <=> 0;
		case xltypeErr:
			return std::strong_ordering::less; // never equal
		//case xltypeMulti:
		//case xltypeSRef:
		//case xltypeRef:
		case xltypeInt:
			return x.val.w <=> y.val.w;
		case xltypeBool:
			return x.val.xbool <=> y.val.xbool;
		default:
			return x.xltype <=> y.xltype;
		}
	}
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
			return x.xltype - y.xltype;
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
			return x.xltype - y.xltype;
		}
	}
}

// works for any compination of XLOPER and OPER.
#if 0
//namespace xll {

template<typename X, typename Y>
requires (std::is_base_of_v<XLOPER, X> && std::is_base_of_v<XLOPER, Y>)
      || (std::is_base_of_v<XLOPER12, X> && std::is_base_of_v<XLOPER12, Y>)
inline auto operator<=>(const X& x, const Y& y)
{
	if constexpr (std::is_convertible_v<X, XLOPER>) {
		return xloper_cmpx<XLOPER, XLOPER>(x, y);
	}
	else {
		return xloper_cmpx<XLOPER12, XLOPER12>(x, y);
	}
}
template<typename X, typename Y>
requires (std::is_base_of_v<XLOPER, X>&& std::is_base_of_v<XLOPER, Y>)
|| (std::is_base_of_v<XLOPER12, X> && std::is_base_of_v<XLOPER12, Y>)
inline auto operator==(const X & x, const Y & y)
{
	if constexpr (std::is_convertible_v<X, XLOPER>) {
		return xloper_cmpx<XLOPER, XLOPER>(x, y) == 0;
	}
	else {
		return xloper_cmpx<XLOPER12, XLOPER12>(x, y) == 0;
	}
}

//}
#endif
// #if 0
template<typename X, typename Y>
requires (std::is_base_of_v<XLOPER,X> && std::is_base_of_v<XLOPER,Y>)
	  || (std::is_base_of_v<XLOPER12, X>&& std::is_base_of_v<XLOPER12, Y>)
inline bool operator==(const X& x, const Y& y)
{
	return xloper_cmp(x, y) == 0;
}
template<typename X, typename Y>
requires (std::is_base_of_v<XLOPER,X> && std::is_base_of_v<XLOPER,Y>)
      || (std::is_base_of_v<XLOPER12, X>&& std::is_base_of_v<XLOPER12, Y>)
inline bool operator!=(const X& x, const Y& y)
{
	return xloper_cmp(x, y) != 0;
}
template<typename X, typename Y>
requires (std::is_base_of_v<XLOPER, X>&& std::is_base_of_v<XLOPER, Y>)
      || (std::is_base_of_v<XLOPER12, X> && std::is_base_of_v<XLOPER12, Y>)
inline bool operator<(const X& x, const Y& y)
{
	return xloper_cmp(x, y) < 0;
}
template<typename X, typename Y>
requires (std::is_base_of_v<XLOPER, X>&& std::is_base_of_v<XLOPER, Y>)
      || (std::is_base_of_v<XLOPER12, X> && std::is_base_of_v<XLOPER12, Y>)
inline bool operator>=(const X& x, const Y& y)
{
	return xloper_cmp(x, y) >= 0;
}
template<typename X, typename Y>
requires (std::is_base_of_v<XLOPER, X>&& std::is_base_of_v<XLOPER, Y>)
      || (std::is_base_of_v<XLOPER12, X> && std::is_base_of_v<XLOPER12, Y>)
inline bool operator>(const X& x, const Y& y)
{
	return xloper_cmp(x, y) > 0;
}
template<typename X, typename Y>
requires (std::is_base_of_v<XLOPER, X>&& std::is_base_of_v<XLOPER, Y>)
      || (std::is_base_of_v<XLOPER12, X> && std::is_base_of_v<XLOPER12, Y>)
inline bool operator<=(const X& x, const Y& y)
{
	return xloper_cmp(x, y) <= 0;
}
// #endif

