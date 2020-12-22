// xloper.h - XLOPER related code
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <compare>
#include <concepts>
#include "traits.h"
#include "sref.h"

namespace xll {

	// convertible to double
	inline static const int xltypeNumeric = (xltypeNum | xltypeBool | xltypeInt);
	// do not involve memory allocation
	inline static const int xltypeScalar = (xltypeNumeric | xltypeErr | xltypeMissing | xltypeNil | xltypeSRef);

	template<class X> requires either_base_of_v<XLOPER, XLOPER12, X>
	inline size_t rows(const X& x)
	{
		// ref type???
		return (x.xltype & xltypeMulti) ? x.val.array.rows
			  : x.xltype == xltypeSRef ? x.val.sref.ref.rwLast - x.val.sref.ref.rwFirst + 1
			  : x.xltype == xltypeNil ? 0
			  : 1;
	}

	template<class X> requires either_base_of_v<XLOPER, XLOPER12, X>
	inline size_t columns(const X& x)
	{
		return (x.xltype & xltypeMulti) ? x.val.array.columns 
			  : x.xltype == xltypeSRef ? x.val.sref.ref.colLast - x.val.sref.ref.colFirst + 1
		      : x.xltype == xltypeNil ? 0
			  : 1;
	}
	template<class X> requires either_base_of_v<XLOPER, XLOPER12, X>
	inline size_t size(const X& x)
	{
		return rows(x) * columns(x);
	}
	template<class X> requires either_base_of_v<XLOPER, XLOPER12, X>
	inline X* begin(X& x)
	{
		return (x.xltype & xltypeMulti) ? static_cast<X*>(x.val.array.lparray) : &x;
	}
	template<class X> requires either_base_of_v<XLOPER, XLOPER12, X>
	inline const X* begin(const X& x)
	{
		return (x.xltype & xltypeMulti) ? static_cast<const X*>(x.val.array.lparray) : &x;
	}
	template<class X> requires either_base_of_v<XLOPER, XLOPER12, X>
	inline X * end(X & x)
	{
		return (x.xltype & xltypeMulti) ? static_cast<X*>(x.val.array.lparray + size(x)) : &x + 1;
	}
	template<class X> requires either_base_of_v<XLOPER, XLOPER12, X>
	inline const X* end(const X & x)
	{
		return (x.xltype & xltypeMulti) ? static_cast<const X*>(x.val.array.lparray + size(x)) : &x + 1;
	}

	// one-dimensional index
	template<class X> requires either_base_of_v<XLOPER, XLOPER12, X>
	X& index(X& x, size_t i)
	{
		ensure(i < size(x));

		if (x.xltype & xltypeMulti) {
			return static_cast<X&>(x.val.array.lparray[xmod(i, size(x))]);
		}
		else {
			ensure(i == 0);
			return x;
		}
	}
	// two-dimensional index
	template<class X> requires either_base_of_v<XLOPER, XLOPER12, X>
	X& index(X& x, size_t rw, size_t col)
	{
		return index(x, columns(x) * xmod(rw, rows(x)) + xmod(col, columns(x)));
	}

	// Missing and Nil
#define X(a, b)                                        \
	template<class T>                                  \
		requires either_base_of_v<XLOPER, XLOPER12, T> \
	inline constexpr T X##a = { .xltype = xltype##a }; \
	inline constexpr XLOPER a##4 = X##a<XLOPER>;       \
	inline constexpr XLOPER12 a##12 = X##a<XLOPER12>;  \
	inline constexpr XLOPERX a = X##a<XLOPERX>;        \

	XLL_NULL_TYPE(X)
#undef X

// Error types xll::ErrNA/4/12, ...
#define X(a, b, c)                                            \
	template<class T>                                         \
		requires either_base_of_v<XLOPER, XLOPER12, T>        \
	inline constexpr T XErr##a =                              \
		{ .val = { .err = xlerr##a }, .xltype = xltypeErr };  \
	inline constexpr XLOPER Err##a##4 = XErr##a<XLOPER>;      \
	inline constexpr XLOPER12 Err##a##12 = XErr##a<XLOPER12>; \
	inline constexpr XLOPERX Err##a = XErr##a<XLOPERX>;       \

	XLL_ERR_TYPE(X)
#undef X
}

// use std::strong_ordering!!!
template<typename X> requires std::is_same_v<XLOPER, X> || std::is_same_v<XLOPER12, X>
inline auto operator<=>(const X& x, const X& y)
{
	auto xtype = (x.xltype & ~(xlbitDLLFree | xlbitXLFree));
	auto ytype = (y.xltype & ~(xlbitDLLFree | xlbitXLFree));

	if (xtype != ytype) {
		return xtype <=> ytype;
	}

	switch (xtype) {
	case xltypeNum: // operator<=> is only a partial ordering on doubles
		// IEEE bits agree with floating point order
		union id {
			int64_t i;
			double d;
		} xid, yid;
		xid.d = x.val.num;
		yid.d = y.val.num;

		return xid.i <=> yid.i;

	case xltypeStr:
		//if (x.val.str[0] != y.val.str[0]) return x.val.str[0] <=> y.val.str[0];
		return x.val.str[0] != y.val.str[0]
			? x.val.str[0] <=> y.val.str[0]
			: xll::traits<X>::cmp(x.val.str + 1, y.val.str + 1, x.val.str[0]) <=> 0;

	case xltypeErr:
		return x.val.err <=> y.val.err;

	case xltypeMulti:
		if (x.val.array.rows != y.val.array.rows)
			return x.val.array.rows <=> y.val.array.rows;
		if (x.val.array.columns != y.val.array.columns)
			return x.val.array.columns <=> y.val.array.columns;
		for (size_t i = 0; i < xll::size(x); ++i)
			if (x.val.array.lparray[i] != y.val.array.lparray[i])
				return x.val.array.lparray[i] <=> y.val.array.lparray[i];
		return std::strong_ordering::equal;

	case xltypeSRef:
		return x.val.sref.ref <=> y.val.sref.ref;

		//case xltypeRef:
	case xltypeInt:
		return x.val.w <=> y.val.w;

	case xltypeBool:
		return x.val.xbool <=> y.val.xbool;
	}

	return xtype <=> ytype;
}

#define XLOPER_CMP(op) \
	inline bool operator##op(const XLOPER& x, const XLOPER& y) { return x <=> y op 0; } \
	inline bool operator##op(const XLOPER12& x, const XLOPER12& y) { return x <=> y op 0; } \

XLOPER_CMP(==)
XLOPER_CMP(!=)
XLOPER_CMP(< )
XLOPER_CMP(<=)
XLOPER_CMP(> )
XLOPER_CMP(>=)

#undef XLOPER_CMP
