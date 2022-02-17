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
	inline static const int xltypeScalar = (xltypeNumeric | xltypeErr | xltypeMissing | xltypeNil | xltypeSRef | xltypeRef);
	// turn off xlbit flags
	inline static const int xlbitmask = ~(xlbitXLFree | xlbitDLLFree);

	template<class X> requires either_base_of_v<XLOPER, XLOPER12, X>
	inline int type(const X& x)
	{
		return x.xltype & xlbitmask;
	}
		
	template<class X> requires either_base_of_v<XLOPER, XLOPER12, X>
	inline unsigned rows(const X& x)
	{
		switch (type(x)) {
		case xltypeMulti:
			return x.val.array.rows;
		case xltypeSRef:
			return height(x.val.sref.ref);
		case xltypeRef:
			return height(x.val.mref.lpmref->reftbl[0]);
		case xltypeNil:
			return 0;
		}
		
		return 1;
	}

	template<class X> requires either_base_of_v<XLOPER, XLOPER12, X>
	inline unsigned columns(const X& x)
	{
		switch (type(x)) {
		case xltypeMulti:
			return x.val.array.columns;
		case xltypeSRef:
			return width(x.val.sref.ref);
		case xltypeRef:
			return width(x.val.mref.lpmref->reftbl[0]);
		case xltypeNil:
			return 0;
		}

		return 1;
	}

	template<class X> requires either_base_of_v<XLOPER, XLOPER12, X>
	inline unsigned size(const X& x)
	{
		return rows(x) * columns(x);
	}

	template<class X> requires either_base_of_v<XLOPER, XLOPER12, X>
	inline X* multi_begin(X& x)
	{
		return type(x) == xltypeMulti ? static_cast<X*>(x.val.array.lparray) : &x;
	}
	template<class X> requires either_base_of_v<XLOPER, XLOPER12, X>
	inline const X* multi_begin(const X& x)
	{
		return type(x) == xltypeMulti ? static_cast<const X*>(x.val.array.lparray) : &x;
	}
	template<class X> requires either_base_of_v<XLOPER, XLOPER12, X>
	inline X * multi_end(X & x)
	{
		return type(x) == xltypeMulti ? static_cast<X*>(x.val.array.lparray + size(x)) : &x + 1;
	}
	template<class X> requires either_base_of_v<XLOPER, XLOPER12, X>
	inline const X* multi_end(const X & x)
	{
		return type(x) == xltypeMulti ? static_cast<const X*>(x.val.array.lparray + size(x)) : &x + 1;
	}

	// ref iterator
	template<class X> requires either_base_of_v<XLOPER, XLOPER12, X>
	inline typename const traits<X>::xref* ref_begin(const X& x)
	{
		return type(x) == xltypeRef ? x.val.mref.lpmref->reftbl : nullptr;
	}
	template<class X> requires either_base_of_v<XLOPER, XLOPER12, X>
	inline typename const traits<X>::xref* ref_end(const X& x)
	{
		return type(x) == xltypeRef ? x.val.mref.lpmref->reftbl + x.val.mref.lpmref->count : nullptr;
	}

	template<class X> requires either_base_of_v<XLOPER, XLOPER12, X>
	inline X* begin(X& x)
	{
		return multi_begin(x);
	}
	template<class X> requires either_base_of_v<XLOPER, XLOPER12, X>
	inline const X* begin(const X& x)
	{
		return multi_begin(x);
	}
	template<class X> requires either_base_of_v<XLOPER, XLOPER12, X>
	inline X* end(X& x)
	{
		return multi_end(x);
	}
	template<class X> requires either_base_of_v<XLOPER, XLOPER12, X>
	inline const X* end(const X& x)
	{
		return multi_end(x);
	}

	// one-dimensional index
	template<class X> requires either_base_of_v<XLOPER, XLOPER12, X>
	inline X& index(X& x, unsigned i)
	{
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
	inline X& index(X& x, unsigned rw, unsigned col)
	{
		return index(x, columns(x) * xmod(rw, rows(x)) + xmod(col, columns(x)));
	}

	// multi-dimensional index???

	// drop from front (n > 0) or back (n < 0)
	template<class X>
		requires (std::is_same_v<XLOPER, X> || std::is_same_v<XLOPER12, X>)
	inline X drop(X x, int n)
	{
		ensure(x.xltype == xltypeMulti);

		if (n > 0) { // front}
			unsigned n_ = n;
			if (rows(x) == 1) {
				if (n_ >= columns(x)) n_ = columns(x);
				x.val.array.lparray += n_;
				x.val.array.columns -= n_;
			}
			else { // top
				if (n_ >= rows(x)) n_ = rows(x);
				x.val.array.lparray += n_ * columns(x);
				x.val.array.rows -= n_;
			}
		}
		else if (n < 0) { // back
			unsigned n_ = -n;
			if (rows(x) == 1) {
				if (n_ >= columns(x)) n_ = columns(x);
				x.val.array.columns -= n_;
			}
			else { // bottom
				if (n_ >= rows(x)) n_ = rows(x);
				x.val.array.rows -= n_;
			}
		}

		return x;
	}

	// take from front (n > 0) or back (n < 0)
	template<class X> 
		requires (std::is_same_v<XLOPER, X> || std::is_same_v<XLOPER12, X>)
	inline X take(X x, int n)
	{
		ensure(x.xltype == xltypeMulti);

		if (n >= 0) { // front
			unsigned n_ = n;
			if (rows(x) == 1) {
				if (n_ > columns(x)) n_ = columns(x);
				x.val.array.columns = n_;
			}
			else { // top
				if (n_ > rows(x)) n_ = rows(x);
				x.val.array.rows = n_;
			}
		}
		else if (n < 0) { // back
			unsigned n_ = -n;
			if (rows(x) == 1) {
				if (n_ > columns(x)) n_ = columns(x);
				x.val.array.lparray = end(x) - n_;
				x.val.array.columns = n_;
			}
			else { // bottom
				if (n_ > rows(x)) n_ = rows(x);
				x.val.array.lparray = end(x) - n_ * columns(x);
				x.val.array.rows = n_;
			}
		}

		return x;
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

	XLL_ERR(X)
#undef X
}

template<typename X> requires std::is_same_v<XLOPER, X> || std::is_same_v<XLOPER12, X>
inline auto operator<=>(const X& x, const X& y)
{
	auto xtype = xll::type(x);
	auto ytype = xll::type(y);

	// xltypeNil is least type
	if (xtype == xltypeNil or ytype == xltypeNil) {
		return (xtype != xltypeNil) <=> (ytype != xltypeNil);
	}

	if (xtype != ytype) {
		return xtype <=> ytype;
	}

	switch (xtype) {
	case xltypeNum: { // operator<=> is only a partial ordering on doubles
		// IEEE bits agree with floating point order
		union id {
			int64_t i;
			double d;
		} xid, yid;
		xid.d = x.val.num;
		yid.d = y.val.num;

		return xid.i <=> yid.i;
	}
	case xltypeStr: {
		return std::lexicographical_compare_three_way(
			x.val.str + 1, x.val.str + 1 + x.val.str[0],
			y.val.str + 1, y.val.str + 1 + y.val.str[0],
			[](int cx, int cy) {
				int ux = std::toupper(cx);
				int uy = std::toupper(cy);
				return ux <=> uy;
			}
		);
	}
	case xltypeErr:
		return x.val.err <=> y.val.err; // false???

	case xltypeMulti: {
		if (auto cmp = x.val.array.rows <=> y.val.array.rows; cmp != 0)
			return cmp;
		if (auto cmp = x.val.array.columns <=> y.val.array.columns; cmp != 0)
			return cmp;
		for (unsigned i = 0; i < xll::size(x); ++i)
			if (auto cmp = x.val.array.lparray[i] <=> y.val.array.lparray[i]; cmp != 0)
				return cmp;

		return std::strong_ordering::equal;
	}
	case xltypeSRef:
		return x.val.sref.ref <=> y.val.sref.ref;
	case xltypeRef:
		if (auto cmp = x.val.mref.idSheet <=> y.val.mref.idSheet; cmp != 0)
			return cmp;
		if (auto cmp = x.val.mref.lpmref->count <=> y.val.mref.lpmref->count; cmp != 0)
			return cmp;
		for (unsigned i = 0; i < x.val.mref.lpmref->count; ++i)
			if (auto cmp = x.val.mref.lpmref->reftbl[i] <=> y.val.mref.lpmref->reftbl[i]; cmp != 0)
				return cmp;

		return std::strong_ordering::equal;
	case xltypeInt:
		return x.val.w <=> y.val.w;
	case xltypeBool:
		return x.val.xbool <=> y.val.xbool;
	case xltypeBigData: // ??? what is bidata.cbData if HANDLE
		if (auto cmp = x.val.bigdata.cbData <=> y.val.bigdata.cbData; cmp != 0)
			return cmp;

		int cmp = std::memcmp(x.val.bigdata.h.lpbData, y.val.bigdata.h.lpbData, x.val.bigdata.cbData);
		
		return cmp <=> 0;
	}

	return std::strong_ordering::equal;
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
