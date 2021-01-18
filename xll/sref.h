// sref.h - simple reference
#pragma once
#include <compare>
#include "traits.h"
#include "concepts.h"

template<class X> requires either_base_of_v<XLREF,XLREF12,X>
inline unsigned height(const X& x)
{
	return x.rwLast - x.rwFirst + 1;
}
template<class X> requires either_base_of_v<XLREF, XLREF12, X>
inline unsigned width(const X& x)
{
	return x.colLast - x.colFirst + 1;
}
template<class X> requires either_base_of_v<XLREF, XLREF12, X>
inline unsigned size(const X& x) // extent???
{
	return height(x) * width(x);
}

//!!! should be a partial order
template<class X> requires std::is_same_v<XLREF, X> || std::is_same_v<XLREF12,X>
inline auto operator<=>(const X & x, const X & y)
{
	if (auto cmp = x.rwFirst <=> y.rwFirst; cmp != 0)
		return cmp;
	if (auto cmp = x.rwLast <=> y.rwLast; cmp != 0)
		return cmp;
	if (auto cmp = x.colFirst <=> y.colFirst; cmp != 0)
		return cmp;
	
	return x.colLast <=> y.colLast;
}

// operator<=> only generates two way comparisons for C++ classes
// inheriting from extern "C" XLREF subverts auto genereration
#define REF_CTW(op) \
	inline bool operator ## op(const XLREF& x, const XLREF& y) { return x <=> y op 0; } \
	inline bool operator ## op(const XLREF12& x, const XLREF12& y) { return x <=> y op 0; } \

REF_CTW(==)
REF_CTW(!=)
REF_CTW(> )
REF_CTW(>=)
REF_CTW(< )
REF_CTW(<=)

#undef REF_CTW

namespace xll {

	template<class X> requires either_base_of_v<XLOPER, XLOPER12, X>
	class XREF : public traits<X>::xref {
		using xrw = typename traits<X>::xrw;
		using xcol = typename traits<X>::xcol;
		using xref = typename traits<X>::xref;
	public:
		XREF()
		{ }
		XREF(xrw row, xcol col, xrw height = 1, xcol width = 1)
			: xref{ row, static_cast<xrw>(row + height - 1), col, static_cast<xcol>(col + width - 1) }
		{ }
		explicit XREF(const xref& r)
			: xref{ r }
		{ }
	};

	using REF4  = XREF<XLOPER>;
	using REF12 = XREF<XLOPER12>;
	using REF   = XREF<XLOPERX>;

} // namespace xll
