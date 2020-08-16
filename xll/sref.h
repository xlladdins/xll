// sref.h - simple reference
#pragma once
#include <compare>
#include "traits.h"
#include "concepts.h"

template<class X>
	//requires std::is_base_of_v<XLREF, X> || std::is_base_of_v<XLREF12, X>
	requires either_base_of_v<XLREF,XLREF12,X>
inline size_t height(const X& x)
{
	return x.rwLast - x.rwFirst + 1;
}
template<class X>
requires either_base_of_v<XLREF, XLREF12, X>
inline size_t width(const X& x)
{
	return x.colLast - x.colFirst + 1;
}
template<class X>
requires either_base_of_v<XLREF, XLREF12, X>
inline size_t size(const X& x) // extent???
{
	return height(x) * width(x);
}

//!!! should be a partial order
template<class X, class Y>
	requires both_derived_from_v<XLREF, X, Y> || both_derived_from_v<XLREF12,X,Y>
inline auto compare_three_way(const X & x, const Y & y)
{
	if (auto cmp = x.rwFirst <=> y.rwFirst; cmp != 0)
		return cmp;
	if (auto cmp = x.rwLast <=> y.rwLast; cmp != 0)
		return cmp;
	if (auto cmp = x.colFirst <=> y.colFirst; cmp != 0)
		return cmp;
	
	return x.colLast <=> y.colLast;
}

// operator<=> only generates two way comparisons for classes
// inheriting from extern "C" XREF subverts auto genereration
#define REF_CTW(op)  template<class X, class Y> \
requires both_derived_from_v<XLREF, X, Y> || both_derived_from_v<XLREF12, X, Y> \
	inline bool operator ## op(const X& x, const Y& y) \
	{ return compare_three_way(x, y) op 0; }

REF_CTW(==)
REF_CTW(!=)
REF_CTW(> )
REF_CTW(>=)
REF_CTW(< )
REF_CTW(<= )

#undef REF_CTW

namespace xll {

	template<class X>
		requires either_base_of_v<XLOPER, XLOPER12, X>
	class XREF : public traits<X>::xref {
		using xrw = typename traits<X>::xrw;
		using xcol = typename traits<X>::xcol;
		using xref = typename traits<X>::xref;
		using xref::xref;
	public:
		XREF()
		{ }
		XREF(xrw row, xcol col, xrw height = 1, xcol width = 1)
			: xref{ row, row + height - 1, col, col + width - 1 }
		{ }
		explicit XREF(const xref& r)
			: xref{ r }
		{ }
		/*
		auto operator<=>(const XREF& r) const
		{
			return static_cast<const xref&>(*this) <=> static_cast<const xref&>(r);
		}
		*/
	};

	using REF4 = XREF<XLOPER>;
	using REF12 = XREF<XLOPER12>;
	using REF = XREF<XLOPERX>;

} // namespace xll
