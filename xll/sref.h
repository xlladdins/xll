// sref.h - simple reference
#pragma once
#include <compare>
#include "traits.h"

// B is base of both D1 and D2
template<class B, class D1, class D2, class>
concept both_base_of = std::is_base_of_v<B, D1> && std::is_base_of_v<B, D2>;
// all_base_iof<B, class... D>

template<class X>
	requires std::is_base_of_v<XLREF, X> || std::is_base_of_v<XLREF12, X>
inline size_t height(const X& x)
{
	return x.rwLast - x.rwFirst + 1;
}
template<class X>
requires std::is_base_of_v<XLREF, X> || std::is_base_of_v<XLREF12, X>
inline size_t width(const X& x)
{
	return x.colLast - x.colFirst + 1;
}
template<class X>
requires std::is_base_of_v<XLREF, X> || std::is_base_of_v<XLREF12, X>
inline size_t size(const X& x) // extent???
{
	return height(x) * width(x);
}

template<class X, class Y>
requires (std::is_base_of_v<XLREF, X>&& std::is_base_of_v<XLREF, Y>)
|| (std::is_base_of_v<XLREF12, X> && std::is_base_of_v<XLREF12, X>)
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

//#define REF_CTW(op) 
//REF_CTW(==)

template<class X, class Y>
	requires (std::is_base_of_v<XLREF, X> && std::is_base_of_v<XLREF, Y>)
	|| (std::is_base_of_v<XLREF12, X> && std::is_base_of_v<XLREF12, X>)
inline bool operator==(const X& x, const Y& y)
{
	return compare_three_way(x, y) == 0;
}
template<class X, class Y>
requires (std::is_base_of_v<XLREF, X>&& std::is_base_of_v<XLREF, Y>)
	|| (std::is_base_of_v<XLREF12, X> && std::is_base_of_v<XLREF12, X>)
inline bool operator!=(const X & x, const Y & y)
{
	return compare_three_way(x, y) != 0;
}
template<class X, class Y>
requires (std::is_base_of_v<XLREF, X>&& std::is_base_of_v<XLREF, Y>)
|| (std::is_base_of_v<XLREF12, X> && std::is_base_of_v<XLREF12, X>)
inline bool operator<(const X & x, const Y & y)
{
	return compare_three_way(x, y) < 0;
}

// need class to auto generate three way compare
struct REFX : public XLREF {
	using XLREF::rwFirst;
	using XLREF::rwLast;
	REFX()
	{ }
	REFX(const XLREF& r)
		: XLREF(r)
	{ }
	auto operator<=>(const REFX&) const = default;
};

namespace xll {

	template<class X>
		requires (std::is_same_v<XLOPER, X> || std::is_same_v<XLOPER12, X>)
	class XREF : public traits<X>::xref {
		using xrw = typename traits<X>::xrw;
		using xcol = typename traits<X>::xcol;
		using xref = typename traits<X>::xref;
		using xref::xref;
	public:
		XREF()
		{ }
		XREF(xrw row, xcol col, xrw height = 1, xcol width = 1)
			: xref{ row, row + height, col, col + width }
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
