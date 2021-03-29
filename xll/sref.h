// sref.h - simple reference
#pragma once
#include <compare>
#include "traits.h"
#include "concepts.h"

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

// operator<=> only generates comparisons for C++ classes
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

	using XLREFX = traits<XLOPERX>::xref;

	template<class X> requires either_base_of_v<XLREF, XLREF12, X>
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
	inline unsigned area(const X& x) 
	{
		return height(x) * width(x);
	}

	// Translate ref by rows and columns
	template<class X> requires either_base_of_v<XLREF, XLREF12, X>
	inline X translate(X ref, int r, int c)
	{
		ref.rwFirst += r;
		ref.rwLast += r;
		ref.colFirst += c;
		ref.colLast += c;

		return ref;
	}

	template<class X> requires either_base_of_v<XLREF, XLREF12, X>
	inline X reshape(X ref, unsigned h, unsigned w)
	{
		ref.rwLast = ref.rwFirst + h - 1;
		ref.colLast = ref.colFirst + w - 1;

		return ref;
	}

	template<class X> requires either_base_of_v<XLOPER, XLOPER12, X>
	class XREF : public traits<X>::xref {
		using xrw = typename traits<X>::xrw;
		using xcol = typename traits<X>::xcol;
		using xref = typename traits<X>::xref;
	public:
		XREF()
		{ }
		XREF(xrw row, xcol col, unsigned height = 1, unsigned width = 1)
			: xref { 
				.rwFirst = row,
				.rwLast = static_cast<xrw>(row + height - 1),
				.colFirst = col, 
				.colLast = static_cast<xcol>(col + width - 1)
			}
		{ }
		XREF(const xref& r)
			: xref{ r }
		{ }
		XREF(const XREF&) = default;
		XREF& operator=(const XREF&) = default;
		~XREF()
		{ }
	};

	using REF4  = XREF<XLOPER>;
	using REF12 = XREF<XLOPER12>;
	using REF   = XREF<XLOPERX>;

} // namespace xll
