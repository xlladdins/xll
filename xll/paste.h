// paste.h - Ctrl-Shift-B and Ctrl-Shift-C for pasting functions.
// Copyright (c) 2012 KALX, LLC. All rights reserved. No warranty is made.
#pragma once
#include "xll.h"

#ifndef _DEBUG
#define ADDIN_URL "https://xladdins.com/addins/"
#endif

namespace xll {

	// iterator for counted index
	template<class C>
	class counted_index {
		const C& c;
		unsigned i, n, i0;
	public:
		counted_index(const C& c, unsigned n, unsigned i = 0, i0 = 0)
			: i(i), n(n), i0(i0), c(c)
		{ }
		bool operator==(const counted_index& c_) const
		{
			return i == c_.i and n == c_.n and i0 == c_.i0;
		}
		explicit operator bool() const
		{
			return i < n;
		}
		counted_index begin()
		{
			return counted_index(c, n, 0, i0);
		}
		counted_index end()
		{
			return counted_index(c, n, n, i0);
		}
		const auto& operator*() const
		{
			ensure(i < n);

			return c(i - i0);
		}
		counted_index& operator++()
		{
			if (i < n) {
				++i;
			}

			return *this;
		}
	};

	template<class X, class C>
	XOPER<X> paste_iterator(const counted_index<C>& cs, XOPER<X> ref = XExcel<X>(xlfActiveCell))
		// , XOPER<X> style = XOPER<X>{}
	{
		ensure(ref.is_sref());

		for (const auto& c : cs) {
			auto rw = rows(c);
			auto col = columns(c);
			auto r = XExcel<X>(xlfOffset, ref, XOPER<X>(0), XOPER<X>(0), XOPER<X>(r), XOPER<X>(c));
			XExcel<X>(xlcFormula, c, r);
			// Excel(xlcApplyStyle, style);
			// move down r rows
			ref = XExcel<X>(xlfOffset, ref, XOPER<X>(r));
		}

		return ref;
	}
}