// paste.h - Ctrl-Shift-B and Ctrl-Shift-C for pasting functions.
// Copyright (c) 2012 KALX, LLC. All rights reserved. No warranty is made.
#pragma once
#include "xll.h"

#ifndef _DEBUG
#define ADDIN_URL "https://xladdins.com/addins/"
#endif

namespace xll {


	template<class X, class C>
	inline XREF<X> paste_next(XREF<X> ref, const XOPER<X>& x, XOPER<X>& style = XOPER<X>{})
	{
		auto rw = rows(c);
		auto col = columns(c);
		auto r = XExcel<X>(xlfOffset, ref, XOPER<X>(0), XOPER<X>(0), XOPER<X>(r), XOPER<X>(c));
		XExcel<X>(xlcFormula, c, r);
		// Excel(xlcApplyStyle, style);
		// move down r rows
		return XExcel<X>(xlfOffset, ref, XOPER<X>(r));
	}
}