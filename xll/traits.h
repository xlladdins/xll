// traits.h - parameterize by XLOPER/XLOPER12 types
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <type_traits>
#include "defines.h"
#include "ensure.h"
#include "utf8.h"

namespace xll {

	// XLOPER/XLOPER12 traits
	template<class X> requires (std::is_same_v<XLOPER, X> || std::is_same_v<XLOPER12, X>)
	struct traits { };
	
	template<>
	struct traits<XLOPER> {
		typedef typename XLOPER xtype;
		typedef typename XLOPER12 typex;
		typedef typename XLREF xref;
		typedef typename CHAR xchar;
		typedef typename const CHAR* xcstr;
		typedef typename short int xint;
		typedef typename WORD xrw;
		typedef typename WORD xcol; // BYTE???
		typedef typename WORD xbool;
		typedef typename _FP xfp;
		static int Excelv(int xlfn, LPXLOPER operRes, int count, LPXLOPER opers[])
		{
			return ::Excel4v(xlfn, operRes, count, opers);
		}
		static int len(const xchar* s)
		{
			return s ? static_cast<int>(strlen(s)) : 0;
		}
		static xchar* cpy(xchar* dest, const xchar* src, size_t n)
		{
			ensure (0 == strncpy_s(dest, n, src, n));

			return dest;
		}
		static int cmp(const xchar* dest, const xchar* src, size_t n)
		{
			return strncmp(dest, src, n);
		}
		// return counted string that must be free'd
		static char* cvt(const wchar_t* ws)
		{
			return utf8::wcstombs(ws);
		}
	};
	
	template<>
	struct traits<XLOPER12> {
		typedef typename XLOPER12 xtype;
		typedef typename XLOPER typex; // not XLOPER12
		typedef typename XLREF12 xref;
		typedef typename XCHAR xchar;
		typedef typename const XCHAR* xcstr;
		typedef typename INT32 xint;
		typedef typename RW xrw;
		typedef typename COL xcol;
		typedef typename _FP12 xfp;
		typedef typename INT32 xbool;
		static int Excelv(int xlfn, LPXLOPER12 operRes, int count, LPXLOPER12 opers[])
		{
			return ::Excel12v(xlfn, operRes, count, opers);
		}
		static int len(const xchar* s)
		{
			return s ? static_cast<int>(wcslen(s)) : 0;
		}
		static xchar* cpy(xchar* dest, const xchar* src, size_t n)
		{
			ensure (0 == wcsncpy_s(dest, n, src, n));

			return dest;
		}
		static int cmp(const xchar* dest, const xchar* src, size_t n)
		{
			return wcsncmp(dest, src, n);
		}
		// return counted string that must be free'd
		static wchar_t* cvt(const char* s)
		{
			return utf8::mbstowcs(s);
		}
	};

}

