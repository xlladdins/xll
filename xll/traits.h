// traits.h - parameterize by XLOPER type
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <type_traits>
#include "defines.h"
#include "ensure.h"

namespace xll {

	/* ??? automate AddIn code
	template<class T> struct type_traits { static constexpr LPCSTR type; };
	template<> struct type_traits<double> { static constexpr LPCSTR type = XLL_DOUBLE; };
	// ...
	template<> struct type_traits<bool> { static constexpr LPCSTR type = XLL_BOOL; };
	*/

	// XLOPER/XLOPER12 traits
	template<class X>
		requires (std::is_same_v<XLOPER, X> || std::is_same_v<XLOPER12, X>)
	struct traits {
	};
	template<>
	struct traits<XLOPER> {
		typedef XLOPER xtype;
		typedef XLOPER12 typex;
		typedef CHAR xchar;
		typedef const CHAR* xcstr;
		typedef WORD xrw;
		typedef WORD xcol; // BYTE???
		typedef short int xint;
		typedef WORD xbool;
		typedef _FP xfp;
		static int Excelv(int xlfn, LPXLOPER operRes, int count, LPXLOPER opers[])
		{
			return ::Excel4v(xlfn, operRes, count, opers);
		}
		static size_t len(const xchar* s)
		{
			return strlen(s);
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
	};
	template<>
	struct traits<XLOPER12> {
		typedef XLOPER12 xtype;
		typedef XLOPER typex; // not XLOPER12
		typedef XCHAR xchar;
		typedef const XCHAR* xcstr;
		typedef INT32 xint;
		typedef RW xrw;
		typedef COL xcol;
		typedef _FP12 xfp;
		typedef INT32 xbool;
		static int Excelv(int xlfn, LPXLOPER12 operRes, int count, LPXLOPER12 opers[])
		{
			return ::Excel12v(xlfn, operRes, count, opers);
		}
		static size_t len(const xchar* s)
		{
			return wcslen(s);
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
	};

}

