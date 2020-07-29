// traits.h - Top level include file to parameterized by XLOPER type
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once

#ifndef XLOPERX
	#define XLOPERX XLOPER12
	#define LPXLOPERX LPXLOPER12
	typedef struct _FP12 _FPX;
	#undef _MBCS

	#ifndef _UNICODE
	#define _UNICODE
	#endif

	#ifndef UNICODE
	#define UNICODE
	#endif
#else
	#define XLOPERX XLOPER
	#define LPXLOPERX LPXLOPER
	typedef struct _FP _FPX;

	#ifndef _MBCS
	#define _MBCS
	#endif

	#undef _UNICODE
	#undef UNICODE
#endif

#define NOMINMAX
#include <Windows.h>
#include <tchar.h>
#include "XLCALL.H"
#include "ensure.h"

#pragma warning(push)
#pragma warning(disable: 4996)

namespace xll {

	// XLOPER/XLOPER12 traits
	template<class X>
		requires (std::is_same_v<XLOPER, X> || std::is_same_v<XLOPER12, X>)
	struct traits {
	};
	template<>
	struct traits<XLOPER> {
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
			return strncpy(dest, src, n);
		}
		static int cmp(const xchar* dest, const xchar* src, size_t n)
		{
			return strncmp(dest, src, n);
		}
	};
	template<>
	struct traits<XLOPER12> {
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
			return wcsncpy(dest, src, n);
		}
		static int cmp(const xchar* dest, const xchar* src, size_t n)
		{
			return wcsncmp(dest, src, n);
		}
	};

}

#pragma warning(pop)
