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
	
#define XLL_TYPE_TRAITS(a,b,c,d,e) typedef decltype(XLOPER::val. ## b) x##c;
	template<>
	struct traits<XLOPER> {
		typedef typename XLOPER xtype;
		typedef typename XLOPER12 typex; // the other type
		typedef typename XLMREF xmref;
		XLL_TYPE_SCALAR(XLL_TYPE_TRAITS)
		XLL_TYPE_ALLOC(XLL_TYPE_TRAITS)
		typedef typename const CHAR* xcstr;
		typedef typename CHAR xchar;
		typedef typename WORD xword;
		typedef typename WORD xrw;
		typedef typename BYTE xcol;
		typedef typename _FP xfp;
		typedef typename std::basic_string<xchar> xstring;
		static const int argmax = 255;
		static const int charmax = 0xFF;
		static int WINAPI MessageBoX(HWND hWnd, LPCSTR lpText, LPCSTR lpCaption, UINT uType)
		{
			return MessageBoxA(hWnd, lpText, lpCaption, uType);
		}
#ifndef _CONSOLE
		static int Excelv(int xlfn, LPXLOPER operRes, int count, LPXLOPER opers[])
		{
			return ::Excel4v(xlfn, operRes, count, opers);
		}
#else
		static int Excelv(int, LPXLOPER, int, LPXLOPER[])
		{
			return 0;
		}
#endif
		static int len(const xchar* s)
		{
			return s ? static_cast<int>(strlen(s)) : 0;
		}
		static xchar* cpy(xchar* dest, const xchar* src, unsigned n)
		{
			ensure(0 == strncpy_s(dest, n, src, n));
	
			return dest;
		}
		static int cmp(const xchar* dest, const xchar* src, unsigned n)
		{
			return strncmp(dest, src, n);
		}
		static int alnum(int c)
		{
			return ::isalnum(c);
		}
		// return counted string that must be free'd
		static char* cvt(const wchar_t* ws, int wn = -1)
		{
			return utf8::wcstombs(ws, wn);
		}
	};
#undef XLL_TYPE_TRAITS

#define XLL_TYPE_TRAITS(a,b,c,d,e) typedef decltype(XLOPER12::val. ## b) x##c;
	template<>
	struct traits<XLOPER12> {
		typedef typename XLOPER12 xtype;
		typedef typename XLOPER typex; // not XLOPER12
		//typedef typename XLREF12 xref;
		typedef typename XLMREF12 xmref;
		XLL_TYPE_SCALAR(XLL_TYPE_TRAITS)
		XLL_TYPE_ALLOC(XLL_TYPE_TRAITS)
		typedef typename XCHAR xchar;
		typedef typename const XCHAR* xcstr;
		//typedef typename INT32 xint;
		typedef typename WORD xword;
		typedef typename RW xrw;
		typedef typename COL xcol;
		//typedef typename INT32 xbool;
		//typedef typename int xerr;
		typedef typename _FP12 xfp;
		typedef typename std::basic_string<xchar> xstring;
		static const int argmax = 255;
		static const int charmax = 0x7FFF;
		static int WINAPI MessageBoX(HWND hWnd, LPCWSTR lpText, LPCWSTR lpCaption, UINT uType)
		{
			return MessageBoxW(hWnd, lpText, lpCaption, uType);
		}
#ifndef _CONSOLE
		static int Excelv(int xlfn, LPXLOPER12 operRes, int count, LPXLOPER12 opers[])
		{
			return ::Excel12v(xlfn, operRes, count, opers);
		}
#else
		static int Excelv(int, LPXLOPER12, int, LPXLOPER12[])
		{
			return 0;
		}
#endif
		static int len(const xchar* s)
		{
			return s ? static_cast<int>(wcslen(s)) : 0;
		}
		static xchar* cpy(xchar* dest, const xchar* src, unsigned n)
		{
			ensure(0 == wcsncpy_s(dest, n, src, n));

			return dest;
		}
		static int cmp(const xchar* dest, const xchar* src, unsigned n)
		{
			return wcsncmp(dest, src, n);
		}
		static int alnum(wint_t c)
		{
			return ::iswalnum(c);
		}
		// return counted string that must be free'd
		static wchar_t* cvt(const char* s, int n = -1)
		{
			return utf8::mbstowcs(s, n);
		}
	};
#undef XLL_TYPE_TRAITS

} // namespace xll

