// traits.h - parameterize by XLOPER/XLOPER12 types
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <type_traits>
#include "defines.h"
#include "ensure.h"
#include "utf8.h"

namespace xll {

	// XLOPER/XLOPER12 traits
	template<class X> requires (
		std::is_same_v<XLOPER, X> || std::is_same_v<XLOPER12, X> || std::is_same_v<char, X> || std::is_same_v<wchar_t, X>
		)
		struct traits { };

#define XLL_TYPE_TRAITS(a,b,c,d,e) typedef decltype(XLOPER::val.##b) x##c;
	template<>
	struct traits<XLOPER> {
		typedef XLOPER xtype;
		typedef XLOPER12 typex; // the other type
		typedef XLMREF xmref;
		XLL_TYPE_SCALAR(XLL_TYPE_TRAITS)
			XLL_TYPE_ALLOC(XLL_TYPE_TRAITS)
			typedef const CHAR* xcstr;
		typedef CHAR xchar;
		typedef WORD xword;
		typedef WORD xrw;
		typedef BYTE xcol;
		typedef _FP xfp;
		typedef std::basic_string<xchar> xstring;
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
		static xchar* cpy(xchar* dest, size_t ndest, const xchar* src, size_t nsrc)
		{
			ensure(0 == strncpy_s(dest, ndest, src, nsrc));

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
		typedef XLOPER12 xtype;
		typedef XLOPER typex; // not XLOPER12
		//typedef XLREF12 xref;
		typedef XLMREF12 xmref;
		XLL_TYPE_SCALAR(XLL_TYPE_TRAITS)
			XLL_TYPE_ALLOC(XLL_TYPE_TRAITS)
			typedef XCHAR xchar;
		typedef const XCHAR* xcstr;
		//typedef INT32 xint;
		typedef WORD xword;
		typedef RW xrw;
		typedef COL xcol;
		//typedef INT32 xbool;
		//typedef int xerr;
		typedef _FP12 xfp;
		typedef std::basic_string<xchar> xstring;
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
		static xchar* cpy(xchar* dest, size_t ndest, const xchar* src, size_t nsrc)
		{
			ensure(0 == wcsncpy_s(dest, ndest, src, nsrc));

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
