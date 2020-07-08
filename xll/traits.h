// traits.h - parameterized by XLOPER type
#pragma once
#include <Windows.h>
#include "XLCALL.H"

#ifndef XLOPERX
#define XLOPERX XLOPER12
#endif

#pragma warning(push)
#pragma warning(disable: 4996)

namespace xll {

	// XLOPER/XLOPER12 traits
	template<class X>
	struct traits {
		typedef X type;
	};
	template<>
	struct traits<XLOPER> {
		typedef CHAR xchar;
		typedef const CHAR* xcstr;
		typedef short int xint;
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
		typedef int xint;
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
