// traits.h - parameterized by XLOPER type
#pragma once
#include <Windows.h>
#include "XLCALL.H"

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
		static size_t len(const xchar* s)
		{
			return strlen(s);
		}
		static xchar* cpy(xchar* dest, const xchar* src, size_t n)
		{
			return strncpy(dest, src, n);
		}
	};
	template<>
	struct traits<XLOPER12> {
		typedef XCHAR xchar;
		typedef const XCHAR* xcstr;
		typedef int xint;
		static size_t len(const xchar* s)
		{
			return wcslen(s);
		}
		static xchar* cpy(xchar* dest, const xchar* src, size_t n)
		{
			return wcsncpy(dest, src, n);
		}
	};

#pragma warning(pop)

}