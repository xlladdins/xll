// xloper.h - XLOPER related code
#pragma once
#include <concepts>
#include <cstring>
#include <type_traits>
#include <Windows.h>
#include "XLCALL.H"

namespace xll {

#pragma warning(push)
#pragma warning(disable: 4996)

	template<typename X>
	concept XLOPERX = requires {
		std::is_same_v<X, XLOPER> || std::is_same_v<X, XLOPER12>
	};
		
	// XLOPER/XLOPER12 traits
	template<class X>
	struct traits {
	};
	template<>
	struct traits<XLOPER> {
		typedef CHAR xchar;
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

	// predefined XLOPERs
	inline const XLOPER Missing
		= { .xltype = xltypeMissing };
	inline const XLOPER12 Missing12
		= { .xltype = xltypeMissing };
	inline const XLOPER Nil
		= { .xltype = xltypeNil };
	inline const XLOPER12 Nil12
		= { .xltype = xltypeNil };

	inline const XLOPER ErrNull
		= { .val = { .err = xlerrNull }, .xltype = xltypeErr };
	inline const XLOPER12 ErrNull12
		= { .val = { .err = xlerrNull }, .xltype = xltypeErr };
	inline const XLOPER ErrDiv0
		= { .val = { .err = xlerrDiv0 }, .xltype = xltypeErr };
	inline const XLOPER12 ErrDiv012
		= { .val = { .err = xlerrDiv0 }, .xltype = xltypeErr };
	inline const XLOPER ErrValue
		= { .val = { .err = xlerrValue }, .xltype = xltypeErr };
	inline const XLOPER12 ErrValue12
		= { .val = { .err = xlerrValue }, .xltype = xltypeErr };
	inline const XLOPER ErrRef
		= { .val = { .err = xlerrRef }, .xltype = xltypeErr };
	inline const XLOPER12 ErrRef12
		= { .val = { .err = xlerrRef }, .xltype = xltypeErr };
	inline const XLOPER ErrName
		= { .val = { .err = xlerrName }, .xltype = xltypeErr };
	inline const XLOPER12 ErrName12
		= { .val = { .err = xlerrName }, .xltype = xltypeErr };
	inline const XLOPER ErrNum
		= { .val = { .err = xlerrNum }, .xltype = xltypeErr };
	inline const XLOPER12 ErrNum12
		= { .val = { .err = xlerrNum }, .xltype = xltypeErr };
	inline const XLOPER ErrNA
		= { .val = { .err = xlerrNA }, .xltype = xltypeErr };
	inline const XLOPER12 ErrNA12
		= { .val = { .err = xlerrNA }, .xltype = xltypeErr };

}
