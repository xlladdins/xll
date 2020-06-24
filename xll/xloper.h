// xloper.h - XLOPER related code
#pragma once
#include <Windows.h>
#include "XLCALL.H"

namespace xll {

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

