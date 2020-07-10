// xloper.h - XLOPER related code
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.

#pragma once
#include "defines.h"

namespace xll {


#define X(a, b) \
	inline const XLOPER a = { .xltype = xltype ## a };         \
	inline const XLOPER12 a ## 12 = { .xltype = xltype ## a }; \

	XLL_NULL_TYPE(X)
#undef X

#define X(a, b) \
	inline XLOPER Err##a = { .val = { .err = xlerr##a }, .xltype = xltypeErr };       \
	inline XLOPER12 Err##a##12 = { .val = { .err = xlerr##a }, .xltype = xltypeErr }; \

	XLL_ERROR_TYPE(X)
#undef X

}

