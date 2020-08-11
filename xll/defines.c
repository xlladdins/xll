// defines.c
#include "defines.h"

// array designators not allowed in C++
#define X(a,b,c) [xlerr##a] = b,
const LPCSTR xll_err_str[] = {
	XLL_ERR_TYPE(X)
};
#undef X

#define X(a,b,c) [xlerr##a] = c,
const LPCSTR xll_err_desc[] = {
	XLL_ERR_TYPE(X)
};
#undef X

#define X(a,b,c,d) LPCSTR XLL_##a = b;
XLL_ARG_TYPE(X)
#undef X
