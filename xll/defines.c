// defines.c
#include "defines.h"

// array designators not allowed in C++
#undef X
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

#define X(a,b,c,d) const LPCSTR XLL_##a##4 = b;
XLL_ARG_TYPE(X)
#undef X
#define X(a,b,c,d) const LPCSTR XLL_##a##12 = c;
XLL_ARG_TYPE(X)
#undef X
#if XLL_VERSION == 4
#define X(a,b,c,d) const LPCSTR XLL_##a = b;
XLL_ARG_TYPE(X)
#else
#define X(a,b,c,d) const LPCSTR XLL_##a = c;
#endif // XLL_VERSION
XLL_ARG_TYPE(X)
#undef X
