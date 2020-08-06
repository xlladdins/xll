// defines.c
#include <windows.h>
#include "XLCALL.H"

#define XLL_ERR_TYPE(X)                                                     \
	X(Null,  "#NULL!",  "intersection of two ranges that do not intersect") \
	X(Div0,  "#DIV/0!", "formula divides by zero")                          \
	X(Value, "#VALUE!", "variable in formula has wrong type")               \
	X(Ref,   "#REF!",   "formula contains an invalid cell reference")       \
	X(Name,  "#NAME?",  "unrecognised formula name or text")                \
	X(Num,   "#NUM!",   "invalid number")                                   \
	X(NA,    "#N/A",    "value not available to a formula.")                \

// array designators not allowed in C++
#define X(a,b,c) [xlerr##a] = b,
LPCSTR xll_err_str[] = {
	XLL_ERR_TYPE(X)
};
#undef X

#define X(a,b,c) [xlerr##a] = L##b,
LPCWSTR xll_err_str12[] = {
	XLL_ERR_TYPE(X)
};
#undef X
