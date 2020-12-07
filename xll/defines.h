// defines.h - top level include
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once

#ifndef XLOPERX
#define XLL_VERSION 12
#define XLOPERX XLOPER12
#define LPXLOPERX LPXLOPER12
typedef struct _FP12 _FPX;

#undef _MBCS

#ifndef _UNICODE
#define _UNICODE
#endif

#ifndef UNICODE
#define UNICODE
#endif

#else

#define XLL_VERSION 4
#define XLOPERX XLOPER
#define LPXLOPERX LPXLOPER
typedef struct _FP _FPX;

#ifndef _MBCS
#define _MBCS
#endif

#undef _UNICODE
#undef UNICODE
#endif

#define NOMINMAX
#define WIN32_LEAN_AND_MEAN
#include <Windows.h>
#include <tchar.h>
#include "XLCALL.H"

#ifdef __cplusplus
// mod with 0 <= x < y
template<typename T>
inline T xmod(T x, T y)
{
	T z = x % y;

	return z >= 0 ? z : z + y;
}
#endif

#define XLL_NULL_TYPE(X)                    \
	X(Missing, "missing function argument") \
	X(Nil,     "empty cell")                \

// xlerrXXX, Excel error string, error description
#define XLL_ERR_TYPE(X)                                                     \
	X(Null,  "#NULL!",  "intersection of two ranges that do not intersect") \
	X(Div0,  "#DIV/0!", "formula divides by zero")                          \
	X(Value, "#VALUE!", "variable in formula has wrong type")               \
	X(Ref,   "#REF!",   "formula contains an invalid cell reference")       \
	X(Name,  "#NAME?",  "unrecognised formula name or text")                \
	X(Num,   "#NUM!",   "invalid number")                                   \
	X(NA,    "#N/A",    "value not available to a formula.")                \

// Defined in defines.c
// Arrays indexed by xlerrXXX
#ifdef __cplusplus 
extern "C" {
#endif
extern const LPCSTR xll_err_str[];  // Excel string
extern const LPCSTR xll_err_desc[]; // Human readable description
#ifdef __cplusplus
}
#endif

// Argument types for Excel Functions
// XLL_XXX, Excel4, Excel12, description
#define XLL_ARG_TYPE(X)                                                      \
X(BOOL,     "A", "A",  "short int used as logical")                          \
X(DOUBLE,   "B", "B",  "double")                                             \
X(CSTRING,  "C", "C%", "XCHAR* to C style NULL terminated unicode string")   \
X(PSTRING,  "D", "D%", "XCHAR* to Pascal style byte counted unicode string") \
X(DOUBLE_,  "E", "E",  "pointer to double")                                  \
X(CSTRING_, "F", "F%", "reference to a null terminated unicode string")      \
X(PSTRING_, "G", "G%", "reference to a byte counted unicode string")         \
X(USHORT,   "H", "H",  "unsigned 2 byte int")                                \
X(WORD,     "H", "H",  "unsigned 2 byte int")                                \
X(SHORT,    "I", "I",  "signed 2 byte int")                                  \
X(LONG,     "J", "J",  "signed 4 byte int")                                  \
X(FP,       "K", "K%", "pointer to struct FP")                               \
X(BOOL_,    "L", "L",  "reference to a boolean")                             \
X(SHORT_,   "M", "M",  "reference to signed 2 byte int")                     \
X(LONG_,    "N", "N",  "reference to signed 4 byte int")                     \
X(LPOPER,   "P", "Q",  "pointer to OPER struct (never a reference type)")    \
X(LPXLOPER, "R", "U",  "pointer to XLOPER struct")                           \
X(VOLATILE, "!", "!",  "called every time sheet is recalced")                \
X(UNCALCED, "#", "#",  "dereferencing uncalced cells returns old value")     \
X(VOID,     "",  ">",  "return type to use for asynchronous functions")      \
X(THREAD_SAFE,  "", "$", "declares function to be thread safe")              \
X(CLUSTER_SAFE, "", "&", "declares function to be cluster safe")             \
X(ASYNCHRONOUS, "", "X", "declares function to be asynchronous")             \

#ifdef __cplusplus 
extern "C" {
#endif
// Defined in defines.c
#define X(a,b,c,d) extern const LPCSTR XLL_##a##4;
XLL_ARG_TYPE(X)
#undef X
#define X(a,b,c,d) extern const LPCSTR XLL_##a##12;
XLL_ARG_TYPE(X)
#undef X
#define X(a,b,c,d) extern const LPCSTR XLL_##a;
XLL_ARG_TYPE(X)
#undef X
#ifdef __cplusplus
}
#endif
