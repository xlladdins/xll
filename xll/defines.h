// defines.h - top level include
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once

#ifndef XLOPERX
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
#include "XLCALL.H"

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

#define xltypeXxx 0 // phony type that is never used
#define xltypePtr -1 // void* pointer

// Argument types for Excel Functions
#define XLL_ARG_TYPE(X)                                                              \
X(BOOL,     "A",  xltypeBool,  "short int used as logical")                          \
X(DOUBLE,   "B",  xltypeNum,   "double")                                             \
X(CSTRING,  "C%", xltypeStr,   "XCHAR* to C style NULL terminated unicode string")   \
X(PSTRING,  "D%", xltypeStr,   "XCHAR* to Pascal style byte counted unicode string") \
X(DOUBLE_,  "E",  xltypePtr,   "pointer to double")                                  \
X(CSTRING_, "F%", xltypeXxx,   "reference to a null terminated unicode string")      \
X(PSTRING_, "G%", xltypeXxx,   "reference to a byte counted unicode string")         \
X(USHORT,   "H",  xltypeNum,   "unsigned 2 byte int")                                \
X(WORD,     "H",  xltypeNum,   "unsigned 2 byte int")                                \
X(SHORT,    "I",  xltypeInt,   "signed 2 byte int")                                  \
X(LONG,     "J",  xltypeNum,   "signed 4 byte int")                                  \
X(FP,       "K%", xltypeXxx,   "pointer to struct FP")                               \
X(BOOL_,    "L",  xltypeXxx,   "reference to a boolean")                             \
X(SHORT_,   "M",  xltypeXxx,   "reference to signed 2 byte int")                     \
X(LONG_,    "N",  xltypeXxx,   "reference to signed 4 byte int")                     \
X(LPOPER,   "Q",  xltypeMulti, "pointer to OPER struct (never a reference type)")    \
X(LPXLOPER, "U",  xltypeMulti, "pointer to XLOPER struct")                           \
X(VOLATILE, "!",  xltypeXxx,   "called every time sheet is recalced")                \
X(UNCALCED, "#",  xltypeXxx,   "dereferencing uncalced cells returns old value")     \
X(VOID,     ">",  xltypeXxx,   "return type to use for asynchronous functions")      \
X(THREAD_SAFE,  "$", xltypeXxx, "declares function to be thread safe")               \
X(CLUSTER_SAFE, "&", xltypeXxx, "declares function to be cluster safe")              \
X(ASYNCHRONOUS, "X", xltypeXxx, "declares function to be asynchronous")              \

#ifdef __cplusplus 
extern "C" {
#endif
// Defined in defines.c
#define X(a,b,c,d) extern LPCSTR  XLL_##a;
XLL_ARG_TYPE(X)
#undef X
#ifdef __cplusplus
}
#endif

