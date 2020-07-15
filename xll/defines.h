// defines.h
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <set>
#include <string>
#include "traits.h"

// 64-bit uses different symbol name decoration
#ifdef _M_X64 
#define XLL_DECORATE(s,n) s
#define XLL_X64(x) x
#define XLL_X32(x)
#else
#define XLL_DECORATE(s,n) "_" s "@" #n
#define XLL_X64(x)	
#define XLL_X32(x) x
#endif

// Used to export undecorated function name from a dll.
#define XLLEXPORT comment(linker, "/export:" __FUNCDNAME__ "=" __FUNCTION__)

#define XLL_NULL_TYPE(X)                    \
	X(Missing, "missing function argument") \
	X(Nil,     "empty cell")

#define XLL_ERROR_TYPE(X)                                          \
	X(Null,    "intersection of two ranges that do not intersect") \
	X(Div0,    "formula divides by zero")                          \
	X(Value,   "variable in formula has wrong type")               \
	X(Ref,     "formula contains an invalid cell reference")       \
	X(Name,    "unrecognised formula name or text")                \
	X(Num,     "invalid number")                                   \
	X(NA,      "value not available to a formula.")                \

// Argument types for Excel Functions
#define XLL_ARG_TYPE(X)                                                 \
X(BOOL,     "A",  "short int used as logical")                          \
X(DOUBLE,   "B",  "double")                                             \
X(CSTRING,  "C%", "XCHAR* to C style NULL terminated unicode string")   \
X(PSTRING,  "D%", "XCHAR* to Pascal style byte counted unicode string") \
X(DOUBLE_,  "E",  "pointer to double")                                  \
X(CSTRING_, "F%", "reference to a null terminated unicode string")      \
X(PSTRING_, "G%", "reference to a byte counted unicode string")         \
X(USHORT,   "H",  "unsigned 2 byte int")                                \
X(WORD,     "H",  "unsigned 2 byte int")                                \
X(SHORT,    "I",  "signed 2 byte int")                                  \
X(LONG,     "J",  "signed 4 byte int")                                  \
X(FP,       "K%", "pointer to struct FP")                               \
X(BOOL_,    "L",  "reference to a boolean")                             \
X(SHORT_,   "M",  "reference to signed 2 byte int")                     \
X(LONG_,    "N",  "reference to signed 4 byte int")                     \
X(LPOPER,   "Q",  "pointer to OPER struct (never a reference type)")    \
X(LPXLOPER, "U",  "pointer to XLOPER struct")                           \
X(VOLATILE, "!",  "called every time sheet is recalced")                \
X(UNCALCED, "#",  "dereferencing uncalced cells returns old value")     \
X(THREAD_SAFE,  "$",  "declares function to be thread safe")            \
X(CLUSTER_SAFE, "&", "declares function to be cluster safe")            \
X(ASYNCHRONOUS, "X", "declares function to be asynchronous")            \
X(VOID,     ">",  "return type to use for asynchronous functions")      \

#define T_(s) _T(s)
#define X(a,b,c)                                                 \
inline const xll::traits<XLOPER>::xchar*   XLL_##a     = b;      \
inline const xll::traits<XLOPER12>::xchar* XLL_##a##12 = L##b;   \
inline const xll::traits<XLOPERX>::xchar*  XLL_##a##X  = T_(b); \

XLL_ARG_TYPE(X)
#undef X

#define X(a,b,c)T_(b), 
inline std::set<std::basic_string<xll::traits<XLOPERX>::xchar>> xll_arg_types {
	XLL_ARG_TYPE(X)
};

#undef X

#undef T_
