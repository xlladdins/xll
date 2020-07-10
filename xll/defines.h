// defines.h
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include "traits.h"

// Build for Excel 2007 and later by default
#ifndef XLL_VERSION
#define XLL_VERSION 12
#else
#define XLL_VERSION 4
#endif

#if XLL_VERSION == 4
#define XLOPERX XLOPER
#else
static_assert(XLL_VERSION == 12);
#define XLOPERX XLOPER12
#endif

// ASCII or wide character strings
#if XLL_VERSION == 4
#define X_(t) t
#define TX_(t) #t
#else
static_assert(XLL_VERSION == 12);
#define X_(t) L##t
#define TX_(t) L#t
#endif

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

#define X(a,b,c) \
inline const char* XLL_##a##4 = b; \
inline const wchar_t* XLL_##a##12 = L#b; \
inline xll::traits<XLOPERX>::xcstr XLL_##a##X = X_(b);

XLL_ARG_TYPE(X)
#undef X

// Excel C data types for Excel12 xlfRegister/AddIn.
#define XLL_BOOL     L"A"  // short int used as logical
#define XLL_DOUBLE   L"B"  // double
#define XLL_CSTRING  L"C%" // XCHAR* to C style NULL terminated unicode string
#define XLL_PSTRING  L"D%" // XCHAR* to Pascal style byte counted unicode string
#define XLL_DOUBLE_  L"E"  // pointer to double
#define XLL_CSTRING_ L"F%" // reference to a null terminated unicode string
#define XLL_PSTRING_ L"G%" // reference to a byte counted unicode string
#define XLL_USHORT   L"H"  // unsigned 2 byte int
#define XLL_WORD     L"H"  // unsigned 2 byte int
#define XLL_SHORT    L"I"  // signed 2 byte int
#define XLL_LONG     L"J"  // signed 4 byte int
#define XLL_FP       L"K%" // pointer to struct FP
#define XLL_BOOL_    L"L"  // reference to a boolean
#define XLL_SHORT_   L"M"  // reference to signed 2 byte int
#define XLL_LONG_    L"N"  // reference to signed 4 byte int
#define XLL_LPOPER   L"Q"  // pointer to OPER struct (never a reference type).
#define XLL_LPXLOPER L"U"  // pointer to XLOPER struct
// Modify add-in function behavior.
#define XLL_VOLATILE L"!"  // called every time sheet is recalced
#define XLL_UNCALCED L"#"  // dereferencing uncalced cells returns old value
#define XLL_THREAD_SAFE L"$"    // declares function to be thread safe
#define XLL_CLUSTER_SAFE L"&"	// declares function to be cluster safe
#define XLL_ASYNCHRONOUS L"X"	// declares function to be asynchronous
#define XLL_VOID     L">"	    // return type to use for asynchronous functions

#define XLL_HANDLE XLL_DOUBLE // pointer to C++ object
#define XLL_DATE   XLL_DOUBLE // Excel Julian date. Days from 12AM 1 Jan 1900
