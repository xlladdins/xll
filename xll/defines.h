// defines.h
#pragma once

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
#define XLL_DATE   XLL_DOUBLE // Excel Julian date
