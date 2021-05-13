// debug.cpp - dump memory leaks
// Copyright (c) 2006-2016 KALX, LLC. All rights reserved. No warranty is made.
#ifdef _DEBUG
#define _CRTDBG_MAP_ALLOC
//#define _CRTDBG_MAP_ALLOC_NEW 
#include <stdlib.h>
#include <crtdbg.h>

struct CrtDbg {
	CrtDbg(int flags = _CRTDBG_ALLOC_MEM_DF)
	{
		_CrtSetDbgFlag (_CrtSetDbgFlag( _CRTDBG_REPORT_FLAG )|flags);
		//_crtBreakAlloc = 1178;
	}
	// When information about a memory block is reported by one of the debug 
	// dump functions, this number is enclosed in braces, such as {36}.
	// Alternatively, set _crtBreakAlloc in the debugger.
	static long SetBreakAlloc(long breakAlloc)
	{
		return breakAlloc ? _CrtSetBreakAlloc(breakAlloc) : 0;
	}
	// Verifies that a specified memory block is in the local heap and that it 
	// has a valid debug heap block type identifier. Any of the last three
	// arguments may be NULL.
	static int IsMemoryBlock(const void* userData, unsigned int size,
		long* requestNumber, char** filename, int* linenumber)
	{
		return _CrtIsMemoryBlock(userData, size, requestNumber, filename, linenumber);
	}
	static long RequestNumber(const void* userData, unsigned int size)
	{
		long req;

		return IsMemoryBlock(userData, size, &req, 0, 0) ? req : 0;
	}
	~CrtDbg()
	{
		if (!(_CrtSetDbgFlag( _CRTDBG_REPORT_FLAG )&_CRTDBG_LEAK_CHECK_DF))
			_CrtDumpMemoryLeaks();
	}
};

// need to construct this before user segment
#pragma warning(disable: 4073)
#pragma init_seg(lib)
CrtDbg crtDbg;

#endif // _DEBUG
