// utf8.t.cpp test utf8.h
#include <cassert>
#include <memory>
#include "../xll/utf8.h"

using std::unique_ptr;
using utf8::mbstowcs;
using utf8::wcstombs;

int test_utf8 = []() {
	{
		char s[] = "abc";
		unique_ptr<wchar_t> ws(mbstowcs(s));
		assert(0 == wcsncmp(ws.get(), L"abc", 3));
	}
	{
		char s[] = "abc";
		auto ws = mbstowcs(s);
		assert(0 == wcsncmp(ws, L"abc", 3));
		free(ws);
	}
	/*
	{
		char s[] = "哈"; // ha
		wchar_t ha[] = L"哈";
		assert(3 == strlen(s));
		wchar_t* ws = mbstowcs(s); // unique_ptr!!!
		assert(0 == wcsncmp(ws, ha, 1));
		free(ws);
	}
	*/
	{
		wchar_t ws[] = L"abc";
		auto s = wcstombs(ws);
		assert(0 == strncmp(s, "abc", 3));
		free(s);
	}
	{
		wchar_t ws[] = L"哈";
		auto s = wcstombs(ws);
		size_t n;
		n = strlen(s);
		assert(0 == strncmp(s, "哈", 3));
		free(s);
	}

	return 0;
}();