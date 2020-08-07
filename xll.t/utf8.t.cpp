// utf8.t.cpp test utf8.h
#include <cassert>
#include "../xll/utf8.h"

using utf8::mbstowcs;
using utf8::wcstombs;

int test_utf8 = []() {
	{
		char s[] = "abc";
		auto ws = mbstowcs(s);
		assert(0 == wcsncmp(ws, L"abc", 3));
		free(ws);
	}
	{
		char s[] = "abc";
		auto ws = mbstowcs(s);
		assert(0 == wcsncmp(ws, L"abc", 3));
		free(ws);
	}
	{
		char s[] = "哈"; // ha
		assert(3 == strlen(s));
		auto ws = mbstowcs(s);
		assert(0 == wcsncmp(ws, L"哈", 1));
		free(ws);
	}
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