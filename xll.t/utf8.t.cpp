// utf8.t.cpp test utf8.h
#include <cassert>
#include "../xll/utf8.h"

int test_utf8 = []() {
	{
		wchar_str s;
		assert(0 == s.length());
		wchar_str s2(std::move(s));
		assert(0 == s2.length());
		assert(nullptr != s2.data());
		assert(0 == s.length());
		assert(nullptr == s.data());
	}
	{
		wchar_str s(L"abc");
		assert(3 == s.length());
		wchar_str s2(std::move(s));
		assert(3 == s2.length());
		assert(0 == s.length());
		assert(nullptr == s.data());
	}
	{
		char s[] = "abc";
		auto ws = Utf8ToWcs(s);
		assert(3 == ws.length());
		assert(0 == wcsncmp(ws.data(), L"abc", 3));
	}
	{
		char s[] = "abc";
		auto ws = Utf8ToWcs(s);
		assert(3 == ws.length());
		assert(0 == wcsncmp(ws.data(), L"abc", 3));
	}
	{
		char s[] = "哈"; // ha
		assert(3 == strlen(s));
		auto ws = Utf8ToWcs(s);
		assert(1 == ws.length());
		assert(0 == wcsncmp(ws.data(), L"哈", 1));
	}

	return 0;
}();