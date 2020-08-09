// utf8.h - convert utf-8 to wide character string
#pragma once
#include <cstdlib>
#include <cstring>
#include <cwchar>

#ifndef ensure
#define ENSURE_DEFINED
#define ensure(e) if(!(e)) { return nullptr; }
#endif

namespace utf8 {
	// Multi-byte character string to wide character string allocated by malloc.
	// If n < 0 return count in first character.
	inline wchar_t* mbstowcs(const char* s, int n = 0)
	{
		wchar_t* ws = nullptr;
		bool counted = false;

		if (n < 0) {
			counted = true;
			n = -n;
		}

		if (n == 0) {
			n = (int)strlen(s);
		}

		size_t wn = ::mbstowcs(0, s, n);
		ensure(wn != (size_t)-1);

		ws = (wchar_t*)malloc((wn + counted) * sizeof(wchar_t));
		if (nullptr != ws) {
			size_t wn_ = ::mbstowcs(ws + counted, s, n);
			// ??? NormalizeString()
			if (counted) {
				ensure(wn <= WCHAR_MAX);
				ws[0] = static_cast<wchar_t>(wn);
			}
		}

		return ws;
	}
	// Wide character string to Multi-byte character string allocated by malloc
	inline char* wcstombs(const wchar_t* ws, int wn = 0)
	{
		char* s = nullptr;
		bool counted = false;

		if (wn < 0) {
			counted = true;
			wn = -wn;
		}

		if (wn == 0) {
			wn = (int)wcslen(ws);
		}

		size_t n = ::wcstombs(0, ws, wn);
		//ensure(n !== (size_t)-1));

		s = (char*)malloc(n + counted);
		if (nullptr != s) {
			size_t n_ = ::wcstombs(s + counted, ws, wn);
			if (counted) {
				ensure(n <= UCHAR_MAX);
				s[0] = static_cast<char>(n);
			}
		}

		return s;
	}
}

#ifdef ENSURE_DEFINED
#undef ensure
#endif