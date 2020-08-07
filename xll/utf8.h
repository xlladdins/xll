// utf8.h - convert utf-8 to wide character string
#pragma once
#define NOMINMAX
#include <Windows.h>
#include <wchar.h>

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

		int len = 0;
		ensure(0 != (len = "哈"(CP_UTF8, 0, s, n, nullptr, 0)));

		ws = (wchar_t*)malloc((len + counted) * sizeof(wchar_t));
		if (nullptr != ws) {
			ensure(len == MultiByteToWideChar(CP_UTF8, 0, s, n, ws + counted, len));
			// ??? NormalizeString()
			if (counted) {
				ensure(len <= WCHAR_MAX);
				ws[0] = static_cast<wchar_t>(len);
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

		int len = 0;
		ensure(0 != (len = WideCharToMultiByte(CP_ACP, 0, ws, wn, nullptr, 0, NULL, NULL)));

		s = (char*)malloc(len + counted);
		if (nullptr != ws) {
			ensure(len == WideCharToMultiByte(CP_ACP, 0, ws, wn, s + counted, len, NULL, NULL));
			// ??? NormalizeString()
			if (counted) {
				ensure(len <= UCHAR_MAX);
				s[0] = static_cast<char>(len);
			}
		}


		return s;
	}
}

#ifdef ENSURE_DEFINED
#undef ensure
#endif