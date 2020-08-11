// utf8.h - convert utf-8 to counted, allocated string
#pragma once
#include <cstdlib>
#include <cstring>
#include <cwchar>
#include <Windows.h>

#ifndef ensure
#define ENSURE_DEFINED
#define ensure(e) if(!(e)) { return nullptr; }
#endif

namespace utf8 {
	// Multi-byte character string to wide character string allocated by malloc.
	// If n < 0 return count in first character.
	inline wchar_t* mbstowcs(const char* s, size_t n = 0)
	{
		wchar_t* ws = nullptr;
		
		if (n == 0) {
			n = strlen(s);
		}

		int wn = 0;
		if (n != 0) {
			wn = MultiByteToWideChar(CP_ACP, 0, s, (int)n, nullptr, 0);
			ensure(wn != 0);
		}

		ws = (wchar_t*)malloc((wn + 1) * sizeof(wchar_t));
		if (nullptr != ws) {
			ensure(wn == MultiByteToWideChar(CP_ACP, 0, s, (int)n, ws + 1, wn));
			// ??? NormalizeString()
			ensure(wn <= WCHAR_MAX);
			ws[0] = static_cast<wchar_t>(wn);
		}

		return ws;
	}
	// Wide character string to Multi-byte character string allocated by malloc
	inline char* wcstombs(const wchar_t* ws, size_t wn = 0)
	{
		char* s = nullptr;

		if (wn == 0) {
			wn = wcslen(ws);
		}

		int n = 0;
		if (wn != 0) {
			n = WideCharToMultiByte(CP_UTF8, 0, ws, (int)wn, NULL, 0, NULL, NULL);
			//ensure(n !== (size_t)-1));
		}

		s = (char*)malloc(n + 1);
		if (nullptr != s) {
			int n_;
			n_ = WideCharToMultiByte(CP_UTF8, 0, ws, (int)wn, s + 1, n, NULL, NULL);
			ensure(n <= UCHAR_MAX);
			s[0] = static_cast<char>(n);
		}

		return s;
	}
}

#ifdef ENSURE_DEFINED
#undef ensure
#endif