// utf8.h - convert utf-8 to counted, allocated string and vice versa
// Loss-free conversion from UTF-16 to MBCS and back
// https://docs.microsoft.com/en-us/archive/msdn-magazine/2016/september/c-unicode-encoding-conversions-with-stl-strings-and-win32-apis

#pragma once
#include <cstdlib>
#include <cstring>
#include <cwchar>
#include <string>
#include <sal.h>
#include <Windows.h>
#include "ensure.h"

namespace utf8 {
		
	// Multi-byte character string to counted wide character string allocated by malloc.
	inline /*_Post_ _Notnull_*/ wchar_t* mbstowcs(const char* s, int n = -1)
	{
		ensure(0 != n);
		wchar_t* ws = nullptr;

		int wn = 0;
		ensure(0 != (wn = MultiByteToWideChar(CP_UTF8, 0, s, n, nullptr, 0)));

		ws = (wchar_t*)malloc((static_cast<size_t>(wn) + 1) * sizeof(wchar_t));
		if (ws) {
			ensure(wn == MultiByteToWideChar(CP_UTF8, 0, s, n, ws + 1, wn));
			ensure(wn <= WCHAR_MAX);
			// MBTWC includes terminating null
			ws[0] = static_cast<wchar_t>(wn - (n == -1));
		}

		return ws;
	}
	inline std::wstring mbstowstring(const char* s, int n = -1)
	{
		std::wstring ws;

		wchar_t* pws = mbstowcs(s, n);
		ensure(pws);
		if (pws) {
			ws.assign(pws + 1, pws[0]);
			free(pws);
		}

		return ws;
	}

	// Wide character string to counted multi-byte character string allocated by malloc
	inline /*_Post_ _Notnull_*/ char* wcstombs(const wchar_t* ws, int wn = -1)
	{
		ensure(0 != wn);
		char* s = nullptr;

		int n = 0;
		ensure(0 != (n = WideCharToMultiByte(CP_UTF8, 0, ws, wn, NULL, 0, NULL, NULL)));

		s = (char*)malloc(static_cast<size_t>(n) + 1);
		if (nullptr != s) {
			ensure(n == WideCharToMultiByte(CP_UTF8, 0, ws, wn, s + 1, n, NULL, NULL));
			// ???NormalizeString
			ensure(n <= UCHAR_MAX);
			// WCTMBS includes terminating null
			s[0] = static_cast<char>(n - (wn == -1));
		}

		return s;
	}
	inline std::string wcstostring(const wchar_t* ws, int wn = -1)
	{
		std::string s;

		char* ps = wcstombs(ws, wn);
		if (nullptr != ps) {
			s.assign(ps + 1, ps[0]);
			free(ps);
		}

		return s;
	}
}
