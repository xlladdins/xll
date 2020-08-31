// utf8.h - convert utf-8 to counted, allocated string and vice versa
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
	inline /*_Post_ _Notnull_*/ wchar_t* mbstowcs(const char* s, size_t n = 0)
	{
		wchar_t* ws = nullptr;

		if (nullptr == s) {
			n = 0;
		}
		else if (n == 0) {
			n = strlen(s);
		}

		int wn = 0;
		if (n != 0) {
			ensure(0 != (wn = MultiByteToWideChar(CP_ACP, 0, s, (int)n, nullptr, 0)));
		}

		ws = (wchar_t*)malloc((wn + 1) * sizeof(wchar_t));
		if (ws) {
			ensure(wn == MultiByteToWideChar(CP_ACP, 0, s, (int)n, ws + 1, wn));
			ensure(wn <= WCHAR_MAX);
			ws[0] = static_cast<wchar_t>(wn);
		}

		return ws;
	}
	inline std::wstring mbstowstring(const char* s, size_t n = 0)
	{
		std::wstring ws;

		wchar_t* pws = mbstowcs(s, n);
		if (pws) {
			ws.assign(pws + 1, pws[0]);
			free(pws);
		}

		return ws;
	}

	// Wide character string to counted multi-byte character string allocated by malloc
	inline /*_Post_ _Notnull_*/ char* wcstombs(const wchar_t* ws, size_t wn = 0)
	{
		char* s = nullptr;

		if (nullptr == ws) {
			wn = 0;
		}
		else if (wn == 0) {
			wn = wcslen(ws);
		}

		int n = 0;
		if (wn != 0) {
			ensure(0 != (n = WideCharToMultiByte(CP_UTF8, 0, ws, (int)wn, NULL, 0, NULL, NULL)));
		}

		s = (char*)malloc(n + 1);
		if (nullptr != s) {
			ensure(n == WideCharToMultiByte(CP_UTF8, 0, ws, (int)wn, s + 1, n, NULL, NULL));
			// ???NormalizeString
			ensure(n <= UCHAR_MAX);
			s[0] = static_cast<char>(n);
		}

		return s;
	}
	inline std::string wcstostring(const wchar_t* ws, size_t wn = 0)
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
