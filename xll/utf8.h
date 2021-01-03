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
	inline /*_Post_ _Notnull_*/ wchar_t* mbstowcs(const char* s, int n = -1)
	{
		wchar_t* ws = nullptr;

		if (nullptr == s) {
			n = 0;
		}
		// n = -1 means null terminated, just like MultiByteToWideChar
		if (n == 0) {
			ws = (wchar_t*)malloc(sizeof(wchar_t));
			if (ws != nullptr) {
				*ws = 0;
			}

			return ws;
		}

		// Loss-free conversion from UTF-16 to MBCS and back
		// https://tinyurl.com/yxf7lvo6?
		int wn = 0;
		if (n != 0) {
			ensure(0 != (wn = MultiByteToWideChar(CP_UTF8, 0, s, (int)n, nullptr, 0)));
		}

		ws = (wchar_t*)malloc((static_cast<size_t>(wn) + 1) * sizeof(wchar_t));
		if (ws) {
			ensure(wn == MultiByteToWideChar(CP_UTF8, 0, s ? s : "", (int)n, ws + 1, wn));
			ensure(wn <= WCHAR_MAX);
			ws[0] = static_cast<wchar_t>(wn - 1);
		}

		return ws;
	}
	inline std::wstring mbstowstring(const char* s, int n = -1)
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
	inline /*_Post_ _Notnull_*/ char* wcstombs(const wchar_t* ws, size_t wn = -1)
	{
		char* s = nullptr;

		if (nullptr == ws) {
			wn = 0;
		}
		// n = -1 means null terminated, just like MultiByteToWideChar
		if (wn == 0) {
			s = (char*)malloc(sizeof(char));
			if (s != nullptr) {
				*s = 0;
			}

			return s;
		}

		int n = 0;
		if (wn != 0) {
			ensure(0 != (n = WideCharToMultiByte(CP_UTF8, 0, ws, (int)wn, NULL, 0, NULL, NULL)));
		}

		s = (char*)malloc(static_cast<size_t>(n) + 1);
		if (nullptr != s) {
			ensure(n == WideCharToMultiByte(CP_UTF8, 0, ws ? ws : L"", (int)wn, s + 1, n, NULL, NULL));
			// ???NormalizeString
			ensure(n <= UCHAR_MAX);
			s[0] = static_cast<char>(n - 1);
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
