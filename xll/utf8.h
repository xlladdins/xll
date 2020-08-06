// utf8.h - convert utf-8 to wide character string
#pragma once
#include <windows.h>
#include <stdexcept>
#include <string>

inline std::wstring Utf8ToWcs(const char* s, size_t n = 0)
{
	std::wstring wcs;

	if (!s || !*s) {
		return wcs;
	}

	if (n == 0) {
		n = strlen(s);
	}

	int len = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, s, n, nullptr, 0);
	if (len == 0) {
		// use FormatMessage!!!
		throw std::runtime_error(__FUNCTION__ ": MultiByteToWideChar get size failed");
	}

	wcs.resize(len);
	len = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, s, n, wcs.data(), len);
	if (len == 0) {
		// use FormatMessage!!!
		throw std::runtime_error(__FUNCTION__ ": MultiByteToWideChar failed");
	}

	return wcs;
}

