// utf8.h - convert utf-8 to wide character string
#pragma once
#include <windows.h>
#include <stdexcept>
#include <cstdlib>
#include <cstring>

// counted wide character string
class wchar_str {
	wchar_t* str_;
public:
	wchar_str(wchar_t* str = nullptr, size_t len = 0)
		: str_(nullptr)
	{
		if (len == 0 && str && *str) {
			len = wcslen(str);
		}
		alloc(len);
		if (len > 0 && nullptr != str_) {
			wcsncpy_s(str_ + 1, len, str, len);
		}
	}
	wchar_str(const wchar_str&) = delete;
	wchar_str& operator=(const wchar_str&) = delete;
	wchar_str(wchar_str&& str) noexcept
	{
		str_ = str.str_;
		str.str_ = nullptr;
	}
	~wchar_str()
	{
		free(str_);
	}
	wchar_t* data()
	{
		return str_ + 1;
	}
	size_t length() const
	{
		return str_[0];
	}
	wchar_t* alloc(size_t n)
	{
		wchar_t* tmp = static_cast<wchar_t*>(realloc(str_, (n + 1)*sizeof(wchar_t)));
		if (nullptr != tmp) {
			str_ = tmp;
			str_[0] = static_cast<wchar_t>(n);
		}

		return str_;
	}
};

inline wchar_str Utf8ToWcs(const char* s, int n = 0)
{
	wchar_str wcs;

	if (!s || !*s) {
		return wcs;
	}

	if (n <= 0) {
		n = static_cast<int>(strlen(s));
	}

	int len = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, s, n, nullptr, 0);
	if (0 == len) {
		// use FormatMessage!!!
		throw std::runtime_error(__FUNCTION__ ": MultiByteToWideChar get size failed");
	}

	if (nullptr != wcs.alloc(len)) {
		len = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, s, n, wcs.data(), len);
		if (0 == len) {
			// use FormatMessage!!!
			throw std::runtime_error(__FUNCTION__ ": MultiByteToWideChar failed");
		}
	}

	return wcs;
}

