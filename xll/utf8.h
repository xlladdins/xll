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

	/// <summary>
	/// View of contiguous memory.
	/// </summary>
	/// <typeparam name="T"></typeparam>
	template<class T>
	class view {
		T* p;
		DWORD n;
	public:
		view(T* p = nullptr, DWORD n = 0)
			: p(p), n(n)
		{ }
		view(const view&) = default;
		view& operator=(const view&) = default;
		virtual ~view()
		{ }

		explicit operator bool() const
		{
			return n != 0;
		}

		operator T* ()
		{
			return p;
		}

		void ptr(T* p_)
		{
			p = p_;
		}

		DWORD size() const
		{
			return n;
		}
		void size(DWORD n_)
		{
			n = n_;
		}
		// subclasses override and do checking
		virtual view& append(const T* p_, DWORD n_)
		{
			std::copy(p_, p_ + n_, p + n);
			n += n_;

			return *this;
		}
	};

	/// <summary>
	/// Memory mapped file class
	/// </summary>
	class mem_view : public view<char> {
		HANDLE h;
	public:
		/// <summary>
		/// Map file or temporary anonymous memory.
		/// </summary>
		/// <param name="h"></param>
		/// <param name="len"></param>
		mem_view(HANDLE h_ = INVALID_HANDLE_VALUE, DWORD len = 1 << 20)
			: h(CreateFileMapping(h_, 0, PAGE_READWRITE, 0, len, nullptr))
		{
			if (h == NULL) {
				throw std::runtime_error("mem_view: CreateFileMapping failed");
			}
			ptr((char*)MapViewOfFile(h, FILE_MAP_ALL_ACCESS, 0, 0, len));
		}
		mem_view(const mem_view&) = delete;
		mem_view& operator=(const mem_view&) = delete;
		~mem_view()
		{
			UnmapViewOfFile(*this);
			CloseHandle(h);
		}
	};

	class str_view : public view<char>
	{
		DWORD N; // maximum string length
	public:
		str_view(char* p, DWORD N)
			: view(p, 0), N(N)
		{ }
		template<size_t N>
		str_view(char (&p)[N])
			: str_view(p, N)
		{ }
		str_view(const str_view&) = delete;
		str_view& operator=(const str_view&) = delete;
		~str_view()
		{ }

		DWORD capacity() const
		{
			return N;
		}

		str_view& append(const char* str, DWORD len) override
		{
			if (len + size() > N) {
				len = N - size(); // ??? throw
			}
			view::append(str, len);

			return *this;
		}
	};

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
