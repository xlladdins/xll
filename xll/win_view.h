// xll_view.h - view of memory mapped file
#pragma once
#include <compare>
#include <stdexcept>
#define NOMINMAX
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <handleapi.h>
#include "fms_view.h"

namespace win {

	/// <summary>
	/// Memory mapped file class
	/// </summary>
	/// <remarks>
	/// Use operating system features to buffer memory.
	/// Default size is 2^20 = 10^6 = 1MB, but does not
	/// allocate memory until used.
	/// </remarks>
	class mem_view : public fms::view<char> {
		HANDLE h;
	public:
		/// <summary>
		/// Map file or temporary anonymous memory.
		/// </summary>
		/// <param name="h">optional handle to file</param>
		/// <param name="len">maximum size of buffer</param>
		mem_view(HANDLE h_ = INVALID_HANDLE_VALUE, DWORD len = 1 << 20)
			: h(CreateFileMapping(h_, 0, PAGE_READWRITE, 0, len, nullptr))
		{
			if (h == NULL) {
				throw std::runtime_error("mem_view: CreateFileMapping failed");
			}
			len = 0;
			buf = (char*)MapViewOfFile(h, FILE_MAP_ALL_ACCESS, 0, 0, len);
			buf += 2; // so first 2 chars can be overwritten for counted strings
		}
		mem_view(const mem_view&) = delete;
		mem_view& operator=(const mem_view&) = delete;
		~mem_view()
		{
			UnmapViewOfFile(buf);
			CloseHandle(h);
		}

		// Write to buffered memory.
		mem_view& append(const char* s, DWORD n)
		{
			CopyMemory(buf + len, s, n);
			len += n;

			return *this;
		}
	};

} // namespace xll
