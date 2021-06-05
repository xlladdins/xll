// view.h - view of contiguous memory
#pragma once
#include <stdexcept>
#define NOMINMAX
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <handleapi.h>

namespace xll {

	/// <summary>
	/// Bare bones view of contiguous memory.
	/// </summary>
	/// <typeparam name="T">type of data</typeparam>
	template<class T>
	struct view {
		T* buf;
		DWORD len;

		view()
			: buf(nullptr), len(0)
		{ }
		view(T* buf, DWORD len)
			: buf(buf), len(len)
		{ }
		template<size_t N>
		view(T(&buf)[N])
			: view(buf, static_cast<DWORD>(N))
		{ }
		view(const view&) = default;
		view& operator=(const view&) = default;
		virtual ~view()
		{ }

		explicit operator bool() const
		{
			return len != 0;
		}
		bool operator==(const view& v) const
		{
			return len == v.len and std::equal(buf, buf + len, v.buf);
		}

		T back() const
		{
			return buf[len - 1];
		}
	};

	template<class T>
	inline view<T> trim_front(view<T> v, T t)
	{
		while (*v.buf == t) {
			++v.buf;
			--v.len;
		}

		return v;
	}
	template<class T>
	inline view<T> trim_back(view<T> v, T t)
	{
		while (v.back() == t) {
			--v.len;
		}

		return v;
	}
	template<class T>
	inline view<T> trim(view<T> v, T c = ' ')
	{
		return trim_front<T>(trim_back<T>(v, c), c);
	}

	/// <summary>
	/// Memory mapped file class
	/// </summary>
	/// <remarks>
	/// Use operating system features to buffer memory.
	/// Default size is 2^20 = 10^6 = 1MB, but does not
	/// allocate memory until used.
	/// </remarks>
	class mem_view : public view<char> {
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
			std::copy(s, s + n, buf + len);
			len += n;

			return *this;
		}
	};

} // namespace xll
