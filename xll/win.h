// win.h - Windows specific code
#pragma once
#include <utility>
#include "defines.h"

namespace Win {

	inline char* GetFormatMessage(DWORD id)
	{
		// not thread safe
		static constexpr DWORD size = 1024;
		static char buf[size];

		FormatMessageA(FORMAT_MESSAGE_FROM_SYSTEM, 0, id, 0, buf, size, 0);

		return buf;
	}

	class Handle {
		HANDLE h;
	public:
		Handle(HANDLE h = INVALID_HANDLE_VALUE)
			: h(h)
		{ }
		Handle(const Handle&) = delete;
		Handle& operator=(const Handle&) = delete;
		Handle(Handle&& h) noexcept
			: h(std::exchange(h.h, INVALID_HANDLE_VALUE))
		{ }
		Handle& operator=(Handle&& h_) noexcept
		{
			if (this != &h_) {
				h = std::exchange(h_.h, INVALID_HANDLE_VALUE);
			}

			return *this;
		}
		~Handle()
		{
			CloseHandle(h);
		}
		operator HANDLE()
		{
			return h;
		}
	};

}
