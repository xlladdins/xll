// win.h - Windows specific code
#pragma once
#include <compare>
#include <utility>
#include <Windows.h>
#include <WinBase.h>

inline auto operator<=>(const FILETIME& a, const FILETIME& b)
{
	auto cmp = a.dwHighDateTime <=> b.dwHighDateTime;

	return cmp != 0 ? cmp : a.dwLowDateTime <=> b.dwLowDateTime;
}

// https://devblogs.microsoft.com/oldnewthing/20071023-00/
inline bool FileExists(PCTSTR file)
{
	DWORD dwAttrib = GetFileAttributes(file);

	return dwAttrib != INVALID_FILE_ATTRIBUTES and !(dwAttrib & FILE_ATTRIBUTE_DIRECTORY);
}
inline bool DirExists(PCTSTR file)
{
	DWORD dwAttrib = GetFileAttributes(file);

	return dwAttrib != INVALID_FILE_ATTRIBUTES and (dwAttrib & FILE_ATTRIBUTE_DIRECTORY);
}

namespace Win {

	inline char* GetFormatMessage(DWORD id = GetLastError())
	{
		// not thread safe
		static constexpr DWORD size = 1024;
		static char buf[size];

		FormatMessageA(FORMAT_MESSAGE_FROM_SYSTEM, 0, id, 0, buf, size, 0);

		return buf;
	}

	// Parameterized handle class
	template<class H, BOOL(WINAPI *D)(H)>
	class Hnd {
		H h;
	public:
		Hnd(H h = INVALID_HANDLE_VALUE)
			: h(h)
		{ }
		Hnd(const Hnd&) = delete;
		Hnd& operator=(const Hnd&) = delete;
		Hnd(Hnd&& h) noexcept
			: h(std::exchange(h.h, INVALID_HANDLE_VALUE))
		{ }
		Hnd& operator=(Hnd&& h_) noexcept
		{
			if (this != &h_) {
				h = std::exchange(h_.h, INVALID_HANDLE_VALUE);
			}

			return *this;
		}
		~Hnd()
		{
			D(h);
		}
		operator H()
		{
			return h;
		}
	};

	using Handle = Hnd<HANDLE, CloseHandle>;

	class File {
		HANDLE h;
	public:
		File(LPCTSTR lpFileName,
			 DWORD   dwDesiredAccess = GENERIC_READ,
			 DWORD   dwCreationDisposition = OPEN_EXISTING,
		 	 DWORD   dwFlagsAndAttributes = FILE_ATTRIBUTE_NORMAL,
			 DWORD   dwShareMode = 0 // no sharing
		)
			: h(CreateFile(lpFileName, dwDesiredAccess, dwShareMode, NULL, dwCreationDisposition, dwFlagsAndAttributes, NULL))
		{ }
		File(const File&) = delete;
		File& operator=(const File&) = delete;
		~File()
		{
			CloseHandle(h);
		}

		operator HANDLE() const
		{
			return h;
		}
	};

	// date and time the file or directory was created.
	inline FILETIME CreateTime(const File& hFile)
	{
		FILETIME create = { 0, 0 };

		GetFileTime(hFile, &create, NULL, NULL);

		return create;
	}
	// last time the file or directory was written, read, or run (if executable).
	inline FILETIME AccessTime(const File& hFile)
	{
		FILETIME access = { 0, 0 };

		GetFileTime(hFile, NULL, &access, NULL);

		return access;
	}
	// date and time the file or directory was last written, truncated, or overwritten 
	inline FILETIME WriteTime(const File& hFile)
	{
		FILETIME write = { 0, 0 };

		GetFileTime(hFile, NULL, NULL, &write);
		
		return write;
	}

}
