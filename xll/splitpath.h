// splitpath.h - _splitpath_s wrapper
#pragma once
#include <stdexcept>
#include <string>
#include <windows.h>
#include <tchar.h>
//#include <fileapi.h>

namespace xll {

	inline errno_t splitpath(const char* path, char* drive, char* dir, char* fname, char* ext)
	{
		return _splitpath_s(path, drive, _MAX_DRIVE, dir, _MAX_DIR, fname, _MAX_FNAME, ext, _MAX_EXT);
	}
	inline errno_t splitpath(const wchar_t* path, wchar_t* drive, wchar_t* dir, wchar_t* fname, wchar_t* ext)
	{
		return _wsplitpath_s(path, drive, _MAX_DRIVE, dir, _MAX_DIR, fname, _MAX_FNAME, ext, _MAX_EXT);
	}

	inline errno_t makepath(char* path, size_t len, const char* drive, const char* dir, const char* fname, const char* ext)
	{
		return _makepath_s(path, len, drive, dir, fname, ext);
	}
	inline errno_t makepath(wchar_t* path, size_t len, const wchar_t* drive, const wchar_t* dir, const wchar_t* fname, const wchar_t* ext)
	{
		return _wmakepath_s(path, len, drive, dir, fname, ext);
	}

	template<class T>
		requires (std::is_same_v<T, char> || std::is_same_v<T, wchar_t>)
	struct path {
		T drive[_MAX_DRIVE];
		T dir[_MAX_DIR];
		T fname[_MAX_FNAME];
		T ext[_MAX_EXT];

		path() noexcept
		{ }
		path(const T* p)
		{ 
			if (p) {
				if (0 != split(p)) {
					throw std::runtime_error(__FUNCTION__ ": splitpath failed");
				}
			}
		}
		path(const path&) = delete;
		path& operator=(const path&) = delete;
		~path()
		{ }

		errno_t split(const T* path)
		{
			return splitpath(path, drive, dir, fname, ext);
		}

		errno_t make(T* path, size_t len)
		{
			return makepath(path, len, drive, dir, fname, ext);
		}

		using string = std::basic_string<T>;

		string dirname() const
		{
			return string(drive).append(dir);
		}
		string basename() const
		{
			return string(fname).append(ext);
		}

	};

}