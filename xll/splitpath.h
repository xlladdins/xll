// splitpath.h - _splitpath_s wrapper
#pragma once
#include <stdexcept>
#include <string>
#include <Windows.h>
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
		// requires (std::is_same_v<T, char> || std::is_same_v<T, wchar_t>)
	struct path {
		T drive[_MAX_DRIVE] = {0};
		T dir[_MAX_DIR] = { 0 };
		T fname[_MAX_FNAME] = { 0 };
		T ext[_MAX_EXT] = { 0 };

		path() noexcept
		{ }
		path(const T* p)
		{ 
			if (p) {
				split(p);
			}
		}
		path(const path&) = delete;
		path& operator=(const path&) = delete;
		~path()
		{ }

		explicit operator bool() const
		{
			return drive[0];
		}

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
		string file() const
		{
			return dirname().append(basename());
		}
	};

}