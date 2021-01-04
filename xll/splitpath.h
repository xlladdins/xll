// splitpath.h - _splitpath_s wrapper
#pragma once
#include <windows.h>
#include <fileapi.h>
#include <stdexcept>
#include <string>

namespace xll {

	struct splitpath {
		char drive[_MAX_DRIVE];
		char dir[_MAX_DIR];
		char fname[_MAX_FNAME];
		char ext[_MAX_EXT];
		splitpath(const char* path)
		{
			errno_t err = _splitpath_s(path, drive, dir, fname, ext);
			if (err != 0) {
				throw std::runtime_error("_splitpath_s failed");
			}
		}
		std::string dirname() const
		{
			return std::string(drive).append(dir);
		}
	};

}