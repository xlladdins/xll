// platform.h - Detect 64 or 32 bit Excel.
#pragma once
#include <Windows.h>

namespace xll {

	// return "x64" or "x84", or "" on failure
	inline const char* Platform(void)
	{
		static char platform[4] = { 0 };

		if (!platform[0]) {
			LPCSTR key = "SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration";
			LPCSTR value = "Platform";
			DWORD len = sizeof(platform);

			RegGetValueA(HKEY_LOCAL_MACHINE, key, value, RRF_RT_REG_SZ, 0, (PVOID)platform, &len);
		}

		return platform;
	}

}
