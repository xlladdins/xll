// addin_manager.h - manage add-in lifecycle
#pragma once
#include <compare>
#include "../registry.h"
#include "../splitpath.h"
#include "xll_macrofun.h"

inline auto operator<=>(const FILETIME& a, const FILETIME& b)
{
	auto cmp = a.dwHighDateTime <=> b.dwHighDateTime;

	return cmp != 0 ? cmp : a.dwLowDateTime <=> b.dwLowDateTime;
}

// date and time the file or directory was last written to, truncated, or overwritten 
inline FILETIME GetFileWriteTime(PCTSTR file)
{
	FILETIME write = { 0, 0 };

	HANDLE hFile = CreateFile(file, 0, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_READONLY, NULL);
	if (hFile != INVALID_HANDLE_VALUE) {
		GetFileTime(hFile, NULL, NULL, &write);
		CloseHandle(hFile);
	}

	return write;
}

inline bool FileExists(PCTSTR file)
{
	DWORD dwAttrib = GetFileAttributes(file);
	
	return dwAttrib != INVALID_FILE_ATTRIBUTES and !(dwAttrib & FILE_ATTRIBUTE_DIRECTORY);
}

namespace xll {

	/// <summary>
	/// Manage add-in lifecycle
	/// </summary>
	/// <remarks>
	/// Excel uses two registry keys for the Add-in Manager:
	/// AIM: HKCU\Software\Microsoft\Office\<version>\Excel\Add-in Manager
	///   with values in AIM having name the <full path> to the add-in.
	/// OPT: HKCU\Software\Microsoft\Office\<version>\Excel\Options
	///   with values in OPT having name OPEN<n> and data "/R <full path>"
	/// </remarks>
	/// https://xlladdins.github.io/Excel4Macros/addin.manager.html
	struct AddinManager {

		OPER xll; // full path to add-in
		xll::path<TCHAR> path; // split path
		
		AddinManager(const OPER& _xll = Excel(xlGetName))
			: xll(_xll), path(xll.as_cstr())
		{
			// xll is NULL terminated
		}

		// Adds an add-in to the working set using the descriptive name in the Add-Ins dialog box.
		// Move AIM value to OPT and add OPEN<n+1> value to load at startup.
		OPER Add(OPER name = Missing) const
		{
			if (name.is_missing()) {
				name = OPER(path.fname);
			}

			return Excel(xlcAddinManager, OPER(1), name);
		}

		// Removes an add-in from the working set using the descriptive name in the Add-Ins dialog box.
		// Move OPT value to AIM. This function is not fast.
		OPER Remove(OPER name = Missing) const
		{
			if (name.is_missing()) {
				name = OPER(path.fname);
			}
			return Excel(xlcAddinManager, OPER(2), name);
		}

		// Adds a new add-in to the list of add-ins that Microsoft Excel knows about. 
		// Create AIM value using xll.
		OPER New() const 
		{
			return Excel(xlcAddinManager, OPER(3), xll);
		}

		// Remove all registry entries. Does not uninstall.
		OPER Delete(OPER name = Missing) const
		{
			if (name.is_missing()) {
				name = OPER(path.fname);
			}

			OPER open = Open(name);
			if (open) {
				Remove(name); // move from OPT to AIM
			}
			OPER known = Known(name);
			if (known) {
				Reg::Key aim(HKEY_CURRENT_USER, Aim());
				ensure(ERROR_SUCCESS == RegDeleteValue(aim, known.as_cstr()));
			}

			return known;
		}

		// Full path if descriptive name matches AIM value.
		OPER Known(OPER name = Missing) const
		{
			OPER aim_value = ErrNA;

			if (!name) {
				name = OPER(path.fname);
			}

			// all known add-ins
			Reg::Key aim(HKEY_CURRENT_USER, Aim());
			for (const Reg::Value& value : aim.Values()) {
				if (value.type == REG_SZ) {
					// value name is the full path to add-in
					xll::path split(value.name.c_str());
					if (OPER(split.fname) == name) {
						aim_value = value.name.c_str();

						break;
					}
				}
			}

			return aim_value;
		}

		// OPT value OPEN<n> name if loaded on startup
		OPER Open(OPER name = Missing) const
		{
			OPER open = ErrNA;

			if (name.is_missing()) {
				name = path.basename().c_str();
			}

			Reg::Key opt(HKEY_CURRENT_USER, Opt());
			for (const Reg::Value& value : opt.Values()) {
				if (value.type == REG_SZ) {
					PCTSTR match = _tcsstr(value, name.as_cstr());
					if (match and 0 == _tcsncmp(match, name.val.str + 1, name.val.str[0])) {
						open = value.name.c_str();

						break;
					}
				}
			}

			return open;
		}

		// Registry entries used by add-in manager.
		inline static const TCHAR OFFICE[] = TEXT("Software\\Microsoft\\Office\\");
		inline static const TCHAR AIM[] = TEXT("\\Excel\\Add-in Manager");
		inline static const TCHAR OPT[] = TEXT("Excel\\Options");
		
		static OPER Version()
		{
			return Excel(xlfGetWorkspace, OPER(2));
		}

		// Add-in Manager registry key
		static const TCHAR* Aim()
		{
			static std::basic_string<TCHAR> aim;

			if (aim.length() == 0) {
				aim = OFFICE;
				const OPER& ver = Version();
				aim.append(ver.val.str + 1, ver.val.str[0]);
				aim.append(AIM);
			}

			return aim.c_str();
		}

		static const TCHAR* Opt()
		{
			static std::basic_string<TCHAR> opt;

			if (opt.length() == 0) {
				opt = OFFICE;
				const OPER& ver = Version();
				opt.append(ver.val.str + 1, ver.val.str[0]);
				opt.append(OPT);
			}

			return opt.c_str();
		}
	};

	// copy to installation dir and let add-in manager know
	class Installer {
		typedef int STATE;
		STATE state;
		OPER dir; // installation directory
	public:
		enum {
			INIT = 0,
			KNOWN = 1, // in AIM
			OPEN = 2,  // in OPT
			EXISTS = 4,// in dir
		};

		Installer(const OPER& dir = OPER(Template()))
			: state(INIT), dir(dir)
		{ }
		STATE State() const
		{
			return state;
		}
		// Install xll in dir
		STATE Install(OPER xll = Excel(xlGetName))
		{
			xll::path path(xll.as_cstr());
			OPER install = dir & OPER("\\") & OPER(path.basename().c_str());
			AddinManager aim(install);

			if (FileExists(install.as_cstr())) {
				state |= EXISTS;
			}
			if (aim.Known()) {
				state |= KNOWN;
			}
			else if (aim.Open()) {
				state |= OPEN;
			}

			if (state & EXISTS) {
				if (GetFileWriteTime(xll.as_cstr()) > GetFileWriteTime(install.as_cstr())) {
					OPER msg = OPER("Replace ") & OPER(path.fname) & OPER(" with new version?");
					OPER result = Excel(xlcAlert, msg, OPER(2));
					if (result) {
						ensure(CopyFile(xll.as_cstr(), install.as_cstr(), FALSE));
					}
				}
			}
			else {
				ensure(CopyFile(xll.as_cstr(), install.as_cstr(), FALSE));
				state |= EXISTS;
			}

			return state; // == EXISTS + KNOWN if new install
		}

		STATE Uninstall(const OPER& name)
		{
			AddinManager aim(dir);

			OPER file = aim.Delete(name);
			if (file) {
				if (FileExists(file.as_cstr())) {
					state |= EXISTS;
				}
				if (DeleteFile(file.as_cstr())) {
					state &= ~EXISTS;
				}
			}

			return state;
		}

		// Excel template directory is trusted location.
		static const OPER& Template()
		{
			static OPER tpl;

			if (!tpl) {
				tpl = getenv("AppData");
				tpl.append("\\Microsoft\\Templates\\");
				tpl.as_cstr(); // add NULL
			}

			return tpl;
		}

	};

} // namespace xll
