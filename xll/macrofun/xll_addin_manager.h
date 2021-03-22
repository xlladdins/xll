// addin_manager.h - manage add-in lifecycle
#pragma once
#include <compare>
#include "../registry.h"
#include "../splitpath.h"
#include "xll_macrofun.h"

inline auto operator<=>(const FILETIME& a, const FILETIME& b)
{
	auto cmp = a.dwHighDateTime <=> b.dwHighDateTime;
	if (cmp != 0) {
		return cmp;
	}

	return a.dwLowDateTime <=> b.dwLowDateTime;
}

// date and time the file or directory was last written to, truncated, or overwritten 
inline FILETIME FileWriteTime(const TCHAR* file)
{
	HANDLE hFile = CreateFile(file, 0, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_READONLY, NULL);
	if (hFile == INVALID_HANDLE_VALUE) {
		throw std::runtime_error(__FUNCTION__ ": CreateFile failed");
	}
	
	FILETIME write;
	auto b = GetFileTime(hFile, NULL, NULL, &write);
	CloseHandle(hFile);
	if (!b) {
		throw std::runtime_error(__FUNCTION__ ": GetFileTime failed");
	}

	return write;
}

namespace xll {

	/// <summary>
	/// Manage add-in lifecycle
	/// </summary>
	/// <remarks>
	/// Excel uses two registry keys for the Add-in Manager.
	/// New keys are put under Aim() with key the <full path>
	/// to the add-in and no value
	/// Add/Remove puts keys under Open() with key OPEN<n>
	/// and value <code>/R "<full path>"</code>.
	/// Add increments <n> by 1 and removes the Aim() entry.
	/// Remove removes the OPEN<n> entry and shifts keys
	/// with larger <n> down and adds the full path under Aim().
	/// </remarks>
	/// https://xlladdins.github.io/Excel4Macros/addin.manager.html
	struct AddinManager {

		/// <summary>
		/// Check if newer add-in is available and install.
		/// </summary>
		/// <remarks>
		/// </remarks>
		
		OPER xll; // full path to add-in
		path<TCHAR> split;
		FILETIME write = { 0, 0 };
		AddinManager()
		{ }
		AddinManager(const OPER& get_name)
			: xll(get_name), split(xll.as_cstr()), write(FileWriteTime(xll.as_cstr()))
		{
		}
		/*
			path_name

			OPER name(Excel(xlGetName));
			path sp(name.as_cstr());
			OPER fname(sp.fname); // descriptive name
			OPER tpl(Template()); // Excel template folder
			tpl.append(sp.basename().c_str());

			// check file times
			if (FileWriteTime(name.as_cstr()) <= FileWriteTime(tpl.as_cstr())) {
				return;
			}

			OPER loaded = Remove(fname); // move to Aim() if loaded
			if (!loaded) {
				if (prompt) {
					OPER msg = OPER("Install ") & fname & OPER("?");
					OPER result = Excel(xlcAlert, msg, OPER(1));
					if (!result) {
						return;
					}
				}
				CopyFile(name.as_cstr(), tpl.as_cstr(), FALSE);
				New(tpl); // add to Aim()
			}
			else { // loaded
				if (prompt) {
					OPER msg = OPER("Replace ") & fname & OPER("?");
					OPER result = Excel(xlcAlert, msg, OPER(1));
					if (!result) {
						return;
					}
				}
				CopyFile(name.as_cstr(), tpl.as_cstr(), FALSE);
				Reg::Key aim(HKEY_CURRENT_USER, Aim());
				RegDeleteKey(aim, tpl.as_cstr());
				New(tpl); // add to Aim()
			}

			if (add) {
				if (prompt) {
					OPER msg = OPER("Load ") & fname & OPER(" when Excel starts?");
					OPER result = Excel(xlcAlert, msg, OPER(1));
					if (!result) {
						return;
					}
				}
				Add(fname);
			}
		}
		*/
		// Adds an add-in to the working set using the descriptive name in the Add-Ins dialog box.
		// HKCU\Software\Microsoft\Office\_version_\Excel\Options\Open<N>
		OPER Add(const OPER& name = OPER())
		{
			return Excel(xlcAddinManager, OPER(1), name ? name : OPER(split.fname));
		}

		// Removes an add-in from the working set using the descriptive name in the Add-Ins dialog box.
		// HKCU\Software\Microsoft\Office\_version_\Excel\Options\Open<N>
		OPER Remove(const OPER& name = OPER())
		{
			// throws an exception!!!
			// do by editing registry
			return Excel(xlcAddinManager, OPER(2), name ? name : OPER(split.fname));
		}

		// Adds a new add-in to the list of add-ins that Microsoft Excel knows about. 
		// HKCU\Software\Microsoft\Office\_version_\Excel\Add-in Manager 
		OPER New(const OPER& get_name = OPER())
		{
			return Excel(xlcAddinManager, OPER(3), get_name ? get_name : xll);
		}

		// get_name more recently written
		bool Newer(OPER get_name) const
		{
			return FileWriteTime(get_name.as_cstr()) >= write;
		}

		// full path if in Aim()
		OPER Exists(OPER name = OPER())
		{
			OPER get_name = ErrNA;

			if (!name) {
				name = OPER(split.fname);
			}

			// all known add-ins
			Reg::Key aim(HKEY_CURRENT_USER, Aim());
			for (const auto& val : aim.Values()) {
				if (val.type == REG_SZ) {
					const TCHAR* sz = val;
					path sp(sz);
					if (OPER(sp.fname) == name) {
						get_name = sz;

						break;
					}
				}
			}

			return get_name;
		}

		// Remove registry entries.
		OPER Delete(const OPER& name)
		{
			try {
				Remove(name); // move from Open() to Aim()
				OPER key = Exists(name);
				if (key) {
					Reg::Key aim(HKEY_CURRENT_USER, Aim());
					RegDeleteKey(aim, key.as_cstr());
				}
			}
			catch (const std::exception& ex) {
				XLL_ERROR(ex.what());

				return OPER(false);
			}
			catch (...) {
				XLL_ERROR("AddinManager::Delete: unknown exception");

				return OPER(false);
			}
			
			return OPER(true);
		}

		inline static const TCHAR OFFICE[] = TEXT("Software\\Microsoft\\Office\\");
		inline static const TCHAR AIM[] = TEXT("\\Excel\\Add-in Manager");
		
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

		// Excel template directory is trusted location.
		static const TCHAR* Template()
		{
			static OPER tpl;
			
			if (!tpl) {
				tpl = getenv("AppData");
				tpl.append("\\Microsoft\\Templates\\");
			}

			return tpl.as_cstr();
		}

	};

} // namespace xll
