// addin_manager.h - manage add-in lifecycle
#pragma once
#include "../registry.h"
#include "../splitpath.h"
#include "xll_macrofun.h"

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

		// Copy add-in to Aim() and inform Excel
		AddinManager(bool prompt = false, bool add = false)
		{
			OPER name(Excel(xlGetName));
			path sp(name.as_cstr());
			OPER fname(sp.fname); // descriptive name
			OPER remove = Remove(fname); // move to Aim()

			if (remove) {
				if (prompt) { // check file times???
					OPER msg = OPER("Replace ") & fname & OPER("?");
					OPER result = Excel(xlcAlert, msg, OPER(1));
					if (!result) {
						Add(fname); // put it back
						return;
					}
				}
				Reg::Key aim(HKEY_CURRENT_USER, Aim());
				RegDeleteKey(aim, name.as_cstr());
			}

			OPER tpl(Template());
			tpl.append(sp.basename().c_str());
			// copy and overwrite
			CopyFile(name.as_cstr(), tpl.as_cstr(), FALSE);
			New(tpl);

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

		// Adds an add-in to the working set using the descriptive name in the Add-Ins dialog box.
		// HKCU\Software\Microsoft\Office\_version_\Excel\Options\Open<N>
		static OPER Add(const OPER& name)
		{
			return Excel(xlcAddinManager, OPER(1), name);
		}

		// Removes an add-in from the working set using the descriptive name in the Add-Ins dialog box.
		// HKCU\Software\Microsoft\Office\_version_\Excel\Options\Open<N>
		static OPER Remove(const OPER& name)
		{
			// throws an exception!!!
			// do by editing registry
			return Excel(xlcAddinManager, OPER(2), name);
		}

		// Adds a new add-in to the list of add-ins that Microsoft Excel knows about. 
		// HKCU\Software\Microsoft\Office\_version_\Excel\Add-in Manager 
		static OPER New(const OPER& get_name = Excel(xlGetName))
		{
			return Excel(xlcAddinManager, OPER(3), get_name);
		}

		// full path if in Aim()
		static OPER Exists(const OPER& name)
		{
			Reg::Key aim(HKEY_CURRENT_USER, Aim());

			for (const auto& key : aim.Keys()) {
				path sp(key);
				if (name == OPER(sp.fname)) {
					return OPER(key);
				}
			}

			return ErrNA;
		}

		// Remove registry entries.
		static OPER Delete(const OPER& name)
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
				tpl.append();
			}

			return tpl.val.str + 1;
		}

	};

} // namespace xll
