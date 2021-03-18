// addin_manager.h - manage add-in lifecycle
#pragma once
#include "xll/xll/registry.h"
#include "xll/xll/splitpath.h"
#include "xll/xll/macrofun/xll_macrofun.h"

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
	struct AddinManager {
//			path<char> sp;
//			sp.split(Excel4(xlGetName).to_string().c_str());
//			name = sp.fname;
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

		// Remove registry entries.
		static OPER Delete(void)
		{
			// Remove();
			// iterate over Add-in Manager entries
			// and match name
			
			return OPER(true);
		}
	private:
		inline static const TCHAR OFFICE[] = TEXT("Software\\Microsoft\\Office\\");
		inline static const TCHAR OPEN[] = TEXT("\\Excel\\Options\\Open");
		inline static const TCHAR AIM[] = TEXT("\\Excel\\Add-in Manager");
		
		static OPER Version()
		{
			return Excel(xlfGetWorkspace, OPER(2));
		}
		// Add/Remove registery key
		static const TCHAR* Open()
		{
			static std::basic_string<TCHAR> open;

			if (open.length() == 0) {
				open = OFFICE;
				const OPER& ver = Version();
				open.append(ver.val.str + 1, ver.val.str[0]);
				open.append(OPEN);
			}

			return open.c_str();
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
	public:
		// Excel template directory
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
