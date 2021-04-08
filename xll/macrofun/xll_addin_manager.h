// addin_manager.h - manage add-in lifecycle
#pragma once
#include <compare>
#include "xll_macrofun.h"
#include "../win.h"
#include "../registry.h"
#include "../splitpath.h"

namespace xll {

	/// Manage Excel add-in lifecycle.
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
		// Excel moves AIM value to OPT and add OPEN<n+1> value to load at startup.
		OPER Add(OPER name = Missing) const
		{
			if (name.is_missing()) {
				name = OPER(path.fname);
			}

			if (Known(name)) {
				ensure(Excel(xlcAddinManager, OPER(1), name));
			}

			return Open(name);
		}

		// Removes an add-in from the working set using the descriptive name in the Add-Ins dialog box.
		// Move OPT value to AIM. This function is not fast.
		OPER Remove(OPER name = Missing) const
		{
			if (name.is_missing()) {
				name = OPER(path.fname);
			}

			if (Open(name)) {
				ensure(Excel(xlcAddinManager, OPER(2), name));
			}

			return Known();
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
				ensure(Remove(name)); // move from OPT to AIM
			}
			OPER known = Known(name);
			if (known) {
				// There is no Excel function for this.
				Reg::Key aim(HKEY_CURRENT_USER, Aim());
				ensure(aim and ERROR_SUCCESS == RegDeleteValue(aim, known.as_cstr()));
			}

			return known;
		}

		// Full path if descriptive name matches AIM value.
		OPER Known(OPER name = Missing) const
		{
			OPER known = ErrNA;

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
						known = value.name.c_str();

						break;
					}
				}
			}

			return known;
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
		inline static const TCHAR OPT[] = TEXT("\\Excel\\Options");
		
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


	class Installer {
		OPER dir; // installation directory
		typedef int STATE;
		STATE state;
	public:
		enum {
			INIT = 0,
			KNOWN = 1, // in AIM
			OPEN = 2,  // in OPT
			EXISTS = 4,// in dir
		};
		static inline const char* install_doc = R"(
The Excel add-in manager uses two registry keys to manage the add-in lifecycle
<dl>
<dt>AIM: <code>HKCU\Software\Microsoft\Office\<i>version</i>\Excel\Add-in Manager</code></dt>
<dd>
Values in AIM have name the <i>full path</i> to the add-in and no data.
</dd>
<dt>OPT: <code>HKCU\Software\Microsoft\Office\<i>version</i>\Excel\Options</code></dt>
<dd>
Values in OPT have name <code>OPEN<i>n</i></code> 
and data <code>"/R \"<i>full path</i>\""</code> where <i>n</i>
is computed by Excel.
</dd>
</dl>
Excel does not write these values to the registry until the Excel session is terminated.
)";
		Installer(const OPER& _dir = OPER(Template()))
			: state(INIT), dir(_dir)
		{
			if (dir.val.str[dir.val.str[0]] != '\\') {
				dir.append("\\");
			}
			ensure(DirExists(dir.as_cstr()));
		}
		STATE State() const
		{
			return state;
		}

		// copy to installation dir and let add-in manager know
		STATE Install(OPER xll = Excel(xlGetName))
		{
			using Win::File;
			using Win::WriteTime;

			xll::path path(xll.as_cstr());
			OPER name(path.fname); // descriptive name
			OPER install = dir & OPER(path.basename().c_str());
			AddinManager aim(install);

			// existing state
			OPER known = aim.Known();
			OPER open = aim.Open();
			if (open) {
				// move from OPT to AIM
				known = aim.Remove();
				ensure(known);
			}

			if (FileExists(install.as_cstr())) {
				state |= EXISTS;
				// ask before clobbering
				if (WriteTime(File(xll.as_cstr())) > WriteTime(File(install.as_cstr()))) {
					OPER msg = OPER("Replace ") & OPER(path.fname) & OPER(".xll with new version?");
					OPER copy = Excel(xlcAlert, msg, OPER(1));
					if (copy) {
						ensure(CopyFile(xll.as_cstr(), install.as_cstr(), FALSE));
						state |= EXISTS;
					}
				}
			}
			else {
				ensure(CopyFile(xll.as_cstr(), install.as_cstr(), FALSE));
				state |= EXISTS;
			}
			ensure(state & EXISTS);

			if (known) {
				ensure(aim.Delete()); // remove old registry entries
			}

			ensure(aim.New()); // add to AIM
			state |= KNOWN;
			if (open) {
				ensure(aim.Add()); // move to OPT
				state |= OPEN;
				state &= ~KNOWN;
			}

			return state; // == EXISTS + KNOWN if new install
		}

		static inline const char* uninstall_doc = R"(
Delete add-in manager entries and installed add-in.
)";
		STATE Uninstall(const OPER& name)
		{
			AddinManager aim(dir & name & OPER(".xll"));

			aim.Delete();

			if (FileExists(aim.xll.as_cstr())) {
				state |= EXISTS;
				if (DeleteFile(aim.xll.as_cstr())) {
					state &= ~EXISTS;
				}
			}

			return state; // should == INIT
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
