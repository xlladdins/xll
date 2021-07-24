// register.h - Register an add-in
#pragma once
#include "shlwapi.h"
#include "splitpath.h"
#include "args.h"

namespace xll{

	inline OPER Register(Args& args)
	{
		OPER procedure = args.Procedure();
		ensure(procedure.xltype == xltypeStr && procedure.val.str[0] > 1);
		if (procedure.val.str[1] == '_') {
			// strip off for C functions
			procedure = OPER(procedure.val.str + 2, procedure.val.str[0] - 1);
		}
		else if (procedure.val.str[1] != '?') {
			// C++ mangled name must start with '?'
			procedure = OPER("?") & procedure;
		}

		OPER moduleText = Excel(xlGetName);
		OPER helpTopic = args.HelpTopic();
		if (!helpTopic and args.Documentation().length() > 0) {
			/*
			path sp(moduleText.as_cstr());
			OPER name(sp.dirname().c_str());
			*/
			OPER name(XLL_URL);
			name.append(args.FunctionText().safe());
			name.append(".html!0");
			name.as_cstr(); // null terminate

			xll::traits<XLOPERX>::xchar buf[2048];
			DWORD len = 2048;
			HRESULT result = UrlCreateFromPath(name.val.str + 1, buf, &len, 0);
			ensure(S_OK == result or S_FALSE == result);

			helpTopic = OPER(buf, len);
		}
		else if (helpTopic.is_str()) {
			// Help URLs must end with "!0"
			const auto& ht = helpTopic.val.str;
			if (ht[0] > 2 and ht[ht[0]] != '0' and ht[ht[0] - 1] != '!') {
				helpTopic &= OPER("!0");
			}
		}

		// indicates function returns a handle
		if (args.isHandle()) {
			ensure(args.isUncalced());
		}

		args.key("moduleText") = moduleText;
		args.key("procedure") = procedure;
		args.key("helpTopic") = helpTopic;

		unsigned count = 11 + static_cast<int>(args.ArgumentCount());
		std::vector<const XLOPERX*> oper(count);
		oper[0] = &args.ModuleText();
		oper[1] = &args.Procedure();
		oper[2] = &args.TypeText();
		oper[3] = &args.FunctionText();
		oper[4] = &args.ArgumentText();
		oper[5] = &args.MacroType();
		oper[6] = &args.Category();
		oper[7] = &args.ShortcutText();
		oper[8] = &args.HelpTopic();
		oper[9] = &args.FunctionHelp();
		for (unsigned i = 1; i <= args.ArgumentCount(); ++i) {
			oper[9 + i] = &args.ArgumentHelp(i);
		}
		// https://docs.microsoft.com/en-us/office/client-developer/excel/known-issues-in-excel-xll-development#argument-description-string-truncation-in-the-function-wizard
		OPER xEmpty("");
		oper.back() = &xEmpty;

		XLOPERX registerId = { .xltype = xltypeNil };
		int ret = traits<XLOPERX>::Excelv(xlfRegister, &registerId, count, (XLOPERX**)&oper[0]);
		if (ret != xlretSuccess || registerId.xltype != xltypeNum) {
			OPER xMsg("Failed to register: ");
			Excel(xlcAlert, xMsg & args.FunctionText(), OPER(2));
		}

		return registerId;
	}

	// Really unregister a function.
	// https://docs.microsoft.com/en-us/office/client-developer/excel/known-issues-in-excel-xll-development#unregistering-xll-commands-and-functions
	// https://stackoverflow.com/questions/15343282/how-to-remove-an-excel-udf-programmatically
	inline OPER Unregister(const OPER& key)
	{
		OPER regid = Excel(xlfEvaluate, key);
		if (!regid) {
			return ErrValue;
		}
		Excel(xlfSetName, key);
		Excel(xlfUnregister, regid);

		regid = Excel(xlfRegister, Excel(xlGetName), 
			OPER("xlAutoRemove"), OPER(XLL_SHORT), key, Missing, OPER(2));
		Excel(xlfSetName, key);
		
		return Excel(xlfUnregister, regid);
	}
}
