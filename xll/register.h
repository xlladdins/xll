// register.h - Register an add-in
#pragma once
#include "shlwapi.h"
#include "splitpath.h"
#include "args.h"

namespace xll{

	inline OPER Register(Args& args)
	{
		args.key("moduleText") = Excel(xlGetName);

		unsigned count = 11 + static_cast<int>(args.ArgumentCount());
		std::vector<const XLOPERX*> oper(count);

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
		args.Procedure(procedure);

		OPER helpTopic = args.HelpTopic();
		if (!helpTopic and args.Documentation().length() > 0) {
			path sp(args.key("moduleText").as_cstr());
			OPER name(sp.dirname().c_str());
			name.append(args.FunctionText().safe());
			name.append(".html!0");
			name.as_cstr(); // null terminate

			xll::traits<XLOPERX>::xchar buf[2048];
			DWORD len = 2048;
			ensure(S_OK == UrlCreateFromPath(name.val.str + 1, buf, &len, 0));

			helpTopic = OPER(buf, len);
		}
		// Help URLs must end with "!0"
		else if (helpTopic.is_str()) {
			const auto& ht = helpTopic.val.str;
			if (ht[0] > 2 and ht[ht[0]] != '0' and ht[ht[0] - 1] != '!') {
				helpTopic &= OPER("!0");
			}
		}
		args.HelpTopic(helpTopic);

		// indicates function returns a handle
		if (args.FunctionText().val.str[1] == '\\') {
			ensure(args.isUncalced());
		}

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
}
