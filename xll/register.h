// register.h - Register an add-in
#pragma once
#include "shlwapi.h"
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

		// indicates function returns a handle
		if (args.FunctionText().val.str[1] == '\\') {
			ensure(args.isUncalced());
		}

		OPER helpTopic = args.HelpTopic();
		if (!helpTopic and args.Documentation().length() > 0) {
			OPER name = args.key("moduleText");
			name.append(); // null terminate

			xll::traits<XLOPERX>::xchar buf[2048];
			DWORD len = 2048;
			ensure(S_OK == UrlCreateFromPath(name.val.str + 1, buf, &len, 0));

			helpTopic = buf;
			unsigned slash = 0; // last index of slash
			for (unsigned i = 1; i <= helpTopic.val.str[0]; ++i) {
				if (helpTopic.val.str[i] == '/') {
					slash = i;
				}
			}
			helpTopic = Excel(xlfLeft, helpTopic, OPER(slash));
			// remove unsafe character
			helpTopic = Excel(xlfSubstitute, helpTopic, OPER("\\"), OPER(""));
			helpTopic.append(args.FunctionText());
			helpTopic.append(".html");
		}
		// Help URLs must end with "!0"
		if (helpTopic.xltype & xltypeStr) { 
			XOPER xFind = Excel(xlfFind, OPER("!"), helpTopic);
			if (xFind.xltype == xltypeErr && xFind.val.err == xlerrValue) {
				helpTopic &= OPER("!0");
			}
		}

		oper[0] = &args.ModuleText();
		oper[1] = &procedure;
		oper[2] = &args.TypeText();
		oper[3] = &args.FunctionText();
		oper[4] = &args.ArgumentText();
		oper[5] = &args.MacroType();
		oper[6] = &args.Category();
		oper[7] = &args.ShortcutText();
		oper[8] = &helpTopic;
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
