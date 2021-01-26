// register.h - Register an add-in
#pragma once
#include "args.h"

namespace xll{

	inline OPER Register(Args& args)
	{
		args.key("moduleText") = Excel(xlGetName);

		int count = 10 + static_cast<int>(args.ArgumentCount());
		std::vector<const XLOPERX*> oper(count + 1);

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
			oper[10 + i] = &args.ArgumentHelp(i);
		}
		// https://docs.microsoft.com/en-us/office/client-developer/excel/known-issues-in-excel-xll-development#argument-description-string-truncation-in-the-function-wizard
		OPER xEmpty("");
		oper[count] = &xEmpty;

		XLOPERX registerId = { .xltype = xltypeNil };
		int ret = traits<XLOPERX>::Excelv(xlfRegister, &registerId, count + 1, (XLOPERX**)&oper[0]);
		if (ret != xlretSuccess || registerId.xltype != xltypeNum) {
			OPER xMsg("Failed to register: ");
			Excel(xlcAlert, xMsg & args.FunctionText(), OPER(2));
		}

		return registerId;
	}
}
