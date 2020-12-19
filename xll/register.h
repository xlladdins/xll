// register.h - Register an add-in
#pragma once
#include "document.h"

namespace xll{

	template<class X>
	inline X Register(const XArgs<X>& args)
	{
		int count = 10 + static_cast<int>(args.ArgumentHelp().size());
		std::vector<const X*> oper(count + 1);

		XOPER<X> moduleText = XExcel<X>(xlGetName);

		XOPER<X> procedure = args.Procedure();
		// C++ mangled name must start with '?'
		ensure(procedure.xltype == xltypeStr && procedure.val.str[0] > 1);
		if (procedure.val.str[1] != '?' && procedure.val.str[1] != '_') {
			procedure = XOPER<X>("?") & procedure;
		}

		XOPER<X> helpTopic = args.HelpTopic();
		//if (helpTopic == XOPER<X>{}) {
			/*helpTopic =*/ xll::Document(args.FunctionText());
		//}
		// Help URLs must end with "!0"
		if (helpTopic.xltype == xltypeStr) {
			XOPER xFind = XExcel<X>(xlfFind, XOPER<X>("!"), helpTopic);
			if (xFind.xltype == xltypeErr && xFind.val.err == xlerrValue) {
				helpTopic &= XOPER<X>("!0");
			}
		}

		oper[0] = &moduleText;
		oper[1] = &procedure;
		oper[2] = &args.TypeText();
		oper[3] = &args.FunctionText();
		oper[4] = &args.ArgumentText();
		oper[5] = &args.MacroType();
		oper[6] = &args.Category();
		oper[7] = &args.ShortcutText();
		oper[8] = &helpTopic;
		oper[9] = &args.FunctionHelp();
		const std::vector<XOPER<X>>& argumentHelp = args.ArgumentHelp();
		for (size_t i = 0; i < argumentHelp.size(); ++i) {
			oper[10 + i] = &argumentHelp[i];
		}
		// https://docs.microsoft.com/en-us/office/client-developer/excel/known-issues-in-excel-xll-development#argument-description-string-truncation-in-the-function-wizard
		XOPER<X> xEmpty("");
		oper[count] = &xEmpty;

		X registerId = { .xltype = xltypeNil };
		int ret = traits<X>::Excelv(xlfRegister, &registerId, count + 1, (X**)&oper[0]);
		if (ret != xlretSuccess || registerId.xltype != xltypeNum) {
			XOPER<X> xMsg("Failed to register: ");
			XExcel<X>(xlcAlert, xMsg & args.FunctionText(), XOPER<X>(2));
		}

		return registerId;
	}
}
