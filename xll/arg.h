// arg.h - Argument for xlfRegister
#pragma once
#include "oper.h"

namespace xll {

	template<class X>
	struct XArg {
		XOPER<X> ModuleText;   // from xlGetName
		XOPER<X> Procedure;    // C function
		XOPER<X> TypeText;     // return type and arg codes 
		XOPER<X> FunctionText; // Excel function
		XOPER<X> ArgumentText; // Ctrl-Shift-A text
		XOPER<X> MacroType;    // function; 2 for macro; 0 for hidden
		XOPER<X> Category;     // for function wizard
		XOPER<X> ShortcutText; // single character for Ctrl-Shift-char shortcut
		XOPER<X> HelpTopic;    // filepath!HelpContextID or http://host/path!0
		XOPER<X> FunctionHelp; // for function wizard
		XOPER<X> ArgumentHelp[10]; // 1-based index...

		struct Arg {
			XOPER<X> type;
			XOPER<X> name;
			XOPER<X> help;
			//XOPER<X> init;
		};

		Args(std::initializer_list<Arg> args)
		{
			for (int i = 0; i < args.size(); ++i) {
				TypeText &= args[i].type;
				if (i > 0) {
					ArgumentText &= ',';
				}
				ArgumentText &= args[i].name;
				ArgumentHelp[i] = args[i].help;
			}
		}

		int Register()
		{
			if (ModuleText.xltype == xltypeMissing) {

			}

		}
	};

}
