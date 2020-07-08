// arg.h - Argument for xlfRegister
#pragma once
#include <initializer_list>
#include <vector>
#include "oper.h"

namespace xll {

	template<class X>
	class XArg {
		using xcstr = typename traits<X>::xcstr;

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
		std::vector<XOPER<X>> ArgumentHelp;

	public:

		XArg(xcstr Type, xcstr Procedure, xcstr FunctionText)
			: TypeText(Type), Procedure(Procedure), FunctionText(FunctionText)
		{
			MacroType = 1; // Function
		}
		XArg(xcstr Procedure, xcstr FunctionText)
			: Procedure(Procedure), FunctionText(FunctionText)
		{
			MacroType = 2; // Macro
		}

		struct Arg {
			xcstr type;
			xcstr name;
			xcstr help;
			//xcstr init;
		};

		XArg& Args(std::initializer_list<Arg> arg)
		{
			/*
			for (int i = 0; i < arg.size(); ++i) {
				TypeText &= arg[i].type;
				if (i > 0) {
					ArgumentText &= ',';
				}
				ArgumentText &= arg[i].name;
				ArgumentHelp[i] = arg[i].help;
			}
			*/

			return *this;
		}
		/*
		XArg& Category(xcstr category)
		{
			//Category = category;

			return *this;
		}
		XArg& FunctionHelp(xcstr functionHelp)
		{
			//FunctionHelp = functionHelp;

			return *this;
		}
		*/
		int Register()
		{
			int count = 10 + ArgumentHelp.size();
			std::vector<traits<X>::type> oper(count);

			if (ModuleText.xltype == xltypeMissing) {
				X x;
				traits::Excel(xlGetName, &x, 0, nullptr);
				ModuleText = x;
			}

			oper[0] = &ModuleText;
			oper[1] = &Procedure;    // C function
			oper[2] = &TypeText;     // return type and arg codes 
			oper[3] = &FunctionText; // Excel function
			oper[4] = &ArgumentText; // Ctrl-Shift-A text
			oper[5] = &MacroType;    // function; 2 for macro; 0 for hidden
			oper[6] = &Category;     // for function wizard
			oper[7] = &ShortcutText; // single character for Ctrl-Shift-char shortcut
			oper[8] = &HelpTopic;    // filepath!HelpContextID or http://host/path!0
			oper[9] = &FunctionHelp; // for function wizard
			for (int i = 0; i < ArgumentHelp.size(); ++i) {
				oper[10 + i] = &ArgumentHelp[i];
			}

			return traits<X>::Excel(xlfRegister, NULL, count, &oper[0]);
		}
	};

	using Function = XArg<XLOPER>;
	using Function12 = XArg<XLOPER12>;
	using FunctionX = XArg<XLOPERX>;
}
