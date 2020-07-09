// arg.h - Argument for xlfRegister
#pragma once
#include <initializer_list>
#include <vector>
#include "oper.h"

namespace xll {

	template<class X>
	class XArg {
		using xcstr = typename traits<X>::xcstr;

		XOPER<X> moduleText;   // from xlGetName
		XOPER<X> procedure;    // C function
		XOPER<X> typeText;     // return type and arg codes 
		XOPER<X> functionText; // Excel function
		XOPER<X> argumentText; // Ctrl-Shift-A text
		XOPER<X> macroType;    // function; 2 for macro; 0 for hidden
		XOPER<X> category;     // for function wizard
		XOPER<X> shortcutText; // single character for Ctrl-Shift-char shortcut
		XOPER<X> helpTopic;    // filepath!HelpContextID or http://host/path!0
		XOPER<X> functionHelp; // for function wizard
		std::vector<XOPER<X>> argumentHelp;

	public:
		XArg()
		{ }
		XArg(xcstr type, xcstr procedure, xcstr functionText)
			: typeText(type), procedure(procedure), functionText(functionText)
		{
			macroType = 1; // Function
		}
		XArg(xcstr procedure, xcstr functionText)
			: procedure(procedure), functionText(functionText)
		{
			macroType = 2; // Macro
		}

		// Used as key in add-in map.
		const XOPER<X>& FunctionText() const
		{
			return functionText;
		}

		struct Arg {
			xcstr type;
			xcstr name;
			xcstr help;
			//xcstr init;
		};

		XArg& Args(std::initializer_list<Arg> arg)
		{
			static xcstr comma{2, ',', ' ' };
			static X xComma = { .val = { .str = comma }, .xltype = xltypeStr };
			
			for (int i = 0; i < arg.size(); ++i) {
				typeText &= arg[i].type;
				if (i > 0) {
					argumentText &= xComma;
				}
				argumentText &= arg[i].name;
				argumentHelp[i] = arg[i].help;
			}
			
			return *this;
		}
		
		XArg& Category(xcstr category)
		{
			this->ategory = category;

			return *this;
		}
		XArg& FunctionHelp(xcstr functionHelp)
		{
			this->unctionHelp = functionHelp;

			return *this;
		}
		
		int Register()
		{
			int count = 10 + argumentHelp.size();
			const X* oper[255];

			if (moduleText.xltype == xltypeNil) {
				X x[1];
				traits<X>::Excelv(xlGetName, &x[0], 0, nullptr);
				moduleText = x[0];
				traits<X>::Excelv(xlFree, 0, 1, (X**)&x);
			}

			oper[0] = &moduleText;
			oper[1] = &procedure;    // C function
			oper[2] = &typeText;     // return type and arg codes 
			oper[3] = &functionText; // Excel function
			oper[4] = &argumentText; // Ctrl-Shift-A text
			oper[5] = &macroType;    // function; 2 for macro; 0 for hidden
			oper[6] = &category;     // for function wizard
			oper[7] = &shortcutText; // single character for Ctrl-Shift-char shortcut
			oper[8] = &helpTopic;    // filepath!HelpContextID or http://host/path!0
			oper[9] = &functionHelp; // for function wizard
			for (size_t i = 0; i < argumentHelp.size(); ++i) {
				oper[10 + i] = &argumentHelp[i];
			}

			X x;
			int ret = traits<X>::Excelv(xlfRegister, &x, count, const_cast<X**>(&oper[0]));
		
			return ret;
		}
	};

	using Function = XArg<XLOPER>;
	using Function12 = XArg<XLOPER12>;
	using FunctionX = XArg<XLOPERX>;

	using Macro = XArg<XLOPER>;
	using Macro12 = XArg<XLOPER12>;
	using MacroX = XArg<XLOPERX>;
}
