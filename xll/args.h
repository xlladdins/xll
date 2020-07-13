// arg.h - Argument for xlfRegister
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.

#pragma once
#include <initializer_list>
#include <vector>
#include "excel.h"

namespace xll {

	template<class X>
	struct XArg {
		using xcstr = typename traits<X>::xcstr;

		xcstr type;
		xcstr name;
		xcstr help;
		//xcstr init;
	};
	using Arg = XArg<XLOPER>;
	using Arg12 = XArg<XLOPER12>;
	using ArgX = XArg<XLOPERX>;


	template<class X>
	class XArgs {
		using xchar = typename traits<X>::xchar;
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
		X registerId = { .val = { .num = 0 }, .xltype = xltypeNum };

	public:
		XArgs()
		{ }
		XArgs(xcstr type, xcstr procedure, xcstr functionText)
			: typeText(type), procedure(procedure), functionText(functionText)
		{
			macroType = 1; // Function
		}
		XArgs(xcstr procedure, xcstr functionText)
			: procedure(procedure), functionText(functionText)
		{
			macroType = 2; // Macro
		}

		// Used as key in add-in map.
		const XOPER<X>& FunctionText() const
		{
			return functionText;
		}

		XArgs& Args(std::initializer_list<XArg<X>> args)
		{
			static xchar comma[3] = {2, ',', ' ' };
			static X xComma = { .val = { .str = comma }, .xltype = xltypeStr };

			for (const auto& arg : args) {
				typeText &= arg.type;
				if (argumentHelp.size() != 0) {
					argumentText &= xComma;
				}
				argumentText &= arg.name;
				argumentHelp.push_back(XOPER<X>(arg.help));
			}
			
			return *this;
		}
		
		XArgs& Category(xcstr _category)
		{
			category = _category;

			return *this;
		}
		XArgs& FunctionHelp(xcstr _functionHelp)
		{
			functionHelp = _functionHelp;

			return *this;
		}
		XArgs& Uncalced()
		{
			typeText &= XLL_XUNCALCED<X>::value;

			return *this;
		}
		XArgs& Volatile()
		{
			typeText &= XLL_XVOLATILE<X>::value;

			return *this;
		}

		// slice since it is xltypeNum
		X RegisterId() const
		{
			return Excel<X>(xlfEvaluate, functionText);
		}

		int Register()
		{
			size_t count = 10 + argumentHelp.size();
			const X* oper[255];

			if (moduleText.xltype == xltypeNil) {
				moduleText = Excel<X>(xlGetName);
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

			int ret = traits<X>::Excelv(xlfRegister, &registerId, static_cast<int>(count), const_cast<X**>(&oper[0]));
			if (registerId.xltype != xltypeNum) {
				Excel<X>(xlcAlert, functionText, XOPER<X>(3)); // error

				return xlretFailed;
			}

			return ret;
		}
		XOPER<X> Unregister() const
		{
			return Excel<X>(xlfUnregister, registerId);
		}
	};

	using Function = XArgs<XLOPER>;
	using Function12 = XArgs<XLOPER12>;
	using FunctionX = XArgs<XLOPERX>;

	using Macro = XArgs<XLOPER>;
	using Macro12 = XArgs<XLOPER12>;
	using MacroX = XArgs<XLOPERX>;

	/*
	template<typename X>
	struct Enum : public XArgs<X> {
		using xcstr = typename traits<X>::xcstr;
		Enum(xcstr name, xcstr value, xcstr category, xcstr description, xcstr doc = 0)
			: ???
	};
	*/
}
