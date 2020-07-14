// args.h - Arguments for xlfRegister
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.

#pragma once
#include <initializer_list>
#include <vector>
#include "excel.h"

namespace xll {

	/// <summary>
	/// Individual argument for add-in function.
	/// Use ArgX({type, name, help}) to construct.
	/// </summary>
	template<class X>
	struct XArg {
		using xchar = typename traits<X>::xchar;
		using xcstr = typename traits<X>::xcstr;

		xcstr type;
		xcstr name;
		xcstr help;
		//xcstr init;
		XArg(xcstr type, xcstr name, xcstr help)
			: type(type), name(name), help(help)
		{
			/*
			if (!xll_arg_types.contains(type)) {
				std::basic_string<xchar> msg{ _T("Unknown Excel argument type: ") };
				msg += type;
				MessageBox(GetForegroundWindow(), msg.c_str(), 0, MB_OK);
			}
			*/
		}
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
			typeText &= XLL_UNCALCEDX; //!!! traits

			return *this;
		}
		XArgs& Volatile()
		{
			typeText &= XLL_VOLATILEX;

			return *this;
		}

		// slice okay since it is xltypeNum
		X RegisterId() const
		{
			return Excel<X>(xlfEvaluate, functionText);
		}

		// Register add-in with Excel
		int Register()
		{
			size_t count = 10 + argumentHelp.size();
			std::vector<X*> oper(count);

			if (moduleText.xltype == xltypeNil) {
				moduleText = Excel<X>(xlGetName);
			}

			oper[0] = &moduleText;
			oper[1] = &procedure;
			oper[2] = &typeText; 
			oper[3] = &functionText;
			oper[4] = &argumentText;
			oper[5] = &macroType;
			oper[6] = &category;
			oper[7] = &shortcutText;
			oper[8] = &helpTopic;
			oper[9] = &functionHelp;
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
		// Unregister add-in.
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
