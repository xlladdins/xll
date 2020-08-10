// args.h - Arguments for xlfRegister
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <initializer_list>
#include <vector>
#include "oper.h"

namespace xll {

	/// <summary>
	/// Individual argument for add-in function.
	/// </summary>
	struct Arg {
		typedef typename const char* cstr;

		cstr type;
		cstr name;
		cstr help;
		//cstr init;
		Arg(cstr type, cstr name, cstr help)
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

	template<class X>
	class XArgs {
		using cstr = const char*;

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
		XArgs(cstr type, cstr procedure, cstr functionText)
			: typeText(type), procedure(procedure), functionText(functionText), macroType(1)
		{
		}
		XArgs(cstr procedure, cstr functionText)
			: procedure(procedure), functionText(functionText), macroType(2)
		{
		}

		// Used as key in add-in map.
		const XOPER<X>& FunctionText() const
		{
			return functionText;
		}

		XArgs& Args(std::initializer_list<Arg> args)
		{
			for (const auto& arg : args) {
				typeText &= arg.type;
				if (argumentHelp.size() != 0) {
					argumentText &= ", ";
				}
				argumentText &= arg.name;
				argumentHelp.push_back(XOPER<X>(arg.help));
			}
			
			return *this;
		}
		
		XArgs& Category(cstr _category)
		{
			category = _category;

			return *this;
		}
		XArgs& FunctionHelp(cstr _functionHelp)
		{
			functionHelp = _functionHelp;

			return *this;
		}
		XArgs& HelpTopic(cstr _helpTopic)
		{
			// !!! If it does not end with '!.+` add a '!0'.
			helpTopic = _helpTopic;

			return *this;
		}
		XArgs& Uncalced()
		{
			typeText &= XLL_UNCALCED; //!!! use traits

			return *this;
		}
		XArgs& Volatile()
		{
			typeText &= XLL_VOLATILE;

			return *this;
		}

		// slice okay since it is xltypeNum/Err
		X RegisterId() const
		{
			return Excel<X>(xlfEvaluate, functionText);
		}

		// Register add-in with Excel
		X* Register()
		{
			size_t count = 10 + argumentHelp.size();
			std::vector<X*> oper(count);

			if (moduleText == XOPER<X>{}) {
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
			if (ret != xlretSuccess || registerId.xltype != xltypeNum) {
				Excel<X>(xlcAlert, functionText, XOPER<X>(2));
			}

			return &registerId;
		}
		// Unregister add-in.
		XOPER<X> Unregister() const
		{
			X* px[1];
			px[0] = &registerId;
			return traits<X>::Excelv(xlfUnregister, 0, 1, px);
		}

		// full name of dll
		const XOPER<X>& ModuleText() const
		{
			if (moduleText == XOPER<X>{}) {
				moduleText = Excel<X>(xlGetName);
			}
			
			return moduleText;
		}
	};

	using Function4 = XArgs<XLOPER>;
	using Function12 = XArgs<XLOPER12>;
	using Function = XArgs<XLOPERX>;

	using Macro4 = XArgs<XLOPER>;
	using Macro12 = XArgs<XLOPER12>;
	using Macro = XArgs<XLOPERX>;

	/*
	template<typename X>
	struct Enum : public XArgs<X> {
		using xcstr = typename traits<X>::xcstr;
		Enum(xcstr name, xcstr value, xcstr category, xcstr description, xcstr doc = 0)
			: ???
	};
	*/
}
