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
		cstr init;
		Arg(cstr type, cstr name, cstr help, cstr init = nullptr)
			: type(type), name(name), help(help), init(init)
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

	/// <summary>
	/// Everything needed to register an add-in.
	/// </summary>
	/// <typeparam name="X">Must be either XLOPER or XLOPER12</typeparam>
	template<class X>
		requires (std::is_same_v<XLOPER, X> || std::is_same_v<XLOPER12, X>)
	class XArgs {
		XOPER<X> procedure;    // C function
		XOPER<X> typeText;     // return type and arg codes 
		XOPER<X> functionText; // Excel function
		XOPER<X> argumentText; // Ctrl-Shift-A text
		XOPER<X> macroType;    // function; 2 for macro; 0 for hidden
		XOPER<X> category;     // for function wizard
		XOPER<X> shortcutText; // single character for Ctrl-Shift-char shortcut
		XOPER<X> helpTopic;    // filepath!HelpContextID or http://host/path!0, ???default to???
		XOPER<X> functionHelp; // for function wizard
		std::vector<XOPER<X>> argumentName;
		std::vector<XOPER<X>> argumentHelp;
		std::vector<XOPER<X>> argumentDefault;
		std::string documentation;
//		X registerId = { .val = { .num = 0 }, .xltype = xltypeNum };

		using cstr = const char*;
	public:
		XArgs()
		{ }
		XArgs(const XArgs&) = default;
		XArgs& operator=(const XArgs&) = default;
		~XArgs()
		{ }
		// Function
		XArgs(cstr type, cstr procedure, cstr functionText)
			: typeText(type), procedure(procedure), functionText(functionText), macroType(1)
		{
		}
		// Macro
		XArgs(cstr procedure, cstr functionText)
			: procedure(procedure), functionText(functionText), macroType(2)
		{
		}

		// list of function arguments
		XArgs& Args(const std::initializer_list<Arg>& args)
		{
			const char* comma = "";
			for (const auto& arg : args) {
				typeText &= arg.type;
				argumentName.push_back(XOPER<X>(arg.name));
				argumentHelp.push_back(XOPER<X>(arg.help));
				argumentDefault.push_back(XOPER<X>(arg.init));
				argumentText &= comma;
				argumentText &= arg.name;
				comma = ", ";
			}

			return *this;
		}

		// list of typeText arguments
		XArgs& Args(const std::vector<cstr>& args)
		{
			return std::initializer_list<Arg>(args.begin(), args.end());
		}

		XArgs& Uncalced()
		{
			typeText &= XLL_UNCALCED;

			return *this;
		}
		XArgs& Volatile()
		{
			typeText &= XLL_VOLATILE;

			return *this;
		}
		XArgs& ThreadSafe()
		{
			typeText &= XLL_THREAD_SAFE;

			return *this;
		}
		XArgs& ClusterSafe()
		{
			typeText &= XLL_CLUSTER_SAFE;

			return *this;
		}
		XArgs& Asynchronous()
		{
			typeText &= XLL_ASYNCHRONOUS;

			return *this;
		}

		// slice okay since it is xltypeNum/Err
		X RegisterId() const
		{
			return XExcel<X>(xlfEvaluate, functionText);
		}

		// C++ function name
		const XOPER<X>& Procedure() const
		{
			return procedure;
		}

		// Excel SDK signature
		const XOPER<X>& TypeText() const
		{
			return typeText;
		}

		// Excel function name used as key for add-in map.
		const XOPER<X>& FunctionText() const
		{
			return functionText;
		}
		// Key used in AddIn::Map.
		const XOPER<X>& Key() const
		{
			return functionText;
		}

		const XOPER<X>& ArgumentText() const
		{
			return argumentText;
		}

		// Type of add-in: 1 for function, 2 for macro, 3 for hidden
		const XOPER<X>& MacroType() const
		{
			return macroType;
		}
		XOPER<X> Type() const
		{
			return OPER(macroType == 1 ? "function"
				: macroType == 2 ? "macro"
				: macroType == 0 ? "hidden"
				: "unknown");
		}
		bool isFunction() const
		{
			return macroType == 1;
		}
		bool isMacro() const
		{
			return macroType == 2;
		}
		bool isHidden() const
		{
			return macroType == 0;
		}

		const XOPER<X>& Category() const
		{
			return category;
		}
		XArgs& Category(cstr _category)
		{
			category = _category;

			return *this;
		}

		const XOPER<X>& ShortcutText() const
		{
			return shortcutText;
		}
		XArgs& ShortcutText(cstr _shortcutText)
		{
			shortcutText = _shortcutText;

			return *this;
		}

		const XOPER<X>& HelpTopic() const
		{
			return helpTopic;
		}
		XArgs& HelpTopic(cstr _helpTopic)
		{
			helpTopic = _helpTopic;

			return *this;
		}

		const XOPER<X>& FunctionHelp() const
		{
			return functionHelp;
		}
		XArgs& FunctionHelp(cstr _functionHelp)
		{
			functionHelp = _functionHelp;

			return *this;
		}

		size_t ArgumentCount() const
		{
			return argumentName.size();
		}

		const std::vector<XOPER<X>>& ArgumentName() const
		{
			return argumentName;
		}
		const XOPER<X>& ArgumentName(size_t i) const
		{
			return argumentName[i];
		}

		const std::vector<XOPER<X>>& ArgumentHelp() const
		{
			return argumentHelp;
		}
		const XOPER<X>& ArgumentHelp(size_t i) const
		{
			return argumentHelp[i];
		}


		// Default value for argument.
		const X& ArgumentDefault(size_t i) const
		{
			return argumentDefault[i];
		}

		XArgs& Documentation(const std::string& s)
		{
			documentation = s;

			return *this;
		}
		const std::string& Documentation() const
		{
			return documentation;
		}

		// never needed
		XOPER<X> Unregister() const
		{
			X* px[1];
			px[0] = RegisterId();
			return traits<X>::Excelv(xlfUnregister, 0, 1, px);
		}
	};

	using Args4  = XArgs<XLOPER>;
	using Args12 = XArgs<XLOPER12>;
	using Args   = XArgs<XLOPERX>;

	using Function4  = XArgs<XLOPER>;
	using Function12 = XArgs<XLOPER12>;
	using Function   = XArgs<XLOPERX>;

	using Macro4  = XArgs<XLOPER>;
	using Macro12 = XArgs<XLOPER12>;
	using Macro   = XArgs<XLOPERX>;

}
