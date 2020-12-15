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
		XOPER<X> moduleText;   // from xlGetName
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
			for (const auto& arg : args) {
				typeText &= arg.type;
				if (argumentHelp.size() != 0) {
					argumentText &= ", ";
				}
				argumentText &= arg.name;
				argumentName.push_back(XOPER<X>(arg.name));
				argumentHelp.push_back(XOPER<X>(arg.help));
				argumentDefault.push_back(XOPER<X>(arg.init));
			}

			return *this;
		}

		// list of typeText arguments
		XArgs& Args(const std::vector<cstr>& args)
		{
			for (const auto& arg : args) {
				typeText &= arg;
				if (argumentHelp.size() != 0) {
					argumentText &= ", ";
				}
				// generic text and help
				argumentText &= "Arg";
				argumentName.push_back(XOPER<X>("Arg")); // Arg<n>???
				argumentHelp.push_back(XOPER<X>("is an argument"));
				argumentDefault.push_back(XOPER<X>());
			}

			return *this;
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

		// Type of add-in: 1 for function, 2 for macro
		const XOPER<X>& MacroType() const
		{
			return macroType;
		}

		const XOPER<X>& FunctionHelp() const
		{
			return functionHelp;
		}

		const XOPER<X>& ArgumentText() const
		{
			return argumentText;
		}

		const std::vector<XOPER<X>>& ArgumentName() const
		{
			return argumentName;
		}

		const std::vector<XOPER<X>>& ArgumentHelp() const
		{
			return argumentHelp;
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
			helpTopic = _helpTopic;

			return *this;
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

		// Register add-in with Excel
		X Register()
		{
			int count = 10 + static_cast<int>(argumentHelp.size());
			std::vector<X*> oper(count + 1);

			if (moduleText == XOPER<X>{}) {
				moduleText = XExcel<X>(xlGetName);
			}

			// C++ mangled name must start with '?'
			ensure(procedure.xltype == xltypeStr && procedure.val.str[0] > 1);
			if (procedure.val.str[1] != '?' && procedure.val.str[1] != '_') {
				procedure = XOPER<X>("?") & procedure;
			}

			// Help URLs must end with "!0"
			if (helpTopic.xltype == xltypeStr) {
				XOPER xFind = XExcel<X>(xlfFind, XOPER<X>("!"), helpTopic);
				if (xFind.xltype == xltypeErr && xFind.val.err == xlerrValue) {
					helpTopic &= XOPER<X>("!0");
				}
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
			// https://docs.microsoft.com/en-us/office/client-developer/excel/known-issues-in-excel-xll-development#argument-description-string-truncation-in-the-function-wizard
			XOPER<X> xEmpty("");
			oper[count] = &xEmpty;

			X registerId = { .xltype = xltypeNil };
			int ret = traits<X>::Excelv(xlfRegister, &registerId, count + 1, &oper[0]);
			if (ret != xlretSuccess || registerId.xltype != xltypeNum) {
				XOPER<X> xMsg("Failed to register: ");
				XExcel<X>(xlcAlert, xMsg & functionText, XOPER<X>(2));
			}

			return registerId;
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
