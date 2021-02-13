// args.h - Arguments for xlfRegister
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <initializer_list>
#include <map>
#include <set>
#include <vector>
#include "excel.h"

// allows add-in argument types
#define XLL_ARG_TYPES(a,b,c,d) b,c,
inline std::set<std::string> xll_arg_types = {
XLL_ARG_TYPE(XLL_ARG_TYPES)
};
#undef XLL_ARG_TYPES

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
			if (!xll_arg_types.contains(type)) {
				std::string err("Unknown add-in argument type: ");
				XLL_ERROR(err.append(type).c_str());
			}
		}
	};

	using xll::Excel;

	/// <summary>
	/// Everything needed to register an add-in.
	/// </summary>
	class Args {
		using cstr = const char*;

		std::map<OPER, OPER> argMap;
		std::string documentation;
	public:
		OPER& operator[](const OPER& name)
		{
			return argMap[name];
		}
		const OPER& operator[](const OPER& name) const
		{
			static OPER o = Nil;
			auto i = argMap.find(name);
			return i != argMap.end() ? i->second : o;
		}
		OPER& key(cstr name)
		{
			return operator[](OPER(name));
		}
		const OPER& key(cstr name) const
		{
			return operator[](OPER(name));
		}

		Args()
		{ }
		Args(const Args&) = default;
		Args& operator=(const Args&) = default;
		Args(Args&&) = default;
		Args& operator=(Args&&) = default;
		~Args()
		{ }
		// Function
		Args(cstr type, cstr procedure, cstr functionText)
		{
			key("typeText") = type;
			key("procedure") = procedure;
			key("functionText") = functionText;
			key("macroType") = 1;
		}
		// Macro
		Args(cstr procedure, cstr functionText)
		{
			key("procedure") = procedure;
			key("functionText") = functionText;
			key("macroType") = 2;
		}

		// list of function arguments
		Args& Arguments(const std::initializer_list<Arg>& args)
		{
			const char* comma = "";
			for (const auto& arg : args) {
				key("typeText") &= arg.type;
				key("argumentName").push_back(OPER(arg.name));
				key("argumentHelp").push_back(OPER(arg.help));
				key("argumentDefault").push_back(OPER(arg.init));
				key("argumentText") &= comma;
				key("argumentText") &= arg.name;
				comma = ", ";
			}

			return *this;
		}
		/*
		// list of typeText arguments
		Args& Arguments(const std::vector<cstr>& args)
		{
			return Arguments(std::initializer_list<Arg>(args.begin(), args.end()));
		}
		*/

		Args& Uncalced()
		{
			key("typeText") &= XLL_UNCALCED;

			return *this;
		}
		bool isUncalced() const
		{
			return !Excel(xlfFind, OPER(XLL_UNCALCED), TypeText()).is_err();
		}

		Args& Volatile()
		{
			key("typeText") &= XLL_VOLATILE;

			return *this;
		}
		bool isVolatile() const
		{
			return !Excel(xlfFind, OPER(XLL_VOLATILE), TypeText()).is_err();
		}

		Args& ThreadSafe()
		{
			key("typeText") &= XLL_THREAD_SAFE;

			return *this;
		}
		bool isThreadSafe() const
		{
			return !Excel(xlfFind, OPER(XLL_THREAD_SAFE), TypeText()).is_err();
		}

		Args& ClusterSafe()
		{
			key("typeText") &= XLL_CLUSTER_SAFE;

			return *this;
		}
		bool isClusterSafe() const
		{
			return !Excel(xlfFind, OPER(XLL_CLUSTER_SAFE), TypeText()).is_err();
		}
		Args& Asynchronous()
		{
			key("typeText") &= XLL_ASYNCHRONOUS;

			return *this;
		}
		bool isAsynchronous() const
		{
			return !Excel(xlfFind, OPER(XLL_ASYNCHRONOUS), TypeText()).is_err();
		}

		OPER RegisterId() const
		{
			return Excel(xlfEvaluate, key("functionText"));
		}

		const OPER& ModuleText() const
		{
			return key("moduleText");
		}
		OPER& ModuleText()
		{
			return key("moduleText");
		}

		// C++ function name
		const OPER& Procedure() const
		{
			return key("procedure");
		}

		// Excel SDK signature
		const OPER& TypeText() const
		{
			return key("typeText");
		}

		// Excel function name used as key for add-in map.
		const OPER& FunctionText() const
		{
			return key("functionText");
		}
		// Key used in AddIn::Map.
		const OPER& Key() const
		{
			return key("functionText");
		}

		const OPER& ArgumentText() const
		{
			return key("argumentText");
		}

		// Type of add-in: 1 for function, 2 for macro, 3 for hidden
		const OPER& MacroType() const
		{
			return key("macroType");
		}
		OPER Type() const
		{
			return OPER(MacroType() == 1 ? "function"
				: MacroType() == 2 ? "macro"
				: MacroType() == 0 ? "hidden"
				: "unknown");
		}
		bool isFunction() const
		{
			return MacroType() == 1;
		}
		bool isMacro() const
		{
			return MacroType() == 2;
		}
		bool isHidden() const
		{
			return MacroType() == 0;
		}

		const OPER& Category() const
		{
			return key("category");
		}
		Args& Category(cstr category)
		{
			key("category") = category;

			return *this;
		}

		const OPER& ShortcutText() const
		{
			return key("shortcutText");
		}
		Args& ShortcutText(cstr shortcutText)
		{
			key("shortcutText") = shortcutText;

			return *this;
		}

		const OPER& HelpTopic() const
		{
			return key("helpTopic");
		}
		Args& HelpTopic(cstr helpTopic)
		{
			key("helpTopic") = helpTopic;

			return *this;
		}

		const OPER& FunctionHelp() const
		{
			return key("functionHelp");
		}
		Args& FunctionHelp(cstr functionHelp)
		{
			key("functionHelp") = functionHelp;

			return *this;
		}

		// 1-based indexing
		unsigned ArgumentCount() const
		{
			return static_cast<unsigned>(ArgumentName().size());
		}

		const OPER& ArgumentName() const
		{
			return key("argumentName");
		}
		const OPER& ArgumentName(unsigned i) const
		{
			if (i == 0) {
				return key("functionText");
			}

			return key("argumentName")[i - 1];
		}

		const OPER& ArgumentHelp() const
		{
			return key("argumentHelp");
		}
		const OPER& ArgumentHelp(unsigned i) const
		{
			if (i == 0) {
				return key("functionHelp");
			}

			return key("argumentHelp")[i - 1];
		}


		// Default value for argument.
		const OPER& ArgumentDefault(unsigned i) const
		{
			if (i == 0) {
				// function call with default arguments
				static OPER arg0, eq("="), lp("("), rp(")"), c(", ");

				arg0 = eq & FunctionText() & lp;
				for (unsigned j = 1; j <= ArgumentCount(); ++j) {
					if (j > 1) {
						arg0 &= c;
					}
					arg0 &= ArgumentDefault(j);
				}
				arg0 &= rp;

				return arg0;
			}

			return key("argumentDefault")[i - 1];
		}

		Args& Documentation(const std::string& s)
		{
			documentation = s;

			return *this;
		}
		const std::string& Documentation() const
		{
			return documentation;
		}

		// never needed
		auto Unregister() const
		{
			XLOPERX* px[1];
			OPER regid = RegisterId();
			px[0] = &regid;
		
			return traits<XLOPERX>::Excelv(xlfUnregister, 0, 1, px);
		}
	};

	using Function = Args;
	using Macro = Args;

}
