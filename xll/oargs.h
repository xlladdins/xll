// args.h - Arguments to register an Excel add-in
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <algorithm>
#include <string>
#include "excel.h"

namespace xll {

	/// See https://msdn.microsoft.com/en-us/library/office/bb687900.aspx
	enum ARG {
		ModuleText,   // from xlGetName
		Procedure,    // C function
		TypeText,     // return type and arg codes 
		FunctionText, // Excel function
		ArgumentText, // Ctrl-Shift-A text
		MacroType,    // function, 2 for macro, 0 for hidden
		Category,     // for function wizard
		ShortcutText, // single character for Ctrl-Shift-char shortcut
		HelpTopic,    // filepath!HelpContextID or http://host/path!0
		FunctionHelp, // for function wizard
		ArgumentHelp  // 1-based index...
	};

	using xcstr = const wchar_t*;

	/// <summary>Prepare an array suitible for <c>xlfRegister</c></summary>
	class Args {
		mutable OPER12 args;
        OPER12 ArgumentName_;
        OPER12 ArgumentDefault_;
        OPER12 Alias_; // alternate names
        std::wstring documentation;
        std::wstring remarks;
        std::wstring examples;
	public:
		/// Name of Excel add-in
		static OPER12& XlGetName()
		{
			static OPER12 hModule;

			if (hModule.type() != xltypeStr) {
				hModule = Excel(xlGetName);
			}

			return hModule;
		}
		
		/// <summary>Number of function arguments</summary>
		int Arity() const
		{
			return ArgumentName_.size();
		}
		OPER12 RegisterId() const
		{
			//return Excel(xlfEvaluate, Excel(xlfConcatenate, OPER12(L"="), args[ARG::FunctionText]));
            return Excel(xlfEvaluate, args[ARG::FunctionText]);
        }
		/// For use as Excelv(xlfRegister, Args(....))
		operator const OPER12&() const
		{
			return args;
		}

		/// Common default.
		Args()
			: args(1, ARG::ArgumentHelp1)
		{
			std::fill(args.begin(), args.end(), OPER12(xltype::Nil));
            args[ARG::MacroType] = OPER(-1); // default is document
		}
		Args(const Args&) = default;
		Args& operator=(const Args&) = default;
		~Args()
		{ }

		/// Macro
		Args(xcstr Procedure, xcstr FunctionText)
			: Args()
		{
			args[ARG::Procedure] = Procedure;
			Alias_ = args[ARG::FunctionText] = FunctionText;
			args[ARG::MacroType] = OPER12(2);
		}
		/// Function
		Args(xcstr TypeText, xcstr Procedure, xcstr FunctionText)
			: Args()
		{
			args[ARG::Procedure] = Procedure;
			args[ARG::TypeText] = TypeText;
            Alias_ = args[ARG::FunctionText] = FunctionText;
			args[ARG::MacroType] = OPER12(1);
		}
        /// Documentation
        Args(xcstr category)
            : Args()
        {
            // needed for Key()
            Alias_ = args[ARG::Category] = category;
            args[ARG::MacroType] = OPER(-1);
        }

		Args& ModuleText(const OPER& moduleText)
		{
			args[ARG::ModuleText] = moduleText;

			return *this;
		}
        const OPER& ModuleText() const
        {
            return args[ARG::ModuleText];
        }

		/// Set the name of the C/C++ function to be called.
		Args& Procedure(xcstr procedure)
		{
			args[ARG::Procedure] = procedure;

			return *this;
		}
        const OPER& Procedure() const
        {
            return args[ARG::Procedure];
        }

		/// Specify the return type and argument types of the function.
		Args& TypeText(xcstr typeText)
		{
			args[ARG::TypeText] = typeText;

			return *this;
		}
		/// Specify the name of the function or macro to be used by Excel.
		Args& FunctionText(xcstr functionText)
		{
			args[ARG::FunctionText] = functionText;

			return *this;
		}
		const OPER& FunctionText() const
		{
			return args[ARG::FunctionText];
		}
		/// Specify the macro type of the function.
		/// Use 1 for functions, 2 for macros, and 0 for hidden functions. 
		Args& MacroType(int macroType)
		{
			args[ARG::MacroType] = macroType;

			return *this;
		}
        bool isFunction() const
        {
            return args[ARG::MacroType] == 1;
        }
        bool isMacro() const
        {
            return args[ARG::MacroType] == 2;
        }
        bool isHidden() const
        {
            return args[ARG::MacroType] == 0;
        }
        bool isDocument() const
        {
            return args[ARG::MacroType] == -1; // special type
        }

        /// Hide the name of the function from Excel.
		Args& Hidden()
		{
			return MacroType(0);
		}
		/// Set the category to be used in the function wizard.
		Args& Category(xcstr category)
		{
			args[ARG::Category] = category;

			return *this;
		}
		const OPER& Category() const
		{
			return args[ARG::Category];
		}
		/// Specify the shortcut text for calling the function.
		Args& ShortcutText(wchar_t shortcutText)
		{
			args[ARG::ShortcutText] = OPER(&shortcutText, 1);

			return *this;
		}
		/// Specify the help topic to be used in the Function Wizard.
		/// This must have the format...
		Args& HelpTopic(xcstr helpTopic)
		{
			args[ARG::HelpTopic] = helpTopic;

			return *this;
		}
		/// Specify the function help displayed in the Functinon Wizard.
		Args& FunctionHelp(xcstr functionHelp)
		{
			args[ARG::FunctionHelp] = functionHelp;

			return *this;
		}
        const OPER& FunctionHelp() const
        {
            return args[ARG::FunctionHelp];
        }
        /// Specify individual argument help in the Function Wizard.
		Args& ArgumentHelp(COL i, xcstr argumentHelp)
		{
			ensure (i != 0);

			if (args.size() < ARG::ArgumentHelp + i)
				args.resize(1, ARG::ArgumentHelp + i);

			auto n = ARG::ArgumentHelp + i - 1;
			args[n] = argumentHelp;

			return *this;
		}
        const OPER& ArgumentHelp(int i) const
        {
            return args[ARG::ArgumentHelp + i - 1];
        }
        OPER ArgumentName(int i) const
        {
			ensure(i <= Arity());

            return ArgumentName_[i - 1];
        }
		OPER ArgumentDefault(int i) const
		{
			ensure(i <= Arity());

			return ArgumentDefault_[i - 1];
		}
		/// Add an individual argument.
		Args& Arg(xcstr type, xcstr name, xcstr helpText = nullptr, xcstr default_ = nullptr)
		{
			OPER& Type = args[ARG::TypeText];
			Type &= type;
			
			OPER& Text = args[ARG::ArgumentText];
			if (Text.isStr())
				Text &= L", ";
			Text &= name;

            ArgumentName_.push_back(OPER(name));
			
            RW n = Arity();
			if (helpText && *helpText) {
				ArgumentHelp(n, helpText);
            }
            if (default_ && *default_) {
                ArgumentDefault_.resize(n,1);
                ArgumentDefault_[n - 1] = default_;
            }

			return *this;
		}
		/// Argument modifiers
		Args& Threadsafe()
		{
			args[ARG::TypeText] &= XLL_THREAD_SAFE;

			return *this;
		}
		int isThreadsafe()
		{
			return Excel(xlfFind, args[ARG::TypeText], OPER(XLL_THREAD_SAFE)).isNum();
		}
		Args& Uncalced()
		{
			args[ARG::TypeText] &= XLL_UNCALCED;

			return *this;
		}
		int isUncalced()
		{
			return Excel(xlfFind, args[ARG::TypeText], OPER(XLL_UNCALCED)).isNum();
		}
		Args& Volatile()
		{
			args[ARG::TypeText] &= XLL_VOLATILE;

			return *this;
		}
		int isVolatile()
		{
			return Excel(xlfFind, args[ARG::TypeText], OPER(XLL_VOLATILE)).isNum();
		}

		/// Convenience function for number types.
        Args& Boolean(xcstr name, xcstr helpText = nullptr, xcstr default_ = nullptr)
        {
            return Arg(XLL_BOOL, name, helpText, default_);
        }
		Args& Number(xcstr name, xcstr helpText = nullptr, xcstr default_ = nullptr)
		{
			return Arg(XLL_DOUBLE, name, helpText, default_);
		}
        Args& Handle(xcstr name, xcstr helpText = nullptr, xcstr default_ = nullptr)
        {
            return Arg(XLL_DOUBLE, name, helpText, default_);
        }
        Args& UShort(xcstr name, xcstr helpText = nullptr, xcstr default_ = nullptr)
        {
            return Arg(XLL_USHORT, name, helpText, default_);
        }
        Args& Short(xcstr name, xcstr helpText = nullptr, xcstr default_ = nullptr)
        {
            return Arg(XLL_SHORT, name, helpText, default_);
        }
        Args& Word(xcstr name, xcstr helpText = nullptr, xcstr default_ = nullptr)
        {
            return Arg(XLL_WORD, name, helpText, default_);
        }
        Args& Long(xcstr name, xcstr helpText = nullptr, xcstr default_ = nullptr)
        {
            return Arg(XLL_LONG, name, helpText, default_);
        }

        // Convenience function for strings
        Args& String(xcstr name, xcstr helpText = nullptr, xcstr default_ = nullptr)
        {
            return Arg(XLL_CSTRING, name, helpText, default_);
        }
        Args& PString(xcstr name, xcstr helpText = nullptr, xcstr default_ = nullptr)
        {
            return Arg(XLL_PSTRING, name, helpText, default_);
        }

        // Two dimensional array of doubles.
        Args& Array(xcstr name, xcstr helpText = nullptr, xcstr default_ = nullptr)
        {
            return Arg(XLL_FP, name, helpText, default_);
        }

        // Convenience function for OPERS
        Args& Range(xcstr name, xcstr helpText = nullptr, xcstr default_ = nullptr)
        {
            return Arg(XLL_LPOPER, name, helpText, default_);
        }
        Args& Reference(xcstr name, xcstr helpText = nullptr, xcstr default_ = nullptr)
        {
            return Arg(XLL_LPXLOPER, name, helpText, default_);
        }

		Args& Documentation(const std::wstring& _documentation)
		{
            documentation = _documentation;

			return *this;
		}
        const std::wstring& Documentation() const
        {
            return documentation;
        }
        Args& Remarks(const std::wstring& _remarks)
        {
            remarks = _remarks;

            return *this;
        }
        const std::wstring& Remarks() const
        {
            return remarks;
        }
        Args& Examples(const std::wstring& _examples)
        {
            examples = _examples;

            return *this;
        }
        const std::wstring& Examples() const
        {
            return examples;
        }
        OPER Syntax() const
        {
            return FunctionText() & OPER(L"(") & args[ARG::ArgumentText] & OPER(L")");
        }

        const OPER& Key() const
        {
			return Alias_;
        }

	    // Simple hash function
        static int hash_string(const wchar_t* s, unsigned n)
        {
            static int A = 54059; /* a prime */
            static int B = 76963; /* another prime */
            static int C = 86969; /* yet another prime */
            static int FIRST = 37; /* also prime */
            int h = FIRST;
            while (n--) {
                h = (h * A) ^ (s[0] * B);
                s++;
            }

            return h > 0 ? h : -h;
        }

        // Integer hash used in help files.
        static OPER TopicId(const OPER& key)
        {
            if (!key.isStr()) {
                return OPER(L"");
            }

            int hash = hash_string(key.val.str + 1, key.val.str[0]);
            std::wstring Hash = std::to_wstring(hash);

            return OPER(Hash);
        }

        static OPER Guid(const OPER& key) 
        {
            if (!key.isStr()) {
                return OPER(L"");
            }

            int hash = hash_string(key.val.str + 1, key.val.str[0]);
            wchar_t buf[40];
            swprintf(buf, 40, L"%08X-0000-0000-0000-000000000000", hash);
            
            return OPER(buf);
        }

        Args& Alias(const std::wstring& alias)
        {
            Alias_.push_back(OPER(alias));

            return *this;
        }
        const OPER& Alias() const
        {
            return Alias_;
        }
        bool isAlias(const OPER& alias) const
        {
            return OPER(xlerr::NA) != Excel(xlfMatch, alias, Alias_, OPER(0)); // exact match
        }
        bool isAlias(const std::wstring& alias) const
        {
            return isAlias(OPER(alias));
        }

		/// Register an add-in function or macro
		OPER Register() const
		{
            OPER oResult;

            if (isDocument()) {
                return OPER(-1); // Do not register if document.
            }

            OPER name = XlGetName();
            args[ARG::ModuleText] = name;
            
            for (const OPER& key : Key()) {
                OPER ft = args[ARG::FunctionText];
                OPER chm = args[ARG::HelpTopic];
                args[ARG::FunctionText] = key;
                if (!documentation.empty()) {
                    OPER ht = Excel(xlfSubstitute, name, OPER(L".xll"), OPER(L".chm!"));
                    ht &= TopicId(key);
                    args[ARG::HelpTopic] = ht;
                }
                OPER oReg = Excelv(xlfRegister, args);
                args[ARG::HelpTopic] = chm;
                args[ARG::FunctionText] = ft;
                if (oReg.isErr()) {
				    OPER oError(L"Failed to register: ");
				    oError &= args[ARG::FunctionText];
				    oError &= L"/";
				    oError &= args[ARG::Procedure];
                    oError &= L"\nDid you forget to #pragma XLLEXPORT a function?";

                    MessageBoxW(0, oError.toStr().c_str(), L"Register Error", MB_OK);
                }
                else {
                    oResult.push_back(oReg);
                }
            }
 
			return oResult;
		}
        /// Unregister an add-in function or macro
        int Unregister() const
        {
            if (isDocument())
                return TRUE;

            return Excel(xlfUnregister, RegisterId()) == true;
        }
	};
	// semantic alias
	using Function = Args;
	using Macro = Args;
    using Documentation = Args;
    using Document = Args;
	// backwards compatibility
	using ArgsX = Args;
	using FunctionX = Args;
	using MacroX = Args;
	using Args12 = Args;
	using Function12 = Args;
	using Macro12 = Args;

	/// Array appropriate for xlfRegister.
	/// Use like <c>Excelv(xlfRegister, Arguments(...))</c>
	inline OPER Arguments(
		xcstr Procedure,        // C function
		xcstr TypeText,         // return type and arg codes 
		xcstr FunctionText,     // Excel function
		xcstr ArgumentText = 0, // Ctrl-Shift-A text
		int   MacroType = 1,    // function, 2 for macro, 0 for hidden
		xcstr Category = 0,     // for function wizard
		xcstr ShortcutText = 0, // single character for Ctrl-Shift-char shortcut
		xcstr HelpTopic = 0,    // filepath!HelpContextID or http://host/path!0
		xcstr FunctionHelp = 0, // for function wizard
		xcstr ArgumentHelp1 = 0,
		xcstr ArgumentHelp2 = 0,
		xcstr ArgumentHelp3 = 0,
		xcstr ArgumentHelp4 = 0,
		xcstr ArgumentHelp5 = 0,
		xcstr ArgumentHelp6 = 0,
		xcstr ArgumentHelp7 = 0,
		xcstr ArgumentHelp8 = 0,
		xcstr ArgumentHelp9 = 0
	)
	{
		OPER args(Args::XlGetName());
		args.push_back(OPER(Procedure));
		args.push_back(OPER(TypeText));
		args.push_back(OPER(FunctionText));
		args.push_back(OPER(ArgumentText));
		args.push_back(OPER(MacroType));
		args.push_back(OPER(Category));
		args.push_back(OPER(ShortcutText));
		args.push_back(OPER(HelpTopic));
		args.push_back(OPER(FunctionHelp));
		args.push_back(OPER(ArgumentHelp1));
		args.push_back(OPER(ArgumentHelp2));
		args.push_back(OPER(ArgumentHelp3));
		args.push_back(OPER(ArgumentHelp4));
		args.push_back(OPER(ArgumentHelp5));
		args.push_back(OPER(ArgumentHelp6));
		args.push_back(OPER(ArgumentHelp7));
		args.push_back(OPER(ArgumentHelp8));
		args.push_back(OPER(ArgumentHelp9));

		return args;
	}
	/*
	template<class R, class... Args>
	Args AutoRegister(R(*)(Args...), xcstr Procedure, xcstr FunctionText)
	{

	}
	*/

	/*
	// Convert __FUNCDNAME__ to arguments for xlfRegister
	inline Args Demangle(const wchar_t* F)
	{
		static std::map<wchar_t,const OPER12> arg_map = {
			{ L'F', OPER12(XLL_SHORT) },
			{ L'G', OPER12(XLL_WORD) }, // also USHORT
			{ L'H', OPER12(XLL_BOOL) },
			{ L'J', OPER12(XLL_LONG) },
			{ L'N', OPER12(XLL_DOUBLE) },
		};

		// C to Excel naming convention
		auto function_text = [](OPER12 o) {
			ensure (o.type() == xltypeStr);
			for (int i = 1; i <= o.val.str[0]; ++i) {
				if (o.val.str[i] == L'_')
					o.val.str[i] = L'.';
				else
					o.val.str[i] = ::towupper(o.val.str[i]);
			}

			return o;
		};

		Args args;
		args.MacroType(1); // function
								  
		// "?foo@@YGNN@Z"

		ensure (F && *F == '?');
		auto E = wcschr(F, L'@');
		ensure (E);
		args.Procedure(OPER12(F, E - F));
		args.FunctionText(function_text(OPER12(F + 1, E - F - 1)));

		F = E;
		ensure (*++F == L'@');
		ensure (*++F == L'Y');
		ensure (*++F == L'G');
		ensure (arg_map.find(*++F) != arg_map.end());
		args.TypeText(arg_map[*F]);
		while (*++F != L'@') {
			// if (*F == L'P') { pointer...
			ensure (arg_map.find(*F) != arg_map.end());
//			type = Excel(xlfConcatenate, type, OPER12(arg_map[*F]));
		}

		return args;
	}
	*/
} // xll
