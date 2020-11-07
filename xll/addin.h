// addin.h - convenience wrapper for Excel add-ins
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <map>
#include "args.h"

namespace xll {

    /// <summary>
    /// Store argument add-in data for xlfRegister
    /// </summary>
    template<class X>
	struct XAddIn {
        // add-ins indexed by Excel name
        static inline std::map<XOPER<X>, XArgs<X>> Map;
 
        XAddIn(const XArgs<X>& args) noexcept
        {
            const auto& key = args.FunctionText();
            auto [_, inserted] = Map.try_emplace(key, std::move(args));
            /*
            // warn if it already exists
            if (!inserted) {
                std::basic_string<TCHAR> msg{ TEXT("AddIn previously defined: ") };
                msg.append(key.val.str + 1, key.val.str[0]);
                MessageBox(GetForegroundWindow(), msg.c_str(), 0, MB_OK);
            }
            */
        }

        // Get arguments using Excel function text name.
        static XArgs<X>& Args(const char* name)
        {
            return Args(XOPER<X>(name));
        }
        static XArgs<X>& Args(const XOPER<X>& name)
        {
            auto i = Map.find(name);
            ensure(i != Map.end());

            return *i;
        }
    };

	using AddIn4  = XAddIn<XLOPER>;
    using AddIn12 = XAddIn<XLOPER12>;
    using AddIn   = XAddIn<XLOPERX>;

} // xll namespace


