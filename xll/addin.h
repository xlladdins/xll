// addin.h - convenience wrapper for Excel add-ins
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <map>
#include "args.h"

namespace xll {

    /// <summary>
    /// Store add-in data for xlfRegister
    /// </summary>
    template<class X>
	struct XAddIn {
        static inline std::map<XOPER<X>, XArgs<X>> Map;
        XAddIn(const XArgs<X>& args)
        {
            Map[args.FunctionText()] = args;
        }
        // Get arguments using Excel function text name.
        static XArgs<X>& Args(typename traits<X>::xcstr name)
        {
            return Map[XOPER<X>(name)];
        }
        static XArgs<X>& Args(const XOPER<X>& name)
        {
            return Map[name];
        }
        /*
        static XArgs<X>& get(const XOPER<X>& key)
        {
            auto i = Map.find(key);
            ensure(i != Map.end());

            return *i;
        }
        static const XArgs<X>& get(const XOPER<X>& key)
        {
            auto i = Map.find(key);
            ensure(i != Map.end());

            return *i;
        }
        */
    };

	using AddIn = XAddIn<XLOPER>;
    using AddIn12 = XAddIn<XLOPER12>;
    using AddInX = XAddIn<XLOPERX>;

} // xll namespace


