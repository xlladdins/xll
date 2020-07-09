// addin.h - convenience wrapper for Excel add-ins
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <cwctype>
#include <map>
#include "auto.h"
#include "arg.h"

namespace xll {

	/// Manage the lifecycle of an Excel add-in.
    template<class X>
	class XAddIn {
    public:
        static inline std::map<XOPER<X>, XArgs<X>> Map;

        XAddIn(const XArgs<X>& args)
        {
            Map[args.FunctionText()] = args;
        }
    };

	using AddIn12 = XAddIn<XLOPER12>;
	using AddIn = XAddIn<XLOPER>;
    using AddInX = XAddIn<XLOPERX>;

} // xll namespace


