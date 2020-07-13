// addin.h - convenience wrapper for Excel add-ins
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <map>
#include "args.h"

namespace xll {

    /// <summary>
    /// Store add-in data for xlAutoOpen/Close
    /// </summary>
    template<class X>
	class XAddIn {
    public:
        static inline std::map<XOPER<X>, XArgs<X>> Map;

        XAddIn(const XArgs<X>& args)
        {
            Map[args.FunctionText()] = args;
        }
    };

	using AddIn = XAddIn<XLOPER>;
    using AddIn12 = XAddIn<XLOPER12>;
    using AddInX = XAddIn<XLOPERX>;

} // xll namespace


