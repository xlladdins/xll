// auto.h - export xlAutoXXX functions
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <algorithm>
#include <functional>
#include <vector>
#include "defines.h"

// Export well known functions implemented by xll lib.
#pragma comment(linker, "/include:" XLL_DECORATE("DllMain", 12))
//#pragma comment(linker, "/export:" XLL_DECORATE("XLCallVer", 0))
#pragma comment(linker, "/export:xlAutoOpen" XLL_X32("@0=xlAutoOpen"))
#pragma comment(linker, "/export:xlAutoClose" XLL_X32("@0=xlAutoClose"))
#pragma comment(linker, "/export:xlAutoAdd" XLL_X32("@0=xlAutoAdd"))
#pragma comment(linker, "/export:xlAutoRemove" XLL_X32("@0=xlAutoRemove"))
#pragma comment(linker, "/export:xlAutoFree12" XLL_X32("@4=xlAutoFree12"))
#pragma comment(linker, "/export:xlAutoRegister12" XLL_X32("@4=xlAutoRegister12"))
//#pragma comment(linker, "/export:xlAddInManagerInfo12@4=xlAddInManagerInfo12")
#ifdef _DEBUG
#pragma comment(linker, "/export:?crtDbg@@3UCrtDbg@@A")
#endif

// Use Auto<XXX> xao_foo(xll_foo) to run xll_foo on event XXX
namespace xll {

    class OpenBefore {};
	class Open {};
    class OpenAfter {};
	class Close {};
	class Add {};
	class Remove {};

	// Register macros to be called in xlAuto functions.
	template<class T>
	class Auto {
	public:
		using macro = std::function<int(void)>;
		Auto(const macro& m)
		{
			macros().push_back(m); // run in the order they are constructed
		}
		static int Call(void)
		{
            for (const auto& m : macros()) {
                if (!m()) {
                    return FALSE;
                }
            }

            return TRUE;
		}
	private:
		static std::vector<macro>& macros(void)
		{
			static std::vector<macro> macros_;

			return macros_;
		}
	};
} // namespace xll
