// auto.h - export xlAutoXXX functions
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <algorithm>
#include <functional>
#include <vector>

// Use Auto<XXX> xao_foo(xll_foo) to run xll_foo on event XXX
namespace xll {

	class Open {};
	class OpenAfter {};
	class CloseBefore {};
	class Close {};
	class Add {};
	class Remove {};

	// Register macros to be called in xlAuto functions.
	template<class T>
	class Auto {
		using macro = std::function<int(void)>;
		inline static std::vector<macro> macros;
	public:
		Auto(const macro& m)
		{
			macros.emplace_back(m); // run in the order they are constructed
		}
		static int Call(void)
		{
			for (const auto& m : macros) {
				if (!m()) {
					return FALSE;
				}
			}

			return TRUE;
		}
	};
} // namespace xll
