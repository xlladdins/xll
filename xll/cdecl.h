// cdecl.h - print C++ declaration to output stream
#pragma once

#include <iostream>
#include "args.h"

namespace xll {

	template<class X, class T = traits<X>::xchar> requires std::is_same_v<XLOPER,X> || std::is_same_v<XLOPER12,X>
	inline std::ostream cdecl(std::basic_ostream<T>& os, const XArgs<X>& args)
	{
		XOPER<X> sig = args.typeText;
		os << " WINAPI ";
		//...
		os << ");\n";

		return os;
	}

}
