// cdecl.h - print C++ declaration to output stream
#pragma once

#include <iostream>
#include "args.h"

namespace xll {

	inline std::ostream cdecl(std::ostream& os, const Args& args)
	{
		os << " WINAPI ";
		//...
		os << ");\n";

		return os;
	}

}
