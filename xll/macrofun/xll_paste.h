// xll_paste.h - paste functions into spreadsheet
#pragma once
#include "xll_macrofun.h"

namespace xll {

	using xll::Excel;

	// Paste formula into active cell and return a reference to what was pasted.
	// Argument is a string that get evaluated.
	// String arguments must be quoted, "\"abc\"".
	// Array arguments have the form "={a,...;b,...}".
	// Formula arguments must start with an equal sign "=1 + rand()".
	inline OPER paste_formula(const OPER& x, Select ref = Select{})
	{
		ensure(x.is_str());

		if (x.val.str[0] == 0) {
			return ref; // missing
		}

		OPER xi = Excel(xlfEvaluate, x);
		ref.Offset(0, 0, xi.rows(), xi.columns());
		ref.Formula(x);

		return ref;
	}

	inline OPER paste_default(const Args& args, unsigned i)
	{
		return paste_formula(args.ArgumentDefault(i));
	}

} // namespace xll