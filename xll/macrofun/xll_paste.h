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
	inline OPER paste_formula(const OPER& x)
	{
		ensure(x.is_str());
		OPER ref = Excel(xlfActiveCell);

		if (x.val.str[0] == 0) {
			return ref; // missing
		}

		OPER xi = Excel(xlfEvaluate, x);
		if (size(xi) == 1) {
			ensure(Excel(xlcFormula, x, ref));
		}
		else {
			auto rw = rows(xi);
			auto col = columns(xi);
			ref = Excel(xlfOffset, ref, OPER(0), OPER(0), OPER(rw), OPER(col));
			ensure(Excel(xlcFormulaArray, x, ref));
		}

		return ref;
	}

	inline OPER paste_default(const Args& args, unsigned i)
	{
		return paste_formula(args.ArgumentDefault(i));
	}

} // namespace xll