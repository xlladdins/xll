// xll_paste.h - paste functions into spreadsheet
#pragma once
#include "xll_macrofun.h"

namespace xll {

	using xll::Excel;

	// Paste argument at active cell and return a reference to what was pasted.
	// Default arguments are always strings. 
	// Arguments of the form "{a,...;b,...}" get pasted as arrays.
	// Arguments that are strings must be quoted, "\"abc\"".
	inline OPER paste_default(const OPER& x)
	{
		ensure(x.is_str());
		OPER ref = Excel(xlfActiveCell);

		if (x.val.str[0] == 0) {
			return ref; // missing
		}

		OPER xi = Excel(xlfEvaluate, x);
		if (size(xi) == 1) {
			ensure(Excel(xlcFormula, xi, ref));
		}
		else {
			auto rw = rows(xi);
			auto col = columns(xi);
			ref = Excel(xlfOffset, ref, OPER(0), OPER(0), OPER(rw), OPER(col));
			ensure(Excel(xlcFormulaArray, xi, ref));
		}

		return ref;
	}
} // namespace xll