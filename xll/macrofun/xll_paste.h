// xll_paste.h - paste functions into spreadsheet
#pragma once
#include "xll_macrofun.h"

namespace xll {

	using xll::Excel;

	// Paste formula into reference and return a reference to what was pasted.
	// Argument is a string that gets evaluated.
	// String arguments must be quoted "\"a string\"".
	// Range arguments must start with left bracket "{1,2;3,4}".
	inline OPER paste_default(const OPER& x, Select ref = Select{})
	{
		ensure(x.is_str());

		if (x.val.str[0] == 0) {
			return ref; // missing
		}

		OPER xi = Excel(xlfEvaluate, x);
		if (xi) {
			ref.Reshape(xi);
			if (x.val.str[0] > 1 and x.val.str[1] == '=') {
				ref.Formula(x); // formula
			}
			else {
				ref.Set(xi);
			}
		}
		else {
			ref.Set(x);
		}

		return ref;
	}

	// paste add-in to active cell and defaults below
	inline OPER xll_paste_default_args(const Args& args)
	{
		Select xAct;

		// call with defaults arguments to get size of output
		OPER xVal = Excel(xlfEvaluate, args.ArgumentDefault(0));
		ensure(xVal);
		OPER xFor = OPER("=") & args.FunctionText() & OPER("(");
		Select sel;
		sel.Move(rows(xVal), 0);

		for (unsigned i = 1; i <= args.ArgumentCount(); ++i) {
			OPER xRel = paste_default(args.ArgumentDefault(i));
			if (i > 1) {
				xFor.append(", ");
			}
			xFor.append(xAct.Relref(xRel));
			sel.Move(rows(xRel), 0);
		}

		xFor.append(")");
		xAct.Formula(xFor);

		return xAct();
	}


} // namespace xll