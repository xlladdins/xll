// xll_paste.h - paste functions into spreadsheet
#pragma once
#include "xll_macrofun.h"

namespace xll {

	using xll::Excel;

	// Paste formula into active cell and return a reference to what was pasted.
	// Argument is a string that gets evaluated.
	// Range arguments must start with left bracket "{1,2;3,4}"
	// Formula arguments must start with an equal sign "=1 + rand()".
	inline OPER paste_formula(const OPER& x, Select ref = Select{})
	{
		ensure(x.is_str());

		if (x.val.str[0] == 0) {
			return ref; // missing
		}

		OPER xi = Excel(xlfEvaluate, x);
		if (x.val.str[1] != '=' and x.val.str[1] != '{') {
			ensure(xi.size() <= 1);
			ref.Set(xi ? xi : x);
		}
		else {
			ref.Reshape(xi);
			if (x.val.str[1] == '=') {
				ref.Formula(x);
			}
			else {
				ensure(xi.size() > 1);
				ref.Set(xi);
			}
		}
			

		return ref;
	}

	// paste add-in to active cell and defaults below
	inline OPER xll_paste_default_args(const Args& args)
	{
		Select xAct;

		// call with defaults arguments to get size of output
		OPER xVal = Excel(xlfEvaluate, args.ArgumentDefault(0));
		OPER xFor = OPER("=") & args.FunctionText() & OPER("(");
		Select sel;
		sel.Move(rows(xVal), 0);

		for (unsigned i = 1; i <= args.ArgumentCount(); ++i) {
			OPER xRel = paste_formula(args.ArgumentDefault(i));
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