// xll_throttle.cpp - Only call on F2-Enter
#include "../xll.h"

using namespace xll;

AddIn xai_throttle(
	Function(XLL_LPXLOPERX, "xll_throttle", "THROTTLE")
	.Arguments({
		Arg(XLL_LPXLOPERX, "formula", "is a formula.")
		})
	.Uncalced()
	.Category("XLL")
	.FunctionHelp("Evaluate formula only if F2, Enter is used on the cell.")
	.Documentation(
		"Block a formula from being called."
	)
);
LPXLOPERX WINAPI xll_throttle(LPXLOPERX px)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		o = Excel(xlCoerce, Excel(xlfCaller));
		if (o[0].is_num() && o[0].as_num() == 0) {
			o = *px; // Excel(xlfEvaluate, *px);
		}
		px = &o;
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR(__FUNCTION__ ": unknown exception");
	}

	return px;
}
