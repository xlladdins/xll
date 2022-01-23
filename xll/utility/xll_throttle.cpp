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
	try {
		OPER o = Excel(xlCoerce, Excel(xlfCaller));
		if (Excel(xlCoerce, Excel(xlfCaller)).as_num() == 0) {
			*(LPOPER)px = Excel(xlfEvaluate, *(LPOPER)px);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR(__FUNCTION__ ": unknown exception");
	}

	return px;
}
