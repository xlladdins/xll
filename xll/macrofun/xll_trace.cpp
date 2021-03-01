// trace.cpp - trace Excel calculation
#include "../xll.h"

using namespace xll;

static AddIn xai_trace(
    Function(XLL_LPXLOPER, "xll_trace", "TRACE")
    .Arguments({
        Arg(XLL_LPXLOPER, "cell", "is the cell to trace."),
    })
    .Category("XLL")
    .FunctionHelp("Alert when cell is called in a calculation.")
    .Uncalced()
);
LPXLOPERX WINAPI
xll_trace(LPXLOPERX px)
{
#pragma XLLEXPORT
    if (px->xltype == xltypeSRef) {
        OPER s;

        s =  OPER("Reference: ") & Excel(xlfReftext, *px, OPER(true));
        s &= OPER("\nFormula: ") & Excel(xlfGetCell, OPER(6), *px);
        s &= OPER("\nContents: ") & Excel(xlfText, Excel(xlfGetCell, OPER(5), *px), OPER("General"));

        MessageBoxA(GetActiveWindow(), s.to_string().c_str(), "TRACE", MB_OK);
    }

    return px;
}

