// trace.cpp - trace Excel calculation
#include "../xll.h"

using namespace xll;

static AddIn xai_trace(
    Function(XLL_LPXLOPER, "xll_trace", "TRACE")
    .Arguments({
        Arg(XLL_LPXLOPER, "ref", "is a reference to a cell to trace."),
    })
    .Category("XLL")
    .FunctionHelp("Alert when a cell is evaluated in a calculation.")
    .Uncalced()
    .Documentation(R"(
The <code>TRACE(ref)</code> function displays a popup dialog when the cell being referenced
is evaluated during a calculation. The dialog displays the
name of the cell reference in A1 format, the formula in the cell that is being
evaluated, and the current contents of the cell.
)")
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

