#include <cmath>
//Uncomment to build for versions of Excel prior to 2007
//#define XLOPERX XLOPER
#include "../xll/xll.h"

using namespace xll;

// Use Alt-F8 then type 'XLL.MACRO' to call 'xll_macro'
AddInX xai_macro(
	// C++ function with '?' prepended, Excel name of macro
	MacroX(X_("?xll_macro"), X_("XLL.MACRO")));
int WINAPI xll_macro(void)
{
#pragma XLLEXPORT
	ExcelX(xlcAlert, 
		OPERX(X_("XLL.MACRO called with active cell: ")) 
			& ExcelX(xlfReftext, ExcelX(xlfActiveCell), OPERX(true)));
	
	return TRUE;
}

AddInX xai_tgamma(
	// Returns double, C++ name with '?' prepended, Excel name of function
	FunctionX(XLL_DOUBLEX, X_("?xll_tgamma"), X_("TGAMMA"))
	// Array of function arguments
	.Args({
		ArgX(XLL_DOUBLEX, X_("x"), X_("is the value for which you want to calculate Gamma."))
	})
	.FunctionHelp(X_("Return the Gamma function value."))
	.Category(X_("Cmath"))
);
double WINAPI xll_tgamma(double x)
{
#pragma XLLEXPORT

	return tgamma(x);
}
