#include <cmath>
//#define XLOPERX XLOPER
#include "../xll/xll.h"

using namespace xll;

// Call using Alt-F8
AddInX xai_macro(MacroX(X_("?xll_macro"), X_("XLL.MACRO")));
int WINAPI xll_macro(void)
{
#pragma XLLEXPORT
	OPERX xMsg(X_("XLL.MACRO called with active cell: "));
	OPERX xActive = Excel<XLOPERX>(xlfActiveCell);
	OPERX xReftext = Excel<XLOPERX>(xlfReftext, xActive, OPERX(true)); // A1 style
	xMsg &= xReftext;
	//Excel<XLOPERX>(xlcAlert, xMsg);

	// same as above
	Excel<XLOPERX>(xlcAlert, 
		OPERX(X_("XLL.MACRO called with active cell: ")) 
			& Excel<XLOPERX>(xlfReftext, Excel<XLOPERX>(xlfActiveCell), OPERX(true)));
	
	return TRUE;
}

AddInX xai_tgamma(
	FunctionX(XLL_DOUBLEX, X_("?xll_tgamma"), X_("TGAMMA"))
	.Args({
		ArgX({ XLL_DOUBLEX, X_("x"), X_("is the value for which you want to calculate Gamma.") })
		})
	.FunctionHelp(X_("Return the Gamma function value."))
	.Category(X_("Cmath"))
);
double WINAPI xll_tgamma(double x)
{
#pragma XLLEXPORT
	return tgamma(x);
}

AddInX xai_jn(
	FunctionX(XLL_DOUBLEX, X_("?xll_jn"), X_("JN"))
	.Args({
		ArgX(XLL_LONGX, X_("n"), X_("is the order of the Bessel function.")),
		ArgX(XLL_DOUBLEX, X_("x"), X_("is the value for which you want to calculate the Bessel function."))
		})
	.FunctionHelp(X_("Return the value of the n-th order Bessel function of the first kind."))
	.Category(X_("Cmath"))
);
double WINAPI xll_jn(LONG n, double x)
{
#pragma XLLEXPORT

	return _jn(n, x);
}


Auto<Open> xai_open([]() { 
	Excel<XLOPERX>(xlcAlert, OPERX(X_("Auto<Open> called")));
	return TRUE;  
});

