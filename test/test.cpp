#include <cmath>
//#define XLOPERX XLOPER
#include "../xll/xll.h"

using namespace xll;

AddInX xai_macro(MacroX(X_("?xll_macro"), X_("XLL.MACRO")));
int WINAPI xll_macro(void)
{
#pragma XLLEXPORT
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

Auto<Open> xai_open([]() { 
	Excel<XLOPERX>(xlcAlert, OPERX(X_("Auto<Open> called")));
	return TRUE;  
});
