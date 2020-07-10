#include <cmath>
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
		ArgX({ XLL_DOUBLEX, X_("x"), X_("is an arg") })
	})
);
double WINAPI xll_tgamma(double x)
{
#pragma XLLEXPORT
	return tgamma(x);
}
