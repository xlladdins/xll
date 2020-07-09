#include <cmath>
#include "../xll/xll.h"

using namespace xll;

AddInX xai_macro(MacroX(L"?xll_macro", L"XLL.MACRO"));
int WINAPI xll_macro(void)
{
#pragma XLLEXPORT
	return TRUE;
}

AddInX xai_tgamma(
	FunctionX(XLL_DOUBLE, L"?xll_tgamma", L"TGAMMA")
	.Args({
		ArgX({ XLL_DOUBLE, L"x", L"is an arg" })
	})
);
double WINAPI xll_tgamma(double x)
{
#pragma XLLEXPORT
	return tgamma(x);
}
