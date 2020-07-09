#include <cmath>
#include "../xll/xll.h"

using namespace xll;

AddInX xai_macro(MacroX(L"?xll_macro", L"XLL.MACRO"));
int WINAPI xll_macro(void)
{
#pragma XLLEXPORT
	return TRUE;
}