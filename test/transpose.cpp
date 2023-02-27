// transpose.cpp - Transpose an array of doubles
#include "../xll/xll.h"

using namespace xll;

AddIn xai_transpose(
	Function(XLL_FPX, "xll_transpose", "XLL.TRANSPOSE")
	.Arguments({
		Arg(XLL_FPX, "array", "is an array of doubles.")
		})
	.FunctionHelp("Transpose an array of doubles.")
	.Category("XLL")
);
_FP12* WINAPI xll_transpose(_FP12* pa)
{
#pragma XLLEXPORT
	// Only transposes if rows or columns is 1.
	std::swap(pa->rows, pa->columns);

	return pa;
}