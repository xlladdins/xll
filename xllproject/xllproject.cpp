// $projectname$.cpp - Sample xll project.
#include <cmath>
//Uncomment to build for versions of Excel prior to 2007
//#define XLOPERX XLOPER
#include "xll/xll/xll.h"

using namespace xll;

AddIn xai_tgamma(
	// Return double, C++ name of function, Excel name.
	Function(XLL_DOUBLE, "xll_tgamma", "TGAMMA")
	// Array of function arguments.
	.Args({
		Arg(XLL_DOUBLE, "x", "is the value for which you want to calculate Gamma.")
	})
	// Function Wizard help.
	.FunctionHelp("Return the Gamma function value.")
	// Function Wizard category.
	.Category("CMATH")
	// URL linked to `Help on this function`.
	.HelpTopic("https://docs.microsoft.com/en-us/cpp/c-runtime-library/reference/tgamma-tgammaf-tgammal")
);
// You **must** specify WINAPI calling convention.
double WINAPI xll_tgamma(double x)
{
#pragma XLLEXPORT

	return tgamma(x);
}

// Press Alt-F8 then type 'XLL.MACRO' to call 'xll_macro'
// See https://github.com/xlladdins/xll/blob/master/docs/Excel4Macros/README.md
AddIn xai_macro(
	// C++ function, Excel name of macro
	Macro("xll_macro", "XLL.MACRO")
);
// All macros must have `int WINAPI (*)(void)` signature.
int WINAPI xll_macro(void)
{
#pragma XLLEXPORT
	Excel(xlcAlert,
		OPER("XLL.MACRO called with active cell: ")
			& Excel(xlfReftext, Excel(xlfActiveCell), 
		OPER(true)) // A1 style instead of default R1C1
	);

	return TRUE;
}
