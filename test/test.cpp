#include <cmath>
// Uncomment to use Excel4 API.
//#define XLOPERX XLOPER
#include "../xll/xll.h"

using namespace xll;

// Use Alt-F8 then type 'XLL.MACRO' to call 'xll_macro'
// See https://github.com/xlladdins/xll/blob/master/docs/Excel4Macros/README.md
// for documentation of ExcelX arguments.
AddInX xai_macro(MacroX(X_("?xll_macro"), X_("XLL.MACRO")));
// All functions called from Excel must be declared with WINAPI.
int WINAPI xll_macro(void)
{
#pragma XLLEXPORT
	//OPERX xMsg(X_("XLL.MACRO called with active cell: "));
	//OPERX xActive = Excel<XLOPERX>(xlfActiveCell);
	//OPERX xReftext = Excel<XLOPERX>(xlfReftext, xActive, OPERX(true)); // A1 style
	//xMsg &= xReftext;
	//Excel<XLOPERX>(xlcAlert, xMsg);

	// same as above
	ExcelX(xlcAlert, 
		ExcelX(xlfConcatenate,
			OPERX(X_("XLL.MACRO called with active cell: ")),
			ExcelX(xlfReftext, ExcelX(xlfActiveCell), OPERX(true))
		),
		OPERX(2), // general information
		OPERX(X_("https://github.com/xlladdins/xll/blob/master/docs/Excel4Macros/ALERT.md!0"))
		// Optional help file link. Note the '!0' appended to the URL.
	);
	
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
	// This results in #NUM! if returned to Excel.
	double result = std::numeric_limits<double>::quiet_NaN();

	try {
		result = _jn(n, x);
	}
	catch (const std::excetion& ex) {
		XLL_ERROR(ex.what());
	}

	return result;
}

Auto<Open> xai_open([]() { 
	Excel<XLOPERX>(xlcAlert, OPERX(X_("Auto<Open> called")));

	return TRUE;  
});

Auto<Close> xai_close([]() {
	Excel<XLOPERX>(xlcAlert, OPERX(X_("Auto<Close> called")));

	return TRUE;
});

AddInX xai_onkey(MacroX(X_("?xll_onkey"), X_("XLL.ONKEY")));
int WINAPI xll_onkey(void)
{
#pragma XLLEXPORT
	ExcelX(xlcAlert, OPERX(X_("You pressed Ctrl-Alt-a")));

	return TRUE;
}

On<Key> xon_key(ON_CTRL ON_ALT X_("a"), X_("XLL.ONKEY"));

AddInX xai_onwindow(MacroX(X_("?xll_onwindow"), X_("XLL.ONWINDOW")));
int WINAPI xll_onwindow(void)
{
#pragma XLLEXPORT
	ExcelX(xlcAlert, OPERX(X_("ONWINDOW called")));

	return TRUE;
}
On<Window> xon_window(X_(""), X_("XLL.ONWINDOW"));

AddInX xai_onsheet(MacroX(X_("?xll_onsheet"), X_("XLL.ONSHEET")));
int WINAPI xll_onsheet(void)
{
#pragma XLLEXPORT
	ExcelX(xlcAlert, OPERX(X_("ONSHEET called")));

	return TRUE;
}
On<Sheet> xon_sheet(X_(""), X_("XLL.ONSHEET"), true);

AddInX xai_get_workspace(
	FunctionX(XLL_LPOPERX, X_("?xll_get_workspace"), X_("GET_WORKSPACE"))
	.Args({
		ArgX(XLL_SHORTX, X_("type_num"), X_("is a number specifying the type of workspace information you want."))
	})
	.Uncalced()
);
LPOPERX WINAPI xll_get_workspace(SHORT type_num)
{
#pragma XLLEXPORT
	static OPERX oResult;

	oResult = ExcelX(xlfGetWorkspace, OPERX(type_num));

	return &oResult;
}