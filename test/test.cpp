#include <cmath>
//#define XLOPERX XLOPER
#include "../xll/xll.h"

using namespace xll;

// Use Alt-F8 then type 'XLL.MACRO' to call 'xll_macro'
AddInX xai_macro(MacroX(X_("?xll_macro"), X_("XLL.MACRO")));
int WINAPI xll_macro(void)
{
#pragma XLLEXPORT
	//OPERX xMsg(X_("XLL.MACRO called with active cell: "));
	//OPERX xActive = Excel<XLOPERX>(xlfActiveCell);
	//OPERX xReftext = Excel<XLOPERX>(xlfReftext, xActive, OPERX(true)); // A1 style
	//xMsg &= xReftext;
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

//!!!not being called!!!
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

AddInX xai_onwindow2(MacroX(X_("?xll_onwindow2"), X_("XLL.ONWINDOW2")));
int WINAPI xll_onwindow2(void)
{
#pragma XLLEXPORT
	ExcelX(xlcAlert, OPERX(X_("ONWINDOW2 called")));
	ExcelX(xlcDefineName, OPERX(X_("foo")), OPERX(1.23));

	return TRUE;
}
On<Window> xon_window2(X_(""), X_("XLL.ONWINDOW2"));

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
		ArgX(XLL_LONGX, X_("type_num"), X_("is a number specifying the type of workspace information you want."))
	})
	.Uncalced()
);
LPOPERX WINAPI xll_get_workspace(LONG type_num)
{
#pragma XLLEXPORT
	static OPERX oResult;

	oResult = ExcelX(xlfGetWorkspace, OPERX(type_num));

	return &oResult;
}