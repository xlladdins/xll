#include <cmath>
// Uncomment to use Excel4 API.
//#define XLOPERX XLOPER
#include "../xll/xll.h"

using namespace xll;

// Use Alt-F8 then type 'XLL.MACRO' to call 'xll_macro'
// See https://github.com/xlladdins/xll/blob/master/docs/Excel4Macros/README.md
// for documentation of Excel arguments.
AddIn xai_macro(Macro("xll_macro", "XLL.MACRO"));
// All functions called from Excel must be declared with WINAPI.
int WINAPI xll_macro(void)
{
#pragma XLLEXPORT
	//OPER xMsg("XLL.MACRO called with active cell: ");
	//OPER xActive = Excel(xlfActiveCell);
	//OPER xReftext = Excel(xlfReftext, xActive, OPER(true)); // A1 style
	//xMsg &= xReftext;
	//Excel(xlcAlert, xMsg);

	// same as above
	Excel(xlcAlert,
		Excel(xlfConcatenate,
			OPER("XLL.MACRO called with active cell: "),
			Excel(xlfReftext, Excel(xlfActiveCell), OPER(true))
		),
		OPER(2), // general information
		OPER("https://github.com/xlladdins/xll/blob/master/docs/Excel4Macros/ALERT.md")
		// Optional help file link.
	);

	return TRUE;
}

AddIn xai_tgamma(
	Function(XLL_DOUBLE, "xll_tgamma", "TGAMMA")
	.Args({
		Arg(XLL_DOUBLE, "x", "is the value for which you want to calculate Gamma.")
		})
	.FunctionHelp("Return the Gamma function value.")
	.Category("CMATH")
	.HelpTopic("https://docs.microsoft.com/en-us/cpp/c-runtime-library/reference/tgamma-tgammaf-tgammal")
);
double WINAPI xll_tgamma(double x)
{
#pragma XLLEXPORT
	return tgamma(x);
}

#if 0
/* AddIn previously defined: TGAMMA */
AddIn xai_tgamma2(
	Function(XLL_DOUBLE, "xll_tgamma2", "TGAMMA")
	.Args({
		Arg(XLL_DOUBLE, "x", "is the value for which you want to calculate Gamma.")
		})
	.FunctionHelp("Return the Gamma function value.")
	.Category("CMATH")
	.HelpTopic("https://docs.microsoft.com/en-us/cpp/c-runtime-library/reference/tgamma-tgammaf-tgammal")
);
double WINAPI xll_tgamma2(double x)
{
#pragma XLLEXPORT
	return tgamma(x);
}
#endif // 0 

AddIn xai_jn(
	Function(XLL_DOUBLE, "xll_jn", "JN")
	.Args({
		Arg(XLL_LONG, "n", "is the order of the Bessel function.", "=1"),
		Arg(XLL_DOUBLE, "x", "is the value for which you want to calculate the Bessel function.", "=1+.1")
		})
	.FunctionHelp("Return the value of the n-th order Bessel function of the first kind.")
	.Category("Cmath")
);
double WINAPI xll_jn(LONG n, double x)
{
#pragma XLLEXPORT
	// This results in #NUM! if returned to Excel.
	double result = std::numeric_limits<double>::quiet_NaN();

	// Catch exeptions.
	try {
		result = _jn(n, x);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR("Unknown exception thrown in: " __FUNCTION__);
	}

	return result;
}

Auto<Open> xai_open([]() {
	Excel(xlcAlert, OPER("Auto<Open> called"));

	return TRUE;
	});

Auto<Close> xai_close([]() {
	Excel(xlcAlert, OPER("Auto<Close> called"));

	return TRUE;
	});

AddIn xai_onkey(Macro("xll_onkey", "XLL.ONKEY"));
int WINAPI xll_onkey(void)
{
#pragma XLLEXPORT
	Excel(xlcAlert, OPER("You pressed Ctrl-Alt-a"));

	return TRUE;
}
On<Key> xon_key(ON_CTRL ON_ALT "a", "XLL.ONKEY");

AddIn xai_onwindow(Macro("xll_onwindow", "XLL.ONWINDOW"));
int WINAPI xll_onwindow(void)
{
#pragma XLLEXPORT
	Excel(xlcAlert, OPER("ONWINDOW called"));

	return TRUE;
}
On<Window> xon_window("", "XLL.ONWINDOW");

AddIn xai_onsheet(Macro("xll_onsheet", "XLL.ONSHEET"));
int WINAPI xll_onsheet(void)
{
#pragma XLLEXPORT
	Excel(xlcAlert, OPER("ONSHEET called"));

	return TRUE;
}
On<Sheet> xon_sheet("", "XLL.ONSHEET", true);

AddIn xai_get_workspace(
	Function(XLL_LPOPER, "xll_get_workspace", "GET_WORKSPACE")
	.Args({
		Arg(XLL_SHORT, "type_num", "is a number specifying the type of workspace information you want.")
		})
	.Uncalced()
);
LPOPER WINAPI xll_get_workspace(SHORT type_num)
{
#pragma XLLEXPORT
	static OPER oResult;

	oResult = Excel(xlfGetWorkspace, OPER(type_num));

	return &oResult;
}

AddIn xai_get_formula(
	Function(XLL_HANDLEX, "xll_get_formula", "GET_FORMULA")
	.Args({
		Arg(XLL_LPXLOPER, "cell", "is a reference to a cell containing a formula.")
		})
	.FunctionHelp("Get formula from cell.")
	.Category("XLL")
	.Uncalced()
);
HANDLEX WINAPI xll_get_formula(LPXLOPERX pCell)
{
#pragma XLLEXPORT
	// if pCall->xltype == xltypeMissing use active cell
	OPER xSS = Excel(xlcSelectSpecial, OPER(9));

	OPER xFormula = Excel(xlfGetFormula, *pCell); // formula references are R1C1

	return HANDLEX{};
}

