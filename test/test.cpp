#include <cmath>
// Uncomment to use Excel4 API. Default is XLOPER12.
//#define XLOPERX XLOPER
#include "../xll/xll.h"

using namespace xll;

// Use Alt-F8 then type 'XLL.MACRO' to call 'xll_macro'
// See https://xlladdins.github.io/Excel4Macros/
// for documentation of Excel arguments.
AddIn xai_macro(Macro("xll_macro", "XLL.MACRO"));
// All functions called from Excel must be declared with WINAPI.
int WINAPI xll_macro(void)
{
// And made available to Excel using a pragma.
#pragma XLLEXPORT

	Excel(xlcAlert,
		Excel(xlfConcatenate,
			OPER("XLL.MACRO called with активный cell: "),
			Excel(xlfReftext, Excel(xlfActiveCell), OPER(true))
		),
		OPER(2), // general information
		OPER("https://xlladdins.github.io/Excel4Macros/alert.html!0") // help
	);

	return TRUE;
}

AddIn xai_tgamma(
	// Return a double by calling xll_tgamma using TGAMMA in Excel.
	Function(XLL_DOUBLE, "xll_tgamma", "TGAMMA")
	// Args are an array of one Arg that is a double. 
	.Args({
		Arg(XLL_DOUBLE, "x", "is the value for which you want to calculate Gamma.")
	})
	.FunctionHelp("Return the Gamma function value.")
	.Category("CMATH")
	.HelpTopic("https://docs.microsoft.com/en-us/cpp/c-runtime-library/reference/tgamma-tgammaf-tgammal")
	.Documentation(R"xyz(
The Gamma function satisfies <math>&Gamma;(<i>n</i> + 1) = <i>n</i>!</math> if <i>n</i> is a natural number.
)xyz")
);
double WINAPI xll_tgamma(double x)
{
#pragma XLLEXPORT
	return tgamma(x);
}

#if 0 // change to 1 to test duplicate function names
// AddIn previously defined: TGAMMA 
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
#endif 

AddIn xai_jn(
	Function(XLL_DOUBLE, "xll_jn", "JN")
	.Args({
		Arg(XLL_LONG, "n", "is the order of the Bessel function.", "=1"),
		Arg(XLL_DOUBLE, "x", "is the value for which you want to calculate the Bessel function.", "=1+.1")
	})
	.FunctionHelp("Return the value of the n-th order Bessel function of the first kind.")
	.Category("CMATH")
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

// lambda gets called on open
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
	.FunctionHelp("Return workspace information.")
	.Category("XLL")
	.HelpTopic("https://xlladdins.github.io/Excel4Macros/get.workspace.html")
);
LPOPER WINAPI xll_get_workspace(SHORT type_num)
{
#pragma XLLEXPORT
	static OPER oResult;

	try {
		oResult = Excel(xlfGetWorkspace, OPER(type_num));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		oResult = ErrValue;
	}

	return &oResult;
}

// Force Excel to use the old API.
AddIn4 xai_set_range(
	Function4(XLL_HANDLEX, "xll_set_range", "SET.RANGE")
	.Args({
		Arg(XLL_LPOPER4, "range", "is a range")
	})
	.FunctionHelp("Return a handle to a range.")
	.Category("XLL")
	.Uncalced()
);
HANDLEX WINAPI xll_set_range(LPOPER4 po) {
#pragma XLLEXPORT
	handle<OPER4> h(new OPER4(*po));

	return h.get();
}
AddIn4 xai_get_range(
	Function4(XLL_LPOPER4, "xll_get_range", "GET.RANGE")
	.Args({
		Arg(XLL_HANDLEX, "handle", "is a handle to a range")
		})
	.FunctionHelp("Return a range corresponding to handle.")
	.Category("XLL")
);
LPOPER4 WINAPI xll_get_range(HANDLEX _h) {
#pragma XLLEXPORT
	handle<OPER4> h(_h);

	h->operator()(0, 0) = "";

	return h.ptr();
}

// UTF-8 test
AddIn xai_utf8(Macro("xll_utf8", "XLL.UTF8"));
int WINAPI xll_utf8(void)
{
#pragma XLLEXPORT
	Excel(xlcAlert,
		OPER("отлично"),
		OPER(2), // general information
		OPER("https://translate.google.com/#view=home&op=translate&sl=auto&tl=en&text=%D0%BE%D1%82%D0%BB%D0%B8%D1%87%D0%BDo!0")
	);

	return TRUE;
}

AddIn xai_file(
	Function(XLL_LPOPER, "xll_file", "XLL.FILE")
	.Args({
		Arg(XLL_LPOPER, "url", "is a URL to retrieve.")
	})
	.Category("XLL")
	.FunctionHelp("Return contents of URL.")
);
LPOPER WINAPI xll_file(const LPOPER po)
{
#pragma XLLEXPORT
	static OPER f;

	f = Excel(xlfWebservice, *po);

	return &f;
}

int test_doc()
{
	Document("TEST");

	return TRUE;
}
Auto<OpenAfter> xaoa_doc(test_doc);