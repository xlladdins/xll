#include <cmath>
// Uncomment to use Excel4 API. Default is XLOPER12.
//#define XLL_VERSION 4
#include "../xll/xll.h"
#include "../xll/macrofun/xll_macrofun.h"

using namespace xll;

#ifdef _DEBUG
Auto<OpenAfter> xaoa_test_doc([]() {
	return Documentation("TEST", "Excel test functions");
});
#endif

/*
Auto<OpenAfter> xaoa_test([]() {
	OPER o;
	o = Workbook::Select();
	o = Workbook::Insert();
	o = Workbook::Select();

	o = OPER{};

	return TRUE;
});
*/

// test for 64-bit excel?
static LSTATUS reg_query = []() {

	LSTATUS status;
	auto HKLM = HKEY_LOCAL_MACHINE;
	LPCSTR key = "SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration";
	LPCSTR val = "Platform";
	DWORD flags = RRF_RT_REG_SZ;
	DWORD type;
	char data[1024];
	DWORD len = 1023;
	status = RegGetValueA(HKLM, key, val, flags, &type, (PVOID)data, &len);

	return status;

}();

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
	.Arguments({
		Arg(XLL_DOUBLE, "x", "is the value for which you want to calculate Gamma.", "10*rand()")
	})
	.FunctionHelp("Return the Gamma function value.")
	.Category("CMATH")
	.HelpTopic("https://docs.microsoft.com/en-us/cpp/c-runtime-library/reference/tgamma-tgammaf-tgammal")
	.Documentation(R"xyz(
The Gamma function is \(\Gamma(\alpha) = \int_0^\infty x^{\alpha - 1} e^{-x}\,dx\).
It satisfies \(\Gamma(n + 1) = n!\) if \(n\) is a natural number.
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
	.Arguments({
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
	.Arguments({
		Arg(XLL_LONG, "n", "is the order of the Bessel function.", "1"),
		Arg(XLL_DOUBLE, "x", "is the value for which you want to calculate the Bessel function.", "1+.1")
	})
	.FunctionHelp("Return the value of the n-th order Bessel function of the first kind.")
	.Category("CMATH")
	.Documentation(R"xyzyx(The \(n\)-th order Bessel function of the first kind
\[
	J_n(x) = \sum_{m=0}^\infty \frac{(-1)^m}{m!\Gamma(m + n + 1)}\left(\frac{x}{2}\right)^{2m + n}.
\]
)xyzyx")
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

#if 0
// lambda gets called on open
Auto<Open> xai_open([]() {
	return !Excel(xlcAlert, OPER("Auto<Open> called")).is_err();
});

Auto<Close> xai_close([]() {
	Excel(xlcAlert, OPER("Auto<Close> called"));

	return TRUE;
});

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
#endif

AddIn xai_onkey(Macro("xll_onkey", "XLL.ONKEY").FunctionHelp("Called when Ctrl-Alt-a is pressed."));
int WINAPI xll_onkey(void)
{
#pragma XLLEXPORT
	Excel(xlcAlert, OPER("You pressed Ctrl-Alt-a"));

	return TRUE;
}
On<Key> xon_key(ON_CTRL ON_ALT "a", "XLL.ONKEY");

AddIn xai_get_workspace(
	Function(XLL_LPOPER, "xll_get_workspace", "GET_WORKSPACE")
	.Arguments({
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
	.Arguments({
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

// ma_n = (x_1 + ... + x_n)/n
//      = (ma_{n-1} + x_n)/n
//      = ma_{n-1} + (- ma_{n-1} + x)/n
class MA {
	size_t n;
	double ma;
public:
	MA()
		: n(0), ma(0)
	{ }
	size_t count() const 
	{
		return n;
	}
	double value() const
	{
		return ma;
	}
	MA& next(double x)
	{
		++n;
		ma += (x - ma) / n;

		return *this;
	}
};

#if 0
AddIn xai_ma(
	Function(XLL_HANDLEX, "xll_ma", "\\XLL.MA")
	.Arguments({
		Arg(XLL_FP, "init", "array of initial data for moving averate."),
	})
	.Uncalced()
	.Category("XLL")
	.FunctionHelp("Compute a moving average.")
	.Documentation("Initialize moving average.")
);
HANDLEX WINAPI xll_ma(_FPX *px)
{
#pragma XLLEXPORT
	handle<MA> ma(new MA{});

	for (unsigned i = 0; i < size(*px); ++i) {
		ma->next(px->array[i]);
	}

	return ma.get();
}

AddIn xai_ma_next(
	Function(XLL_FP, "xll_ma_next", "XLL.MA.NEXT")
	.Arguments({
		Arg(XLL_HANDLEX, "handle", "is a handle to a moving average."),
		Arg(XLL_DOUBLE, "x", "is the value whos moving averate is to be computed."),
		Arg(XLL_BOOL, "reset", "reset the moving average computation."),
		})
	.Category("XLL")
	.FunctionHelp("Compute a moving average.")
	.Documentation("Update moving average.")
);
_FPX* WINAPI xll_ma_next(HANDLEX ma, double x, BOOL reset)
{
#pragma XLLEXPORT
	static xll::FPX_<1, 2> result;

	handle<MA> ma_(ma);
	ensure(ma_);
	if (reset) {
		*ma_ = MA{};
	}
	else {
		ma_->next(x);
	}

	result[0] = ma_->value();
	result[1] = static_cast<double>(ma_->count());

	return (_FPX*)&result;
}
#endif