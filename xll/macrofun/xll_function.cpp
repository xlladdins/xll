// xll_function.cpp - xlfXXX macro functions
#include "xll_macrofun.h"

using namespace xll;

using xchar = traits<XLOPERX>::xchar;

AddIn xai_eval(
	Function(XLL_LPOPER, "xll_eval", "EVAL")
	.Arguments({
		{XLL_LPOPER, "cell", "is a cell to be evaluated.", "=\"1 + 2\""}
		})
	.FunctionHelp("Call xlfEvaluate on cell.")
	.Category("XLL")
	.Documentation(R"(
<code>EVAL</code> calls the Excel function <code>xlfEvaluate</code> 
to use the Excel engine to evaluate its argument. 
A naked string like <code>abc</code> is interpreted as
a named range. To get <code>EVAL</code> to treat it like
a string it must be enclosed in quotes, <code>\"abc\"</code>.
<p>
Two dimensional ranges are enclosed in curly braces using commas for
field seperators and semi-colons for record seperators. 
For example <code>"{1,\"a\";FALSE,2.34}"</code>.
Individual range elements are not evaluated.
)")
);
LPOPER WINAPI xll_eval(const LPOPER pcell)
{
#pragma XLLEXPORT
	static OPER result;

	try {
		result = Excel(xlfEvaluate, *pcell);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		result = ErrNA;
	}
	catch (...) {
		XLL_ERROR("unknown error");

		result = ErrNA;
	}

	return &result;
}

#ifdef _DEBUG

int test_eval()
{
	try {
		{
			OPER o(1.23);
			OPER e = *xll_eval(&o);
			ensure(e == 1.23);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}

	return TRUE;
}
Auto<OpenAfter> xaoa_test_eval(test_eval);

#endif // _DEBUG

#if 0
AddIn xai_absref(
	Function(XLL_LPOPER, "xll_absref", "XLL.ABSREF")
	.Arguments({
		{XLL_PSTRING, "R1C1", "an R1C1-style relative reference in the form of text", "\"A1\""},
		{XLL_LPXLOPER, "ref", "is a reference.", "$A$1"}
		})
	.Uncalced()
	.FunctionHelp("Converts a reference to an absolute reference in the form of text.")
	.HelpTopic("https://xlladdins.github.io/Excel4Macros/absref.html")
	.Category("XLL")
	.Documentation(R"(
Returns the absolute reference of the cells that are offset from a reference 
by a specified amount. The offset is relative to the upper-left corner of <c>ref</c>. 
You should generally use <c>OFFSET</c> instead of <c>ABSREF</c>. 
This function is provided for users who prefer to supply an absolute reference in text form.
)")
);
LPOPER WINAPI xll_absref(xchar* r1c1, LPOPER pref)
{
#pragma XLLEXPORT
	static OPER result;

	try {
		XLOPERX R1C1;
		R1C1.xltype = xltypeStr;
		R1C1.val.str = r1c1;
		result = Excel(xlfAbsref, R1C1, *pref);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
		result = ErrNA;
	}
	catch (...) {
		XLL_ERROR("unknown exception");
		result = ErrNA;
	}

	return &result;
}
AddIn xai_reftext(
	Function(XLL_LPOPER, "xll_reftext", "XLL.REFTEXT")
	.Arguments({
		{XLL_LPXLOPER, "ref", "is a cell reference.", "\"A1\""},
		{XLL_BOOL, "A1", "is a boolean specifying A1 or R1C1-style references.", "=\"FALSE\""}
		})
	.Uncalced()
	.FunctionHelp("Converts a reference to an absolute reference in the form of text.")
	.HelpTopic("https://xlladdins.github.io/Excel4Macros/reftext.html")
	.Category("XLL")
	.Documentation(R"(
Use <c>REFTEXT</c> when you need to manipulate references with text functions. 
After manipulating the reference text, you can convert it back into a normal reference 
by using <c>TEXTREF</c>.
)")
);
LPOPER WINAPI xll_reftext(LPOPER pref, BOOL a1)
{
#pragma XLLEXPORT
	static OPER result;

	try {
		result = Excel(xlfReftext, *pref, OPER(a1));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
		result = ErrNA;
	}
	catch (...) {
		XLL_ERROR("unknown exception");
		result = ErrNA;
	}

	return &result;
}
#endif // 0