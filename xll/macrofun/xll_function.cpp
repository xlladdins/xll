// xll_function.cpp - xlfXXX macro functions
#include "xll_macrofun.h"

using namespace xll;

using xchar = traits<XLOPERX>::xchar;

AddIn xai_eval(
	Function(XLL_LPOPER, "xll_eval", "EVAL")
	.Arguments({
		{XLL_PSTRING, "str", "is a string to be evaluated.", "1 + 2*rand()"}
		})
	.FunctionHelp("Call xlfEvaluate on string.")
	.Category("XLL")
	.Documentation(R"(
The function <code>xlfEvaluate</code> evaluates
its string argument, just like pressing <code>F9</code> evaluates selected text
in the formula bar. A naked string like <code>abc</code> is interpreted as
a named range and returns <code>#NAME?</code> if it is not defined.
To get <code>EVAL</code> to treat it like
a string it must be enclosed in quotes, <code>"abc"</code>. The naked
string without quotes is the result.
<p>
Two dimensional ranges start with curly braces and use commas for
field separators and semi-colons for record separators, 
for example <code>{1,\"a\";FALSE,2.34}</code>.
The range can only contain constants, not formulas, just like
<a href="https://docs.microsoft.com/en-us/office/client-developer/excel/xlset"><code>xlSet</code></a>.
)")
);
LPOPER WINAPI xll_eval(xchar* str)
{
#pragma XLLEXPORT
	static OPER result;

	try {
		XLOPERX Str;
		Str.val.str = str;
		Str.xltype = xltypeStr;

		result = Excel(xlfEvaluate, Str);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
		result = ErrNA;
	}
	catch (...) {
		XLL_ERROR(__FUNCTION__ "unknown error");
		result = ErrNA;
	}

	return &result;
}

#if 0
AddIn xai_absref(
	Function(XLL_LPOPER, "xll_absref", "XLL.ABSREF")
	.Arguments({
		{XLL_PSTRING, "R1C1", "an R1C1-style relative reference in the form of text", "\"R[1]C[-1]\""},
		{XLL_LPXLOPER, "ref", "is a reference.", "$A$1"}
		})
	.Uncalced()
	.FunctionHelp("Converts a reference to an absolute reference in the form of text.")
	.HelpTopic("https://xlladdins.github.io/Excel4Macros/absref.html")
	.Category("XLL")
	.Documentation(R"(
Returns the absolute reference of the cells that are offset from a reference 
by a specified amount. The offset is relative to the upper-left corner of <code>ref</code>. 
You should generally use <code>OFFSET</code> instead of <code>ABSREF</code>. 
This function is provided for users who prefer to supply an absolute reference in text form.
)")
);
LPXLOPERX WINAPI xll_absref(xchar* r1c1, LPXLOPERX pref)
{
#pragma XLLEXPORT
	static XLOPERX result;

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
		XLL_ERROR(__FUNCTION__ ": unknown exception");
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