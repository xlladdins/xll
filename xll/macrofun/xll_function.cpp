// xll_function.cpp - xlfXXX macro functions
#include "xll_macrofun.h"

using namespace xll;

using xchar = traits<XLOPERX>::xchar;

#ifdef _DEBUG
// int breakme = []() { return _crtBreakAlloc = 1295; }();
#endif 

AddIn xai_eval(
	Function(XLL_LPOPER, "xll_eval", "EVAL")
	.Arguments({
		{XLL_LPOPER, "cell", "is a cell to be evaluated.", "=\"1 + 2\""}
		})
	.FunctionHelp("Call xlfEvaluate on cell.")
	.Category("XLL")
	.Documentation(R"xyzyx(
The Excel function <code>xlfEvaluate</code> uses the Excel engine to evaluate
its argument, just like pressing <code>F9</code> evaluates selected text
in the formula bar. A naked string like <code>abc</code> is interpreted as
a named range and it's corresponding value is returned. 
To get <code>EVAL</code> to treat it like
a string it must be enclosed in quotes, <code>"abc"</code>.
If a string is a case-insensitive match with <code>TRUE</code> or <code>FALSE</code>
it is converted to the appropriate boolean value. If the string matches a known
Excel error then the string is converted to an error type. If the string
looks like a function call then Excel calls the function and returns the result.
Use an initial equal sign (<code>=</code>) to force Excel to evaluate the
string as a function. To parse a string as a date use the <code>VALUE()</code> function.
<p>
Two dimensional ranges are enclosed in curly braces, use commas for
field seperators, and semi-colons for record seperators. 
For example, evaluating the string <code>{1.23,"abc";fAlSe,#N/A}</code>
results in the 2x2 range consisting of the number <code>1.23</code>,
the string <code>abc</code>, the boolean <code>FALSE</code> value,
and a "not available" error type. Excel will not attempt to evaluate
any item in a multi-dimensional range as a function.
)xyzyx")
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

#if 0

int test_eval()
{
	try {
		{
			OPER o(1.23);
			OPER e = *xll_eval(&o);
			ensure(e == 1.23);
		}
		bool fail = false;
		try {
			{
				OPER o("abc");
				ensure(*xll_eval(&o)); // should fail
			}
		}
		catch (const std::exception&) {
			fail = true;
		}
		ensure(fail);
		{
			OPER o("\"abc\"");
			OPER e = *xll_eval(&o);
			ensure(e == OPER("abc"));
		}
		{
			OPER o(true);
			OPER e = *xll_eval(&o);
			ensure(e == true);
		}
		{
			OPER o(false);
			OPER e = *xll_eval(&o);
			ensure(e == false);
		}
		// error strings evaluate to corresponding error values
#define X(a,b,c) { OPER o(b); OPER e = *xll_eval(&o); ensure(e.val.err == xlerr##a); } 
		XLL_ERR(X)
#undef X
		{
			// error strings in multis evaluate to errors
			OPER o("{1.23, \"foo\";TRUE, #N/A}");
			OPER e = *xll_eval(&o);
			ensure(e.rows() == 2);
			ensure(e.columns() == 2);
			ensure(e(0, 0) == 1.23);
			ensure(e(0, 1) == "foo");
			ensure(e(1, 0) == true);
			ensure(e(1, 1).xltype == xltypeErr);
			ensure(e(1, 1).val.err == xlerrNA);
		}
		{
			OPER o("=SIN(0)");
			OPER e = *xll_eval(&o);
			ensure(e == 0);
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