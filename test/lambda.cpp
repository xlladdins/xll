#include "../xll/xll.h"
using namespace xll;

// ,LPXLOPERX, LPXLOPERX, ...
#define LPXLOPERX1 ,LPXLOPERX
#define LPXLOPERX2 LPXLOPERX1 LPXLOPERX1
#define LPXLOPERX4 LPXLOPERX2 LPXLOPERX2
#define LPXLOPERX8 LPXLOPERX4 LPXLOPERX4
#define LPXLOPERX16 LPXLOPERX8 LPXLOPERX8
#define LPXLOPERX32 LPXLOPERX16 LPXLOPERX16

namespace xll {

	// first missing arg determines length
	template<class X>
	inline size_t len(X** ppa)
	{
		auto missing = [](X* pa) { return pa->xltype == xltypeMissing; };
		auto i = std::find_if(ppa, ppa + traits<X>::argmax - 1, missing);

		return i - ppa;
	}

	// 1.2, {"foo", "bar"}, true, ...
	// _1 = 1.2, foo = "bar", _2 = true, ...
	template<class X>
	inline void SetNames(X** ppa)
	{
		int j = 1;
		for (int i = 0; i < len(ppa); ++i) {
			XOPER<X> ki, vi; // key, value
			const XOPER<X>& ai = *ppa[i];

			if (ai.size() == 1) {
				ki = Excel<X>(xlfText, XOPER<X>(j), XOPER<X>("\"_\"0"));
				vi = ai;
				++j;
			}
			else if (ai.size() == 2) {
				ki = ai[0];
				vi = ai[1];
				ensure(ki.xltype == xltypeStr);
			}
			else {
				XLL_ERROR("XLL.EVAL: args must be 1 or 2 dimensional");
			}

			ensure(XExcel<X>(xlfSetName, ki, vi));
			//ensure(XExcel<X>(xlfEvaluate, XExcel<X>(xlfGetName, ki)) == vi);
		}
	}
	template<class X>
	inline void UnsetNames(X** ppa)
	{
		int j = 1;
		for (int i = 0; i < len(ppa); ++i) {
			XOPER<X> ki, vi; // key, value
			const XOPER<X>& ai = *ppa[i];

			if (ai.size() == 1) {
				ki = Excel<X>(xlfText, XOPER<X>(j), XOPER<X>("\"_\"0"));
				++j;
			}
			else if (ai.size() == 2) {
				ki = ai[0];
				ensure(ki.xltype == xltypeStr);
			}
			else {
				XLL_ERROR("XLL.EVAL: args must be 1 or 2 dimensional");
			}

			ensure(XExcel<X>(xlfSetName, ki));
		}
	}
} // namespace xll

static std::vector<const char*> as(30, XLL_LPOPER); // use argmax?

AddIn xai_udf(
	Function(XLL_LPOPER, "xll_udf", "XLL.UDF")
	.Args(as)
	.Uncalced()
	.FunctionHelp("Call user defined function.")
	.Category("XLL")
	.Documentation(R"(
Call a user defined function given its arguments.
Either the string or register id of the udf may be used.
	)")
);
LPXLOPERX WINAPI xll_udf(LPXLOPERX pa LPXLOPERX32)  
{
#pragma XLLEXPORT
	static XLOPERX r = ErrValue;

	try {
		XLOPERX** ppa = &pa;

		r = Excelv(xlfEvaluate, len(ppa), ppa);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &r;
}

AddIn xai_eval(
	Function(XLL_LPOPER, "xll_eval", "XLL.EVAL")
	.Args(as)
	.Uncalced()
	.FunctionHelp("Evaluate a function.")
	.Category("XLL")
	.Documentation(R"(
The first argument is a string expression to evaluate.
Occurences of _1, _2, ... are replaced by the following arguments.
If the argument is a key-value pair the key is replaced by the value in the string expression.
To use a cell reference or existing name prepend an exclamation point.		
	)")
);
LPXLOPERX WINAPI xll_eval(LPXLOPERX pa LPXLOPERX32)
{
#pragma XLLEXPORT
	static XLOPERX r = ErrValue;

	try {
		XLOPERX** ppa = &pa;

		SetNames(ppa + 1);
		r = Excelv(xlfEvaluate, 1, ppa);
		UnsetNames(ppa + 1);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &r;
}

