#include "../xll/xll.h"
using namespace xll;

static std::vector<const char*> as(30, XLL_LPOPER); // 1 + 29
using OPERX = xll::traits<XLOPERX>::xtype;
using LPOPERX = OPERX*;

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
LPOPERX WINAPI xll_udf( 
	LPOPERX pa, LPOPERX, LPOPERX, LPOPERX,
	LPOPERX, LPOPERX, LPOPERX, LPOPERX, // 8
	LPOPERX, LPOPERX, LPOPERX, LPOPERX, // 12
	LPOPERX, LPOPERX, LPOPERX, LPOPERX, // 16
	LPOPERX, LPOPERX, LPOPERX, LPOPERX, // 20
	LPOPERX, LPOPERX, LPOPERX, LPOPERX, // 24
	LPOPERX, LPOPERX, LPOPERX, LPOPERX, // 28
	LPOPERX, LPOPERX                  // 30
	)
{
#pragma XLLEXPORT
	static OPERX r = ErrValue;

	try {
		OPERX** ppa = &pa;

		int i;
		for (i = 0; i < 32; ++i) {
			if (ppa[i]->xltype == xltypeMissing) {
				break;
			}
		}
		int ret;
		ret = xll::traits<XLOPERX>::Excelv(xlUDF, (LPOPERX)&r, i, (LPOPERX*)ppa);
		ret = ret;
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
LPOPERX WINAPI xll_eval(
	LPOPERX pa, LPOPERX, LPOPERX, LPOPERX,
	LPOPERX, LPOPERX, LPOPERX, LPOPERX, // 8
	LPOPERX, LPOPERX, LPOPERX, LPOPERX, // 12
	LPOPERX, LPOPERX, LPOPERX, LPOPERX, // 16
	LPOPERX, LPOPERX, LPOPERX, LPOPERX, // 20
	LPOPERX, LPOPERX, LPOPERX, LPOPERX, // 24
	LPOPERX, LPOPERX, LPOPERX, LPOPERX, // 28
	LPOPERX, LPOPERX                  // 30
)
{
#pragma XLLEXPORT
	static OPERX r = ErrValue;

	try {
		OPERX** ppa = &pa;

		int i, j = 1;
		for (i = 1; i < 32; ++i) {
			if (ppa[i]->xltype == xltypeMissing) {
				break;
			}
			OPER ai, vi; // name, value
			const OPER& pai = *ppa[i];
			if (pai.size() == 1) {
				ai = Excel(xlfText, OPER(j), OPER("\"_\"0"));
				vi = pai;
				++j;
			}
			else if (pai.size() == 2) {
				ensure(pai[0].xltype == xltypeStr);
				ai = pai[0];
				vi = pai[1];
			}
			else {
				XLL_ERROR("XLL.EVAL: args must be 1 or 2 diminsional");

				return &r;
			}
			Excel(xlfSetName, ai, vi);
			//r = Excel(xlGetName, ai);
		}
		int ret;
		ret = xll::traits<XLOPERX>::Excelv(xlfEvaluate, (LPOPERX)&r, 1, (LPOPERX*)ppa);
		for (i = 1; i < 32; ++i) {
			if (ppa[i]->xltype == xltypeMissing) {
				break;
			}
			OPER ai; // name
			const OPER& pai = *ppa[i];
			if (pai.size() == 1) {
				ai = Excel(xlfText, OPER(j), OPER("\"_\"0"));
				++j;
			}
			else if (pai.size() == 2) {
				ensure(pai[0].xltype == xltypeStr);
				ai = pai[0];
			}
			else {
				XLL_ERROR("XLL.EVAL: args must be 1 or 2 diminsional");

				return &r;
			}
			Excel(xlfSetName, ai);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &r;
}

