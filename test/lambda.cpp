#include "../xll/xll.h"
using namespace xll;

struct lambda {
	OPER body;

	lambda(const OPER& body)
		: body(body)
	{ }
	OPER evaluate(LPXLOPERX* ppa)
	{
		int n = SetNames(ppa);
		OPER result = Excel(xlfEvaluate, body);
		UnsetNames(n);

		return result;
	}
private:
	static OPER arg(int i)
	{
		return Excel(xlfConcatenate, OPER("_"), OPER(i));
	}
	int SetNames(LPXLOPERX* ppa)
	{
		int i = 0;

		while((*ppa)->xltype != xltypeMissing) {
			Excel(xlfSetName, arg(i), **ppa);
			++ppa;
			++i;
		}

		return i;
	}
	void UnsetNames(int n)
	{
		for (int i = 0; i < n; ++i) {
			Excel(xlfSetName, arg(i));
		}
	}
};

AddIn xai_lambda(
	Function(XLL_LPOPER, "xll_lambda", "\\")
	.Arguments({
		Arg(XLL_LPOPER, "body", "is the body or a handle"),
		Arg(XLL_LPOPER, "_arg0", "is the optional first argument"),
		Arg(XLL_LPOPER, "_arg1", "is the optional second argument"),
		Arg(XLL_LPOPER, "_arg2", "is the optional third argument"),
		Arg(XLL_LPOPER, "_arg3", "is the optional fourth argument"),
		})
	.Uncalced()
	.FunctionHelp("Return handle to a lambda or evaluate when supplied with arguments.")
	.Category("XLL")
	.Documentation(R"xyzyx(
Call <code>\(body)<code> to create a handle to a function.
Call <code>\(lambda, args, ...)</code> to evaluate.
)xyzyx")
);
LPXLOPERX WINAPI xll_lambda(LPXLOPERX body, LPXLOPERX pa0, LPXLOPERX, LPXLOPERX, LPXLOPERX)
{
#pragma XLLEXPORT
	static OPER result;

	try {
		if (body->xltype == xltypeStr) {
			ensure(pa0->xltype == xltypeMissing);
			handle<lambda> h_(new lambda(*body));
			result = h_.get();
		}
		else if (body->xltype == xltypeNum) {
			handle<lambda> h_(body->val.num);
			ensure(h_);
			result = h_->evaluate(&pa0);
		}
		else {
			result = ErrNA;
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &result;
}

