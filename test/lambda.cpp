//
// Implementing lambda
// 
// h=\LAMBDA(arg, ..., body)
// =\(h, arg, ...)
//
#include "../xll/xll.h"
using namespace xll;

struct lambda {
	OPER body;
	OPER args;
	OPER pre;

	lambda(const OPER& body, LPXLOPERX* ppa)
		: body(body), pre("__") // <- change to unique string
	{
		while ((*ppa)->xltype != xltypeMissing) {
			args.push_back(**ppa);
			this->body = Excel(xlfSubstitute, this->body, **ppa, pre & OPER(**ppa));
			++ppa;
		}
	}
	OPER evaluate(LPXLOPERX pargs)
	{
		OPER result;

		SetNames(&pargs);
		result = Excel(xlfEvaluate, body);
		UnsetNames();

		return result;
	}
private:
	void SetNames(LPXLOPERX* arg)
	{
		for (unsigned i = 0; i < args.size(); ++i) {
			Excel(xlfSetName, pre & args[i], *arg[i]);
		}
	}
	void UnsetNames()
	{
		for (unsigned i = 0; i < args.size(); ++i) {
			Excel(xlfSetName, pre & args[i]);
		}
	}
};

// arg followed by n generic arguments
std::vector<Arg> args(const Arg& arg, unsigned n)
{
	std::vector<Arg> va(1, arg);

	for (unsigned i = 1; i <= n; ++i) {
		va.push_back(Arg(XLL_LPXLOPER, "arg", "is an arg."));
	}

	return va;
}

#define LPXLOPERX2 ,LPXLOPERX,LPXLOPERX
#define LPXLOPERX4 LPXLOPERX2 LPXLOPERX2
#define LPXLOPERX8 LPXLOPERX4 LPXLOPERX4
#define LPXLOPERX16 LPXLOPERX8 LPXLOPERX8

AddIn xai_lambda(
	Function(XLL_HANDLEX, "xll_lambda", "\\LAMBDA")
	.Arguments(args(Arg(XLL_CSTRING, "body", "is the function body."), 16))
	.Uncalced()
	.FunctionHelp("Return handle to function.")
	.Category("XLL")
	.Documentation(R"xyzyx(
Call <code>\LAMBDA(body, arg, ...)<code> to create a handle to a function.
)xyzyx")
);
HANDLEX WINAPI xll_lambda(LPCTSTR body, LPXLOPERX pa LPXLOPERX16)
{
#pragma XLLEXPORT
	HANDLEX h = INVALID_HANDLEX;

	try {
		handle<lambda> h_(new lambda(OPER(body), &pa));
		h = h_.get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
}

AddIn xai_lambda_eval(
	Function(XLL_LPOPER, "xll_lambda_eval", "\\")
	.Arguments(args(Arg(XLL_HANDLEX, "lambda", "is a handle."), 16))
	.Uncalced()
	.FunctionHelp("Evaluate a lambda.")
	.Category("XLL")
);
LPOPER WINAPI  xll_lambda_eval(HANDLEX h, LPXLOPERX pa LPXLOPERX16)
{
#pragma XLLEXPORT
	static OPER result;
	result = ErrNA;

	handle<lambda> h_(h);
	if (h_) {
		result = h_->evaluate(pa);
	}

	return &result;
}
#if 0

#define ARGMAX 30

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
	inline unsigned len(X** ppa)
	{
		auto missing = [](X* pa) { return pa->xltype == xltypeMissing; };
		auto i = std::find_if(ppa, ppa + ARGMAX, missing);

		return static_cast<unsigned>(i - ppa);
	}

	// 1.2, {"foo", "bar"}, true, ...
	// _1 = 1.2, foo = "bar", _2 = true, ...
	template<class X>
	inline unsigned SetNames(X** ppa)
	{
		//int j = 1;
		for (int i = 0, j = 1; i < ARGMAX; ++i) {
			XOPER<X> ki, vi; // key, value
			const XOPER<X>& ai = *ppa[i];

			if (ai.xltype == xltypeMissing) {
				return i;
			}

			if (ai.size() == 1) {
				ki = XExcel<X>(xlfText, XOPER<X>(j), XOPER<X>("\"_\"0"));
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

		return static_cast<unsigned>(-1);
	}
	template<class X>
	inline void UnsetNames(X** ppa, unsigned len)
	{
		for (unsigned i = 0, j = 1; i < len; ++i) {
			XOPER<X> ki, vi; // key, value
			const XOPER<X>& ai = *ppa[i];

			if (ai.size() == 1) {
				ki = XExcel<X>(xlfText, XOPER<X>(j), XOPER<X>("\"_\"0"));
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

std::vector args(ARGMAX, XLL_LPOPER);

AddIn xai_udf(
	Function(XLL_LPOPER, "xll_udf", "XLL.UDF")
	.Arguments({
		{XLL_LPXLOPER, "arg", "is an argument."},
		{XLL_LPXLOPER, "arg", "is an argument."},
		{XLL_LPXLOPER, "arg", "is an argument."},
		{XLL_LPXLOPER, "arg", "is an argument."},
		})
	.Uncalced()
	.FunctionHelp("Call user defined function.")
	.Category("XLL")
	/*
	.Documentation(R"(
Call a user defined function given its arguments.
Either the string or register id of the udf may be used.
	)")
	*/
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

AddIn xai_evaluate(
	Function(XLL_LPOPER, "xll_evaluate", "XLL.EVAL")
	.Arguments({
		{XLL_LPXLOPER, "arg", "is an argument."},
		{XLL_LPXLOPER, "arg", "is an argument."},
		{XLL_LPXLOPER, "arg", "is an argument."},
		{XLL_LPXLOPER, "arg", "is an argument."},
		})
	.Uncalced()
	.FunctionHelp("Evaluate a function.")
	.Category("XLL")
	/*
	.Documentation(R"(
The first argument is a string expression to evaluate.
Occurences of _1, _2, ... are replaced by the following arguments.
If the argument is a key-value pair the key is replaced by the value in the string expression.
To use a cell reference or existing name prepend an exclamation point.		
	)")
	*/
);
LPXLOPERX WINAPI xll_evaluate(LPXLOPERX pa LPXLOPERX32)
{
#pragma XLLEXPORT
	static XLOPERX r = ErrValue;

	try {
		XLOPERX** ppa = &pa;

		unsigned n = SetNames(ppa + 1);
		r = Excelv(xlfEvaluate, 1, ppa);
		UnsetNames(ppa + 1, n);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &r;
}

#endif // 0
