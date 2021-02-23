// xll_range.cpp - two-dimensional range of cells
#include "xll.h"

using namespace xll;

static AddIn xai_range_set(
	Function(XLL_HANDLEX, "xll_range_set", "\\RANGE.SET")
	.Arguments({
		Arg(XLL_LPOPER, "range", "is the range to set.")
		})
	.Uncalced()
	.FunctionHelp("Return a handle to a range.")
	.Category("XLL")
	.Documentation(R"(Create a handle to a two dimensional range of cells)")
);
HANDLEX WINAPI xll_range_set(LPOPER px)
{
#pragma XLLEXPORT
	handle<OPER> h(new OPER(*px));

	return h.get();
}

static AddIn xai_range_get(
	Function(XLL_LPOPER, "xll_range_get", "RANGE.GET")
	.Arguments({
		Arg(XLL_HANDLEX, "handle", "is a handle returned by RANGE.SET.")
	})
	.FunctionHelp("Return the range held by a handle.")
	.Category("XLL")
	.Documentation(R"(Return a two dimensional range of cells)")
);
LPOPER WINAPI xll_range_get(HANDLEX h)
{
#pragma XLLEXPORT
	handle<OPER> h_(h, false);

	if (!h_) {
		XLL_WARNING("RANGE.GET: unknown handle");
	}

	return h_.ptr();
}
