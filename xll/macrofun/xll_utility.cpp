// utility.cpp - generally useful functions
#include "../xll.h"

using namespace xll;

AddIn xai_range_set(
	Function(XLL_HANDLEX, "xll_range_set_", "\\RANGE")
	.Arguments({
		Arg(XLL_LPOPER, "range", "is the range to set.", "{1,2;3,4}")
		})
	.Uncalced()
	.FunctionHelp("Return a handle to a range.")
	.Category("XLL")
	.Documentation(R"(
Create a handle to a two dimensional range of cells.
)")
);
HANDLEX WINAPI xll_range_set_(LPOPER px)
{
#pragma XLLEXPORT
	handle<OPER> h(new OPER(*px));

	return h.get();
}

AddIn xai_range_get(
	Function(XLL_LPOPER, "xll_range_get", "RANGE.GET")
	.Arguments({
		Arg(XLL_HANDLEX, "handle", "is a handle returned by RANGE.SET.", "\\RANGE.SET({0,1;2,3})")
		})
	.FunctionHelp("Return the range held by a handle.")
	.Category("XLL")
	.Documentation(R"(
Return a two dimensional range of cells.
)")
);
LPOPER WINAPI xll_range_get(HANDLEX h)
{
#pragma XLLEXPORT
	handle<OPER> h_(h);

	if (!h_) {
		XLL_ERROR("RANGE.GET: unknown handle");

		return nullptr;
	}

	return h_.ptr();
}

static constexpr size_t N = 30;
static const auto& args = []() {
	static Arg arg[N];

	static const char* type = XLL_LPOPER;
	static const char* name = "cell";
	static const char* help = "is a cell to collect.";

	for (unsigned i = 0; i < N; ++i) {
		arg[i] = Arg(type, name, help);
	}

	return arg;
}();

AddIn xai_collect(
	Function(XLL_LPOPER, "xll_collect", "COLLECT")
	.Arguments(std::initializer_list(args, args + N))
	.Uncalced()
	.Category("XLL")
	.FunctionHelp("Collect cells into a range")
	.Documentation(R"(
Assemble up to 30 spreadsheet cells into a range. If all arguments to
<code>COLLECT</code> have the same number of columns then the
ranges are stacked.
)")
);
LPOPER WINAPI xll_collect(LPOPER po,
	LPOPER, LPOPER, LPOPER, LPOPER, LPOPER, LPOPER,
	LPOPER, LPOPER, LPOPER, LPOPER, LPOPER, LPOPER,
	LPOPER, LPOPER, LPOPER, LPOPER, LPOPER, LPOPER,
	LPOPER, LPOPER, LPOPER, LPOPER, LPOPER, LPOPER,
	LPOPER, LPOPER, LPOPER, LPOPER, LPOPER, LPOPER)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		LPOPER* p = &po;
		o = **p;
		++p;
		for (size_t n = 1; n < N and !(*p)->is_missing(); ++n, ++p) {
			o.push_back(**p);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
		o = ErrNA;
	}
	catch (...) {
		XLL_ERROR("COLLECT: unknown exception");
		o = ErrNA;
	}

	return &o;
}

AddIn xai_depends(
	Function(XLL_LPOPER, "xll_depends", "DEPENDS")
	.Arguments({
		Arg(XLL_LPXLOPER, "cell", "is the cell to return.", "A1"),
		Arg(XLL_LPXLOPER, "dependent", "is the dependent cell.", "A2"),
		})
		.FunctionHelp("Return cell after dependent is calculated.")
	.Category("XLL")
	.Documentation(R"(
The <code>DEPENDS</code> function is used to force the calculation order of cells.
The first argument <code>cell</code> will be calculated after the <code>dependent</code>
cell is evaluated. This is useful when using handles since some functions return
unchanged handles so Excel cannot keep track of dependencies in the usual way.
)")
);
LPXLOPER WINAPI xll_depends(LPXLOPER cell, LPXLOPER)
{
#pragma XLLEXPORT
	return cell;
}

AddIn xai_replace_eq_by_eq(
	Macro("xll_replace_eq_by_eq", "REPLACE.EQ.BY.EQ")
	.FunctionHelp("Recalculate all formulas in a workbook. (Ctrl-Shift-F9)")
	.Category("XLL")
	.Documentation(R"(
Calls Find and Replace (<code>Ctrl-h</code>) to replace '=' by '=' in all formulas.
This causes Excel to recalculate all formulas in the workbook.
)")
);
int WINAPI xll_replace_eq_by_eq()
{
#pragma XLLEXPORT
	static OPER eq("=");

	return Excel(xlcFormulaReplace, eq, eq) == true;
}
// Bind Ctrl-Shift-F9 to replace = by =
On<Key> xok_replace_eq_by_eq(ON_CTRL ON_SHIFT ON_F9, "REPLACE.EQ.BY.EQ");

static AddIn xai_this(
	Function(XLL_LPXLOPER, "xll_this", "THIS")
	.Uncalced()
	.Category("XLL")
	.FunctionHelp("Return the contents of the calling cell.")
	.Documentation(
		"The contents are the previously calculated value for the cell."
	)
);
LPXLOPERX WINAPI xll_this(void)
{
#pragma XLLEXPORT
	static XLOPERX x;

	x = Excel(xlCoerce, Excel(xlfCaller));

	return &x;
}