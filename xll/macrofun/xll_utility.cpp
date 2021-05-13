// utility.cpp - generally useful functions
#include "../xll.h"

using namespace xll;

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
Assemble up to 30 spreadsheet cells into a range.
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
	.Documentation(R"(Force the calculation order of cells.)")
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
Calls Find and Replace (<code>Ctrl-H</code>) to replace '=' by '=' in all formulas.
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
	x.xltype |= xlbitXLFree;

	return &x;
}