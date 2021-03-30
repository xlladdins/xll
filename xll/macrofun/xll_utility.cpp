// utility.cpp - generally useful functions
#include "../xll.h"

using namespace xll;

AddIn xai_replace_eq_by_eq(
	Macro("xll_replace_eq_by_eq", "REPLACE.EQ.BY.EQ")
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
