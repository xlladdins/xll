// utility.cpp - generally useful functions
#include "xll.h"

using namespace xll;

static AddIn xai_replace_eq_by_eq(Macro(XLL_DECORATE("_xll_replace_eq_by_eq", 0), "REPLACE.EQ.BY.EQ"));
extern "C" __declspec(dllexport) int WINAPI _xll_replace_eq_by_eq()
{
	static OPER eq("=");

	return Excel(xlcFormulaReplace, eq, eq) == true;
}
// Bind Ctrl-Shift-F9 to replace = by =
On<Key> xok_replace_eq_by_eq(ON_CTRL ON_SHIFT ON_F9, "REPLACE.EQ.BY.EQ");


