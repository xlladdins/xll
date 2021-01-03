// paste.cpp - Paste arguments into Excel given function name or register id.
// If it is a number, just paste the default arguments and do not define names.
// If it is a string, use that as a rdb prefix for defined names.

#include "xll.h"

using namespace xll;
using xll::Excel;

typedef traits<XLOPERX>::xword xword;

static AddIn xai_range_set(
	Function(XLL_HANDLEX, "xll_range_set", "RANGE.SET")
	.Args({
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
	.Args({
		Arg(XLL_HANDLEX, "handle", "is a handle returned by RANGE.SET.")
		})
	.FunctionHelp("Return the range held by a handle.")
	.Category("XLL")
	.Documentation(R"(Return a two dimensional range of cells)")
);
LPOPER WINAPI xll_range_get(HANDLEX h)
{
#pragma XLLEXPORT
	handle<OPER> h_(h);

	if (!h_) {
		XLL_WARNING("RANGE.GET: unknown handle");
	}

	return h_.ptr();
}

inline void Move(short r, short c)
{
	Excel(xlcSelect, Excel(xlfOffset, Excel(xlfActiveCell), OPER(r), OPER(c)));
}

inline void
ApplyStyle(const char* style)
{
	Excel(xlcDefineStyle, OPER(style));
	Excel(xlcApplyStyle, OPER(style));
}

// paste argument into xActi, return reference for formula
template<class X>
inline XOPER<X> paste_default(const XOPER<X>& xAct, const XOPER<X>& xActi, const XOPER<X>& xDef)
{
	XOPER<X> xRel = XExcel<X>(xlfRelref, xActi, xAct);

	if (xDef && xDef.xltype == xltypeStr && xDef.val.str[0] > 0 && xDef.val.str[1] == '=') {
		XOPER<X> xEval = XExcel<X>(xlfEvaluate, xDef);
		XOPER<X> xSize = XExcel<X>(xlfLen, xEval);
		if (xSize.val.num > 1) {
			XOPER<X> xOff = XExcel<X>(xlfOffset, xActi, XOPER<X>(0), XOPER<X>(0), XOPER<X>(1), xSize);
			XExcel<X>(xlSet, xOff, xEval);

			xRel = XExcel<X>(xlfRelref, xOff, xAct);
		}
		else {
			XExcel<X>(xlcFormula, xDef);
		}
	}
	else {
		XExcel<X>(xlSet, xActi, xDef);
	}

	return xRel;
}

template<class X>
void xll_paste_regid(const XArgs<X>& args)
{
	XOPER<X> xAct = XExcel<X>(xlfActiveCell);

	XOPER<X> xFor = XOPER<X>("=") & args.FunctionText() & XOPER<X>("(");

	for (unsigned short i = 0; i < args.ArgumentCount(); ++i) {
		Move(1,0);

		XOPER<X> xActi = XExcel<X>(xlfActiveCell);
		XOPER<X> xDef = args.ArgumentDefault(i);
		XOPER<X> xRel = paste_default(xAct, xActi, xDef);

		if (i > 0) {
			xFor.append(XOPER<X>(", "));
		}
		xFor.append(xRel);
	}
	xFor.append(XOPER<X>(")"));

	XExcel<X>(xlcSelect, xAct);
	XExcel<X>(xlcFormula, xFor);
}

void xll_paste_regidx(void)
{
	double regid = Excel(xlCoerce, Excel(xlfActiveCell)).val.num;

	const Args4& arg4 = AddIn4::Args(regid);
	if (arg4.isFunction()) {
		xll_paste_regid(arg4);

		return;
	}

	const Args12& arg12 = AddIn12::Args(regid);
	if (arg12.isFunction()) {
		xll_paste_regid(arg12);

		return;
	}

	throw std::runtime_error("XLL.PASTE.FUNCTION: register id not found");
}

template<class X>
void xll_paste_name(const XArgs<X>& args)
{
	XOPER<X> xAct = XExcel<X>(xlfActiveCell);
	XOPER<X> xPre = XExcel<X>(xlCoerce, xAct);

	XOPER<X> xFor = XOPER<X>("=") & args.FunctionText() & XOPER<X>("(");

	for (unsigned short i = 0; i < args.ArgumentCount(); ++i) {
		Move(1,0);

		XExcel<X>(xlcFormula, args.ArgumentName(i));
		XOPER<X> xNamei = xPre & args.ArgumentName(i);

		Move(0,1);
		XOPER<X> xActi = XExcel<X>(xlfActiveCell);
		XOPER<X> xDef = args.ArgumentDefault(i);
		XOPER<X> xRel = paste_default(xAct, xActi, xDef);
		XExcel<X>(xlcDefineName, xNamei, XExcel<X>(xlfAbsref, xRel, xAct));
		Move(0, -1);

		if (i > 0) {
			xFor.append(", ");
		}
		xFor.append(xNamei);
	}
	xFor.append(XOPER<X>(")"));

	XExcel<X>(xlcSelect, xAct);
	Move(0, -1);
	XExcel<X>(xlcFormula, xFor);
	Move(0, 1);
}

void xll_paste_namex(void)
{
	double regid = Excel(xlCoerce, Excel(xlfOffset, Excel(xlfActiveCell), OPER(0), OPER(1))).val.num;

	const Args4& arg4 = AddIn4::Args(regid);
	if (arg4.isFunction()) {
		xll_paste_name(arg4);
		
		return;
	}

	const Args12& arg12 = AddIn12::Args(regid);
	if (arg12.isFunction()) {
		xll_paste_name(arg12);

		return;
	}

	throw std::runtime_error("XLL.PASTE.FUNCTION: register id not found");
}
/*
static AddIn xai_xll_paste(
	MacroX("_xll_paste@0", "XLL.PASTE.FUNCTION")
	.Category(CATEGORY)
	.FunctionHelp("Paste a function into Excel. ")
	.Documentation(
		"When the active cell is a register id, paste the default arguments below the active cell "
		"and replace the current cell with a call to the function using relative arguments. "
		"When the active cell is a string, look for the register id or a call to the correponding function "
		"in the cell to its right, paste the argument names below the active cell, and the default values "
		"to their right and name them using the active cell contents as a prefix. "
		"Replace the register id/function call with a function call using the named ranges. "
	)
);
extern "C" int __declspec(dllexport) WINAPI
xll_paste(void)
{
	int result(0);

	try {
		Excel(xlcEcho, XOPER<X>(false));

		XOPER<X> oAct = Excel(xlCoerce, Excel(xlfActiveCell));

		if (oAct.xltype == xltypeNum)
			xll_paste_regidx();
		else if (oAct.xltype == xltypeStr)
			xll_paste_namex();
		else
			throw std::runtime_error("XLL.PASTE.FUNCTION: Active cell must be a number or a string. ");

		Excel(xlcEcho, XOPER<X>(true));

	}
	catch (const std::exception& ex) {
		Excel(xlcEcho, XOPER<X>(true));
		XLL_ERROR(ex.what());

		return 0;
	}

	return result;
}

int
xll_paste_close(void)
{
	try {
		if (Excel(xlfGetBar, XOPER<X>(7), XOPER<X>(4), XOPER<X>("Paste Function")))
			Excel(xlfDeleteCommand, XOPER<X>(7), XOPER<X>(4), XOPER<X>("Paste Function"));
	}
	catch (const std::exception& ex) {
		XLL_INFO(ex.what());

		return 0;
	}

	return 1;
}
static Auto<Close> xac_paste(xll_paste_close);

int
xll_paste_open(void)
{
	try {
		// Try to add just above first menu separator.
		XOPER<X> oPos;
		oPos = Excel(xlfGetBar, XOPER<X>(7), XOPER<X>(4), XOPER<X>("-"));
		oPos = 5;

		XOPER<X> oAdj = Excel(xlfGetBar, XOPER<X>(7), XOPER<X>(4), XOPER<X>("Paste Function"));
		if (oAdj == Err(xlerrNA)) {
			XOPER<X> oAdj(1, 5);
			oAdj(0, 0) = "Paste Function";
			oAdj(0, 1) = "XLL.PASTE.FUNCTION";
			oAdj(0, 3) = "Paste function under cursor into spreadsheet.";
			Excel(xlfAddCommand, XOPER<X>(7), XOPER<X>(4), oAdj, oPos);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}
static Auto<Open> xao_paste(xll_paste_open);
*/


/*
static AddInX xai_paste_basic(
	MacroX("_xll_paste_basic@0"), _T("XLL.PASTE.BASIC")
	.Category("Utility")
	.FunctionHelp("Paste a function with default arguments. Shortcut Ctrl-Shift-B.")
	.Documentation("Does not define names. ")
);
*/

static AddIn xai_paste_basic(
	Macro(XLL_DECORATE("_xll_paste_basic", 0), "XLL.PASTE.BASIC")
	.FunctionHelp("Paste a function with default arguments. Shortcut Ctrl-Shift-B.")
	.Category("XLL")
	//.ShortcutText("^+B")
	.Documentation("Shortcut Ctrl-Shift-B. Does not define names.")
);
extern "C" int __declspec(dllexport) WINAPI
xll_paste_basic(void)
{
	int result = TRUE;

	Excel(xlfEcho, OPER(false));
	try {
		xll_paste_regidx();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		result = FALSE;
	}
	Excel(xlfEcho, OPER(true));

	return result;
}
// Ctrl-Shift-B
Auto<Open> xaoa_paste_basic([]() {
	try {
		On<Key> xok_paste_basic(ON_CTRL ON_SHIFT "B", "XLL.PASTE.BASIC");
	}
	catch (...) {
		XLL_ERROR("On<Key> failed to assign Ctrl-Shift-B to XLL.PASTE.BASIC");

		return FALSE;
	}

	return TRUE;
});

// create named ranges for arguments
/*
static AddInX xai_paste_create(
	MacroX("_xll_paste_create@0"), _T("XLL.PASTE.CREATE")
	.Category("Utility")
	.FunctionHelp("Paste a function and create named ranges for arguments. Shortcut Ctrl-Shift-C")
	.Documentation("Uses current cell as a prefix. ")
);
*/
static AddIn xai_paste_create(Macro(XLL_DECORATE("_xll_paste_create", 0), "XLL.PASTE.CREATE"));
extern "C" int __declspec(dllexport) WINAPI
xll_paste_create(void)
{
	int result = TRUE;

	Excel(xlfEcho, OPER(false));
	try {
		OPER xAct = Excel(xlfActiveCell);
		OPER xFor = Excel(xlfGetCell, OPER(6), xAct); // formula
		ensure(xFor.is_str() and xFor.val.str[1] == '=');

		// extract "=Function("
		OPER xFind = Excel(xlfFind, OPER("("), xFor);
		if (xFind.is_num()) {
			xFor = Excel(xlfLeft, xFor, OPER(xFind.val.num - 1));
		}

		// get regid
		double regid = Excel(xlfEvaluate, xFor).val.num;

		const Args& args = AddIn::Args(regid);

		if (!args.isFunction()) {
			XLL_WARNING("XLL.PASTE.CREATE: could not find register id of function");

			return 0;
		}

		xFor &= OPER("(");

		for (unsigned short i = 0; i < args.ArgumentCount(); ++i) {
			Move(1, 0);

			OPER xNamei = args.ArgumentName(i);
			Excel(xlcFormula, xNamei);
			Excel(xlcAlignment, OPER(4)); // align right

			Move(0, 1);

			// paste default argument
			OPER xEval = Excel(xlfEvaluate, args.ArgumentDefault(i));
			if (xEval.size() > 1) {
				OPER xFor2 = Excel(xlfConcatenate,
					OPER("=RANGE.SET("),
					args.ArgumentDefault(i),
					OPER(")"));
				Excel(xlcFormula, xFor2);
				xNamei = OPER("RANGE.GET(") & xNamei & OPER(")");
			}
			else {
				Excel(xlcFormula, xEval);
			}
			Excel(xlcDefineName, xNamei);

			// style
			if (xNamei.val.str[1] == '[') {
				ApplyStyle("Optional");
			}
			else {
				ApplyStyle("Input");
			}

			if (i > 0) {
				xFor &= OPER(", ");
			}
			xFor &= xNamei;

			Move(0, -1);
		}
		xFor &= OPER(")");

		Excel(xlcSelect, xAct);
		Move(0, 1);
		Excel(xlcFormula, xFor);
		ApplyStyle("Output");
		Move(0, -1);

		Excel(xlcFormula, args.FunctionText());
		Excel(xlcAlignment, OPER(4)); // align right

		// select range for RDB.DEFINE???
		// Excel(xlcSelect, Excel(xlfOffset, xAct, OPER(0), OPER(0), OPER(args.ArgumentCount() + 1), OPER(2)));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		result = FALSE;
	}

	Excel(xlfEcho, OPER(true));

	return result;
}
// Ctrl-Shift-C
static On<Key> xok_paste_create("^+C", "XLL.PASTE.CREATE");
