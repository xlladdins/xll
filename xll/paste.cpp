// paste.cpp - Paste arguments into Excel given function name or register id.
// If it is a number, just paste the default arguments and do not define names.
// If it is a string, use that as a rdb prefix for defined names.

#include "xll.h"

using namespace xll;
using xll::Excel;

typedef traits<XLOPERX>::xword xword;

template<class X> const XOPER<X> up = XOPER<X>("R[-1]C[0]");
template<class X> const XOPER<X> down = XOPER<X>("R[1]C[0]");
template<class X> const XOPER<X> right = XOPER<X>("R[0]C[1]");
template<class X> const XOPER<X> left = XOPER<X>("R[0]C[-1]");
template<class X> const XOPER<X> alert = XOPER<X>("Would you like to replace the existing definition of ");
template<class X> const XOPER<X> input = XOPER<X>("Enter the range name.");
template<class X> const XOPER<X> r_ = XOPER<X>("R[");
template<class X> const XOPER<X> c0 = XOPER<X>("]C[0]");
template<class X> const XOPER<X> range_set = XOPER<X>("=RANGE.SET(");
template<class X> const XOPER<X> range_get = XOPER<X>("=RANGE.GET(");

#define UP XExcel<X>(xlcSelect, up<X>)
#define DOWN XExcel<X>(xlcSelect, down<X>)
#define RIGHT XExcel<X>(xlcSelect, right<X>)
#define LEFT XExcel<X>(xlcSelect, left<X>)

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

	for (unsigned short i = 1; i < args.ArgumentCount(); ++i) {
		DOWN;

		XOPER<X> xActi = XExcel<X>(xlfActiveCell);
		XOPER<X> xDef = args.ArgumentDefault(i);
		XOPER<X> xRel = paste_default(xAct, xActi, xDef);

		if (i > 1) {
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

	for (unsigned short i = 1; i < args.ArgumentCount(); ++i) {
		Move(1,0);

		XExcel<X>(xlcFormula, args.ArgumentName(i));
		XOPER<X> xNamei = XExcel<X>(xlfConcatenate, xPre, args.ArgumentName(i));

		RIGHT;
		XOPER<X> xActi = XExcel<X>(xlfActiveCell);
		XOPER<X> xDef = args.ArgumentDefault(i);
		XOPER<X> xRel = paste_default(xAct, xActi, xDef);
		XExcel<X>(xlcDefineName, xNamei, XExcel<X>(xlfAbsref, xRel, xAct));
		LEFT;

		if (i > 1) {
			xFor.append(XOPER<X>(","));
			xFor.append(XOPER<X>(" "));
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

#ifdef _M_X64
static AddIn4 xai_paste_basic(Macro4("xll_paste_basic", "XLL.PASTE.BASIC"));
#else
static AddIn12 xai_paste_basic(Macro12("_xll_paste_basic@0", "XLL.PASTE.BASIC"));
#endif
extern "C" int __declspec(dllexport) WINAPI
xll_paste_basic(void)
{
	Excel(xlfEcho, OPER(false));

	try {
		xll_paste_regidx();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}
	Excel(xlfEcho, OPER(true));

	return 1;
}
// Ctrl-Shift-B
static On<Key> xok_paste_basic("^+B", "XLL.PASTE.BASIC");

// create named ranges for arguments
/*
static AddInX xai_paste_create(
	MacroX("_xll_paste_create@0"), _T("XLL.PASTE.CREATE")
	.Category("Utility")
	.FunctionHelp("Paste a function and create named ranges for arguments. Shortcut Ctrl-Shift-C")
	.Documentation("Uses current cell as a prefix. ")
);
*/
#ifdef _M_X64
static AddIn xai_paste_create(Macro("xll_paste_create", "XLL.PASTE.CREATE"));
#else
static AddIn xai_paste_create(Macro("_xll_paste_create@0", "XLL.PASTE.CREATE"));
#endif
extern "C" int __declspec(dllexport) WINAPI
xll_paste_create(void)
{
	Excel(xlfEcho, OPER(false));

	try {
		OPER xAct = Excel(xlfActiveCell);
		OPER xPre = Excel(xlCoerce, xAct);

		// use cell to right if in cell containing handle
		if (xPre.xltype == xltypeNum) {
			Move(0, -1);
			xAct = Excel(xlfActiveCell);
			xPre = Excel(xlCoerce, xAct);
		}

		if (xPre.xltype == xltypeStr) {
			Excel(xlcAlignment, OPER(4)); // align right

			xPre = Excel(xlfConcatenate, xPre, OPER("."));
		}
		else {
			xPre = "";
		}

		OPER xFor = Excel(xlfGetCell, OPER(6), Excel(xlfOffset, xAct, OPER(0), OPER(1))); // formula
		ensure(xFor.xltype == xltypeStr);
		ensure(xFor.val.str[1] == '=');

		// extract "=Function"
		OPER xFind = Excel(xlfFind, OPER("("), xFor);
		if (xFind.xltype == xltypeNum)
			xFor = Excel(xlfLeft, xFor, OPER(xFind.val.num - 1));

		// get regid
		xFor = Excel(xlfEvaluate, xFor);
		ensure(xFor.xltype == xltypeNum);
		double regid = xFor.val.num;

		const Args& args = AddIn::Args(regid);

		if (!args.isFunction()) {
			XLL_WARNING("XLL.PASTE.CREATE: could not find register id of function");

			return 0;
		}

		xFor = Excel(xlfConcatenate, OPER("="), args.FunctionText(), OPER("("));

		for (unsigned short i = 1; i < args.ArgumentCount(); ++i) {
			Move(1, 0);

			Excel(xlcFormula, args.ArgumentName(i));
			Excel(xlcAlignment, OPER(4)); // align right
			OPER xNamei = Excel(xlfConcatenate, xPre, args.ArgumentName(i));

			Move(0, 1);
			Excel(xlcDefineName, xNamei);

			// paste default argument
			OPER xDef = args.ArgumentDefault(i);
			if (xDef.xltype == xltypeStr && xDef.val.str[1] == '=') {
				OPER xEval = Excel(xlfEvaluate, xDef);
				if (xEval.size() > 1) {
					OPER xFor2 = Excel(xlfConcatenate,
						OPER("=RANGE.SET("),
						OPER(xDef.val.str + 2, xDef.val.str[0] - 1),
						OPER(")"));
					Excel(xlcFormula, xFor2);
					xNamei = OPER("RANGE.GET(") & xNamei & OPER(")");
				}
				else {
					Excel(xlcFormula, xDef);
				}
			}
			else {
				Excel(xlcFormula, xDef);
			}
			// style
			if (args.ArgumentName(i).val.str[1] == '[') {
				ApplyStyle("Optional");
			}
			else {
				ApplyStyle("Output");
			}

			if (i > 1) {
				xFor = Excel(xlfConcatenate, xFor, OPER(", "));
			}
			xFor = Excel(xlfConcatenate, xFor, xNamei);

			Move(0, -1);
		}
		xFor = Excel(xlfConcatenate, xFor, OPER(")"));

		Excel(xlcSelect, xAct);
		Move(0, 1);
		Excel(xlcFormula, xFor);
		Move(0, -1);

		Excel(xlcSelect, Excel(xlfOffset, xAct, OPER(0), OPER(0), OPER(args.ArgumentCount() + 1), OPER(2))); // select range for RDB.DEFINE
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	Excel(xlfEcho, OPER(true));

	return 1;
}
// Ctrl-Shift-C
static On<Key> xok_paste_create("^+C", "XLL.PASTE.CREATE");
