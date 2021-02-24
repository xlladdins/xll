// paste.cpp - Paste arguments into Excel given function name or register id.
// If it is a number, just paste the default arguments and do not define names.
// If it is a string, use that as a rdb prefix for defined names.
#include "../splitpath.h"
#include "xll_paste.h"

using namespace xll;
using xll::Excel;

// Lookup Args using name or register id.
inline Args* xll_arguments(const OPER& xRef = Excel(xlfActiveCell))
{
	Args* pargs = nullptr;

	OPER xCaller = Excel(xlCoerce, xRef);
	if (xCaller.is_str()) {
		pargs = AddIn::Arguments(xCaller);
	}
	else if (xCaller.is_num()) {
		pargs = AddIn::Arguments(xCaller.as_num());
	}
	else {
		XLL_ERROR("xll_arguments: cell must contain string or number");
	}

	return pargs;
}

AddIn xai_paste_args(
	Macro(XLL_DECORATE("_xll_paste_args", 0), "XLL.PASTE.ARGS")
	.FunctionHelp("Paste function using default arguments.")
	.Category("XLL")
	//.ShortcutText(ON_CTRL ON_SHIFT "2")
	.Documentation(R"(
This macro works like <c>Ctrl-Shift-A</c> except it pastes argument defaults
instead of argument names.
)")
);
extern "C" __declspec(dllexport) int WINAPI xll_paste_args()
{
	int result = FALSE;

	try {
		const OPER& xRef = Excel(xlfActiveCell);
		Args* pargs = xll_arguments(xRef);
		ensure(pargs || !"XLL.PASTE.ARGS: name or register id not found");
		const OPER& xDef0 = pargs->ArgumentDefault(0);
		ensure(Excel(xlcFormula, xDef0, xRef)); // FormulaArray???
		result = TRUE;
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR("XLL.PASTE.ARGS: unknown exception");
	}

	return result;
}
// Ctrl-Alt-@ is like Ctrl-Shift-A but paste default values.
On<Key> xok_paste_args(ON_CTRL ON_SHIFT "2", "XLL.PASTE.ARGS");
/*
Auto<OpenAfter> xaoa_paste_args([]() {
	On<Key> xok_paste_args(ON_CTRL ON_SHIFT "2", "XLL.PASTE.ARGS");
	return TRUE;
	});
*/

#if 0

template<class X>
void xll_paste_name(const XArgs& args)
{
	Select xAct;
	OPER xFor = OPER("=") & args->FunctionText() & OPER("(");

	for (unsigned short i = 1; i <= args->ArgumentCount(); ++i) {
		Move(1, 0);
		OPER xNamei = args->ArgumentName(i);
		Excel(xlcFormula, xNamei);

		Move(0, 1);
		Name namei(xNamei);
		namei.Define(Select::Special(Select::Type::Current), true);

		Move(0, -1);
		if (i > 1) {
			xFor.append(", ");
		}
		xFor.append(xNamei);
	}
	xFor.append(OPER(")"));

	Excel(xlcSelect, xAct.selection);
	Move(0, -1);
	Excel(xlcFormula, xFor);
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
		Excel(xlcEcho, OPER(false));

		OPER oAct = Excel(xlCoerce, Excel(xlfActiveCell));

		if (oAct.xltype == xltypeNum)
			xll_paste_regidx();
		else if (oAct.xltype == xltypeStr)
			xll_paste_namex();
		else
			throw std::runtime_error("XLL.PASTE.FUNCTION: Active cell must be a number or a string. ");

		Excel(xlcEcho, OPER(true));

	}
	catch (const std::exception& ex) {
		Excel(xlcEcho, OPER(true));
		XLL_ERROR(ex.what());

		return 0;
	}

	return result;
}

int
xll_paste_close(void)
{
	try {
		if (Excel(xlfGetBar, OPER(7), OPER(4), OPER("Paste Function")))
			Excel(xlfDeleteCommand, OPER(7), OPER(4), OPER("Paste Function"));
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
		OPER oPos;
		oPos = Excel(xlfGetBar, OPER(7), OPER(4), OPER("-"));
		oPos = 5;

		OPER oAdj = Excel(xlfGetBar, OPER(7), OPER(4), OPER("Paste Function"));
		if (oAdj == Err(xlerrNA)) {
			OPER oAdj(1, 5);
			oAdj(0, 0) = "Paste Function";
			oAdj(0, 1) = "XLL.PASTE.FUNCTION";
			oAdj(0, 3) = "Paste function under cursor into spreadsheet.";
			Excel(xlfAddCommand, OPER(7), OPER(4), oAdj, oPos);
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

#endif // 0

// paste basic function call in active cell and arguments below
void xll_paste_regid(const Args* args)
{
	// call with defaults arguments to get size of output
	OPER xDef0 = args->ArgumentDefault(0);
	OPER xVal = Excel(xlfEvaluate, xDef0);
	if (!xVal) {
		OPER xErr = OPER("Failed to evaluate: ") & xDef0;
		XLL_ERROR(xErr.to_string().c_str());
	}
	
	OPER xAct = Excel(xlfActiveCell);
	OPER xFor = OPER("=") & args->FunctionText() & OPER("(");

	Down(rows(xVal));
	for (unsigned short i = 1; i <= args->ArgumentCount(); ++i) {
		OPER xRel = paste_default(args->ArgumentDefault(i));
		/*
		OPER xDefi = args->ArgumentDefault(i);
		ensure(xDefi.is_str());
		if (xDefi.val.str[0] != 0 and xDefi.val.str[1] == '{') {
			xDefi = Excel(xlfEvaluate, xDefi);
		}
		OPER xRel = Excel(xlcFormula, xDefi, Excel(xlfActiveCell));
		ensure(xRel);
		xRel = Select::Special(Select::Type::Current);
		ensure(xRel);
		xRel = Excel(xlfActiveCell);
		*/
		if (i > 1) {
			xFor.append(OPER(", "));
		}
		xFor.append(Excel(xlfReftext, xRel));
		Down(rows(xRel));
	}

	xFor.append(OPER(")"));
	Excel(xlcSelect, xAct);
	OPER xRef = Excel(xlfOffset, xAct,
		OPER(0), OPER(0),
		Excel(xlfRows, xVal), Excel(xlfColumns, xVal)
	);
	if (size(xRef) == 1) {
		Excel(xlcFormula, xFor, xRef);
	}
	else {
		Excel(xlcFormulaArray, xFor, xRef);
	}
}

void xll_paste_regidx(void)
{
	double regid = Excel(xlCoerce, Excel(xlfActiveCell)).val.num;

	const Args* args = AddIn::Arguments(regid);
	if (args) {
		xll_paste_regid(args);
	}
	else {
		throw std::runtime_error("XLL.PASTE.BASIC: register id not found");
	}
}
static AddIn xai_paste_basic(
	Macro(XLL_DECORATE("_xll_paste_basic", 0), "XLL.PASTE.BASIC")
	.FunctionHelp("Paste a function with default arguments. Shortcut Ctrl-Shift-B.")
	.Category("XLL")
	//.ShortcutText("^+B")
	.Documentation(R"xyzyx(
Paste a basic call to an add-in function.
<p>
Excel's built-in Ctrl-Shift-A shortcut pastes the function text after typing <code> =FUNCTION</code>
while in edit mode.
This macro requires you to first enter <code> =FUNCTION&lt;Enter&gt;</code> in the cell to get the
<a href="https://docs.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1#property-valuereturn-value">register id</a>
of the function. Ctrl-Shift-B uses the register id to look up the function, paste the default
arguments in the cells below, and replace the register id with a call to the function using the default arguments. 
</p>
)xyzyx")
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
On<Key> xok_paste_basic(ON_CTRL ON_SHIFT "B", "XLL.PASTE.BASIC");

#if 0
// create named ranges for arguments
static AddIn xai_paste_create(
	Macro(XLL_DECORATE("_xll_paste_create", 0), "XLL.PASTE.CREATE")
	.Category("XLL")
	.FunctionHelp("Paste a function and create named ranges for arguments. Shortcut Ctrl-Shift-C")
	.Documentation(R"xyzyx(Paste a call to an add-in function and create argument names.
<p>
Excel's built-in Ctrl-Shift-A shortcut pastes the function text after typing <code> =FUNCTION</code>
while in edit mode.
This macro requires you to first enter <code> =FUNCTION&lt;Enter&gt;</code> in the cell to get the
<a href="https://docs.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1#property-valuereturn-value">register id</a>
of the function. Ctrl-Shift-C uses this to look up the function and pastes the default
argument names in the cells below and default values to the immediate right.
The default values are named then the register id is replaced by the function name
and the cell to the immediate right has a call to the function using the default argument names. 
</p>
)xyzyx"));
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

		if (!args->isFunction()) {
			XLL_WARNING("XLL.PASTE.CREATE: could not find register id of function");

			return 0;
		}

		xFor &= OPER("(");

		for (unsigned i = 1; i <= args->ArgumentCount(); ++i) {
			Move<XLOPER>(1, 0);

			OPER xNamei = args->ArgumentName(i);
			Excel(xlcFormula, xNamei);
			Excel(xlcAlignment, OPER(4)); // align right

			Move<XLOPER>(0, 1);

			// paste default argument
			OPER xEval = Excel(xlfEvaluate, args->ArgumentDefault(i));
			if (xEval.size() > 1) {
				OPER xFor2 = Excel(xlfConcatenate,
					OPER("=RANGE.SET("),
					args->ArgumentDefault(i),
					OPER(")"));
				Excel(xlcFormula, xFor2);
				xNamei = OPER("RANGE.GET(") & xNamei & OPER(")");
			}
			else {
				Excel(xlcFormula, xEval);
			}
			Excel(xlcDefineName, xNamei);

			// style
			/*
			if (xNamei.val.str[1] == '[') {
				ApplyStyle("Optional");
			}
			else {
				ApplyStyle("Input");
			}
			*/
			if (i > 1) {
				xFor &= OPER(", ");
			}
			xFor &= xNamei;

			Move<XLOPER>(0, -1);
		}
		xFor &= OPER(")");

		Excel(xlcSelect, xAct);
		Move<XLOPER>(0, 1);
		Excel(xlcFormula, xFor);
		// ApplyStyle("Output");
		Move<XLOPER>(0, -1);

		Excel(xlcFormula, args->FunctionText());
		Excel(xlcAlignment, OPER(4)); // align right

		// select range for RDB.DEFINE???
		// Excel(xlcSelect, Excel(xlfOffset, xAct, OPER(0), OPER(0), OPER(args->ArgumentCount() + 1), OPER(2)));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		result = FALSE;
	}

	Excel(xlfEcho, OPER(true));

	return result;
}

// Ctrl-Shift-C
/*
Auto<OpenAfter> xaoa_paste_create([]() {
	try {
		On<Key> xok_paste_create(ON_CTRL ON_SHIFT "C", "XLL.PASTE.CREATE");
	}
	catch (...) {
		XLL_ERROR("On<Key> failed to assign Ctrl-Shift-C to XLL.PASTE.CREATE");

		return FALSE;
	}

	return TRUE;
});
*/
#endif // 0