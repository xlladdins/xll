// spreadsheet.cpp
#include "splitpath.h"
#include "xll.h"
#include "spreadsheet.h"

using namespace xll;

// create sample spreadsheet
bool Spreadsheet(const char* description = "", bool release = false)
{
	// get dir and filename extension
	splitpath sp(Excel4(xlGetName).to_string().c_str());
	OPER dir(release ? ADDIN_URL : sp.dirname().c_str());

	OPER addin = OPER(sp.fname) & OPER(sp.ext);
	OPER Name(Excel(xlfUpper, OPER(sp.fname)));
	Rename(Name);

	// sort by category then function text/key
	std::map<OPER, std::map<OPER, Args*>> cat_key;
	for (auto& [key, args] : AddIn::Map) {
		if (args.isFunction() and args.Documentation().length() > 0) {
			cat_key[args.Category()][args.Key()] = &args;
		}
	}

	std::vector<std::pair<OPER, Args*>> tabs;
	for (auto& [cat, key] : cat_key) {
		for (auto& [text, args] : key) {
			tabs.push_back(std::pair(text, args));
		}
	}
	// reverse order since inserts happen on the left
	std::reverse(tabs.begin(), tabs.end());

	for (auto& [tab, args] : tabs) {
		OPER old = GetSheet();
		Excel(xlcWorkbookSelect, old);
		Excel(xlcWorkbookInsert, OPER(1)); // insert Worksheet
		Rename(tab);

		Excel(xlcSelect, OPER(REF(1, 1)));
		Excel(xlcFormula, tab);
		Excel(xlcFormatFont, XLL_H1);
		//Excel(xlcAlignment, XLL_ALIGN_CENTER);

		Move(1, 0);
		Excel(xlcFormula, args->FunctionHelp());

		Excel(xlcSelect, OPER(REF(4, 1)));
		Excel(xlcFormula, tab);
		Excel(xlcAlignment, XLL_ALIGN_RIGHT);

		Move(0, 1);
		OPER xFor = Excel(xlfActiveCell);
		Move(1, -1);
		for (int i = 0; i < args->ArgumentCount(); ++i) {
			Excel(xlcFormula, args->ArgumentName(i));
			Excel(xlcAlignment, XLL_ALIGN_RIGHT);
			Excel(xlcFontProperties, OPER(XLL_FONT), OPER("Bold"));
			Move(0, 1);
			Excel(xlcFormula, args->ArgumentDefault(i));
			// Define name only for this sheet.
			DefineName(args->ArgumentName(i), true);
			Move(0, 1);
			Excel(xlcFormula, args->ArgumentHelp(i));
			Move(1, -2);
		}
		Excel(xlcSelect, xFor);
		Excel(xlcFormula, OPER("=") & args->FunctionText() & OPER("(") & args->ArgumentText() & OPER(")"));
		//Excel(xlcWorkbookMove, text); // move to last tab
		//old = Excel(xlfGetDocument, OPER(76)); // [Book]Sheet
		//Excel(xlcAlert, old, OPER(3));
	}

	OPER book = GetBook();
	Excel(xlcWorkbookMove, Name, book, OPER(1));
	Excel(xlcWorkbookSelect, Name);

	// "=HYPERLINK(\"https://xlladdins.com/addins/xll_math.xll\", \"xll_math\")"));
	Excel(xlcSelect, OPER(REF(1, 1)));
	Excel(xlcFormula, OPER("=HYPERLINK(\"") & dir & addin & OPER("\", \"") & Name & OPER("\")"));
	Excel(xlcFormatFont, XLL_H1);
	Excel(xlcAlignment, XLL_ALIGN_RIGHT);

	Excel(xlcSelect, OPER(REF(1, 2)));
	Excel(xlcFormula, OPER(description));

	Excel(xlcSelect, OPER(REF(3, 1)));

	for (auto& [cat, key] : cat_key) {
		Excel(xlcFormula, OPER("Category ") & cat);
		Excel(xlcFormatFont, XLL_H2);
		Move(2, 0);
		for (auto& [text, args] : key) {
			Excel(xlcFormula, text);
			Excel(xlcFormatFont, XLL_ARG);
			Excel(xlcAlignment, XLL_ALIGN_RIGHT);
			Move(0, 1);
			Excel(xlcFormula, args->FunctionHelp());
			Move(1, -1);
		}
		Move(1, 0);
	}

	// select cell containing the hyperlink
	Excel(xlcSelect, OPER(REF(1, 1)));

	return true;
}

static AddIn xai_spreadsheet(
	Macro(XLL_DECORATE("_xll_spreadsheet", 0), "DOC")
	.Category("XLL")
);
extern "C" int __declspec(dllexport) WINAPI 
xll_spreadsheet(void)
{
#pragma XLLEXPORT
	int result = FALSE;

	try {
		result = Spreadsheet(R"(All files and macros having documentation.)", true);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR("unknown exception in xll::Spreadsheet");
	}

	return result;
}
