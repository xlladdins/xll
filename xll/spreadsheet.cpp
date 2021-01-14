// spreadsheet.cpp
#include <list>
#include "splitpath.h"
#include "xll.h"
#include "spreadsheet.h"

using namespace xll;

// create sample spreadsheet
bool Spreadsheet(const char* description = "", bool release = false)
{
	Colors();

	// sort by category, macro type, then function text
	std::map<OPER, std::map<OPER, std::map<OPER, Args*>>> cat_type_text;
	for (auto& [key, args] : AddIn::Map) {
		cat_type_text[args.Category()][args.Type()][args.FunctionText()] = &args;
	}

	// get dir and filename extension
	splitpath sp(Excel4(xlGetName).to_string().c_str());
	OPER dir(release ? ADDIN_URL : sp.dirname().c_str());

	OPER Name(Excel(xlfUpper, OPER(sp.fname)));
	Rename(Name);
	Header();

	std::list<std::pair<OPER, Args*>> tabs;
	for (auto& [cat, type_text] : cat_type_text) {
		for (auto& [type, text] : type_text) {
			for (auto& [key, args] : text) {
				tabs.push_front(std::pair(key, args)); // insert happens to the left
			}
		}
	}

	// insert sheets with sample call
	for (auto& [tab, args] : tabs) {
		OPER old = GetSheet();
		Excel(xlcWorkbookSelect, old);
		Excel(xlcWorkbookInsert, OPER(1)); // insert Worksheet
		Excel(xlcDisplay, OPER(false), OPER(false)); // gridlines off
		Rename(tab);
		Header();

		Excel(xlcSelect, OPER(REF(1, 1)));
		Excel(xlcFormula, tab & OPER(" ") & args->Type());

		Move(1, 0);
		Excel(xlcFormula, args->FunctionHelp());
		Excel(xlcFormatFont, XLL_HDR);

		Excel(xlcSelect, OPER(REF(4, 1)));
		Excel(xlcFormula, tab);
		Excel(xlcAlignment, XLL_ALIGN_RIGHT);

		Move(0, 1);
		OPER xFor = Excel(xlfActiveCell);
		Move(1, -1);
		if (args->isFunction()) {
			for (size_t i = 0; i < args->ArgumentCount(); ++i) {
				Excel(xlcFormula, args->ArgumentName(i));
				Excel(xlcAlignment, XLL_ALIGN_RIGHT);
				Bold();
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
		}
	}

	OPER book = GetBook();
	Excel(xlcWorkbookMove, Name, book, OPER(1)); // move main page to first position
	Excel(xlcWorkbookSelect, Name);

	// "=HYPERLINK(\"https://xlladdins.com/addins/xll_math.xll\", \"xll_math\")"));
	Excel(xlcSelect, OPER(REF(1, 1)));
	Excel(xlcFormula, OPER("=HYPERLINK(\"") & dir & OPER(sp.fname) & OPER(" & WinXX() & ") & OPER(sp.ext) & OPER("\", \"") & Name & OPER("\")"));
	Excel(xlcAlignment, XLL_ALIGN_RIGHT);
	//Excel(xlcNote, OPER("user: addinuser pass: @addm3"));

	Excel(xlcSelect, OPER(REF(1, 2)));
	Excel(xlcFormula, OPER(description));

	Excel(xlcSelect, OPER(REF(4, 1)));

	for (auto& [cat, type_text] : cat_type_text) {
		Category();
		Excel(xlcFormula, OPER("Category ") & cat);
		Move(1, 0);
		for (auto& [type, text] : type_text) {
			TableHeader();

			Excel(xlcFormula, Excel(xlfProper, type));
			Excel(xlcAlignment, HALIGN_RIGHT, OPER(), VALIGN_CENTER);
			Excel(xlcFormatFont, XLL_HDR);
			Excel(xlcPatterns, OPER(1), xBlue, OPER(0));

			Move(0, 1);
			Excel(xlcFormula, OPER("Description"));
			Excel(xlcAlignment, HALIGN_LEFT, OPER(), VALIGN_CENTER);
			Excel(xlcFormatFont, XLL_HDR);
			Excel(xlcPatterns, OPER(1), xBlue, OPER(0));

			Move(1, -1);
			int i = 0;
			for (auto& [key, args] : text) {
				TableRow();
				Excel(xlcFormula, args->FunctionText());
				Excel(xlcAlignment, XLL_ALIGN_RIGHT);
				Excel(xlcFormatFont, XLL_ARG);

				Move(0, 1);
				Excel(xlcFormula, args->FunctionHelp());
				// PATTERNS(apattern, afore, aback, newui)
				if (i % 2 == 1) {
					Background(xGrey);
				}
				Move(1, -1);
				++i;
			}
		}
		Move(1, 0);
	}
	// select column C and fit text width
	// COLUMN.WIDTH(width_num, reference, standard, type_num, standard_num)
	Excel(xlcColumnWidth, OPER(), OPER("C3:C3"), OPER(), OPER(3));

	// select cell containing the hyperlink
	Excel(xlcSelect, OPER(REF(1, 1)));

	return true;
}

static AddIn xai_spreadsheet_doc(
	Macro(XLL_DECORATE("_xll_spreadsheet_doc", 0), "DOC")
	.Category("XLL")
);
extern "C" __declspec(dllexport) int WINAPI
xll_spreadsheet_doc(void)
{
	int result = FALSE;

	try {
		Excel(xlcEcho, OPER(false));
		result = Spreadsheet(R"(All functions and macros having documentation.)", true);
		Excel(xlcEcho, OPER(true));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR("unknown exception in xll::Spreadsheet");
	}

	return result;
}
