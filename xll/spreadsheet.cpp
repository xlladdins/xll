// spreadsheet.cpp
#include "splitpath.h"
#include "xll.h"
#include "spreadsheet.h"

using namespace xll;

// system default
const OPER xBlack(1);
const OPER xWhite(2);

// custom colors
const Color grey(0xF0, 0xF0, 0xF0);
static OPER xGrey;
const Color green(0x00, 0xD1, 0xB1);
static OPER xGreen;
const Color blue(0x31, 0x8B, 0xCE);
static OPER xBlue;

inline Style H1()
{
	FormatFont format;
	format.color = xWhite;
	format.size_num = 24;

	Alignment align;
	align.vert_align = (int)Alignment::Vertical::Center;
	align.horiz_align = (int)Alignment::Horizontal::Left;

	Style s("H1");
	s.FormatFont(format).Alignment(align).Pattern(xGreen, xWhite);

	return s;
}

#pragma warning(disable: 100)
// create sample spreadsheet
bool Spreadsheet(const char* description = "", bool release = false)
{
	using cat_type_text_map = std::map<OPER, std::map<OPER, std::map<OPER, Args*>>>;

	// custom colors
	xGrey = grey.Edit();
	xGreen = green.Edit();
	xBlue = blue.Edit();

	// get dir and filename extension
	splitpath sp(Excel4(xlGetName).to_string().c_str());
	OPER dir(release ? ADDIN_URL : sp.dirname().c_str());
	OPER Name(Excel(xlfUpper, OPER(sp.fname)));

	// sort by category, macro type, then function text
	cat_type_text_map cat_type_text;
	for (auto& [key, args] : AddIn::Map) {
		cat_type_text[args.Category()][args.Type()][args.FunctionText()] = &args;
	}

	std::list<std::pair<OPER, Args*>> tabs;
	for (auto& [cat, type_text] : cat_type_text) {
		for (auto& [type, text] : type_text) {
			for (auto& [key, args] : text) {
				tabs.push_front(std::pair(key, args)); // insert happens to the left
			}
		}
	}

	// global settings
	//OptionsView global;
	//global.gridlines = false;
	//global.View();

	Workbook::Rename(Name);

	// insert sheets with sample call
	for (auto& [tab, args] : tabs) {
		if (args->Documentation().size() == 0)
			continue;
		
		Workbook::Select();
		Workbook::Insert();
		Workbook::Rename(tab);

		Select sel("R1:R1");
		H1().Apply();

		sel = Select(REF(0,1)); // B1
		sel.Set(tab & OPER(" ") & args->Type());
		sel.Move(1, 0);
		// ??? H2.Apply() 
		sel.Set(args->FunctionHelp());
		
		if (args->isFunction()) {
			sel.Move(1, 0);
			sel.Set(args->FunctionText());
			Alignment::Align(Alignment::Horizontal::Right);
			sel.Move(0, 1);
			sel.Formula(OPER("=") & args->FunctionText() & OPER("(") & args->ArgumentText() & OPER(")"));

			sel.Move(1, -1);
			for (unsigned i = 1; i <= args->ArgumentCount(); ++i) {
				sel.Set(args->ArgumentName(i));
				Alignment::Align(Alignment::Horizontal::Right);
				FormatFont::Bold();
				
				sel.Move(0, 1);
				OPER ref = paste_default(sel.selection, args, i);
				xll::Name namei(args->ArgumentName(i));
				namei.Define(ref, true); // local name
				sel.Move(0, 1);
				sel.Set(args->ArgumentHelp(i));
				sel.Move(rows(ref), -2);
			}
		}

	}

	OPER book = Document::Sheet();

	Excel(xlcWorkbookMove, Name, book, OPER(1)); // move main page to first position
	Excel(xlcWorkbookSelect, Name);
	Select sel("R1:R1");
	H1().Apply();

	sel = Select(REF(0, 1)); // B1

	// "=HYPERLINK(\"https://xlladdins.com/addins/xll_math.xll\", \"xll_math\")"));
	sel.Formula(OPER("=HYPERLINK(\"") & dir & OPER(sp.fname) & OPER(" & WinXX() & ") & OPER(sp.ext) & OPER("\", \"") & Name & OPER("\")"));
	Alignment::Align(Alignment::Horizontal::Left);
	//Excel(xlcNote, OPER("user: addinuser pass: @addm3"));

	sel.Move(1, 0);
	sel.Set(OPER(description));

	sel.Move(1, 0);

	for (auto& [cat, type_text] : cat_type_text) {
		sel.Set(OPER("Category ") & cat);
		FormatFont::SizeNum(16);
		//FormatFont::Bold();
		sel.Move(1, 0);
		for (auto& [type, text] : type_text) {

			sel.Set(Excel(xlfProper, type));
			Alignment::Align(Alignment::Horizontal::Right);
			//Alignment::Align(Alignment::Vertical::Center);
			FormatFont::SizeNum(13);
			FormatFont::Color(xWhite);
			Patterns(xBlue, xWhite);

			// table header
			sel.Move(0, 1);
			sel.Set(OPER("Description"));
			Alignment::Align(Alignment::Horizontal::Left);
			//Alignment::Align(Alignment::Vertical::Center);
			FormatFont::SizeNum(13);
			FormatFont::Color(xWhite);
			Patterns(xBlue, xWhite);

			sel.Move(1, -1);
			int i = 1;
			for (auto& [key, args] : text) {
				if (!args->FunctionHelp()) {
					continue;
				}
				// table row
				Excel(xlcFormula, args->FunctionText());
				Alignment::Align(Alignment::Horizontal::Right);
				FormatFont::Bold();

				sel.Move(0, 1);
				Excel(xlcFormula, args->FunctionHelp());
				// PATTERNS(apattern, afore, aback, newui)
				//if (i % 2 == 1) {
					//Background(xGrey);
				//}
				sel.Move(1, -1);
				++i;
			}
		}
		sel.Move(1, 0);
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

	//Excel(xlcEcho, OPER(false));
	OptionsCalculation oc;
	oc.type_num = OPER((int)OptionsCalculation::Type::Manual);
	oc.Calculation();
	try {
		result = Spreadsheet(R"(All functions and macros having documentation.)", true);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR("unknown exception in xll::Spreadsheet");
	}
	oc.type_num = OPER((int)OptionsCalculation::Type::Automatic);
	oc.Calculation();
	//Excel(xlcEcho, OPER(true));

	return result;
}
