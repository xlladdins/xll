// spreadsheet.h - tools for automating spreadsheet creation
// 
// For xlcMacroFun having many arguments we define struct MacroFun.
// It has OPER members with exactly the same name as the documentation
// and a member OPER Fun() that calls Excel(xlcMacroFun, args...),
// and possible come convenience functions making it easier to use.
// We throw in some enums from the documentation for the macro function
// and provide static members for common operations.
#pragma once
#define XLOPERX XLOPER12
#include "xll.h"

#define ADDIN_URL "https://xlladdins.com/addins/"
#define XLL_FONT "Calibri"

using xll::Excel;

namespace xll {

	// RGB colors in the palette
	class Color {
		inline static int count = 57; // largest index
	public:
		OPER r, g, b;
		Color(int _r, int _g, int _b)
			: r(_r), g(_g), b(_b)
		{
			ensure(0 <= _r and _r <= 255);
			ensure(0 <= _g and _g <= 255);
			ensure(0 <= _b and _b <= 255);
		}
		// Add to color palette
		OPER Edit() const
		{
			ensure(--count > 10);
			ensure(xll::Excel(xlcEditColor, OPER(count), OPER(r), OPER(g), OPER(b)));

			return OPER(count);
		}
	};

	// https://xlladdins.github.io/Excel4Macros/alignment.html
	// selection alignment
	struct Alignment {
		OPER horiz_align = Missing; 
		OPER wrap = Missing;
		OPER vert_align = Missing;
		OPER orientation = Missing;
		OPER add_indent = Missing;
		OPER shrink_to_fit = Missing;
		OPER merge_cells = Missing;

		enum class Horizontal {
			General = 1,
			Left,
			Center,
			Right,
			Fill,
			Justify,
			CenterCells
		};
		enum class Vertical {
			Top = 1,
			Center,
			Bottom,
			Justify
		};
		enum class Orientation {
			Horizontal = 0,
			Vertical,
			Upward,
			Downward
		};

		OPER Align()
		{
			return xll::Excel(xlcAlignment, horiz_align, wrap, vert_align, orientation);
		}
		static OPER Align(enum Horizontal e)
		{
			Alignment align; align.horiz_align = (int)e; return align.Align();
		}
		static OPER Align(enum Vertical e)
		{
			Alignment align; align.vert_align = (int)e; return align.Align();
		}
		static OPER Align(enum Orientation e)
		{
			Alignment align; align.orientation = (int)e; return align.Align();
		}
	};

	// https://xlladdins.github.io/Excel4Macros/format.font.html
	// FORMAT.FONT(name_text, size_num, bold, italic, underline, strike, color, outline, shadow)
	struct FormatFont {
		OPER name_text = Missing; // font name, e.g., "Calibri"
		OPER size_num = Missing;  // font size in points
		OPER bold = Missing;
		OPER italic = Missing;
		OPER underline = Missing;
		OPER strike = Missing;
		OPER color = Missing;
		OPER outline = Missing;
		OPER shadow = Missing;

		OPER Format()
		{
			return xll::Excel(xlcFormatFont, name_text, size_num,
				bold, italic, underline, strike,
				color, outline, shadow);
		}

		static OPER NameText(const char* name)
		{
			FormatFont ff; ff.name_text = name; return ff.Format();
		}
		static OPER SizeNum(int size)
		{
			FormatFont ff; ff.size_num = size; return ff.Format();
		}
		static OPER Color(const OPER& color)
		{
			FormatFont ff; ff.color = color; return ff.Format();
		}
		static OPER Bold(bool b = true)
		{
			FormatFont ff; ff.bold = b; return ff.Format();
		}
		static OPER Italic(bool b = true)
		{
			FormatFont ff; ff.italic = b; return ff.Format();
		}
		static OPER Underline(bool b = true)
		{
			FormatFont ff; ff.underline = b; return ff.Format();
		}
		static OPER Strike(bool b = true)
		{
			FormatFont ff; ff.strike = b; return ff.Format();
		}
	};

	// https://xlladdins.github.io/Excel4Macros/options.calculation.html
	struct OptionsCalculation {
		enum class Type {
			Automatic = 1,
			AutomaticExceptTables = 2,
			Manual = 3
		};
		OPER type_num;
		OPER iter = OPER(false);
		OPER max_num = OPER(100);
		OPER max_change = OPER(0.001);
		OPER update = OPER(true);
		OPER precision = OPER(false);
		OPER date_1904 = OPER(false);
		OPER calc_save = OPER(true);
		OPER save_values = OPER(true);

		OptionsCalculation(enum Type type = Type::Automatic)
		{
			type_num = OPER((int)type);
		}

		OPER Calculation()
		{
			return xll::Excel(xlcOptionsCalculation, type_num, iter, max_num, max_change, 
				update, precision, date_1904, calc_save, save_values);
		}
	};
	// https://xlladdins.github.io/Excel4Macros/options.view.html
	struct OptionsView {
		OPER formula = OPER(true);
		OPER status = OPER(true);
		OPER notes = OPER(false);
		OPER show_info = Missing;
		OPER object_num = Missing;
		OPER page_breaks = OPER(false);
		OPER formulas = OPER(false);
		OPER gridlines = OPER(true);
		OPER color_num = OPER(0);
		OPER headers = OPER(true);
		OPER outline = OPER(true);
		OPER zeros = OPER(true);
		OPER hor_scroll = OPER(true);
		OPER vert_scroll = OPER(true);
		OPER sheet_tabs = OPER(true);

		OPER View()
		{
			return xll::Excel(xlcOptionsView, formula, status, notes, 
				show_info, object_num, page_breaks, 
				formulas, gridlines, color_num, headers, outline, 
				zeros, hor_scroll, vert_scroll, sheet_tabs);
		}

		OptionsView& Gridlines(bool b)
		{
			gridlines = b; return *this;
		}
	};


	// Cell information from xlcGetCell
	struct Cell {
		OPER ref;
		Cell(const OPER& ref = xll::Excel(xlfActiveCell))
			: ref(ref)
		{ }
		Cell(REF ref)
			: ref(OPER(ref))
		{ }
		Cell(const char* ref)
			: ref(xll::Excel(xlfTextref, OPER(ref), OPER(false))) // R1C1
		{ }

		// generic get
		OPER Get(int i) const
		{
			return xll::Excel(xlfGetCell, OPER(i));
		}

		// Absolute reference of the upper-left cell in reference, as text in the current workspace reference style.
		OPER AbsRef() const
		{
			return xll::Excel(xlfGetCell, OPER(1), ref);
		}
		// Same as TYPE(reference).
		OPER Type() const
		{
			return xll::Excel(xlfType, ref);
		}
		// Contents of reference.
		OPER Contents() const
		{
			return xll::Excel(xlfGetCell, OPER(5), ref);
		}
		// Formula in reference, as text, in either A1 or R1C1 style depending on the workspace setting.
		OPER Formula() const
		{
			return xll::Excel(xlfGetCell, OPER(6), ref);
		}
		// Number format of the cell
		OPER FormatNumber() const
		{
			return xll::Excel(xlfGetCell, OPER(7), ref);
		}
		Alignment Align() const
		{
			Alignment align;

			align.horiz_align = xll::Excel(xlfGetCell, OPER(7), ref);
			align.vert_align = xll::Excel(xlfGetCell, OPER(50), ref);
			align.orientation = xll::Excel(xlfGetCell, OPER(51), ref);

			return align;
		}
		// 14 	If the cell is locked, returns TRUE; otherwise, returns FALSE.
		// 15 	If the cell's formula is hidden, returns TRUE; otherwise, returns FALSE.
		// Row height of cell, in points.
		OPER RowHeight() const
		{
			return xll::Excel(xlfGetCell, OPER(17), ref);
		}
		// GET.CELL(44) - GET.CELL(42) to determine the width.
		OPER Width() const
		{
			return OPER(Get(44).val.num - Get(42).val.num);
		}
		// Name of font.
		OPER FontName() const
		{
			return xll::Excel(xlfGetCell, OPER(18), ref);
		}
		// Size of font, in points.
		OPER FontSize() const
		{
			return xll::Excel(xlfGetCell, OPER(19), ref);
		}
		// Style of the cell.
		OPER Style() const
		{
			return xll::Excel(xlfGetCell, OPER(40), ref);
		}
		// 46 	If the cell contains a text note, returns TRUE; otherwise, returns FALSE.
		// This returns the note text
		OPER Note() const
		{
			return xll::Excel(xlfGetNote, ref);
		}
		// Returns the name of the sheet in the form "[Book1]Sheet1".
		OPER Name() const
		{
			return xll::Excel(xlfGetCell, OPER(62), ref);
		}
	};

	// prefer styles to explicit formatting
	struct Style {
		OPER name;
		Style(const char* name)
			: name(name)
		{ }
		OPER Apply()
		{
			return xll::Excel(xlcApplyStyle, name);
		}
		OPER Delete()
		{
			return xll::Excel(xlcDeleteStyle);
		}
		// define styles
		Style& FormatNumber(const char* format)
		{
			ensure(xll::Excel(xlcDefineStyle, name, OPER(2), OPER(format)));

			return *this;
		}
		Style& FormatFont(const xll::FormatFont& format)
		{
			ensure(xll::Excel(xlcDefineStyle, name, OPER(3),
				format.name_text, 
				format.size_num, 
				format.bold, 
				format.italic, 
				format.underline, 
				format.strike, 
				format.color));
			
			return *this;
		}
		Style& Alignment(const xll::Alignment& align)
		{
			ensure(xll::Excel(xlcDefineStyle, name, OPER(4),
				align.horiz_align,
				align.wrap,
				align.vert_align,
				align.orientation));

			return *this;
		}
		//!!! 4-7
		Style& Pattern(const OPER& fore, const OPER& back = Missing)
		{
			ensure(xll::Excel(xlcDefineStyle, name, OPER(6), OPER(1), fore, back));

			return *this;
		}
	};

	// foreground and background cell colors
	inline OPER Patterns(OPER fore, OPER back)
	{
		return xll::Excel(xlcPatterns, OPER(1), fore, back);
	}

	// https://xlladdins.github.io/Excel4Macros/selection.html
	// https://xlladdins.github.io/Excel4Macros/select.html
	struct Select {
		OPER selection; // ??? must be an sref
		Select(const OPER& _selection = xll::Excel(xlfActiveCell))
			: selection(_selection)
		{
			if (selection.xltype & xltypeRef) {
				selection = OPER(selection.val.mref.lpmref->reftbl[0]);
			}

			ensure(Excel(xlcSelect, selection));
		}
		Select(const REF& selection)
			: Select(OPER(selection))
		{ }
		Select(const char* selection)
			: Select(xll::Excel(xlfTextref, OPER(selection), OPER(false))) // R1C1
		{ }

		OPER Offset(int r, int c, int h = 1, int w = 1) const
		{
			return xll::Excel(xlfOffset, selection, OPER(r), OPER(c), OPER(w), OPER(h));
		}

		// move r rows and c columns
		Select& Move(int r, int c)
		{
			selection = Offset(r, c);
			
			ensure(Excel(xlcSelect, selection));

			return *this;
		}
		Select& Down(int r = 1)
		{
			return Move(r, 0);
		}
		Select& Up(int r = 1)
		{
			return Move(-r, 0);
		}
		Select& Right(int c = 1)
		{
			return Move(0, c);
		}
		Select& Left(int c = 1)
		{
			return Move(0, -c);
		}

		// Enters numbers, text, references, and formulas in a worksheet
		OPER Formula(const OPER& formula)
		{
			return xll::Excel(xlcFormula, formula, selection);
		}
		// https://docs.microsoft.com/en-us/office/client-developer/excel/xlset
		OPER Set(const OPER& set)
		{
			return xll::Excel(xlSet, selection, set);
		}
		OPER RowHeight(int points)
		{
			return xll::Excel(xlcRowHeight, OPER(points), selection);
		}
		OPER RowVisible(bool visible = true)
		{
			return xll::Excel(xlcRowHeight, OPER(), selection, OPER(), OPER(visible ? 2 : 1));
		}
		// Sets the row selection to a best-fit height
		OPER RowFit()
		{
			return xll::Excel(xlcRowHeight, OPER(), selection, OPER(), OPER(3));
		}
		OPER ColumnWidth(int points)
		{
			return xll::Excel(xlcColumnWidth, OPER(points), selection);
		}
		OPER ColumnVisible(bool visible = true)
		{
			return xll::Excel(xlcColumnWidth, OPER(), selection, OPER(), OPER(visible ? 2 : 1));
		}
		// Sets the column selection to a best-fit width
		OPER ColumnFit()
		{
			return xll::Excel(xlcColumnWidth, OPER(), selection, OPER(), OPER(3));
		}
		// Creates a comment 
		OPER Note(const char* note)
		{
			return xll::Excel(xlcNote, OPER(note), selection);
		}
		enum Type {
			Notes = 1,
			Constants = 2,
			Formulas = 3,
			Blanks = 4,
			CurrentRegion = 5,
			Current = 6, // current array
			RowDifferences = 7,
			ColumnDifferences = 8,
			Precedents = 9,
			Dependents = 10,
			LastCell = 11,
			VisibleCellsOnly = 12, // (outlining)
			AllObjects = 13
		};
		enum ValueType {
			Numbers = 1,
			Text = 2,
			Logical = 4, 
			Error = 16,
			Default = Numbers + Text + Logical + Error
		};
		enum Level {
			Direct = 1,
			All = 2
		};
		// uses active cell implicitly
		static OPER Special(Type type, ValueType value = ValueType::Default, Level level = Level::Direct)
		{
			OPER type_num = type;
			OPER value_type = Missing;
			OPER levels = Missing;

			if (type == Type::Constants or type == Type::Formulas) {
				value_type = value;
			}
			else if (type == Type::Precedents or type == Type::Dependents) {
				levels = level;
			}

			return xll::Excel(xlcSelectSpecial, type_num, value_type, levels);
		}
	};

	struct Name {
		OPER name;
		Name(const OPER& name)
			: name(name)
		{ }
		Name(const char* name)
			: name(name)
		{ }
		OPER Define(const OPER& ref, bool local)
		{
			return xll::Excel(xlcDefineName, name, ref, OPER(3), Missing, OPER(false), Missing, OPER(local));
		}
		OPER Delete()
		{
			return xll::Excel(xlcDeleteName, name);
		}
		OPER Get()
		{
			return xll::Excel(xlfGetName, name);
		}
		// xlcSetName is only for macro sheets
	};

	struct Document {
		static OPER Book()
		{
			return xll::Excel(xlfGetDocument, OPER(88)); // Book
		}

		static OPER Sheet()
		{
			return xll::Excel(xlfGetDocument, OPER(76)); // [Book]Sheet
		}
	};

	struct Workbook {
		static OPER Rename(const OPER& name)
		{
			return xll::Excel(xlcWorkbookName, Document::Book(), name);
		}
		static OPER Select(const OPER& name = Document::Sheet())
		{
			return xll::Excel(xlcWorkbookSelect, name);
		}
		// https://xlladdins.github.io/Excel4Macros/workbook.insert.html
		enum class Type {
			Worksheet = 1,
			Chart = 2,
			MacroSheet = 3,
			InternationallMacroSheet = 4,
			VisualBasicModule = 6,
			Dialog = 7
		};
		static OPER Insert(enum Type type = Type::Worksheet)
		{
			return xll::Excel(xlcWorkbookInsert, OPER((int)type));
		}
		static OPER Move(const OPER& name, int position = 1)
		{
			return xll::Excel(xlcWorkbookMove, name, Document::Sheet(), OPER(position));
		}
	};

	/*
	inline void Header()
	{
		xll::Excel(xlcDisplay, OPER(false), OPER(false)); // no gridlines
		xll::Excel(xlcRowHeight, OPER(24), OPER("R2:R2"));
		xll::Excel(xlcSelect, OPER("R2:R2"));
		FormatFont ff(OPER(XLL_FONT), OPER(14), xWhite);
		ff.Bold().Set();
		xll::Excel(xlcSelect, OPER("R1:R3"));
		Excel(xlcPatterns, OPER(1), xGreen, OPER(0));
	}
	inline void Category()
	{
		Excel(xlcRowHeight, OPER(20), OPER("R[0]:R[0]"));
		FormatFont ff(OPER(XLL_FONT), OPER(14), xBlack);
		ff.Bold().Set();
		//Excel(xlcAlignment, HALIGN_RIGHT, OPER(), VALIGN_CENTER);
	}
	inline void TableHeader()
	{
		Excel(xlcRowHeight, OPER(17.5), OPER("R[0]:R[0]"));
	}
	inline void TableRow()
	{
		Excel(xlcRowHeight, OPER(15), OPER("R[0]:R[0]"));
	}




	inline void Move(short r, short c)
	{
		Excel(xlcSelect, Excel(xlfOffset, Excel(xlfActiveCell), OPER(r), OPER(c)));
	}
	*/
	// paste default argument at ref and return reference to what was pasted
	template<class X>
	inline OPER paste_default(OPER ref, const XArgs<X>* pa, unsigned i)
	{
		const OPER& x = pa->ArgumentDefault(i);

		if (x.is_str() and x.val.str[1] == '=') {
			// formula
			OPER xi = XExcel<X>(xlfEvaluate, x);
			ensure(xi);
			if (size(xi) == 1) {
				ensure(XExcel<X>(xlcFormula, x, ref));
			}
			else {
				auto rw = rows(xi);
				auto col = columns(xi);
				ref = XExcel<X>(xlfOffset, ref, OPER(0), OPER(0), OPER(rw), OPER(col));
				ensure(XExcel<X>(xlcFormulaArray, x, ref))
			}
		}
		else {
			auto rw = rows(x);
			auto col = columns(x);
			ref = XExcel<X>(xlfOffset, ref, OPER(0), OPER(0), OPER(rw), OPER(col));
			ensure(XExcel<X>(xlcFormula, x, ref));
		}
		
		return ref;
	}

}
