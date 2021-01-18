// spreadsheet.h - tools for automating spreadsheet creation
#pragma once
#include "xll.h"

#define ADDIN_URL "https://xlladdins.com/addins/"
#define XLL_FONT "Calibri"

using xll::Excel;

namespace xll {

	// Add to color palette at index
	// xGreen = EditColor(56, 0, 0xFF, 0);
	inline OPER EditColor(int index, int r, int g, int b)
	{
		ensure(1 <= index and index <= 56);
		ensure(0 <= r and r <= 255);
		ensure(0 <= g and g <= 255);
		ensure(0 <= b and b <= 255);

		ensure(xll::Excel(xlcEditColor, OPER(index), OPER(r), OPER(g), OPER(b)));

		return OPER(index);
	}

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
			Alignment align;

			align.horiz_align = (int)e;

			return align.Align();
		}
		static OPER Align(enum Vertical e)
		{
			Alignment align;

			align.vert_align = (int)e;

			return align.Align();
		}
		static OPER Align(enum Orientation e)
		{
			Alignment align;

			align.orientation = (int)e;

			return align.Align();
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

		void Format()
		{
			ensure(xll::Excel(xlcFormatFont, name_text, size_num,
				bold, italic, underline, strike,
				color, outline, shadow));
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
	};


	// Cell information and settings
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
			ensure(xll::Excel(xlcDefineStyle, name, OPER(6), OPER(0), fore, back));

			return *this;
		}
	};

	// foreground and background cell colors
	inline OPER Patterns(OPER fore, OPER back)
	{
		return xll::Excel(xlcPatterns, OPER(0), fore, back);
	}

	class Selection {
		OPER sel; // ??? must be an sref
	public:
		Selection(const OPER& sel = xll::Excel(xlfActiveCell))
			: sel(sel)
		{
			Select();
		}
		Selection(const REF& sel)
			: sel(OPER(sel))
		{ }
		Selection(const char* sel)
			: sel(xll::Excel(xlfTextref, OPER(sel), OPER(false))) // R1C1
		{ }

		OPER Select()
		{
			return xll::Excel(xlcSelect, sel);
		}

		// move r rows and c columns
		OPER Move(int r, int c)
		{
			sel = xll::Excel(xlfOffset, sel, OPER(r), OPER(c));

			return Select();
		}

		// Enters numbers, text, references, and formulas in a worksheet
		OPER Formula(const OPER& formula)
		{
			return xll::Excel(xlcFormula, formula, sel);
		}
		// https://docs.microsoft.com/en-us/office/client-developer/excel/xlset
		OPER Set(const OPER& set)
		{
			return xll::Excel(xlSet, sel, set);
		}
		OPER RowHeight(int points)
		{
			return xll::Excel(xlcRowHeight, OPER(points), sel);
		}
		OPER RowVisible(bool visible = true)
		{
			return xll::Excel(xlcRowHeight, OPER(), sel, OPER(), OPER(visible ? 2 : 1));
		}
		// Sets the row selection to a best-fit height
		OPER RowFit()
		{
			return xll::Excel(xlcRowHeight, OPER(), sel, OPER(), OPER(3));
		}
		OPER ColumnWidth(int points)
		{
			return xll::Excel(xlcColumnWidth, OPER(points), sel);
		}
		OPER ColumnVisible(bool visible = true)
		{
			return xll::Excel(xlcColumnWidth, OPER(), sel, OPER(), OPER(visible ? 2 : 1));
		}
		// Sets the column selection to a best-fit width
		OPER ColumnFit()
		{
			return xll::Excel(xlcColumnWidth, OPER(), sel, OPER(), OPER(3));
		}
		// Creates a comment 
		OPER Note(const char* note)
		{
			return xll::Excel(xlcNote, OPER(note), sel);
		}

	};

	struct Name {
		OPER name;
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
		static void Rename(const OPER& name)
		{
			ensure(xll::Excel(xlcWorkbookName, Document::Book(), name));
		}
		static void Select(const OPER& name = Document::Sheet())
		{
			ensure(xll::Excel(xlcWorkbookSelect, name));
		}
		static void Insert()
		{
			ensure(xll::Excel(xlcWorkbookInsert, OPER(1))); // insert Worksheet
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
	// paste range x at ref and return next ref below
	template<class X, class C>
	inline XREF<X> paste_next(XREF<X> ref, const XOPER<X>& x, XOPER<X>& style = XOPER<X>{})
	{
		auto rw = rows(x);
		auto col = columns(x);
		auto r = XExcel<X>(xlfOffset, XOPER<X>(ref), XOPER<X>(0), XOPER<X>(0), XOPER<X>(rw), XOPER<X>(col));
		XExcel<X>(xlcFormula, x, r);
		if (style) {
			Excel(xlcApplyStyle, style);
		}
		// move down r rows
		return XExcel<X>(xlfOffset, ref, XOPER<X>(r)).val.sref;
	}

}
