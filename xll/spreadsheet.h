// spreadsheet.h - tools for automating spreadsheet creation
#pragma once

#define ADDIN_URL "https://xlladdins.com/addins/"
#define XLL_FONT "Calibri"

// https://xlladdins.github.io/Excel4Macros/format.font.html
// FORMAT.FONT(name_text, size_num, bold, italic, underline, strike, color, outline, shadow)
#define XLL_ARG OPER(XLL_FONT), OPER(11), OPER(true)
#define XLL_DEF OPER(XLL_FONT), OPER(11)
#define XLL_HDR OPER(XLL_FONT), OPER(11), OPER(true), OPER(false), OPER(false), OPER(false), xWhite

// https://xlladdins.github.io/Excel4Macros/alignment.html
// ALIGNMENT(horiz_align, wrap, vert_align, orientation, add_indent, shrink_to_fit, merge_cells)
#define XLL_ALIGN_LEFT   OPER(2)
#define XLL_ALIGN_CENTER OPER(3)
#define XLL_ALIGN_RIGHT  OPER(4)

// https://xlladdins.github.io/Excel4Macros/options.view.html
// OPTIONS.VIEW(formula, status, notes, show_info, object_num, page_breaks, formulas, gridlines, color_num, headers, outline, zeros, hor_scroll, vert_scroll, sheet_tabs)

namespace xll {

	inline const OPER HALIGN_GENERAL(1);
	inline const OPER HALIGN_LEFT(2);
	inline const OPER HALIGN_CENTER(3);
	inline const OPER HALIGN_RIGHT(4);
	inline const OPER HALIGN_FILL(5);
	inline const OPER HALIGN_JUSTIFY(6);

	inline const OPER VALIGN_TOP(1);
	inline const OPER VALIGN_CENTER(2);
	inline const OPER VALIGN_BOTTOM(3);
	inline const OPER VALIGN_JUSTIFY(4);

	inline const OPER xBlack(1);
	inline const OPER xWhite(2);
	inline const OPER xGrey(56);
	inline const OPER xGreen(55);
	inline const OPER xBlue(54);

	// call to set colors
	inline void Colors()
	{
		Excel(xlcEditColor, xGrey,  OPER(0xF0), OPER(0xF0), OPER(0xF0));
		Excel(xlcEditColor, xGreen, OPER(0x00), OPER(0xD1), OPER(0xB1));
		Excel(xlcEditColor, xBlue,  OPER(0x31), OPER(0x8B), OPER(0xCE));
	}

	class FormatFont {
		OPER name_text; // font name, e.g., "Calibri"
		OPER size_num;  // font size in points
		OPER bold; 
		OPER italic; 
		OPER underline; 
		OPER strike; 
		OPER color; 
		OPER outline; 
		OPER shadow;
		// FORMAT.FONT(name_text, size_num, bold, italic, underline, strike, color, outline, shadow)
	public:
		FormatFont(const OPER& name_text = OPER(), const OPER& size_num = OPER(), const OPER& color = OPER())
			: name_text(name_text), size_num(size_num), color(color)
		{ }

		FormatFont& Set()
		{
			Excel(xlcFormatFont, name_text, size_num, bold, italic, underline, strike, color, outline, shadow);

			return *this;
		}

		FormatFont& Color(const OPER& _color)
		{
			color = _color;

			return *this;
		}
		FormatFont& Bold()
		{
			bold = true;

			return *this;
		}
		FormatFont& Italic()
		{
			italic = true;

			return *this;
		}
		FormatFont& Underline()
		{
			underline = true;

			return *this;
		}

	};

	class Selection {
		OPER xSel;
	public:
		Selection(const OPER& xSel = Excel(xlfActiveCell))
			: xSel(xSel)
		{ }
		// return old selection
		Selection Select()
		{
			Selection oSel(Excel(xlfSelection));
			Excel(xlcSelect, xSel);

			return oSel;
		}
		Selection& Offset(int r, int c)
		{
			xSel = Excel(xlfOffset, xSel, OPER(r), OPER(c));

			return *this;
		}
	};

	inline void Background(const OPER& color)
	{
		Excel(xlcPatterns, OPER(1), color, OPER(0));
	}
	inline void Foreground(const OPER& color)
	{
		FormatFont ff(OPER(), OPER(), color);
		ff.Set();
	}
	inline void Bold()
	{
		FormatFont ff;
		ff.Bold().Set();
	}

	inline void Header()
	{
		Excel(xlcDisplay, OPER(false), OPER(false)); // no gridlines
		Excel(xlcRowHeight, OPER(24), OPER("R2:R2"));
		Excel(xlcSelect, OPER("R2:R2"));
		FormatFont ff(OPER(XLL_FONT), OPER(14), xWhite);
		ff.Bold().Set();
		Excel(xlcSelect, OPER("R1:R3"));
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

	inline void ApplyStyle(const char* style)
	{
		Excel(xlcDefineStyle, OPER(style));
		Excel(xlcApplyStyle, OPER(style));
	}

	// define name for active cell
	inline void DefineName(const OPER& name, bool local)
	{
		Excel(xlcDefineName, name, Excel(xlfActiveCell), OPER(3), Missing, OPER(false), Missing, OPER(local));
	}

	inline OPER GetBook()
	{
		return Excel(xlfGetDocument, OPER(88)); // Book
	}

	inline OPER GetSheet()
	{
		return Excel(xlfGetDocument, OPER(76)); // [Book]Sheet
	}

	inline void Move(short r, short c)
	{
		Excel(xlcSelect, Excel(xlfOffset, Excel(xlfActiveCell), OPER(r), OPER(c)));
	}

	// rename current tab
	inline void Rename(const OPER& name)
	{
		Excel(xlcWorkbookName, GetBook(), name);
	}


}
