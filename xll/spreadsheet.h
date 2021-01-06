// spreadsheet.h - tools for automating spreadsheet creation
#pragma once

#define ADDIN_URL "https://xlladdins.com/addins/"
#define XLL_FONT "Calibri"

// https://xlladdins.github.io/Excel4Macros/format.font.html
// FORMAT.FONT(name_text, size_num, bold, italic, underline, strike, color, outline, shadow)
#define XLL_H1 OPER(XLL_FONT), OPER(14), OPER(true)
#define XLL_H2 OPER(XLL_FONT), OPER(13), OPER(true)
#define XLL_ARG OPER(XLL_FONT), OPER(11), OPER(true)

// https://xlladdins.github.io/Excel4Macros/alignment.html
// ALIGNMENT(horiz_align, wrap, vert_align, orientation, add_indent, shrink_to_fit, merge_cells)
#define XLL_ALIGN_LEFT   OPER(2)
#define XLL_ALIGN_CENTER OPER(3)
#define XLL_ALIGN_RIGHT  OPER(4)

// https://xlladdins.github.io/Excel4Macros/options.view.html
// OPTIONS.VIEW(formula, status, notes, show_info, object_num, page_breaks, formulas, gridlines, color_num, headers, outline, zeros, hor_scroll, vert_scroll, sheet_tabs)

namespace xll {

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
		Excel(xlcWorkbookName, GetBook(), name); // [Book]Sheet
	}

}
