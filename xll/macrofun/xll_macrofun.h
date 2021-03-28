// xll_macrofun.h - tools for automating spreadsheet creation
// 
// This header makes it convenient to call macro functions 
// documented in https://xlladdins.github.io/Excel4Macros/
// 
// For `xlcMacroFun` having many arguments we define `struct MacroFun`.
// It has members with exactly the same name as in the documentation
// and a member function `Fun()` that calls `Excel(xlcMacroFun, args...)`,
// and possibly come convenience functions making it easier to use.
// We throw in some enums from the documentation for the macro function
// and provide static members for common operations.
//
// Create a `MacroFun` object, set members as desired, then call `Fun()`.

#pragma once
//#define XLL_VERSION 4
#include "../xll.h"

namespace xll {

	using xll::Excel;

	// https://support.microsoft.com/en-us/office/offset-function-c8de19ae-dd79-4b9b-a14e-b4d906d11b66	
	inline OPER Offset(const REF& ref, const REF& off)
	{
		return Excel(xlfOffset, OPER(ref), 
			OPER(off.rwFirst), OPER(off.colFirst), 
			OPER(height(off)), OPER(width(off))
		);
	}

	// https://xlladdins.github.io/Excel4Macros/selection.html
	// https://xlladdins.github.io/Excel4Macros/select.html
	class Select {
		OPER selection; 
	public:
		// selection must be sref or ref
		Select(const OPER& sel = Excel(xlfActiveCell))
			: selection(sel.as_sref())
		{
			Excel(xlcSelect, selection);
		}
		/*
		Select(const REF& ref)
			: Select(OPER(ref))
		{ }
		*/
		// selection is R1C1 by default
		Select(const char* selection, bool A1 = false)
			: Select(Excel(xlfTextref, OPER(selection), OPER(A1)))
		{ }
		Select(const Select&) = delete;
		Select& operator=(const Select&) = delete;
		~Select()
		{ }

		operator const OPER& () const
		{
			return selection;
		}

		// select selection
		Select& operator()()
		{
			ensure(Excel(xlcSelect, selection));

			return *this;
		}
		// select ref
		Select& operator()(const REF& ref)
		{
			selection = Select(ref);

			return *this;
		}
		// select row r column c
		Select& operator()(int r, int c, unsigned h = 0, unsigned w = 0)
		{
			if (h == 0) {
				h = selection.rows();
			}
			if (w == 0) {
				w = selection.columns();
			}

			return operator()(REF(r, c, h, w));
		}

		/*
		Select& Offset(const REF& ref)
		{
			selection = Select(xll::Offset(selection.as_sref(), ref));

			return *this;
		}
		Select& Offset(int r, int c, unsigned h = 0, unsigned w = 0)
		{
			if (h == 0) {
				h = selection.rows();
			}
			if (w == 0) {
				w = selection.columns();
			}

			return Offset(REF(r, c, h, w));
		}
		*/

		// preserves upper left corner
		Select& Reshape(unsigned h, unsigned w)
		{
			selection = Select(reshape(selection.as_sref(), h, w));

			return *this;
		}
		Select& Reshape(const OPER& ref)
		{
			return Reshape(rows(ref), columns(ref));
		}

		// translate by r rows and c columns
		Select& Move(int r, int c)
		{
			selection = Select(translate(selection.as_sref(), r, c));

			return *this;
		}

		// Enters numbers, text, references, and formulas in a worksheet
		// https://xlladdins.github.io/Excel4Macros/formula.html
		// https://xlladdins.github.io/Excel4Macros/formula.array.html
		OPER Formula(const OPER& formula) const
		{
			return Excel(selection.size() == 1 ? xlcFormula : xlcFormulaArray, formula, selection);
		}

		// The reference is given as an R1C1-style relative reference in the form of text, such as "R[1]C[1]".
		// https://xlladdins.github.io/Excel4Macros/relref.html
		OPER Relref(const OPER& ref)
		{
			return Excel(xlfRelref, ref, selection);
		}
		// The reference must be an R1C1-style relative reference in the form of text, such as "R[1]C[1]".
		// https://xlladdins.github.io/Excel4Macros/absref.html
		OPER Absref(const OPER& ref)
		{
			return Excel(xlfAbsref, ref, selection);
		}

		// https://docs.microsoft.com/en-us/office/client-developer/excel/xlset
		OPER Set(const OPER& set)
		{
			return Excel(xlSet, selection, set);
		}

		// class Row, Column
		// https://xlladdins.github.io/Excel4Macros/row.height.html
		OPER RowHeight(int points)
		{
			return Excel(xlcRowHeight, OPER(points), selection);
		}
		OPER RowVisible(bool visible = true)
		{
			return Excel(xlcRowHeight, OPER(), selection, OPER(), OPER(visible ? 2 : 1));
		}
		// Sets the row selection to a best-fit height
		OPER RowFit()
		{
			return Excel(xlcRowHeight, OPER(), selection, OPER(), OPER(3));
		}
		// https://xlladdins.github.io/Excel4Macros/column.width.html
		OPER ColumnWidth(int points)
		{
			return Excel(xlcColumnWidth, OPER(points), selection);
		}
		OPER ColumnVisible(bool visible = true)
		{
			return Excel(xlcColumnWidth, OPER(), selection, OPER(), OPER(visible ? 2 : 1));
		}
		// Sets the column selection to a best-fit width
		OPER ColumnFit()
		{
			return Excel(xlcColumnWidth, OPER(), selection, OPER(), OPER(3));
		}
		// Creates a comment 
		OPER Note(const char* note)
		{
			return Excel(xlcNote, OPER(note), selection);
		}

		// https://xlladdins.github.io/Excel4Macros/select.special.html
		enum class Type {
			Notes = 1,
			Constants = 2,
			Formulas = 3,
			Blanks = 4,
			Current = 5, // region
			CurrentArray = 6,
			RowDifferences = 7,
			ColumnDifferences = 8,
			Precedents = 9,
			Dependents = 10,
			LastCell = 11,
			VisibleCellsOnly = 12, // (outlining)
			AllObjects = 13
		};
		enum class ValueType {
			Numbers = 1,
			Text = 2,
			Logical = 4,
			Error = 16,
			Default = Numbers + Text + Logical + Error
		};
		enum class Level {
			Direct = 1,
			All = 2
		};
		// Uses active cell implicitly. Select sel; sel().Special(...);
		static OPER Special(Type type, ValueType value = ValueType::Default, Level level = Level::Direct)
		{
			OPER type_num = OPER((int)type);
			OPER value_type = Missing;
			OPER levels = Missing;

			if (type == Type::Constants or type == Type::Formulas) {
				value_type = OPER((int)value);
			}
			else if (type == Type::Precedents or type == Type::Dependents) {
				levels = OPER((int)level);
			}

			return Excel(xlcSelectSpecial, type_num, value_type, levels);
		}
	};

	//
	// Parameterized functions
	//

	// move relative to active cell
	inline OPER Move(int r, int c)
	{
		Select sel; // active cell

		return sel.Move(r, c);
	}
	// move down r rows
	inline OPER Down(int r = 1)
	{
		return Move(r, 0);
	}
	// move up r rows
	inline OPER Up(int r = 1)
	{
		return Move(-r, 0);
	}
	// move right c columns
	inline OPER Right(int c = 1)
	{
		return Move(0, c);
	}
	// move left c columns
	inline OPER Left(int c = 1)
	{
		return Move(0, -c);
	}

	//
	// MacroFun objects
	//

	// ALIGNMENT(horiz_align, wrap, vert_align, orientation, add_indent, shrink_to_fit, merge_cells)
	// https://xlladdins.github.io/Excel4Macros/alignment.html
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
			return Excel(xlcAlignment, horiz_align, wrap, vert_align, orientation);
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

	// RGB colors in the palette
	// https://xlladdins.github.io/Excel4Macros/edit.color.html
	class EditColor {
		inline static int count = 57; // largest index + 1
	public:
		OPER r, g, b;
		EditColor(int _r, int _g, int _b)
			: r(_r), g(_g), b(_b)
		{
			ensure(0 <= _r and _r <= 255);
			ensure(0 <= _g and _g <= 255);
			ensure(0 <= _b and _b <= 255);
		}
		// Add from end of color palette
		OPER Color(int index = 0) const
		{
			if (index == 0) {
				ensure(--count > 1);
				index = count;
			}
			ensure(0 <= index and index <= 56);
			ensure(Excel(xlcEditColor, OPER(index), OPER(r), OPER(g), OPER(b)));

			return OPER(count);
		}
		// Microsoft color palette
		static OPER MSred()
		{
			EditColor _(0xF6, 0x53, 0x14); return _.Color();
		}
		static OPER MSgreen()
		{
			EditColor _(0x7C, 0xBB, 0x00); return _.Color();
		}
		static OPER MSblue()
		{
			EditColor _(0x00, 0xA1, 0xF1); return _.Color();
		}
		static OPER MSorange()
		{
			EditColor _(0xFF, 0xBB, 0x00); return _.Color();
		}
	};

	// FORMAT.FONT(name_text, size_num, bold, italic, underline, strike, color, outline, shadow)
	// https://xlladdins.github.io/Excel4Macros/format.font.html
	struct FormatFont {
		OPER name_text = Missing; // font name, e.g., "Calibri"
		OPER size_num = Missing;  // font size in points
		OPER bold = Missing;      // boolean
		OPER italic = Missing;    // boolean
		OPER underline = Missing; // boolean
		OPER strike = Missing;    // boolean
		OPER color = Missing;     // color palette index
		OPER outline = Missing;   // boolean
		OPER shadow = Missing;    // boolean

		OPER Font()
		{
			return Excel(xlcFormatFont, name_text, size_num,
				bold, italic, underline, strike,
				color, outline, shadow);
		}

		static OPER NameText(const char* name)
		{
			FormatFont ff; ff.name_text = name; return ff.Font();
		}
		static OPER SizeNum(int size)
		{
			FormatFont ff; ff.size_num = size; return ff.Font();
		}
		static OPER Color(const OPER& color)
		{
			FormatFont ff; ff.color = color; return ff.Font();
		}
		static OPER Bold(bool b = true)
		{
			FormatFont ff; ff.bold = b; return ff.Font();
		}
		static OPER Italic(bool b = true)
		{
			FormatFont ff; ff.italic = b; return ff.Font();
		}
		static OPER Underline(bool b = true)
		{
			FormatFont ff; ff.underline = b; return ff.Font();
		}
		static OPER Strike(bool b = true)
		{
			FormatFont ff; ff.strike = b; return ff.Font();
		}
	};

	// Factor out Object class with Polygon and Chart sublasses. 
	// Factor out Input class.
	// https://xlladdins.github.io/Excel4Macros/create.object.html
	struct CreateObject {
		OPER obj_type;
		OPER ref1;
		OPER x_offset1 = OPER(0);
		OPER y_offset1 = OPER(0);
		OPER ref2;
		OPER x_offset2 = OPER(0);
		OPER y_offset2 = OPER(0);
		OPER text;
		OPER fill = Missing;
		OPER editable = Missing;

		enum class Type {
			Line = 1,
			Rectangle = 2,
			Oval = 3,
			Arc = 4,
			EmbeddedChart = 5,
			TextBox = 6,
			Button = 7,
			ClosedPolygon = 9,
			OpenPolygon = 10,
			CheckBox = 11,
			OptionButton = 12,
			EditBox = 13,
			Label = 14,
			DialogFrame = 15,
			Spinner = 16,
			ScrollBar = 17,
			ListBox = 18,
			GroupBox = 19,
			DropDown = 20,
		};

		enum class Chart {
			Area = 1,
			Bar = 2,
			Column = 3,
			Line = 4,
			Pie = 5,
			Radar = 6,
			XY = 7,
			Combination = 8,
			Area3D = 9,
			Bar3D = 10,
			Column3D = 11,
			Line3D = 12,
			Pie3D = 13,
			Surface3D = 14,
			Doughnut = 15,
		};
		// CREATE.OBJECT(obj_type, ref1, x_offset1, y_offset1, ref2, x_offset2, y_offset2,
		//	text, fill, editable)
		CreateObject(enum Type type, const OPER& ref1, const OPER& ref2)
			: obj_type(static_cast<int>(type)), ref1(ref1), ref2(ref2)
		{ }
		// Polygon
		// CREATE.OBJECT(obj_type, ref1, x_offset1, y_offset1, ref2, x_offset2, y_offset2, 
		//	array, fill)
		CreateObject(const OPER& ref1, const OPER& ref2, const OPER& array, bool closed = false)
			: CreateObject(closed ? Type::ClosedPolygon : Type::OpenPolygon, ref1, ref2)
		{
			text = array;
		}
		// Embedded Chart
		// CREATE.OBJECT(obj_type, ref1, x_offset1, y_offset1, ref2, x_offset2, y_offset2,
		//	xy_series, fill, gallery_num, type_num, plot_visible)
		CreateObject(enum Chart gallery, const OPER& ref1, const OPER& ref2, const OPER& xy)
			: CreateObject(Type::EmbeddedChart, ref1, ref2)
		{
			text = xy;
			editable = static_cast<int>(gallery);
		}
		OPER Object()
		{
			return Excel(xlfCreateObject, obj_type, ref1, x_offset1, y_offset1,
				ref2, x_offset2, y_offset2, text, fill, editable);
		}

		OPER Offset1(int x, int y)
		{
			x_offset1 = x;
			y_offset1 = y;
		}
		OPER Offset2(int x, int y)
		{
			x_offset2 = x;
			y_offset2 = y;
		}
		// Call immediately after creating object to assign macro.
		OPER Assign(const OPER& macro)
		{
			return Excel(xlcAssignToObject, macro);
		}
		enum class Placement {
			Sized = 1,        // Moved and sized with cells.
			NotSized = 2,     // Moved but not sized with cells.
			FreeFloating = 3, // Not affected by moving and sizing cells.
		};
		// Object must be selected.
		OPER Properties(enum Placement placement, bool print = true)
		{
			return Excel(xlcObjectProperties, OPER(static_cast<int>(placement)), OPER(print));
		}

		enum class GetObject {
			Locked = 2, // If the object is locked, returns TRUE; otherwise FALSE.
			Zorder = 3, // Z-order position (layering) of the object; that is, the relative position of the overlapping objects, starting with 1 for the object that is most under the others.
			Ref1 = 4, // Reference of the cell under the upper-left corner of the object as text in R1C1 reference style; for a line or arc, returns the start point.
			x1 = 5, // X offset from the upper-left corner of the cell under the upper-left corner of the object, measured in points.
			y1 = 6, // Y offset from the upper-left corner of the cell under the upper-left corner of the object, measured in points.
			Ref2 = 7, // Reference of the cell under the lower-right corner of the object as text in R1C1 reference style; for a line or arc, returns the end point.
			x2 = 8, // X offset from the upper-left corner of the cell under the lower-right corner of the object, measured in points.
			y2 = 9, // Y offset from the upper-left corner of the cell under the lower-right corner of the object, measured in points.
			MacroName = 10, // Name, including the filename, of the macro assigned to the object. If no macro is assigned, returns FALSE.
			Placement = 11, // Number indicating how the object moves and sizes:
			//= 1 = Object moves and sizes with cells, // 
			//= 2 = Object moves with cells, // 
			//= 3 = Object is fixed, // 
			//=, // 
			//= Values 12 to 21 for type_num apply only to text boxesand buttons.If another type of object is selected, GET.OBJECT returns the #VALUE!error value., // 
			//= Type_num, // Returns
			Text = 12, // Text starting at start_num for count_num characters.
			Font = 13, // Font name of all text starting at start_num for count_num characters. If the text contains more than one font name, returns the #N/A error value.
			FontSize = 14, // Font size of all text starting at start_num for count_num characters. If the text contains more than one font size, returns the #N/A error value.
			Bold = 15, // If all text starting at start_num for count_num characters is bold, returns TRUE. If text contains only partial bold formatting, returns the #N/A error value.
			Italic = 16, // If all text starting at start_num for count_num characters is italic, returns TRUE. If text contains only partial italic formatting, returns the #N/A error value.
			Underlined = 17, // If all text starting at start_num for count_num characters is underlined, returns TRUE. If text contains only partial underline formatting, returns the #N/A error value.
			Strike = 18, // If all text starting at start_num for count_num characters is struck through, returns TRUE. If text contains only partial struck-through formatting, returns the #N/A error value.
			Outline = 19, // In Microsoft Excel for the Macintosh, if all text starting at start_num for count_num characters is outlined, returns TRUE. If text contains only partial outline formatting, returns the #N/A error value. Always returns FALSE in Microsoft Excel for Windows.
			Shadow = 20, // In Microsoft Excel for the Macintosh, if all text starting at start_num for count_num characters is shadowed, returns TRUE. If text contains only partial shadow formatting, returns the #N/A error value. Always returns FALSE in Microsoft Excel for Windows.
			Color = 21, // Number from 0 to 56 indicating the color of all text starting at start_num for count_num characters; if color is automatic, returns 0. If more than one color is used, returns the #N/A error value.
			//=, // 
			//= Values 22 to 25 for type_num also apply only to text boxesand buttons.If another type of object is selected, GET.OBJECT returns the #N / A error value., // 
			//= Type_num, // Returns
			AlignmentVertical = 23, // Number indicating the vertical alignment of text:
			//= 1 = Top, // 
			//= 2 = Center, // 
			//= 3 = Bottom, // 
			//= 4 = Justified, // 
			Orientation = 24, // Number indicating the orientation of text:
			//= 0 = Horizontal, // 
			//= 1 = Vertical, // 
			//= 2 = Upward, // 
			//= 3 = Downward, // 
			Automatic = 25, // If button or text box is set to automatic sizing, returns TRUE; otherwise FALSE.
			//=, // 
			//= The following values for type_num apply to all objects, except where indicated., // 
			//= Type_num, // Returns
			Border = 27, // Number indicating the type of the border or line:
			//= 0 = Custom, // 
			//= 1 = Automatic, // 
			//= 2 = None, // 
			BorderStyle = 28, // Number indicating the style of the border or line as shown in the Patterns tab in the Format Objects dialog box:
			//= 0 = None, // 
			//= 1 = Solid line, // 
			//= 2 = Dashed line, // 
			//= 3 = Dotted line, // 
			//= 4 = Dashed dotted line, // 
			//= 5 = Dashed double - dotted line, // 
			//= 6 = 50 % gray line, // 
			//= 7 = 75 % gray line, // 
			//= 8 = 25 % gray line, // 
			BorderColor = 29, // Number from 0 to 56 indicating the color of the border or line; if the border is automatic, returns 0.
			BorderWeight = 30, // Number indicating the weight of the border or line:
			//= 1 = Hairline, // 
			//= 2 = Thin, // 
			//= 3 = Medium, // 
			// 4 = Thick, // 
			Fill = 31, // Number indicating the type of fill:
			//= 0 = Custom, // 
			//= 1 = Automatic, // 
			//= 2 = None, // 
			FillPattern = 32, // Number from 1 to 18 indicating the fill pattern as shown in the Format Object dialog box.
			ForegroundColor = 33, // Number from 0 to 56 indicating the foreground color of the fill pattern; if the fill is automatic, returns 0. If the object is a line, returns the #N/A error value.
			BackgroundColor = 34, // Number from 0 to 56 indicating the background color of the fill pattern; if the fill is automatic, returns 0. If the object is a line, returns the #N/A error value.
			ArrowheadWidth = 35, // Number indicating the width of the arrowhead:
			//= 1 = Narrow, // 
			//= 2 = Medium, // 
			//= 3 = Wide, // 
			//= If the object is not a line, returns the #N / A error value., // 
			ArrowheadLength = 36, // Number indicating the length of the arrowhead:
			//= 1 = Short, // 
			//= 2 = Medium, // 
			//= 3 = Long, // 
			//= If the object is not a line, returns the #N / A error value., // 
			ArrowheadStyle = 37, // Number indicating the style of the arrowhead:
			//= 1 = No head, // 
			//= 2 = Open head, // 
			//= 3 = Closed head, // 
			//= 4 = Open double - ended head, // 
			//= 5 = Closed double - ended head, // 
			//= If the object is not a line, returns the #N / A error value., // 
			BorderRound = 38, // If the border has round corners, returns TRUE; if the corners are square, returns FALSE. If the object is a line, returns the #N/A error value.
			BorderShadow = 39, // If the border has a shadow, returns TRUE; if the border has no shadow, returns FALSE. If the object is a line, returns the #N/A error value.
			LockText = 40, // If the Lock Text check box in the Protection Tab of the Format Object dialog box is selected, returns TRUE; otherwise FALSE.
			Printed = 41, // If objects are set to be printed, returns TRUE; otherwise FALSE.
			//= 42, // The horizontal distance, measured in points, from the left edge of the active window to the left edge of the object. May be a negative number if the window is scrolled beyond the object.
			//= 43, // The vertical distance, measured in points, from the top edge of the active window to the top edge of the object. May be a negative number if the window is scrolled beyond the object.
			//= 44, // The horizontal distance, measured in points, from the left edge of the active window to the right edge of the object. May be a negative number if the window is scrolled beyond the object.
			//= 45, // The vertical distance, measured in points, from the top edge of the active window to the bottom edge of the object. May be a negative number if the window is scrolled beyond the object.
			//= 46, // The number of vertices in a polygon, or the #N/A error value if the object is not a polygon.
			PolygonVertices = 47, // A count_num by 2 array of vertex coordinates starting at start_num in a polygon's array of vertices.
			TextBoxReference = 48, // If the object is a text box, returns the cell reference that the text box is linked to. If the object is a control on a worksheet, returns the cell reference that the control's value is linked to. This information is returned as a string.
			ID = 49, // Returns the ID number of the object. For example, "Rectangle 5" returns 5. Note that the name of the object may not have this index in it if the object has been renamed by the user.
			Class = 50, // Returns the object's classname. For example, "Rectangle".
			Name = 51, // Returns the object name. By default, object names are the classname followed by the ID. For example, "Rectangle 1" is an object name, of which "Rectangle" is the classname, and 1 is the ID number. The object can also be renamed, in which case the name picked by the user is returned.
			//= 52, // Returns the distance from cell A1 to the Left of the object bounding rectangle in points
			//= 53, // Returns the distance from Cell A1 to the top of the object bounding rectangle in points
			Width = 54, // Returns the width of object bounding rectangle in points
			Height = 55, // Returns the height of object bounding rectangle in points
			Enabled = 56, // If the object is enabled, returns TRUE; otherwise, it returns FALSE.
			Shortcut = 57, // Returns the shortcut key assignment for the control object, as text.
			Default = 58, // Returns TRUE is the button control on a dialog sheet is the default button of the dialog; otherwise, returns FALSE
			//= 59, // Returns TRUE if the button control on the dialog sheet is clicked when the user presses the ESCAPE Key; otherwise, returns FALSE.
			//= 60, // Returns TRUE if the button control on a dialog sheet will close the dialog box when pressed; otherwise, returns FALSE
			//= 61, // Returns TRUE if the button control on a dialog sheet will be clicked when the user presses F1.
			Value = 62, // Returns the value of the control. For a check box or radio button, Returns 1 if it is selected, zero if it is not selected, or 2 if mixed. For a List box or dropdown box, returns the index number of the selected item, or zero if no item is selected. For a scroll bar, returns the numeric value of the scroll bar.
			ValueMinimum = 63, // Returns the minimum value that a scroll bar or spinner button can have
			ValueMaximum = 64, // Returns the maximum value that a scroll bar or spinner button can have
			ValueIncrement = 65, // Returns the step increment value added or subtracted from the value of a scroll bar or spinner. This value is used when the arrow buttons are pressed on the control.
			//= 66, // Returns the large, or "page" step increment value added or subtracted from the value of a scroll bar when it is clicked in the region between the thumb and the arrow buttons.
			InputType = 67, // Returns the input type allowed in an edit box control:
			//= 1 = Text, // 
			//= 2 = Integer, // 
			//= 3 = Number(what type), // 
			//= 4 = Cell reference, // 
			//= 5 = Formula, // 
			EditMultiline = 68, // Returns TRUE if the edit box control allows multi-line editing with wrapped text; otherwise, it returns FALSE.
			EditVerticalScroll = 69, // Returns TRUE if the edit box has a vertical scroll bar; otherwise, it returns FALSE.
			BoxID = 70, // Returns the object ID of the object that is linked to a list box or edit box. For a dropdown combo box that has an editable entry field, returns the object ID of itself. A dropdown box that can't be edited, returns FALSE.
			BoxEntries = 71, // Returns the number of entries in a List box, dropdown List box, or dropdown combo box.
			BoxText = 72, // Returns the text of the selected entry in a List box, dropdown List box, or dropdown combo box.
			BoxRange = 73, // Returns the range used to fill the entries in a List box, dropdown List box, or dropdown combo box, as text. If an empty string is returned, then the control isn't filled from a range.
			DrowdownLines = 74, // Returns the number of list lines displayed when a dropdown control is dropped.
			Is3D = 75, // Returns TRUE the object is displayed as 3-D; otherwise, it returns FALSE.
			//= 76, // Returns the Far East phonetic accelerator key as text. Used for Far East versions of Microsoft Excel.
			Selected = 77, // Returns the select status of the list box:
			//= 0 = single, // 
			//= 1 = simple multi - select, // 
			//= 2 = extended multi - select, // 
			ListBoxSelected = 78, // Returns an array of TRUE and FALSE values indicating which items are selected in a list box. If TRUE, the item is selected; If FALSE, the item is not selected.
			//= 79, // Returns TRUE if the add indent attribute is on for alignment. Returns FALSE if the add indent attribute is off for alignment. Used for only Far East versions of Microsoft Excel.
		};
	};

	// OPTIONS.CALCULATION(type_num, iter, max_num, max_change, update, precision, date_1904, calc_save, save_values)
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
			return Excel(xlcOptionsCalculation, type_num, iter, max_num, max_change,
				update, precision, date_1904, calc_save, save_values);
		}
	};

	// OPTIONS.VIEW(formula, status, notes, show_info, object_num, page_breaks, formulas, gridlines, color_num, headers, outline, zeros, hor_scroll, vert_scroll, sheet_tabs)
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
			return Excel(xlcOptionsView, formula, status, notes,
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
	// https://xlladdins.github.io/Excel4Macros/get.cell.html
	struct Cell {
		OPER ref;
		Cell(const OPER& ref = Excel(xlfActiveCell))
			: ref(ref)
		{ }
		Cell(REF ref)
			: ref(OPER(ref))
		{ }
		Cell(const char* ref)
			: ref(Excel(xlfTextref, OPER(ref), OPER(false))) // R1C1
		{ }

		// generic get
		OPER Get(int i) const
		{
			return Excel(xlfGetCell, OPER(i));
		}

		// Absolute reference of the upper-left cell in reference, as text in the current workspace reference style.
		OPER AbsRef() const
		{
			return Excel(xlfGetCell, OPER(1), ref);
		}
		// Same as TYPE(reference).
		OPER Type() const
		{
			return Excel(xlfType, ref);
		}
		// Contents of reference.
		OPER Contents() const
		{
			return Excel(xlfGetCell, OPER(5), ref);
		}
		// Formula in reference, as text, in either A1 or R1C1 style depending on the workspace setting.
		OPER Formula() const
		{
			return Excel(xlfGetCell, OPER(6), ref);
		}
		// Number format of the cell
		OPER FormatNumber() const
		{
			return Excel(xlfGetCell, OPER(7), ref);
		}
		Alignment Align() const
		{
			Alignment align;

			align.horiz_align = Excel(xlfGetCell, OPER(7), ref);
			align.vert_align = Excel(xlfGetCell, OPER(50), ref);
			align.orientation = Excel(xlfGetCell, OPER(51), ref);

			return align;
		}
		// 14 	If the cell is locked, returns TRUE; otherwise, returns FALSE.
		// 15 	If the cell's formula is hidden, returns TRUE; otherwise, returns FALSE.
		// Row height of cell, in points.
		OPER RowHeight() const
		{
			return Excel(xlfGetCell, OPER(17), ref);
		}
		// GET.CELL(44) - GET.CELL(42) to determine the width.
		OPER Width() const
		{
			return OPER(Get(44).val.num - Get(42).val.num);
		}
		// Name of font.
		OPER FontName() const
		{
			return Excel(xlfGetCell, OPER(18), ref);
		}
		// Size of font, in points.
		OPER FontSize() const
		{
			return Excel(xlfGetCell, OPER(19), ref);
		}
		// Style of the cell.
		OPER Style() const
		{
			return Excel(xlfGetCell, OPER(40), ref);
		}
		// 46 	If the cell contains a text note, returns TRUE; otherwise, returns FALSE.
		// This returns the note text
		OPER Note() const
		{
			return Excel(xlfGetNote, ref);
		}
		// Returns the name of the sheet in the form "[Book1]Sheet1".
		OPER Name() const
		{
			return Excel(xlfGetCell, OPER(62), ref);
		}
	};

	// prefer styles to explicit formatting
	struct Style {
		OPER name;
		Style(const char* name)
			: name(name)
		{ }
		// https://xlladdins.github.io/Excel4Macros/apply.style.html
		OPER Apply() const
		{
			return Excel(xlcApplyStyle, name);
		}
		static OPER Apply(const OPER& name)
		{
			return Excel(xlcApplyStyle, name);
		}
		// https://xlladdins.github.io/Excel4Macros/delete.style.html
		OPER Delete() const 
		{
			return Excel(xlcDeleteStyle, name);
		}
		static OPER Delete(const OPER& name)
		{
			return Excel(xlcDeleteStyle, name);
		}
		// define styles
		Style& FormatNumber(const char* format)
		{
			ensure(Excel(xlcDefineStyle, name, OPER(2), OPER(format)));

			return *this;
		}
		Style& FormatFont(const xll::FormatFont& format)
		{
			ensure(Excel(xlcDefineStyle, name, OPER(3),
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
			ensure(Excel(xlcDefineStyle, name, OPER(4),
				align.horiz_align,
				align.wrap,
				align.vert_align,
				align.orientation));

			return *this;
		}
		//!!! 4-7
		Style& Pattern(const OPER& fore, const OPER& back = Missing)
		{
			ensure(Excel(xlcDefineStyle, name, OPER(6), OPER(1), fore, back));

			return *this;
		}
	};

	// foreground and background cell colors
	// https://xlladdins.github.io/Excel4Macros/patterns.html
	inline OPER Patterns(OPER fore, OPER back)
	{
		return Excel(xlcPatterns, OPER(1), fore, back);
	}


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
			return Excel(xlcDefineName, name, ref, OPER(3), Missing, OPER(false), Missing, OPER(local));
		}
		OPER Delete()
		{
			return Excel(xlcDeleteName, name);
		}
		OPER Get()
		{
			return Excel(xlfGetName, name);
		}
		// xlcSetName is only for macro sheets
	};

	struct Document {
		static OPER Book()
		{
			return Excel(xlfGetDocument, OPER(88)); // Book
		}

		static OPER Sheet()
		{
			return Excel(xlfGetDocument, OPER(76)); // [Book]Sheet
		}
	};

	struct Workbook {
		// https://xlladdins.github.io/Excel4Macros/workbook.delete.html
		static OPER Delete(const OPER& name = Document::Sheet())
		{
			return Excel(xlcWorkbookDelete, name);
		}
		enum class Type {
			Worksheet = 1,
			Chart = 2,
			MacroSheet = 3,
			InternationallMacroSheet = 4,
			VisualBasicModule = 6,
			Dialog = 7
		};
		// https://xlladdins.github.io/Excel4Macros/workbook.insert.html
		static OPER Insert(enum Type type = Type::Worksheet)
		{
			OPER insert = Excel(xlcWorkbookInsert, OPER((int)type));
			if (!insert) {
				// empty workbook
				insert = New(type);
			}

			return insert;
		}
		// https://xlladdins.github.io/Excel4Macros/workbook.move.html
		static OPER Move(const OPER& name, int position = 1)
		{
			return Excel(xlcWorkbookMove, name, Document::Sheet(), OPER(position));
		}
		// https://xlladdins.github.io/Excel4Macros/new.html
		static OPER New(enum Type type = Type::Worksheet)
		{
			return Excel(xlcNew, OPER((int)type), Missing, Missing);
		}
		// https://xlladdins.github.io/Excel4Macros/workbook.name.html
		static OPER Rename(const OPER& name)
		{
			return Excel(xlcWorkbookName, Document::Sheet(), name);
		}
		// https://xlladdins.github.io/Excel4Macros/workbook.select.html
		static OPER Select(const OPER& name = Document::Sheet())
		{
			return Excel(xlcWorkbookSelect, name);
		}
	};
}
