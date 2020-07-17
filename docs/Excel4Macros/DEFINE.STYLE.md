DEFINE.STYLE
===========================

Equivalent to clicking the Define button in the Style dialog box, which
appears when you click the Style command on the Format menu. Creates and
changes cell styles. Use one of the following syntax forms of
DEFINE.STYLE to select cell formats for a new style or to alter the
formats of an existing style. Use syntax 1 of DEFINE.STYLE to define
styles based on the format of the active cell.

**Syntax 2**

Number format, using the arguments from the FORMAT.NUMBER function

**DEFINE.STYLE**(**style\_text, attribute\_num**, format\_text)

**Syntax 3**

Font format, using the arguments from the FORMAT.FONT and
FONT.PROPERTIES functions

**DEFINE.STYLE**(**style\_text, attribute\_num**, name\_text, size\_num,
bold, italic, underline, strike, color, outline, shadow, superscript,
subscript)

**Syntax 4**

Alignment, using the arguments from the ALIGNMENT function

**DEFINE.STYLE**(**style\_text**, **attribute\_num**, horiz\_align,
wrap, vert\_align, orientation)

**Syntax 5**

Border, using the arguments from the BORDER function

**DEFINE.STYLE**(**style\_text**, **attribute\_num**, left, right, top,
bottom, left\_color, right\_color, top\_color, bottom\_color)

**Syntax 6**

Pattern, using the arguments from the cell form of the PATTERNS function

**DEFINE.STYLE**(**style\_text, attribute\_num**, apattern, afore,
aback)

**Syntax 7**

Cell protection, using the arguments from the CELL.PROTECTION function

**DEFINE.STYLE**(**style\_text, attribute\_num**, locked, hidden)

Style\_text    is the name, as text, that you want to assign to the
style.

Attribute\_num    is a number from 2 to 7 that specifies which attribute
of the style, such as its font, alignment, or number format, you want to
designate with this function.

  -------------------- -----------------
  **Attribute\_num**   **Specifies**
  2                    Number format
  3                    Font format
  4                    Alignment
  5                    Border
  6                    Pattern
  7                    Cell protection
  -------------------- -----------------

**Remarks**

-   The remaining arguments are different for each form and are
    > identical to arguments in the corresponding function. For example,
    > form 2 of DEFINE.STYLE defines the number format of a style and
    > corresponds to the FORMAT.NUMBER function. The exception is form
    > 5, which does not include every argument for BORDER. For details
    > on the values you can use for these arguments, see the description
    > under the corresponding function.

-   If you define a style using one of these forms, then any attributes
    > you don\'t explicitly define are not changed.

>  

**Related Functions**

DEFINE.STYLE Syntax 1

ALIGNMENT   Aligns or wraps text in cells

APPLY.STYLE   Applies a style to the selection

BORDER   Adds a border to the selected cell or object

CELL.PROTECTION   Allows you to control cell protection and display

DELETE.STYLE   Deletes a cell style

FONT.PROPERTIES   Applies a font to the selection

FORMAT.NUMBER   Formats numbers, dates, and times in the selected cells

MERGE.STYLES   Imports styles from another workbook into the active
workbook

PATTERNS   Changes the appearance of the selected object

Return to [top](#A)

DELETE.ARROW
