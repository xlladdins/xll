# DEFINE.STYLE

Equivalent to clicking the Define button in the Style dialog box, which
appears when you click the Style command on the Format menu. Creates and
changes cell styles. There are seven syntax forms of this function. Use
syntax 1 of DEFINE.STYLE to define styles based on the format of the
active cell. To create a style by specifying number, font, and other
formats, use syntaxes 2 through 7 of DEFINE.STYLE.

Syntax 1

Syntaxes 2-7


# DEFINE.STYLE Syntax 1

Equivalent to clicking the Define button in the Style dialog box, which
appears when you click the Style command on the Format menu. Creates and
changes cell styles. There are seven syntax forms of this function. Use
syntax 1 of DEFINE.STYLE to define styles based on the format of the
active cell. To create a style by specifying number, font, and other
formats, use syntaxes 2 through 7 of DEFINE.STYLE.

**Syntax**

**DEFINE.STYLE**(**style\_text**, number, font, alignment, border,
pattern, protection)

**DEFINE.STYLE**?(style\_text, number, font, alignment, border, pattern,
protection)

Style\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name, as text, that you want
to assign to the style.

The following arguments are logical values corresponding to check boxes
in the Style dialog box. If an argument is TRUE, Microsoft Excel selects
the check box and uses the corresponding format of the active cell in
the style; if FALSE, Microsoft Excel clears the check box and omits
formatting descriptions for that attribute. If style\_text is omitted
and all selected cells have identical formatting, the default is TRUE;
if cells have different formatting, the default is FALSE.

Number&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Number check box.

Font&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Font check box.

Alignment&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Alignment check box.

Border&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Border check box.

Pattern&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Pattern check box.

Protection&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Protection check
box.

**Related Functions**

[DEFINE.STYLE](DEFINE.STYLE.md) Syntaxes 2-7

[APPLY.STYLE](APPLY.STYLE.md)&nbsp;&nbsp;&nbsp;Applies a style to the selection

[DELETE.STYLE](DELETE.STYLE.md)&nbsp;&nbsp;&nbsp;Deletes a cell style

[MERGE.STYLES](MERGE.STYLES.md)&nbsp;&nbsp;&nbsp;Imports styles from another workbook into
the active workbook


# DEFINE.STYLE Syntaxes 2 - 7

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

Style\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name, as text, that you want
to assign to the style.

Attribute\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 2 to 7 that
specifies which attribute of the style, such as its font, alignment, or
number format, you want to designate with this function.

|                    |                 |
| ------------------ | --------------- |
| **Attribute\_num** | **Specifies**   |
| 2                  | Number format   |
| 3                  | Font format     |
| 4                  | Alignment       |
| 5                  | Border          |
| 6                  | Pattern         |
| 7                  | Cell protection |

**Remarks**

  - > The remaining arguments are different for each form and are
    > identical to arguments in the corresponding function. For example,
    > form 2 of DEFINE.STYLE defines the number format of a style and
    > corresponds to the FORMAT.NUMBER function. The exception is form
    > 5, which does not include every argument for BORDER. For details
    > on the values you can use for these arguments, see the description
    > under the corresponding function.

  - > If you define a style using one of these forms, then any
    > attributes you don't explicitly define are not changed.


**Related Functions**

[DEFINE.STYLE](DEFINE.STYLE.md) Syntax 1

[ALIGNMENT](ALIGNMENT.md)&nbsp;&nbsp;&nbsp;Aligns or wraps text in cells

[APPLY.STYLE](APPLY.STYLE.md)&nbsp;&nbsp;&nbsp;Applies a style to the selection

[BORDER](BORDER.md)&nbsp;&nbsp;&nbsp;Adds a border to the selected cell or object

[CELL.PROTECTION](CELL.PROTECTION.md)&nbsp;&nbsp;&nbsp;Allows you to control cell protection
and display

[DELETE.STYLE](DELETE.STYLE.md)&nbsp;&nbsp;&nbsp;Deletes a cell style

[FONT.PROPERTIES](FONT.PROPERTIES.md)&nbsp;&nbsp;&nbsp;Applies a font to the selection

[FORMAT.NUMBER](FORMAT.NUMBER.md)&nbsp;&nbsp;&nbsp;Formats numbers, dates, and times in the
selected cells

[MERGE.STYLES](MERGE.STYLES.md)&nbsp;&nbsp;&nbsp;Imports styles from another workbook into
the active workbook

[PATTERNS](PATTERNS.md)&nbsp;&nbsp;&nbsp;Changes the appearance of the selected object



Return to [README](README.md#D)

# DEFINE.STYLE

Equivalent to clicking the Define button in the Style dialog box, which
appears when you click the Style command on the Format menu. Creates and
changes cell styles. There are seven syntax forms of this function. Use
syntax 1 of DEFINE.STYLE to define styles based on the format of the
active cell. To create a style by specifying number, font, and other
formats, use syntaxes 2 through 7 of DEFINE.STYLE.

Syntax 1

Syntaxes 2-7


# DEFINE.STYLE Syntax 1

Equivalent to clicking the Define button in the Style dialog box, which
appears when you click the Style command on the Format menu. Creates and
changes cell styles. There are seven syntax forms of this function. Use
syntax 1 of DEFINE.STYLE to define styles based on the format of the
active cell. To create a style by specifying number, font, and other
formats, use syntaxes 2 through 7 of DEFINE.STYLE.

**Syntax**

**DEFINE.STYLE**(**style\_text**, number, font, alignment, border,
pattern, protection)

**DEFINE.STYLE**?(style\_text, number, font, alignment, border, pattern,
protection)

Style\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name, as text, that you want
to assign to the style.

The following arguments are logical values corresponding to check boxes
in the Style dialog box. If an argument is TRUE, Microsoft Excel selects
the check box and uses the corresponding format of the active cell in
the style; if FALSE, Microsoft Excel clears the check box and omits
formatting descriptions for that attribute. If style\_text is omitted
and all selected cells have identical formatting, the default is TRUE;
if cells have different formatting, the default is FALSE.

Number&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Number check box.

Font&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Font check box.

Alignment&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Alignment check box.

Border&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Border check box.

Pattern&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Pattern check box.

Protection&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Protection check
box.

**Related Functions**

[DEFINE.STYLE](DEFINE.STYLE.md) Syntaxes 2-7

[APPLY.STYLE](APPLY.STYLE.md)&nbsp;&nbsp;&nbsp;Applies a style to the selection

[DELETE.STYLE](DELETE.STYLE.md)&nbsp;&nbsp;&nbsp;Deletes a cell style

[MERGE.STYLES](MERGE.STYLES.md)&nbsp;&nbsp;&nbsp;Imports styles from another workbook into
the active workbook


# DEFINE.STYLE Syntaxes 2 - 7

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

Style\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name, as text, that you want
to assign to the style.

Attribute\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 2 to 7 that
specifies which attribute of the style, such as its font, alignment, or
number format, you want to designate with this function.

|                    |                 |
| ------------------ | --------------- |
| **Attribute\_num** | **Specifies**   |
| 2                  | Number format   |
| 3                  | Font format     |
| 4                  | Alignment       |
| 5                  | Border          |
| 6                  | Pattern         |
| 7                  | Cell protection |

**Remarks**

  - > The remaining arguments are different for each form and are
    > identical to arguments in the corresponding function. For example,
    > form 2 of DEFINE.STYLE defines the number format of a style and
    > corresponds to the FORMAT.NUMBER function. The exception is form
    > 5, which does not include every argument for BORDER. For details
    > on the values you can use for these arguments, see the description
    > under the corresponding function.

  - > If you define a style using one of these forms, then any
    > attributes you don't explicitly define are not changed.


**Related Functions**

[DEFINE.STYLE](DEFINE.STYLE.md) Syntax 1

[ALIGNMENT](ALIGNMENT.md)&nbsp;&nbsp;&nbsp;Aligns or wraps text in cells

[APPLY.STYLE](APPLY.STYLE.md)&nbsp;&nbsp;&nbsp;Applies a style to the selection

[BORDER](BORDER.md)&nbsp;&nbsp;&nbsp;Adds a border to the selected cell or object

[CELL.PROTECTION](CELL.PROTECTION.md)&nbsp;&nbsp;&nbsp;Allows you to control cell protection
and display

[DELETE.STYLE](DELETE.STYLE.md)&nbsp;&nbsp;&nbsp;Deletes a cell style

[FONT.PROPERTIES](FONT.PROPERTIES.md)&nbsp;&nbsp;&nbsp;Applies a font to the selection

[FORMAT.NUMBER](FORMAT.NUMBER.md)&nbsp;&nbsp;&nbsp;Formats numbers, dates, and times in the
selected cells

[MERGE.STYLES](MERGE.STYLES.md)&nbsp;&nbsp;&nbsp;Imports styles from another workbook into
the active workbook

[PATTERNS](PATTERNS.md)&nbsp;&nbsp;&nbsp;Changes the appearance of the selected object



Return to [README](README.md#D)

