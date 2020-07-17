DEFINE.STYLE Syntax 1
=====================

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

Style\_text    is the name, as text, that you want to assign to the
style.

The following arguments are logical values corresponding to check boxes
in the Style dialog box. If an argument is TRUE, Microsoft Excel selects
the check box and uses the corresponding format of the active cell in
the style; if FALSE, Microsoft Excel clears the check box and omits
formatting descriptions for that attribute. If style\_text is omitted
and all selected cells have identical formatting, the default is TRUE;
if cells have different formatting, the default is FALSE.

Number    corresponds to the Number check box.

Font    corresponds to the Font check box.

Alignment    corresponds to the Alignment check box.

Border    corresponds to the Border check box.

Pattern    corresponds to the Pattern check box.

Protection    corresponds to the Protection check box.

**Related Functions**

DEFINE.STYLE Syntaxes 2-7

APPLY.STYLE   Applies a style to the selection

DELETE.STYLE   Deletes a cell style

MERGE.STYLES   Imports styles from another workbook into the active
workbook

Return to [top](#A)

DEFINE.STYLE Syntaxes 2 - 7
