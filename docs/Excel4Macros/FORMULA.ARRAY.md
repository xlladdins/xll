FORMULA.ARRAY

Enters a formula as an array formula in the range specified or in the
current selection. Equivalent to entering an array formula while
pressing CTRL+SHIFT+ENTER in Microsoft Excel for Windows or
COMMAND+ENTER in Microsoft Excel for the Macintosh.

**Syntax**

**FORMULA.ARRAY**(**formula\_text**, reference)

Formula\_text&nbsp;&nbsp;&nbsp;&nbsp;is the text you want to enter in
the array. For more information on formula\_text, see the first form of
FORMULA.

Reference&nbsp;&nbsp;&nbsp;&nbsp;specifies where formula\_text is
entered. It can be a reference to a cell on the active worksheet or an
external reference to a named workbook. Reference must be a R1C1-style
reference in text form. If reference is omitted, formula\_text is
entered in the active cell.

**Examples**

If the selection is D25:E25, the following macro formula enters the
array formula {=D22:E22+D23:E23} in the range D25:E25:

FORMULA.ARRAY("=R\[-3\]C:R\[-3\]C\[1\]+R\[-2\]C:R\[-2\]C\[1\]")

Regardless of the selection, the following macro formula enters the
array formula {=D22:E22+D23:E23} in the range D25:E25:

FORMULA.ARRAY("=R\[-3\]C:R\[-3\]C\[1\]+R\[-2\]C:R\[-2\]C\[1\]",
"R25C4:R25C5")

To use FORMULA.ARRAY to put an array in a specific workbook, specify the
name of the workbook as an external reference in the reference argument.
Using "\[SALES.XLS\]North\!R25C3:R25C4" as the reference argument in the
preceding example would enter the array in cells C25:D25 on the
worksheet named North in the workbook SALES.XLS. Using
"SALES\!R25C3:R25C4" as the reference argument would enter the array in
the same cells in the worksheet named SALES.

**Related Functions**

[FORMULA](FORMULA.md)&nbsp;&nbsp;&nbsp;Enters values into a cell or range or onto a
chart

[FORMULA.FILL](FORMULA.FILL.md)&nbsp;&nbsp;&nbsp;Enters a formula in the specified range



Return to [README](README.md)

