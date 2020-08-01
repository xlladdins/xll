# FORMULA

Enters a formula in the active cell or in a reference. There are two
syntax forms of this function. Use syntax 1 to enter numbers, text,
references, and formulas in a worksheet. Although syntax 1 can also be
used to enter values on a macro sheet, you will not generally use
FORMULA for this purpose. Use syntax 2 to enter a formula in a chart.
For information about setting values on a macro sheet, see "Remarks" in
the following topics.

Syntax 1&nbsp;&nbsp;&nbsp;Enters numbers, text, references, and formulas
in a worksheet

Syntax 2&nbsp;&nbsp;&nbsp;Enters formulas in a chart


# FORMULA Syntax 1

Enters a formula in the active cell or in a reference. If the active
sheet is a worksheet, using FORMULA is equivalent to entering
formula\_text in the cell specified by reference. Formula\_text is
entered just as if you typed it in the formula bar.

There are two syntax forms of this function. Use syntax 1 to enter
numbers, text, references, and formulas in a worksheet. Although syntax
1 can also be used to enter values on a macro sheet, you will not
generally use FORMULA for this purpose. Use syntax 2 to enter a formula
in a chart. For information about setting values on a macro sheet, see
"Remarks" later in this topic.

**Syntax**

**FORMULA**(**formula\_text**, reference)

Formula\_text&nbsp;&nbsp;&nbsp;&nbsp;can be text, a number, a reference,
or a formula in the form of text, or a reference to a cell containing
any of the above.

  - > If formula\_text contains references, they must be R1C1-style
    > references, such as "=RC\[1\]\*(1+R1C1)". If you are recording a
    > macro when you enter a formula, Microsoft Excel converts A1-style
    > references to R1C1-style references. For example, if you enter the
    > formula =B2\*(1+$A$1) in cell C2 while recording, Microsoft Excel
    > records that action as =FORMULA("=RC\[-1\]\*(1+R1C1)").

  - > If formula\_text is a formula, the formula is entered. Text
    > arguments must be surrounded by double sets of quotation marks.
    > For example, to enter the formula =IF($A$1="Hello World", 1, 0) in
    > the active cell with the FORMULA function, you would use the
    > formula FORMULA("=IF(R1C1=""Hello World"", 1, 0)")

  - > If formula\_text is a number, text, or logical value, the value is
    > entered as a constant.

Reference&nbsp;&nbsp;&nbsp;&nbsp;specifies where formula\_text is to be
entered. It can be a reference to a cell in the active workbook or an
external reference to a workbook. If reference is omitted, formula\_text
is entered in the active cell.

**Remarks**

Consider the following guidelines as you choose a function to set values
on a worksheet or macro sheet:

  - > Use FORMULA to enter formulas and change values in a worksheet
    > cell.

  - > SET.VALUE changes values on the macro sheet. Use SET.VALUE to
    > assign initial values to a reference and to store values during
    > the calculation of the macro.

  - > SET.NAME creates names on the macro sheet. Use SET.NAME to create
    > a name and immediately assign a value to the name.


**Examples**

If the active sheet is a worksheet, the following macro formula enters
the number constant 523 in the active cell:

FORMULA(523)

If the active sheet is a worksheet, the following macro formula enters
the result of the INPUT function in cell A5:

FORMULA(INPUT("Enter a formula:", 0), \!$A$5)

If you're using R1C1-style references and the active sheet is a
worksheet, the following macro formula enters the formula
=RC\[-1\]\*(1+R1C1) in the active cell:

FORMULA("=RC\[-1\]\*(1+R1C1)")

If the active sheet is a worksheet, the following macro formulas enter
the number 1000 in the cell two rows down and three columns right from
the active cell. The R1C1-style formula is shorter, but the OFFSET
method may provide faster performance in larger macro sheets.

FORMULA(1000, OFFSET(ACTIVE.CELL(), 2, 3))

FORMULA(1000, "R\[2\]C\[3\]")

The following macro formula enters the phrase "Year to Date" in cell B4
on the sheet named SALES 1993:

FORMULA("Year to Date", 'SALES 1993'\!B4)

**Related Functions**

[FORMULA.ARRAY](FORMULA.ARRAY.md)&nbsp;&nbsp;&nbsp;Enters an array

[FORMULA.FILL](FORMULA.FILL.md)&nbsp;&nbsp;&nbsp;Enters a formula in the specified range

[SET.VALUE](SET.VALUE.md)&nbsp;&nbsp;&nbsp;Sets the value of a cell on a macro sheet

[FORMULA](FORMULA.md) Syntax 2&nbsp;&nbsp;&nbsp;Enters formulas in a chart


# FORMULA Syntax 2

Enters a text label or SERIES formula in a chart. To enter formulas on a
worksheet or macro sheet, use syntax 1 of this function.

**Syntax**

**FORMULA**(**formula\_text**)

Formula\_text&nbsp;&nbsp;&nbsp;&nbsp;is the text label or SERIES formula
you want to enter into the chart.

|                                                                                                                             |                                                             |
| --------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------- |
| **If**                                                                                                                      | **Then**                                                    |
| Formula\_text can be treated as a text label and the current selection is a text label                                      | The selected text label is replaced with formula\_text.     |
| Formula\_text can be treated as a text label and there is no current selection or the current selection is not a text label | Formula\_text creates a new unattached text label.          |
| Formula\_text can be treated as a SERIES formula and the current selection is a SERIES formula                              | The selected SERIES formula is replaced with formula\_text. |
| Formula\_text can be treated as a SERIES formula and the current selection is not a SERIES formula                          | Formula\_text creates a new SERIES formula.                 |

**Remarks**

You would normally use the EDIT.SERIES function to create or edit a
chart series. For more information, see EDIT.SERIES.

**Example**

The following macro formula enters a SERIES formula on the chart. If the
current selection is a SERIES formula, it is replaced:

FORMULA("=SERIES(""Title"", , {1, 2, 3}, 1)")

**Related Functions**

[EDIT.SERIES](EDIT.SERIES.md)&nbsp;&nbsp;&nbsp;Creates or changes a chart series

[FORMULA](FORMULA.md), Syntax 1&nbsp;&nbsp;&nbsp;Enters numbers, text, references, and
formulas in a worksheet



Return to [README](README.md#F)

# FORMULA

Enters a formula in the active cell or in a reference. There are two
syntax forms of this function. Use syntax 1 to enter numbers, text,
references, and formulas in a worksheet. Although syntax 1 can also be
used to enter values on a macro sheet, you will not generally use
FORMULA for this purpose. Use syntax 2 to enter a formula in a chart.
For information about setting values on a macro sheet, see "Remarks" in
the following topics.

Syntax 1&nbsp;&nbsp;&nbsp;Enters numbers, text, references, and formulas
in a worksheet

Syntax 2&nbsp;&nbsp;&nbsp;Enters formulas in a chart


# FORMULA Syntax 1

Enters a formula in the active cell or in a reference. If the active
sheet is a worksheet, using FORMULA is equivalent to entering
formula\_text in the cell specified by reference. Formula\_text is
entered just as if you typed it in the formula bar.

There are two syntax forms of this function. Use syntax 1 to enter
numbers, text, references, and formulas in a worksheet. Although syntax
1 can also be used to enter values on a macro sheet, you will not
generally use FORMULA for this purpose. Use syntax 2 to enter a formula
in a chart. For information about setting values on a macro sheet, see
"Remarks" later in this topic.

**Syntax**

**FORMULA**(**formula\_text**, reference)

Formula\_text&nbsp;&nbsp;&nbsp;&nbsp;can be text, a number, a reference,
or a formula in the form of text, or a reference to a cell containing
any of the above.

  - > If formula\_text contains references, they must be R1C1-style
    > references, such as "=RC\[1\]\*(1+R1C1)". If you are recording a
    > macro when you enter a formula, Microsoft Excel converts A1-style
    > references to R1C1-style references. For example, if you enter the
    > formula =B2\*(1+$A$1) in cell C2 while recording, Microsoft Excel
    > records that action as =FORMULA("=RC\[-1\]\*(1+R1C1)").

  - > If formula\_text is a formula, the formula is entered. Text
    > arguments must be surrounded by double sets of quotation marks.
    > For example, to enter the formula =IF($A$1="Hello World", 1, 0) in
    > the active cell with the FORMULA function, you would use the
    > formula FORMULA("=IF(R1C1=""Hello World"", 1, 0)")

  - > If formula\_text is a number, text, or logical value, the value is
    > entered as a constant.

Reference&nbsp;&nbsp;&nbsp;&nbsp;specifies where formula\_text is to be
entered. It can be a reference to a cell in the active workbook or an
external reference to a workbook. If reference is omitted, formula\_text
is entered in the active cell.

**Remarks**

Consider the following guidelines as you choose a function to set values
on a worksheet or macro sheet:

  - > Use FORMULA to enter formulas and change values in a worksheet
    > cell.

  - > SET.VALUE changes values on the macro sheet. Use SET.VALUE to
    > assign initial values to a reference and to store values during
    > the calculation of the macro.

  - > SET.NAME creates names on the macro sheet. Use SET.NAME to create
    > a name and immediately assign a value to the name.


**Examples**

If the active sheet is a worksheet, the following macro formula enters
the number constant 523 in the active cell:

FORMULA(523)

If the active sheet is a worksheet, the following macro formula enters
the result of the INPUT function in cell A5:

FORMULA(INPUT("Enter a formula:", 0), \!$A$5)

If you're using R1C1-style references and the active sheet is a
worksheet, the following macro formula enters the formula
=RC\[-1\]\*(1+R1C1) in the active cell:

FORMULA("=RC\[-1\]\*(1+R1C1)")

If the active sheet is a worksheet, the following macro formulas enter
the number 1000 in the cell two rows down and three columns right from
the active cell. The R1C1-style formula is shorter, but the OFFSET
method may provide faster performance in larger macro sheets.

FORMULA(1000, OFFSET(ACTIVE.CELL(), 2, 3))

FORMULA(1000, "R\[2\]C\[3\]")

The following macro formula enters the phrase "Year to Date" in cell B4
on the sheet named SALES 1993:

FORMULA("Year to Date", 'SALES 1993'\!B4)

**Related Functions**

[FORMULA.ARRAY](FORMULA.ARRAY.md)&nbsp;&nbsp;&nbsp;Enters an array

[FORMULA.FILL](FORMULA.FILL.md)&nbsp;&nbsp;&nbsp;Enters a formula in the specified range

[SET.VALUE](SET.VALUE.md)&nbsp;&nbsp;&nbsp;Sets the value of a cell on a macro sheet

[FORMULA](FORMULA.md) Syntax 2&nbsp;&nbsp;&nbsp;Enters formulas in a chart


# FORMULA Syntax 2

Enters a text label or SERIES formula in a chart. To enter formulas on a
worksheet or macro sheet, use syntax 1 of this function.

**Syntax**

**FORMULA**(**formula\_text**)

Formula\_text&nbsp;&nbsp;&nbsp;&nbsp;is the text label or SERIES formula
you want to enter into the chart.

|                                                                                                                             |                                                             |
| --------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------- |
| **If**                                                                                                                      | **Then**                                                    |
| Formula\_text can be treated as a text label and the current selection is a text label                                      | The selected text label is replaced with formula\_text.     |
| Formula\_text can be treated as a text label and there is no current selection or the current selection is not a text label | Formula\_text creates a new unattached text label.          |
| Formula\_text can be treated as a SERIES formula and the current selection is a SERIES formula                              | The selected SERIES formula is replaced with formula\_text. |
| Formula\_text can be treated as a SERIES formula and the current selection is not a SERIES formula                          | Formula\_text creates a new SERIES formula.                 |

**Remarks**

You would normally use the EDIT.SERIES function to create or edit a
chart series. For more information, see EDIT.SERIES.

**Example**

The following macro formula enters a SERIES formula on the chart. If the
current selection is a SERIES formula, it is replaced:

FORMULA("=SERIES(""Title"", , {1, 2, 3}, 1)")

**Related Functions**

[EDIT.SERIES](EDIT.SERIES.md)&nbsp;&nbsp;&nbsp;Creates or changes a chart series

[FORMULA](FORMULA.md), Syntax 1&nbsp;&nbsp;&nbsp;Enters numbers, text, references, and
formulas in a worksheet



Return to [README](README.md#F)

