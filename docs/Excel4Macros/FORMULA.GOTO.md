FORMULA.GOTO

Equivalent to clicking the Go To command on the Edit menu or to pressing
F5. Scrolls through the worksheet and selects a named area or reference.
Use FORMULA.GOTO to select a range on any open workbook; use SELECT to
select a range on the active workbook.

**Syntax**

**FORMULA.GOTO**(reference, corner)

**FORMULA.GOTO**?(reference, corner)

Reference    specifies where to scroll and what to select.

  - > Reference should be either an external reference to a workbook, an
    > R1C1-style reference in the form of text (see the second example
    > following), or a name.

  - > If the Go To command has already been carried out, reference is
    > optional. If reference is omitted, it is assumed to be the
    > reference of the cells you selected before the previous Go To
    > command or FORMULA.GOTO macro function was carried out. This
    > feature distinguishes FORMULA.GOTO from SELECT.

>  

Corner    is a logical value that specifies whether to scroll through
the window so that the upper-left cell in reference is in the upper-left
corner of the active window. If corner is TRUE, Microsoft Excel places
reference in the upper-left corner of the window; if FALSE or omitted,
Microsoft Excel scrolls through normally.

**Tip**   Microsoft Excel keeps a list of the cells you've selected with
previous FORMULA.GOTO functions or Go To commands. When you use
FORMULA.GOTO with GET.WORKSPACE(41), which returns a horizontal array of
previous Go To selections, you can backtrack through multiple previous
selections. See the last example below.

**Remarks**

  - > If you are recording a macro when you click the Go To command, the
    > reference you enter in the Reference box of the Go To dialog box
    > is recorded as text in the R1C1 reference style.

  - > If you are recording a macro when you double-click a cell that has
    > precedents on another worksheet, Microsoft Excel records a
    > FORMULA.GOTO function.

**Examples**

Each of the following macro formulas goes to cell A1 on the active
worksheet:

FORMULA.GOTO(\!$A$1)

FORMULA.GOTO("R1C1")

Each of the following macro formulas goes to the cells named Sales on
the active worksheet and scrolls through the worksheet so that the
upper-left corner of Sales is in the upper-left corner of the window:

FORMULA.GOTO(\!Sales, TRUE)

FORMULA.GOTO("Sales", TRUE)

The following macro formula goes to the cells that were selected by the
third most recent FORMULA.GOTO function or Go To command:

FORMULA.GOTO(INDEX(GET.WORKSPACE(41), 1, 3))

**Related Functions**

[GOTO](GOTO.md)   Directs macro execution to another cell

[HSCROLL](HSCROLL.md)   Horizontally scrolls through a sheet by percentage or by
column or row number

[SELECT](SELECT.md)   Selects a cell, worksheet object, or chart item

[VSCROLL](VSCROLL.md)   Vertically scrolls through a sheet by percentage or by column
or row number



Return to [README](README.md)

