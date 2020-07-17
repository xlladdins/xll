SELECT Syntax 1
===============

Equivalent to selecting cells or changing the active cell. There are
three syntax forms of SELECT. Use syntax 1 to select a cell on a
worksheet or macro sheet; use one of the other syntax forms to select
worksheet or macro sheet objects or chart items.

**Syntax**

**SELECT**(selection, active\_cell)

Selection    is the cell or range of cells you want to select. Selection
can be a reference to the active worksheet, such as !\$A\$1:\$A\$3 or
!Sales, or an R1C1-style reference to a cell or range relative to the
active cell in the current selection, such as
\"R\[-1\]C\[-1\]:R\[1\]C\[1\]\". The reference must be in text form. If
selection is omitted, the current selection is used.

Active\_cell    is the cell in selection you want to make the active
cell. Active\_cell can be a reference to a single cell on the active
worksheet, such as !\$A\$1, or an R1C1-style reference relative to the
active cell, such as \"R\[-1\]C\[-1\]\". The reference must be in text
form. If active\_cell is omitted, SELECT makes the cell in the
upper-left corner of selection the active cell.

**Remarks**

-   Active\_cell must be within selection. If it is not, an error
    > message is displayed and SELECT returns the \#VALUE! error value.

-   If you are recording a macro using relative references, Microsoft
    > Excel records the action using R1C1-style relative references in
    > the form of text.

-   If you are recording using absolute references, Microsoft Excel
    > records the action using R1C1-style absolute references in the
    > form of text.

-   You cannot give an external reference to a specific sheet as the
    > selection argument. The sheet on which you want to make a
    > selection must be active when you use SELECT. Use FORMULA.GOTO to
    > make a selection on another sheet in the same workbook or in
    > another workbook.

>  

**Tip**   You can enter data in a cell without selecting the cell by
using the reference arguments to the CUT, COPY, or FORMULA functions.

**Examples**

The following macro formula selects cells C3:E5 on the active worksheet
and makes C5 the active cell:

SELECT(!\$C\$3:\$E\$5, !\$C\$5)

If the active cell is C3, the following macro formula selects cells
E5:G7 and makes cell F6 the active cell in the selection:

SELECT(\"R\[2\]C\[2\]:R\[4\]C\[4\]\", \"R\[1\]C\[1\]\")

You can also make multiple nonadjacent selections with SELECT. The
following macro formula selects a number of nonadjacent ranges:

SELECT(\"R1C1, R3C2:R4C3, R8C4:R10C5\")

The following sequence of macro formulas moves the active cell right,
left, down, and up within the selection, just as TAB, SHIFT+TAB, ENTER,
and SHIFT+ENTER do:

SELECT(, \"RC\[1\]\")

SELECT(, \"RC\[-1\]\")

SELECT(, \"R\[1\]C\")

SELECT(, \"R\[-1\]C\")

Use SELECT with the OFFSET function to select a new range a specified
distance away from the current range. For example, the following macro
formula selects a range that is the same size as the current range, one
column over:

SELECT(OFFSET(SELECTION(), 0, 1))

**Related Functions**

ACTIVE.CELL   Returns the reference of the active cell

SELECT.SPECIAL   Selects a group of cells belonging to a category

SELECTION   Returns the reference of the selection

SELECT Syntax 2   Selects objects on worksheets

SELECT Syntax 3   Selects chart objects

Return to [top](#Q)

SELECT Syntax 2
