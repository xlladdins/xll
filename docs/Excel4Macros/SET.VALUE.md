SET.VALUE

Changes the value of a cell or cells on the macro sheet (not the
worksheet) without changing any formulas entered in those cells. Use
SET.VALUE to assign initial values and to store values during the
calculation of a macro. SET.VALUE is especially useful for initializing
a dialog box and the conditional test in a WHILE loop. SET.VALUE assigns
values to a specific reference or to the name of a reference that has
already been defined. For information about creating a new name or
entering data on a worksheet, see "Remarks" later in this topic.

**Syntax**

**SET.VALUE**(**reference**, values)

Reference&nbsp;&nbsp;&nbsp;&nbsp;specifies the cell or cells on the
macro sheet to which you want to assign a new value or values. If the
cell is empty, enters the value in the cell.

  - > If a cell in reference previously contained a formula, the formula
    > is not changed, but the value of the cell might change. See the
    > second example following.

  - > If reference is a reference to a range of cells, rather than to a
    > single cell, then values should be an array of the same size. If
    > not, Microsoft Excel expands it into multiple values using the
    > normal rules for expanding arrays. See the third example
    > following.

> &nbsp;

Values&nbsp;&nbsp;&nbsp;&nbsp;is the value or set of values to which you
want to assign the cell or cells in reference.

**Remarks**

Consider the following guidelines as you choose a function to set values
on a worksheet or macro sheet:

  - > Use SET.VALUE to assign initial values to a reference (including
    > names that have already been defined) on a macro sheet, and to
    > store values during the calculation of a macro.

  - > Use FORMULA to enter values in a worksheet cell.

  - > Use SET.NAME to change the value of a name on a macro sheet (the
    > name is created if it does not already exist). For more
    > information, see SET.NAME.

  - > Use DEFINE.NAME to create or change the value of a name on a
    > worksheet.

> &nbsp;

**Examples**

The following macro formula changes the value of cell A1 on the macro
sheet to 1:

SET.VALUE($A$1, 1)

Suppose the name TempAverage refers to a cell containing the formula
AVERAGE(Temp1, Temp2, Temp3). The following formula assigns the value 99
to this cell, even if the average of the arguments is not 99, without
changing the formula in TempAverage:

SET.VALUE(TempAverage, 99)

The preceding formula is useful if a WHILE loop or some other
conditional test depends on TempAverage and you want to force the
conditional test to have a particular result. Of course, TempAverage is
restored to its correct value as soon as it is recalculated. (Recall
that unlike formulas in a worksheet, formulas in a macro sheet are not
recalculated until the macro actually uses them.)

The following macro formula stores the values 1, 2, 3, and 4 in cells
A1:B2:

SET.VALUE($A$1:$B$2, {1, 2;3, 4})

**Related Functions**

[DEFINE.NAME](DEFINE.NAME.md)&nbsp;&nbsp;&nbsp;Defines a name on the active worksheet or
macro sheet

[FORMULA](FORMULA.md)&nbsp;&nbsp;&nbsp;Enters values into a cell or range or onto a
chart

[SET.NAME](SET.NAME.md)&nbsp;&nbsp;&nbsp;Defines a name as a value



Return to [README](README.md)

