# TABLE

Equivalent to clicking the Table command on the Data menu. Creates a
table based on the input values and formulas you define on a worksheet.
Use data tables to perform a "what-if" analysis by changing certain
constant values in your workbook to see how values in other cells are
affected.

**Syntax**

**TABLE**(row\_ref, column\_ref)

**TABLE**?(row\_ref, column\_ref)

Row\_ref&nbsp;&nbsp;&nbsp;&nbsp;specifies the one cell to use as the row
input for your table.

  - > Row\_ref should be either an external reference to a single cell
    > on the active worksheet, such as \!$A$1 or \!Price, or an
    > R1C1-style reference to a single cell in the form of text, such as
    > "R1C1", "R\[-1\]C\[-1\]", or "Price".

  - > If row\_ref is an R1C1-style reference, it is assumed to be
    > relative to the active cell in the selection.


Column\_ref&nbsp;&nbsp;&nbsp;&nbsp;specifies the one cell to use as the
column input for your table. Column\_ref is subject to the same
restrictions as row\_ref.



Return to [README](README.md#T)

