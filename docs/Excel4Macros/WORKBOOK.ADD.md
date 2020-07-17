WORKBOOK.ADD

Macro SheetsOnly

Equivalent to clicking the Add button in the workbook contents window in
Microsoft Excel version 4.0. Moves a sheet between workbooks. For
Microsoft Excel version 5.0 or later use WORKBOOK.MOVE.

**Syntax**

**WORKBOOK.ADD**(**name\_array**, dest\_book, position\_num)

**WORKBOOK.ADD**?(name\_array, dest\_book, position\_num)

Name\_array    is the name of a sheet, or an array of names of sheets,
that you want to move.

Dest\_book     is the name of the workbook to which you want to add
name\_array. If dest\_book is omitted, it is assumed to be the active
workbook.

Position\_num    is a number that specifies the position of the sheet
within the workbook.

**Related Functions**

WORKBOOK.MOVE   Moves one or more sheets between workbooks or changes a
sheet's position within a workbook

WORKBOOK.COPY   Copies one or more documents from their current workbook
to another workbook


