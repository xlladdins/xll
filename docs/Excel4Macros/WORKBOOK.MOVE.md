WORKBOOK.MOVE

Equivalent to clicking the Move or Copy Sheet command on the Edit menu.
Moves one or more sheets between workbooks or changes a sheet's position
within a workbook.

**Syntax**

**WORKBOOK.MOVE**(**name\_array**, dest\_book, position\_num)

**WORKBOOK.MOVE**?(name\_array, dest\_book, position\_num)

Name\_array    is the name of a sheet or an array of names of sheets in
the active workbook that you want to move.

Dest\_book    is the name of the workbook to which you want to move
name\_array. If dest\_book is omitted, WORKBOOK.MOVE moves the sheet out
of the workbook and puts it in a new separate workbook. If dest\_book is
the same as the current book, then the sheet is moved within the
workbook.

Position\_num    is a number that specifies the target position for the
sheet within dest\_book. The first position is 1.

  - > If position\_num is specified, Microsoft Excel inserts the sheet
    > at the specified position in the workbook.

  - > If position\_num is omitted, Microsoft Excel moves the sheet to
    > the last position in the workbook. If you move the last sheet out
    > of a workbook, the workbook closes.

>  

**Related Function**

[WORKBOOK.COPY](WORKBOOK.COPY.md)   Copies one or more documents from their current workbook
into another workbook



Return to [README](README.md)

