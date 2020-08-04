# VSCROLL

Vertically scrolls through the active sheet by percentage or by row
number.

**Syntax**

**VSCROLL**(**position**, row\_logical)

Position&nbsp;&nbsp;&nbsp;&nbsp;specifies the row you want to scroll to.
Position can be an integer representing the row number or a fraction or
percentage representing the vertical position of the row in the sheet.
If position is 0, VSCROLL scrolls through your sheet to its top edge,
which is row 1. If position is 1, VSCROLL scrolls through your sheet to
its bottom edge, which is row 16, 384 in Microsoft Excel 95 or earlier,
or row 65,536 in Microsoft Excel 97 or later. For charts that do not
size with the window, use a fraction or percentage.

Row\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying how
the function scrolls.

  - > If row\_logical is TRUE, VSCROLL scrolls through the sheet to row
    > position.

  - > If row\_logical is FALSE or omitted, VSCROLL scrolls through the
    > sheet to the vertical position represented by the fraction
    > position.


**Remarks**

  - > To scroll to a specific row n, either use VSCROLL(n, TRUE) or
    > VSCROLL(n/16384) in Microsoft Excel 95 or earlier; in Microsoft
    > Excel 97 or later, you should use VSCROLL(n/65536). To scroll to
    > row 138, for example, enter VSCROLL(138, TRUE) (in any version) or
    > VSCROLL(138/16384) in earlier versions of Microsoft Excel or
    > VSCROLL(138/65536) in Microsoft Excel 97 or later

  - > If you are recording a macro and move the scroll box several times
    > in a row, the recorder only records the final location of the
    > scroll box, omitting any intermediate steps. Remember that
    > scrolling does not change the active cell or the selection.


**Related Functions**

[FORMULA.GOTO](FORMULA.GOTO.md)&nbsp;&nbsp;&nbsp;Selects a named area or reference on any
open workbook

[HLINE](HLINE.md)&nbsp;&nbsp;&nbsp;Horizontally scrolls through the active window by
columns

[HPAGE](HPAGE.md)&nbsp;&nbsp;&nbsp;Horizontally scrolls through the active window
one window at a time

[HSCROLL](HSCROLL.md)&nbsp;&nbsp;&nbsp;Horizontally scrolls through a sheet by
percentage or by column number

[SELECT](SELECT.md)&nbsp;&nbsp;&nbsp;Selects a cell, object, or chart item

[VLINE](VLINE.md)&nbsp;&nbsp;&nbsp;Vertically scrolls through the active window by
rows

[VPAGE](VPAGE.md)&nbsp;&nbsp;&nbsp;Vertically scrolls through the active window one
window at a time



Return to [README](README.md#V)

