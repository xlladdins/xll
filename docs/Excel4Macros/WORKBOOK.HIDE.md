WORKBOOK.HIDE

Equivalent to clicking the Sheet command on the Format menu, and then
clicking Hide on the Sheet submenu. Hides sheets in the active workbook.

**Syntax**

**WORKBOOK.HIDE**(sheet\_text, very\_hidden)

Sheet\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of the sheet to hide. If
omitted, the currently selected sheet(s) are hidden.

Very\_hidden&nbsp;&nbsp;&nbsp;&nbsp;specifies how the sheet is hidden.
If TRUE, then the sheet name does not appear in the Unhide dialog box.
After using this argument, use WORKBOOK.UNHIDE to unhide the sheet. If
FALSE or omitted, hides the sheet but does not prevent the sheet's name
from appearing in the Unhide dialog box.

**Remarks**

  - > If the structure of the workbook is protected, you cannot hide any
    > sheets in the workbook.

  - > You cannot hide the last visible sheet in a workbook.

  - > To hide Sheet1:Sheet10, select them first with the WORKBOOK.SELECT
    > function. You can also place the sheets in an array first, as in
    > {"Sheet1", "Sheet2", "Sheet3",...}.



Return to [README](README.md)

