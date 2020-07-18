# QUIT

Equivalent to clicking the Exit command on the File menu in Microsoft
Excel for Windows. Equivalent to clicking the Quit command on the File
menu in Microsoft Excel for the Macintosh. Quits Microsoft Excel and
closes any open workbooks. If open workbooks have unsaved changes,
Microsoft Excel displays a message asking if you want to save them. You
can use QUIT in an Auto\_Close macro to force Microsoft Excel to quit
when a particular workbook is closed.

**Syntax**

**QUIT**( )

**Caution**&nbsp;&nbsp;&nbsp;**nbsp;&nbsp;&nbsp;&nbsp;**nbsp;&nbsp;&nbsp;&nbsp;**nbsp;If you have cleared error-checking with an
ERROR(FALSE) function, QUIT will not ask whether you want to save
changes.

**Remarks**

When you use the QUIT function, Microsoft Excel does not run any
Auto\_Close macros before closing the workbook.

**Examples**

The following function displays a confirmation alert and quits Microsoft
Excel if the user clicks OK:

IF(ALERT("Are you sure you want to quit Microsoft Excel?",1), QUIT(),)

**Related Function**

[FILE.CLOSE](FILE.CLOSE.md)&nbsp;&nbsp;&nbsp;Closes the active workbook



Return to [README](README.md)

