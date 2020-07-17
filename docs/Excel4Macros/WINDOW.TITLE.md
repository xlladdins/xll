WINDOW.TITLE

Changes the title of the active window to the title you specify. The
title appears at the top of the workbook window. Use WINDOW.TITLE to
control window titles when you're using Microsoft Excel to create a
custom application.

**Syntax**

**WINDOW.TITLE**(text)

Text    is the title you want to assign to the window. If text is
omitted, it is assumed to be the name of the workbook as it is stored on
your disk. Empty text ("") specifies no title.

**Important**   WINDOW.TITLE changes the name of the window, not the
actual name of the workbook as it is stored on your disk. To change the
name of the workbook, use the SAVE.AS function.

**Remarks**

  - > The window name you create using WINDOW.TITLE will appear on the
    > Window menu, and will be returned by the WINDOWS function, but not
    > by the DOCUMENTS function. You must use the new window name in
    > theACTIVATE function and the ON.WINDOW function.

  - > If you want to communicate with a Microsoft Excel workbook using
    > DDE functions like INITIATE or REQUEST, you must specify the
    > filename of the workbook and not the window title specified with
    > the WINDOW.TITLE function.

  - > If you use NEW.WINDOW to create new windows on the workbook, the
    > window title will be restored to its original name.

>  

**Example**

The following macro formula changes the title of the active window to
First Quarter.

WINDOW.TITLE("First Quarter")

**Related Functions**

[APP.TITLE](APP.TITLE.md)   Changes the title of the application workspace

[SAVE.AS](SAVE.AS.md)   Specifies a new filename, file type, protection password, or
write-reservation password, or to create a backup file


