WORKBOOK.NAME
=============

Equivalent to clicking the Rename command on the Sheet submenu of the
Format menu. Renames a sheet in a workbook.

**Syntax**

**WORKBOOK.NAME**(**oldname\_text**, **newname\_text**)

**WORKBOOK.NAME**?(oldname\_text, newname\_text)

Oldname\_text    is the name of the sheet that you want to rename.

Newname\_text    is the new name of the sheet.

**Remarks**

-   If you try to rename a sheet using a sheet name that already exists
    > in the workbook, Microsoft Excel displays an error message and
    > interrupts the macro.

-   If the structure of the workbook is protected, you cannot rename any
    > of the sheets in the workbook.

Return to [top](#T)

WORKBOOK.NEW
