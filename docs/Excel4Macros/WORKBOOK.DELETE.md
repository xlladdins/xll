WORKBOOK.DELETE
===============

Equivalent to clicking the Delete Sheet command on the Edit menu.
Deletes a sheet or group of sheets from the current workbook.

**Syntax**

**WORKBOOK.DELETE**(sheet\_text)

Sheet\_text    is the name of the sheet to delete. If omitted, the
currently active sheet or sheets is deleted.

**Remarks**

-   This function prompts for confirmation. To suppress the prompt, use
    > the ERROR function. For example, =ERROR(FALSE).

-   If the structure of the workbook is protected, you cannot delete any
    > of its sheets.

-   If you want to delete Sheet1:Sheet10, you must select them first
    > with WORKBOOK.SELECT(). You can also place the sheets in an array
    > first, as in {\"Sheet1\", \"Sheet2\", \"Sheet3\",\...}.

-   You cannot delete the last visible sheet in a workbook.

Return to [top](#T)

WORKBOOK.HIDE
