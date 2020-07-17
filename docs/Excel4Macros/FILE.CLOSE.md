FILE.CLOSE

Equivalent to clicking the Close command on the File menu. Closes the
active workbook.

**Syntax**

**FILE.CLOSE**(save\_logical, route\_logical)

Save\_logical    is a logical value specifying whether to save the file
before closing it.

|                   |                                                                                                       |
| ----------------- | ----------------------------------------------------------------------------------------------------- |
| **Save\_logical** | **Result**                                                                                            |
| TRUE              | Saves the workbook                                                                                    |
| FALSE             | Does not save the workbook                                                                            |
| Omitted           | If you've made changes to the workbook, displays a dialog box asking if you want to save the workbook |

Route\_logical    is a logical value that specifies whether to route the
file after closing it. This argument is ignored if there is not a
routing slip present.

|                    |                         |
| ------------------ | ----------------------- |
| **Route\_logical** | **Result**              |
| TRUE               | Routes the file         |
| FALSE              | Does not route the file |

Omitted   If you've specified recipients for routing, displays a dialog
box asking if you want to save the file

**Remarks**

If you make any changes to the structure of a workbook, such as the name
of sheets, their order, and so on, then a message will be displayed
reminding you that there are unsaved changes, regardless of the
save\_logical value.

**Note   **When you use the FILE.CLOSE function, Microsoft Excel does
not run any Auto\_Close macros before closing the workbook.

**Related Functions**

CLOSE   Closes the active window

CLOSE.ALL   Closes all unprotected windows

FCLOSE   Closes a text file


