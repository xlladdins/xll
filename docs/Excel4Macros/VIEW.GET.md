# VIEW.GET

Equivalent to displaying a list of views in the Custom Views dialog box,
which appears when you click the Custom Views command on the View menu.
In Microsoft Excel 97 or later, the Custom Views command replaces the
View Manager command that was available in Microsoft Excel 95 and
earlier versions. Returns an array of views from the active workbook.

If this function is not available in Microsoft Excel 95 or in earlier
versions, you must install the View Manager add-in.

**Syntax**

**VIEW.GET**(**type\_num**, view\_name)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 3 that specifies
the type of information to return, as shown in the following table.

|               |                                                                                                                                                                                                                               |
| ------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Type\_num** | **Result**                                                                                                                                                                                                                    |
| 1             | Returns an array of views from all the sheets in the active workbook or the \#N/A error value if none are defined.                                                                                                            |
| 2             | Returns TRUE if print settings are included in the specified view. Returns FALSE if print settings are not included. Returns the \#VALUE\! error value if the name is invalid or the workbook is protected.                   |
| 3             | Returns TRUE if row and column settings are included in the specified view. Returns FALSE if row and column settings are not included. Returns the \#VALUE\! error value if the name is invalid or the workbook is protected. |

View\_name&nbsp;&nbsp;&nbsp;&nbsp;is text enclosed in quotation marks
and specifies the name of a view in the active workbook. View\_name is
required if type\_num is 2 or 3.

**Example**

The following macro formula returns an array of views from the active
workbook:

VIEW.GET(1)

**Related Functions**

[VIEW.DEFINE](VIEW.DEFINE.md)&nbsp;&nbsp;&nbsp;Creates or replaces a view

[VIEW.DELETE](VIEW.DELETE.md)&nbsp;&nbsp;&nbsp;Removes a view from the active workbook

[VIEW.SHOW](VIEW.SHOW.md)&nbsp;&nbsp;&nbsp;Shows a view



Return to [README](README.md#V)

