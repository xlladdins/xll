# NEW

Equivalent to clicking the New command on the File menu. Creates a new
Microsoft Excel workbook or opens a template.

**Syntax**

**NEW**(type\_num, xy\_series, add\_logical)

**NEW**?(type\_num, xy\_series, add\_logical)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies the type of workbook to
create, as shown in the following table. Type\_num is most often 5 or
quoted text; other values are mainly for compatibility with Microsoft
Excel version 4.0.

|               |                                                                               |
| ------------- | ----------------------------------------------------------------------------- |
| **Type\_num** | **Workbook**                                                                  |
| Omitted       | New workbook with a single worksheet of the same type as the active worksheet |
| 1             | New workbook with one worksheet                                               |
| 2             | New workbook with one chart based on the current selection                    |
| 3             | New workbook with one macro sheet                                             |
| 4             | New workbook with one international macro sheet                               |
| 5             | New workbook with 16 worksheets or based on the default workbook              |
| 6             | New workbook with one Visual Basic module                                     |
| 7             | New workbook with one dialog sheet                                            |
| Quoted text   | Template.                                                                     |

Xy\_series&nbsp;&nbsp;&nbsp;&nbsp;is a number from 0 to 3 that specifies
how data is arranged in a chart.

|                |                                                                                         |
| -------------- | --------------------------------------------------------------------------------------- |
| **Xy\_series** | **Result**                                                                              |
| 0              | Displays a dialog box if the selection is ambiguous.                                    |
| 1 or omitted   | The first row/column is the first data series.                                          |
| 2              | The first row/column contains the category (x) axis labels.                             |
| 3              | The first row/column contains the x-values; the created chart is an xy (scatter) chart. |

Add\_logical&nbsp;&nbsp;&nbsp;&nbsp;specifies whether or not to add the
sheet type to the open workbook. If add\_logical is TRUE, the sheet type
is inserted before the current sheet; if FALSE or omitted, it is not
inserted. This argument is for compatibility with Microsoft Excel
version 4.0.

Add\_logical is ignored if type\_num is 5.

**Remarks**

You can also use NEW to create new sheets from templates that exist in
the startup directory or folder, using for type\_num the text that
appears in the File New list box. To create new sheets from any template
that is not in the start-up directory, use the OPEN function.

**Related Functions**

[NEW.WINDOW](NEW.WINDOW.md)&nbsp;&nbsp;&nbsp;Creates a new window for an existing
worksheet or macro sheet

[OPEN](OPEN.md)&nbsp;&nbsp;&nbsp;Opens a workbook



Return to [README](README.md#N)

# NEW

Equivalent to clicking the New command on the File menu. Creates a new
Microsoft Excel workbook or opens a template.

**Syntax**

**NEW**(type\_num, xy\_series, add\_logical)

**NEW**?(type\_num, xy\_series, add\_logical)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies the type of workbook to
create, as shown in the following table. Type\_num is most often 5 or
quoted text; other values are mainly for compatibility with Microsoft
Excel version 4.0.

|               |                                                                               |
| ------------- | ----------------------------------------------------------------------------- |
| **Type\_num** | **Workbook**                                                                  |
| Omitted       | New workbook with a single worksheet of the same type as the active worksheet |
| 1             | New workbook with one worksheet                                               |
| 2             | New workbook with one chart based on the current selection                    |
| 3             | New workbook with one macro sheet                                             |
| 4             | New workbook with one international macro sheet                               |
| 5             | New workbook with 16 worksheets or based on the default workbook              |
| 6             | New workbook with one Visual Basic module                                     |
| 7             | New workbook with one dialog sheet                                            |
| Quoted text   | Template.                                                                     |

Xy\_series&nbsp;&nbsp;&nbsp;&nbsp;is a number from 0 to 3 that specifies
how data is arranged in a chart.

|                |                                                                                         |
| -------------- | --------------------------------------------------------------------------------------- |
| **Xy\_series** | **Result**                                                                              |
| 0              | Displays a dialog box if the selection is ambiguous.                                    |
| 1 or omitted   | The first row/column is the first data series.                                          |
| 2              | The first row/column contains the category (x) axis labels.                             |
| 3              | The first row/column contains the x-values; the created chart is an xy (scatter) chart. |

Add\_logical&nbsp;&nbsp;&nbsp;&nbsp;specifies whether or not to add the
sheet type to the open workbook. If add\_logical is TRUE, the sheet type
is inserted before the current sheet; if FALSE or omitted, it is not
inserted. This argument is for compatibility with Microsoft Excel
version 4.0.

Add\_logical is ignored if type\_num is 5.

**Remarks**

You can also use NEW to create new sheets from templates that exist in
the startup directory or folder, using for type\_num the text that
appears in the File New list box. To create new sheets from any template
that is not in the start-up directory, use the OPEN function.

**Related Functions**

[NEW.WINDOW](NEW.WINDOW.md)&nbsp;&nbsp;&nbsp;Creates a new window for an existing
worksheet or macro sheet

[OPEN](OPEN.md)&nbsp;&nbsp;&nbsp;Opens a workbook



Return to [README](README.md#N)

