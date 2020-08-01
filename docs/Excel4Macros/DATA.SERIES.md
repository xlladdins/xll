# DATA.SERIES

Equivalent to clicking the Series command on the Fill submenu of the
Edit menu. Use DATA.SERIES to enter an interpolated or incrementally
increasing or decreasing series of numbers or dates on a sheet or macro
sheet.

**Syntax**

**DATA.SERIES**(rowcol, type\_num, date\_num, step\_value, stop\_value,
trend)

**DATA.SERIES**?(rowcol, type\_num, date\_num, step\_value, stop\_value,
trend)

Rowcol&nbsp;&nbsp;&nbsp;&nbsp;is a number that specifies where the
series should be entered. If rowcol is omitted, the default value is
based on the size and shape of the current selection.

|            |                     |
| ---------- | ------------------- |
| **Rowcol** | **Enter series in** |
| 1          | Rows                |
| 2          | Columns             |

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 4 that specifies
the type of series.

|               |                    |
| ------------- | ------------------ |
| **Type\_num** | **Type of series** |
| 1 or omitted  | Linear             |
| 2             | Growth             |
| 3             | Date               |
| 4             | AutoFill           |

Date\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 4 that specifies
the date unit of the series, as shown in the following table. To use the
date\_num argument, the type\_num argument must be 3.

|               |               |
| ------------- | ------------- |
| **Date\_num** | **Date unit** |
| 1 or omitted  | Day           |
| 2             | Weekday       |
| 3             | Month         |
| 4             | Year          |

Step\_value&nbsp;&nbsp;&nbsp;&nbsp;is a number that specifies the step
value for the series. If step\_value is omitted, it is assumed to be 1.

Stop\_value&nbsp;&nbsp;&nbsp;&nbsp;is a number that specifies the stop
value for the series. If stop\_value is omitted, DATA.SERIES continues
filling the series until the end of the selected range.

Trend&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to the
Trend check box. If trend is TRUE, Microsoft Excel generates a linear or
exponential trend; if FALSE or omitted, Microsoft Excel generates a
standard data series.

**Remarks**

  - > If you specify a positive value for stop\_value that is lower than
    > the value in the active cell of the selection, DATA.SERIES takes
    > no action.

  - > If type\_num is 4 (AutoFill), Microsoft Excel performs an AutoFill
    > operation just as if you had filled the selection by dragging the
    > fill selection handle or had used the FILL.AUTO macro function.


**Related Function**

[FILL.AUTO](FILL.AUTO.md)&nbsp;&nbsp;&nbsp;Copies cells or automatically fills a
selection



Return to [README](README.md#D)

# DATA.SERIES

Equivalent to clicking the Series command on the Fill submenu of the
Edit menu. Use DATA.SERIES to enter an interpolated or incrementally
increasing or decreasing series of numbers or dates on a sheet or macro
sheet.

**Syntax**

**DATA.SERIES**(rowcol, type\_num, date\_num, step\_value, stop\_value,
trend)

**DATA.SERIES**?(rowcol, type\_num, date\_num, step\_value, stop\_value,
trend)

Rowcol&nbsp;&nbsp;&nbsp;&nbsp;is a number that specifies where the
series should be entered. If rowcol is omitted, the default value is
based on the size and shape of the current selection.

|            |                     |
| ---------- | ------------------- |
| **Rowcol** | **Enter series in** |
| 1          | Rows                |
| 2          | Columns             |

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 4 that specifies
the type of series.

|               |                    |
| ------------- | ------------------ |
| **Type\_num** | **Type of series** |
| 1 or omitted  | Linear             |
| 2             | Growth             |
| 3             | Date               |
| 4             | AutoFill           |

Date\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 4 that specifies
the date unit of the series, as shown in the following table. To use the
date\_num argument, the type\_num argument must be 3.

|               |               |
| ------------- | ------------- |
| **Date\_num** | **Date unit** |
| 1 or omitted  | Day           |
| 2             | Weekday       |
| 3             | Month         |
| 4             | Year          |

Step\_value&nbsp;&nbsp;&nbsp;&nbsp;is a number that specifies the step
value for the series. If step\_value is omitted, it is assumed to be 1.

Stop\_value&nbsp;&nbsp;&nbsp;&nbsp;is a number that specifies the stop
value for the series. If stop\_value is omitted, DATA.SERIES continues
filling the series until the end of the selected range.

Trend&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to the
Trend check box. If trend is TRUE, Microsoft Excel generates a linear or
exponential trend; if FALSE or omitted, Microsoft Excel generates a
standard data series.

**Remarks**

  - > If you specify a positive value for stop\_value that is lower than
    > the value in the active cell of the selection, DATA.SERIES takes
    > no action.

  - > If type\_num is 4 (AutoFill), Microsoft Excel performs an AutoFill
    > operation just as if you had filled the selection by dragging the
    > fill selection handle or had used the FILL.AUTO macro function.


**Related Function**

[FILL.AUTO](FILL.AUTO.md)&nbsp;&nbsp;&nbsp;Copies cells or automatically fills a
selection



Return to [README](README.md#D)

# DATA.SERIES

Equivalent to clicking the Series command on the Fill submenu of the
Edit menu. Use DATA.SERIES to enter an interpolated or incrementally
increasing or decreasing series of numbers or dates on a sheet or macro
sheet.

**Syntax**

**DATA.SERIES**(rowcol, type\_num, date\_num, step\_value, stop\_value,
trend)

**DATA.SERIES**?(rowcol, type\_num, date\_num, step\_value, stop\_value,
trend)

Rowcol&nbsp;&nbsp;&nbsp;&nbsp;is a number that specifies where the
series should be entered. If rowcol is omitted, the default value is
based on the size and shape of the current selection.

|            |                     |
| ---------- | ------------------- |
| **Rowcol** | **Enter series in** |
| 1          | Rows                |
| 2          | Columns             |

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 4 that specifies
the type of series.

|               |                    |
| ------------- | ------------------ |
| **Type\_num** | **Type of series** |
| 1 or omitted  | Linear             |
| 2             | Growth             |
| 3             | Date               |
| 4             | AutoFill           |

Date\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 4 that specifies
the date unit of the series, as shown in the following table. To use the
date\_num argument, the type\_num argument must be 3.

|               |               |
| ------------- | ------------- |
| **Date\_num** | **Date unit** |
| 1 or omitted  | Day           |
| 2             | Weekday       |
| 3             | Month         |
| 4             | Year          |

Step\_value&nbsp;&nbsp;&nbsp;&nbsp;is a number that specifies the step
value for the series. If step\_value is omitted, it is assumed to be 1.

Stop\_value&nbsp;&nbsp;&nbsp;&nbsp;is a number that specifies the stop
value for the series. If stop\_value is omitted, DATA.SERIES continues
filling the series until the end of the selected range.

Trend&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to the
Trend check box. If trend is TRUE, Microsoft Excel generates a linear or
exponential trend; if FALSE or omitted, Microsoft Excel generates a
standard data series.

**Remarks**

  - > If you specify a positive value for stop\_value that is lower than
    > the value in the active cell of the selection, DATA.SERIES takes
    > no action.

  - > If type\_num is 4 (AutoFill), Microsoft Excel performs an AutoFill
    > operation just as if you had filled the selection by dragging the
    > fill selection handle or had used the FILL.AUTO macro function.


**Related Function**

[FILL.AUTO](FILL.AUTO.md)&nbsp;&nbsp;&nbsp;Copies cells or automatically fills a
selection



Return to [README](README.md#D)

