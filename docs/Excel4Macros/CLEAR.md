# CLEAR

Equivalent to clicking the Clear command on the Edit menu. Clears
contents, formats, notes, or all of these from the active worksheet or
macro sheet. Clears series or formats from the active chart.

**Syntax**

**CLEAR**(type\_num)

**CLEAR**?(type\_num)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 4 specifying what
to clear. Only values 1, 2, and 3 are valid if the selected item is a
chart.

On a worksheet or macro sheet, or if an entire chart is selected, the
following occurs.

|               |                                                                  |
| ------------- | ---------------------------------------------------------------- |
| **Type\_num** | **Clears**                                                       |
| 1             | All                                                              |
| 2             | Formats (if a chart, clears the chart format or clears pictures) |
| 3             | Contents (if a chart, clears all data series)                    |
| 4             | Comments (this does not apply to charts)                         |

On a chart, if a single point, an entire data series, error bars, or a
trend line is selected, the following occurs.

|               |                                                                 |
| ------------- | --------------------------------------------------------------- |
| **Type\_num** | **Clears**                                                      |
| 1             | Selected series, error bars, or trend line                      |
| 2             | Format in the selected point, series, error bars, or trend line |

If type\_num is omitted, the default values are set as shown in the
following table.

|                            |                           |
| -------------------------- | ------------------------- |
| **Active sheet**           | **Type\_num**             |
| Worksheet                  | 3                         |
| Macro sheet                | 3                         |
| Chart (with no selection)  | 1                         |
| Chart (with item selected) | Deletes the selected item |

**Related Function**

[EDIT.DELETE](EDIT.DELETE.md)&nbsp;&nbsp;&nbsp;Removes cells from a sheet



Return to [README](README.md#C)

# CLEAR

Equivalent to clicking the Clear command on the Edit menu. Clears
contents, formats, notes, or all of these from the active worksheet or
macro sheet. Clears series or formats from the active chart.

**Syntax**

**CLEAR**(type\_num)

**CLEAR**?(type\_num)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 4 specifying what
to clear. Only values 1, 2, and 3 are valid if the selected item is a
chart.

On a worksheet or macro sheet, or if an entire chart is selected, the
following occurs.

|               |                                                                  |
| ------------- | ---------------------------------------------------------------- |
| **Type\_num** | **Clears**                                                       |
| 1             | All                                                              |
| 2             | Formats (if a chart, clears the chart format or clears pictures) |
| 3             | Contents (if a chart, clears all data series)                    |
| 4             | Comments (this does not apply to charts)                         |

On a chart, if a single point, an entire data series, error bars, or a
trend line is selected, the following occurs.

|               |                                                                 |
| ------------- | --------------------------------------------------------------- |
| **Type\_num** | **Clears**                                                      |
| 1             | Selected series, error bars, or trend line                      |
| 2             | Format in the selected point, series, error bars, or trend line |

If type\_num is omitted, the default values are set as shown in the
following table.

|                            |                           |
| -------------------------- | ------------------------- |
| **Active sheet**           | **Type\_num**             |
| Worksheet                  | 3                         |
| Macro sheet                | 3                         |
| Chart (with no selection)  | 1                         |
| Chart (with item selected) | Deletes the selected item |

**Related Function**

[EDIT.DELETE](EDIT.DELETE.md)&nbsp;&nbsp;&nbsp;Removes cells from a sheet



Return to [README](README.md#C)

