# ATTACH.TEXT

Attaches text to certain parts of the selected chart. Use ATTACH.TEXT to
attach text as a title or as a label for an axis or data point.

**Syntax**

**ATTACH.TEXT**(**attach\_to\_num**, series\_num, point\_num)

**ATTACH.TEXT**?(attach\_to\_num, series\_num, point\_num)

Attach\_to\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies which item on a chart
to attach text to. Attach\_to\_num is different for 2-D and 3-D charts.
Attach\_to\_num values for 2-D charts are shown in the following table.

|                     |                             |
| ------------------- | --------------------------- |
| **Attach\_to\_num** | **Attaches text to**        |
| 1                   | Chart title                 |
| 2                   | Value (y) axis              |
| 3                   | Category (x) axis           |
| 4                   | Series and data point       |
| 5                   | Secondary value (y) axis    |
| 6                   | Secondary category (x) axis |

Attach\_to\_num values for 3-D charts are shown in the following table.

|                     |                       |
| ------------------- | --------------------- |
| **Attach\_to\_num** | **Attaches text to**  |
| 1                   | Chart title           |
| 2                   | Value (z) axis        |
| 3                   | Series (y) axis       |
| 4                   | Category (x) axis     |
| 5                   | Series and data point |

Series\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies the series number if
attach\_to\_num specifies a series or data point. If attach\_to\_num
specifies a series or data point and series\_num is omitted, the macro
is interrupted.

Point\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies the number of the data
point, but only if you specify a series number. Point\_num is required
if series\_num is specified, unless the chart is an area chart.

**Remarks**

When you record adding an axis title or a chart title, Microsoft Excel
records both an ATTACH.TEXT function to attach the text and a
FONT.PROPERTIES function to make the text bold.

**Example**

The following macro functions attach the text "Quarterly Sales" to the x
(category) axis of the selected chart:

ATTACH.TEXT(3)

FORMULA("Quarterly Sales")

**Related Functions**

[DATA.LABEL](DATA.LABEL.md)&nbsp;&nbsp;&nbsp;Assigns text labels to point on a chart

[FORMULA](FORMULA.md)&nbsp;&nbsp;&nbsp;Enters values into a cell or range or onto a
chart



Return to [README](README.md#A)

# ATTACH.TEXT

Attaches text to certain parts of the selected chart. Use ATTACH.TEXT to
attach text as a title or as a label for an axis or data point.

**Syntax**

**ATTACH.TEXT**(**attach\_to\_num**, series\_num, point\_num)

**ATTACH.TEXT**?(attach\_to\_num, series\_num, point\_num)

Attach\_to\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies which item on a chart
to attach text to. Attach\_to\_num is different for 2-D and 3-D charts.
Attach\_to\_num values for 2-D charts are shown in the following table.

|                     |                             |
| ------------------- | --------------------------- |
| **Attach\_to\_num** | **Attaches text to**        |
| 1                   | Chart title                 |
| 2                   | Value (y) axis              |
| 3                   | Category (x) axis           |
| 4                   | Series and data point       |
| 5                   | Secondary value (y) axis    |
| 6                   | Secondary category (x) axis |

Attach\_to\_num values for 3-D charts are shown in the following table.

|                     |                       |
| ------------------- | --------------------- |
| **Attach\_to\_num** | **Attaches text to**  |
| 1                   | Chart title           |
| 2                   | Value (z) axis        |
| 3                   | Series (y) axis       |
| 4                   | Category (x) axis     |
| 5                   | Series and data point |

Series\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies the series number if
attach\_to\_num specifies a series or data point. If attach\_to\_num
specifies a series or data point and series\_num is omitted, the macro
is interrupted.

Point\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies the number of the data
point, but only if you specify a series number. Point\_num is required
if series\_num is specified, unless the chart is an area chart.

**Remarks**

When you record adding an axis title or a chart title, Microsoft Excel
records both an ATTACH.TEXT function to attach the text and a
FONT.PROPERTIES function to make the text bold.

**Example**

The following macro functions attach the text "Quarterly Sales" to the x
(category) axis of the selected chart:

ATTACH.TEXT(3)

FORMULA("Quarterly Sales")

**Related Functions**

[DATA.LABEL](DATA.LABEL.md)&nbsp;&nbsp;&nbsp;Assigns text labels to point on a chart

[FORMULA](FORMULA.md)&nbsp;&nbsp;&nbsp;Enters values into a cell or range or onto a
chart



Return to [README](README.md#A)

# ATTACH.TEXT

Attaches text to certain parts of the selected chart. Use ATTACH.TEXT to
attach text as a title or as a label for an axis or data point.

**Syntax**

**ATTACH.TEXT**(**attach\_to\_num**, series\_num, point\_num)

**ATTACH.TEXT**?(attach\_to\_num, series\_num, point\_num)

Attach\_to\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies which item on a chart
to attach text to. Attach\_to\_num is different for 2-D and 3-D charts.
Attach\_to\_num values for 2-D charts are shown in the following table.

|                     |                             |
| ------------------- | --------------------------- |
| **Attach\_to\_num** | **Attaches text to**        |
| 1                   | Chart title                 |
| 2                   | Value (y) axis              |
| 3                   | Category (x) axis           |
| 4                   | Series and data point       |
| 5                   | Secondary value (y) axis    |
| 6                   | Secondary category (x) axis |

Attach\_to\_num values for 3-D charts are shown in the following table.

|                     |                       |
| ------------------- | --------------------- |
| **Attach\_to\_num** | **Attaches text to**  |
| 1                   | Chart title           |
| 2                   | Value (z) axis        |
| 3                   | Series (y) axis       |
| 4                   | Category (x) axis     |
| 5                   | Series and data point |

Series\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies the series number if
attach\_to\_num specifies a series or data point. If attach\_to\_num
specifies a series or data point and series\_num is omitted, the macro
is interrupted.

Point\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies the number of the data
point, but only if you specify a series number. Point\_num is required
if series\_num is specified, unless the chart is an area chart.

**Remarks**

When you record adding an axis title or a chart title, Microsoft Excel
records both an ATTACH.TEXT function to attach the text and a
FONT.PROPERTIES function to make the text bold.

**Example**

The following macro functions attach the text "Quarterly Sales" to the x
(category) axis of the selected chart:

ATTACH.TEXT(3)

FORMULA("Quarterly Sales")

**Related Functions**

[DATA.LABEL](DATA.LABEL.md)&nbsp;&nbsp;&nbsp;Assigns text labels to point on a chart

[FORMULA](FORMULA.md)&nbsp;&nbsp;&nbsp;Enters values into a cell or range or onto a
chart



Return to [README](README.md#A)

