# FORMAT.OVERLAY

Equivalent to clicking the Overlay command on the Format menu in
Microsoft Excel version 4.0. Formats the overlay chart according to the
arguments you specify.

**Syntax**

**FORMAT.OVERLAY**(**type\_num**, view, overlap, gap\_width, vary, drop,
hilo, angle, series\_dist, series\_num, up\_down, series\_line, labels)

**FORMAT.OVERLAY**?(type\_num, view, overlap, gap\_width, vary, drop,
hilo, angle, series\_dist, series\_num, up\_down, series\_line, labels)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the type of
chart.

|               |              |
| ------------- | ------------ |
| **Type\_num** | **Chart**    |
| 1             | Area         |
| 2             | Bar          |
| 3             | Column       |
| 4             | Line         |
| 5             | Pie          |
| 6             | XY (Scatter) |
| 11            | Radar        |
| 14            | Doughnut     |

View&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying one of the views in
the Data View box in the Overlay dialog box. The view varies depending
on the type of chart.

Overlap&nbsp;&nbsp;&nbsp;&nbsp;is a number from -100 to 100 specifying
how you want bars or columns to be positioned. It corresponds to the
Overlap box in the Overlay dialog box. Overlap is ignored if type\_num
is not 2 or 3 (bar or column chart).

  - > If overlap is positive, it specifies the percentage of overlap you
    > want for bars or columns. For example, 50 would cause one-half of
    > a bar or column to be covered by an adjacent bar or column.

  - > If overlap is negative, then bars or columns are separated by the
    > specified percentage of the maximum available distance between any
    > two bars or columns.

  - > If overlap is omitted, it is assumed to be 0 (bars or columns do
    > not overlap), or it is unchanged if a value was previously set.


Gap\_width&nbsp;&nbsp;&nbsp;&nbsp;is a number from 0 to 500 specifying
the space between bar or column clusters as a percentage of the width of
a bar or column.

  - > Gap\_width is ignored if type\_num is not 2 or 3 (bar or column
    > chart).

  - > If gap\_width is omitted, it is assumed to be 50, or it is
    > unchanged if a value was previously set.


Several of the following arguments are logical values corresponding to
check boxes in the Overlay dialog box. If an argument is TRUE, Microsoft
Excel selects the corresponding check box; if FALSE, Microsoft Excel
clears the check box. If an argument is omitted, the setting is
unchanged.

Vary&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Vary By Categories check
box. Vary is not available for area charts.

Drop&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Drop Lines check box.
Drop is available only for area and line charts.

Hilo&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Hi-Lo Lines check box.
Hilo is available only for line charts.

Angle&nbsp;&nbsp;&nbsp;&nbsp;is a number from 0 to 360 specifying the
angle of the first pie slice (in degrees) if the chart is a pie chart.
If angle is omitted, it is assumed to be 0, or it is unchanged if a
value was previously set.

Series\_dist&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 and specifies
automatic or manual series distribution.

  - > If series\_dist is 1 or omitted, Microsoft Excel uses automatic
    > series distribution.

  - > If series\_dist is 2, Microsoft Excel uses manual series
    > distribution, and you must specify which series is first in the
    > distribution by using the series\_num argument.


Series\_num&nbsp;&nbsp;&nbsp;&nbsp;is the number of the first series in
the overlay chart and corresponds to the First Overlay Series box in the
Overlay dialog box. If series\_dist is 1 (automatic series
distribution), this argument is ignored.

Up\_down&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Up/Down Bars check
box. Up\_down is available only for line charts.

Series\_line&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Series Lines
check box. Series\_line is available only for stacked bar and column
charts.

Labels&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Radar Axis Labels check
box. Labels is available only for radar charts.

**Related Functions**

[DELETE.OVERLAY](DELETE.OVERLAY.md)&nbsp;&nbsp;&nbsp;Deletes the overlay on a chart

[FORMAT.CHART](FORMAT.CHART.md)&nbsp;&nbsp;&nbsp;Formats a chart



Return to [README](README.md#F)

