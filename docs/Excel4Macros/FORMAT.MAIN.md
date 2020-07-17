FORMAT.MAIN

Equivalent to clicking the Main Chart command on the Format menu in
Microsoft Excel version 4.0. Formats a chart according to the arguments
you specify. This function is included for compatibility with Microsoft
Excel version 4.0. In Microsoft Excel version 5.0 or later, this is
equivalent to clicking the Chart Type command on the Format menu. You
can also use the FORMAT.CHART function.

**Syntax**

**FORMAT.MAIN**(**type\_num**, view, overlap, gap\_width, vary, drop,
hilo, angle, gap\_depth, chart\_depth, up\_down, series\_line, labels,
doughnut\_size)

**FORMAT.MAIN**?(type\_num, view, overlap, gap\_width, vary, drop, hilo,
angle, gap\_depth, chart\_depth, up\_down, series\_line, labels,
doughnut\_size)

Type\_num    is a number specifying the type of chart.

|               |              |
| ------------- | ------------ |
| **Type\_num** | **Chart**    |
| 1             | Area         |
| 2             | Bar          |
| 3             | Column       |
| 4             | Line         |
| 5             | Pie          |
| 6             | XY (Scatter) |
| 7             | 3-D Area     |
| 8             | 3-D Column   |
| 9             | 3-D Line     |
| 10            | 3-D Pie      |
| 11            | Radar        |
| 12            | 3-D Bar      |
| 13            | 3-D Surface  |
| 14            | Doughnut     |

View    is a number specifying one of the views in the Data View box in
the Main Chart dialog box. The view varies depending on the type of
chart.

Overlap    is a number from -100 to 100 specifying how you want bars or
columns to be positioned. It corresponds to the Overlap box in the Main
Chart dialog box. Overlap is ignored if type\_num is not 2 or 3 (bar or
column chart).

  - > If overlap is positive, it specifies the percentage of overlap you
    > want for bars or columns. For example, 50 would cause one-half of
    > a bar or column to be covered by an adjacent bar or column. A
    > value of zero prevents bars or columns from overlapping.

  - > If overlap is negative, then bars or columns are separated by the
    > specified percentage of the maximum available distance between any
    > two bars or columns.

  - > If overlap is omitted, it is assumed to be 0 (bars or columns do
    > not overlap), or it is unchanged if a value was previously set.

Gap\_width    is a number from 0 to 500 specifying the space between bar
or column clusters as a percentage of the width of a bar or column. It
corresponds to the Gap Width box in the Main Chart dialog box.

  - > Gap\_width is ignored if type\_num is not 2, 3, 8, or 12 (bar or
    > column chart).

  - > If gap\_width is omitted, it is assumed to be 50, or it is
    > unchanged if a value was previously set.

Several of the following arguments are logical values corresponding to
check boxes in the Main Chart dialog box. If an argument is TRUE,
Microsoft Excel selects the corresponding check box; if FALSE, Microsoft
Excel clears the check box. If an argument is omitted, the setting is
unchanged.

Vary    corresponds to the Vary By Categories check box. Vary applies
only to charts with one data series and is not available for area
charts.

Drop    corresponds to the Drop Lines check box. Drop is available only
for area and line charts.

Hilo    corresponds to the Hi-Lo Lines check box. Hilo is available only
for line charts.

Angle    is a number from 0 to 360 specifying the angle of the first pie
slice (in degrees) if the chart is a pie chart. If angle is omitted, it
is assumed to be 0, or it is unchanged if a value was previously set.

The next two arguments are for 3-D charts only.

Gap\_depth    is a number from 0 to 500 specifying the depth of the gap
in front of and behind a bar, column, area, or line as a percentage of
the depth of the bar, column, area, or line.

  - > Gap\_depth is ignored if the chart is a pie chart or if it is not
    > a 3-D chart.

  - > If gap\_depth is omitted and the chart is a 3-D chart, gap\_depth
    > is assumed to be 50, or it is unchanged if a value was previously
    > set. If gap\_depth is omitted and the view is side-by-side,
    > stacked, or stacked 100%, gap\_depth is assumed to be 0, or it is
    > unchanged if a value was previously set.

>  

Chart\_depth    is a number from 20 to 2000 specifying the visual depth
of the chart as a percentage of the width of the chart. Chart\_depth
corresponds to the Chart Depth box in the Main Chart dialog box.

  - > Chart\_depth is ignored if the chart is not a 3-D chart.

  - > If chart\_depth is omitted, it is assumed to be 100, or it is
    > unchanged if a value was previously set.

>  

The next three arguments are logical values corresponding to check boxes
in the Main Chart dialog box. If an argument is TRUE, Microsoft Excel
selects the corresponding check box; if FALSE, Microsoft Excel clears
the check box. If an argument is omitted, the setting is unchanged. The
final argument is for compatibility with Microsoft Excel version 4.0.

Up\_down    corresponds to the Up/Down Bars check box. Up\_down is
available only for line charts.

Series\_line    corresponds to the Series Lines check box. Series\_line
is available only for stacked bar and column charts.

Labels    corresponds to the Radar Axis Labels check box. Labels is
available only for radar charts.

Doughnut\_size    specifies the size of the hole in a doughnut chart.
Can be a value from 10% - 90%. Default is 50%

**Related Functions**

[FORMAT.CHART](FORMAT.CHART.md)   Formats a chart

[FORMAT.OVERLAY](FORMAT.OVERLAY.md)   Formats an overlay chart



Return to [README.md](README.md)

FORMAT.MAIN

Equivalent to clicking the Main Chart command on the Format menu in
Microsoft Excel version 4.0. Formats a chart according to the arguments
you specify. This function is included for compatibility with Microsoft
Excel version 4.0. In Microsoft Excel version 5.0 or later, this is
equivalent to clicking the Chart Type command on the Format menu. You
can also use the FORMAT.CHART function.

**Syntax**

**FORMAT.MAIN**(**type\_num**, view, overlap, gap\_width, vary, drop,
hilo, angle, gap\_depth, chart\_depth, up\_down, series\_line, labels,
doughnut\_size)

**FORMAT.MAIN**?(type\_num, view, overlap, gap\_width, vary, drop, hilo,
angle, gap\_depth, chart\_depth, up\_down, series\_line, labels,
doughnut\_size)

Type\_num    is a number specifying the type of chart.

|               |              |
| ------------- | ------------ |
| **Type\_num** | **Chart**    |
| 1             | Area         |
| 2             | Bar          |
| 3             | Column       |
| 4             | Line         |
| 5             | Pie          |
| 6             | XY (Scatter) |
| 7             | 3-D Area     |
| 8             | 3-D Column   |
| 9             | 3-D Line     |
| 10            | 3-D Pie      |
| 11            | Radar        |
| 12            | 3-D Bar      |
| 13            | 3-D Surface  |
| 14            | Doughnut     |

View    is a number specifying one of the views in the Data View box in
the Main Chart dialog box. The view varies depending on the type of
chart.

Overlap    is a number from -100 to 100 specifying how you want bars or
columns to be positioned. It corresponds to the Overlap box in the Main
Chart dialog box. Overlap is ignored if type\_num is not 2 or 3 (bar or
column chart).

  - > If overlap is positive, it specifies the percentage of overlap you
    > want for bars or columns. For example, 50 would cause one-half of
    > a bar or column to be covered by an adjacent bar or column. A
    > value of zero prevents bars or columns from overlapping.

  - > If overlap is negative, then bars or columns are separated by the
    > specified percentage of the maximum available distance between any
    > two bars or columns.

  - > If overlap is omitted, it is assumed to be 0 (bars or columns do
    > not overlap), or it is unchanged if a value was previously set.

Gap\_width    is a number from 0 to 500 specifying the space between bar
or column clusters as a percentage of the width of a bar or column. It
corresponds to the Gap Width box in the Main Chart dialog box.

  - > Gap\_width is ignored if type\_num is not 2, 3, 8, or 12 (bar or
    > column chart).

  - > If gap\_width is omitted, it is assumed to be 50, or it is
    > unchanged if a value was previously set.

Several of the following arguments are logical values corresponding to
check boxes in the Main Chart dialog box. If an argument is TRUE,
Microsoft Excel selects the corresponding check box; if FALSE, Microsoft
Excel clears the check box. If an argument is omitted, the setting is
unchanged.

Vary    corresponds to the Vary By Categories check box. Vary applies
only to charts with one data series and is not available for area
charts.

Drop    corresponds to the Drop Lines check box. Drop is available only
for area and line charts.

Hilo    corresponds to the Hi-Lo Lines check box. Hilo is available only
for line charts.

Angle    is a number from 0 to 360 specifying the angle of the first pie
slice (in degrees) if the chart is a pie chart. If angle is omitted, it
is assumed to be 0, or it is unchanged if a value was previously set.

The next two arguments are for 3-D charts only.

Gap\_depth    is a number from 0 to 500 specifying the depth of the gap
in front of and behind a bar, column, area, or line as a percentage of
the depth of the bar, column, area, or line.

  - > Gap\_depth is ignored if the chart is a pie chart or if it is not
    > a 3-D chart.

  - > If gap\_depth is omitted and the chart is a 3-D chart, gap\_depth
    > is assumed to be 50, or it is unchanged if a value was previously
    > set. If gap\_depth is omitted and the view is side-by-side,
    > stacked, or stacked 100%, gap\_depth is assumed to be 0, or it is
    > unchanged if a value was previously set.

>  

Chart\_depth    is a number from 20 to 2000 specifying the visual depth
of the chart as a percentage of the width of the chart. Chart\_depth
corresponds to the Chart Depth box in the Main Chart dialog box.

  - > Chart\_depth is ignored if the chart is not a 3-D chart.

  - > If chart\_depth is omitted, it is assumed to be 100, or it is
    > unchanged if a value was previously set.

>  

The next three arguments are logical values corresponding to check boxes
in the Main Chart dialog box. If an argument is TRUE, Microsoft Excel
selects the corresponding check box; if FALSE, Microsoft Excel clears
the check box. If an argument is omitted, the setting is unchanged. The
final argument is for compatibility with Microsoft Excel version 4.0.

Up\_down    corresponds to the Up/Down Bars check box. Up\_down is
available only for line charts.

Series\_line    corresponds to the Series Lines check box. Series\_line
is available only for stacked bar and column charts.

Labels    corresponds to the Radar Axis Labels check box. Labels is
available only for radar charts.

Doughnut\_size    specifies the size of the hole in a doughnut chart.
Can be a value from 10% - 90%. Default is 50%

**Related Functions**

[FORMAT.CHART](FORMAT.CHART.md)   Formats a chart

[FORMAT.OVERLAY](FORMAT.OVERLAY.md)   Formats an overlay chart



Return to [README.md](README.md)

