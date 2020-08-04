# FORMAT.CHART

Equivalent to choosing the Options button in the Chart Type dialog box,
which is available when you choose the Chart Type command from the
Format menu when a chart is active. Formats the chart according to the
arguments you specify.

**Syntax**

**FORMAT.CHART**(layer\_num, view, overlap, angle, gap\_width,
gap\_depth, chart\_depth, doughnut\_size, axis\_num, drop, hilo,
up\_down, series\_line, labels, vary)

**FORMAT.CHART**?(layer\_num, view, overlap, angle, gap\_width,
gap\_depth, chart\_depth, doughnut\_size, axis\_num, drop, hilo,
up\_down, series\_line, labels, vary)

Several of the following arguments are logical values corresponding to
check boxes in the Options tab of Format (chart type) Group dialog box.
If an argument is TRUE, Microsoft Excel selects the corresponding check
box; if FALSE, Microsoft Excel clears the check box. If an argument is
omitted, the setting is unchanged.

Layer\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying which chart you
want to change.

View&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying one of the subtypes
in the Subtype tab of the Format (type) Group dialog box. The subtype
varies depending on the type of chart.

Overlap&nbsp;&nbsp;&nbsp;&nbsp;is a number from -100 to 100 specifying
how you want bars or columns to be positioned. It corresponds to the
Overlap edit box in the Options tab on the Format Bar Group Dialog box,
which appears when you choose the Bar Group from the Format menu.
Overlap is ignored if type\_num is not 2 or 3 (bar or column chart).

  - > If overlap is positive, it specifies the percentage of overlap you
    > want for bars or columns. For example, 50 would cause one-half of
    > a bar or column to be covered by an adjacent bar or column. A
    > value of zero prevents bars or columns from overlapping.

  - > If overlap is negative, then bars or columns are separated by the
    > specified percentage of the maximum available distance between any
    > two bars or columns.

  - > If overlap is omitted, it is assumed to be 0 (bars or columns do
    > not overlap), or it is unchanged if a value was previously set.

Angle&nbsp;&nbsp;&nbsp;&nbsp;is a number from 0 to 360 specifying the
angle of the first pie or doughnut slice (in degrees) if the chart is a
pie or doughnut chart. If angle is omitted, it is assumed to be 0, or it
is unchanged if a value was previously set.

Gap\_width&nbsp;&nbsp;&nbsp;&nbsp;is a number from 0 to 500 specifying
the space between bar or column clusters as a percentage of the width of
a bar or column. It corresponds to the Gap Width edit box in the Options
tab on the Format Bar Group Dialog box, which appears when you choose
the Bar Group from the Format menu.

  - > Gap\_width is ignored if type\_num is not 2, 3, 8, or 12 (bar or
    > column chart).

  - > If Gap\_width is omitted, it is assumed to be 50, or it is
    > unchanged if a value was previously set.


The next two arguments are for 3-D charts only, and correspond to check
boxes in the Options tab of Format (chart type) Group dialog box.

Gap\_depth&nbsp;&nbsp;&nbsp;&nbsp;is a number from 0 to 500 specifying
the depth of the gap in front of and behind a bar, column, area, or line
as a percentage of the depth of the bar, column, area, or line.

  - > Gap\_depth is ignored if the chart is a pie chart or if it is not
    > a 3-D chart.

  - > If gap\_depth is omitted and the chart is a 3-D chart, gap\_depth
    > is assumed to be 50, or it is unchanged if a value was previously
    > set. If gap\_depth is omitted and the view is side-by-side,
    > stacked, or stacked 100%, gap\_depth is assumed to be 0, or it is
    > unchanged if a value was previously set.


Chart\_depth&nbsp;&nbsp;&nbsp;&nbsp;is a number from 20 to 2000
specifying the visual depth of the chart as a percentage of the width of
the chart.

  - > Chart\_depth is ignored if the chart is not a 3-D chart.

  - > If Chart\_depth is omitted, it is assumed to be 100, or it is
    > unchanged if a value was previously set.


Doughnut\_size&nbsp;&nbsp;&nbsp;&nbsp;specifies the size of the hole in
a doughnut chart. Can be a value from 10% - 90%. Default is 50%.

Axis\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying whether to plot
the chart on the primary axis or the secondary axis.

Drop&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Drop Lines check box.
Drop is available only for area and line charts.

Hilo&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Hi-Lo Lines check box.
Hilo is available only for 2-D line charts.

The next four arguments are logical values corresponding to check boxes
in the Options tab of the Format (chart type) Group dialog box. If an
argument is TRUE, Microsoft Excel selects the corresponding check box;
if FALSE, Microsoft Excel clears the check box. If an argument is
omitted, the setting is unchanged.

Up\_down&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Up/Down Bars check
box. Up\_down is available only for 2-D line charts.

Series\_line&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Series Lines
check box. Series\_line is available only for 2-D stacked bar and column
charts.

Labels&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Radar Axis Labels check
box. Labels is available only for radar charts.

Vary&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Vary Colors By Point
check box. Vary applies only to charts with one data series and is not
available for area charts.

**Related Functions**

[FORMAT.MAIN](FORMAT.MAIN.md)&nbsp;&nbsp;&nbsp;Formats a chart according to the arguments
you specify

[FORMAT.OVERLAY](FORMAT.OVERLAY.md)&nbsp;&nbsp;&nbsp;Formats an overlay chart



Return to [README](README.md#F)

