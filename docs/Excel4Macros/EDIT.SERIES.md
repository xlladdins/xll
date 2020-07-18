EDIT.SERIES

Equivalent to clicking the Edit Series command on the Chart menu in
Microsoft Excel version 4.0. Creates or changes chart series by adding a
new SERIES formula or modifying an existing SERIES formula in the
topmost chart type. Chart types are displayed in the following order
from top to bottom: XY (Scatter), Line, Column, Bar, Area.

**Syntax**

**EDIT.SERIES**(**series\_num**, **name\_ref**, **x\_ref**, **y\_ref**,
**z\_ref**, **plot\_order**)

**EDIT.SERIES**?(series\_num, name\_ref, x\_ref, y\_ref, z\_ref,
plot\_order)

Series\_num**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is the number of the series you want
to change. If series\_num is 0 or omitted, Microsoft Excel creates a new
data series.

Name\_ref**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is the name of the data series. It can
be an external reference to a single cell, a name defined as a single
cell, or a name defined as a sequence of characters. Name\_ref can also
be text (for example, "Projected Sales").

X\_ref**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is an external reference to the name of
the sheet and the cells that contain one of the following sets of data:

  - > Category labels for all charts except xy (scatter) charts

  - > X-coordinate data for xy (scatter) charts

Y\_ref**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is an external reference to the name of
the sheet and the cells that contain values (or y-coordinate data in xy
(scatter) charts) for all 2-D charts. Y\_ref is required in 2-D charts
but does not apply to 3-D charts.

Z\_ref**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is an external reference to the name of
the sheet and the cells that contain values for all 3-D charts. Z\_ref
is required in 3-D charts but does not apply to 2-D charts.

Plot\_order**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a number specifying whether the
data series is plotted first, second, and so on, in the chart type.

  - > If you assign a plot\_order to a series, Microsoft Excel plots
    > that series in the order you specify, and the series that
    > previously had that plot order (and any series following it) has
    > its plot order increased by one.

  - > If you add a series to a chart with an overlay, the number of
    > series in the main chart does not change, so if the series is
    > added to the main chart, then the series that was plotted last in
    > the main chart will be plotted first in the overlay chart. To
    > change which series is plotted first in the overlay chart, use the
    > (chart type) Group command from the Format menu, and then select
    > the Series Order tab in the Format (chart type) Group dialog box.
    > You can also use the FORMAT.OVERLAY function.

  - > If you omit plot\_order when you add a new series, then Microsoft
    > Excel plots that series last and assigns it the correct
    > plot\_order value.

  - > The maximum value for plot\_order is 255.

**Remarks**

To change where a series is plotted within a chart, you can change the
chart type, using the FORMAT.CHART function, or the plot order. Plot
order affects where the series appears within the chart type only.

X\_ref, y\_ref, and z\_ref can be arrays or references to a nonadjacent
selection, although they cannot be names that refer to a nonadjacent
selection. If you specify a nonadjacent selection for any of these
arguments, make sure to enclose the reference to the selection in
parentheses so that Microsoft Excel does not treat the components of the
references as separate arguments.

**Tip****&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;To delete a data series, use the SELECT("Sn")
macro function, where n is the series number, followed by the
FORMULA("") macro function. You can also use the CLEAR function instead
of FORMULA.

**Related Function**

[FORMAT.CHART](FORMAT.CHART.md)



Return to [README](README.md)

