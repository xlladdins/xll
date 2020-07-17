SERIES
======

Charts Only

Represents a data series in the active chart. SERIES is used only in
charts; you cannot enter it on a sheet or macro sheet. You normally
create or change data series by using the Chart Wizard or EDIT.SERIES
macro function, which is equivalent to the Edit Series command on the
Chart menu in Microsoft Excel version 4.0. However, you can edit a data
series manually by selecting it, switching to the formula bar, and
typing the changes.

**Syntax**

**SERIES**(name\_ref, categories, **values, plot\_order**)

Name\_ref    is the name of the data series. It can be an external
reference to a single cell or a name defined as a single cell. Name\_ref
can also be text enclosed in quotation marks (for example, \"Projected
Sales\").

Categories    is an external reference to the name of the workbook and
to the cells that contain one of the following sets of data:

-   Category labels for all charts except xy (scatter) charts

-   X-coordinate data for xy (scatter) charts

>  

Values    is an external reference to the name of the workbook and to
the cells that contain values (or y-coordinate data in scatter charts).

Plot\_order    is an integer specifying whether the series is plotted
first, second, or third, and so on, in the chart. No two series can have
the same plot\_order.

**Remarks**

-   Categories and values can be arrays or references to a multiple
    > selection, although they cannot be names that refer to a multiple
    > selection. If you specify a multiple selection for any of these
    > arguments, make sure you include the necessary sets of parentheses
    > so that Microsoft Excel does not treat the components of the
    > references as separate arguments.

-   If either categories or values is a multiple selection, then all
    > areas in that selection must be either vertical (more rows than
    > columns) or horizontal (more columns than rows).

>  

**Related Functions**

CHART.WIZARD   Creates and formats a chart

EDIT.SERIES   Creates or changes a chart series

Return to [top](#Q)

SERIES.AXES
