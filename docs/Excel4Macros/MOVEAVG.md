# MOVEAVG

Projects values in a forecast period, based on the average value of the
variable over a specific number of preceding periods.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**MOVEAVG**(**inprng**, outrng, interval, stderrs, chart, labels)

**MOVEAVG**?(inprng, outrng, interval, stderrs, chart, labels)

Inprng&nbsp;&nbsp;&nbsp;&nbsp;is the input range.

Outrng&nbsp;&nbsp;&nbsp;&nbsp;is the first cell (the upper-left cell) in
the output table or the name, as text, of a new sheet to contain the
output table. If FALSE, blank, or omitted, places the output table in a
new workbook.

Interval&nbsp;&nbsp;&nbsp;&nbsp;is the number of values to include in
the moving average. If omitted, interval is 3.

Stderrs&nbsp;&nbsp;&nbsp;&nbsp;is a logical value.

  - > If stderrs is TRUE, standard error values are included in the
    > output table.

  - > If stderrs is FALSE or omitted, standard errors are not included
    > in the output table.


Chart&nbsp;&nbsp;&nbsp;&nbsp;is a logical value.

  - > If chart is TRUE, then MOVEAVG generates a chart for the actual
    > and forecast values.

  - > If chart is FALSE or omitted, the chart is not generated.


Labels&nbsp;&nbsp;&nbsp;&nbsp;is a logical value.

  - > If labels is TRUE, then the first row or column of inprng contains
    > labels.

  - > If labels is FALSE or omitted, all cells in inprng are considered
    > data. Microsoft Excel generates appropriate data labels for the
    > output table.



Return to [README](README.md#M)

