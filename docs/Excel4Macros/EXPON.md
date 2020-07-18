EXPON

Predicts a value based on the forecast for the prior period, adjusted
for the error in that prior forecast.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**EXPON**(**inprng**, outrng, damp, stderrs, chart)

Inprng**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is the input range.

Outrng**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is the first cell (the upper-left cell) in
the output table or the name, as text, of a new sheet to contain the
output table. If FALSE, blank, or omitted, places the output table in a
new workbook.

Damp**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is the damping factor. If omitted, damp is
0.3.

Stderrs**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a logical value. If TRUE, standard
error values are included in the output table. If FALSE, standard errors
are not included.

Chart**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a logical value. If TRUE, EXPON
generates a chart for the actual and forecast values. If FALSE, the
chart is not generated.

**Related Function**

[MOVEAVG](MOVEAVG.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Returns values along a moving average trend



Return to [README](README.md)

