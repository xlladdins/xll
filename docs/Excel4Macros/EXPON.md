EXPON

Predicts a value based on the forecast for the prior period, adjusted
for the error in that prior forecast.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**EXPON**(**inprng**, outrng, damp, stderrs, chart)

Inprng    is the input range.

Outrng    is the first cell (the upper-left cell) in the output table or
the name, as text, of a new sheet to contain the output table. If FALSE,
blank, or omitted, places the output table in a new workbook.

Damp    is the damping factor. If omitted, damp is 0.3.

Stderrs    is a logical value. If TRUE, standard error values are
included in the output table. If FALSE, standard errors are not
included.

Chart    is a logical value. If TRUE, EXPON generates a chart for the
actual and forecast values. If FALSE, the chart is not generated.

**Related Function**

[MOVEAVG](MOVEAVG.md)   Returns values along a moving average trend



Return to [README](README.md)

