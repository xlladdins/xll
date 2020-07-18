MCOVAR

Returns a covariance matrix that measures the covariance between two or
more data sets.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**MCOVAR**(**inprng**, outrng, grouped, labels)

**MCOVAR**?(inprng, outrng, grouped, labels)

Inprng**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is the input range.

Outrng**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is the first cell (the upper-left cell) in
the output table or the name, as text, of a new sheet to contain the
output table. If FALSE, blank, or omitted, places the output table in a
new workbook.

Grouped**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a text character that indicates
whether the data in the input range is organized by row or column.

  - > If grouped is "C" or omitted, then the data is organized by
    > column.

  - > If grouped is "R", then the data is organized by row.


Labels**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a logical value that describes where
the labels are located in the input range, as shown in the following
table:

|                  |             |                                                   |
| ---------------- | ----------- | ------------------------------------------------- |
| **Labels**       | **Grouped** | **Labels are in**                                 |
| TRUE             | "C"         | First row of the input range                      |
| TRUE             | "R"         | First column of the input range                   |
| FALSE or omitted | (ignored)   | No labels. All cells in the input range are data. |

**Related Function**

[MCORREL](MCORREL.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Returns the correlation coefficient of two or
more data sets that are scaled to be independent of the unit of
measurement



Return to [README](README.md)

