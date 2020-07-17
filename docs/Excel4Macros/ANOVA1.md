ANOVA1

Performs single-factor analysis of variance, which tests the hypothesis
that means from several samples are equal.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**ANOVA1**(**inprng**, outrng, grouped, labels, alpha)

**ANOVA1**?(inprng, outrng, grouped, labels, alpha)

Inprng    is the input range.

Outrng    is the first cell (the upper-left cell) in the output table or
the name, as text, of a new sheet to contain the output table. If FALSE,
blank, or omitted, places the output table in a new workbook.

Grouped    is a text character that indicates whether the data in the
input range is organized by row or column.

  - > If grouped is "C" or omitted, then the data is organized by
    > column.

  - > If grouped is "R", then the data is organized by row.

Labels    is a logical value that describes where the labels are located
in the input range, as shown in the following table:

|                  |             |                                                   |
| ---------------- | ----------- | ------------------------------------------------- |
| **Labels**       | **Grouped** | **Labels are in**                                 |
| TRUE             | "C"         | First row of the input range.                     |
| TRUE             | "R"         | First column of the input range.                  |
| FALSE or omitted | (ignored)   | No labels. All cells in the input range are data. |

Alpha    is the significance level at which to evaluate critical values
for the F statistic. If omitted, alpha is 0.05.

**Related Functions**

[ANOVA](ANOVA.md)2   Performs two-factor analysis of variance with replication

[ANOVA](ANOVA.md)3   Performs two-factor analysis of variance without replication



Return to [README](README.md)

