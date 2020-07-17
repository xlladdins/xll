ANOVA2

Performs two-factor analysis of variance with replication.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**ANOVA2**(**inprng**, outrng, **sample\_rows**, alpha)

**ANOVA2**?(inprng, outrng, sample\_rows, alpha)

Inprng    is the input range. The input range should contain labels in
the first row and column.

Outrng    is the first cell (the upper-left cell) in the output table or
the name, as text, of a new sheet to contain the output table. If FALSE,
blank, or omitted, places the output table in a new workbook.

Sample\_rows    is the number of rows in each sample.

Alpha    is the significance level at which to evaluate critical values
for the F statistic. If omitted, alpha is 0.05.

**Related Functions**

[ANOVA](ANOVA.md)1   Performs single-factor analysis of variance

[ANOVA](ANOVA.md)3   Performs two-factor analysis of variance without replication



Return to [README](README.md)

