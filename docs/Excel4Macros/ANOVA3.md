ANOVA3
======

Performs two-factor analysis of variance without replication.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**ANOVA3**(**inprng**, outrng, labels, alpha)

**ANOVA3**?(inprng, outrng, labels, alpha)

Inprng    is the input range.

Outrng    is the first cell (the upper-left cell) in the output table or
the name, as text, of a new sheet to contain the output table. If FALSE,
blank, or omitted, places the output table in a new workbook.

Labels    is a logical value.

-   If labels is TRUE, then the first row and column of the input range
    > contain labels.

-   If labels is FALSE or omitted, all cells in inprng are considered
    > data. Microsoft Excel will then generate the appropriate data
    > labels for the output table.

>  

Alpha    is the significance level at which to evaluate critical values
for the F statistic. If omitted, alpha is 0.05.

**Related Functions**

ANOVA1   Performs single-factor analysis of variance

ANOVA2   Performs two-factor analysis of variance with replication

Return to [top](#A)

APP.ACTIVATE
