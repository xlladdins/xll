PTTESTM

Performs a paired two-sample Student's t-Test for means.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**PTTESTM**(**inprng1, inprng2**, outrng, labels, alpha, difference)

**PTTESTM**?(inprng1, inprng2, outrng, labels, alpha, difference)

Inprng1    is the input range for the first data set.

Inprng2    is the input range for the second data set.

Outrng    is the first cell (the upper-left cell) in the output table or
the name, as text, of a new sheet to contain the output table. If FALSE,
blank, or omitted, places the output table in a new workbook.

Labels    is a logical value.

  - > If labels is TRUE, then labels are in the first row or column of
    > the input ranges.

  - > If labels is FALSE or omitted, all cells in inprng1 and inprng2
    > are considered data. The output table will include default row or
    > column headings.

>  

Alpha    is the confidence level for the test. If omitted, alpha is
0.05.

Difference    is the hypothesized mean difference. If omitted,
difference is 0.

**Related Functions**

PTTESTV   Performs a two-sample Student's t-Test, assuming unequal
variances

TTESTM   Performs a two-sample Student's t-Test for means, assuming
equal variances


