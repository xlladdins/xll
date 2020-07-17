SAMPLE
======

Samples data.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**SAMPLE**(**inprng**, outrng, **method, rate**, labels)

**SAMPLE**?(inprng, outrng, method, rate, labels)

Inprng    is the input range.

Outrng    is the first cell (the upper-left cell) in the output column
or the name, as text, of a new sheet to contain the output column. If
FALSE, blank, or omitted, places the output table in a new workbook.

Method    is a text character that indicates the type of sampling.

-   If method is \"P\", then periodic sampling is used. The input range
    > is sampled every nth cell, where n = rate.

-   If method is \"R\", then random sampling is used. The output column
    > will contain rate samples.

>  

Rate    is the sampling rate, if method is \"P\" (periodic sampling).
Rate is the number of samples to take if method is \"R\" (random
sampling).

Labels    is a logical value.

-   If labels is TRUE, then the first row or column of inprng contains
    > labels.

-   If labels is FALSE or omitted, all cells in inprng are considered
    > data. Microsoft Excel generates appropriate data labels for the
    > output table.

Return to [top](#Q)

SAVE
