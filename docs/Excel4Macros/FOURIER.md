# FOURIER

Performs a Fourier transform.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**FOURIER**(**inprng**, outrng, inverse, labels)

**FOURIER**?(inprng, outrng, inverse, labels)

Inprng&nbsp;&nbsp;&nbsp;&nbsp;is the input range. The number of cells in
the input range must be equal to a power of two (2, 4, 8, 16, ...).

Outrng&nbsp;&nbsp;&nbsp;&nbsp;is the first cell in the output range or
the name, as text, of a new sheet to contain the output table. If FALSE,
blank, or omitted, places the output table in a new workbook.

Inverse&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If TRUE, an inverse
Fourier transform is performed. If FALSE or omitted, a forward Fourier
transform is performed.

Labels&nbsp;&nbsp;&nbsp;&nbsp;is a logical value.

  - > If labels is TRUE, then the first row or column of inprng contains
    > labels.

  - > If labels is FALSE or omitted, all cells in inprng are considered
    > data. Microsoft Excel generates appropriate data labels for the
    > output table.


**Related Function**

[SAMPLE](SAMPLE.md)&nbsp;&nbsp;&nbsp;Samples data



Return to [README](README.md)

