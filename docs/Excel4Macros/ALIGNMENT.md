ALIGNMENT

Equivalent to clicking the Alignment tab in the Format Cells dialog box,
which is displayed when you click the Cells command on the Format menu.
Aligns the contents of the selected cells.

**Syntax**

**ALIGNMENT**(horiz\_align, wrap, vert\_align, orientation, add\_indent,
shrink\_to\_fit, merge\_cells)

**ALIGNMENT**?(horiz\_align, wrap, vert\_align, orientation,
add\_indent, shrink\_to\_fit, merge\_cells)

Horiz\_align    is a number from 1 to 7 specifying the type of
horizontal alignment, as shown in the following table. If horiz\_align
is omitted, horizontal alignment does not change.

|                  |                          |
| ---------------- | ------------------------ |
| **Horiz\_align** | **Horizontal alignment** |
| 1                | General                  |
| 2                | Left                     |
| 3                | Center                   |
| 4                | Right                    |
| 5                | Fill                     |
| 6                | Justify                  |
| 7                | Center across selection  |

Wrap    is a logical value corresponding to the Wrap Text check box in
the Alignment tab. If wrap is TRUE, Microsoft Excel selects the check
box and wraps text in cells; if FALSE, Microsoft Excel clears the check
box and does not wrap text. If wrap is omitted, wrapping does not
change.

Vert\_align    is a number from 1 to 4 specifying the vertical alignment
of the text. If vert\_align is omitted, vertical alignment does not
change.

|                 |                        |
| --------------- | ---------------------- |
| **Vert\_align** | **Vertical alignment** |
| 1               | Top                    |
| 2               | Center                 |
| 3               | Bottom                 |
| 4               | Justify                |

Orientation    is a number from 0 to 4 specifying the orientation of the
text. If orientation is omitted, text orientation does not change.

|                 |                                               |
| --------------- | --------------------------------------------- |
| **Orientation** | **Text orientation**                          |
| 0               | Horizontal                                    |
| 1               | Vertical                                      |
| 2               | Upward                                        |
| 3               | Downward                                      |
| 4               | Automatic (applies to only chart tick labels) |

Add\_indent     This argument is for only Far East versions of Microsoft
Excel.

Shrink\_to\_fit    is a logical value corresponding to the Shrink To Fit
check box in the Alignment tab.

Merge\_cells    is a logical value corresponding to the Merge Cells
check box in the Alignment tab. If merge\_cells is TRUE, Microsoft Excel
selects the check box and merges the selected cells; the merged cell
contains the value of the left-most cell that was merged. If FALSE,
Microsoft Excel clears the check box and unmerges the selected cells;
the left-most cell takes the formula or value of the cell that was
unmerged. If merge\_cells is omitted, cell mergers do not change.

**Related Function**

[FORMAT.TEXT](FORMAT.TEXT.md)   Formats a worksheet text box or a chart text item


