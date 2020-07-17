CHART.ADD.DATA

Equivalent to dragging data from a worksheet onto a chart. Adds data to
an existing chart.

**Syntax**

**CHART.ADD.DATA**(**ref**, rowcol, titles, categories, replace, series)

Ref    is the cell reference for the data that is being dragged onto the
chart

Rowcol    is the number 1 or 2 and specifies whether the values
corresponding to a particular data series are in rows or columns. Enter
1 for rows or 2 for columns.

Titles    is a logical value corresponding to the Series Names In First
Column check box (or First Row, depending on the value of rowcol) in the
Paste Special dialog box.

  - > If titles is TRUE, Microsoft Excel selects the check box and uses
    > the contents of the cell in the first column of each row (or first
    > row of each column) as the name of the data series in that row (or
    > column).

  - > If titles is FALSE, Microsoft Excel clears the check box and uses
    > the contents of the cell in the first column of each row (or first
    > row of each column) as the first data point of the data series.

Categories    is a logical value corresponding to the Categories (X
Labels) In First Row (or First Column, depending on the value of rowcol)
check box in the Paste Special dialog box.

  - > If categories is TRUE, Microsoft Excel selects the check box and
    > uses the contents of the first row (or column) of the selection as
    > the categories for the chart.

  - > If categories is FALSE, Microsoft Excel clears the check box and
    > uses the contents of the first row (or column) as the first data
    > series in the chart.

>  

Replace    is a logical value corresponding to the Replace Existing
Categories check box in the Paste Special dialog box.

  - > If replace is TRUE, Microsoft Excel selects the check box and
    > applies categories while replacing existing categories with
    > information from the copied cell range.

  - > If replace is FALSE, Microsoft Excel clears the check box and
    > applies new categories without replacing any old ones.

Series    is a number specifying how cells are added to a chart.

|            |              |
| ---------- | ------------ |
| **Series** | **Added as** |
| 1          | New series   |
| 2          | New point(s) |



Return to [README](README.md)

