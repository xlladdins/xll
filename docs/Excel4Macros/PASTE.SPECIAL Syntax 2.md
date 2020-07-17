PASTE.SPECIAL Syntax 2
======================

Equivalent to clicking the Paste Special command on the Edit menu on the
Chart menu bar. Pastes the specified components from the copy area into
a chart. The PASTE.SPECIAL function has four syntax forms. Use syntax 2
if you have copied from a sheet and are pasting into a chart.

**Syntax**

**PASTE.SPECIAL**(rowcol, titles, categories, replace, series)

**PASTE.SPECIAL**?(rowcol, titles, categories, replace, series)

Rowcol    is the number 1 or 2 and specifies whether the values
corresponding to a particular data series are in rows or columns. Enter
1 for rows or 2 for columns.

Titles    is a logical value corresponding to the Series Names In First
Column check box (or First Row, depending on the value of rowcol) in the
Paste Special dialog box.

-   If series is TRUE, Microsoft Excel selects the check box and uses
    > the contents of the cell in the first column of each row (or first
    > row of each column) as the name of the data series in that row (or
    > column).

-   If series is FALSE, Microsoft Excel clears the check box and uses
    > the contents of the cell in the first column of each row (or first
    > row of each column) as the first data point of the data series.

>  

Categories    is a logical value corresponding to the Categories (X
Labels) In First Row (or First Column, depending on the value of rowcol)
check box in the Paste Special dialog box.

-   If categories is TRUE, Microsoft Excel selects the check box and
    > uses the contents of the first row (or column) of the selection as
    > the categories for the chart.

-   If categories is FALSE, Microsoft Excel clears the check box and
    > uses the contents of the first row (or column) as the first data
    > series in the chart.

Replace    is a logical value corresponding to the Replace Existing
Categories check box in the Paste Special dialog box.

-   If replace is TRUE, Microsoft Excel selects the check box and
    > applies categories while replacing existing categories with
    > information from the copied cell range.

-   If replace is FALSE, Microsoft Excel clears the check box and
    > applies new categories without replacing any old ones.

Series    is a number specifying how cells are added to a chart.

  ------------ --------------
  **Series**   **Added as**
  1            New series
  2            New point(s)
  ------------ --------------

**Related Functions**

Syntax 1   Pasting into a sheet or macro sheet

Syntax 3   Copying and pasting between charts

Syntax 4   Pasting information from another application

Return to [top](#H)

PASTE.SPECIAL Syntax 3
