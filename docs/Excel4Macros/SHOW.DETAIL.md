SHOW.DETAIL
===========

Expands or collapses the detail under the specified expand or collapse
button.

**Syntax**

**SHOW.DETAIL**(**rowcol, rowcol\_num**, expand, show\_field)

Rowcol    is a number that specifies whether to operate on rows or
columns of data.

+--------------+------------------------------------------------------+
| > **Rowcol** | > **Operates on**                                    |
+--------------+------------------------------------------------------+
| > 1          | > Rows                                               |
+--------------+------------------------------------------------------+
| > 2          | > Columns                                            |
+--------------+------------------------------------------------------+
| > 3          | > The current cell\'s row or column. The second      |
|              | > argument, rowcol\_num, is then ignored.            |
+--------------+------------------------------------------------------+

Rowcol\_num    is a number that specifies the row or column to expand or
collapse. If you are in A1 mode, you must still give the column as a
number. If rowcol\_num is not a summary row or column, SHOW.DETAIL
returns the \#VALUE! error value and interrupts the macro.

Expand    is a logical value that specifies whether to expand or
collapse the detail under the row or column. If expand is TRUE,
Microsoft Excel expands the detail under the row or column; if FALSE, it
collapses the detail under the row or column. If expand is omitted, the
detail is expanded if it is currently collapsed and collapsed if it is
currently expanded.

Show\_Field    is a string specifying the name of the field to add to a
PivotTable report, if the selection is inside a PivotTable report. The
new field is added as the new innermost field. Available for only
innermost row or column fields.

**Related Function**

SHOW.LEVELS   Displays a specific number of levels of an outline

Return to [top](#Q)

SHOW.DIALOG
