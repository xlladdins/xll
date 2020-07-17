SORT
====

Equivalent to clicking the Sort command on the Data menu. Sorts the rows
or columns of the selection according to the contents of a key row or
column within the selection. Use SORT to rearrange information into
ascending or descending order.

**Syntax 1**

For Worksheet and macro sheets

**SORT**(orientation, key1, order1, key2, order2, key3, order3, header,
custom, case)

**SORT**?(orientation, key1, order1, key2, order2, key3, order3, header,
custom, case)

**Syntax 2**

For PivotTable reports

**SORT**(orientation, key1, order1, type, custom)

**SORT**?(orientation, key1, order1, type, custom)

Orientation    is a number specifying whether to sort by rows or
columns. Enter 1 to sort top to bottom or 2 to sort left to right.

Key1    is a reference to the cell or cells you want to use as the first
sort key. The sort key identifies which column to sort by when sorting
rows or which row to sort by when sorting columns. For a PivotTable
report, if type is 1, then key1 is a cell reference which indicates what
value to sort by. There are two ways to specify sort keys:

+----------------------------------+----------------------------------+
| > **Type of key**                | > **Examples**                   |
+----------------------------------+----------------------------------+
| > An R1C1-style reference in the | > \"C2\" or \"C\[1\]\" or        |
| > form of text. If the reference | > \"Price\"                      |
| > is relative, it is assumed to  |                                  |
| > be relative to the active cell |                                  |
| > in the selection.              |                                  |
+----------------------------------+----------------------------------+

Order1    specifies whether to sort the row or column containing key1 in
ascending or descending order. Enter 1 to sort in ascending order or 2
to sort in descending order.

Key2, order2, key3, and order3    are similar to key1 and order1. Key2
specifies the second sort key, and order2 specifies whether to sort the
row or column containing key2 in ascending or descending order. Key3 and
order3 work similarly.

Header    is a number indicating how Microsoft Excel is to handle
headers on list.

+----------------+-------------------------------------------------------+
| > **Header**   | > **Defined**                                         |
+----------------+-------------------------------------------------------+
| > 0            | > Microsoft Excel will guess if there is a header     |
+----------------+-------------------------------------------------------+
| > 1            | > Forces Microsoft Excel to assume there is a header  |
+----------------+-------------------------------------------------------+
| > 2 or omitted | > Forces Microsoft Excel to assume there is no header |
+----------------+-------------------------------------------------------+

Type    is a number specifying whether to sort the field by labels or
values. Use one to sort by values or two to sort by labels.

Custom    is a number that specifies what kind of custom sorting you
want. This corresponds to the First Key Sort Order drop-down box in the
Sort Options dialog box. For a PivotTable report, custom is a number
indicating what custom sort order to use when sorting labels.

+--------------+--------------------------------------------------------------+
| > **Number** | > **Type of sort**                                           |
+--------------+--------------------------------------------------------------+
| > 1          | > Normal                                                     |
+--------------+--------------------------------------------------------------+
| > 2          | > Weekdays in abbreviated form (\"Sun\", \"Mon\", and so on) |
+--------------+--------------------------------------------------------------+
| > 3          | > Weekdays                                                   |
+--------------+--------------------------------------------------------------+
| > 4          | > Months in abbreviated form (\"Jan\" \"Feb\", and so on)    |
+--------------+--------------------------------------------------------------+
| > 5          | > Months                                                     |
+--------------+--------------------------------------------------------------+

Case    is a logical value that determines whether the sort is case
sensitive. If TRUE, the sort is case sensitive. If FALSE or omitted, the
sort will not be case sensitive.

**Tip   **If you want to sort using more than three keys, then sort the
data three keys at a time, starting with the least important group of
keys and progressing to the most important group, but listing the most
important key first within each group.

**Remarks**

In the dialog box form of this function, if the header argument is
omitted, then Microsoft Excel will guess whether or not there are
headers.

Return to [top](#Q)

SOUND.NOTE