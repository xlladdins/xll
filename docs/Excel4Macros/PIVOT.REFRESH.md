PIVOT.REFRESH
=============

Refreshes a PivotTable report.

**Syntax**

**PIVOT.REFRESH**(name)

Name    is the name of the PivotTable report the user would like to
refresh with current data. If name is omitted, Microsoft Excel will use
the PivotTable report containing the active cell.

**Remarks**

-   If the function is successful, it returns TRUE; otherwise, it
    > returns the \#VALUE! error value.

-   If name is not a valid PivotTable name, then the \#VALUE! error
    > value is returned.

**Related Functions**

PIVOT.ADD.DATA   Adds a field to a PivotTable report as a data field

PIVOT.ADD.FIELDS   Adds fields to a PivotTable report

PIVOT.FIELD   Pivots fields within a PivotTable report

PIVOT.FIELD.GROUP   Creates groups within a PivotTable report

PIVOT.FIELD.PROPERTIES   Changes the properties of a field inside a
PivotTable report

PIVOT.FIELD.UNGROUP   Ungroups all selected groups within a PivotTable
report

PIVOT.ITEM   Moves an item within a PivotTable report

PIVOT.ITEM.PROPERTIES   Changes the properties of an item within a
header field

PIVOT.SHOW.PAGES   Creates new sheets in the workbook containing the
active cell

PIVOT.TABLE.WIZARD   Creates an empty PivotTable report

Return to [top](#H)

PIVOT.SHOW.PAGES
