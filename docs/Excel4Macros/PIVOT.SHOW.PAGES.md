PIVOT.SHOW.PAGES

Creates new sheets in the workbook containing the active cell. The
function will iterate through each item in page\_field and create a new
PivotTable report on a new sheet with the page field set to that
particular item.

**Syntax**

**PIVOT.SHOW.PAGES**(name, page\_field)

Name    is the name of the target PivotTable report. If name is omitted,
Microsoft Excel will use the PivotTable report containing the active
cell.

Page\_field    is the name of a page field in the PivotTable report
specified by the name argument.

**Remarks**

  - > If the function is successful, it returns TRUE; otherwise, it
    > returns the \#VALUE\! error value.

  - > If name is not a valid PivotTable name then the \#VALUE\! error
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

PIVOT.REFRESH   Refreshes a PivotTable report

PIVOT.TABLE.WIZARD   Creates an empty PivotTable report


