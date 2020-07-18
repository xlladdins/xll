# PIVOT.SHOW.PAGES

Creates new sheets in the workbook containing the active cell. The
function will iterate through each item in page\_field and create a new
PivotTable report on a new sheet with the page field set to that
particular item.

**Syntax**

**PIVOT.SHOW.PAGES**(name, page\_field)

Name&nbsp;&nbsp;&nbsp;&nbsp;is the name of the target PivotTable report.
If name is omitted, Microsoft Excel will use the PivotTable report
containing the active cell.

Page\_field&nbsp;&nbsp;&nbsp;&nbsp;is the name of a page field in the
PivotTable report specified by the name argument.

**Remarks**

  - > If the function is successful, it returns TRUE; otherwise, it
    > returns the \#VALUE\! error value.

  - > If name is not a valid PivotTable name then the \#VALUE\! error
    > value is returned.

**Related Functions**

[PIVOT.ADD.DATA](PIVOT.ADD.DATA.md)&nbsp;&nbsp;&nbsp;Adds a field to a PivotTable report as a
data field

[PIVOT.ADD.FIELDS](PIVOT.ADD.FIELDS.md)&nbsp;&nbsp;&nbsp;Adds fields to a PivotTable report

[PIVOT.FIELD](PIVOT.FIELD.md)&nbsp;&nbsp;&nbsp;Pivots fields within a PivotTable report

[PIVOT.FIELD.GROUP](PIVOT.FIELD.GROUP.md)&nbsp;&nbsp;&nbsp;Creates groups within a PivotTable
report

[PIVOT.FIELD.PROPERTIES](PIVOT.FIELD.PROPERTIES.md)&nbsp;&nbsp;&nbsp;Changes the properties of a
field inside a PivotTable report

[PIVOT.FIELD.UNGROUP](PIVOT.FIELD.UNGROUP.md)&nbsp;&nbsp;&nbsp;Ungroups all selected groups within
a PivotTable report

[PIVOT.ITEM](PIVOT.ITEM.md)&nbsp;&nbsp;&nbsp;Moves an item within a PivotTable report

[PIVOT.ITEM.PROPERTIES](PIVOT.ITEM.PROPERTIES.md)&nbsp;&nbsp;&nbsp;Changes the properties of an item
within a header field

[PIVOT.REFRESH](PIVOT.REFRESH.md)&nbsp;&nbsp;&nbsp;Refreshes a PivotTable report

[PIVOT.TABLE.WIZARD](PIVOT.TABLE.WIZARD.md)&nbsp;&nbsp;&nbsp;Creates an empty PivotTable report



Return to [README](README.md)

