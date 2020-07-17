PIVOT.ADD.FIELDS

Add fields onto a PivotTable report.

**Syntax**

**PIVOT.ADD.FIELDS**(name, row\_array, column\_array, page\_array,
add\_to\_table)

Name    is the name of the PivotTable report to which the user wants to
add fields. If name is omitted, Microsoft Excel will use the PivotTable
report containing the active cell.

Row\_array    is an array of text constants consisting of the names of
the fields which the user would like to add to the PivotTable report as
row fields.

Column\_array    is an array of text constants consisting of the names
of the fields which the user would like to add to the PivotTable report
as column fields.

Page\_array    is an array of text constants consisting of the names of
the fields which the user would like to add to the PivotTable report as
page fields.

Add\_to\_table    is a logical value which if TRUE adds the fields
specified by row\_array, column\_array and page\_array to the existing
fields on the PivotTable report. If add\_to\_table is FALSE, Microsoft
Excel will replace the fields already along the rows, columns and pages
with the fields specified by row\_array, column\_array and page\_array.

**Remarks**

If name is not a valid PivotTable name, then the \#VALUE\! error value
is returned.

**Related Functions**

[PIVOT.ADD.DATA](PIVOT.ADD.DATA.md)   Adds a field to a PivotTable report as a data field

[PIVOT.FIELD](PIVOT.FIELD.md)   Pivots fields within a PivotTable report

[PIVOT.FIELD.GROUP](PIVOT.FIELD.GROUP.md)   Creates groups within a PivotTable report

[PIVOT.FIELD.PROPERTIES](PIVOT.FIELD.PROPERTIES.md)   Changes the properties of a field inside a
[P](P.md)ivotTable report

[PIVOT.FIELD.UNGROUP](PIVOT.FIELD.UNGROUP.md)   Ungroups all selected groups within a PivotTable
report

[PIVOT.ITEM](PIVOT.ITEM.md)   Moves an item within a PivotTable report

[PIVOT.ITEM.PROPERTIES](PIVOT.ITEM.PROPERTIES.md)   Changes the properties of an item within a
header field

[PIVOT.REFRESH](PIVOT.REFRESH.md)   Refreshes a PivotTable report

[PIVOT.SHOW.PAGES](PIVOT.SHOW.PAGES.md)   Creates new sheets in the workbook containing the
active cell

[PIVOT.TABLE.WIZARD](PIVOT.TABLE.WIZARD.md)   Creates an empty PivotTable report



Return to [README.md](README.md)

PIVOT.ADD.FIELDS

Add fields onto a PivotTable report.

**Syntax**

**PIVOT.ADD.FIELDS**(name, row\_array, column\_array, page\_array,
add\_to\_table)

Name    is the name of the PivotTable report to which the user wants to
add fields. If name is omitted, Microsoft Excel will use the PivotTable
report containing the active cell.

Row\_array    is an array of text constants consisting of the names of
the fields which the user would like to add to the PivotTable report as
row fields.

Column\_array    is an array of text constants consisting of the names
of the fields which the user would like to add to the PivotTable report
as column fields.

Page\_array    is an array of text constants consisting of the names of
the fields which the user would like to add to the PivotTable report as
page fields.

Add\_to\_table    is a logical value which if TRUE adds the fields
specified by row\_array, column\_array and page\_array to the existing
fields on the PivotTable report. If add\_to\_table is FALSE, Microsoft
Excel will replace the fields already along the rows, columns and pages
with the fields specified by row\_array, column\_array and page\_array.

**Remarks**

If name is not a valid PivotTable name, then the \#VALUE\! error value
is returned.

**Related Functions**

[PIVOT.ADD.DATA](PIVOT.ADD.DATA.md)   Adds a field to a PivotTable report as a data field

[PIVOT.FIELD](PIVOT.FIELD.md)   Pivots fields within a PivotTable report

[PIVOT.FIELD.GROUP](PIVOT.FIELD.GROUP.md)   Creates groups within a PivotTable report

[PIVOT.FIELD.PROPERTIES](PIVOT.FIELD.PROPERTIES.md)   Changes the properties of a field inside a
[P](P.md)ivotTable report

[PIVOT.FIELD.UNGROUP](PIVOT.FIELD.UNGROUP.md)   Ungroups all selected groups within a PivotTable
report

[PIVOT.ITEM](PIVOT.ITEM.md)   Moves an item within a PivotTable report

[PIVOT.ITEM.PROPERTIES](PIVOT.ITEM.PROPERTIES.md)   Changes the properties of an item within a
header field

[PIVOT.REFRESH](PIVOT.REFRESH.md)   Refreshes a PivotTable report

[PIVOT.SHOW.PAGES](PIVOT.SHOW.PAGES.md)   Creates new sheets in the workbook containing the
active cell

[PIVOT.TABLE.WIZARD](PIVOT.TABLE.WIZARD.md)   Creates an empty PivotTable report



Return to [README.md](README.md)

