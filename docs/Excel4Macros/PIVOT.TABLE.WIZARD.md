PIVOT.TABLE.WIZARD
==================

Creates an empty PivotTable report.

**Syntax**

**PIVOT.TABLE.WIZARD**(type, source, destination, name, row\_grand,
col\_grand, save\_data, apply\_auto\_format, autopage)

**PIVOT.TABLE.WIZARD**?(type, source, destination, name, row\_grand,
col\_grand, save\_data, apply\_auto\_format, autopage)

Type    is a number specifying the type of source data used to create
the PivotTable report.

  ----------- ----------------------------------
  **Value**   **Type of source data**
  1           Microsoft Excel list or database
  2           External data source
  3           Multiple consolidation ranges
  4           Another PivotTable report
  ----------- ----------------------------------

Source    can be one of four things. If type is 1, then source is a cell
reference or name to the range to be used as the PivotTable source. If
type is 2, then source is a one-dimensional array describing the
external database to be used as the PivotTable source. If type is 3,
then source is a multi-dimensional array listing the cell ranges and
associated page field items describing the consolidation PivotTable
source. If type is 4, then source is the name of another PivotTable
report with which to share its source.

Destination    is a cell reference or name. The upper-left corner of
this range will act as the upper-left corner of the PivotTable report
which will be created. If destination is omitted, Microsoft Excel will
create the PivotTable report on a new sheet.

Name    is the name of the PivotTable report to be created given as a
text. If name is omitted, Microsoft Excel uses a default name.

Row\_grand    is a logical value which if TRUE displays a Grand Total
for each row on the PivotTable report. If FALSE, a Grand Total for each
row is not displayed.

Col\_grand    is a logical value which if TRUE displays a Grand Total
for each column. If FALSE, a Grand Total for each column is not
displayed.

Save\_data    is a logical value which if TRUE causes the data for the
PivotTable report to be saved along with the PivotTable definition. If
FALSE, the data is not saved along with the PivotTable definition.

Apply\_auto\_format    is a logical value which if TRUE causes
autoformatting upon pivotting or refreshing. If FALSE, the PivotTable
report will not be formatted automatically upon pivoting or refreshing.

Autopage    Applies only to type 3. This argument is a logical value
which if TRUE or omitted causes Microsoft Excel to create a page field
automatically. If FALSE, the page field must be created manually.

**Remarks**

-   The function will return TRUE if successful; otherwise, returns the
    > \#VALUE! error value.

-   If destination is not a valid Microsoft Excel reference, then
    > \#VALUE! error value is returned.

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

PIVOT.REFRESH   Refreshes a PivotTable report

PIVOT.SHOW.PAGES   Creates new sheets in the workbook containing the
active cell

Return to [top](#H)

PLACEMENT
