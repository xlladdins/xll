PIVOT.ITEM
==========

Moves an item within a PivotTable report.

**Syntax**

**PIVOT.ITEM**(name, pivot\_field\_name, pivot\_item\_name,position)

Name    is the name of the PivotTable report within which an item will
be repositioned. If name is omitted, Microsoft Excel will use the
PivotTable report containing the active cell.

Pivot\_field\_name    is the name of the field within which an item will
be repositioned, given as a text string. If pivot\_field\_name is
omitted, Microsoft Excel will use the field containing the active cell.
If the active cell is not within a field, then this argument is
required.

Pivot\_item\_name    is the name of the item to be repositioned in its
field (given as a text constant). If it is omitted, Microsoft Excel uses
the item containing the active cell. If the active cell is not contained
within an item, then this argument is required.

Position    is a number representing where in the field the items will
be moved. Position 1 is the topmost position within the row field and
the leftmost position within the column field and the highest position
within the page field. If the position argument is omitted, it will
default to the last position in the field.

**Remarks**

-   If an item is set to be visible, but its display is suppressed
    > because there is no data, this item still occupies a valid
    > position.

-   If name is not a valid PivotTable name then the \#VALUE! error value
    > is returned.

-   If pivot\_field\_name is not a text string, or if pivot\_field\_name
    > is not a text string within a valid field name, then \#VALUE! is
    > returned.

-   If pivot\_item\_name is an item which is not currently showing in
    > the PivotTable report because it does not exist in the field
    > pivot\_field\_name, the \#VALUE! error value is returned.

**Related Functions**

PIVOT.ADD.DATA   Adds a field to a PivotTable report as a data field

PIVOT.ADD.FIELDS   Adds fields to a PivotTable report

PIVOT.FIELD   Pivots fields within a PivotTable report

PIVOT.FIELD.GROUP   Creates groups within a PivotTable report

PIVOT.FIELD.PROPERTIES   Changes the properties of a field inside a
PivotTable report

PIVOT.FIELD.UNGROUP   Ungroups all selected groups within a PivotTable
report

PIVOT.ITEM.PROPERTIES   Changes the properties of an item within a
header field

PIVOT.REFRESH   Refreshes a PivotTable report

PIVOT.SHOW.PAGES   Creates new sheets in the workbook containing the
active cell

PIVOT.TABLE.WIZARD   Creates an empty PivotTable report

Return to [top](#H)

PIVOT.ITEM.PROPERTIES