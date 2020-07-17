PIVOT.FIELD

Pivots a field within a PivotTable report.

**Syntax**

**PIVOT.FIELD**(name, pivot\_field\_name, orientation, position)

Name    is the name of the PivotTable report in which the user wants to
pivot fields. If name is omitted, Microsoft Excel will use the
PivotTable report containing the active cell.

Pivot\_field\_name    is the name of the field which the user wishes to
pivot to another part of the PivotTable report. This argument is given
as a text constant or a reference to a text constant. If field\_name is
omitted, Microsoft Excel uses the field containing the active cell.

Orientation    is an integer representing the destination of the field
which is being pivoted. If this argument is omitted, then the
orientation remains unchanged. The integers refer to orientations as
follows:

|           |                 |
| --------- | --------------- |
| **Value** | **Orientation** |
| 0         | Hidden          |
| 1         | Row             |
| 2         | Column          |
| 3         | Page            |
| 4         | Data            |

Position    is an integer representing where in the orientation the
fields will be positioned. Position 1 is the leftmost header position in
the row header and the topmost position in the column header. This
argument is ignored if orientation is set to 0. If the position argument
is omitted, it will default to the last position in the field.

**Remarks**

  - > The function returns TRUE if successful.

  - > If name is not a valid PivotTable name then the \#VALUE\! error
    > value is returned.

  - > If pivot\_field\_name is not a text constant or contains text
    > which is not a valid field name for the PivotTable report then the
    > \#VALUE\! error value is returned.

  - > If destination is not an integer between 0 and 4, then the
    > \#VALUE\! error value is returned.

>  

**Related Functions**

PIVOT.ADD.DATA   Adds a field to a PivotTable report as a data field

PIVOT.ADD.FIELDS   Adds fields to a PivotTable report

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

PIVOT.TABLE.WIZARD   Creates an empty PivotTable report


