PIVOT.ADD.DATA

Adds a field to a PivotTable report.

**Syntax**

**PIVOT.ADD.DATA**(name, pivot\_field\_name, new\_name, position,
function, calculation, base\_field, base\_item, format\_text)

Name    is the name of the PivotTable report to which the user wants to
add as a data field. If name is omitted, Microsoft Excel will use the
PivotTable report containing the active cell.

Pivot\_field\_name    is the name of a field which the user would like
to add to the PivotTable report as data or as text.

New\_name    is the name you would like to give to the new field once it
is added to your PivotTable report. If this argument is omitted,
Microsoft Excel will pick a default name for you. This function returns
new\_name or the name Microsoft Excel gives the field.

Position    is the position within all the Data fields you would like to
place the new data field. If position is omitted, the field will be
added as the last data field.

Function    is a number from 2 to 2048 specifying how the new field is
to be calculated. To compute the value to place in this column use one
value from the following table. If function is omitted, SUM will be
used. If the field is a numeric field or text field, COUNTA will be
used.

|           |              |
| --------- | ------------ |
| **Value** | **Function** |
| 2         | SUM          |
| 4         | COUNTA       |
| 8         | COUNT        |
| 16        | AVERAGE      |
| 32        | MAX          |
| 64        | MIN          |
| 128       | PRODUCT      |
| 256       | STDEV        |
| 512       | STDEVP       |
| 1024      | VAR          |
| 2048      | VARP         |

Calculation    is a number between 1 and 9 representing which custom
calculation you would like to apply to this data field. This corresponds
to the Show Data As drop-down box on the PivotTable Field dialog box. If
this argument is omitted, no special calculation will be applied to the
data field.

|           |                   |
| --------- | ----------------- |
| **Value** | **Calculation**   |
| 1         | Normal            |
| 2         | Difference From   |
| 3         | % Of Item         |
| 4         | % Difference From |
| 5         | Running Total In  |
| 6         | % of Row          |
| 7         | % of Column       |
| 8         | % of Total        |
| 9         | Index             |

Base\_Field    is the field on which you want to base the calculation.

Base\_Item    is the item within base\_field on which you want to base
the calculation.

Format\_text    is the type of number format you want to apply to the
PivotTable data. Corresponds to the number button in the PivotTable
Field dialog box, which appears when you click the PivotTable Field
command on the Data menu when the selection is in a data field.

**Remarks**

  - > If name is not a valid PivotTable name, then the \#VALUE\! error
    > value is returned.

  - > If field\_name is not a valid field for the current PivotTable
    > report then the \#VALUE\! error value is returned.

>  

**Related Functions**

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

PIVOT.TABLE.WIZARD   Creates an empty PivotTable report


