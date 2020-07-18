# PIVOT.FIELD.PROPERTIES

Changes the properties of a field inside a PivotTable report.

**Syntax**

**PIVOT.FIELD.PROPERTIES**(name, pivot\_field\_name, new\_name,
orientation, function, formats)

**PIVOT.FIELD.PROPERTIES**?(name, pivot\_field\_name, new\_name,
orientation, function, formats)

Name&nbsp;&nbsp;&nbsp;&nbsp;is the name of the PivotTable report
containing the field which the user wants to edit. If name is omitted,
Microsoft Excel will use the PivotTable report containing the active
cell.

Pivot\_field\_name&nbsp;&nbsp;&nbsp;&nbsp;is the name of a field in the
PivotTable report which the user would like to edit, as text. If it is
omitted, Microsoft Excel uses the field containing the active cell.

New\_name&nbsp;&nbsp;&nbsp;&nbsp;is the name which you would like to
rename the current field. If it is omitted, the name of the current
field will not change.

Orientation&nbsp;&nbsp;&nbsp;&nbsp;is a number between 0 and 4
specifying which area will show the field containing the active cell. If
zero, then the field is deleted and all other arguments to this function
are ignored. If this argument is omitted, the orientation of the field
will not change.

|           |                   |
| --------- | ----------------- |
| **Value** | **Orientation**   |
| 0         | Delete            |
| 1         | Display as row    |
| 2         | Display as column |
| 3         | Display as page   |
| 4         | Display as data   |

Function&nbsp;&nbsp;&nbsp;&nbsp;is a number between 0 and 4094
specifying which calculation or subtotals to apply to the field. If you
will be showing the field in the header (orientation 1, 2, or 3), add up
the values from the table corresponding to the subtotals you would like
to show. If you will be showing the field as a data field (orientation
4), use one value from the table. If an entry in this column is left
blank, Microsoft Excel will not change the calculation or subtotal which
are currently attached to the field.

|           |              |
| --------- | ------------ |
| **Value** | **Function** |
| 0         | NO SUBTOTALS |
| 1         | AUTOMATIC    |
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

Formats&nbsp;&nbsp;&nbsp;&nbsp;is either a one- or a two- dimensional
array, depending on whether the field is a header field or a data field.

  - > If the active field is a header field (orientation argument is 1,
    > 2 or 3) then this is a two-dimensional array. Each row of the
    > array should consist of two entries. The first is a text string
    > corresponding to the item whose property is being changed. The
    > second element specifies whether the item will be hidden. If this
    > argument is TRUE, the item will be hidden and therefore will not
    > be displayed in the PivotTable report. If the argument is FALSE,
    > then the item will be displayed in the PivotTable report.

  - > If the active field is a data field, then the array is a
    > one-dimensional array with four elements. The first element is a
    > number between 1 and 9 specifying which calculation you wish to
    > apply to the current data field. This corresponds to the Show Data
    > As drop-down box on the PivotTable Field dialog box.

|           |                  |
| --------- | ---------------- |
| **Value** | **Format**       |
| 1         | Normal           |
| 2         | Difference From  |
| 3         | %Of Item         |
| 4         | %Difference From |
| 5         | Running Total In |
| 6         | %Of Row          |
| 7         | %Of Column       |
| 8         | %Of Subtotal     |
| 9         | Index            |

&nbsp;

  - > The second element contains a text string representing the field
    > to which your data field is related. This argument is not
    > necessary for the Normal calculation. If omitted, Microsoft Excel
    > will use the first field that would appear in the Base Field list
    > box.

  - > The third element must contain a text string representing an item
    > in the base field on which to base your calculation. Note that
    > this argument is not necessary for calculations like Running Total
    > In which relies only on a Base Field. If omitted, Microsoft Excel
    > will use the first item that would appear in the Base Item list
    > box.

  - > The fourth element is a text string representing the number format
    > you wish to apply to the data field.


**Remarks**

  - > If pivot\_field\_name is not a valid field name for the PivotTable
    > report then the \#VALUE\! error value is returned.

  - > If name is not a valid PivotTable name, then the \#VALUE\! error
    > value is returned.

  - > If the orientation and function arguments do not contain numbers
    > or if these arguments contain numbers which are out of range then
    > the \#VALUE\! error value is returned.


**Related Functions**

[PIVOT.ADD.DATA](PIVOT.ADD.DATA.md)&nbsp;&nbsp;&nbsp;Adds a field to a PivotTable report as a
data field

[PIVOT.ADD.FIELDS](PIVOT.ADD.FIELDS.md)&nbsp;&nbsp;&nbsp;Adds fields to a PivotTable report

[PIVOT.FIELD](PIVOT.FIELD.md)&nbsp;&nbsp;&nbsp;Pivots fields within a PivotTable report

[PIVOT.FIELD.GROUP](PIVOT.FIELD.GROUP.md)&nbsp;&nbsp;&nbsp;Creates groups within a PivotTable
report

[PIVOT.FIELD.UNGROUP](PIVOT.FIELD.UNGROUP.md)&nbsp;&nbsp;&nbsp;Ungroups all selected groups within
a PivotTable report

[PIVOT.ITEM](PIVOT.ITEM.md)&nbsp;&nbsp;&nbsp;Moves an item within a PivotTable report

[PIVOT.ITEM.PROPERTIES](PIVOT.ITEM.PROPERTIES.md)&nbsp;&nbsp;&nbsp;Changes the properties of an item
within a header field

[PIVOT.REFRESH](PIVOT.REFRESH.md)&nbsp;&nbsp;&nbsp;Refreshes a PivotTable report

[PIVOT.SHOW.PAGES](PIVOT.SHOW.PAGES.md)&nbsp;&nbsp;&nbsp;Creates new sheets in the workbook
containing the active cell

[PIVOT.TABLE.WIZARD](PIVOT.TABLE.WIZARD.md)&nbsp;&nbsp;&nbsp;Creates an empty PivotTable report



Return to [README](README.md)

