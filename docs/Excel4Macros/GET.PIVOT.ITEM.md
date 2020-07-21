# GET.PIVOT.ITEM

Returns information about an item in a PivotTable report.

**Syntax**

**GET.PIVOT.ITEM**(**type\_num**, pivot\_item\_name, pivot\_field\_name,
pivot\_table\_name)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a value from 1 to 9 the represents
the type of information you want about an item in a PivotTable report.

|               |                                                                                                                                                                                                                                                                                   |
| ------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Type\_num** | **Information**                                                                                                                                                                                                                                                                   |
| 1             | Returns the position of the item in its field. Returns \#N/A if pivot\_field\_name is a data field. Returns \#N/A\! if the item is hidden.                                                                                                                                        |
| 2             | Returns the reference to all the cells in the PivotTable header currently containing pivot\_item\_name. This reference is returned as text. If pivot\_item\_name is currently not showing in the PivotTable report, \#N/A\! is returned.                                          |
| 3             | Returns the reference to all the data in the PivotTable report which is qualified by pivot\_item\_name. This reference is returned as text. If pivot\_item\_name is currently not showing in the PivotTable report, \#N/A\! is returned.                                          |
| 4             | Returns an array of text constants representing the children of pivot\_item\_name if pivot\_item\_name is a parent. Otherwise the function returns \#N/A\!.                                                                                                                       |
| 5             | Returns a text constant representing the parent of pivot\_item\_name, if pivot\_item\_name exists as part of a group. Otherwise the function returns \#N/A\!.                                                                                                                     |
| 6             | Returns TRUE if pivot\_item\_name is a member of a group which is currently expanded to show detail. Returns FALSE if pivot\_item\_name is a member of a group currently collapsed to hide detail. If pivot\_item\_name is not a member of a group, the function returns \#N/A\!. |
| 7             | Returns TRUE if pivot\_item\_name is expanded to show detail. Returns FALSE if pivot\_item\_name is collapsed to hide detail.                                                                                                                                                     |
| 8             | Returns TRUE if the item pivot\_item\_name is currently visible, FALSE if it is hidden.                                                                                                                                                                                           |
| 9             | Returns the name of the item as it appeared in the original at a source. This will differ from the current item name only if the user changes the name of the item after creating the PivotTable report.                                                                          |

Pivot\_item\_name&nbsp;&nbsp;&nbsp;&nbsp;is the name of the item that
you want information about. If there is no item named pivot\_item\_name
in the PivotTable report, returns \#VALUE\!.

Pivot\_field\_name&nbsp;&nbsp;&nbsp;&nbsp;is the name of the field that
you want information about. If there is no field named
pivot\_field\_name in the PivotTable report, returns \#VALUE\!.

Pivot\_table\_name&nbsp;&nbsp;&nbsp;&nbsp;is the name of a PivotTable
report containing the field that you want information about. If omitted,
uses the PivotTable report containing the active cell. If the active
cell is not in a PivotTable report, the \#VALUE\! error value is
returned.

**Related Functions**

[GET.PIVOT.FIELD](GET.PIVOT.FIELD.md)&nbsp;&nbsp;&nbsp;Returns information about an item in a
[P](P.md)ivotTable report.

[GET.PIVOT.TABLE](GET.PIVOT.TABLE.md)&nbsp;&nbsp;&nbsp;Returns information about a PivotTable
report.



Return to [README](README.md#G)

