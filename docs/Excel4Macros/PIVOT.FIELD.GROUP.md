PIVOT.FIELD.GROUP

Creates groups within a PivotTable report.

**Syntax**

**PIVOT.FIELD.GROUP**(start, end, by, periods)

**PIVOT.FIELD.GROUP**?(start, end, by, periods)

Start    is the beginning date of the range to be grouped. If start is
TRUE or omitted, it is assumed to be the first value in the field.

End    is the ending date of the range to be grouped. If end is TRUE or
omitted it is assumed to be the last value in the field.

By    is the size of the groups to be created. If by is omitted,
Microsoft Excel uses a default group size. If grouping a date field and
if periods is not 8(days) then by is ignored.

Periods    is a number between 1 and 127. It is calculated by summing
the values in the following table corresponding to the periods into
which you want to group your dates. This argument is ignored if the
field is not a date field. This argument takes precedence over By if
they are both specified for a date field.

|           |             |
| --------- | ----------- |
| **Value** | **Periods** |
| 1         | Seconds     |
| 2         | Minutes     |
| 4         | Hours       |
| 8         | Days        |
| 16        | Months      |
| 32        | Quarters    |
| 64        | Years       |

**Remarks**

  - > This function returns TRUE if the grouping is successful. The
    > \#N/A error value is returned if the grouping failed.

  - > If no arguments are specified and multiple items within the header
    > field are selected then this function groups those selected items.

  - > If no arguments are specified and a single item within the header
    > field is selected then the \#VALUE\! error value is returned.

>  

**Related Functions**

PIVOT.ADD.DATA   Adds a field to a PivotTable report as a data field

PIVOT.ADD.FIELDS   Adds fields to a PivotTable report

PIVOT.FIELD   Pivots fields within a PivotTable report

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


PIVOT.FIELD.GROUP

Creates groups within a PivotTable report.

**Syntax**

**PIVOT.FIELD.GROUP**(start, end, by, periods)

**PIVOT.FIELD.GROUP**?(start, end, by, periods)

Start    is the beginning date of the range to be grouped. If start is
TRUE or omitted, it is assumed to be the first value in the field.

End    is the ending date of the range to be grouped. If end is TRUE or
omitted it is assumed to be the last value in the field.

By    is the size of the groups to be created. If by is omitted,
Microsoft Excel uses a default group size. If grouping a date field and
if periods is not 8(days) then by is ignored.

Periods    is a number between 1 and 127. It is calculated by summing
the values in the following table corresponding to the periods into
which you want to group your dates. This argument is ignored if the
field is not a date field. This argument takes precedence over By if
they are both specified for a date field.

|           |             |
| --------- | ----------- |
| **Value** | **Periods** |
| 1         | Seconds     |
| 2         | Minutes     |
| 4         | Hours       |
| 8         | Days        |
| 16        | Months      |
| 32        | Quarters    |
| 64        | Years       |

**Remarks**

  - > This function returns TRUE if the grouping is successful. The
    > \#N/A error value is returned if the grouping failed.

  - > If no arguments are specified and multiple items within the header
    > field are selected then this function groups those selected items.

  - > If no arguments are specified and a single item within the header
    > field is selected then the \#VALUE\! error value is returned.

>  

**Related Functions**

PIVOT.ADD.DATA   Adds a field to a PivotTable report as a data field

PIVOT.ADD.FIELDS   Adds fields to a PivotTable report

PIVOT.FIELD   Pivots fields within a PivotTable report

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


PIVOT.FIELD.GROUP

Creates groups within a PivotTable report.

**Syntax**

**PIVOT.FIELD.GROUP**(start, end, by, periods)

**PIVOT.FIELD.GROUP**?(start, end, by, periods)

Start    is the beginning date of the range to be grouped. If start is
TRUE or omitted, it is assumed to be the first value in the field.

End    is the ending date of the range to be grouped. If end is TRUE or
omitted it is assumed to be the last value in the field.

By    is the size of the groups to be created. If by is omitted,
Microsoft Excel uses a default group size. If grouping a date field and
if periods is not 8(days) then by is ignored.

Periods    is a number between 1 and 127. It is calculated by summing
the values in the following table corresponding to the periods into
which you want to group your dates. This argument is ignored if the
field is not a date field. This argument takes precedence over By if
they are both specified for a date field.

|           |             |
| --------- | ----------- |
| **Value** | **Periods** |
| 1         | Seconds     |
| 2         | Minutes     |
| 4         | Hours       |
| 8         | Days        |
| 16        | Months      |
| 32        | Quarters    |
| 64        | Years       |

**Remarks**

  - > This function returns TRUE if the grouping is successful. The
    > \#N/A error value is returned if the grouping failed.

  - > If no arguments are specified and multiple items within the header
    > field are selected then this function groups those selected items.

  - > If no arguments are specified and a single item within the header
    > field is selected then the \#VALUE\! error value is returned.

>  

**Related Functions**

PIVOT.ADD.DATA   Adds a field to a PivotTable report as a data field

PIVOT.ADD.FIELDS   Adds fields to a PivotTable report

PIVOT.FIELD   Pivots fields within a PivotTable report

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


PIVOT.FIELD.GROUP

Creates groups within a PivotTable report.

**Syntax**

**PIVOT.FIELD.GROUP**(start, end, by, periods)

**PIVOT.FIELD.GROUP**?(start, end, by, periods)

Start    is the beginning date of the range to be grouped. If start is
TRUE or omitted, it is assumed to be the first value in the field.

End    is the ending date of the range to be grouped. If end is TRUE or
omitted it is assumed to be the last value in the field.

By    is the size of the groups to be created. If by is omitted,
Microsoft Excel uses a default group size. If grouping a date field and
if periods is not 8(days) then by is ignored.

Periods    is a number between 1 and 127. It is calculated by summing
the values in the following table corresponding to the periods into
which you want to group your dates. This argument is ignored if the
field is not a date field. This argument takes precedence over By if
they are both specified for a date field.

|           |             |
| --------- | ----------- |
| **Value** | **Periods** |
| 1         | Seconds     |
| 2         | Minutes     |
| 4         | Hours       |
| 8         | Days        |
| 16        | Months      |
| 32        | Quarters    |
| 64        | Years       |

**Remarks**

  - > This function returns TRUE if the grouping is successful. The
    > \#N/A error value is returned if the grouping failed.

  - > If no arguments are specified and multiple items within the header
    > field are selected then this function groups those selected items.

  - > If no arguments are specified and a single item within the header
    > field is selected then the \#VALUE\! error value is returned.

>  

**Related Functions**

PIVOT.ADD.DATA   Adds a field to a PivotTable report as a data field

PIVOT.ADD.FIELDS   Adds fields to a PivotTable report

PIVOT.FIELD   Pivots fields within a PivotTable report

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


PIVOT.FIELD.GROUP

Creates groups within a PivotTable report.

**Syntax**

**PIVOT.FIELD.GROUP**(start, end, by, periods)

**PIVOT.FIELD.GROUP**?(start, end, by, periods)

Start    is the beginning date of the range to be grouped. If start is
TRUE or omitted, it is assumed to be the first value in the field.

End    is the ending date of the range to be grouped. If end is TRUE or
omitted it is assumed to be the last value in the field.

By    is the size of the groups to be created. If by is omitted,
Microsoft Excel uses a default group size. If grouping a date field and
if periods is not 8(days) then by is ignored.

Periods    is a number between 1 and 127. It is calculated by summing
the values in the following table corresponding to the periods into
which you want to group your dates. This argument is ignored if the
field is not a date field. This argument takes precedence over By if
they are both specified for a date field.

|           |             |
| --------- | ----------- |
| **Value** | **Periods** |
| 1         | Seconds     |
| 2         | Minutes     |
| 4         | Hours       |
| 8         | Days        |
| 16        | Months      |
| 32        | Quarters    |
| 64        | Years       |

**Remarks**

  - > This function returns TRUE if the grouping is successful. The
    > \#N/A error value is returned if the grouping failed.

  - > If no arguments are specified and multiple items within the header
    > field are selected then this function groups those selected items.

  - > If no arguments are specified and a single item within the header
    > field is selected then the \#VALUE\! error value is returned.

>  

**Related Functions**
**Related Functions**

PIVOT.ADD.DATA   Adds a field to a PivotTable report as a data field

PIVOT.ADD.FIELDS   Adds fields to a PivotTable report

PIVOT.FIELD   Pivots fields within a PivotTable report

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


PIVOT.FIELD.GROUP

Creates groups within a PivotTable report.

**Syntax**

**PIVOT.FIELD.GROUP**(start, end, by, periods)

**PIVOT.FIELD.GROUP**?(start, end, by, periods)

Start    is the beginning date of the range to be grouped. If start is
TRUE or omitted, it is assumed to be the first value in the field.

End    is the ending date of the range to be grouped. If end is TRUE or
omitted it is assumed to be the last value in the field.

By    is the size of the groups to be created. If by is omitted,
Microsoft Excel uses a default group size. If grouping a date field and
if periods is not 8(days) then by is ignored.

Periods    is a number between 1 and 127. It is calculated by summing
the values in the following table corresponding to the periods into
which you want to group your dates. This argument is ignored if the
field is not a date field. This argument takes precedence over By if
they are both specified for a date field.

|           |             |
| --------- | ----------- |
| **Value** | **Periods** |
| 1         | Seconds     |
| 2         | Minutes     |
| 4         | Hours       |
| 8         | Days        |
| 16        | Months      |
| 32        | Quarters    |
| 64        | Years       |

**Remarks**

  - > This function returns TRUE if the grouping is successful. The
    > \#N/A error value is returned if the grouping failed.

  - > If no arguments are specified and multiple items within the header
    > field are selected then this function groups those selected items.

  - > If no arguments are specified and a single item within the header
    > field is selected then the \#VALUE\! error value is returned.

>  

**Related Functions**

PIVOT.ADD.DATA   Adds a field to a PivotTable report as a data field

PIVOT.ADD.FIELDS   Adds fields to a PivotTable report

PIVOT.FIELD   Pivots fields within a PivotTable report

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


PIVOT.FIELD.GROUP

Creates groups within a PivotTable report.

**Syntax**

**PIVOT.FIELD.GROUP**(start, end, by, periods)

**PIVOT.FIELD.GROUP**?(start, end, by, periods)

Start    is the beginning date of the range to be grouped. If start is
TRUE or omitted, it is assumed to be the first value in the field.

End    is the ending date of the range to be grouped. If end is TRUE or
omitted it is assumed to be the last value in the field.

By    is the size of the groups to be created. If by is omitted,
Microsoft Excel uses a default group size. If grouping a date field and
if periods is not 8(days) then by is ignored.

Periods    is a number between 1 and 127. It is calculated by summing
the values in the following table corresponding to the periods into
which you want to group your dates. This argument is ignored if the
field is not a date field. This argument takes precedence over By if
they are both specified for a date field.

|           |             |
| --------- | ----------- |
| **Value** | **Periods** |
| 1         | Seconds     |
| 2         | Minutes     |
| 4         | Hours       |
| 8         | Days        |
| 16        | Months      |
| 32        | Quarters    |
| 64        | Years       |

**Remarks**

  - > This function returns TRUE if the grouping is successful. The
    > \#N/A error value is returned if the grouping failed.

  - > If no arguments are specified and multiple items within the header
    > field are selected then this function groups those selected items.

  - > If no arguments are specified and a single item within the header
    > field is selected then the \#VALUE\! error value is returned.

>  

**Related Functions**

PIVOT.ADD.DATA   Adds a field to a PivotTable report as a data field

PIVOT.ADD.FIELDS   Adds fields to a PivotTable report

PIVOT.FIELD   Pivots fields within a PivotTable report

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


PIVOT.FIELD.GROUP

Creates groups within a PivotTable report.

**Syntax**

**PIVOT.FIELD.GROUP**(start, end, by, periods)

**PIVOT.FIELD.GROUP**?(start, end, by, periods)

Start    is the beginning date of the range to be grouped. If start is
TRUE or omitted, it is assumed to be the first value in the field.

End    is the ending date of the range to be grouped. If end is TRUE or
omitted it is assumed to be the last value in the field.

By    is the size of the groups to be created. If by is omitted,
Microsoft Excel uses a default group size. If grouping a date field and
if periods is not 8(days) then by is ignored.

Periods    is a number between 1 and 127. It is calculated by summing
the values in the following table corresponding to the periods into
which you want to group your dates. This argument is ignored if the
field is not a date field. This argument takes precedence over By if
they are both specified for a date field.

|           |             |
| --------- | ----------- |
| **Value** | **Periods** |
| 1         | Seconds     |
| 2         | Minutes     |
| 4         | Hours       |
| 8         | Days        |
| 16        | Months      |
| 32        | Quarters    |
| 64        | Years       |

**Remarks**

  - > This function returns TRUE if the grouping is successful. The
    > \#N/A error value is returned if the grouping failed.

  - > If no arguments are specified and multiple items within the header
    > field are selected then this function groups those selected items.

  - > If no arguments are specified and a single item within the header
    > field is selected then the \#VALUE\! error value is returned.

>  

**Related Functions**

PIVOT.ADD.DATA   Adds a field to a PivotTable report as a data field

PIVOT.ADD.FIELDS   Adds fields to a PivotTable report

PIVOT.FIELD   Pivots fields within a PivotTable report

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


PIVOT.FIELD.GROUP

Creates groups within a PivotTable report.

**Syntax**

**PIVOT.FIELD.GROUP**(start, end, by, periods)

**PIVOT.FIELD.GROUP**?(start, end, by, periods)

Start    is the beginning date of the range to be grouped. If start is
TRUE or omitted, it is assumed to be the first value in the field.

End    is the ending date of the range to be grouped. If end is TRUE or
omitted it is assumed to be the last value in the field.

By    is the size of the groups to be created. If by is omitted,
Microsoft Excel uses a default group size. If grouping a date field and
if periods is not 8(days) then by is ignored.

Periods    is a number between 1 and 127. It is calculated by summing
the values in the following table corresponding to the periods into
which you want to group your dates. This argument is ignored if the
field is not a date field. This argument takes precedence over By if
they are both specified for a date field.

|           |             |
| --------- | ----------- |
| **Value** | **Periods** |
| 1         | Seconds     |
| 2         | Minutes     |
| 4         | Hours       |
| 8         | Days        |
| 16        | Months      |
| 32        | Quarters    |
| 64        | Years       |

**Remarks**

  - > This function returns TRUE if the grouping is successful. The
    > \#N/A error value is returned if the grouping failed.

  - > If no arguments are specified and multiple items within the header
    > field are selected then this function groups those selected items.

  - > If no arguments are specified and a single item within the header
    > field is selected then the \#VALUE\! error value is returned.

>  

**Related Functions**

[PIVOT.ADD.DATA](PIVOT.ADD.DATA.md)   Adds a field to a PivotTable report as a data field

[PIVOT.ADD.FIELDS](PIVOT.ADD.FIELDS.md)   Adds fields to a PivotTable report

[PIVOT.FIELD](PIVOT.FIELD.md)   Pivots fields within a PivotTable report

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


