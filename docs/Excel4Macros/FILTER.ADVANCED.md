FILTER.ADVANCED

Equivalent to choosing the Advanced Filter command from the Filter
submenu on the Data menu. Lets you set options for filtering a list.

**Syntax**

**FILTER.ADVANCED**(**operation**, **list\_ref**, criteria\_ref,
copy\_ref, unique)

**FILTER.ADVANCED**?(operation, list\_ref, criteria\_ref, copy\_ref,
unique)

Operation**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a number specifying whether to copy
the filter list to a new location. To filter a list without copying, use
1; to copy the filter list to a new location, use 2.

List\_ref**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;specifies the location of the list to
be filtered. If operation is 1, then list\_ref must be on the active
sheet.

Criteria\_ref**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a reference to a range
containing criteria for filtering the list. If omitted, uses "All" as
the criteria.

Copy\_ref**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a reference on the active sheet
where you want the filtered list copied. Ignored if operation is 1.

Unique**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a logical value that specifies whether
only unique records are displayed. To display only unique records, use
TRUE. To display all records that match the criteria, use FALSE or omit
this argument.

**Related Function**

[FILTER](FILTER.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Filters lists of data one column at a time



Return to [README](README.md)

