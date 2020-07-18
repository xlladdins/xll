FILTER

Filters lists of data one column at a time. Only one list can be
filtered on any one sheet at a time.

**Syntax**

**FILTER**(field\_num, criteria1, operation, criteria2)

**FILTER**?(field\_num, criteria1, operation, criteria2)

Field\_num**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is the number of the field that you
want to filter. Fields are numbered from left to right starting with 1.

Criteria1**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a text string specifying criteria
for filtering a list, such as "\>2". If you want to include all items in
the list, omit this argument.

Operation**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a number that specifies how you want
criteria2 used with criteria1:

|            |                    |
| ---------- | ------------------ |
| **Number** | **Operation Used** |
| 1          | AND                |
| 2          | OR                 |

Criteria2**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a text string specifying criteria
for filtering a list, such as "\>2". If you include this argument,
operation is required.

**Remarks**

If you omit all arguments, FILTER toggles the display of filter arrows.

**Related Function**

[FILTER.ADVANCED](FILTER.ADVANCED.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Lets you set options for filtering a
list



Return to [README](README.md)

