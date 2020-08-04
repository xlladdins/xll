# EDIT.DELETE

Equivalent to clicking the Delete command on the Edit menu. Removes the
selected cells from the worksheet and shifts other cells to close up the
space.

**Syntax**

**EDIT.DELETE**(shift\_num)

**EDIT.DELETE**?(shift\_num)

Shift\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 4 specifying
whether to shift cells left or up after deleting the current selection
or else to delete the entire row or column.

|                |                       |
| -------------- | --------------------- |
| **Shift\_num** | **Result**            |
| 1              | Shifts cells left     |
| 2              | Shifts cells up       |
| 3              | Deletes entire row    |
| 4              | Deletes entire column |

&nbsp;

  - > If shift\_num is omitted and if one cell or a horizontal range is
    > selected, EDIT.DELETE shifts cells up.

  - > If shift\_num is omitted and a vertical range is selected,
    > EDIT.DELETE shifts cells left.


**Related Function**

[CLEAR](CLEAR.md)&nbsp;&nbsp;&nbsp;Clears specified information from the selected
cells or chart



Return to [README](README.md#E)

