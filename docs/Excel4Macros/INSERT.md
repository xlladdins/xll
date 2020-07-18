INSERT

Inserts a blank cell or range of cells or pastes cells from the
Clipboard into a sheet. Shifts the selected cells to accommodate the new
ones. The size and shape of the inserted range are the same as those of
the current selection.

**Syntax**

**INSERT**(shift\_num)

**INSERT**?(shift\_num)

Shift\_num**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a number from 1 to 4 specifying
which way to shift the cells. If an entire row or column is selected,
shift\_num is ignored. If shift\_num is omitted, Microsoft Excel shifts
cells in the logical direction based on the selection.

|                |                     |
| -------------- | ------------------- |
| **Shift\_num** | **Direction**       |
| 1              | Shift cells right   |
| 2              | Shift cells down    |
| 3              | Shift entire row    |
| 4              | Shift entire column |

**Remarks**

If you have just cut or copied information to the Clipboard, INSERT
performs both an insert and a paste operation. First, Microsoft Excel
inserts new blank cells into the sheet; then, Microsoft Excel pastes
information from the Clipboard into the newly inserted cells. If you
have used the INSERT function in macros written for Microsoft Excel
version 2.2 or earlier, make sure you consider this feature when you use
your old macros with later versions of Microsoft Excel.

**Related Functions**

[COPY](COPY.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Copies and pastes data or objects

[CUT](CUT.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Cuts or moves data or objects

[EDIT.DELETE](EDIT.DELETE.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Removes cells from a sheet

[PASTE](PASTE.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Pastes cut or copied data



Return to [README](README.md)

