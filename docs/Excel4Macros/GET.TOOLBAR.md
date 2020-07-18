GET.TOOLBAR

Returns information about one toolbar or all toolbars. Use GET.TOOLBAR
to get information about a toolbar to use with functions that add,
delete, or alter toolbars.

**Syntax**

**GET.TOOLBAR**(**type\_num**, bar\_id)

Type\_num**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;specifies what type of information to
return. If type\_num is 8 or 9, GET.TOOLBAR returns an array of names or
numbers of all visible or hidden toolbars. Otherwise, bar\_id is
required, and GET.TOOLBAR returns the requested information about the
specified toolbar.

|               |                                                                                                                                                   |
| ------------- | ------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Type\_num** | **Returns**                                                                                                                                       |
| 1             | A horizontal array of all tool IDs on the toolbar, ordered by position. Gaps are represented by zeros.                                            |
| 2             | Number indicating the horizontal position (x-coordinate) of the toolbar in the docked or floating region. For more information, see SHOW.TOOLBAR. |
| 3             | Number indicating the vertical position (y-coordinate) of the toolbar in the docked or floating region.                                           |
| 4             | Number indicating the width of the toolbar in points.                                                                                             |
| 5             | Number indicating the height of the toolbar in points.                                                                                            |
| 6             | Number indicating the toolbar location:                                                                                                           |
| &nbsp;        | 1 = Top dock in the workspace                                                                                                                     |
| &nbsp;        | 2 = Left dock in the workspace                                                                                                                    |
| &nbsp;        | 3 = Right dock in the workspace                                                                                                                   |
| &nbsp;        | 4 = Bottom dock in the workspace                                                                                                                  |
| &nbsp;        | 5 = Floating                                                                                                                                      |
| 7             | If the toolbar is visible, returns TRUE. If the toolbar is hidden, returns FALSE.                                                                 |
| 8             | An array of toolbar IDs (names or numbers in the bar\_id array) for all toolbars, visible and hidden.                                             |
| 9             | An array of toolbar IDs (names or numbers in the bar\_id array) for all visible toolbars.                                                         |
| 10            | If the toolbar is visible in full-screen mode, returns TRUE; otherwise, returns FALSE.                                                            |

Bar\_id**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;specifies the number or name of a toolbar
for which you want information. If type\_num is 8 or 9, Microsoft Excel
ignores bar\_id. For detailed information about bar\_id, see ADD.TOOL.

**Remarks**

If you request position information for a hidden toolbar, Microsoft
Excel returns the position where the toolbar would appear if shown.

**Examples**

The following macro formula returns information about the width of
Toolbar1:

GET.TOOLBAR(4, "Toolbar1")

When the following macro formula is entered as an array with
CTRL+SHIFT+ENTER, the IDs of all visible toolbars are returned, and the
array is named All\_Bar\_Ids:

SET.NAME("All\_Bar\_Ids", GET.TOOLBAR(9))

**Related Functions**

[ADD.TOOL](ADD.TOOL.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Adds one or more buttons to a toolbar

[ADD.TOOLBAR](ADD.TOOLBAR.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Creates a new toolbar with the specified
tools

[DELETE.TOOLBAR](DELETE.TOOLBAR.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Deletes custom toolbars

[GET.TOOL](GET.TOOL.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Returns information about a tool or tools on a
toolbar

[SHOW.TOOLBAR](SHOW.TOOLBAR.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Hides or displays a toolbar



Return to [README](README.md)

