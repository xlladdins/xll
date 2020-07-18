# ARRANGE.ALL

Equivalent to clicking the Arrange command on the Window menu.
Rearranges open windows and icons and resizes open windows. Also can be
used to synchronize scrolling of windows of the active sheet.

**Syntax**

**ARRANGE.ALL**(arrange\_num, active\_doc, sync\_horiz, sync\_vert)

**ARRANGE.ALL**?(arrange\_num, active\_doc, sync\_horiz, sync\_vert)

Arrange\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 7 specifying
how to arrange the windows.

|                  |                                                                                                             |
| ---------------- | ----------------------------------------------------------------------------------------------------------- |
| **Arrange\_num** | **Result**                                                                                                  |
| 1 or omitted     | Tiled (also used to arrange icons in Microsoft Excel for Windows)                                           |
| 2                | Horizontal                                                                                                  |
| 3                | Vertical                                                                                                    |
| 4                | None                                                                                                        |
| 5                | Horizontally arranges and sizes the windows based on the position of the active cell.                       |
| 6                | Vertically arranges and sizes the windows based on the position of the active cell.                         |
| 7                | Arranges windows so that they cascade from the upper left to the bottom right of the application workspace. |

If you want to change whether the windows are synchronized for scrolling
but not how they are arranged, make sure arrange\_num is 4.

Active\_doc&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying which
windows to arrange. If active\_doc is TRUE, Microsoft Excel arranges
only windows on the active workbook; if FALSE or omitted, all open
windows are arranged.

Sync\_horiz&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to
the Sync Horizontal check box in Microsoft Excel version 4.0.

  - > If sync\_horiz is TRUE, Microsoft Excel selects the check box and
    > synchronizes horizontal scrolling.

  - > If sync\_horiz is FALSE or omitted, Microsoft Excel clears the
    > check box, and windows will not be synchronized when you scroll
    > horizontally.

  - > This argument is used only when active\_doc is TRUE.


Sync\_vert&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to
the Sync Vertical check box in Microsoft Excel version 4.0.

  - > If sync\_vert is TRUE, Microsoft Excel selects the check box and
    > synchronizes vertical scrolling.

  - > If sync\_vert is FALSE or omitted, Microsoft Excel clears the
    > check box, and windows will not be synchronized when you scroll
    > vertically.

  - > This argument is used only when active\_doc is TRUE.


**Note**&nbsp;&nbsp;&nbsp;If arguments are omitted in the dialog box
form of this function, the default values are the previous settings, if
any; otherwise the default values are as described above.

**Remarks**

  - > After arranging windows, the top or leftmost window is active.

  - > In Microsoft Excel for Windows, if all windows are minimized,
    > ARRANGE.ALL ignores its arguments, if any, and arranges the
    > corresponding icons horizontally along the bottom of the
    > workspace.


**Tip**&nbsp;&nbsp;&nbsp;You can use synchronized horizontal or vertical
scrolling when you need to scroll while viewing macro formulas in one
window and corresponding macro values in another window of the same
macro sheet.

**Related Function**

[ACTIVATE](ACTIVATE.md)&nbsp;&nbsp;&nbsp;Switches to a window



Return to [README](README.md)

