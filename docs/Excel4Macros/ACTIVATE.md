ACTIVATE

Switches to a window if more than one window is open, or a pane of a
window if the window is split and its panes are not frozen. Switching to
a pane is useful with functions such as VSCROLL, HSCROLL, and GOTO,
which operate only on the active pane.

**Syntax**

**ACTIVATE**(window\_text, pane\_num)

**ACTIVATE**?(window\_text, pane\_num)

Window\_text    is text specifying the name of a window to switch to:
for example, "Book1" or "Book1:2".

  - > If a workbook is displayed in more than one window and
    > window\_text does not specify which window to switch to, the first
    > window containing that workbook is switched to.

  - > If window\_text is omitted, the active window is not changed.

>  

Pane\_num    is a number from 1 to 4 specifying which pane to switch to.
If pane\_num is omitted and the window has more than one pane, the
active pane is not changed.

|               |                                                                                                                                                                                                                          |
| ------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| **Pane\_num** | **Activates**                                                                                                                                                                                                            |
| 1             | Upper-left pane of the active sheet. If the window is not split, this is the only pane. If the window is split only horizontally, this is the upper pane. If the window is split only vertically, this is the left pane. |
| 2             | Upper-right pane of the active sheet. If the window is split only vertically, this is the right pane. If the window is split only horizontally, an error occurs.                                                         |
| 3             | Lower-left pane of the active sheet. If the window is split only horizontally, this is the lower pane. If the window is split only vertically, an error occurs.                                                          |
| 4             | Lower-right pane of the active sheet. If the window is split into only two panes either vertically or horizontally, an error occurs.                                                                                     |

**Related Functions**

[ACTIVATE.NEXT](ACTIVATE.NEXT.md)   Switches to the next window, or switches to the next
sheet in a workbook

[ACTIVATE.PREV](ACTIVATE.PREV.md)   Switches to the previous window, or switches to the
previous sheet in a workbook

[DOCUMENTS](DOCUMENTS.md)   Returns the names of the specified open workbooks

[FREEZE.PANES](FREEZE.PANES.md)   Freezes the panes of a window so that they do not scroll

[ON.WINDOW](ON.WINDOW.md)   Runs a macro when you switch to a window

[SPLIT](SPLIT.md)   Splits a window

[WINDOWS](WINDOWS.md)   Returns the names of all open windows

[WORKBOOK.SELECT](WORKBOOK.SELECT.md)   Select a sheet in a workbook



Return to [README](README.md)

