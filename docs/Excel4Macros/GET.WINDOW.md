GET.WINDOW

Returns information about a window. Use GET.WINDOW in a macro that
requires the status of a window, such as its name, size, position, and
display options.

**Syntax**

**GET.WINDOW**(**type\_num**, window\_text)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number that specifies what type of
window information you want. The following list shows the possible
values of type\_num and the corresponding results:

|               |                                                                                                                                                                                                                                                                                                                               |
| ------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Type\_num** | **Returns**                                                                                                                                                                                                                                                                                                                   |
| 1             | Name of the workbook and sheet in the window as text. For compatibility with Microsoft Excel version 4.0, if the window contains only a single sheet that has the same name as the workbook without its extension, returns only the name of the book. Otherwise, returns the name of the sheet in the form "\[Book1\]Sheet1". |
| 2             | Number of the window.                                                                                                                                                                                                                                                                                                         |
| 3             | X position, measured in points from the left edge of the workspace (in Microsoft Excel for Windows) or screen (in Microsoft Excel for the Macintosh) to the left edge of the window.                                                                                                                                          |
| 4             | Y position, measured in points from the bottom edge of the formula bar to the top edge of the window.                                                                                                                                                                                                                         |
| 5             | Width, measured in points.                                                                                                                                                                                                                                                                                                    |
| 6             | Height, measured in points.                                                                                                                                                                                                                                                                                                   |
| 7             | If window is hidden, returns TRUE; otherwise, returns FALSE.                                                                                                                                                                                                                                                                  |

The rest of the values for type\_num apply only to worksheets and macro
sheets, except where indicated:

|               |                                                                                                                                                                       |
| ------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Type\_num** | **Returns**                                                                                                                                                           |
| 8             | If formulas are displayed, returns TRUE; otherwise, returns FALSE.                                                                                                    |
| 9             | If gridlines are displayed, returns TRUE; otherwise, returns FALSE.                                                                                                   |
| 10            | If row and column headings are displayed, returns TRUE; otherwise, returns FALSE.                                                                                     |
| 11            | If zeros are displayed, returns TRUE; otherwise, returns FALSE.                                                                                                       |
| 12            | Gridline and heading color as a number in the range 1 to 56, corresponding to the colors in the View tab of the Options dialog box; if color is automatic, returns 0. |

Values 13 to 16 for type\_num return arrays that specify which rows or
columns are at the top and left edges of the panes in the window and the
widths and heights of those panes.

  - > The first number in the array corresponds to the first pane, the
    > second number to the second pane, and so on.

  - > If the edge of the pane occurs at the boundary between rows or
    > columns, the number returned is an integer.

  - > If the edge of the pane occurs within a row or column, the number
    > returned has a fractional part that represents the fraction of the
    > row or column visible within the pane.

  - > The numbers can be used as arguments to the SPLIT function to
    > split a window at specific locations.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>13</td>
<td>Leftmost column number of each pane, in a horizontal numeric array</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Top row number of each pane, in a horizontal numeric array.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Number of columns in each pane, in a horizontal numeric array.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Number of rows in each pane, in a horizontal numeric array.</td>
</tr>
<tr class="even">
<td>17</td>
<td><p>Number indicating the active pane:</p>
<p>1 = Upper, left, or upper-left</p>
<p>2 = Right or upper-right</p>
<p>3 = Lower or lower-left</p>
<p>4 = Lower-right</p></td>
</tr>
<tr class="odd">
<td>18</td>
<td>If window has a vertical split, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>19</td>
<td>If window has a horizontal split, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>If window is maximized, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>21</td>
<td>Reserved</td>
</tr>
<tr class="odd">
<td>22</td>
<td>If the Outline Symbols check box is selected in the View tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>23</td>
<td><p>Number indicating the size of the window (including charts):</p>
<p>1 = Restored</p>
<p>2 = Minimized (displayed as an icon)</p>
<p>3 = Maximized</p></td>
</tr>
<tr class="odd">
<td>24</td>
<td>If panes are frozen on the active window, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>25</td>
<td>The numeric magnification of the active window (as a percentage of normal size) as set in the Zoom dialog box, or 100 if none is specified.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>Returns TRUE if horizontal scrollbars are displayed in the active window; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>27</td>
<td>Returns TRUE if vertical scrollbars are displayed in the active window; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>28</td>
<td>Returns the tab ratio of workbook tabs to horizontal scrollbar, from 0 to 1. The default is .6.</td>
</tr>
<tr class="even">
<td>29</td>
<td>Returns TRUE if workbook tabs are displayed in the active window; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>30</td>
<td>Returns the title of the active sheet in the window in the form "[Book1]Sheet1".</td>
</tr>
<tr class="even">
<td>31</td>
<td>Returns the name of a workbook only, without read/write indicated. For example, if Book1.xls is read only, then "Book.xls" will be returned without "[Read Only]" appended.</td>
</tr>
</tbody>
</table>

Window\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name that appears in the
title bar of the window that you want information about. If window\_text
is omitted, it is assumed to be the active window.

**Examples**

If the active window contains the workbook Book1, then:

GET.WINDOW(1) equals "Book1"

If the title of the active window is Macro1:3, then:

GET.WINDOW(2) equals 3

In Microsoft Excel for Windows, the following macro formula returns the
gridline and heading color of REPORT.XLS:

GET.WINDOW(12, "REPORT.XLS")

In Microsoft Excel for the Macintosh, the following macro formula
returns the gridline and heading color of REPORT MASTER:

GET.WINDOW(12, "REPORT MASTER")

**Related Functions**

[GET.DOCUMENT](GET.DOCUMENT.md)&nbsp;&nbsp;&nbsp;Returns information about a workbook

[GET.WORKSPACE](GET.WORKSPACE.md)&nbsp;&nbsp;&nbsp;Returns information about the workspace



Return to [README](README.md)

