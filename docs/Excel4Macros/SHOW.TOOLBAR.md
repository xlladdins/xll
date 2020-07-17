SHOW.TOOLBAR
============

Equivalent to selecting the check box corresponding to a toolbar on the
Toolbars tab in the Customize dialog box, which appears when you select
the Customize command (View menu, Toolbars submenu). Hides or displays a
toolbar. Use SHOW.TOOLBAR to display or hide a menu bar you have created
with the ADD.BAR function or to display a built-in Microsoft Excel 95 or
earlier version toolbar.

**Syntax**

**SHOW.TOOLBAR**(**bar\_id, visible**, dock, x\_pos, y\_pos, width,
protect, tool\_tips, large\_buttons, color\_buttons)

Bar\_id    is a number or name of a toolbar corresponding to the
toolbars you want to display. For detailed information about bar\_id,
see ADD.TOOL.

Visible    is a logical value that, if TRUE, specifies that the toolbar
is visible or, if FALSE, specifies that the toolbar is hidden.

Dock    specifies the docking location of the toolbar.

+------------+---------------------------+
| > **Dock** | > **Position of toolbar** |
+------------+---------------------------+
| > 1        | > Top of workspace        |
+------------+---------------------------+
| > 2        | > Left edge of workspace  |
+------------+---------------------------+
| > 3        | > Right edge of workspace |
+------------+---------------------------+
| > 4        | > Bottom of workspace     |
+------------+---------------------------+
| > 5        | > Floating (not docked)   |
+------------+---------------------------+

X\_pos    specifies the horizontal position of the toolbar.

-   If the toolbar is docked (not floating), x\_pos is measured
    > horizontally from the left edge of the toolbar to the left edge of
    > the toolbar\'s docking area.

-   If the toolbar is floating, x\_pos is measured horizontally from the
    > left edge of the toolbar to the right edge of the rightmost
    > toolbar in the left docking area.

-   X\_pos is measured in points. A point is 1/72nd of an inch.

Y\_pos    specifies the vertical position of the toolbar.

-   If the toolbar is docked, y\_pos is measured vertically from the top
    > edge of the toolbar to the top edge of the toolbar\'s docking
    > area.

-   If the toolbar is floating, y\_pos is measured vertically from the
    > top edge of the toolbar to the top edge of the Microsoft Excel
    > workspace.

-   Y\_pos is measured in points.

>  

Width    specifies the width of the toolbar and is measured in points.
If you omit width, Microsoft Excel uses the existing width setting.

Protect    is a number specifying the degree to which you can modify a
toolbar and its buttons. Each succeeding protect number retains the
protection status of its previous numbers. For example, a protect status
of 3 (a toolbar cannot become docked if it is floating) assumes the
protection status of 0, 1, and 2 as well.

+---------------+-----------------------------------------------------+
| > **Protect** | > **Description**                                   |
+---------------+-----------------------------------------------------+
| > 0           | > Default. Toolbars can be re-shaped, docked, and   |
|               | > floating. Toolbar buttons can be removed from and |
|               | > moved to the toolbar.                             |
+---------------+-----------------------------------------------------+
| > 1           | > Toolbars can be re-shaped, docked, and floating.  |
|               | > Toolbar buttons can not be removed from nor moved |
|               | > to the toolbar.                                   |
+---------------+-----------------------------------------------------+
| > 2           | > A floating toolbar cannot be re-shaped. It can be |
|               | > docked.                                           |
+---------------+-----------------------------------------------------+
| > 3           | > A floating toolbar cannot be docked. If it is     |
|               | > already docked, it cannot become floating.        |
+---------------+-----------------------------------------------------+
| > 4           | > The toolbar cannot be moved at all. If it is      |
|               | > already floating, it cannot be re-shaped or       |
|               | > moved. If it is docked, it cannot become          |
|               | > un-docked.                                        |
+---------------+-----------------------------------------------------+

Tool\_tips    is a logical value that corresponds to the Show Screentips
On Toolbars check box on the Options tab. If TRUE, ScreenTips will be
displayed. If FALSE, ScreenTips will not be displayed.

Large Buttons    is a logical value that corresponds to the Large Icons
check box on the Options tab. If TRUE, large icons will be displayed. If
FALSE, large icons will not be displayed.

Color\_buttons    is a logical value that corresponds to the Color
Toolbars check box. If TRUE, the toolbar buttons will be displayed in
color. If FALSE, the toolbar buttons will not be displayed in color.
This argument is for compatibility with Microsoft Excel version 5.0.

**Related Functions**

ADD.BAR   Adds a menu bar

ADD.TOOLBAR   Creates a new toolbar with the specified tools

Return to [top](#Q)

SIZE
