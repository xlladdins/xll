SHOW.BAR
========

Displays the specified menu bar. Use SHOW.BAR to display a menu bar you
have created with the ADD.BAR function or to display a built-in
Microsoft Excel 95 or earlier version menu bar.

**Syntax**

**SHOW.BAR**(bar\_num)

Bar\_num    is the number of the menu bar you want to display. It can be
the number of one of the Microsoft Excel built-in menu bars, the number
returned by a previously executed ADD.BAR function, or a reference to a
cell containing a previously executed ADD.BAR function.

If bar\_num is omitted, Microsoft Excel displays the appropriate menu
bar for the active workbook as shown in the following table.

+----------------+----------------------------------------------------+
| > **Bar\_num** | > **Bar displayed**                                |
+----------------+----------------------------------------------------+
| > 1            | > A sheet or macro sheet (Microsoft Excel version  |
|                | > 4.0)                                             |
+----------------+----------------------------------------------------+
| > 2            | > A chart (Microsoft Excel version 4.0)            |
+----------------+----------------------------------------------------+
| > 3            | > No active window                                 |
+----------------+----------------------------------------------------+
| > 4            | > The Info window (Microsoft Excel 95 or earlier   |
|                | > versions)                                        |
+----------------+----------------------------------------------------+
| > 5            | > A sheet or macro sheet (short menus)             |
+----------------+----------------------------------------------------+
| > 6            | > A chart (short menus)                            |
+----------------+----------------------------------------------------+
| > 7            | > Shortcut menus 1 (for Cells, Workbook tabs,      |
|                | > Toolbars, VB Windows)                            |
+----------------+----------------------------------------------------+
| > 8            | > Shortcut menus 2 (for objects)                   |
+----------------+----------------------------------------------------+
| > 9            | > Shortcut menus 3 (for chart elements)            |
+----------------+----------------------------------------------------+
| > 10           | > A sheet or macro sheet                           |
+----------------+----------------------------------------------------+
| > 11           | > A chart                                          |
+----------------+----------------------------------------------------+
| > 12           | > A Visual Basic module                            |
+----------------+----------------------------------------------------+
| > 13-35        | > Reserved for use by shortcut menus. These        |
|                | > numbers will return an error if a macro tries to |
|                | > do anything with them.                           |
+----------------+----------------------------------------------------+
| > 37-51        | > Custom menu bar for macro use                    |
+----------------+----------------------------------------------------+

**Remarks**

-   When displaying a built-in menu bar, you can display only bars 1 or
    > 5 if a sheet or macro sheet is active, bars 2 or 6 if a chart is
    > active, and so on. If you try to display a chart menu bar while a
    > sheet or macro sheet is active, SHOW.BAR returns an error and
    > interrupts the current macro.

-   Displaying a custom menu bar disables automatic menu-bar switching
    > when different types of sheets are selected. For example, if a
    > custom menu bar is displayed and you switch to a chart, neither of
    > the two chart menus is automatically displayed as it would be when
    > you are using the built-in menu bars. Automatic menu-bar switching
    > is reenabled when a built-in bar is displayed using SHOW.BAR.

>  

**Example**

The following macro formula displays short menus on a worksheet or macro
sheet:

SHOW.BAR(5)

**Related Functions**

ADD.BAR   Adds a menu bar

DELETE.BAR   Deletes a menu bar

SHOW.TOOLBAR   Hides or displays a toolbar

Return to [top](#Q)

SHOW.CLIPBOARD
