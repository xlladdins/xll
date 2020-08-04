# ADD.COMMAND

Adds a command to a menu. ADD.COMMAND returns the position number on the
menu of the added command. Use ADD.COMMAND to add one or more custom
menu commands to a menu on a built-in or custom menu bar. You can also
use ADD.COMMAND to restore a deleted built-in command to its original
menu.

**Syntax**

**ADD.COMMAND**(**bar\_num, menu, command\_ref**, position1, position2)

Bar\_num&nbsp;&nbsp;&nbsp;&nbsp;is the number corresponding to a menu
bar or a type of shortcut menu to which you want to add a command.

  - > Bar\_num can be the ID number of a built-in or custom menu bar.
    > The ID number of a custom menu bar is the number returned by the
    > ADD.BAR function.

  - > Bar\_num can also refer to a type of shortcut menu; use menu to
    > identify the specific shortcut menu.


The ID numbers of the built-in menu bars and the types of shortcut menus
are listed in the following tables. Short menus are abbreviated versions
of the normal Microsoft Excel menus. To turn on short menus, use the
SHORT.MENUS function.

|              |                                                                          |
| ------------ | ------------------------------------------------------------------------ |
| **Bar\_num** | **Built-in menu bar**                                                    |
| 1            | Worksheet and macro sheet (Microsoft Excel 4.0 or later)                 |
| 2            | Chart (Microsoft Excel 4.0 or later)                                     |
| 3            | Null (the menu displayed when no workbooks are open)                     |
| 4            | Info                                                                     |
| 5            | Worksheet and macro sheet (short menus, Microsoft Excel 3.0 and earlier) |
| 6            | Chart (short menus, Microsoft Excel 3.0 and earlier)                     |
| 7            | Cell, toolbar, and workbook (shortcut menus)                             |
| 8            | Object (shortcut menus)                                                  |
| 9            | Chart (Microsoft Excel 4.0 or later shortcut menus)                      |
| 10           | Worksheet and macro sheet                                                |
| 11           | Chart                                                                    |
| 12           | Visual Basic                                                             |

Menu&nbsp;&nbsp;&nbsp;&nbsp;is the menu to which you want the new
command added.

  - > Menu can be either the name of a menu as text or the number of a
    > menu.

  - > If bar\_num is 1 through 6, menus are numbered starting with 1
    > from the left of the menu bar.

  - > If bar\_num is 7, 8, or 9, menu refers to a built-in shortcut
    > menu. The combination of bar\_num and menu determines which
    > shortcut menu to modify, as shown in the following table.


|              |          |                                                                    |
| ------------ | -------- | ------------------------------------------------------------------ |
| **Bar\_num** | **Menu** | **Shortcut menu**                                                  |
| 7            | 1        | Toolbars                                                           |
| 7            | 2        | Toolbar buttons                                                    |
| 7            | 3        | Workbook paging icons in Microsoft Excel 4.0                       |
| 7            | 4        | Cells (worksheet)                                                  |
| 7            | 5        | Column selections                                                  |
| 7            | 6        | Row selections                                                     |
| 7            | 7        | Workbook tabs                                                      |
| 7            | 8        | Cells (macro sheet)                                                |
| 7            | 9        | Workbook title bar                                                 |
| 7            | 10       | Desktop (Microsoft Excel for Windows only)                         |
| 7            | 11       | Module                                                             |
| 7            | 12       | Watch pane                                                         |
| 7            | 13       | Immediate pane                                                     |
| 7            | 14       | Debug code pane                                                    |
| 8            | 1        | Drawn or imported objects on worksheets, dialog sheets, and charts |
| 8            | 2        | Buttons on sheets                                                  |
| 8            | 3        | Text boxes                                                         |
| 8            | 4        | Dialog sheet                                                       |
| 9            | 1        | Chart series                                                       |
| 9            | 2        | Chart and axis titles                                              |
| 9            | 3        | Chart plot area and walls                                          |
| 9            | 4        | Entire chart                                                       |
| 9            | 5        | Chart axes                                                         |
| 9            | 6        | Chart gridlines                                                    |
| 9            | 7        | Chart floor and arrows                                             |
| 9            | 8        | Chart legend                                                       |

**Note**&nbsp;&nbsp;&nbsp;Any commands that you add to the toolbar
buttons, watch pane, immediate pane or debug code pane shortcut menus
will be dimmed.

Command\_ref&nbsp;&nbsp;&nbsp;&nbsp;is an array or a reference to an
area on the macro sheet that describes the new command or commands.

  - > Command\_ref must be at least two columns wide. The first column
    > specifies command names; the second specifies macro names.
    > Optional columns can be specified for shortcut keys (in Microsoft
    > Excel for the Macintosh), status bar messages, and custom Help
    > topics, in that order.

  - > Command\_ref is similar to menu\_ref in ADD.MENU. For more
    > information about command\_ref, see the description of menu\_ref
    > in ADD.MENU.

  - > Command\_ref can be the name, as text, of a previously deleted
    > built-in command that you want to restore. You can also use the
    > value returned by the DELETE.COMMAND formula that deleted the
    > command.


Position1&nbsp;&nbsp;&nbsp;&nbsp;specifies the placement of the new
command.

  - > Use a hyphen (-) to represent a line separating commands on a
    > menu. If you want to place a command before the second separator
    > on a menu, use two hyphens (--), three hyphens for the third
    > separator, and so on.

  - > Position1 can be a number indicating the position of the command
    > on the menu. Commands are numbered from the top of the menu
    > starting with 1.

  - > Position1 can be the name of an existing command, as text, above
    > which you want to add the new command.

  - > If position1 is omitted, the command is added to the bottom of the
    > menu.

  - > For the toolbar shortcut menu (bar\_num 7, menu 1) and the
    > shortcut menu for workbook paging icons in Microsoft Excel version
    > 4.0 (bar\_num 7, menu 3), you cannot add commands to the middle of
    > the toolbar name list or the middle of the workbook contents list.


Position2&nbsp;&nbsp;&nbsp;&nbsp;specifies the placement of the new
command on a submenu.

  - > Position2 can be a number indicating the position of the command
    > on the submenu. Commands are numbered from the top of the menu
    > starting with 1.

  - > Position2 can be the name of an existing command, as text, above
    > which you want to add the new command.

  - > If position2 is omitted, the command is added to the main menu,
    > not the submenu.

  - > To add a command to the bottom of a submenu, use 0 for position2.

**Tip**&nbsp;&nbsp;&nbsp;In general, use menu and command names rather
than numbers for arguments. The numbers assigned to menus and commands
change as you add and delete menus and commands. Using names ensures
that your menu and command macro functions always refer to the correct
items.

**Example**

The following macro formula adds the command described in cells G16:J16
to the bottom of the worksheet cells shortcut menu:

ADD.COMMAND(7, 4, G16:J16)

**Related Functions**

[ADD.BAR](ADD.BAR.md)&nbsp;&nbsp;&nbsp;Adds a menu bar

[ADD.MENU](ADD.MENU.md)&nbsp;&nbsp;&nbsp;Adds a menu to a menu bar

[ADD.TOOL](ADD.TOOL.md)&nbsp;&nbsp;&nbsp;Adds one or more buttons to a toolbar

[ADD.TOOLBAR](ADD.TOOLBAR.md)&nbsp;&nbsp;&nbsp;Creates a toolbar with the specified tools

[DELETE.COMMAND](DELETE.COMMAND.md)&nbsp;&nbsp;&nbsp;Deletes a command from a menu

[ENABLE.COMMAND](ENABLE.COMMAND.md)&nbsp;&nbsp;&nbsp;Enables or disables a menu or custom
command

[GET.TOOLBAR](GET.TOOLBAR.md)&nbsp;&nbsp;&nbsp;Retrieves information about a toolbar

[RENAME.COMMAND](RENAME.COMMAND.md)&nbsp;&nbsp;&nbsp;Changes the name of a command or menu



Return to [README](README.md#A)

