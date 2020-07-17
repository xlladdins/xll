ADD.MENU

Adds a menu to a menu bar. Use ADD.MENU to add a custom menu to a
built-in or custom menu bar. You can also use ADD.MENU to restore
built-in menus you have deleted with DELETE.MENU. ADD.MENU returns the
position number in the menu bar of the new menu.

**Syntax**

**ADD.MENU**(**bar\_num, menu\_ref**, position1, position2)

Bar\_num    is the menu bar to which you want a menu added. Bar\_num can
be the ID number of a built-in or custom menu bar. See ADD.COMMAND for a
list of ID numbers for built-in menu bars.

Menu\_ref    is an array or a reference to an area on the macro sheet
that describes the new menu or the name of a deleted built-in menu you
want to restore.

  - > Menu\_ref must be made up of at least two rows and two columns of
    > cells. The upper-left cell of menu\_ref specifies the menu title,
    > which is displayed in the menu bar. In the following example, the
    > range A3:E10 is a valid menu\_ref.

![](media/image1.png)

> The rest of the first column indicates the names of the commands. The
> corresponding rows in the second column give the names of the macros
> that run when the commands are chosen.

  - > You can also specify status-bar text and Help topics in the fourth
    > and fifth columns of menu\_ref. In Microsoft Excel for the
    > Macintosh, you can specify shortcut keys in the third column of
    > menu\_ref.

>  

Position1    specifies the placement of the new menu. Position can be
the name of a menu, as text, or the number of a menu. Menus are numbered
from left to right starting with 1. Menus are added to the left of the
position specified.

  - > Use a hyphen (-) to represent a line separating commands on a
    > menu. If you want to place a command before the second separator
    > on a menu, use two hyphens (--), three hyphens for the third
    > separator, and so on.

  - > If position1 is omitted, the menu is added to the end of the menu
    > bar.

  - > If there is already a menu at position1, that menu is shifted to
    > the right and the new menu is added in its place.

  - > If you are using ADD.MENU to restore a deleted built-in menu, you
    > can use the position argument to put it back in its original place
    > on the menu bar. For example, to restore the Data menu on the
    > worksheet and macro sheet menu bar, use position 7. If position1
    > is omitted, the menu is added to the right of the last menu
    > restored.

>  

Position2    specifies the placement of a submenu.

  - > Use a hyphen (-) to represent a line separating commands on a
    > menu. If you want to place a command before the second separator
    > on a menu, use two hyphens (--), three hyphens for the third
    > separator, and so on.

  - > Position2 can be a number indicating the position of the submenu
    > on the menu. Commands are numbered from the top of the menu
    > starting with 1 and include separators.

  - > Position2 can also be the name, as text, of an existing command
    > above which you want to add the new command.

  - > If position2 is omitted, the command is added to the main menu,
    > not the submenu.

>  

**Example**

The following macro formula adds a new menu to the end of the worksheet
menu bar, where A10:B15 is the menu\_ref describing the menu:

ADD.MENU(1, A10:B15)

**Related Functions**

[ADD.BAR](ADD.BAR.md)   Adds a menu bar

[ADD.COMMAND](ADD.COMMAND.md)   Adds a command to a menu

[DELETE.MENU](DELETE.MENU.md)   Deletes a menu

[ENABLE.COMMAND](ENABLE.COMMAND.md)   Enables or disables a menu or custom command



Return to [README](README.md)

