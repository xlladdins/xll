DELETE.MENU

Deletes a menu or submenu. Use DELETE.MENU to delete menus you have
added to menu bars when the supporting macro sheet is closed (using an
Auto\_Close macro), or any time you want to remove a menu.

**Syntax**

**DELETE.MENU**(**bar\_num, menu**, submenu)

Bar\_num    is the menu bar from which you want to delete the menu.
Bar\_num can be the number of a Microsoft Excel built-in menu bar or the
number returned by a previously run ADD.BAR function. For a list of ID
numbers for built-in menu bars, see ADD.COMMAND.

Menu    is the menu you want to delete. Menu can be either the name of a
menu as text or the number of a menu. Menus are numbered starting with 1
from the left of the screen. If the specified menu does not exist,
DELETE.MENU returns the \#VALUE\! error value and interrupts the macro.
After a menu is deleted, the menu number for each menu to the right of
that menu is decreased by 1.

Submenu    is the name of the submenu you want to delete or the number
of the menu in the list of commands.

**Remarks**

You cannot delete a shortcut menu. Instead, use ENABLE.COMMAND to
prevent the user from accessing a shortcut menu.

**Example**

The following macro formula deletes the Reports menu from the custom
menu bar created by the ADD.BAR function in a cell named Financials:

DELETE.MENU(Financials, "Reports")

**Related Functions**

[ADD.MENU](ADD.MENU.md)   Adds a menu to a menu bar

[ADD.BAR](ADD.BAR.md)   Adds a menu bar

[DELETE.BAR](DELETE.BAR.md)   Deletes a menu bar

[DELETE.COMMAND](DELETE.COMMAND.md)   Deletes a command from a menu

[ENABLE.COMMAND](ENABLE.COMMAND.md)   Enables or disables a menu or custom command


