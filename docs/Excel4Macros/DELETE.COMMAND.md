DELETE.COMMAND

Deletes a command from a custom or built-in menu. Use DELETE.COMMAND to
remove commands you don't want the user to have access to or to remove
custom commands that you have added.

**Syntax**

**DELETE.COMMAND**(**bar\_num, menu, command**, subcommand)

Bar\_num    is the menu bar from which you want to delete the command.
Bar\_num can be the ID number of a built-in or custom menu bar. See
ADD.COMMAND for a list of ID numbers for built-in menu bars and shortcut
menus.

Menu    is the menu from which you want to delete the command. Menu can
be the name of a menu as text or the number of a menu. Menus are
numbered starting with 1 from the left of the screen.

Command    is the command you want to delete, or the name of a submenu.
Command can be the name of the command as text or the number of the
command; the first command on a menu is in position 1.

Subcommand    is the command you want to delete from a submenu. If you
use subcommand, you must use command as the name of the submenu.

**Remarks**

  - > If the specified command does not exist, DELETE.COMMAND returns
    > the \#VALUE\! error value and interrupts the macro.

  - > After a command is deleted, the command number for all commands
    > below that command is decreased by one.

  - > When you delete a built-in command, DELETE.COMMAND returns a
    > unique ID number for that command. You can use this ID number with
    > ADD.COMMAND to restore the built-in command to the original menu.

>  

**Example**

The following macro formula removes the Compile Reports command from the
Reports menu on a custom menu bar created by the ADD.BAR function in a
cell named Financials.

DELETE.COMMAND(Financials, "Reports", "Compile Reports...")

**Related Functions**

[ADD.COMMAND](ADD.COMMAND.md)   Adds a command to a menu

[CHECK.COMMAND](CHECK.COMMAND.md)   Adds or deletes a check mark to or from a command

[ENABLE.COMMAND](ENABLE.COMMAND.md)   Enables or disables a menu or custom command

[RENAME.COMMAND](RENAME.COMMAND.md)   Changes the name of a command or menu


