CHECK.COMMAND
=============

Adds or removes a check mark to or from a command name on a menu. A
check mark beside a command indicates that the command has been chosen.

**Syntax**

**CHECK.COMMAND**(**bar\_num**, **menu**, **command**, **check**,
position)

Bar\_num    is the menu bar containing the command. Bar\_num can be the
ID number of a built-in or custom menu bar.

Menu    is the menu containing the command. Menu can be either the name
of a menu as text or the number of a menu. Menus are numbered starting
with 1 from the left of the screen.

Command    is the command you want to check or the submenu containing
the command you want to check. Command can be the name of the command as
text or the number of the command; the first command on a menu is in
position 1.

Check    is a logical value corresponding to the check mark. If check is
TRUE, Microsoft Excel adds a check mark to the command; if FALSE,
Microsoft Excel removes the check mark.

position    is the name of a command on a submenu that you want to
check.

**Remarks**

-   The check mark doesn\'t affect execution of the command. Microsoft
    > Excel automatically adds and deletes check marks to some commands,
    > such as the name of the active workbook in the Window menu. If you
    > have assigned a check mark to a built-in command that Microsoft
    > Excel automatically changes in response to the user\'s actions,
    > the check mark will be added or removed as appropriate, and any
    > check marks you have added or deleted with CHECK.COMMAND will be
    > ignored.

-   If you use CHECK.COMMAND with a command on a Microsoft Excel version
    > 4.0 menu bar, the corresponding command on the Microsoft Excel
    > version 5.0 or later menu bar will not be effected.

>  

**Example**

The following macro formula adds a check mark to the Sales command on
the Weekly menu on a custom menu bar created by the ADD.BAR function in
a cell named Reports:

CHECK.COMMAND(Reports, \"Weekly\", \"Sales\", TRUE)

**Related Functions**

ADD.COMMAND   Adds a command to a menu

DELETE.COMMAND   Deletes a command from a menu

ENABLE.COMMAND   Enables or disables a menu or custom command

RENAME.COMMAND   Changes the name of a command or menu

Return to [top](#A)

CLEAR
