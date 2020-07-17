RENAME.COMMAND

Changes the name of a built-in or custom menu command or the name of a
menu. Use RENAME.COMMAND to change the name of a command on a menu, for
example, when you create two custom commands that toggle on the menu.
Examples of two built-in commands that toggle are the Page Break and
Remove Page Break commands on the Insert menu.

**Syntax**

**RENAME.COMMAND**(**bar\_num, menu, command, name\_text**, position)

Bar\_num    can be either the number of one of the Microsoft Excel
built-in menu bars or the number returned by a previously run ADD.BAR
function. See ADD.COMMAND for a list of ID numbers for built-in menu
bars.

Menu    can be either the name of a menu as text or the number of a
menu. Menus are numbered starting with 1 from the left of the screen.

Command    can be either the name of the command as text or the number
of the command to be renamed (the first command on a menu is command 1).
If command is 0, RENAME.COMMAND renames the menu instead of the command.
Because other macros can change the position of custom menu commands,
you should use the name of the command rather than a number whenever
possible.

If the specified menu bar, menu, or command does not exist,
RENAME.COMMAND returns the \#VALUE\! error value and interrupts the
macro.

Name\_text    is the new name for the command.

Position    is the name of the command on a submenu that you want to
rename. If you use position, you must use command as the name of the
submenu.

**Tip**   To specify an access key for the new name, precede the
character you want to use with an ampersand (&). The access key is
indicated by an underline under one letter of a menu or command name. In
Microsoft Excel for the Macintosh, you can use the General tab in the
Options dialog box to turn command underlining on or off. To see the
Options dialog box, click Options on the Tools menu.

**Example**

To rename the Save All command as Global Save, and to make the letter
"G" in Global Save an access key, use the following macro formula:

RENAME.COMMAND(10, "File", "Save All", "\&Global Save")

**Related Functions**

[ADD.BAR](ADD.BAR.md)   Adds a menu bar

[ADD.COMMAND](ADD.COMMAND.md)   Adds a command to a menu

[CHECK.COMMAND](CHECK.COMMAND.md)   Adds or deletes a check mark to or from a command

[DELETE.COMMAND](DELETE.COMMAND.md)   Deletes a command from a menu

[ENABLE.COMMAND](ENABLE.COMMAND.md)   Enables or disables a menu or custom command


