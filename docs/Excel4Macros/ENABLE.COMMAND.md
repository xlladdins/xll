# ENABLE.COMMAND

Enables or disables a custom command or menu. Disabled commands appear
dimmed and can't be chosen. Use ENABLE.COMMAND to control which commands
the user can click in a menu bar.

**Syntax**

**ENABLE.COMMAND**(**bar\_num**, **menu**, **command**, **enable**,
subcommand)

Bar\_num&nbsp;&nbsp;&nbsp;&nbsp;is the menu bar in which a command
resides. Bar\_num can be the number of a built-in menu bar or the number
returned by a previously run ADD.BAR function. See ADD.COMMAND for a
list of the built-in menu bar numbers.

Menu&nbsp;&nbsp;&nbsp;&nbsp;is the menu on which the command resides.
Menu can be either the name of a menu as text or the number of a menu.
Menus are numbered starting with 1 from the left of the screen.

Command&nbsp;&nbsp;&nbsp;&nbsp;is the command you want to enable or
disable. Command can be either the name of the command as text or the
number of the command. The top command on a menu is command 1. If
command is 0, ENABLE.COMMAND enables or disables the entire menu.

Enable&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying whether the
command should be enabled or disabled. If enable is TRUE, Microsoft
Excel enables the command; if FALSE, it disables the command.

Subcommand&nbsp;&nbsp;&nbsp;&nbsp;is the name of the command on a
submenu that you want to enable. If you use subcommand, you must use
command as the name of the submenu. Use subcommand 0 to enable an entire
submenu.

**Remarks**

  - > You cannot disable built-in commands. If the specified command is
    > a built-in command or does not exist, ENABLE.COMMAND returns the
    > \#VALUE\! error value and interrupts the macro.

  - > You can hide any shortcut menu from users by using ENABLE.COMMAND
    > with command set to 0.


**Example**

The following macro formula disables a custom command that had been
added previously to the View menu on the worksheet and macro sheet menu
bar:

ENABLE.COMMAND(10, "View", "Audit...", FALSE)

**Related Functions**

[ADD.BAR](ADD.BAR.md)&nbsp;&nbsp;&nbsp;Adds a menu bar

[ADD.COMMAND](ADD.COMMAND.md)&nbsp;&nbsp;&nbsp;Adds a command to a menu

[CHECK.COMMAND](CHECK.COMMAND.md)&nbsp;&nbsp;&nbsp;Adds or deletes a check mark to or from a
command

[DELETE.COMMAND](DELETE.COMMAND.md)&nbsp;&nbsp;&nbsp;Deletes a command from a menu

[RENAME.COMMAND](RENAME.COMMAND.md)&nbsp;&nbsp;&nbsp;Changes the name of a command or menu



Return to [README](README.md#E)

# ENABLE.COMMAND

Enables or disables a custom command or menu. Disabled commands appear
dimmed and can't be chosen. Use ENABLE.COMMAND to control which commands
the user can click in a menu bar.

**Syntax**

**ENABLE.COMMAND**(**bar\_num**, **menu**, **command**, **enable**,
subcommand)

Bar\_num&nbsp;&nbsp;&nbsp;&nbsp;is the menu bar in which a command
resides. Bar\_num can be the number of a built-in menu bar or the number
returned by a previously run ADD.BAR function. See ADD.COMMAND for a
list of the built-in menu bar numbers.

Menu&nbsp;&nbsp;&nbsp;&nbsp;is the menu on which the command resides.
Menu can be either the name of a menu as text or the number of a menu.
Menus are numbered starting with 1 from the left of the screen.

Command&nbsp;&nbsp;&nbsp;&nbsp;is the command you want to enable or
disable. Command can be either the name of the command as text or the
number of the command. The top command on a menu is command 1. If
command is 0, ENABLE.COMMAND enables or disables the entire menu.

Enable&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying whether the
command should be enabled or disabled. If enable is TRUE, Microsoft
Excel enables the command; if FALSE, it disables the command.

Subcommand&nbsp;&nbsp;&nbsp;&nbsp;is the name of the command on a
submenu that you want to enable. If you use subcommand, you must use
command as the name of the submenu. Use subcommand 0 to enable an entire
submenu.

**Remarks**

  - > You cannot disable built-in commands. If the specified command is
    > a built-in command or does not exist, ENABLE.COMMAND returns the
    > \#VALUE\! error value and interrupts the macro.

  - > You can hide any shortcut menu from users by using ENABLE.COMMAND
    > with command set to 0.


**Example**

The following macro formula disables a custom command that had been
added previously to the View menu on the worksheet and macro sheet menu
bar:

ENABLE.COMMAND(10, "View", "Audit...", FALSE)

**Related Functions**

[ADD.BAR](ADD.BAR.md)&nbsp;&nbsp;&nbsp;Adds a menu bar

[ADD.COMMAND](ADD.COMMAND.md)&nbsp;&nbsp;&nbsp;Adds a command to a menu

[CHECK.COMMAND](CHECK.COMMAND.md)&nbsp;&nbsp;&nbsp;Adds or deletes a check mark to or from a
command

[DELETE.COMMAND](DELETE.COMMAND.md)&nbsp;&nbsp;&nbsp;Deletes a command from a menu

[RENAME.COMMAND](RENAME.COMMAND.md)&nbsp;&nbsp;&nbsp;Changes the name of a command or menu



Return to [README](README.md#E)

# ENABLE.COMMAND

Enables or disables a custom command or menu. Disabled commands appear
dimmed and can't be chosen. Use ENABLE.COMMAND to control which commands
the user can click in a menu bar.

**Syntax**

**ENABLE.COMMAND**(**bar\_num**, **menu**, **command**, **enable**,
subcommand)

Bar\_num&nbsp;&nbsp;&nbsp;&nbsp;is the menu bar in which a command
resides. Bar\_num can be the number of a built-in menu bar or the number
returned by a previously run ADD.BAR function. See ADD.COMMAND for a
list of the built-in menu bar numbers.

Menu&nbsp;&nbsp;&nbsp;&nbsp;is the menu on which the command resides.
Menu can be either the name of a menu as text or the number of a menu.
Menus are numbered starting with 1 from the left of the screen.

Command&nbsp;&nbsp;&nbsp;&nbsp;is the command you want to enable or
disable. Command can be either the name of the command as text or the
number of the command. The top command on a menu is command 1. If
command is 0, ENABLE.COMMAND enables or disables the entire menu.

Enable&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying whether the
command should be enabled or disabled. If enable is TRUE, Microsoft
Excel enables the command; if FALSE, it disables the command.

Subcommand&nbsp;&nbsp;&nbsp;&nbsp;is the name of the command on a
submenu that you want to enable. If you use subcommand, you must use
command as the name of the submenu. Use subcommand 0 to enable an entire
submenu.

**Remarks**

  - > You cannot disable built-in commands. If the specified command is
    > a built-in command or does not exist, ENABLE.COMMAND returns the
    > \#VALUE\! error value and interrupts the macro.

  - > You can hide any shortcut menu from users by using ENABLE.COMMAND
    > with command set to 0.


**Example**

The following macro formula disables a custom command that had been
added previously to the View menu on the worksheet and macro sheet menu
bar:

ENABLE.COMMAND(10, "View", "Audit...", FALSE)

**Related Functions**

[ADD.BAR](ADD.BAR.md)&nbsp;&nbsp;&nbsp;Adds a menu bar

[ADD.COMMAND](ADD.COMMAND.md)&nbsp;&nbsp;&nbsp;Adds a command to a menu

[CHECK.COMMAND](CHECK.COMMAND.md)&nbsp;&nbsp;&nbsp;Adds or deletes a check mark to or from a
command

[DELETE.COMMAND](DELETE.COMMAND.md)&nbsp;&nbsp;&nbsp;Deletes a command from a menu

[RENAME.COMMAND](RENAME.COMMAND.md)&nbsp;&nbsp;&nbsp;Changes the name of a command or menu



Return to [README](README.md#E)

