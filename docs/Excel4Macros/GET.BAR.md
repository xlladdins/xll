GET.BAR

Returns the number of the active menu bar. There are two syntax forms of
GET.BAR. Use syntax 1 to return information that you can use with other
functions that manipulate menu bars. Use syntax 2 to return information
that you can use with functions that add, delete, or alter menu
commands.

Syntax 1**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Returns the number of the active menu bar

Syntax 2**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Returns the name or position number of a
specified command on a menu or of a specified menu on a menu bar



Return to [README](README.md)

GET.BAR Syntax 1

Returns the number of the active menu bar. There are two syntax forms of
GET.BAR. Use syntax 1 to return information that you can use with other
functions that manipulate menu bars. For a list of the ID numbers for
Microsoft Excel's built-in menu bars, see ADD.COMMAND.

**Syntax**

**GET.BAR**( )

**Example**

The following macro formula assigns the name OldBar to the number of the
active menu bar. This is useful if you will need to restore the current
menu bar after displaying another custom menu bar.

SET.NAME("OldBar", GET.BAR())

**Related Functions**

[ADD.BAR](ADD.BAR.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Adds a menu bar

[SHOW.BAR](SHOW.BAR.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Displays a menu bar

[GET.BAR](GET.BAR.md) Syntax 2**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Returns the name or position number of
a specified command on a menu or of a specified menu on a menu bar



Return to [README](README.md)

GET.BAR Syntax 2

Returns the name or position number of a specified command on a menu or
of a specified menu on a menu bar. There are two syntax forms of
GET.BAR. Use syntax 2 to return information that you can use with
functions that add, delete, or alter menu commands.

**Syntax**

**GET.BAR**(**bar\_num, menu, command**, subcommand)

Bar\_num**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is the number of a menu bar containing
the menu or command about which you want information. Bar\_num can be
the number of a built-in menu bar or the number returned by a previously
run ADD.BAR function. For a list of the ID numbers for Microsoft Excel's
built-in menu bars, see ADD.COMMAND.

Menu**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is the menu on which the command resides or
the menu whose name or position you want. Menu can be the name of the
menu as text or the number of the menu. Menus are numbered starting with
1 from the left of the menu bar.

Command**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is the command or submenu whose name or
number you want returned. Command can be the name of the command from
the menu as text, in which case the number is returned, or the number of
the command from the menu, in which case the name is returned. Commands
are numbered starting with 1 from the top of the menu. If command is 0,
the name or position number of the menu is returned. If an ellipsis
(...) follows a command name, such as the Open... command on the File
menu, then you must include the ellipsis when referring to that command.
See the following examples.

Subcommand**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;returns the name (if number is used
for subcommand) or position (if name is used for subcommand) of a
command on a submenu. If the command argument refers to an empty
submenu, or is a command instead of a submenu, then using subcommand
returns \#N/A.

**Remarks**

  - > If an ampersand is used to indicate the access key in the name of
    > a custom command, the ampersand is included in the name returned
    > by GET.BAR. All built-in commands have an ampersand before the
    > letter used as the access key.

  - > If the command name or position specified does not exist, GET.BAR
    > returns the \#N/A error value.


**Examples**

In the default worksheet and macro sheet menu bar:

GET.BAR(10, "File", "Print...") equals 14

GET.BAR(10, "File", 14) equals "\&Print...^tCTRL+P" (where ^t is a tab
character)

GET.BAR(10, 1, "Open") equals \#N/A

GET.BAR(10, 1, "Open...") equals 2

**Related Functions**

[ADD.COMMAND](ADD.COMMAND.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Adds a command to a menu

[DELETE.COMMAND](DELETE.COMMAND.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Deletes a command from a menu

[GET.TOOLBAR](GET.TOOLBAR.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Retrieves information about a toolbar

[RENAME.COMMAND](RENAME.COMMAND.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Changes the name of a command or menu

[GETBAR](GETBAR.md) Syntax 1**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Returns the number of the active menu
bar



Return to [README](README.md)

