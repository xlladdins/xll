GET.BAR Syntax 2
================

Returns the name or position number of a specified command on a menu or
of a specified menu on a menu bar. There are two syntax forms of
GET.BAR. Use syntax 2 to return information that you can use with
functions that add, delete, or alter menu commands.

**Syntax**

**GET.BAR**(**bar\_num, menu, command**, subcommand)

Bar\_num    is the number of a menu bar containing the menu or command
about which you want information. Bar\_num can be the number of a
built-in menu bar or the number returned by a previously run ADD.BAR
function. For a list of the ID numbers for Microsoft Excel\'s built-in
menu bars, see ADD.COMMAND.

Menu    is the menu on which the command resides or the menu whose name
or position you want. Menu can be the name of the menu as text or the
number of the menu. Menus are numbered starting with 1 from the left of
the menu bar.

Command    is the command or submenu whose name or number you want
returned. Command can be the name of the command from the menu as text,
in which case the number is returned, or the number of the command from
the menu, in which case the name is returned. Commands are numbered
starting with 1 from the top of the menu. If command is 0, the name or
position number of the menu is returned. If an ellipsis (\...) follows a
command name, such as the Open\... command on the File menu, then you
must include the ellipsis when referring to that command. See the
following examples.

Subcommand    returns the name (if number is used for subcommand) or
position (if name is used for subcommand) of a command on a submenu. If
the command argument refers to an empty submenu, or is a command instead
of a submenu, then using subcommand returns \#N/A.

**Remarks**

-   If an ampersand is used to indicate the access key in the name of a
    > custom command, the ampersand is included in the name returned by
    > GET.BAR. All built-in commands have an ampersand before the letter
    > used as the access key.

-   If the command name or position specified does not exist, GET.BAR
    > returns the \#N/A error value.

>  

**Examples**

In the default worksheet and macro sheet menu bar:

GET.BAR(10, \"File\", \"Print\...\") equals 14

GET.BAR(10, \"File\", 14) equals \"&Print\...\^tCTRL+P\" (where \^t is a
tab character)

GET.BAR(10, 1, \"Open\") equals \#N/A

GET.BAR(10, 1, \"Open\...\") equals 2

**Related Functions**

ADD.COMMAND   Adds a command to a menu

DELETE.COMMAND   Deletes a command from a menu

GET.TOOLBAR   Retrieves information about a toolbar

RENAME.COMMAND   Changes the name of a command or menu

GETBAR Syntax 1   Returns the number of the active menu bar

Return to [top](#E)

GET.CELL