GET.BAR Syntax 1
================

Returns the number of the active menu bar. There are two syntax forms of
GET.BAR. Use syntax 1 to return information that you can use with other
functions that manipulate menu bars. For a list of the ID numbers for
Microsoft Excel\'s built-in menu bars, see ADD.COMMAND.

**Syntax**

**GET.BAR**( )

**Example**

The following macro formula assigns the name OldBar to the number of the
active menu bar. This is useful if you will need to restore the current
menu bar after displaying another custom menu bar.

SET.NAME(\"OldBar\", GET.BAR())

**Related Functions**

ADD.BAR   Adds a menu bar

SHOW.BAR   Displays a menu bar

GET.BAR Syntax 2   Returns the name or position number of a specified
command on a menu or of a specified menu on a menu bar

Return to [top](#E)

GET.BAR Syntax 2
