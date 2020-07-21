# ADD.BAR

Creates a new menu bar and returns the bar ID number. Use the bar ID
number to identify the menu in functions that display and add menus and
commands to the menu bar. You can also use ADD.BAR to restore a built-in
menu bar with its original menus and commands.

**Syntax**

**ADD.BAR**(bar\_num)

Bar\_num&nbsp;&nbsp;&nbsp;&nbsp;is the number of a built-in menu bar
that you want to restore. Use ADD.BAR(bar\_num) to restore an unaltered
version of a built-in menu bar after you have made changes to the menu
bar's menus and commands. See ADD.COMMAND for a list of ID numbers for
built-in menu bars.

**Important**&nbsp;&nbsp;&nbsp;Restoring a built-in menu bar will remove
menus and commands added by other macros. Use ADD.COMMAND and ADD.MENU
to restore individual commands and menus.

**Remarks**

  - > ADD.BAR just creates a new menu bar; it does not display it. Use
    > SHOW.BAR to display a menu bar. The argument to the SHOW.BAR
    > function should be the number returned by ADD.BAR or a reference
    > to the cell containing ADD.BAR.

  - > You can define up to 15 custom menu bars at one time. If you carry
    > out an ADD.BAR function when more than 15 custom menu bars are
    > already defined, Microsoft Excel returns the \#VALUE\! error
    > value.

**Example**

The following formula creates a new menu bar and returns a bar ID
number:

ADD.BAR()

**Related Functions**

[ADD.COMMAND](ADD.COMMAND.md)&nbsp;&nbsp;&nbsp;Adds a command to a menu

[ADD.MENU](ADD.MENU.md)&nbsp;&nbsp;&nbsp;Adds a menu to a menu bar

[DELETE.BAR](DELETE.BAR.md)&nbsp;&nbsp;&nbsp;Deletes a menu bar

[SHOW.BAR](SHOW.BAR.md)&nbsp;&nbsp;&nbsp;Displays a menu bar



Return to [README](README.md#A)

