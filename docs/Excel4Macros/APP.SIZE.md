APP.SIZE
========

Equivalent to choosing the Size command from the Control menu for the
application window. Changes the size of the Microsoft Excel window.

**Syntax**

**APP.SIZE**(x\_num, y\_num)

**APP.SIZE**?(x\_num, y\_num)

**Note   **This function is only for Microsoft Excel for Windows. You
can use this function in macros created with Microsoft Excel for the
Macintosh, but it will return the \#N/A error value.

X\_num    specifies the width of the Microsoft Excel window in points.

Y\_num    specifies the height of the Microsoft Excel window in points.

APP.SIZE?, the dialog-box form of the function, doesn\'t display a
dialog box. Instead, it is equivalent to pressing ALT, SPACEBAR, S or to
dragging a window border with the mouse. Using APP.SIZE?, you can size
the window with the keyboard or mouse. If you specify x\_num and/or
y\_num in the dialog-box form of the function, the window is sized
according to the specified coordinates, and you are left in size mode.

**Related Functions**

APP.ACTIVATE   Switches to an application

APP.MAXIMIZE   Maximizes the Microsoft Excel application window

APP.MINIMIZE   Minimizes the Microsoft Excel application window

APP.MOVE   Moves the Microsoft Excel application window

APP.RESTORE   Restores the Microsoft Excel application window

Return to [top](#A)

APP.TITLE
