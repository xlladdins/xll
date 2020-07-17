APP.MOVE

Equivalent to clicking the Move command on the Control menu for the
application window. Moves the Microsoft Excel window. In Microsoft Excel
for Windows, if the application window is already maximized APP.MOVE
returns the \#VALUE\! error value and interrupts the macro.

**Syntax**

**APP.MOVE**(x\_num, y\_num)

**APP.MOVE**?(x\_num, y\_num)

**Note**   This function is only for Microsoft Excel for Windows. You
can use this function in macros created with Microsoft Excel for the
Macintosh, but it will return the \#N/A error value.

X\_num    specifies the horizontal position of the Microsoft Excel
window measured in points from the left edge of your screen to the left
side of the Microsoft Excel window.

Y\_num    specifies the vertical position of the Microsoft Excel window
measured in points from the top edge of your screen to the top of the
Microsoft Excel window.

**Remarks**

  - > APP.MOVE?, the dialog-box form of the function, doesn't display a
    > dialog box. Instead, it is equivalent to pressing ALT + SPACEBAR,
    > M or to dragging the title bar with the mouse. With APP.MOVE?, you
    > can move the window with the keyboard or mouse.

  - > If you specify x\_num and/or y\_num in the dialog-box form of the
    > function, the window is moved according to the specified
    > coordinates, and you are left in move mode.

**Related Functions**

[APP.ACTIVATE](APP.ACTIVATE.md)   Switches to an application

[APP.MAXIMIZE](APP.MAXIMIZE.md)   Maximizes the Microsoft Excel application window

[APP.MINIMIZE](APP.MINIMIZE.md)   Minimizes the Microsoft Excel application window

[APP.RESTORE](APP.RESTORE.md)   Restores the Microsoft Excel application window

[APP.SIZE](APP.SIZE.md)   Changes the size of the Microsoft Excel application window



Return to [README](README.md)

