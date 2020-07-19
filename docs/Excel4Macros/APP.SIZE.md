# APP.SIZE

Equivalent to choosing the Size command from the Control menu for the
application window. Changes the size of the Microsoft Excel window.

**Syntax**

**APP.SIZE**(x\_num, y\_num)

**APP.SIZE**?(x\_num, y\_num)

**Note**&nbsp;&nbsp;&nbsp;This function is only for Microsoft Excel for
Windows. You can use this function in macros created with Microsoft
Excel for the Macintosh, but it will return the \#N/A error value.

X\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies the width of the Microsoft Excel
window in points.

Y\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies the height of the Microsoft
Excel window in points.

APP.SIZE?, the dialog-box form of the function, doesn't display a dialog
box. Instead, it is equivalent to pressing ALT, SPACEBAR, S or to
dragging a window border with the mouse. Using APP.SIZE?, you can size
the window with the keyboard or mouse. If you specify x\_num and/or
y\_num in the dialog-box form of the function, the window is sized
according to the specified coordinates, and you are left in size mode.

**Related Functions**

[APP.ACTIVATE](APP.ACTIVATE.md)&nbsp;&nbsp;&nbsp;Switches to an application

[APP.MAXIMIZE](APP.MAXIMIZE.md)&nbsp;&nbsp;&nbsp;Maximizes the Microsoft Excel application
window

[APP.MINIMIZE](APP.MINIMIZE.md)&nbsp;&nbsp;&nbsp;Minimizes the Microsoft Excel application
window

[APP.MOVE](APP.MOVE.md)&nbsp;&nbsp;&nbsp;Moves the Microsoft Excel application window

[APP.RESTORE](APP.RESTORE.md)&nbsp;&nbsp;&nbsp;Restores the Microsoft Excel application
window



Return to [README](README.md)

