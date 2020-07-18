WINDOW.SIZE

Equivalent to clicking the Size command on the Control menu or to
adjusting the sizing borders (in Microsoft Excel for Windows) or the
sizing box (in Microsoft Excel for the Macintosh) of the window with the
mouse. Changes the size of the active window by moving its lower-right
corner so that the window has the width and height you specify.
WINDOW.SIZE does not change the position of the upper-left corner of the
window, nor does it affect whether the specified window is active or
inactive.

**Syntax**

**WINDOW.SIZE**(**width, height**, window\_text)

**WINDOW.SIZE**?(width, height, window\_text)

Width&nbsp;&nbsp;&nbsp;&nbsp;specifies the width of the window and is
measured in points. A point is 1/72nd of an inch.

Height&nbsp;&nbsp;&nbsp;&nbsp;specifies the height of the window and is
measured in points.

Window\_text&nbsp;&nbsp;&nbsp;&nbsp;specifies which window to size.

  - > Window\_text is text enclosed in quotation marks or a reference to
    > a cell containing text.

  - > If window\_text is omitted, it is assumed to be the name of the
    > active window.


**Remarks**

  - > In Microsoft Excel for Windows, an error occurs if you try to
    > resize a window that has already been minimized to an icon or
    > enlarged to its maximum size. You must first restore the window to
    > its original size using the WINDOW.RESTORE function. For more
    > information, see WINDOW.RESTORE.

  - > WINDOW.SIZE replaces SIZE in earlier versions of Microsoft Excel.


**Related Functions**

[FORMAT.SIZE](FORMAT.SIZE.md)&nbsp;&nbsp;&nbsp;Sizes an object

[WINDOW.MAXIMIZE](WINDOW.MAXIMIZE.md)&nbsp;&nbsp;&nbsp;Maximizes a window

[WINDOW.MINIMIZE](WINDOW.MINIMIZE.md)&nbsp;&nbsp;&nbsp;Minimizes a window

[WINDOW.MOVE](WINDOW.MOVE.md)&nbsp;&nbsp;&nbsp;Moves a window

[WINDOW.RESTORE](WINDOW.RESTORE.md)&nbsp;&nbsp;&nbsp;Restores a window to its previous size



Return to [README](README.md)

