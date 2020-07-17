FORMAT.MOVE Syntax 2
====================

Equivalent to moving an object with the mouse. Moves the base of the
selected object to the specified position and, if successful, returns
TRUE. If the selected object cannot be moved, FORMAT.MOVE returns FALSE.
There are three syntax forms of this function. Use syntax 3 to move
pie-chart and doughnut-chart items. Use syntax 1 to move worksheet
objects. It is generally easier to use the macro recorder to enter this
function on your macro sheet.

**Syntax**

**FORMAT.MOVE**(**x\_pos**, **y\_pos**)

**FORMAT.MOVE**?(x\_pos, y\_pos)

X\_pos    specifies the horizontal position to which you want to move
the object and is measured in points from the base of the object to the
lower-left corner of the window. A point is 1/72nd of an inch.

Y\_pos    specifies the vertical position to which you want to move the
object and is measured in points from the base of the object to the
lower-left corner of the window.

**Remarks**

-   The base of a text label on a chart is the lower-left corner of the
    > text rectangle.

-   The base of an arrow is the end without the arrowhead.

-   The base of a pie slice is the point.

>  

**Example**

On a chart, the following macro formula moves the base of the selected
chart object 10 points to the right of and 20 points above the
lower-left corner of the window:

FORMAT.MOVE(10, 20)

**Related Functions**

FORMAT.SIZE   Sizes an object

WINDOW.MOVE   Moves a window

Syntax 1   Moves worksheet items

Syntax 3   Moves pie-chart and doughnut-chart items

Return to [top](#E)

FORMAT.MOVE Syntax 3
