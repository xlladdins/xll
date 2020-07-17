FORMAT.MOVE Syntax 1
====================

Equivalent to moving an object with the mouse. Moves the selected object
to the specified position and, if successful, returns TRUE. If the
selected object cannot be moved, FORMAT.MOVE returns FALSE. There are
three syntax forms of this function. Use syntax 1 to move worksheet
objects. Use syntax 2 to move chart items. Use syntax 3 to move
pie-chart and doughnut-chart items. It is generally easier to use the
macro recorder to enter this function on your macro sheet.

**Syntax**

**FORMAT.MOVE**(**x\_offset**, **y\_offset**, reference)

**FORMAT.MOVE**?(x\_offset, y\_offset, reference)

X\_offset    specifies the horizontal position to which you want to move
the object and is measured in points from the upper-left corner of the
object to the upper-left corner of the cell specified by reference. A
point is 1/72nd of an inch.

Y\_offset    specifies the vertical position to which you want to move
the object and is measured in points from the upper-left corner of the
object to the upper-left corner of the cell specified by reference.

Reference    specifies which cell or range of cells to place the object
in relation to.

-   If reference is a range of cells, only the upper-left cell is used.

-   If reference is omitted, it is assumed to be cell A1.

>  

**Remarks**

The position of an object is based on its upper-left corner. For ovals
and arcs, the position is based on the upper-left corner of the bounding
rectangle of the object.

**Example**

The following macro formula moves an object on the active worksheet so
that it is 10 points horizontally offset and 15 points vertically offset
from cell D4:

FORMAT.MOVE(10, 15, !\$D\$4)

**Related Functions**

CREATE.OBJECT   Creates an object

FORMAT.SIZE   Sizes an object

WINDOW.MOVE   Moves a window

Syntax 2   Moves chart items

Syntax 3   Moves pie-chart and doughnut-chart items

Return to [top](#E)

FORMAT.MOVE Syntax 2
