FORMAT.MOVE

Equivalent to moving an object with the mouse. Moves the selected object
to the specified position and, if successful, returns TRUE. If the
selected object cannot be moved, FORMAT.MOVE returns FALSE. There are
three syntax forms of this function. Use syntax 1 to move worksheet
objects. Use syntax 2 to move chart items. Use syntax 3 to move
pie-chart and doughnut-chart items. It is generally easier to use the
macro recorder to enter this function on your macro sheet.

Syntax 1&nbsp;&nbsp;&nbsp;Moves worksheet items

Syntax 2&nbsp;&nbsp;&nbsp;Moves chart items

Syntax 3&nbsp;&nbsp;&nbsp;Moves pie-chart and doughnut-chart items



Return to [README](README.md)

FORMAT.MOVE Syntax 1

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

X\_offset&nbsp;&nbsp;&nbsp;&nbsp;specifies the horizontal position to
which you want to move the object and is measured in points from the
upper-left corner of the object to the upper-left corner of the cell
specified by reference. A point is 1/72nd of an inch.

Y\_offset&nbsp;&nbsp;&nbsp;&nbsp;specifies the vertical position to
which you want to move the object and is measured in points from the
upper-left corner of the object to the upper-left corner of the cell
specified by reference.

Reference&nbsp;&nbsp;&nbsp;&nbsp;specifies which cell or range of cells
to place the object in relation to.

  - > If reference is a range of cells, only the upper-left cell is
    > used.

  - > If reference is omitted, it is assumed to be cell A1.

> &nbsp;

**Remarks**

The position of an object is based on its upper-left corner. For ovals
and arcs, the position is based on the upper-left corner of the bounding
rectangle of the object.

**Example**

The following macro formula moves an object on the active worksheet so
that it is 10 points horizontally offset and 15 points vertically offset
from cell D4:

FORMAT.MOVE(10, 15, \!$D$4)

**Related Functions**

[CREATE.OBJECT](CREATE.OBJECT.md)&nbsp;&nbsp;&nbsp;Creates an object

[FORMAT.SIZE](FORMAT.SIZE.md)&nbsp;&nbsp;&nbsp;Sizes an object

[WINDOW.MOVE](WINDOW.MOVE.md)&nbsp;&nbsp;&nbsp;Moves a window

[S](S.md)yntax 2&nbsp;&nbsp;&nbsp;Moves chart items

[S](S.md)yntax 3&nbsp;&nbsp;&nbsp;Moves pie-chart and doughnut-chart items



Return to [README](README.md)

FORMAT.MOVE Syntax 2

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

X\_pos&nbsp;&nbsp;&nbsp;&nbsp;specifies the horizontal position to which
you want to move the object and is measured in points from the base of
the object to the lower-left corner of the window. A point is 1/72nd of
an inch.

Y\_pos&nbsp;&nbsp;&nbsp;&nbsp;specifies the vertical position to which
you want to move the object and is measured in points from the base of
the object to the lower-left corner of the window.

**Remarks**

  - > The base of a text label on a chart is the lower-left corner of
    > the text rectangle.

  - > The base of an arrow is the end without the arrowhead.

  - > The base of a pie slice is the point.

> &nbsp;

**Example**

On a chart, the following macro formula moves the base of the selected
chart object 10 points to the right of and 20 points above the
lower-left corner of the window:

FORMAT.MOVE(10, 20)

**Related Functions**

[FORMAT.SIZE](FORMAT.SIZE.md)&nbsp;&nbsp;&nbsp;Sizes an object

[WINDOW.MOVE](WINDOW.MOVE.md)&nbsp;&nbsp;&nbsp;Moves a window

[S](S.md)yntax 1&nbsp;&nbsp;&nbsp;Moves worksheet items

[S](S.md)yntax 3&nbsp;&nbsp;&nbsp;Moves pie-chart and doughnut-chart items



Return to [README](README.md)

FORMAT.MOVE Syntax 3

Equivalent to exploding by moving a pie-chart or doughnut-chart slice
with the mouse. Sets the percentage of pie-chart or doughnut-chart slice
explosion, and, if successful, returns TRUE. If the selected object
cannot be exploded, returns FALSE. There are three syntax forms of this
function. Use syntax 1 to move worksheet items. Use syntax 2 to move
chart items. It is generally easier to use the macro recorder to enter
this function on your macro sheet.

**Syntax**

**FORMAT.MOVE**(**explosion\_num**)

Explosion\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the
explosion percentage for the selected pie slice or the entire chart (if
the series is selected). Zero is no explosion (the tip of the slice is
in the center of the pie).

**Related Functions**

[FORMAT.SIZE](FORMAT.SIZE.md)&nbsp;&nbsp;&nbsp;Sizes an object

[S](S.md)yntax 1&nbsp;&nbsp;&nbsp;Moves worksheet items

[S](S.md)yntax 2&nbsp;&nbsp;&nbsp;Moves chart items

[WINDOW.MOVE](WINDOW.MOVE.md)&nbsp;&nbsp;&nbsp;Moves a window



Return to [README](README.md)

