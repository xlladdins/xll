FORMAT.SIZE

Equivalent to sizing an object with the mouse. Sizes the selected object
and returns TRUE. If the selected chart object cannot be sized,
FORMAT.SIZE returns FALSE. There are two syntax forms of this function.
Use syntax 1 to size worksheet objects and chart items absolutely. Use
syntax 2 relative to a cell or range of cells to size only worksheet
objects. It is generally easier to use the macro recorder to enter this
function on your macro sheet.

Syntax 1&nbsp;&nbsp;&nbsp;Sizes worksheet objects and chart items

Syntax 2&nbsp;&nbsp;&nbsp;Sizes worksheet objects relative to a cell or
range



Return to [README](README.md)

FORMAT.SIZE Syntax 1

Equivalent to sizing an object with the mouse. Sizes the selected object
and returns TRUE. If the selected chart object cannot be sized,
FORMAT.SIZE returns FALSE. There are two syntax forms of this function.
Use syntax 1 to size worksheet objects and chart items absolutely. Use
syntax 2 relative to a cell or range of cells to size only worksheet
objects. It is generally easier to use the macro recorder to enter this
function on your macro sheet.

**Syntax**

**FORMAT.SIZE**(width, height)

**FORMAT.SIZE**?(width, height)

Width&nbsp;&nbsp;&nbsp;&nbsp;specifies the width of the selected object,
measured in points. A point is 1/72nd of an inch.

Height&nbsp;&nbsp;&nbsp;&nbsp;specifies the height of the selected
object, measured in points.

You do not always have to use both arguments. For example, if you
specify height and not width, the height changes but the width does not.

**Remarks**

  - > The base of a text label on a chart is the lower-left corner of
    > the text rectangle.

  - > The base of an arrow is the end without the arrowhead.


**Related Functions**

[FORMAT.MOVE](FORMAT.MOVE.md)&nbsp;&nbsp;&nbsp;Moves the selected object

[SIZE](SIZE.md)&nbsp;&nbsp;&nbsp;Changes the size of a window

[S](S.md)yntax 2&nbsp;&nbsp;&nbsp;Sizes worksheet objects relative to a cell or
range



Return to [README](README.md)

FORMAT.SIZE Syntax 2

Equivalent to sizing an object with the mouse. Sizes the selected
worksheet object and returns TRUE. If the selected object cannot be
sized, FORMAT.SIZE returns FALSE. There are two syntax forms of this
function. Use syntax 2 to size worksheet objects relative to a cell or
range of cells. Use syntax 1 to size worksheet objects and chart items.
It is generally easier to use the macro recorder to enter this function
on your macro sheet.

**Syntax**

**FORMAT.SIZE**(x\_off, y\_off, reference)

**FORMAT.SIZE**?(x\_off, y\_off, reference)

X\_off&nbsp;&nbsp;&nbsp;&nbsp;specifies the width of the selected object
and is measured in points from the lower-right corner of the object to
the upper-left corner of reference. A point is 1/72nd of an inch. If
omitted, x\_off is assumed to be 0. If reference is omitted, x\_off
specifies the horizontal size.

Y\_off&nbsp;&nbsp;&nbsp;&nbsp;specifies the height of the selected
object and is measured in points from the lower-right corner of the
object to the upper-left corner of reference. If omitted, y\_off is
assumed to be 0. If reference is omitted, y\_off specifies the vertical
size.

Reference&nbsp;&nbsp;&nbsp;&nbsp;specifies the cell or range of cells to
use as the basis for the offset and for sizing. If reference is a range
of cells, only the upper-left cell in the range is used.

**Related Functions**

[FORMAT.MOVE](FORMAT.MOVE.md)&nbsp;&nbsp;&nbsp;Moves the selected object

[SIZE](SIZE.md)&nbsp;&nbsp;&nbsp;Changes the size of a window

[S](S.md)yntax 1&nbsp;&nbsp;&nbsp;Sizes worksheet objects and chart items



Return to [README](README.md)

