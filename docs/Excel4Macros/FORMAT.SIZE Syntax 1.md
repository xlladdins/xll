FORMAT.SIZE Syntax 1
====================

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

Width    specifies the width of the selected object, measured in points.
A point is 1/72nd of an inch.

Height    specifies the height of the selected object, measured in
points.

You do not always have to use both arguments. For example, if you
specify height and not width, the height changes but the width does not.

**Remarks**

-   The base of a text label on a chart is the lower-left corner of the
    > text rectangle.

-   The base of an arrow is the end without the arrowhead.

>  

**Related Functions**

FORMAT.MOVE   Moves the selected object

SIZE   Changes the size of a window

Syntax 2   Sizes worksheet objects relative to a cell or range

Return to [top](#E)

FORMAT.SIZE Syntax 2
