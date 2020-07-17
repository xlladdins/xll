FORMAT.SIZE Syntax 2
====================

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

X\_off    specifies the width of the selected object and is measured in
points from the lower-right corner of the object to the upper-left
corner of reference. A point is 1/72nd of an inch. If omitted, x\_off is
assumed to be 0. If reference is omitted, x\_off specifies the
horizontal size.

Y\_off    specifies the height of the selected object and is measured in
points from the lower-right corner of the object to the upper-left
corner of reference. If omitted, y\_off is assumed to be 0. If reference
is omitted, y\_off specifies the vertical size.

Reference    specifies the cell or range of cells to use as the basis
for the offset and for sizing. If reference is a range of cells, only
the upper-left cell in the range is used.

**Related Functions**

FORMAT.MOVE   Moves the selected object

SIZE   Changes the size of a window

Syntax 1   Sizes worksheet objects and chart items

Return to [top](#E)

FORMAT.TEXT
