FORMAT.MOVE
====================

Equivalent to exploding by moving a pie-chart or doughnut-chart slice
with the mouse. Sets the percentage of pie-chart or doughnut-chart slice
explosion, and, if successful, returns TRUE. If the selected object
cannot be exploded, returns FALSE. There are three syntax forms of this
function. Use syntax 1 to move worksheet items. Use syntax 2 to move
chart items. It is generally easier to use the macro recorder to enter
this function on your macro sheet.

**Syntax**

**FORMAT.MOVE**(**explosion\_num**)

Explosion\_num    is a number specifying the explosion percentage for
the selected pie slice or the entire chart (if the series is selected).
Zero is no explosion (the tip of the slice is in the center of the pie).

**Related Functions**

FORMAT.SIZE   Sizes an object

Syntax 1   Moves worksheet items

Syntax 2   Moves chart items

WINDOW.MOVE   Moves a window

Return to [top](#E)

FORMAT.NUMBER
