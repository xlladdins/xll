EXTEND.POLYGON
==============

Adds vertices to a polygon. This function must immediately follow a
CREATE.OBJECT function or another EXTEND.POLYGON function. Use multiple
EXTEND.POLYGON functions to create arbitrarily complex polygons. It is
generally easier to use the macro recorder to enter this function on
your macro sheet.

**Syntax**

**EXTEND.POLYGON**(**array**)

Array    is an array of values, or a reference to a range of cells
containing values, that indicate the position of vertices in the
polygon. The position is measured in points and is relative to the
upper-left corner of the polygon\'s bounding rectangle.

-   A vertex is a point. Each vertex is defined by a pair of coordinates
    > in one row of array.

-   The polygon is defined by the array argument to the CREATE.OBJECT
    > function and to all the immediately following EXTEND.POLYGON
    > functions.

-   If the polygon contains many vertices, one array may not be
    > sufficient to define it. If the number of elements in the formula
    > exceeds 1024, you must include additional EXTEND.POLYGON
    > functions. If you\'re recording a macro, Microsoft Excel
    > automatically records additional EXTEND.POLYGON functions as
    > needed.

**Related Functions**

CREATE.OBJECT   Creates an object

FORMAT.SHAPE   Inserts, moves, or deletes vertices of the selected
polygon

Return to [top](#E)

EXTRACT
