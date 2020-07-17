FORMAT.SHAPE

Equivalent to clicking the reshape button on the Drawing toolbar and
then inserting, moving, or deleting vertices of the selected polygon. A
vertex is a point defined by a pair of coordinates in one row of the
array that defines the polygon. The array is created by CREATE.OBJECT
and EXTEND.POLYGON functions.

**Syntax**

**FORMAT.SHAPE**(**vertex\_num, insert**, reference, x\_offset,
y\_offset)

Vertex\_num    is a number corresponding to the vertex you want to
insert, move, or delete.

Insert    is a logical value specifying whether to insert a vertex, or
move or delete a vertex.

  - > If insert is TRUE, Microsoft Excel inserts a vertex between the
    > vertices vertex\_num and vertex\_num-1. The number of the new
    > vertex then becomes vertex\_num. The number of the vertex
    > previously identified by vertex\_num becomes vertex\_num+1, and so
    > on.

  - > If insert is FALSE, Microsoft Excel deletes the vertex (if the
    > remaining arguments are omitted) or moves the vertex to the
    > position specified by the remaining arguments.

>  

Reference    is the reference from which the vertex you are inserting or
moving is measured; that is, the cell or range of cells to use as the
basis for the x and y offsets.

  - > If reference is a range of cells, only the upper-left cell is
    > used.

  - > If reference is omitted, the vertex is measured from the
    > upper-left corner of the polygon's bounding rectangle.

>  

X\_offset    is the horizontal distance from the upper-left corner of
reference to the vertex. X\_offset is measured in points. A point is
1/72nd of an inch. If reference is omitted, x\_offset specifies the
horizontal distance from the upper-left corner of the polygon bounding
rectangle.

Y\_offset    is the vertical distance from the upper-left corner of
reference to the vertex. Y\_offset is measured in points. If reference
is omitted, y\_offset specifies the vertical distance from the
upper-left corner of the polygon bounding rectangle.

**Remarks**

You cannot delete a vertex if only two vertices remain.

**Examples**

The following macro formula deletes the second vertex of the selected
polygon:

FORMAT.SHAPE(2, FALSE)

The following macro formula moves the thirteenth vertex 6 points to the
right and 4 points below the upper-left corner of cell B5 on the active
worksheet:

FORMAT.SHAPE(13, FALSE, \!$B$5, 6, 4)

The following macro formula inserts a new vertex between vertices 2 and
3. The new vertex is 60 points to the right and 75 points below the
upper-left corner of the polygon's bounding rectangle:

FORMAT.SHAPE(3, TRUE, , 60, 75)

**Related Functions**

CREATE.OBJECT   Creates an object

EXTEND.POLYGON   Adds vertices to a polygon


FORMAT.SHAPE

Equivalent to clicking the reshape button on the Drawing toolbar and
then inserting, moving, or deleting vertices of the selected polygon. A
vertex is a point defined by a pair of coordinates in one row of the
array that defines the polygon. The array is created by CREATE.OBJECT
and EXTEND.POLYGON functions.

**Syntax**

**FORMAT.SHAPE**(**vertex\_num, insert**, reference, x\_offset,
y\_offset)

Vertex\_num    is a number corresponding to the vertex you want to
insert, move, or delete.

Insert    is a logical value specifying whether to insert a vertex, or
move or delete a vertex.

  - > If insert is TRUE, Microsoft Excel inserts a vertex between the
    > vertices vertex\_num and vertex\_num-1. The number of the new
    > vertex then becomes vertex\_num. The number of the vertex
    > previously identified by vertex\_num becomes vertex\_num+1, and so
    > on.

  - > If insert is FALSE, Microsoft Excel deletes the vertex (if the
    > remaining arguments are omitted) or moves the vertex to the
    > position specified by the remaining arguments.

>  

Reference    is the reference from which the vertex you are inserting or
moving is measured; that is, the cell or range of cells to use as the
basis for the x and y offsets.

  - > If reference is a range of cells, only the upper-left cell is
    > used.

  - > If reference is omitted, the vertex is measured from the
    > upper-left corner of the polygon's bounding rectangle.

>  

X\_offset    is the horizontal distance from the upper-left corner of
reference to the vertex. X\_offset is measured in points. A point is
1/72nd of an inch. If reference is omitted, x\_offset specifies the
horizontal distance from the upper-left corner of the polygon bounding
rectangle.

Y\_offset    is the vertical distance from the upper-left corner of
reference to the vertex. Y\_offset is measured in points. If reference
is omitted, y\_offset specifies the vertical distance from the
upper-left corner of the polygon bounding rectangle.

**Remarks**

You cannot delete a vertex if only two vertices remain.

**Examples**

The following macro formula deletes the second vertex of the selected
polygon:

FORMAT.SHAPE(2, FALSE)

The following macro formula moves the thirteenth vertex 6 points to the
right and 4 points below the upper-left corner of cell B5 on the active
worksheet:

FORMAT.SHAPE(13, FALSE, \!$B$5, 6, 4)

The following macro formula inserts a new vertex between vertices 2 and
3. The new vertex is 60 points to the right and 75 points below the
upper-left corner of the polygon's bounding rectangle:

FORMAT.SHAPE(3, TRUE, , 60, 75)

**Related Functions**

CREATE.OBJECT   Creates an object

EXTEND.POLYGON   Adds vertices to a polygon


FORMAT.SHAPE

Equivalent to clicking the reshape button on the Drawing toolbar and
then inserting, moving, or deleting vertices of the selected polygon. A
vertex is a point defined by a pair of coordinates in one row of the
array that defines the polygon. The array is created by CREATE.OBJECT
and EXTEND.POLYGON functions.

**Syntax**

**FORMAT.SHAPE**(**vertex\_num, insert**, reference, x\_offset,
y\_offset)

Vertex\_num    is a number corresponding to the vertex you want to
insert, move, or delete.

Insert    is a logical value specifying whether to insert a vertex, or
move or delete a vertex.

  - > If insert is TRUE, Microsoft Excel inserts a vertex between the
    > vertices vertex\_num and vertex\_num-1. The number of the new
    > vertex then becomes vertex\_num. The number of the vertex
    > previously identified by vertex\_num becomes vertex\_num+1, and so
    > on.

  - > If insert is FALSE, Microsoft Excel deletes the vertex (if the
    > remaining arguments are omitted) or moves the vertex to the
    > position specified by the remaining arguments.

>  

Reference    is the reference from which the vertex you are inserting or
moving is measured; that is, the cell or range of cells to use as the
basis for the x and y offsets.

  - > If reference is a range of cells, only the upper-left cell is
    > used.

  - > If reference is omitted, the vertex is measured from the
    > upper-left corner of the polygon's bounding rectangle.

>  

X\_offset    is the horizontal distance from the upper-left corner of
reference to the vertex. X\_offset is measured in points. A point is
1/72nd of an inch. If reference is omitted, x\_offset specifies the
horizontal distance from the upper-left corner of the polygon bounding
rectangle.

Y\_offset    is the vertical distance from the upper-left corner of
reference to the vertex. Y\_offset is measured in points. If reference
is omitted, y\_offset specifies the vertical distance from the
upper-left corner of the polygon bounding rectangle.

**Remarks**

You cannot delete a vertex if only two vertices remain.

**Examples**

The following macro formula deletes the second vertex of the selected
polygon:

FORMAT.SHAPE(2, FALSE)

The following macro formula moves the thirteenth vertex 6 points to the
right and 4 points below the upper-left corner of cell B5 on the active
worksheet:

FORMAT.SHAPE(13, FALSE, \!$B$5, 6, 4)

The following macro formula inserts a new vertex between vertices 2 and
3. The new vertex is 60 points to the right and 75 points below the
upper-left corner of the polygon's bounding rectangle:

FORMAT.SHAPE(3, TRUE, , 60, 75)

**Related Functions**

CREATE.OBJECT   Creates an object

EXTEND.POLYGON   Adds vertices to a polygon


FORMAT.SHAPE

Equivalent to clicking the reshape button on the Drawing toolbar and
then inserting, moving, or deleting vertices of the selected polygon. A
vertex is a point defined by a pair of coordinates in one row of the
array that defines the polygon. The array is created by CREATE.OBJECT
and EXTEND.POLYGON functions.

**Syntax**

**FORMAT.SHAPE**(**vertex\_num, insert**, reference, x\_offset,
y\_offset)

Vertex\_num    is a number corresponding to the vertex you want to
insert, move, or delete.

Insert    is a logical value specifying whether to insert a vertex, or
move or delete a vertex.

  - > If insert is TRUE, Microsoft Excel inserts a vertex between the
    > vertices vertex\_num and vertex\_num-1. The number of the new
    > vertex then becomes vertex\_num. The number of the vertex
    > previously identified by vertex\_num becomes vertex\_num+1, and so
    > on.

  - > If insert is FALSE, Microsoft Excel deletes the vertex (if the
    > remaining arguments are omitted) or moves the vertex to the
    > position specified by the remaining arguments.

>  

Reference    is the reference from which the vertex you are inserting or
moving is measured; that is, the cell or range of cells to use as the
basis for the x and y offsets.

  - > If reference is a range of cells, only the upper-left cell is
    > used.

  - > If reference is omitted, the vertex is measured from the
    > upper-left corner of the polygon's bounding rectangle.

>  

X\_offset    is the horizontal distance from the upper-left corner of
reference to the vertex. X\_offset is measured in points. A point is
1/72nd of an inch. If reference is omitted, x\_offset specifies the
horizontal distance from the upper-left corner of the polygon bounding
rectangle.

Y\_offset    is the vertical distance from the upper-left corner of
reference to the vertex. Y\_offset is measured in points. If reference
is omitted, y\_offset specifies the vertical distance from the
upper-left corner of the polygon bounding rectangle.

**Remarks**

You cannot delete a vertex if only two vertices remain.

**Examples**

The following macro formula deletes the second vertex of the selected
polygon:

FORMAT.SHAPE(2, FALSE)

The following macro formula moves the thirteenth vertex 6 points to the
right and 4 points below the upper-left corner of cell B5 on the active
worksheet:

FORMAT.SHAPE(13, FALSE, \!$B$5, 6, 4)

The following macro formula inserts a new vertex between vertices 2 and
3. The new vertex is 60 points to the right and 75 points below the
upper-left corner of the polygon's bounding rectangle:

FORMAT.SHAPE(3, TRUE, , 60, 75)

**Related Functions**

CREATE.OBJECT   Creates an object

EXTEND.POLYGON   Adds vertices to a polygon


FORMAT.SHAPE

Equivalent to clicking the reshape button on the Drawing toolbar and
then inserting, moving, or deleting vertices of the selected polygon. A
vertex is a point defined by a pair of coordinates in one row of the
array that defines the polygon. The array is created by CREATE.OBJECT
and EXTEND.POLYGON functions.

**Syntax**

**FORMAT.SHAPE**(**vertex\_num, insert**, reference, x\_offset,
y\_offset)

Vertex\_num    is a number corresponding to the vertex you want to
insert, move, or delete.

Insert    is a logical value specifying whether to insert a vertex, or
move or delete a vertex.

  - > If insert is TRUE, Microsoft Excel inserts a vertex between the
    > vertices vertex\_num and vertex\_num-1. The number of the new
    > vertex then becomes vertex\_num. The number of the vertex
    > previously identified by vertex\_num becomes vertex\_num+1, and so
    > on.

  - > If insert is FALSE, Microsoft Excel deletes the vertex (if the
    > remaining arguments are omitted) or moves the vertex to the
    > position specified by the remaining arguments.

>  

Reference    is the reference from which the vertex you are inserting or
moving is measured; that is, the cell or range of cells to use as the
basis for the x and y offsets.

  - > If reference is a range of cells, only the upper-left cell is
    > used.

  - > If reference is omitted, the vertex is measured from the
    > upper-left corner of the polygon's bounding rectangle.

>  

X\_offset    is the horizontal distance from the upper-left corner of
reference to the vertex. X\_offset is measured in points. A point is
1/72nd of an inch. If reference is omitted, x\_offset specifies the
horizontal distance from the upper-left corner of the polygon bounding
rectangle.

Y\_offset    is the vertical distance from the upper-left corner of
reference to the vertex. Y\_offset is measured in points. If reference
is omitted, y\_offset specifies the vertical distance from the
upper-left corner of the polygon bounding rectangle.

**Remarks**

You cannot delete a vertex if only two vertices remain.

**Examples**

The following macro formula deletes the second vertex of the selected
polygon:

FORMAT.SHAPE(2, FALSE)

The following macro formula moves the thirteenth vertex 6 points to the
right and 4 points below the upper-left corner of cell B5 on the active
worksheet:

FORMAT.SHAPE(13, FALSE, \!$B$5, 6, 4)

The following macro formula inserts a new vertex between vertices 2 and
3. The new vertex is 60 points to the right and 75 points below the
upper-left corner of the polygon's bounding rectangle:

FORMAT.SHAPE(3, TRUE, , 60, 75)

**Related Functions**
**Related Functions**

CREATE.OBJECT   Creates an object

EXTEND.POLYGON   Adds vertices to a polygon


FORMAT.SHAPE

Equivalent to clicking the reshape button on the Drawing toolbar and
then inserting, moving, or deleting vertices of the selected polygon. A
vertex is a point defined by a pair of coordinates in one row of the
array that defines the polygon. The array is created by CREATE.OBJECT
and EXTEND.POLYGON functions.

**Syntax**

**FORMAT.SHAPE**(**vertex\_num, insert**, reference, x\_offset,
y\_offset)

Vertex\_num    is a number corresponding to the vertex you want to
insert, move, or delete.

Insert    is a logical value specifying whether to insert a vertex, or
move or delete a vertex.

  - > If insert is TRUE, Microsoft Excel inserts a vertex between the
    > vertices vertex\_num and vertex\_num-1. The number of the new
    > vertex then becomes vertex\_num. The number of the vertex
    > previously identified by vertex\_num becomes vertex\_num+1, and so
    > on.

  - > If insert is FALSE, Microsoft Excel deletes the vertex (if the
    > remaining arguments are omitted) or moves the vertex to the
    > position specified by the remaining arguments.

>  

Reference    is the reference from which the vertex you are inserting or
moving is measured; that is, the cell or range of cells to use as the
basis for the x and y offsets.

  - > If reference is a range of cells, only the upper-left cell is
    > used.

  - > If reference is omitted, the vertex is measured from the
    > upper-left corner of the polygon's bounding rectangle.

>  

X\_offset    is the horizontal distance from the upper-left corner of
reference to the vertex. X\_offset is measured in points. A point is
1/72nd of an inch. If reference is omitted, x\_offset specifies the
horizontal distance from the upper-left corner of the polygon bounding
rectangle.

Y\_offset    is the vertical distance from the upper-left corner of
reference to the vertex. Y\_offset is measured in points. If reference
is omitted, y\_offset specifies the vertical distance from the
upper-left corner of the polygon bounding rectangle.

**Remarks**

You cannot delete a vertex if only two vertices remain.

**Examples**

The following macro formula deletes the second vertex of the selected
polygon:

FORMAT.SHAPE(2, FALSE)

The following macro formula moves the thirteenth vertex 6 points to the
right and 4 points below the upper-left corner of cell B5 on the active
worksheet:

FORMAT.SHAPE(13, FALSE, \!$B$5, 6, 4)

The following macro formula inserts a new vertex between vertices 2 and
3. The new vertex is 60 points to the right and 75 points below the
upper-left corner of the polygon's bounding rectangle:

FORMAT.SHAPE(3, TRUE, , 60, 75)

**Related Functions**

CREATE.OBJECT   Creates an object

EXTEND.POLYGON   Adds vertices to a polygon


FORMAT.SHAPE

Equivalent to clicking the reshape button on the Drawing toolbar and
then inserting, moving, or deleting vertices of the selected polygon. A
vertex is a point defined by a pair of coordinates in one row of the
array that defines the polygon. The array is created by CREATE.OBJECT
and EXTEND.POLYGON functions.

**Syntax**

**FORMAT.SHAPE**(**vertex\_num, insert**, reference, x\_offset,
y\_offset)

Vertex\_num    is a number corresponding to the vertex you want to
insert, move, or delete.

Insert    is a logical value specifying whether to insert a vertex, or
move or delete a vertex.

  - > If insert is TRUE, Microsoft Excel inserts a vertex between the
    > vertices vertex\_num and vertex\_num-1. The number of the new
    > vertex then becomes vertex\_num. The number of the vertex
    > previously identified by vertex\_num becomes vertex\_num+1, and so
    > on.

  - > If insert is FALSE, Microsoft Excel deletes the vertex (if the
    > remaining arguments are omitted) or moves the vertex to the
    > position specified by the remaining arguments.

>  

Reference    is the reference from which the vertex you are inserting or
moving is measured; that is, the cell or range of cells to use as the
basis for the x and y offsets.

  - > If reference is a range of cells, only the upper-left cell is
    > used.

  - > If reference is omitted, the vertex is measured from the
    > upper-left corner of the polygon's bounding rectangle.

>  

X\_offset    is the horizontal distance from the upper-left corner of
reference to the vertex. X\_offset is measured in points. A point is
1/72nd of an inch. If reference is omitted, x\_offset specifies the
horizontal distance from the upper-left corner of the polygon bounding
rectangle.

Y\_offset    is the vertical distance from the upper-left corner of
reference to the vertex. Y\_offset is measured in points. If reference
is omitted, y\_offset specifies the vertical distance from the
upper-left corner of the polygon bounding rectangle.

**Remarks**

You cannot delete a vertex if only two vertices remain.

**Examples**

The following macro formula deletes the second vertex of the selected
polygon:

FORMAT.SHAPE(2, FALSE)

The following macro formula moves the thirteenth vertex 6 points to the
right and 4 points below the upper-left corner of cell B5 on the active
worksheet:

FORMAT.SHAPE(13, FALSE, \!$B$5, 6, 4)

The following macro formula inserts a new vertex between vertices 2 and
3. The new vertex is 60 points to the right and 75 points below the
upper-left corner of the polygon's bounding rectangle:

FORMAT.SHAPE(3, TRUE, , 60, 75)

**Related Functions**

CREATE.OBJECT   Creates an object

EXTEND.POLYGON   Adds vertices to a polygon


FORMAT.SHAPE

Equivalent to clicking the reshape button on the Drawing toolbar and
then inserting, moving, or deleting vertices of the selected polygon. A
vertex is a point defined by a pair of coordinates in one row of the
array that defines the polygon. The array is created by CREATE.OBJECT
and EXTEND.POLYGON functions.

**Syntax**

**FORMAT.SHAPE**(**vertex\_num, insert**, reference, x\_offset,
y\_offset)

Vertex\_num    is a number corresponding to the vertex you want to
insert, move, or delete.

Insert    is a logical value specifying whether to insert a vertex, or
move or delete a vertex.

  - > If insert is TRUE, Microsoft Excel inserts a vertex between the
    > vertices vertex\_num and vertex\_num-1. The number of the new
    > vertex then becomes vertex\_num. The number of the vertex
    > previously identified by vertex\_num becomes vertex\_num+1, and so
    > on.

  - > If insert is FALSE, Microsoft Excel deletes the vertex (if the
    > remaining arguments are omitted) or moves the vertex to the
    > position specified by the remaining arguments.

>  

Reference    is the reference from which the vertex you are inserting or
moving is measured; that is, the cell or range of cells to use as the
basis for the x and y offsets.

  - > If reference is a range of cells, only the upper-left cell is
    > used.

  - > If reference is omitted, the vertex is measured from the
    > upper-left corner of the polygon's bounding rectangle.

>  

X\_offset    is the horizontal distance from the upper-left corner of
reference to the vertex. X\_offset is measured in points. A point is
1/72nd of an inch. If reference is omitted, x\_offset specifies the
horizontal distance from the upper-left corner of the polygon bounding
rectangle.

Y\_offset    is the vertical distance from the upper-left corner of
reference to the vertex. Y\_offset is measured in points. If reference
is omitted, y\_offset specifies the vertical distance from the
upper-left corner of the polygon bounding rectangle.

**Remarks**

You cannot delete a vertex if only two vertices remain.

**Examples**

The following macro formula deletes the second vertex of the selected
polygon:

FORMAT.SHAPE(2, FALSE)

The following macro formula moves the thirteenth vertex 6 points to the
right and 4 points below the upper-left corner of cell B5 on the active
worksheet:

FORMAT.SHAPE(13, FALSE, \!$B$5, 6, 4)

The following macro formula inserts a new vertex between vertices 2 and
3. The new vertex is 60 points to the right and 75 points below the
upper-left corner of the polygon's bounding rectangle:

FORMAT.SHAPE(3, TRUE, , 60, 75)

**Related Functions**

CREATE.OBJECT   Creates an object

EXTEND.POLYGON   Adds vertices to a polygon


FORMAT.SHAPE

Equivalent to clicking the reshape button on the Drawing toolbar and
then inserting, moving, or deleting vertices of the selected polygon. A
vertex is a point defined by a pair of coordinates in one row of the
array that defines the polygon. The array is created by CREATE.OBJECT
and EXTEND.POLYGON functions.

**Syntax**

**FORMAT.SHAPE**(**vertex\_num, insert**, reference, x\_offset,
y\_offset)

Vertex\_num    is a number corresponding to the vertex you want to
insert, move, or delete.

Insert    is a logical value specifying whether to insert a vertex, or
move or delete a vertex.

  - > If insert is TRUE, Microsoft Excel inserts a vertex between the
    > vertices vertex\_num and vertex\_num-1. The number of the new
    > vertex then becomes vertex\_num. The number of the vertex
    > previously identified by vertex\_num becomes vertex\_num+1, and so
    > on.

  - > If insert is FALSE, Microsoft Excel deletes the vertex (if the
    > remaining arguments are omitted) or moves the vertex to the
    > position specified by the remaining arguments.

>  

Reference    is the reference from which the vertex you are inserting or
moving is measured; that is, the cell or range of cells to use as the
basis for the x and y offsets.

  - > If reference is a range of cells, only the upper-left cell is
    > used.

  - > If reference is omitted, the vertex is measured from the
    > upper-left corner of the polygon's bounding rectangle.

>  

X\_offset    is the horizontal distance from the upper-left corner of
reference to the vertex. X\_offset is measured in points. A point is
1/72nd of an inch. If reference is omitted, x\_offset specifies the
horizontal distance from the upper-left corner of the polygon bounding
rectangle.

Y\_offset    is the vertical distance from the upper-left corner of
reference to the vertex. Y\_offset is measured in points. If reference
is omitted, y\_offset specifies the vertical distance from the
upper-left corner of the polygon bounding rectangle.

**Remarks**

You cannot delete a vertex if only two vertices remain.

**Examples**

The following macro formula deletes the second vertex of the selected
polygon:

FORMAT.SHAPE(2, FALSE)

The following macro formula moves the thirteenth vertex 6 points to the
right and 4 points below the upper-left corner of cell B5 on the active
worksheet:

FORMAT.SHAPE(13, FALSE, \!$B$5, 6, 4)

The following macro formula inserts a new vertex between vertices 2 and
3. The new vertex is 60 points to the right and 75 points below the
upper-left corner of the polygon's bounding rectangle:

FORMAT.SHAPE(3, TRUE, , 60, 75)

**Related Functions**

[CREATE.OBJECT](CREATE.OBJECT.md)   Creates an object

[EXTEND.POLYGON](EXTEND.POLYGON.md)   Adds vertices to a polygon


