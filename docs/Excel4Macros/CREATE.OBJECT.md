CREATE.OBJECT
=============

Draws an object on a sheet or macro sheet and returns a value
identifying the object created. It is generally easier to use the macro
recorder to enter this function on your macro sheet.

**Syntax 1**

Lines, rectangles, ovals, arcs, pictures, text boxes, and buttons

**CREATE.OBJECT**(**obj\_type**, **ref1**, x\_offset1, y\_offset1,
**ref2**, x\_offset2, y\_offset2, text, fill, editable)

**Syntax 2**

Polygons

**CREATE.OBJECT**(**obj\_type**, **ref1**, x\_offset1, y\_offset1,
**ref2**, x\_offset2,\
y\_offset2, **array**, fill)

**Syntax 3**

Embedded charts

**CREATE.OBJECT**(**obj\_type**, **ref1**, x\_offset1, y\_offset1,
**ref2**, x\_offset2,\
y\_offset2, xy\_series, fill, gallery\_num, type\_num, plot\_visible)

Obj\_type    is a number specifying the type of object to create.

  --------------- ----------------------------------------
  **Obj\_type**   **Object**
  1               Line
  2               Rectangle
  3               Oval
  4               Arc
  5               Embedded chart
  6               Text box
  7               Button
  8               Picture (created with the camera tool)
  9               Closed polygon
  10              Open polygon
  11              Check box
  12              Option button
  13              Edit box
  14              Label
  15              Dialog frame
  16              Spinner
  17              Scroll bar
  18              List box
  19              Group box
  20              Drop down list box
  --------------- ----------------------------------------

Ref1    is a reference to the cell from which the upper-left corner of
the object is drawn, or from which the upper-left corner of the
object\'s bounding rectangle is defined.

X\_offset1    is the horizontal distance from the upper-left corner of
ref1 to the upper-left corner of the object or to the upper-left corner
of the object\'s bounding rectangle. X\_offset1 is measured in points. A
point is 1/72nd of an inch. If x\_offset1 is omitted, it is assumed to
be 0.

Y\_offset1    is the vertical distance from the upper-left corner of
ref1 to the upper-left corner of the object or to the upper-left corner
of the object\'s bounding rectangle. Y\_offset1 is measured in points.
If y\_offset1 is omitted, it is assumed to be 0.

Ref2    is a reference to the cell from which the lower-right corner of
the object is drawn, or from which the lower-right corner of the
object\'s bounding rectangle is defined.

X\_offset2    is the horizontal distance from the upper-left corner of
ref2 to the lower-right corner of the object or to the lower-right
corner of the object\'s bounding rectangle. X\_offset2 is measured in
points. If x\_offset2 is omitted, it is assumed to be 0.

Y\_offset2    is the vertical distance from the upper-left corner of
ref2 to the lower-right corner of the object or to the lower-right
corner of the object\'s bounding rectangle. Y\_offset2 is measured in
points. If y\_offset2 is omitted, it is assumed to be 0.

Text    specifies the text that appears in a text box or button. If text
is omitted for a button, the button is named \"Button n\", where n is a
number. If obj\_type is not 6 or 7, text is ignored.

Fill    is a logical value specifying whether the object is filled or
transparent. If fill is TRUE, the object is filled; if FALSE, the object
is transparent; if omitted, the object is filled with an applicable
pattern for the object being created.

Array    is an n by 2 array of values, or a reference to a range of
cells containing values, that indicate the position of each vertex in a
polygon, relative to the upper-left corner of the polygon\'s bounding
rectangle.

-   A vertex is a point that is defined by a pair of coordinates in one
    > row of array.

-   If the polygon contains many vertices, one array may not be
    > sufficient to define it. If the number of characters in the
    > formula exceeds 1024, you must include one or more EXTEND.POLYGON
    > functions. If you\'re recording a macro, Microsoft Excel
    > automatically records EXTEND.POLYGON functions as needed. For more
    > information, see EXTEND.POLYGON.

>  

Xy\_series    is a number from 0 to 3 that specifies how data is
arranged in a chart and corresponds to options in the Paste Special
dialog box.

  ---------------- ------------------------------------------------------------------------------------
  **Xy\_series**   **Result**
  0                Displays a dialog box if the selection is ambiguous
  1 or omitted     First row/column is the first data series
  2                First row/column contains the category (x) axis labels
  3                First row/column contains the x-values; the created chart is an xy (scatter) chart
  ---------------- ------------------------------------------------------------------------------------

 

-   Xy\_series is ignored unless obj\_type is 5 (chart).

-   If you want more control over how the data is arranged, use the
    > plot\_by, categories, and ser\_titles arguments to the
    > CHART.WIZARD function. For more information, see CHART.WIZARD.

>  

Gallery\_num    is a number from 1 to 15 specifying the type of embedded
chart you want to create.

  ------------------ --------------
  **Gallery\_num**   **Chart**
  1                  Area
  2                  Bar
  3                  Column
  4                  Line
  5                  Pie
  6                  Radar
  7                  XY (scatter)
  8                  Combination
  9                  3-D area
  10                 3-D bar
  11                 3-D column
  12                 3-D line
  13                 3-D pie
  14                 3-D surface
  15                 Doughnut
  ------------------ --------------

Type\_num    is a number identifying a formatting option for a chart.
The formatting options are shown in the dialog box of the AutoFormat
command that corresponds to the type of chart you\'re creating. The
first formatting option in any gallery is 1.

Plot\_visible    is a logical value that corresponds to the Plot Visible
Cells Only checkbox in the Chart tab of the Options dialog box. If FALSE
or omitted, all values are plotted.

Editable    is a logical value that determines whether the drop down
list box is editable or not. If TRUE, the drop down list box is
editable. If FALSE, the drop down list box is not editable. If obj\_type
is not 20, this argument is ignored.

**Remarks**

-   CREATE.OBJECT returns the object identifier of the object it
    > created. Object identifiers include text describing the object,
    > such as \"Text\" or \"Oval\", and a number indicating the order in
    > which the object was created. For example, CREATE.OBJECT returns
    > \"Oval 3\" after creating an oval that is the third object in the
    > workbook.

-   If the offsets are not specified, the object is drawn from the
    > upper-left corner of ref1 to the upper-left corner of ref2.

-   If the object is not a picture and either ref1 or ref2 is omitted,
    > CREATE.OBJECT returns the \#VALUE! error value and does not create
    > the object.

-   CREATE.OBJECT also selects the object.

-   You must use the COPY function before the CREATE.OBJECT function to
    > create a chart or a picture.

>  

**Tip**   To assign a macro to an object, use the ASSIGN.TO.OBJECT
function immediately after creating the object.

**Related Functions**

ASSIGN.TO.OBJECT   Assigns a macro to an object

EXTEND.POLYGON   Adds vertices to a polygon

FORMAT.MOVE   Moves the selected object

FORMAT.SHAPE   Inserts, moves, or deletes vertices of the selected
polygon

FORMAT.SIZE   Sizes an object

GET.OBJECT   Returns information about an object

OBJECT.PROPERTIES   Determines an object\'s relationship to underlying
cells

TEXT.BOX   Replaces text in a text box

Return to [top](#A)

CREATE.PUBLISHER
