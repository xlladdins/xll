GET.OBJECT
==========

Returns information about the specified object. Use GET.OBJECT to return
information you can use in other macro formulas that manipulate objects.

**Syntax**

**GET.OBJECT**(**type\_num**, object\_id\_text, start\_num, count\_num,
item\_index)

Type\_num    is a number specifying the type of information you want
returned about an object. GET.OBJECT returns the \#VALUE! error value
(and the macro is halted) if an object isn\'t specified or if more than
one object is selected.

  --------------- ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
  **Type\_num**   **Returns**

  2               If the object is locked, returns TRUE; otherwise FALSE.

  3               Z-order position (layering) of the object; that is, the relative position of the overlapping objects, starting with 1 for the object that is most under the others.

  4               Reference of the cell under the upper-left corner of the object as text in R1C1 reference style; for a line or arc, returns the start point.

  5               X offset from the upper-left corner of the cell under the upper-left corner of the object, measured in points.

  6               Y offset from the upper-left corner of the cell under the upper-left corner of the object, measured in points.

  7               Reference of the cell under the lower-right corner of the object as text in R1C1 reference style; for a line or arc, returns the end point.

  8               X offset from the upper-left corner of the cell under the lower-right corner of the object, measured in points.

  9               Y offset from the upper-left corner of the cell under the lower-right corner of the object, measured in points.

  10              Name, including the filename, of the macro assigned to the object. If no macro is assigned, returns FALSE.

  11              Number indicating how the object moves and sizes:\
                  1 = Object moves and sizes with cells\
                  2 = Object moves with cells\
                  3 = Object is fixed
  --------------- ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

Values 12 to 21 for type\_num apply only to text boxes and buttons. If
another type of object is selected, GET.OBJECT returns the \#VALUE!
error value.

  --------------- -------------
  **Type\_num**   **Returns**
  --------------- -------------

  ---- ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  12   Text starting at start\_num for count\_num characters.
  13   Font name of all text starting at start\_num for count\_num characters. If the text contains more than one font name, returns the \#N/A error value.
  14   Font size of all text starting at start\_num for count\_num characters. If the text contains more than one font size, returns the \#N/A error value.
  15   If all text starting at start\_num for count\_num characters is bold, returns TRUE. If text contains only partial bold formatting, returns the \#N/A error value.
  16   If all text starting at start\_num for count\_num characters is italic, returns TRUE. If text contains only partial italic formatting, returns the \#N/A error value.
  17   If all text starting at start\_num for count\_num characters is underlined, returns TRUE. If text contains only partial underline formatting, returns the \#N/A error value.
  18   If all text starting at start\_num for count\_num characters is struck through, returns TRUE. If text contains only partial struck-through formatting, returns the \#N/A error value.
  19   In Microsoft Excel for the Macintosh, if all text starting at start\_num for count\_num characters is outlined, returns TRUE. If text contains only partial outline formatting, returns the \#N/A error value. Always returns FALSE in Microsoft Excel for Windows.
  20   In Microsoft Excel for the Macintosh, if all text starting at start\_num for count\_num characters is shadowed, returns TRUE. If text contains only partial shadow formatting, returns the \#N/A error value. Always returns FALSE in Microsoft Excel for Windows.
  21   Number from 0 to 56 indicating the color of all text starting at start\_num for count\_num characters; if color is automatic, returns 0. If more than one color is used, returns the \#N/A error value.
  ---- ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Values 22 to 25 for type\_num also apply only to text boxes and buttons.
If another type of object is selected, GET.OBJECT returns the \#N/A
error value.

  --------------- ----------------------------------------------------------------------------------
  **Type\_num**   **Returns**

  23              Number indicating the vertical alignment of text:\
                  1 = Top\
                  2 = Center\
                  3 = Bottom\
                  4 = Justified

  24              Number indicating the orientation of text:\
                  0 = Horizontal\
                  1 = Vertical\
                  2 = Upward\
                  3 = Downward

  25              If button or text box is set to automatic sizing, returns TRUE; otherwise FALSE.
  --------------- ----------------------------------------------------------------------------------

The following values for type\_num apply to all objects, except where
indicated.

  --------------- --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  **Type\_num**   **Returns**

  27              Number indicating the type of the border or line:\
                  0 = Custom\
                  1 = Automatic\
                  2 = None

  28              Number indicating the style of the border or line as shown in the Patterns tab in the Format Objects dialog box:\
                  0 = None\
                  1 = Solid line\
                  2 = Dashed line\
                  3 = Dotted line\
                  4 = Dashed dotted line\
                  5 = Dashed double-dotted line\
                  6 = 50% gray line\
                  7 = 75% gray line\
                  8 = 25% gray line

  29              Number from 0 to 56 indicating the color of the border or line; if the border is automatic, returns 0.

  30              Number indicating the weight of the border or line:\
                  1 = Hairline\
                  2 = Thin\
                  3 = Medium\
                  4 = Thick

  31              Number indicating the type of fill:\
                  0 = Custom\
                  1 = Automatic\
                  2 = None

  32              Number from 1 to 18 indicating the fill pattern as shown in the Format Object dialog box.

  33              Number from 0 to 56 indicating the foreground color of the fill pattern; if the fill is automatic, returns 0. If the object is a line, returns the \#N/A error value.

  34              Number from 0 to 56 indicating the background color of the fill pattern; if the fill is automatic, returns 0. If the object is a line, returns the \#N/A error value.

  35              Number indicating the width of the arrowhead:\
                  1 = Narrow\
                  2 = Medium\
                  3 = Wide\
                  If the object is not a line, returns the \#N/A error value.

  36              Number indicating the length of the arrowhead:\
                  1 = Short\
                  2 = Medium\
                  3 = Long\
                  If the object is not a line, returns the \#N/A error value.

  37              Number indicating the style of the arrowhead:\
                  1 = No head\
                  2 = Open head\
                  3 = Closed head\
                  4 = Open double-ended head\
                  5 = Closed double-ended head\
                  If the object is not a line, returns the \#N/A error value.

  38              If the border has round corners, returns TRUE; if the corners are square, returns FALSE. If the object is a line, returns the \#N/A error value.

  39              If the border has a shadow, returns TRUE; if the border has no shadow, returns FALSE. If the object is a line, returns the \#N/A error value.

  40              If the Lock Text check box in the Protection Tab of the Format Object dialog box is selected, returns TRUE; otherwise FALSE.

  41              If objects are set to be printed, returns TRUE; otherwise FALSE.

  42              The horizontal distance, measured in points, from the left edge of the active window to the left edge of the object. May be a negative number if the window is scrolled beyond the object.

  43              The vertical distance, measured in points, from the top edge of the active window to the top edge of the object. May be a negative number if the window is scrolled beyond the object.

  44              The horizontal distance, measured in points, from the left edge of the active window to the right edge of the object. May be a negative number if the window is scrolled beyond the object.

  45              The vertical distance, measured in points, from the top edge of the active window to the bottom edge of the object. May be a negative number if the window is scrolled beyond the object.

  46              The number of vertices in a polygon, or the \#N/A error value if the object is not a polygon.

  47              A count\_num by 2 array of vertex coordinates starting at start\_num in a polygon\'s array of vertices.

  48              If the object is a text box, returns the cell reference that the text box is linked to. If the object is a control on a worksheet, returns the cell reference that the control\'s value is linked to. This information is returned as a string.

  49              Returns the ID number of the object. For example, \"Rectangle 5\" returns 5. Note that the name of the object may not have this index in it if the object has been renamed by the user.

  50              Returns the object\'s classname. For example, \"Rectangle\".

  51              Returns the object name. By default, object names are the classname followed by the ID. For example, \"Rectangle 1\" is an object name, of which \"Rectangle\" is the classname, and 1 is the ID number. The object can also be renamed, in which case the name picked by the user is returned.

  52              Returns the distance from cell A1 to the Left of the object bounding rectangle in points

  53              Returns the distance from Cell A1 to the top of the object bounding rectangle in points

  54              Returns the width of object bounding rectangle in points

  55              Returns the height of object bounding rectangle in points

  56              If the object is enabled, returns TRUE; otherwise, it returns FALSE.

  57              Returns the shortcut key assignment for the control object, as text.

  58              Returns TRUE is the button control on a dialog sheet is the default button of the dialog; otherwise, returns FALSE

  59              Returns TRUE if the button control on the dialog sheet is clicked when the user presses the ESCAPE Key; otherwise, returns FALSE.

  60              Returns TRUE if the button control on a dialog sheet will close the dialog box when pressed; otherwise, returns FALSE

  61              Returns TRUE if the button control on a dialog sheet will be clicked when the user presses F1.

  62              Returns the value of the control. For a check box or radio button, Returns 1 if it is selected, zero if it is not selected, or 2 if mixed. For a List box or dropdown box, returns the index number of the selected item, or zero if no item is selected. For a scroll bar, returns the numeric value of the scroll bar.

  63              Returns the minimum value that a scroll bar or spinner button can have

  64              Returns the maximum value that a scroll bar or spinner button can have

  65              Returns the step increment value added or subtracted from the value of a scroll bar or spinner. This value is used when the arrow buttons are pressed on the control.

  66              Returns the large, or \"page\" step increment value added or subtracted from the value of a scroll bar when it is clicked in the region between the thumb and the arrow buttons.

  67              Returns the input type allowed in an edit box control:\
                  1 = Text\
                  2 = Integer\
                  3 = Number (what type)\
                  4 = Cell reference\
                  5 = Formula

  68              Returns TRUE if the edit box control allows multi-line editing with wrapped text; otherwise, it returns FALSE.

  69              Returns TRUE if the edit box has a vertical scroll bar; otherwise, it returns FALSE.

  70              Returns the object ID of the object that is linked to a list box or edit box. For a dropdown combo box that has an editable entry field, returns the object ID of itself. A dropdown box that can\'t be edited, returns FALSE.

  71              Returns the number of entries in a List box, dropdown List box, or dropdown combo box.

  72              Returns the text of the selected entry in a List box, dropdown List box, or dropdown combo box.

  73              Returns the range used to fill the entries in a List box, dropdown List box, or dropdown combo box, as text. If an empty string is returned, then the control isn\'t filled from a range.

  74              Returns the number of list lines displayed when a dropdown control is dropped.

  75              Returns TRUE the object is displayed as 3-D; otherwise, it returns FALSE.

  76              Returns the Far East phonetic accelerator key as text. Used for Far East versions of Microsoft Excel.

  77              Returns the select status of the list box:\
                  0 = single\
                  1 = simple multi-select\
                  2 = extended multi-select

  78              Returns an array of TRUE and FALSE values indicating which items are selected in a list box. If TRUE, the item is selected; If FALSE, the item is not selected.

  79              Returns TRUE if the add indent attribute is on for alignment. Returns FALSE if the add indent attribute is off for alignment. Used for only Far East versions of Microsoft Excel.
  --------------- --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Object\_id\_text    is the name and number, or number alone, of the
object you want information about. Object\_id\_text is the text
displayed in the reference area when the object is selected. If
object\_id\_text is omitted, it is assumed to be the selected object. If
object\_id\_text is omitted and no object is selected, GET.OBJECT
returns the \#REF! error value and interrupts the macro.

Start\_num    is the number of the first character in the text box or
button or the first vertex in a polygon you want information about.
Start\_num is ignored unless a text box, button, or polygon is specified
by type\_num and object\_id\_text. If start\_num is omitted, it is
assumed to be 1.

Count\_num    is the number of characters in a text box or button, or
the number of vertices in a polygon, starting at start\_num, that you
want information about. Count\_num is ignored unless a text box, button,
or polygon is specified by type\_num and object\_id\_text. If count\_num
is omitted, it is assumed to be 255.

Item\_index    is the index number or position of the item in the list
box or drop-down box that you want information about, ranging from 1 to
the number of items in the list box or drop-down box.

**Tip   **Use GET.OBJECT(45) - GET.OBJECT(43) to determine the height of
an object and GET.OBJECT(44) - GET.OBJECT(42) to determine the width.

**Examples**

The following macro formula returns the reference of the cell under the
upper-left corner of the object Oval 3 (assume the cell is E2):

GET.OBJECT(4, \"Oval 3\") returns \"R2C5\"

The following macro formula changes the protection status of the object
Rectangle 2 if it is locked:

IF(GET.OBJECT(2, \"Rectangle 2\"), OBJECT.PROTECTION(FALSE))

The following macro formula returns characters 25 through 185 from the
object Text 5:

GET.OBJECT(12, \"Text 5\", 25, 160)

**Related Functions**

CREATE.OBJECT   Creates an object

FONT.PROPERTIES   Applies a font to the selection

OBJECT.PROTECTION   Controls how an object is protected

PLACEMENT   Determines an object\'s relationship to underlying cells

Return to [top](#E)

GET.PIVOT.FIELD
