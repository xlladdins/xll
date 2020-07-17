HIDE.OBJECT

Hides or displays the specified object.

**Syntax**

**HIDE.OBJECT**(object\_id\_text, hide)

Object\_id\_text    is the name and number, or number alone, of the
object, as text, as it appears in the reference area when the object is
selected. The name of the object is also the text returned by the
CREATE.OBJECT function, so object\_id\_text can be a reference to a cell
containing CREATE.OBJECT. To give the name of more than one object, use
the following format for object\_id\_text:

"oval 3, text 2, arc 5"

 

If object\_id\_text is omitted, the function operates on all selected
objects. If no object is selected or if the object specified by
object\_id\_text does not exist, HIDE.OBJECT returns the \#VALUE\! error
value.

Hide    is a logical value that specifies whether to hide or display the
specified object. If hide is TRUE or omitted, Microsoft Excel hides the
object; if FALSE, Microsoft Excel displays the object.

**Remarks**

Objects are not automatically selected after they are unhidden.

**Examples**

The following macro formula hides the selected object:

HIDE.OBJECT(, TRUE)

The following macro formula displays the object named Oval 3:

HIDE.OBJECT("Oval 3", FALSE)

The following macro formula displays the three specified objects:

HIDE.OBJECT("oval 3, text 2, arc 5", FALSE)

**Related Functions**

CREATE.OBJECT   Creates an object

DISPLAY   Controls how an object is displayed


