PRESS.TOOL
==========

Formats a button so that it appears either normal or depressed into the
screen.

**Syntax**

**PRESS.TOOL**(**bar\_id, position**, down)

Bar\_id    specifies the number or name of the toolbar in which you want
to change the button appearance. For detailed information about bar\_id,
see ADD.TOOL.

Position    specifies the position of the button within the toolbar.
Position starts with 1 at the left side (if horizontal) or at the top
(if vertical).

Down    is a logical value specifying the appearance of the button. If
down is TRUE, the button appears depressed into the screen; if FALSE or
omitted, it appears normal (up).

**Remarks**

This function applies only to custom buttons to which macros have
already been assigned. An error occurs if you try to process any other
type of button.

**Example**

The following macro formula sets the third button image on Toolbar4 to
normal (up).

PRESS.TOOL(\"Toolbar4\", 3, FALSE)

**Related Functions**

ADD.TOOL   Adds one or more buttons to a toolbar

DELETE.TOOL   Deletes a button from a toolbar

Return to [top](#H)

PRINT
