MOVE.TOOL
=========

Moves or copies a button from one toolbar to another.

**Syntax**

**MOVE.TOOL**(**from\_bar\_id, from\_bar\_position**, to\_bar\_id,
to\_bar\_position, copy, width)

From\_bar\_id    specifies the number or name of a toolbar from which
you want to move or copy the button. For detailed information, see the
description of bar\_id in ADD.TOOL.

From\_bar\_position    specifies the current position of the button
within the toolbar. From\_bar\_position starts with 1 at the left side
(if horizontal) or at the top (if vertical).

To\_bar\_id    specifies the number or name of a toolbar to which you
want to move or paste the button. For detailed information, see the
description of bar\_id in ADD.TOOL. To\_bar\_id is optional if you are
moving a button within the same toolbar.

To\_bar\_position    specifies where you want to move or paste the
button within the toolbar. To\_bar\_position starts with 1 at the left
side (if horizontal) or at the top (if vertical). To\_bar\_position is
optional if you are only adjusting the width of a drop-down list.

Copy    is a logical value specifying whether to copy the button. If
copy is TRUE, the button is copied; if FALSE or omitted, the button is
moved.

Width    is the width, measured in points, of a drop-down list. If the
button you are moving is not a drop-down list, width is ignored.

**Related Functions**

ADD.TOOL   Adds one or more buttons to a toolbar

COPY.TOOL   Copies a button face to the Clipboard

GET.TOOL   Returns information about a button or buttons on a toolbar

Return to [top](#H)

NAMES
