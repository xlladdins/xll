# SCROLLBAR.PROPERTIES

Sets the properties of the scroll bar and spinner button on a worksheet
or dialog sheet.

**Syntax**

**SCROLLBAR.PROPERTIES**(value, min, max, inc, page, link, 3d\_shading)

**SCROLLBAR.PROPERTIES**?(value, min, max, inc, page, link, 3d\_shading)

Value&nbsp;&nbsp;&nbsp;&nbsp;is the value of the control, and can range
from min to max, inclusive. It designates where the scroll bar button is
positioned along the scroll bar.

Min&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the minimum value that
the scroll bar can have. This number ranges from 0 to 30,000, but cannot
be greater than the maximum value given in max.

Max&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the maximum value that
the scroll bar can have. This number ranges from 0 to 30,000.

Inc&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the increment that the
value is adjusted by when the scrollbar arrow is clicked.

Page&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the increment that
the value is adjusted by when the page scroll region of a scroll bar is
clicked.

Link&nbsp;&nbsp;&nbsp;&nbsp;is the cell on the macro sheet to which the
scroll bar value is linked. Whenever the scroll bar control is changed,
the value of the control is entered into the cell. Similarly, whenever
the value in the cell is changed, the setting for the scroll bar is also
changed. To clear the link, set this value to an empty string.

3d\_shading&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that specifies
whether the scroll bar or spinner button appears as 3-D. If TRUE, the
scroll bar or spinner button will appear as 3-D. If FALSE or omitted,
the scroll bar or spinner button will not be 3-D. The argument is
available for only worksheets.

**Related Functions**

[PUSHBUTTON.PROPERTIES](PUSHBUTTON.PROPERTIES.md)&nbsp;&nbsp;&nbsp;Sets the properties of the push
button control

[EDITBOX.PROPERTIES](EDITBOX.PROPERTIES.md)&nbsp;&nbsp;&nbsp;Sets the properties of an edit box
on a worksheet or dialog sheet



Return to [README](README.md#S)

