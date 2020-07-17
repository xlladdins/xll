OPTIONS.EDIT

Equivalent to clicking the Options command on the tools menu and then
clicking the Edit tab in the Options dialog box. Sets various worksheet
editing options.

**Syntax**

**OPTIONS.EDIT**(incell\_edit, drag\_drop, alert, entermove, fixed,
decimals, copy\_objects, update\_links, move\_direction, autocomplete,
animations)

**OPTIONS.EDIT**?(incell\_edit, drag\_drop, alert, entermove, fixed,
decimals, copy\_objects, update\_links, move\_direction, autocomplete,
animations)

Incell\_edit    is a logical value corresponding to the Edit Directly In
Cell check box, which if TRUE allows In Cell Editing. If FALSE, editing
directly in cells is not allowed. If omitted, the dialog box setting is
not changed.

Drag\_drop    is a logical value corresponding to the Allow Cell Drag
And Drop check box, which if TRUE allows drag and dropping on sheets. If
FALSE, drag and drop is not allowed. If omitted, the dialog box setting
is not changed.

Alert    is a logical value corresponding to the Alert Before
Overwriting Cells check box, which if TRUE displays an alert message
warning you that cells containing values are about to be overwritten. If
FALSE, an alert will not be displayed if your cells are about to
overwritten. If omitted, the dialog box setting is not changed.

Entermove    is a logical value corresponding to the Move Selection
After Enter check box, which if TRUE moves the selection after the ENTER
key is pressed. If FALSE, the selection is not moved. If omitted, the
dialog box setting is not changed.

Fixed    is a logical value corresponding to the Fixed Decimal check
box, which if TRUE fixes the decimal place according to decimals. If
FALSE, the decimal places are not fixed. If omitted, the dialog box
setting is not changed.

Decimals    is a number specifying the number of decimal places.
Decimals is ignored if fixed is FALSE or omitted.

Copy\_objects     is a logical value corresponding to the Cut, Copy, And
Sort Objects With Cells check box. If TRUE allows objects to be cut,
copied and sorted with their cells. If FALSE, objects are not cut,
copied or sorted with cells. If omitted, the dialog box setting is not
changed.

Update\_links    is a logical value corresponding to the Ask To Update
Automatic Links check box, which if TRUE will prompt the user when the
workbook is opened that has links to other documents. If FALSE, the
prompt will not be displayed. If omitted, the dialog box setting is not
changed.

Move\_direction    is a number specifying the direction to move the
selection when the ENTER key is pressed and Entermove is TRUE. Setting
this number to 1 moves down one cell, 2 moves right one cell, 3 moves up
on cell, and 4 moves down one cell..

Autocomplete    is a logical value corresponding to the Enable
AutoComplete For Cell Values check box, which, if TRUE, will enable the
AutoComplete feature of Microsoft Excel 95 and later versions.

Animations    is a logical value corresponding to the Provide Feedback
With Animation check box, which if TRUE will enable the worksheet
Animations feature of Microsoft Excel 95 and later versions. Deleted
worksheet rows and columns will slowly disappear, and inserted worksheet
rows and columns will slowly appear.

**Related Function**

OPTIONS.GENERAL   Sets various general Microsoft Excel settings


