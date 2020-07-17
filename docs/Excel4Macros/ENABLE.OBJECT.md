ENABLE.OBJECT
=============

Enables or disables a drawing object or the selected drawing object. A
disabled object will not run any macro events assigned to it, and the
controls will be grayed out.

**Syntax**

**ENABLE.OBJECT**(object\_id\_text, **enable\_logical**)

Object\_id\_text    is the name of the object(s) as text. If omitted,
the selected object(s) are assumed.

Enable\_logical    is a logical value that specifies whether the object
is to be enabled. If TRUE, the object is enabled. If FALSE, the object
is disabled.

**Examples**

ENABLE.OBJECT(\"Button 2\",FALSE) disables the button with object name
Button 2 on the dialog box.

**Related Function**

SET.CONTROL.VALUE   Changes the value of the active control

Return to [top](#E)

ENABLE.TIPWIZARD
