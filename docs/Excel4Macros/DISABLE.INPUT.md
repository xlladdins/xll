# DISABLE.INPUT

Blocks all input from the keyboard and mouse to Microsoft Excel (except
input to displayed dialog boxes). Use DISABLE.INPUT to prevent input
from the user or from other applications.

**Syntax**

**DISABLE.INPUT**(**logical**)

Logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying whether
input is currently disabled. If logical is TRUE, input is disabled; if
FALSE, input is reenabled.

**Remarks**

Disabling input can be useful if you are using dynamic data exchange
(DDE) to communicate with Microsoft Excel from another application.

**Important**&nbsp;&nbsp;&nbsp;Be sure to end any macro that uses
DISABLE.INPUT(TRUE) with a DISABLE.INPUT(FALSE) function. If you do not
include DISABLE.INPUT(FALSE) to allow non-dialog-box input, you will not
be able to take any actions on your computer after the macro has
finished.

**Related Functions**

[CANCEL.KEY](CANCEL.KEY.md)&nbsp;&nbsp;&nbsp;Disables macro interruption

[ENTER.DATA](ENTER.DATA.md)&nbsp;&nbsp;&nbsp;Turns Data Entry mode on and off

[WORKSPACE](WORKSPACE.md)&nbsp;&nbsp;&nbsp;Changes workspace settings



Return to [README](README.md)

