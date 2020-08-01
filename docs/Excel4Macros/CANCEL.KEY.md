# CANCEL.KEY

Disables macro interruption, or specifies a macro to run when a macro is
interrupted. Use CANCEL.KEY to control what happens when a macro is
interrupted.

**Syntax**

**CANCEL.KEY**(**enable**, macro\_ref)

Enable&nbsp;&nbsp;&nbsp;&nbsp;specifies whether the macro can be
interrupted by pressing ESC in Microsoft Excel for Windows or ESC or
COMMAND+PERIOD in Microsoft Excel for the Macintosh.

|                                  |                                                           |
| -------------------------------- | --------------------------------------------------------- |
| **If enable is**                 | **Then**                                                  |
| FALSE                            | Pressing ESC or COMMAND+PERIOD does not interrupt a macro |
| TRUE and macro\_ref is omitted   | Pressing ESC or COMMAND+PERIOD interrupts a macro         |
| TRUE and macro\_ref is specified | Macro\_ref runs when ESC or COMMAND+PERIOD is pressed     |

Macro\_ref&nbsp;&nbsp;&nbsp;&nbsp;is a reference to a macro, as a cell
reference or a name, that runs when enable is TRUE and ESC or
COMMAND+PERIOD is pressed.

**Remarks**

  - > CANCEL.KEY affects only the macro that is currently running. Once
    > the macro is stopped by a RETURN or HALT function, ESC or
    > COMMAND+PERIOD is reactivated.

  - > When CANCEL.KEY is in effect, users can still cancel a dialog box
    > displayed while the macro is running.


**Examples**

The following macro formula prevents the macro from being interrupted by
pressing ESC or COMMAND+PERIOD:

CANCEL.KEY(FALSE)

The following macro formula reactivates ESC or COMMAND+PERIOD to cancel
macro execution:

CANCEL.KEY(TRUE)

The following line in a macro runs CheckCancel when ESC or
COMMAND+PERIOD is pressed:

CANCEL.KEY(TRUE, CheckCancel)

**Related Functions**

[ERROR](ERROR.md)&nbsp;&nbsp;&nbsp;Specifies an action to take if an error occurs
while a macro is running

[ON.KEY](ON.KEY.md)&nbsp;&nbsp;&nbsp;Runs a macro when a specified key is pressed

[ON.TIME](ON.TIME.md)&nbsp;&nbsp;&nbsp;Runs a macro at a specified time



Return to [README](README.md#C)

# CANCEL.KEY

Disables macro interruption, or specifies a macro to run when a macro is
interrupted. Use CANCEL.KEY to control what happens when a macro is
interrupted.

**Syntax**

**CANCEL.KEY**(**enable**, macro\_ref)

Enable&nbsp;&nbsp;&nbsp;&nbsp;specifies whether the macro can be
interrupted by pressing ESC in Microsoft Excel for Windows or ESC or
COMMAND+PERIOD in Microsoft Excel for the Macintosh.

|                                  |                                                           |
| -------------------------------- | --------------------------------------------------------- |
| **If enable is**                 | **Then**                                                  |
| FALSE                            | Pressing ESC or COMMAND+PERIOD does not interrupt a macro |
| TRUE and macro\_ref is omitted   | Pressing ESC or COMMAND+PERIOD interrupts a macro         |
| TRUE and macro\_ref is specified | Macro\_ref runs when ESC or COMMAND+PERIOD is pressed     |

Macro\_ref&nbsp;&nbsp;&nbsp;&nbsp;is a reference to a macro, as a cell
reference or a name, that runs when enable is TRUE and ESC or
COMMAND+PERIOD is pressed.

**Remarks**

  - > CANCEL.KEY affects only the macro that is currently running. Once
    > the macro is stopped by a RETURN or HALT function, ESC or
    > COMMAND+PERIOD is reactivated.

  - > When CANCEL.KEY is in effect, users can still cancel a dialog box
    > displayed while the macro is running.


**Examples**

The following macro formula prevents the macro from being interrupted by
pressing ESC or COMMAND+PERIOD:

CANCEL.KEY(FALSE)

The following macro formula reactivates ESC or COMMAND+PERIOD to cancel
macro execution:

CANCEL.KEY(TRUE)

The following line in a macro runs CheckCancel when ESC or
COMMAND+PERIOD is pressed:

CANCEL.KEY(TRUE, CheckCancel)

**Related Functions**

[ERROR](ERROR.md)&nbsp;&nbsp;&nbsp;Specifies an action to take if an error occurs
while a macro is running

[ON.KEY](ON.KEY.md)&nbsp;&nbsp;&nbsp;Runs a macro when a specified key is pressed

[ON.TIME](ON.TIME.md)&nbsp;&nbsp;&nbsp;Runs a macro at a specified time



Return to [README](README.md#C)

