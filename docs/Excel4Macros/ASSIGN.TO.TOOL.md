ASSIGN.TO.TOOL

Assigns a macro to be run when a button is clicked with the mouse.

**Syntax**

**ASSIGN.TO.TOOL**(**bar\_id, position**, macro\_ref)

Bar\_id**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;specifies the number or name of a toolbar
to which you want to assign a macro. For more information about bar\_id,
see ADD.TOOL.

Position**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;specifies the position of the button
within the toolbar. Position starts with 1 at the left side (if
horizontal) or at the top (if vertical).

Macro\_ref**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is the name of, or a reference to, the
macro you want to run when the button is clicked. If macro\_ref is
omitted, Microsoft Excel no longer runs the previously specified macro.
After canceling the macro, if the button is a built-in button, Microsoft
Excel performs the normal default action when the button is clicked. If
the button is a custom button, Microsoft Excel displays the Assign Macro
dialog box when the button is clicked.

**Related Functions**

[ADD.TOOL](ADD.TOOL.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Adds one or more buttons to a toolbar

[GET.TOOL](GET.TOOL.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Returns information about a button or buttons
on a toolbar



Return to [README](README.md)

