GET.TOOL
========

Returns information about a button or buttons on a toolbar. Use GET.TOOL
to get information about a button to use with functions that add,
delete, or alter buttons.

**Syntax**

**GET.TOOL**(**type\_num, bar\_id, position**)

Type\_num    specifies what type of information you want GET.TOOL to
return.

  --------------- --------------------------------------------------------------------------------------------------------------------------
  **Type\_num**   **Returns**
  1               The button\'s ID number. Gaps are represented by zeros.
  2               The reference of the macro assigned to the button. If no macro is assigned, GET.TOOL returns the \#N/A error value.
  3               If the button is down, returns TRUE. If the button is up, returns FALSE.
  4               If the button is enabled, returns TRUE. If the button is disabled, returns FALSE.
  5               A logical value indicating the type of the face on the button:
                  TRUE = bitmap
                  FALSE = a default button face
  6               The help\_text reference associated with the custom button. If the button is built-in, returns \#N/A.
  7               The balloon\_text reference associated with the custom button. If the button is built-in, returns the \#N/A error value.
  8               The Help context string associated with the custom button.
  9               The Tip\_text associated with the custom button.
  --------------- --------------------------------------------------------------------------------------------------------------------------

Bar\_id    specifies the number or name of the toolbar for which you
want information. For detailed information about bar\_id, see ADD.TOOL.

Position    specifies the position of the button on the toolbar.
Position starts with 1 at the left side (if horizontal) or at the top
(if vertical). A position can be occupied by a button or a gap.

**Example**

The following macro formula requests the help text associated with the
third button in Toolbar2:

GET.TOOL(6, \"Toolbar2\", 3)

**Related Functions**

ADD.TOOL   Adds one or more buttons to a toolbar

DELETE.TOOL   Deletes a button from a toolbar

ENABLE.TOOL   Enables or disables a button on a toolbar

GET.TOOLBAR   Retrieves information about a toolbar

Return to [top](#E)

GET.TOOLBAR
