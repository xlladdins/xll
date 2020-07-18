EDITBOX.PROPERTIES

Sets the properties of an edit box on a dialog sheet.

**Syntax**

**EDITBOX.PROPERTIES**(validation\_num, multiline\_logical,
vscroll\_logical, password\_logical)

**EDITBOX.PROPERTIES**?(validation\_num, multiline\_logical,
vscroll\_logical, password\_logical)

Validation\_num&nbsp;&nbsp;&nbsp;&nbsp;is the validation applied to the
edit box when the dialog is dismissed. If the edit box contains a value
other than the type specified (or validation), an error is returned.

|                     |                                |
| ------------------- | ------------------------------ |
| **Validation\_num** | **Type**                       |
| 1                   | Text                           |
| 2                   | Integer                        |
| 3                   | Number (allows floating point) |
| 4                   | Reference                      |
| 5                   | Formula                        |

Multiline\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying
whether word wrapping is allowed in the edit box control. If TRUE, word
wrapping is allowed. If FALSE, word wrapping is not allowed

Vscroll\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying
whether edit box displays a vertical scrollbar. If TRUE, a scrollbar is
displayed. If FALSE, a scrollbar is not displayed.

Password\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying
whether edit box displays characters as the user types. If TRUE,
asterisks (\*) are displayed as the user types. If FALSE, no asterisks
are displayed.

**Related Functions**

[CHECKBOX.PROPERTIES](CHECKBOX.PROPERTIES.md)&nbsp;&nbsp;&nbsp;Sets various properties of check
box and option box controls

[PUSHBUTTON.PROPERTIES](PUSHBUTTON.PROPERTIES.md)&nbsp;&nbsp;&nbsp;Sets the properties of the push
button control



Return to [README](README.md)

