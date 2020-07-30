# SLIDE.EDIT

This function should not be used in Microsoft Excel 95 or later because
the Slide Show add-in is available only in Microsoft Excel version 5.0
or earlier versions.

Equivalent to clicking the Edit button in a slide show sheet. Gives the
currently selected slide the attributes you specify.

If this function is not available, you must install the Slide Show
add-in.

**Syntax**

**SLIDE.EDIT**(effect\_num, speed\_num, advance\_rate\_num,
soundfile\_text)

**SLIDE.EDIT**?(effect\_num, speed\_num, advance\_rate\_num,
soundfile\_text)

For a description of the arguments, see SLIDE.PASTE.

**Remarks**

  - > SLIDE.EDIT returns TRUE if it successfully edits the slide, or
    > FALSE if you click the Cancel button when using the dialog-box
    > form. If the active sheet is not a slide show or is protected,
    > SLIDE.EDIT returns the \#N/A error value. If the current selection
    > is not a valid slide, SLIDE.EDIT returns the \#VALUE error value.


**Related Function**

[SLIDE.PASTE](SLIDE.PASTE.md)&nbsp;&nbsp;&nbsp;Pastes the contents of the Clipboard onto a
slide



Return to [README](README.md#S)

