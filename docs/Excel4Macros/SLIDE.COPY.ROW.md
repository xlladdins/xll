SLIDE.COPY.ROW

This function should not be used in Microsoft Excel 95 or later because
the Slide Show add-in is available only in Microsoft Excel version 5.0
or earlier versions.

Equivalent to clicking the Copy Row button on a slide show sheet. Copies
the selected slides, each of which is defined on a single row, to the
Clipboard.

If this function is not available, you must install the Slide Show
add-in.

**Syntax**

**SLIDE.COPY.ROW**( )

**Remarks**

  - > SLIDE.COPY.ROW, SLIDE.CUT.ROW, SLIDE.DELETE.ROW, and
    > SLIDE.PASTE.ROW return TRUE if successful, or FALSE if not
    > successful. If the active sheet is not a slide show or is
    > protected, these functions return the \#N/A error value. If the
    > current selection is not valid, these functions return the
    > \#VALUE\! error value.


**Related Functions**

[SLIDE.CUT.ROW](SLIDE.CUT.ROW.md)&nbsp;&nbsp;&nbsp;Cuts the selected slides and pastes them
onto the Clipboard

[SLIDE.DEFAULTS](SLIDE.DEFAULTS.md)&nbsp;&nbsp;&nbsp;Specifies default values for the active
slide show sheet

[SLIDE.DELETE.ROW](SLIDE.DELETE.ROW.md)&nbsp;&nbsp;&nbsp;Deletes the selected slides

[SLIDE.EDIT](SLIDE.EDIT.md)&nbsp;&nbsp;&nbsp;Changes the attributes of the selected slide

[SLIDE.GET](SLIDE.GET.md)&nbsp;&nbsp;&nbsp;Returns information about a slide or slide
show

[SLIDE.PASTE](SLIDE.PASTE.md)&nbsp;&nbsp;&nbsp;Pastes the contents of the Clipboard onto a
slide

[SLIDE.PASTE.ROW](SLIDE.PASTE.ROW.md)&nbsp;&nbsp;&nbsp;Pastes previously cut or copied slides
onto the current selection

[SLIDE.SHOW](SLIDE.SHOW.md)&nbsp;&nbsp;&nbsp;Starts a slide show in the active sheet



Return to [README](README.md)

