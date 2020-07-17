OPTIONS.VIEW
============

Equivalent to clicking the Options command on the Tools menu and then
clicking the View tab in the Options dialog box. Sets various view
settings.

**Syntax**

**OPTIONS.VIEW**(formula, status, notes, show\_info, object\_num,
page\_breaks, formulas, gridlines, color\_num, headers, outline, zeros,
hor\_scroll, vert\_scroll, sheet\_tabs)

**OPTIONS.VIEW**?(formula, status, notes, show\_info, object\_num,
page\_breaks, formulas, gridlines, color\_num, headers, outline, zeros,
hor\_scroll, vert\_scroll, sheet\_tabs)

Arguments correspond to check boxes and text boxes in the View tab on
the Options dialog box. Arguments corresponding to check boxes are
logical values. If an argument is TRUE, the check box is selected; if
FALSE, the check box is cleared; if omitted, the current setting is not
changed.

Formula    is a logical value corresponding to the Formula Bar check
box. If TRUE, displays the formula bar. If FALSE, the formula bar is not
displayed.

Status    is a logical value corresponding to the Status Bar check box.
If TRUE, the status bar is displayed. If FALSE, the status bar is not
displayed.

Notes    is a logical value corresponding to the Comment & Indicator
check box. If TRUE, comments and comment indicators will be displayed.
If FALSE, comments and indicators will not be displayed.

Show\_info    is a logical value corresponding to the Info Window check
box (only in Microsoft Excel 95 and earlier versions). If TRUE, displays
the Info Window. The Info Window is not available in Microsoft Excel 97
or later.

Object\_num    is a number from 1 to 3 corresponding to the display
options in the Objects box.

  ----------------- --------------------
  **Object\_num**   **Corresponds to**
  1 or omitted      Show All
  2                 Show Placeholders
  3                 Hide
  ----------------- --------------------

Page\_breaks    is a logical value corresponding to the Page Breaks
check box. If TRUE, page breaks will appear. If FALSE, page breaks will
not appear.

Formulas    is a logical value corresponding to the Formulas check box.
If TRUE, formulas will appear in the cells. If FALSE, formulas will not
appear in the cells. The default is FALSE on worksheets and TRUE on
macro sheets.

Gridlines    is a logical value corresponding to the Gridlines check
box. If TRUE, gridlines will be displayed. If FALSE, gridlines will not
appear. The default is TRUE.

Color\_num    is a number from 0 to 56 corresponding to gridline color.
Zero corresponds to automatic color and is the default value.

Headings    is a logical value corresponding to the Row & Column Headers
check box. If TRUE, row and column headers will be displayed. If FALSE,
they will not be displayed. The default is TRUE.

Outline    is a logical value corresponding to the Outline Symbols check
box. If TRUE, outline symbols will appear. If FALSE, they will not
appear. The default is TRUE.

Zeros    is a logical value corresponding to the Zero Values check box.
If TRUE, zero values will appear, If FALSE, zero values will not appear.
The default is TRUE.

Hor\_scroll    is a logical value corresponding to the Horizontal Scroll
Bar checkbox. If TRUE, the horizontal scroll bar will be displayed. If
FALSE, it will not be displayed. The default is TRUE.

Vert\_scroll    is a logical value corresponding to the Vertical Scroll
Bar checkbox. If TRUE, the vertical scroll bar will be displayed. If
FALSE, it will not be displayed. The default is TRUE.

Sheet\_tabs    is a logical value corresponding to the Sheet Tabs check
box. If TRUE, sheet tabs will be displayed. If FALSE, sheet tabs will
not be displayed. The default is TRUE.

**Related Functions**

OPTIONS.LISTS.GET   Returns contents of custom AutoFill lists

OPTIONS.LISTS.DELETE   Deletes a custom list

Return to [top](#H)

OUTLINE
