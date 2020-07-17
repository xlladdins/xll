DISPLAY Syntax 1
================

Controls whether the screen displays formulas, gridlines, row and column
headings, and other screen attributes. There are two syntax forms of
this function. Use syntax 1 to control screen display. This function is
provided for compatibility with Microsoft Excel version 4.0. To control
screen display in Microsoft Excel version 5.0 or later, see
OPTIONS.VIEW.

Arguments for this syntax form correspond to options and check boxes in
the Display Options dialog box in Microsoft Excel version 4.0. Arguments
that correspond to check boxes are logical values. If an argument is
TRUE, Microsoft Excel selects the check box; if FALSE, Microsoft Excel
clears the check box. If an argument is omitted, no action is taken.

**Syntax**

**DISPLAY**(formulas, gridlines, headings, zeros, color\_num, reserved,
outline, page\_breaks, object\_num)

**DISPLAY**?(formulas, gridlines, headings, zeros, color\_num, reserved,
outline, page\_breaks, object\_num)

Formulas    corresponds to the Formulas check box. The default is FALSE
on worksheets and TRUE on macro sheets.

Gridlines    corresponds to the Gridlines check box. The default is
TRUE.

Headings    corresponds to the Row & Column Headings check box. The
default is TRUE.

Zeros    corresponds to the Zero Values check box. The default is TRUE.

Color\_num    is a number from 0 to 56 corresponding to the gridline and
heading colors in the Display Options dialog box; 0 corresponds to
automatic color and is the default value.

Reserved    is reserved for certain international versions of Microsoft
Excel.

Outline    corresponds to the Outline Symbols check box. The default is
TRUE.

Page\_breaks    corresponds to the Automatic Page Breaks check box. The
default is FALSE.

Object\_num    is a number from 1 to 3 corresponding to the display
options in the Object box.

  ----------------- --------------------
  **Object\_num**   **Corresponds to**
  1 or omitted      Show All
  2                 Show Placeholders
  3                 Hide
  ----------------- --------------------

**Related Functions**

OPTIONS.VIEW   Controls display

WORKSPACE   Changes workspace settings

ZOOM   Enlarges or reduces a sheet in the active window

Syntax 2   Controls display of Info Window

Return to [top](#A)

DISPLAY Syntax 2
