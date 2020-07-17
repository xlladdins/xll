DISPLAY

Controls whether the screen displays formulas, gridlines, row and column
headings, and other screen attributes. There are two syntax forms of
this function. Use syntax 1 to control screen display. Use syntax 2 to
control the display of the Info Window.

Syntax 1   Controls screen display

Syntax 2   Controls display of Info Window


DISPLAY Syntax 1

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

|                 |                    |
| --------------- | ------------------ |
| **Object\_num** | **Corresponds to** |
| 1 or omitted    | Show All           |
| 2               | Show Placeholders  |
| 3               | Hide               |

**Related Functions**

OPTIONS.VIEW   Controls display

WORKSPACE   Changes workspace settings

ZOOM   Enlarges or reduces a sheet in the active window

Syntax 2   Controls display of Info Window


DISPLAY Syntax 2

Equivalent to clicking the commands on the Info menu when the Info
Window is active. Controls which commands on the Info Window are in
effect. There are two syntax forms of this function. Use syntax 2 to
control the display of the Info Window. The Info Window must be active
to use this form of DISPLAY. This function is included for compatibility
with Microsoft Excel 95 or earlier; the Info Window is not available in
Microsoft Excel 97 or later.

Arguments in this syntax form correspond to commands on the Info menu
with the same names.

For these arguments:

  - > If the argument is TRUE, Microsoft Excel displays the
    > corresponding Info item.

  - > If the argument is FALSE, Microsoft Excel does not display the
    > corresponding Info item.

  - > If the argument is omitted, the status of the item is unchanged.

>  

**Syntax**

For controlling Info Window display

**DISPLAY**(cell, formula, value, format, protection, names, precedents,
dependents, note)

Cell    is a logical value that corresponds to the Cell command and
controls the display of cell information in the Info Window. If TRUE,
cell information will be displayed; if FALSE, cell information will not
be displayed.

Formula    is a logical value that corresponds to the Formula command
and controls the display of formula information in the Info Window. If
TRUE, formula information will be displayed; if FALSE, formula
information will not be displayed.

Value    is a logical value that corresponds to the Value command and
controls the display of value information in the Info Window. If TRUE,
value information will be displayed; if FALSE, value information will
not be displayed.

Format    is a logical value that corresponds to the Format command and
controls the display of format information in the Info Window. If TRUE,
format information will be displayed; if FALSE, format information will
not be displayed.

Protection    is a logical value that corresponds to the Protection
command and controls the display of protection information in the Info
Window. If TRUE, protection information will be displayed; if FALSE,
protection information will not be displayed.

Names    is a logical value that corresponds to the Names command and
controls the display of name information in the Info Window. If TRUE,
name information will be displayed; if FALSE, name information will not
be displayed.

Precedents    is a number from 1 to 3 that specifies which precedents to
list, according to the following table.

Dependents    is a number from 1 to 3 that specifies which dependents to
list, according to the following table.

|                              |             |
| ---------------------------- | ----------- |
| **Precedents or dependents** | **List**    |
| 0                            | None        |
| 1                            | Direct only |
| 2                            | All levels  |

Note    is a logical value that corresponds to the Note command and
controls the display of note information in the Info Window. If TRUE,
note information will be displayed; if FALSE, note information will not
be displayed.

**Related Functions**

SHOW.INFO   Controls the display of the Info Window

ZOOM   Enlarges or reduces a sheet in the active window

Syntax 1   Controls screen display


DISPLAY

Controls whether the screen displays formulas, gridlines, row and column
headings, and other screen attributes. There are two syntax forms of
this function. Use syntax 1 to control screen display. Use syntax 2 to
control the display of the Info Window.

Syntax 1   Controls screen display

Syntax 2   Controls display of Info Window


DISPLAY Syntax 1

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

|                 |                    |
| --------------- | ------------------ |
| **Object\_num** | **Corresponds to** |
| 1 or omitted    | Show All           |
| 2               | Show Placeholders  |
| 3               | Hide               |

**Related Functions**

OPTIONS.VIEW   Controls display

WORKSPACE   Changes workspace settings

ZOOM   Enlarges or reduces a sheet in the active window

Syntax 2   Controls display of Info Window


DISPLAY Syntax 2

Equivalent to clicking the commands on the Info menu when the Info
Window is active. Controls which commands on the Info Window are in
effect. There are two syntax forms of this function. Use syntax 2 to
control the display of the Info Window. The Info Window must be active
to use this form of DISPLAY. This function is included for compatibility
with Microsoft Excel 95 or earlier; the Info Window is not available in
Microsoft Excel 97 or later.

Arguments in this syntax form correspond to commands on the Info menu
with the same names.

For these arguments:

  - > If the argument is TRUE, Microsoft Excel displays the
    > corresponding Info item.

  - > If the argument is FALSE, Microsoft Excel does not display the
    > corresponding Info item.

  - > If the argument is omitted, the status of the item is unchanged.

>  

**Syntax**

For controlling Info Window display

**DISPLAY**(cell, formula, value, format, protection, names, precedents,
dependents, note)

Cell    is a logical value that corresponds to the Cell command and
controls the display of cell information in the Info Window. If TRUE,
cell information will be displayed; if FALSE, cell information will not
be displayed.

Formula    is a logical value that corresponds to the Formula command
and controls the display of formula information in the Info Window. If
TRUE, formula information will be displayed; if FALSE, formula
information will not be displayed.

Value    is a logical value that corresponds to the Value command and
controls the display of value information in the Info Window. If TRUE,
value information will be displayed; if FALSE, value information will
not be displayed.

Format    is a logical value that corresponds to the Format command and
controls the display of format information in the Info Window. If TRUE,
format information will be displayed; if FALSE, format information will
not be displayed.

Protection    is a logical value that corresponds to the Protection
command and controls the display of protection information in the Info
Window. If TRUE, protection information will be displayed; if FALSE,
protection information will not be displayed.

Names    is a logical value that corresponds to the Names command and
controls the display of name information in the Info Window. If TRUE,
name information will be displayed; if FALSE, name information will not
be displayed.

Precedents    is a number from 1 to 3 that specifies which precedents to
list, according to the following table.

Dependents    is a number from 1 to 3 that specifies which dependents to
list, according to the following table.

|                              |             |
| ---------------------------- | ----------- |
| **Precedents or dependents** | **List**    |
| 0                            | None        |
| 1                            | Direct only |
| 2                            | All levels  |

Note    is a logical value that corresponds to the Note command and
controls the display of note information in the Info Window. If TRUE,
note information will be displayed; if FALSE, note information will not
be displayed.

**Related Functions**

SHOW.INFO   Controls the display of the Info Window

ZOOM   Enlarges or reduces a sheet in the active window

Syntax 1   Controls screen display


DISPLAY

Controls whether the screen displays formulas, gridlines, row and column
headings, and other screen attributes. There are two syntax forms of
this function. Use syntax 1 to control screen display. Use syntax 2 to
control the display of the Info Window.

Syntax 1   Controls screen display

Syntax 2   Controls display of Info Window


DISPLAY Syntax 1

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

|                 |                    |
| --------------- | ------------------ |
| **Object\_num** | **Corresponds to** |
| 1 or omitted    | Show All           |
| 2               | Show Placeholders  |
| 3               | Hide               |

**Related Functions**

OPTIONS.VIEW   Controls display

WORKSPACE   Changes workspace settings

ZOOM   Enlarges or reduces a sheet in the active window

Syntax 2   Controls display of Info Window


DISPLAY Syntax 2

Equivalent to clicking the commands on the Info menu when the Info
Window is active. Controls which commands on the Info Window are in
effect. There are two syntax forms of this function. Use syntax 2 to
control the display of the Info Window. The Info Window must be active
to use this form of DISPLAY. This function is included for compatibility
with Microsoft Excel 95 or earlier; the Info Window is not available in
Microsoft Excel 97 or later.

Arguments in this syntax form correspond to commands on the Info menu
with the same names.

For these arguments:

  - > If the argument is TRUE, Microsoft Excel displays the
    > corresponding Info item.

  - > If the argument is FALSE, Microsoft Excel does not display the
    > corresponding Info item.

  - > If the argument is omitted, the status of the item is unchanged.

>  

**Syntax**

For controlling Info Window display

**DISPLAY**(cell, formula, value, format, protection, names, precedents,
dependents, note)

Cell    is a logical value that corresponds to the Cell command and
controls the display of cell information in the Info Window. If TRUE,
cell information will be displayed; if FALSE, cell information will not
be displayed.

Formula    is a logical value that corresponds to the Formula command
and controls the display of formula information in the Info Window. If
TRUE, formula information will be displayed; if FALSE, formula
information will not be displayed.

Value    is a logical value that corresponds to the Value command and
controls the display of value information in the Info Window. If TRUE,
value information will be displayed; if FALSE, value information will
not be displayed.

Format    is a logical value that corresponds to the Format command and
controls the display of format information in the Info Window. If TRUE,
format information will be displayed; if FALSE, format information will
not be displayed.

Protection    is a logical value that corresponds to the Protection
command and controls the display of protection information in the Info
Window. If TRUE, protection information will be displayed; if FALSE,
protection information will not be displayed.

Names    is a logical value that corresponds to the Names command and
controls the display of name information in the Info Window. If TRUE,
name information will be displayed; if FALSE, name information will not
be displayed.

Precedents    is a number from 1 to 3 that specifies which precedents to
list, according to the following table.

Dependents    is a number from 1 to 3 that specifies which dependents to
list, according to the following table.

|                              |             |
| ---------------------------- | ----------- |
| **Precedents or dependents** | **List**    |
| 0                            | None        |
| 1                            | Direct only |
| 2                            | All levels  |

Note    is a logical value that corresponds to the Note command and
controls the display of note information in the Info Window. If TRUE,
note information will be displayed; if FALSE, note information will not
be displayed.

**Related Functions**

SHOW.INFO   Controls the display of the Info Window

ZOOM   Enlarges or reduces a sheet in the active window

Syntax 1   Controls screen display


DISPLAY

Controls whether the screen displays formulas, gridlines, row and column
headings, and other screen attributes. There are two syntax forms of
this function. Use syntax 1 to control screen display. Use syntax 2 to
control the display of the Info Window.

Syntax 1   Controls screen display

Syntax 2   Controls display of Info Window


DISPLAY Syntax 1

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

|                 |                    |
| --------------- | ------------------ |
| **Object\_num** | **Corresponds to** |
| 1 or omitted    | Show All           |
| 2               | Show Placeholders  |
| 3               | Hide               |

**Related Functions**

OPTIONS.VIEW   Controls display

WORKSPACE   Changes workspace settings

ZOOM   Enlarges or reduces a sheet in the active window

Syntax 2   Controls display of Info Window


DISPLAY Syntax 2

Equivalent to clicking the commands on the Info menu when the Info
Window is active. Controls which commands on the Info Window are in
effect. There are two syntax forms of this function. Use syntax 2 to
control the display of the Info Window. The Info Window must be active
to use this form of DISPLAY. This function is included for compatibility
with Microsoft Excel 95 or earlier; the Info Window is not available in
Microsoft Excel 97 or later.

Arguments in this syntax form correspond to commands on the Info menu
with the same names.

For these arguments:

  - > If the argument is TRUE, Microsoft Excel displays the
    > corresponding Info item.

  - > If the argument is FALSE, Microsoft Excel does not display the
    > corresponding Info item.

  - > If the argument is omitted, the status of the item is unchanged.

>  

**Syntax**

For controlling Info Window display

**DISPLAY**(cell, formula, value, format, protection, names, precedents,
dependents, note)

Cell    is a logical value that corresponds to the Cell command and
controls the display of cell information in the Info Window. If TRUE,
cell information will be displayed; if FALSE, cell information will not
be displayed.

Formula    is a logical value that corresponds to the Formula command
and controls the display of formula information in the Info Window. If
TRUE, formula information will be displayed; if FALSE, formula
information will not be displayed.

Value    is a logical value that corresponds to the Value command and
controls the display of value information in the Info Window. If TRUE,
value information will be displayed; if FALSE, value information will
not be displayed.

Format    is a logical value that corresponds to the Format command and
controls the display of format information in the Info Window. If TRUE,
format information will be displayed; if FALSE, format information will
not be displayed.

Protection    is a logical value that corresponds to the Protection
command and controls the display of protection information in the Info
Window. If TRUE, protection information will be displayed; if FALSE,
protection information will not be displayed.

Names    is a logical value that corresponds to the Names command and
controls the display of name information in the Info Window. If TRUE,
name information will be displayed; if FALSE, name information will not
be displayed.

Precedents    is a number from 1 to 3 that specifies which precedents to
list, according to the following table.

Dependents    is a number from 1 to 3 that specifies which dependents to
list, according to the following table.

|                              |             |
| ---------------------------- | ----------- |
| **Precedents or dependents** | **List**    |
| 0                            | None        |
| 1                            | Direct only |
| 2                            | All levels  |

Note    is a logical value that corresponds to the Note command and
controls the display of note information in the Info Window. If TRUE,
note information will be displayed; if FALSE, note information will not
be displayed.

**Related Functions**

SHOW.INFO   Controls the display of the Info Window

ZOOM   Enlarges or reduces a sheet in the active window

Syntax 1   Controls screen display


DISPLAY

Controls whether the screen displays formulas, gridlines, row and column
headings, and other screen attributes. There are two syntax forms of
this function. Use syntax 1 to control screen display. Use syntax 2 to
control the display of the Info Window.

Syntax 1   Controls screen display

Syntax 2   Controls display of Info Window


DISPLAY Syntax 1

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

|                 |                    |
| --------------- | ------------------ |
| **Object\_num** | **Corresponds to** |
| 1 or omitted    | Show All           |
| 2               | Show Placeholders  |
| 3               | Hide               |

**Related Functions**
**Related Functions**

OPTIONS.VIEW   Controls display

WORKSPACE   Changes workspace settings

ZOOM   Enlarges or reduces a sheet in the active window

Syntax 2   Controls display of Info Window


DISPLAY Syntax 2

Equivalent to clicking the commands on the Info menu when the Info
Window is active. Controls which commands on the Info Window are in
effect. There are two syntax forms of this function. Use syntax 2 to
control the display of the Info Window. The Info Window must be active
to use this form of DISPLAY. This function is included for compatibility
with Microsoft Excel 95 or earlier; the Info Window is not available in
Microsoft Excel 97 or later.

Arguments in this syntax form correspond to commands on the Info menu
with the same names.

For these arguments:

  - > If the argument is TRUE, Microsoft Excel displays the
    > corresponding Info item.

  - > If the argument is FALSE, Microsoft Excel does not display the
    > corresponding Info item.

  - > If the argument is omitted, the status of the item is unchanged.

>  

**Syntax**

For controlling Info Window display

**DISPLAY**(cell, formula, value, format, protection, names, precedents,
dependents, note)

Cell    is a logical value that corresponds to the Cell command and
controls the display of cell information in the Info Window. If TRUE,
cell information will be displayed; if FALSE, cell information will not
be displayed.

Formula    is a logical value that corresponds to the Formula command
and controls the display of formula information in the Info Window. If
TRUE, formula information will be displayed; if FALSE, formula
information will not be displayed.

Value    is a logical value that corresponds to the Value command and
controls the display of value information in the Info Window. If TRUE,
value information will be displayed; if FALSE, value information will
not be displayed.

Format    is a logical value that corresponds to the Format command and
controls the display of format information in the Info Window. If TRUE,
format information will be displayed; if FALSE, format information will
not be displayed.

Protection    is a logical value that corresponds to the Protection
command and controls the display of protection information in the Info
Window. If TRUE, protection information will be displayed; if FALSE,
protection information will not be displayed.

Names    is a logical value that corresponds to the Names command and
controls the display of name information in the Info Window. If TRUE,
name information will be displayed; if FALSE, name information will not
be displayed.

Precedents    is a number from 1 to 3 that specifies which precedents to
list, according to the following table.

Dependents    is a number from 1 to 3 that specifies which dependents to
list, according to the following table.

|                              |             |
| ---------------------------- | ----------- |
| **Precedents or dependents** | **List**    |
| 0                            | None        |
| 1                            | Direct only |
| 2                            | All levels  |

Note    is a logical value that corresponds to the Note command and
controls the display of note information in the Info Window. If TRUE,
note information will be displayed; if FALSE, note information will not
be displayed.

**Related Functions**
**Related Functions**

SHOW.INFO   Controls the display of the Info Window

ZOOM   Enlarges or reduces a sheet in the active window

Syntax 1   Controls screen display


DISPLAY

Controls whether the screen displays formulas, gridlines, row and column
headings, and other screen attributes. There are two syntax forms of
this function. Use syntax 1 to control screen display. Use syntax 2 to
control the display of the Info Window.

Syntax 1   Controls screen display

Syntax 2   Controls display of Info Window


DISPLAY Syntax 1

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

|                 |                    |
| --------------- | ------------------ |
| **Object\_num** | **Corresponds to** |
| 1 or omitted    | Show All           |
| 2               | Show Placeholders  |
| 3               | Hide               |

**Related Functions**

OPTIONS.VIEW   Controls display

WORKSPACE   Changes workspace settings

ZOOM   Enlarges or reduces a sheet in the active window

Syntax 2   Controls display of Info Window


DISPLAY Syntax 2

Equivalent to clicking the commands on the Info menu when the Info
Window is active. Controls which commands on the Info Window are in
effect. There are two syntax forms of this function. Use syntax 2 to
control the display of the Info Window. The Info Window must be active
to use this form of DISPLAY. This function is included for compatibility
with Microsoft Excel 95 or earlier; the Info Window is not available in
Microsoft Excel 97 or later.

Arguments in this syntax form correspond to commands on the Info menu
with the same names.

For these arguments:

  - > If the argument is TRUE, Microsoft Excel displays the
    > corresponding Info item.

  - > If the argument is FALSE, Microsoft Excel does not display the
    > corresponding Info item.

  - > If the argument is omitted, the status of the item is unchanged.

>  

**Syntax**

For controlling Info Window display

**DISPLAY**(cell, formula, value, format, protection, names, precedents,
dependents, note)

Cell    is a logical value that corresponds to the Cell command and
controls the display of cell information in the Info Window. If TRUE,
cell information will be displayed; if FALSE, cell information will not
be displayed.

Formula    is a logical value that corresponds to the Formula command
and controls the display of formula information in the Info Window. If
TRUE, formula information will be displayed; if FALSE, formula
information will not be displayed.

Value    is a logical value that corresponds to the Value command and
controls the display of value information in the Info Window. If TRUE,
value information will be displayed; if FALSE, value information will
not be displayed.

Format    is a logical value that corresponds to the Format command and
controls the display of format information in the Info Window. If TRUE,
format information will be displayed; if FALSE, format information will
not be displayed.

Protection    is a logical value that corresponds to the Protection
command and controls the display of protection information in the Info
Window. If TRUE, protection information will be displayed; if FALSE,
protection information will not be displayed.

Names    is a logical value that corresponds to the Names command and
controls the display of name information in the Info Window. If TRUE,
name information will be displayed; if FALSE, name information will not
be displayed.

Precedents    is a number from 1 to 3 that specifies which precedents to
list, according to the following table.

Dependents    is a number from 1 to 3 that specifies which dependents to
list, according to the following table.

|                              |             |
| ---------------------------- | ----------- |
| **Precedents or dependents** | **List**    |
| 0                            | None        |
| 1                            | Direct only |
| 2                            | All levels  |

Note    is a logical value that corresponds to the Note command and
controls the display of note information in the Info Window. If TRUE,
note information will be displayed; if FALSE, note information will not
be displayed.

**Related Functions**

SHOW.INFO   Controls the display of the Info Window

ZOOM   Enlarges or reduces a sheet in the active window

Syntax 1   Controls screen display


DISPLAY

Controls whether the screen displays formulas, gridlines, row and column
headings, and other screen attributes. There are two syntax forms of
this function. Use syntax 1 to control screen display. Use syntax 2 to
control the display of the Info Window.

Syntax 1   Controls screen display

Syntax 2   Controls display of Info Window


DISPLAY Syntax 1

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

|                 |                    |
| --------------- | ------------------ |
| **Object\_num** | **Corresponds to** |
| 1 or omitted    | Show All           |
| 2               | Show Placeholders  |
| 3               | Hide               |

**Related Functions**

OPTIONS.VIEW   Controls display

WORKSPACE   Changes workspace settings

ZOOM   Enlarges or reduces a sheet in the active window

Syntax 2   Controls display of Info Window


DISPLAY Syntax 2

Equivalent to clicking the commands on the Info menu when the Info
Window is active. Controls which commands on the Info Window are in
effect. There are two syntax forms of this function. Use syntax 2 to
control the display of the Info Window. The Info Window must be active
to use this form of DISPLAY. This function is included for compatibility
with Microsoft Excel 95 or earlier; the Info Window is not available in
Microsoft Excel 97 or later.

Arguments in this syntax form correspond to commands on the Info menu
with the same names.

For these arguments:

  - > If the argument is TRUE, Microsoft Excel displays the
    > corresponding Info item.

  - > If the argument is FALSE, Microsoft Excel does not display the
    > corresponding Info item.

  - > If the argument is omitted, the status of the item is unchanged.

>  

**Syntax**

For controlling Info Window display

**DISPLAY**(cell, formula, value, format, protection, names, precedents,
dependents, note)

Cell    is a logical value that corresponds to the Cell command and
controls the display of cell information in the Info Window. If TRUE,
cell information will be displayed; if FALSE, cell information will not
be displayed.

Formula    is a logical value that corresponds to the Formula command
and controls the display of formula information in the Info Window. If
TRUE, formula information will be displayed; if FALSE, formula
information will not be displayed.

Value    is a logical value that corresponds to the Value command and
controls the display of value information in the Info Window. If TRUE,
value information will be displayed; if FALSE, value information will
not be displayed.

Format    is a logical value that corresponds to the Format command and
controls the display of format information in the Info Window. If TRUE,
format information will be displayed; if FALSE, format information will
not be displayed.

Protection    is a logical value that corresponds to the Protection
command and controls the display of protection information in the Info
Window. If TRUE, protection information will be displayed; if FALSE,
protection information will not be displayed.

Names    is a logical value that corresponds to the Names command and
controls the display of name information in the Info Window. If TRUE,
name information will be displayed; if FALSE, name information will not
be displayed.

Precedents    is a number from 1 to 3 that specifies which precedents to
list, according to the following table.

Dependents    is a number from 1 to 3 that specifies which dependents to
list, according to the following table.

|                              |             |
| ---------------------------- | ----------- |
| **Precedents or dependents** | **List**    |
| 0                            | None        |
| 1                            | Direct only |
| 2                            | All levels  |

Note    is a logical value that corresponds to the Note command and
controls the display of note information in the Info Window. If TRUE,
note information will be displayed; if FALSE, note information will not
be displayed.

**Related Functions**

SHOW.INFO   Controls the display of the Info Window

ZOOM   Enlarges or reduces a sheet in the active window

Syntax 1   Controls screen display


DISPLAY

Controls whether the screen displays formulas, gridlines, row and column
headings, and other screen attributes. There are two syntax forms of
this function. Use syntax 1 to control screen display. Use syntax 2 to
control the display of the Info Window.

Syntax 1   Controls screen display

Syntax 2   Controls display of Info Window


DISPLAY Syntax 1

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

|                 |                    |
| --------------- | ------------------ |
| **Object\_num** | **Corresponds to** |
| 1 or omitted    | Show All           |
| 2               | Show Placeholders  |
| 3               | Hide               |

**Related Functions**

OPTIONS.VIEW   Controls display

WORKSPACE   Changes workspace settings

ZOOM   Enlarges or reduces a sheet in the active window

Syntax 2   Controls display of Info Window


DISPLAY Syntax 2

Equivalent to clicking the commands on the Info menu when the Info
Window is active. Controls which commands on the Info Window are in
effect. There are two syntax forms of this function. Use syntax 2 to
control the display of the Info Window. The Info Window must be active
to use this form of DISPLAY. This function is included for compatibility
with Microsoft Excel 95 or earlier; the Info Window is not available in
Microsoft Excel 97 or later.

Arguments in this syntax form correspond to commands on the Info menu
with the same names.

For these arguments:

  - > If the argument is TRUE, Microsoft Excel displays the
    > corresponding Info item.

  - > If the argument is FALSE, Microsoft Excel does not display the
    > corresponding Info item.

  - > If the argument is omitted, the status of the item is unchanged.

>  

**Syntax**

For controlling Info Window display

**DISPLAY**(cell, formula, value, format, protection, names, precedents,
dependents, note)

Cell    is a logical value that corresponds to the Cell command and
controls the display of cell information in the Info Window. If TRUE,
cell information will be displayed; if FALSE, cell information will not
be displayed.

Formula    is a logical value that corresponds to the Formula command
and controls the display of formula information in the Info Window. If
TRUE, formula information will be displayed;DISPLAY

Controls whether the screen displays formulas, gridlines, row and column
headings, and other screen attributes. There are two syntax forms of
this function. Use syntax 1 to control screen display. Use syntax 2 to
control the display of the Info Window.

Syntax 1   Controls screen display

Syntax 2   Controls display of Info Window


DISPLAY Syntax 1

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

|                 |                    |
| --------------- | ------------------ |
| **Object\_num** | **Corresponds to** |
| 1 or omitted    | Show All           |
| 2               | Show Placeholders  |
| 3               | Hide               |

**Related Functions**

OPTIONS.VIEW   Controls display

WORKSPACE   Changes workspace settings

ZOOM   Enlarges or reduces a sheet in the active window

Syntax 2   Controls display of Info Window


DISPLAY Syntax 2

Equivalent to clicking the commands on the Info menu when the Info
Window is active. Controls which commands on the Info Window are in
effect. There are two syntax forms of this function. Use syntax 2 to
control the display of the Info Window. The Info Window must be active
to use this form of DISPLAY. This function is included for compatibility
with Microsoft Excel 95 or earlier; the Info Window is not available in
Microsoft Excel 97 or later.

Arguments in this syntax form correspond to commands on the Info menu
with the same names.

For these arguments:

  - > If the argument is TRUE, Microsoft Excel displays the
    > corresponding Info item.

  - > If the argument is FALSE, Microsoft Excel does not display the
    > corresponding Info item.

  - > If the argument is omitted, the status of the item is unchanged.

>  

**Syntax**

For controlling Info Window display

**DISPLAY**(cell, formula, value, format, protection, names, precedents,
dependents, note)

Cell    is a logical value that corresponds to the Cell command and
controls the display of cell information in the Info Window. If TRUE,
cell information will be displayed; if FALSE, cell information will not
be displayed.

Formula    is a logical value that corresponds to the Formula command
and controls the display of formula information in the Info Window. If
TRUE, formula information will be displayed; if FALSE, formula
information will not be displayed.

Value    is a logical value that corresponds to the Value command and
controls the display of value information in the Info Window. If TRUE,
value information will be displayed; if FALSE, value information will
not be displayed.

Format    is a logical value that corresponds to the Format command and
controls the display of format information in the Info Window. If TRUE,
format information will be displayed; if FALSE, format information will
not be displayed.

Protection    is a logical value that corresponds to the Protection
command and controls the display of protection information in the Info
Window. If TRUE, protection information will be displayed; if FALSE,
protection information will not be displayed.

Names    is a logical value that corresponds to the Names command and
controls the display of name information in the Info Window. If TRUE,
name information will be displayed; if FALSE, name information will not
be displayed.

Precedents    is a number from 1 to 3 that specifies which precedents to
list, according to the following table.

Dependents    is a number from 1 to 3 that specifies which dependents to
list, according to the following table.

|                              |             |
| ---------------------------- | ----------- |
| **Precedents or dependents** | **List**    |
| 0                            | None        |
| 1                            | Direct only |
| 2                            | All levels  |

Note    is a logical value that corresponds to the Note command and
controls the display of note information in the Info Window. If TRUE,
note information will be displayed; if FALSE, note information will not
be displayed.

**Related Functions**

SHOW.INFO   Controls the display of the Info Window

ZOOM   Enlarges or reduces a sheet in the active window

Syntax 1   Controls screen display


DISPLAY

Controls whether the screen displays formulas, gridlines, row and column
headings, and other screen attributes. There are two syntax forms of
this function. Use syntax 1 to control screen display. Use syntax 2 to
control the display of the Info Window.

Syntax 1   Controls screen display

Syntax 2   Controls display of Info Window


DISPLAY Syntax 1

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

|                 |                    |
| --------------- | ------------------ |
| **Object\_num** | **Corresponds to** |
| 1 or omitted    | Show All           |
| 2               | Show Placeholders  |
| 3               | Hide               |

**Related Functions**

OPTIONS.VIEW   Controls display

WORKSPACE   Changes workspace settings

ZOOM   Enlarges or reduces a sheet in the active window

Syntax 2   Controls display of Info Window


DISPLAY Syntax 2

Equivalent to clicking the commands on the Info menu when the Info
Window is active. Controls which commands on the Info Window are in
effect. There are two syntax forms of this function. Use syntax 2 to
control the display of the Info Window. The Info Window must be active
to use this form of DISPLAY. This function is included for compatibility
with Microsoft Excel 95 or earlier; the Info Window is not available in
Microsoft Excel 97 or later.

Arguments in this syntax form correspond to commands on the Info menu
with the same names.

For these arguments:

  - > If the argument is TRUE, Microsoft Excel displays the
    > corresponding Info item.

  - > If the argument is FALSE, Microsoft Excel does not display the
    > corresponding Info item.

  - > If the argument is omitted, the status of the item is unchanged.

>  

**Syntax**

For controlling Info Window display

**DISPLAY**(cell, formula, value, format, protection, names, precedents,
dependents, note)

Cell    is a logical value that corresponds to the Cell command and
controls the display of cell information in the Info Window. If TRUE,
cell information will be displayed; if FALSE, cell information will not
be displayed.

Formula    is a logical value that corresponds to the Formula command
and controls the display of formula information in the Info Window. If
TRUE, formula information will be displayed;DISPLAY

Controls whether the screen displays formulas, gridlines, row and column
headings, and other screen attributes. There are two syntax forms of
this function. Use syntax 1 to control screen display. Use syntax 2 to
control the display of the Info Window.

Syntax 1   Controls screen display

Syntax 2   Controls display of Info Window


DISPLAY Syntax 1

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

|                 |                    |
| --------------- | ------------------ |
| **Object\_num** | **Corresponds to** |
| 1 or omitted    | Show All           |
| 2               | Show Placeholders  |
| 3               | Hide               |

**Related Functions**

OPTIONS.VIEW   Controls display

WORKSPACE   Changes workspace settings

ZOOM   Enlarges or reduces a sheet in the active window

Syntax 2   Controls display of Info Window


DISPLAY Syntax 2

Equivalent to clicking the commands on the Info menu when the Info
Window is active. Controls which commands on the Info Window are in
effect. There are two syntax forms of this function. Use syntax 2 to
control the display of the Info Window. The Info Window must be active
to use this form of DISPLAY. This function is included for compatibility
with Microsoft Excel 95 or earlier; the Info Window is not available in
Microsoft Excel 97 or later.

Arguments in this syntax form correspond to commands on the Info menu
with the same names.

For these arguments:

  - > If the argument is TRUE, Microsoft Excel displays the
    > corresponding Info item.

  - > If the argument is FALSE, Microsoft Excel does not display the
    > corresponding Info item.

  - > If the argument is omitted, the status of the item is unchanged.

>  

**Syntax**

For controlling Info Window display

**DISPLAY**(cell, formula, value, format, protection, names, precedents,
dependents, note)

Cell    is a logical value that corresponds to the Cell command and
controls the display of cell information in the Info Window. If TRUE,
cell information will be displayed; if FALSE, cell information will not
be displayed.

Formula    is a logical value that corresponds to the Formula command
and controls the display of formula information in the Info Window. If
TRUE, formula information will be displayed;DISPLAY

Controls whether the screen displays formulas, gridlines, row and column
headings, and other screen attributes. There are two syntax forms of
this function. Use syntax 1 to control screen display. Use syntax 2 to
control the display of the Info Window.

Syntax 1   Controls screen display

Syntax 2   Controls display of Info Window


DISPLAY Syntax 1

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

|                 |                    |
| --------------- | ------------------ |
| **Object\_num** | **Corresponds to** |
| 1 or omitted    | Show All           |
| 2               | Show Placeholders  |
| 3               | Hide               |

**Related Functions**

OPTIONS.VIEW   Controls display

WORKSPACE   Changes workspace settings

ZOOM   Enlarges or reduces a sheet in the active window

Syntax 2   Controls display of Info Window


DISPLAY Syntax 2

Equivalent to clicking the commands on the Info menu when the Info
Window is active. Controls which commands on the Info Window are in
effect. There are two syntax forms of this function. Use syntax 2 to
control the display of the Info Window. The Info Window must be active
to use this form of DISPLAY. This function is included for compatibility
with Microsoft Excel 95 or earlier; the Info Window is not available in
Microsoft Excel 97 or later.

Arguments in this syntax form correspond to commands on the Info menu
with the same names.

For these arguments:

  - > If the argument is TRUE, Microsoft Excel displays the
    > corresponding Info item.

  - > If the argument is FALSE, Microsoft Excel does not display the
    > corresponding Info item.

  - > If the argument is omitted, the status of the item is unchanged.

>  

**Syntax**

For controlling Info Window display

**DISPLAY**(cell, formula, value, format, protection, names, precedents,
dependents, note)

Cell    is a logical value that corresponds to the Cell command and
controls the display of cell information in the Info Window. If TRUE,
cell information will be displayed; if FALSE, cell information will not
be displayed.

Formula    is a logical value that corresponds to the Formula command
and controls the display of formula information in the Info Window. If
TRUE, formula information will be displayed;DISPLAY

Controls whether the screen displays formulas, gridlines, row and column
headings, and other screen attributes. There are two syntax forms of
this function. Use syntax 1 to control screen display. Use syntax 2 to
control the display of the Info Window.

Syntax 1   Controls screen display

Syntax 2   Controls display of Info Window


DISPLAY Syntax 1

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

|                 |                    |
| --------------- | ------------------ |
| **Object\_num** | **Corresponds to** |
| 1 or omitted    | Show All           |
| 2               | Show Placeholders  |
| 3               | Hide               |

**Related Functions**

OPTIONS.VIEW   Controls display

WORKSPACE   Changes workspace settings

ZOOM   Enlarges or reduces a sheet in the active window

Syntax 2   Controls display of Info Window


DISPLAY Syntax 2

Equivalent to clicking the commands on the Info menu when the Info
Window is active. Controls which commands on the Info Window are in
effect. There are two syntax forms of this function. Use syntax 2 to
control the display of the Info Window. The Info Window must be active
to use this form of DISPLAY. This function is included for compatibility
with Microsoft Excel 95 or earlier; the Info Window is not available in
Microsoft Excel 97 or later.

Arguments in this syntax form correspond to commands on the Info menu
with the same names.

For these arguments:

  - > If the argument is TRUE, Microsoft Excel displays the
    > corresponding Info item.

  - > If the argument is FALSE, Microsoft Excel does not display the
    > corresponding Info item.

  - > If the argument is omitted, the status of the item is unchanged.

>  

**Syntax**

For controlling Info Window display

**DISPLAY**(cell, formula, value, format, protection, names, precedents,
dependents, note)

Cell    is a logical value that corresponds to the Cell command and
controls the display of cell information in the Info Window. If TRUE,
cell information will be displayed; if FALSE, cell information will not
be displayed.

Formula    is a logical value that corresponds to the Formula command
and controls the display of formula information in the Info Window. If
TRUE, formula information will be displayed;DISPLAY

Controls whether the screen displays formulas, gridlines, row and column
headings, and other screen attributes. There are two syntax forms of
this function. Use syntax 1 to control screen display. Use syntax 2 to
control the display of the Info Window.

Syntax 1   Controls screen display

Syntax 2   Controls display of Info Window


DISPLAY Syntax 1

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

|                 |                    |
| --------------- | ------------------ |
| **Object\_num** | **Corresponds to** |
| 1 or omitted    | Show All           |
| 2               | Show Placeholders  |
| 3               | Hide               |

**Related Functions**

OPTIONS.VIEW   Controls display

WORKSPACE   Changes workspace settings

ZOOM   Enlarges or reduces a sheet in the active window

Syntax 2   Controls display of Info Window


DISPLAY Syntax 2

Equivalent to clicking the commands on the Info menu when the Info
Window is active. Controls which commands on the Info Window are in
effect. There are two syntax forms of this function. Use syntax 2 to
control the display of the Info Window. The Info Window must be active
to use this form of DISPLAY. This function is included for compatibility
with Microsoft Excel 95 or earlier; the Info Window is not available in
Microsoft Excel 97 or later.

Arguments in this syntax form correspond to commands on the Info menu
with the same names.

For these arguments:

  - > If the argument is TRUE, Microsoft Excel displays the
    > corresponding Info item.

  - > If the argument is FALSE, Microsoft Excel does not display the
    > corresponding Info item.

  - > If the argument is omitted, the status of the item is unchanged.

>  

**Syntax**

For controlling Info Window display

**DISPLAY**(cell, formula, value, format, protection, names, precedents,
dependents, note)

Cell    is a logical value that corresponds to the Cell command and
controls the display of cell information in the Info Window. If TRUE,
cell information will be displayed; if FALSE, cell information will not
be displayed.

Formula    is a logical value that corresponds to the Formula command
and controls the display of formula information in the Info Window. If
TRUE, formula information will be displayed;DISPLAY

Controls whether the screen displays formulas, gridlines, row and column
headings, and other screen attributes. There are two syntax forms of
this function. Use syntax 1 to control screen display. Use syntax 2 to
control the display of the Info Window.

Syntax 1   Controls screen display

Syntax 2   Controls display of Info Window


DISPLAY Syntax 1

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

|                 |                    |
| --------------- | ------------------ |
| **Object\_num** | **Corresponds to** |
| 1 or omitted    | Show All           |
| 2               | Show Placeholders  |
| 3               | Hide               |

**Related Functions**

OPTIONS.VIEW   Controls display

WORKSPACE   Changes workspace settings

ZOOM   Enlarges or reduces a sheet in the active window

Syntax 2   Controls display of Info Window


DISPLAY Syntax 2

Equivalent to clicking the commands on the Info menu when the Info
Window is active. Controls which commands on the Info Window are in
effect. There are two syntax forms of this function. Use syntax 2 to
control the display of the Info Window. The Info Window must be active
to use this form of DISPLAY. This function is included for compatibility
with Microsoft Excel 95 or earlier; the Info Window is not available in
Microsoft Excel 97 or later.

Arguments in this syntax form correspond to commands on the Info menu
with the same names.

For these arguments:

  - > If the argument is TRUE, Microsoft Excel displays the
    > corresponding Info item.

  - > If the argument is FALSE, Microsoft Excel does not display the
    > corresponding Info item.

  - > If the argument is omitted, the status of the item is unchanged.

>  

**Syntax**

For controlling Info Window display

**DISPLAY**(cell, formula, value, format, protection, names, precedents,
dependents, note)

Cell    is a logical value that corresponds to the Cell command and
controls the display of cell information in the Info Window. If TRUE,
cell information will be displayed; if FALSE, cell information will not
be displayed.

Formula    is a logical value that corresponds to the Formula command
and controls the display of formula information in the Info Window. If
TRUE, formula information will be displayed;DISPLAY

Controls whether the screen displays formulas, gridlines, row and column
headings, and other screen attributes. There are two syntax forms of
this function. Use syntax 1 to control screen display. Use syntax 2 to
control the display of the Info Window.

Syntax 1   Controls screen display

Syntax 2   Controls display of Info Window


DISPLAY Syntax 1

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

|                 |                    |
| --------------- | ------------------ |
| **Object\_num** | **Corresponds to** |
| 1 or omitted    | Show All           |
| 2               | Show Placeholders  |
| 3               | Hide               |

**Related Functions**

[OPTIONS.VIEW   Controls](OPTIONS.VIEW   Controls.md) display

[WORKSPACE   Changes](WORKSPACE   Changes.md) workspace settings

[ZOOM   Enlarges](ZOOM   Enlarges.md) or reduces a sheet in the active window

[Syntax](Syntax.md) 2   Controls display of Info Window


DISPLAY Syntax 2

Equivalent to clicking the commands on the Info menu when the Info
Window is active. Controls which commands on the Info Window are in
effect. There are two syntax forms of this function. Use syntax 2 to
control the display of the Info Window. The Info Window must be active
to use this form of DISPLAY. This function is included for compatibility
with Microsoft Excel 95 or earlier; the Info Window is not available in
Microsoft Excel 97 or later.

Arguments in this syntax form correspond to commands on the Info menu
with the same names.

For these arguments:

  - > If the argument is TRUE, Microsoft Excel displays the
    > corresponding Info item.

  - > If the argument is FALSE, Microsoft Excel does not display the
    > corresponding Info item.

  - > If the argument is omitted, the status of the item is unchanged.

>  

**Syntax**

For controlling Info Window display

**DISPLAY**(cell, formula, value, format, protection, names, precedents,
dependents, note)

Cell    is a logical value that corresponds to the Cell command and
controls the display of cell information in the Info Window. If TRUE,
cell information will be displayed; if FALSE, cell information will not
be displayed.

Formula    is a logical value that corresponds to the Formula command
and controls thDISPLAY

Controls whether the screen displays formulas, gridlines, row and column
headings, and other screen attributes. There are two syntax forms of
this function. Use syntax 1 to control screen display. Use syntax 2 to
control the display of the Info Window.

Syntax 1   Controls screen display

Syntax 2   Controls display of Info Window


DISPLAY Syntax 1

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

|                 |                    |
| --------------- | ------------------ |
| **Object\_num** | **Corresponds to** |
| 1 or omitted    | Show All           |
| 2               | Show Placeholders  |
| 3               | Hide               |

**Related Functions**

[OPTIONS.VIEW   Controls](OPTIONS.VIEW   Controls.md) display

[WORKSPACE   Changes](WORKSPACE   Changes.md) workspace settings

[ZOOM   Enlarges](ZOOM   Enlarges.md) or reduces a sheet in the active window

[Syntax](Syntax.md) 2   Controls display of Info Window


DISPLAY Syntax 2

Equivalent to clicking the commands on the Info menu when the Info
Window is active. Controls which commands on the Info Window are in
effect. There are two syntax forms of this function. Use syntax 2 to
control the display of the Info Window. The Info Window must be active
to use this form of DISPLAY. This function is included for compatibility
with Microsoft Excel 95 or earlier; the Info Window is not available in
Microsoft Excel 97 or later.

Arguments in this syntax form correspond to commands on the Info menu
with the same names.

For these arguments:

  - > If the argument is TRUE, Microsoft Excel displays the
    > corresponding Info item.

  - > If the argument is FALSE, Microsoft Excel does not display the
    > corresponding Info item.

  - > If the argument is omitted, the status of the item is unchanged.

>  

**Syntax**

For controlling Info Window display

**DISPLAY**(cell, formula, value, format, protection, names, precedents,
dependents, note)

Cell    is a logical value that corresponds to the Cell command and
controls the display of cell information in the Info Window. If TRUE,
cell information will be displayed; if FALSE, cell information will not
be displayed.

Formula    is a logical value that corresponds to the Formula command
and controls thDISPLAY

Controls whether the screen displays formulas, gridlines, row and column
headings, and other screen attributes. There are two syntax forms of
this function. Use syntax 1 to control screen display. Use syntax 2 to
control the display of the Info Window.

Syntax 1   Controls screen display

Syntax 2   Controls display of Info Window


DISPLAY Syntax 1

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

|                 |                    |
| --------------- | ------------------ |
| **Object\_num** | **Corresponds to** |
| 1 or omitted    | Show All           |
| 2               | Show Placeholders  |
| 3               | Hide               |

**Related Functions**

[OPTIONS](OPTIONS.md).VIEW   Controls display

[WORKSPACE](WORKSPACE.md)   Changes workspace settings

[ZOOM](ZOOM.md)   Enlarges or reduces a sheet in the active window

[S](S.md)yntax 2   Controls display of Info Window


DISPLAY Syntax 2

Equivalent to clicking the commands on the Info menu when the Info
Window is active. Controls which commands on the Info Window are in
effect. There are two syntax forms of this function. Use syntax 2 to
control the display of the Info Window. The Info Window must be active
to use this form of DISPLAY. This function is included for compatibility
with Microsoft Excel 95 or earlier; the Info Window is not available in
Microsoft Excel 97 or later.

Arguments in this syntax form correspond to commands on the Info menu
with the same names.

For these arguments:

  - > If the argument is TRUE, Microsoft Excel displays the
    > corresponding Info item.

  - > If the argument is FALSE, Microsoft Excel does not display the
    > corresponding Info item.

  - > If the argument is omitted, the status of the item is unchanged.

>  

**Syntax**

For controlling Info Window display

**DISPLAY**(cell, formula, value, format, protection, names, precedents,
dependents, note)

Cell    is a logical value that corresponds to the Cell command and
controls the display of cell information in the Info Window. If TRUE,
cell information will be displayed; if FALSE, cell information will not
be displayed.

Formula    is a logical value that corresponds to the Formula command
and controls the display of formula information in the Info WindowDISPLAY

Controls whether the screen displays formulas, gridlines, row and column
headings, and other screen attributes. There are two syntax forms of
this function. Use syntax 1 to control screen display. Use syntax 2 to
control the display of the Info Window.

Syntax 1   Controls screen display

Syntax 2   Controls display of Info Window


DISPLAY Syntax 1

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

|                 |                    |
| --------------- | ------------------ |
| **Object\_num** | **Corresponds to** |
| 1 or omitted    | Show All           |
| 2               | Show Placeholders  |
| 3               | Hide               |

**Related Functions**

[OPTIONS.VIEW](OPTIONS.VIEW.md)   Controls display

[WORKSPACE](WORKSPACE.md)   Changes workspace settings

[ZOOM](ZOOM.md)   Enlarges or reduces a sheet in the active window

[S](S.md)yntax 2   Controls display of Info Window


DISPLAY Syntax 2

Equivalent to clicking the commands on the Info menu when the Info
Window is active. Controls which commands on the Info Window are in
effect. There are two syntax forms of this function. Use syntax 2 to
control the display of the Info Window. The Info Window must be active
to use this form of DISPLAY. This function is included for compatibility
with Microsoft Excel 95 or earlier; the Info Window is not available in
Microsoft Excel 97 or later.

Arguments in this syntax form correspond to commands on the Info menu
with the same names.

For these arguments:

  - > If the argument is TRUE, Microsoft Excel displays the
    > corresponding Info item.

  - > If the argument is FALSE, Microsoft Excel does not display the
    > corresponding Info item.

  - > If the argument is omitted, the status of the item is unchanged.

>  

**Syntax**

For controlling Info Window display

**DISPLAY**(cell, formula, value, format, protection, names, precedents,
dependents, note)

Cell    is a logical value that corresponds to the Cell command and
controls the display of cell information in the Info Window. If TRUE,
cell information will be displayed; if FALSE, cell information will not
be displayed.

Formula    is a logical value that corresponds to the Formula command
and controls the display of formula information in the Info WDISPLAY

Controls whether the screen displays formulas, gridlines, row and column
headings, and other screen attributes. There are two syntax forms of
this function. Use syntax 1 to control screen display. Use syntax 2 to
control the display of the Info Window.

Syntax 1   Controls screen display

Syntax 2   Controls display of Info Window


DISPLAY Syntax 1

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

|                 |                    |
| --------------- | ------------------ |
| **Object\_num** | **Corresponds to** |
| 1 or omitted    | Show All           |
| 2               | Show Placeholders  |
| 3               | Hide               |

**Related Functions**

[OPTIONS.VIEW](OPTIONS.VIEW.md)   Controls display

[WORKSPACE](WORKSPACE.md)   Changes workspace settings

[ZOOM](ZOOM.md)   Enlarges or reduces a sheet in the active window

[S](S.md)yntax 2   Controls display of Info Window


DISPLAY Syntax 2

Equivalent to clicking the commands on the Info menu when the Info
Window is active. Controls which commands on the Info Window are in
effect. There are two syntax forms of this function. Use syntax 2 to
control the display of the Info Window. The Info Window must be active
to use this form of DISPLAY. This function is included for compatibility
with Microsoft Excel 95 or earlier; the Info Window is not available in
Microsoft Excel 97 or later.

Arguments in this syntax form correspond to commands on the Info menu
with the same names.

For these arguments:

  - > If the argument is TRUE, Microsoft Excel displays the
    > corresponding Info item.

  - > If the argument is FALSE, Microsoft Excel does not display the
    > corresponding Info item.

  - > If the argument is omitted, the status of the item is unchanged.

>  

**Syntax**

For controlling Info Window display

**DISPLAY**(cell, formula, value, format, protection, names, precedents,
dependents, note)

Cell    is a logical value that corresponds to the Cell command and
controls the display of cell information in the Info Window. If TRUE,
cell information will be displayed; if FALSE, cell information will not
be displayed.

Formula    is a logical value that corresponds to the Formula command
and controls the display of formula information in the Info Window. If
TRUE, formula information will be displayed; if FALSE, formula
information will not be displayed.

Value    is a logical value that corresponds to the Value command and
controls the display of value information in the Info Window. If TRUE,
value information will be displayed; if FALSE, value information will
not be displayed.

Format    is a logical value that corresponds to the Format command and
controls the display of format information in the Info Window. If TRUE,
format information will be displayed; if FALSE, format information will
not be displayed.

Protection    is a logical value that corresponds to the Protection
command and controls the display of protection information in the Info
Window. If TRUE, protection information will be displayed; if FALSE,
protection information will not be displayed.

Names    is a logical value that corresponds to the Names command and
controls the display of name information in the Info Window. If TRUE,
name information will be displayed; if FALSE, name information will not
be displayed.

Precedents    is a number from 1 to 3 that specifies which precedents to
list, according to the following table.

Dependents    is a number from 1 to 3 that specifies which dependents to
list, according to the following table.

|                              |             |
| ---------------------------- | ----------- |
| **Precedents or dependents** | **List**    |
| 0                            | None        |
| 1                            | Direct only |
| 2                            | All levels  |

Note    is a logical value that corresponds to the Note command and
controls the display of note information in the Info Window. If TRUE,
note information will be displayed; if FALSE, note information will not
be displayed.

**Related Functions**

[SHOW.INFO](SHOW.INFO.md)   Controls the display of the Info Window

[ZOOM](ZOOM.md)   Enlarges or reduces a sheet in the active window

[S](S.md)yntax 1   Controls screen display


