DISPLAY Syntax 2
================

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

-   If the argument is TRUE, Microsoft Excel displays the corresponding
    > Info item.

-   If the argument is FALSE, Microsoft Excel does not display the
    > corresponding Info item.

-   If the argument is omitted, the status of the item is unchanged.

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

  ------------------------------ -------------
  **Precedents or dependents**   **List**
  0                              None
  1                              Direct only
  2                              All levels
  ------------------------------ -------------

Note    is a logical value that corresponds to the Note command and
controls the display of note information in the Info Window. If TRUE,
note information will be displayed; if FALSE, note information will not
be displayed.

**Related Functions**

SHOW.INFO   Controls the display of the Info Window

ZOOM   Enlarges or reduces a sheet in the active window

Syntax 1   Controls screen display

Return to [top](#A)

DOCUMENTS
