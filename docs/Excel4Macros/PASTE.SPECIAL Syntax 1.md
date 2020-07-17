PASTE.SPECIAL Syntax 1
======================

Equivalent to clicking the Paste Special command on the Edit menu.
Pastes the specified components from the copy area into the current
selection. The PASTE.SPECIAL function has four syntax forms. Use syntax
1 if you are pasting into a sheet or macro sheet.

**Syntax**

**PASTE.SPECIAL**(paste\_num, operation\_num, skip\_blanks, transpose)

**PASTE.SPECIAL**?(paste\_num, operation\_num, skip\_blanks, transpose)

Paste\_num    is a number from 1 to 6 specifying what to paste.
Paste\_num can also be quoted text of the object you want to paste.

  ---------------- --------------------
  **Paste\_num**   **Pastes**
  1                All
  2                Formulas
  3                Values
  4                Formats
  5                Comments
  6                All except borders
  ---------------- --------------------

Operation\_num    is a number from 1 to 5 specifying which operation to
perform when pasting.

  -------------------- ------------
  **Operation\_num**   **Action**
  1                    None
  2                    Add
  3                    Subtract
  4                    Multiply
  5                    Divide
  -------------------- ------------

Skip\_blanks    is a logical value corresponding to the Skip Blanks
check box in the Paste Special dialog box.

-   If skip\_blanks is TRUE, Microsoft Excel skips blanks in the copy
    > area when pasting.

-   If skip\_blanks is FALSE, Microsoft Excel pastes normally.

>  

Transpose    is a logical value corresponding to the Transpose check box
in the Paste Special dialog box.

-   If transpose is TRUE, Microsoft Excel transposes rows and columns
    > when pasting.

-   If transpose is FALSE, Microsoft Excel pastes normally.

>  

**Related Functions**

FORMULA   Enters values into a cell or range or onto a chart

PASTE   Pastes cut or copied data

PASTE.LINK   Pastes copied data and establishes a link to its source

Syntax 2   Copying from a sheet and pasting into a chart.

Syntax 3   Copying and pasting between charts

Syntax 4   Pasting information from another application.

Return to [top](#H)

PASTE.SPECIAL Syntax 2
