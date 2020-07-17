FORMAT.AUTO
===========

Equivalent to clicking the AutoFormat command on the Format menu when a
worksheet is active or clicking the AutoFormat button. Formats the
selected range of cells from a built-in gallery of formats.

**Syntax**

**FORMAT.AUTO**(format\_num, number, font, alignment, border, pattern,
width)

**FORMAT.AUTO**?(format\_num, number, font, alignment, border, pattern,
width)

Format\_num    is a number from 1 to 17 corresponding to the formats in
the Table Format list box in the AutoFormat dialog box.

  ----------------- -----------------------------------------------------
  **Format\_num**   **Table Format**
  0                 None
  1 or omitted      Classic 1
  2                 Classic 2
  3                 Classic 3
  4                 Accounting 1
  5                 Accounting 2
  6                 Accounting 3
  7                 Colorful 1
  8                 Colorful 2
  9                 Colorful 3
  10                List 1
  11                List 2
  12                List 3
  13                3D Effects 1
  14                3D Effects 2
  15                Japan 1 (Far East versions of Microsoft Excel only)
  16                Japan 2 (Far East versions of Microsoft Excel only)
  17                Accounting 4
  18                Simple
  ----------------- -----------------------------------------------------

The following arguments are logical values corresponding to the Formats
To Apply check boxes in the AutoFormat dialog box. If an argument is
TRUE or omitted, Microsoft Excel selects the check box; if FALSE,
Microsoft Excel clears the check box.

Number    corresponds to the Number check box.

Font    corresponds to the Font check box.

Alignment    corresponds to the Alignment check box.

Border    corresponds to the Border check box.

Pattern    corresponds to the Pattern check box.

Width    corresponds to the Column Width/Row Height check box.

**Related Functions**

ALIGNMENT   Aligns or wraps text in cells

BORDER   Adds a border to the selected cell or object

FONT.PROPERTIES   Applies a font to the selection

FORMAT.NUMBER   Applies a number format to the selection

PATTERNS   Changes the appearance of the selected object

Return to [top](#E)

FORMAT.CHART
