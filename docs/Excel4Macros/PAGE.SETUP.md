# PAGE.SETUP

Equivalent to clicking the Page Setup command on the File menu. Use
PAGE.SETUP to control the printed appearance of your sheets.

There are three syntax forms of PAGE.SETUP. Syntax 1 applies if a sheet
or macro sheet is active; syntax 2 applies if a chart is active; syntax
three applies to Visual Basic modules and the info Window.

Arguments correspond to check boxes and text boxes in the Page Setup
dialog box. Arguments that correspond to check boxes are logical values.
If an argument is TRUE, Microsoft Excel selects the check box; if FALSE,
Microsoft Excel clears the check box. Arguments for margins are always
in inches, regardless of your country setting.

**Syntax 1**

Worksheets and macro sheets

**PAGE.SETUP**(head, foot, left, right, top, bot, hdng, grid, h\_cntr,
v\_cntr, orient, paper\_size, scale, pg\_num, pg\_order, bw\_cells,
quality, head\_margin, foot\_margin, notes, draft)

**PAGE.SETUP**?(head, foot, left, right, top, bot, hdng, grid, h\_cntr,
v\_cntr, orient, paper\_size, scale, pg\_num, pg\_order, bw\_cells,
quality, head\_margin, foot\_margin, notes, draft)

**Syntax 2**

Charts

**PAGE.SETUP**(head, foot, left, right, top, bot, size, h\_cntr,
v\_cntr, orient, paper\_size, scale, pg\_num, bw\_chart, quality,
head\_margin, foot\_margin, draft)

**PAGE.SETUP**?(head, foot, left, right, top, bot, size, h\_cntr,
v\_cntr, orient, paper\_size, scale, pg\_num, bw\_chart, quality,
head\_margin, foot\_margin, draft)

**Syntax 3**

Visual Basic Modules and the Info Window

**PAGE.SETUP**(head, foot, left, right, top, bot, orient, paper\_size,
scale, quality, head\_margin, foot\_margin, pg\_num)

**PAGE.SETUP**?(head, foot, left, right, top, bot, orient, paper\_size,
scale, quality, head\_margin, foot\_margin, pg\_num)

Head&nbsp;&nbsp;&nbsp;&nbsp;specifies the text and formatting codes for
the header for the current sheet . For information about formatting
codes, see "Remarks" later in this topic.

Foot&nbsp;&nbsp;&nbsp;&nbsp;specifies the text and formatting codes for
the workbook footer.

Left&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Left box and is a number
specifying the left margin.

Right&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Right box and is a
number specifying the right margin.

Top&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Top box and is a number
specifying the top margin.

Bot&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Bottom box and is a number
specifying the bottom margin.

Hdng&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Row & Column Headings
check box. Hdng is available only in the sheet and macro sheet form of
the function.

Grid&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Cell Gridlines check box.
Grid is available only in the sheet and macro sheet form of the
function.

H\_cntr&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Center Horizontally
check box in the Margins panel of the Page Setup dialog box.

V\_cntr&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Center Vertically
check box in the Margins panel of the Page Setup dialog box.

Orient&nbsp;&nbsp;&nbsp;&nbsp;determines the direction in which your
workbook is printed.

|            |                  |
| ---------- | ---------------- |
| **Orient** | **Print format** |
| 1          | Portrait         |
| 2          | Landscape        |

Paper\_size&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 26 that
specifies the size of the paper.

|                 |                |
| --------------- | -------------- |
| **Paper\_size** | **Paper type** |
| 1               | Letter         |
| 2               | Letter (small) |
| 3               | Tabloid        |
| 4               | Ledger         |
| 5               | Legal          |
| 6               | Statement      |
| 7               | Executive      |
| 8               | A3             |
| 9               | A4             |
| 10              | A4 (small)     |
| 11              | A5             |
| 12              | B4             |
| 13              | B5             |
| 14              | Folio          |
| 15              | Quarto         |
| 16              | 10x14          |
| 17              | 11x17          |
| 18              | Note           |
| 19              | ENV9           |
| 20              | ENV10          |
| 21              | ENV11          |
| 22              | ENV12          |
| 23              | ENV14          |
| 24              | C Sheet        |
| 25              | D Sheet        |
| 26              | E Sheet        |

Scale&nbsp;&nbsp;&nbsp;&nbsp;is a number representing the percentage to
increase or decrease the size of the sheet. All scaling retains the
aspect ratio of the original.

  - > To specify a percentage of reduction or enlargement, set scale to
    > the percentage.

  - > For worksheets and macros, you can specify the number of pages
    > that the printout should be scaled to fit. Set scale to a two-item
    > horizontal array, with the first item equal to the width and the
    > second item equal to the height. If no constraint is necessary in
    > one direction, you can set the corresponding value to \#N/A.

  - > Scale can also be a logical value. To fit the print area on a
    > single page, set scale to TRUE.


Pg\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies the number of the first page.
If zero, sets first page to zero. If "Auto" is used, then the page
numbering is set to automatic. If omitted, PAGE.SETUP retains the
existing pg\_num.

Pg\_order&nbsp;&nbsp;&nbsp;&nbsp;specifies whether pagination is
left-to-right and then down, or top-to-bottom and then right.

|               |                           |
| ------------- | ------------------------- |
| **Pg\_order** | **Pagination**            |
| 1             | Top-to-bottom, then right |
| 2             | Left-to-right, then down  |

Bw\_cells&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that specifies
whether to print cells and all graphic objects, such as text boxes and
buttons, in color.

  - > If bw\_cells is TRUE, Microsoft Excel prints cell text and borders
    > in black and cell backgrounds in white.

  - > If bw\_cells is FALSE , Microsoft Excel prints cell text, borders,
    > and background patterns in color (or in gray scale).


Bw\_chart&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that specifies
whether to print chart in color.

Size&nbsp;&nbsp;&nbsp;&nbsp;is a number corresponding to the options in
the Chart Size box, and determines how you want the chart printed on the
page within the margins. Size is available only in the chart form of the
function.

|          |                             |
| -------- | --------------------------- |
| **Size** | **Size to print the chart** |
| 1        | Screen size                 |
| 2        | Fit to page                 |
| 3        | Full page                   |

Quality&nbsp;&nbsp;&nbsp;&nbsp;specifies the print quality in
dots-per-inch. To specify both horizontal and vertical print quality,
use an array of two values.

Head\_margin&nbsp;&nbsp;&nbsp;&nbsp;is the placement, in inches, of the
running head margin from the edge of the page.

Foot\_margin&nbsp;&nbsp;&nbsp;&nbsp;is the placement, in inches, of the
running foot margin from the edge of the page.

Draft&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Draft Quality checkbox
in the Sheet tab and in the Chart tab of the Page Setup dialog box. If
FALSE or omitted, graphics are printed with the sheet. If TRUE, no
graphics are printed.

Notes&nbsp;&nbsp;&nbsp;&nbsp;specifies whether to print cell notes with
the sheet. If TRUE, both the sheet and the cell notes are printed. If
FALSE or omitted, just the sheet is printed.

**Remarks**

Microsoft Excel no longer requires you to enter formatting codes to
format headers and footers, but the codes are still supported and
recorded by the macro recorder. You can include these codes as part of
the head and foot text strings to align portions of the header or footer
to the left, right, or center; to include the page number, date, time,
or workbook name; and to print the header or footer in bold or italic.

|                         |                                                                                                                                                                                             |
| ----------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Formatting code**     | **Result**                                                                                                                                                                                  |
| \&L                     | Left-aligns the characters that follow.                                                                                                                                                     |
| \&C                     | Centers the characters that follow.                                                                                                                                                         |
| \&R                     | Right-aligns the characters that follow.                                                                                                                                                    |
| \&B                     | Turns bold printing on or off (now obsolete).                                                                                                                                               |
| \&I                     | Turns italic printing on or off.                                                                                                                                                            |
| \&U                     | Turns single underlining printing on or off.                                                                                                                                                |
| \&S                     | Turns strikethrough printing on or off.                                                                                                                                                     |
| \&O                     | Turns outline printing on or off (Macintosh only).                                                                                                                                          |
| \&H                     | Turns shadow printing on or off (Macintosh only).                                                                                                                                           |
| \&D                     | Prints the current date.                                                                                                                                                                    |
| \&T                     | Prints the current time.                                                                                                                                                                    |
| \&A                     | Prints the name of the sheet                                                                                                                                                                |
| \&F                     | Prints the name of the workbook.                                                                                                                                                            |
| \&P                     | Prints the page number.                                                                                                                                                                     |
| \&P+number              | Prints the page number plus number.                                                                                                                                                         |
| \&P-number              | Prints the page number minus number.                                                                                                                                                        |
| &&                      | Prints a single ampersand.                                                                                                                                                                  |
| & "fontname, fontstyle" | Prints the characters that follow in the specified font and style. Be sure to include a comma immediately following the fontname, and double quotation marks around fontname and fontstyle. |
| \&nn                    | Prints the characters that follow in the specified font size. Use a two-digit number to specify a size in points.                                                                           |
| \&N                     | Prints the total number of pages in the workbook.                                                                                                                                           |
| \&E                     | Prints a double underline                                                                                                                                                                   |
| \&X                     | Prints the character as superscript                                                                                                                                                         |
| \&Y                     | Prints the character as subscript                                                                                                                                                           |

**Related Functions**

[DISPLAY](DISPLAY.md)&nbsp;&nbsp;&nbsp;Controls screen and Info Window display

[GET.DOCUMENT](GET.DOCUMENT.md)&nbsp;&nbsp;&nbsp;Returns information about a workbook

[PRINT](PRINT.md)&nbsp;&nbsp;&nbsp;Prints the active workbook

[WORKSPACE](WORKSPACE.md)&nbsp;&nbsp;&nbsp;Changes workspace settings



Return to [README](README.md#P)

# PAGE.SETUP

Equivalent to clicking the Page Setup command on the File menu. Use
PAGE.SETUP to control the printed appearance of your sheets.

There are three syntax forms of PAGE.SETUP. Syntax 1 applies if a sheet
or macro sheet is active; syntax 2 applies if a chart is active; syntax
three applies to Visual Basic modules and the info Window.

Arguments correspond to check boxes and text boxes in the Page Setup
dialog box. Arguments that correspond to check boxes are logical values.
If an argument is TRUE, Microsoft Excel selects the check box; if FALSE,
Microsoft Excel clears the check box. Arguments for margins are always
in inches, regardless of your country setting.

**Syntax 1**

Worksheets and macro sheets

**PAGE.SETUP**(head, foot, left, right, top, bot, hdng, grid, h\_cntr,
v\_cntr, orient, paper\_size, scale, pg\_num, pg\_order, bw\_cells,
quality, head\_margin, foot\_margin, notes, draft)

**PAGE.SETUP**?(head, foot, left, right, top, bot, hdng, grid, h\_cntr,
v\_cntr, orient, paper\_size, scale, pg\_num, pg\_order, bw\_cells,
quality, head\_margin, foot\_margin, notes, draft)

**Syntax 2**

Charts

**PAGE.SETUP**(head, foot, left, right, top, bot, size, h\_cntr,
v\_cntr, orient, paper\_size, scale, pg\_num, bw\_chart, quality,
head\_margin, foot\_margin, draft)

**PAGE.SETUP**?(head, foot, left, right, top, bot, size, h\_cntr,
v\_cntr, orient, paper\_size, scale, pg\_num, bw\_chart, quality,
head\_margin, foot\_margin, draft)

**Syntax 3**

Visual Basic Modules and the Info Window

**PAGE.SETUP**(head, foot, left, right, top, bot, orient, paper\_size,
scale, quality, head\_margin, foot\_margin, pg\_num)

**PAGE.SETUP**?(head, foot, left, right, top, bot, orient, paper\_size,
scale, quality, head\_margin, foot\_margin, pg\_num)

Head&nbsp;&nbsp;&nbsp;&nbsp;specifies the text and formatting codes for
the header for the current sheet . For information about formatting
codes, see "Remarks" later in this topic.

Foot&nbsp;&nbsp;&nbsp;&nbsp;specifies the text and formatting codes for
the workbook footer.

Left&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Left box and is a number
specifying the left margin.

Right&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Right box and is a
number specifying the right margin.

Top&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Top box and is a number
specifying the top margin.

Bot&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Bottom box and is a number
specifying the bottom margin.

Hdng&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Row & Column Headings
check box. Hdng is available only in the sheet and macro sheet form of
the function.

Grid&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Cell Gridlines check box.
Grid is available only in the sheet and macro sheet form of the
function.

H\_cntr&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Center Horizontally
check box in the Margins panel of the Page Setup dialog box.

V\_cntr&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Center Vertically
check box in the Margins panel of the Page Setup dialog box.

Orient&nbsp;&nbsp;&nbsp;&nbsp;determines the direction in which your
workbook is printed.

|            |                  |
| ---------- | ---------------- |
| **Orient** | **Print format** |
| 1          | Portrait         |
| 2          | Landscape        |

Paper\_size&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 26 that
specifies the size of the paper.

|                 |                |
| --------------- | -------------- |
| **Paper\_size** | **Paper type** |
| 1               | Letter         |
| 2               | Letter (small) |
| 3               | Tabloid        |
| 4               | Ledger         |
| 5               | Legal          |
| 6               | Statement      |
| 7               | Executive      |
| 8               | A3             |
| 9               | A4             |
| 10              | A4 (small)     |
| 11              | A5             |
| 12              | B4             |
| 13              | B5             |
| 14              | Folio          |
| 15              | Quarto         |
| 16              | 10x14          |
| 17              | 11x17          |
| 18              | Note           |
| 19              | ENV9           |
| 20              | ENV10          |
| 21              | ENV11          |
| 22              | ENV12          |
| 23              | ENV14          |
| 24              | C Sheet        |
| 25              | D Sheet        |
| 26              | E Sheet        |

Scale&nbsp;&nbsp;&nbsp;&nbsp;is a number representing the percentage to
increase or decrease the size of the sheet. All scaling retains the
aspect ratio of the original.

  - > To specify a percentage of reduction or enlargement, set scale to
    > the percentage.

  - > For worksheets and macros, you can specify the number of pages
    > that the printout should be scaled to fit. Set scale to a two-item
    > horizontal array, with the first item equal to the width and the
    > second item equal to the height. If no constraint is necessary in
    > one direction, you can set the corresponding value to \#N/A.

  - > Scale can also be a logical value. To fit the print area on a
    > single page, set scale to TRUE.


Pg\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies the number of the first page.
If zero, sets first page to zero. If "Auto" is used, then the page
numbering is set to automatic. If omitted, PAGE.SETUP retains the
existing pg\_num.

Pg\_order&nbsp;&nbsp;&nbsp;&nbsp;specifies whether pagination is
left-to-right and then down, or top-to-bottom and then right.

|               |                           |
| ------------- | ------------------------- |
| **Pg\_order** | **Pagination**            |
| 1             | Top-to-bottom, then right |
| 2             | Left-to-right, then down  |

Bw\_cells&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that specifies
whether to print cells and all graphic objects, such as text boxes and
buttons, in color.

  - > If bw\_cells is TRUE, Microsoft Excel prints cell text and borders
    > in black and cell backgrounds in white.

  - > If bw\_cells is FALSE , Microsoft Excel prints cell text, borders,
    > and background patterns in color (or in gray scale).


Bw\_chart&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that specifies
whether to print chart in color.

Size&nbsp;&nbsp;&nbsp;&nbsp;is a number corresponding to the options in
the Chart Size box, and determines how you want the chart printed on the
page within the margins. Size is available only in the chart form of the
function.

|          |                             |
| -------- | --------------------------- |
| **Size** | **Size to print the chart** |
| 1        | Screen size                 |
| 2        | Fit to page                 |
| 3        | Full page                   |

Quality&nbsp;&nbsp;&nbsp;&nbsp;specifies the print quality in
dots-per-inch. To specify both horizontal and vertical print quality,
use an array of two values.

Head\_margin&nbsp;&nbsp;&nbsp;&nbsp;is the placement, in inches, of the
running head margin from the edge of the page.

Foot\_margin&nbsp;&nbsp;&nbsp;&nbsp;is the placement, in inches, of the
running foot margin from the edge of the page.

Draft&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Draft Quality checkbox
in the Sheet tab and in the Chart tab of the Page Setup dialog box. If
FALSE or omitted, graphics are printed with the sheet. If TRUE, no
graphics are printed.

Notes&nbsp;&nbsp;&nbsp;&nbsp;specifies whether to print cell notes with
the sheet. If TRUE, both the sheet and the cell notes are printed. If
FALSE or omitted, just the sheet is printed.

**Remarks**

Microsoft Excel no longer requires you to enter formatting codes to
format headers and footers, but the codes are still supported and
recorded by the macro recorder. You can include these codes as part of
the head and foot text strings to align portions of the header or footer
to the left, right, or center; to include the page number, date, time,
or workbook name; and to print the header or footer in bold or italic.

|                         |                                                                                                                                                                                             |
| ----------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Formatting code**     | **Result**                                                                                                                                                                                  |
| \&L                     | Left-aligns the characters that follow.                                                                                                                                                     |
| \&C                     | Centers the characters that follow.                                                                                                                                                         |
| \&R                     | Right-aligns the characters that follow.                                                                                                                                                    |
| \&B                     | Turns bold printing on or off (now obsolete).                                                                                                                                               |
| \&I                     | Turns italic printing on or off.                                                                                                                                                            |
| \&U                     | Turns single underlining printing on or off.                                                                                                                                                |
| \&S                     | Turns strikethrough printing on or off.                                                                                                                                                     |
| \&O                     | Turns outline printing on or off (Macintosh only).                                                                                                                                          |
| \&H                     | Turns shadow printing on or off (Macintosh only).                                                                                                                                           |
| \&D                     | Prints the current date.                                                                                                                                                                    |
| \&T                     | Prints the current time.                                                                                                                                                                    |
| \&A                     | Prints the name of the sheet                                                                                                                                                                |
| \&F                     | Prints the name of the workbook.                                                                                                                                                            |
| \&P                     | Prints the page number.                                                                                                                                                                     |
| \&P+number              | Prints the page number plus number.                                                                                                                                                         |
| \&P-number              | Prints the page number minus number.                                                                                                                                                        |
| &&                      | Prints a single ampersand.                                                                                                                                                                  |
| & "fontname, fontstyle" | Prints the characters that follow in the specified font and style. Be sure to include a comma immediately following the fontname, and double quotation marks around fontname and fontstyle. |
| \&nn                    | Prints the characters that follow in the specified font size. Use a two-digit number to specify a size in points.                                                                           |
| \&N                     | Prints the total number of pages in the workbook.                                                                                                                                           |
| \&E                     | Prints a double underline                                                                                                                                                                   |
| \&X                     | Prints the character as superscript                                                                                                                                                         |
| \&Y                     | Prints the character as subscript                                                                                                                                                           |

**Related Functions**

[DISPLAY](DISPLAY.md)&nbsp;&nbsp;&nbsp;Controls screen and Info Window display

[GET.DOCUMENT](GET.DOCUMENT.md)&nbsp;&nbsp;&nbsp;Returns information about a workbook

[PRINT](PRINT.md)&nbsp;&nbsp;&nbsp;Prints the active workbook

[WORKSPACE](WORKSPACE.md)&nbsp;&nbsp;&nbsp;Changes workspace settings



Return to [README](README.md#P)

# PAGE.SETUP

Equivalent to clicking the Page Setup command on the File menu. Use
PAGE.SETUP to control the printed appearance of your sheets.

There are three syntax forms of PAGE.SETUP. Syntax 1 applies if a sheet
or macro sheet is active; syntax 2 applies if a chart is active; syntax
three applies to Visual Basic modules and the info Window.

Arguments correspond to check boxes and text boxes in the Page Setup
dialog box. Arguments that correspond to check boxes are logical values.
If an argument is TRUE, Microsoft Excel selects the check box; if FALSE,
Microsoft Excel clears the check box. Arguments for margins are always
in inches, regardless of your country setting.

**Syntax 1**

Worksheets and macro sheets

**PAGE.SETUP**(head, foot, left, right, top, bot, hdng, grid, h\_cntr,
v\_cntr, orient, paper\_size, scale, pg\_num, pg\_order, bw\_cells,
quality, head\_margin, foot\_margin, notes, draft)

**PAGE.SETUP**?(head, foot, left, right, top, bot, hdng, grid, h\_cntr,
v\_cntr, orient, paper\_size, scale, pg\_num, pg\_order, bw\_cells,
quality, head\_margin, foot\_margin, notes, draft)

**Syntax 2**

Charts

**PAGE.SETUP**(head, foot, left, right, top, bot, size, h\_cntr,
v\_cntr, orient, paper\_size, scale, pg\_num, bw\_chart, quality,
head\_margin, foot\_margin, draft)

**PAGE.SETUP**?(head, foot, left, right, top, bot, size, h\_cntr,
v\_cntr, orient, paper\_size, scale, pg\_num, bw\_chart, quality,
head\_margin, foot\_margin, draft)

**Syntax 3**

Visual Basic Modules and the Info Window

**PAGE.SETUP**(head, foot, left, right, top, bot, orient, paper\_size,
scale, quality, head\_margin, foot\_margin, pg\_num)

**PAGE.SETUP**?(head, foot, left, right, top, bot, orient, paper\_size,
scale, quality, head\_margin, foot\_margin, pg\_num)

Head&nbsp;&nbsp;&nbsp;&nbsp;specifies the text and formatting codes for
the header for the current sheet . For information about formatting
codes, see "Remarks" later in this topic.

Foot&nbsp;&nbsp;&nbsp;&nbsp;specifies the text and formatting codes for
the workbook footer.

Left&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Left box and is a number
specifying the left margin.

Right&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Right box and is a
number specifying the right margin.

Top&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Top box and is a number
specifying the top margin.

Bot&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Bottom box and is a number
specifying the bottom margin.

Hdng&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Row & Column Headings
check box. Hdng is available only in the sheet and macro sheet form of
the function.

Grid&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Cell Gridlines check box.
Grid is available only in the sheet and macro sheet form of the
function.

H\_cntr&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Center Horizontally
check box in the Margins panel of the Page Setup dialog box.

V\_cntr&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Center Vertically
check box in the Margins panel of the Page Setup dialog box.

Orient&nbsp;&nbsp;&nbsp;&nbsp;determines the direction in which your
workbook is printed.

|            |                  |
| ---------- | ---------------- |
| **Orient** | **Print format** |
| 1          | Portrait         |
| 2          | Landscape        |

Paper\_size&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 26 that
specifies the size of the paper.

|                 |                |
| --------------- | -------------- |
| **Paper\_size** | **Paper type** |
| 1               | Letter         |
| 2               | Letter (small) |
| 3               | Tabloid        |
| 4               | Ledger         |
| 5               | Legal          |
| 6               | Statement      |
| 7               | Executive      |
| 8               | A3             |
| 9               | A4             |
| 10              | A4 (small)     |
| 11              | A5             |
| 12              | B4             |
| 13              | B5             |
| 14              | Folio          |
| 15              | Quarto         |
| 16              | 10x14          |
| 17              | 11x17          |
| 18              | Note           |
| 19              | ENV9           |
| 20              | ENV10          |
| 21              | ENV11          |
| 22              | ENV12          |
| 23              | ENV14          |
| 24              | C Sheet        |
| 25              | D Sheet        |
| 26              | E Sheet        |

Scale&nbsp;&nbsp;&nbsp;&nbsp;is a number representing the percentage to
increase or decrease the size of the sheet. All scaling retains the
aspect ratio of the original.

  - > To specify a percentage of reduction or enlargement, set scale to
    > the percentage.

  - > For worksheets and macros, you can specify the number of pages
    > that the printout should be scaled to fit. Set scale to a two-item
    > horizontal array, with the first item equal to the width and the
    > second item equal to the height. If no constraint is necessary in
    > one direction, you can set the corresponding value to \#N/A.

  - > Scale can also be a logical value. To fit the print area on a
    > single page, set scale to TRUE.


Pg\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies the number of the first page.
If zero, sets first page to zero. If "Auto" is used, then the page
numbering is set to automatic. If omitted, PAGE.SETUP retains the
existing pg\_num.

Pg\_order&nbsp;&nbsp;&nbsp;&nbsp;specifies whether pagination is
left-to-right and then down, or top-to-bottom and then right.

|               |                           |
| ------------- | ------------------------- |
| **Pg\_order** | **Pagination**            |
| 1             | Top-to-bottom, then right |
| 2             | Left-to-right, then down  |

Bw\_cells&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that specifies
whether to print cells and all graphic objects, such as text boxes and
buttons, in color.

  - > If bw\_cells is TRUE, Microsoft Excel prints cell text and borders
    > in black and cell backgrounds in white.

  - > If bw\_cells is FALSE , Microsoft Excel prints cell text, borders,
    > and background patterns in color (or in gray scale).


Bw\_chart&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that specifies
whether to print chart in color.

Size&nbsp;&nbsp;&nbsp;&nbsp;is a number corresponding to the options in
the Chart Size box, and determines how you want the chart printed on the
page within the margins. Size is available only in the chart form of the
function.

|          |                             |
| -------- | --------------------------- |
| **Size** | **Size to print the chart** |
| 1        | Screen size                 |
| 2        | Fit to page                 |
| 3        | Full page                   |

Quality&nbsp;&nbsp;&nbsp;&nbsp;specifies the print quality in
dots-per-inch. To specify both horizontal and vertical print quality,
use an array of two values.

Head\_margin&nbsp;&nbsp;&nbsp;&nbsp;is the placement, in inches, of the
running head margin from the edge of the page.

Foot\_margin&nbsp;&nbsp;&nbsp;&nbsp;is the placement, in inches, of the
running foot margin from the edge of the page.

Draft&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Draft Quality checkbox
in the Sheet tab and in the Chart tab of the Page Setup dialog box. If
FALSE or omitted, graphics are printed with the sheet. If TRUE, no
graphics are printed.

Notes&nbsp;&nbsp;&nbsp;&nbsp;specifies whether to print cell notes with
the sheet. If TRUE, both the sheet and the cell notes are printed. If
FALSE or omitted, just the sheet is printed.

**Remarks**

Microsoft Excel no longer requires you to enter formatting codes to
format headers and footers, but the codes are still supported and
recorded by the macro recorder. You can include these codes as part of
the head and foot text strings to align portions of the header or footer
to the left, right, or center; to include the page number, date, time,
or workbook name; and to print the header or footer in bold or italic.

|                         |                                                                                                                                                                                             |
| ----------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Formatting code**     | **Result**                                                                                                                                                                                  |
| \&L                     | Left-aligns the characters that follow.                                                                                                                                                     |
| \&C                     | Centers the characters that follow.                                                                                                                                                         |
| \&R                     | Right-aligns the characters that follow.                                                                                                                                                    |
| \&B                     | Turns bold printing on or off (now obsolete).                                                                                                                                               |
| \&I                     | Turns italic printing on or off.                                                                                                                                                            |
| \&U                     | Turns single underlining printing on or off.                                                                                                                                                |
| \&S                     | Turns strikethrough printing on or off.                                                                                                                                                     |
| \&O                     | Turns outline printing on or off (Macintosh only).                                                                                                                                          |
| \&H                     | Turns shadow printing on or off (Macintosh only).                                                                                                                                           |
| \&D                     | Prints the current date.                                                                                                                                                                    |
| \&T                     | Prints the current time.                                                                                                                                                                    |
| \&A                     | Prints the name of the sheet                                                                                                                                                                |
| \&F                     | Prints the name of the workbook.                                                                                                                                                            |
| \&P                     | Prints the page number.                                                                                                                                                                     |
| \&P+number              | Prints the page number plus number.                                                                                                                                                         |
| \&P-number              | Prints the page number minus number.                                                                                                                                                        |
| &&                      | Prints a single ampersand.                                                                                                                                                                  |
| & "fontname, fontstyle" | Prints the characters that follow in the specified font and style. Be sure to include a comma immediately following the fontname, and double quotation marks around fontname and fontstyle. |
| \&nn                    | Prints the characters that follow in the specified font size. Use a two-digit number to specify a size in points.                                                                           |
| \&N                     | Prints the total number of pages in the workbook.                                                                                                                                           |
| \&E                     | Prints a double underline                                                                                                                                                                   |
| \&X                     | Prints the character as superscript                                                                                                                                                         |
| \&Y                     | Prints the character as subscript                                                                                                                                                           |

**Related Functions**

[DISPLAY](DISPLAY.md)&nbsp;&nbsp;&nbsp;Controls screen and Info Window display

[GET.DOCUMENT](GET.DOCUMENT.md)&nbsp;&nbsp;&nbsp;Returns information about a workbook

[PRINT](PRINT.md)&nbsp;&nbsp;&nbsp;Prints the active workbook

[WORKSPACE](WORKSPACE.md)&nbsp;&nbsp;&nbsp;Changes workspace settings



Return to [README](README.md#P)

