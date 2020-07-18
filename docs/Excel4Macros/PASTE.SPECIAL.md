PASTE.SPECIAL

Equivalent to clicking the Paste Special command on the Edit menu.
Pastes the specified components from the copy area into the current
selection. The PASTE.SPECIAL function has four syntax forms.

Syntax 1&nbsp;&nbsp;&nbsp;Pasting into a sheet or macro sheet

Syntax 2&nbsp;&nbsp;&nbsp;Copying from a sheet and pasting into a chart.

Syntax 3&nbsp;&nbsp;&nbsp;Copying and pasting between charts

Syntax 4&nbsp;&nbsp;&nbsp;Pasting information from another application.



Return to [README](README.md)

PASTE.SPECIAL Syntax 1

Equivalent to clicking the Paste Special command on the Edit menu.
Pastes the specified components from the copy area into the current
selection. The PASTE.SPECIAL function has four syntax forms. Use syntax
1 if you are pasting into a sheet or macro sheet.

**Syntax**

**PASTE.SPECIAL**(paste\_num, operation\_num, skip\_blanks, transpose)

**PASTE.SPECIAL**?(paste\_num, operation\_num, skip\_blanks, transpose)

Paste\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 6 specifying
what to paste. Paste\_num can also be quoted text of the object you want
to paste.

|                |                    |
| -------------- | ------------------ |
| **Paste\_num** | **Pastes**         |
| 1              | All                |
| 2              | Formulas           |
| 3              | Values             |
| 4              | Formats            |
| 5              | Comments           |
| 6              | All except borders |

Operation\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 5 specifying
which operation to perform when pasting.

|                    |            |
| ------------------ | ---------- |
| **Operation\_num** | **Action** |
| 1                  | None       |
| 2                  | Add        |
| 3                  | Subtract   |
| 4                  | Multiply   |
| 5                  | Divide     |

Skip\_blanks&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to
the Skip Blanks check box in the Paste Special dialog box.

  - > If skip\_blanks is TRUE, Microsoft Excel skips blanks in the copy
    > area when pasting.

  - > If skip\_blanks is FALSE, Microsoft Excel pastes normally.


Transpose&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to the
Transpose check box in the Paste Special dialog box.

  - > If transpose is TRUE, Microsoft Excel transposes rows and columns
    > when pasting.

  - > If transpose is FALSE, Microsoft Excel pastes normally.


**Related Functions**

[FORMULA](FORMULA.md)&nbsp;&nbsp;&nbsp;Enters values into a cell or range or onto a
chart

[PASTE](PASTE.md)&nbsp;&nbsp;&nbsp;Pastes cut or copied data

[PASTE.LINK](PASTE.LINK.md)&nbsp;&nbsp;&nbsp;Pastes copied data and establishes a link to
its source

[S](S.md)yntax 2&nbsp;&nbsp;&nbsp;Copying from a sheet and pasting into a chart.

[S](S.md)yntax 3&nbsp;&nbsp;&nbsp;Copying and pasting between charts

[S](S.md)yntax 4&nbsp;&nbsp;&nbsp;Pasting information from another application.



Return to [README](README.md)

PASTE.SPECIAL Syntax 2

Equivalent to clicking the Paste Special command on the Edit menu on the
Chart menu bar. Pastes the specified components from the copy area into
a chart. The PASTE.SPECIAL function has four syntax forms. Use syntax 2
if you have copied from a sheet and are pasting into a chart.

**Syntax**

**PASTE.SPECIAL**(rowcol, titles, categories, replace, series)

**PASTE.SPECIAL**?(rowcol, titles, categories, replace, series)

Rowcol&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 and specifies whether
the values corresponding to a particular data series are in rows or
columns. Enter 1 for rows or 2 for columns.

Titles&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to the
Series Names In First Column check box (or First Row, depending on the
value of rowcol) in the Paste Special dialog box.

  - > If series is TRUE, Microsoft Excel selects the check box and uses
    > the contents of the cell in the first column of each row (or first
    > row of each column) as the name of the data series in that row (or
    > column).

  - > If series is FALSE, Microsoft Excel clears the check box and uses
    > the contents of the cell in the first column of each row (or first
    > row of each column) as the first data point of the data series.


Categories&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to
the Categories (X Labels) In First Row (or First Column, depending on
the value of rowcol) check box in the Paste Special dialog box.

  - > If categories is TRUE, Microsoft Excel selects the check box and
    > uses the contents of the first row (or column) of the selection as
    > the categories for the chart.

  - > If categories is FALSE, Microsoft Excel clears the check box and
    > uses the contents of the first row (or column) as the first data
    > series in the chart.

Replace&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to the
Replace Existing Categories check box in the Paste Special dialog box.

  - > If replace is TRUE, Microsoft Excel selects the check box and
    > applies categories while replacing existing categories with
    > information from the copied cell range.

  - > If replace is FALSE, Microsoft Excel clears the check box and
    > applies new categories without replacing any old ones.

Series&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying how cells are added
to a chart.

|            |              |
| ---------- | ------------ |
| **Series** | **Added as** |
| 1          | New series   |
| 2          | New point(s) |

**Related Functions**

[S](S.md)yntax 1&nbsp;&nbsp;&nbsp;Pasting into a sheet or macro sheet

[S](S.md)yntax 3&nbsp;&nbsp;&nbsp;Copying and pasting between charts

[S](S.md)yntax 4&nbsp;&nbsp;&nbsp;Pasting information from another application



Return to [README](README.md)

PASTE.SPECIAL Syntax 3

Equivalent to clicking the Paste Special command on the Edit menu on the
Chart menu bar. Pastes the specified components from the copy area into
a chart. The PASTE.SPECIAL function has four syntax forms. Use syntax 3
if you have copied from a chart and are pasting into a chart.

**Syntax**

**PASTE.SPECIAL**(paste\_num)

**PASTE.SPECIAL**?(paste\_num)

Paste\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 3 specifying
what to paste.

|                |                               |
| -------------- | ----------------------------- |
| **Paste\_num** | **Pastes**                    |
| 1              | All (formats and data series) |
| 2              | Formats only                  |
| 3              | Formulas (data series) only   |

**Related Functions**

[S](S.md)yntax 1&nbsp;&nbsp;&nbsp;Pasting into a sheet or macro sheet

[S](S.md)yntax 2&nbsp;&nbsp;&nbsp;Copying from a sheet and pasting into a chart

[S](S.md)yntax 4&nbsp;&nbsp;&nbsp;Pasting information from another application



Return to [README](README.md)

PASTE.SPECIAL Syntax 4

Equivalent to clicking the Paste Special command on the Edit menu when
pasting data from another application into Microsoft Excel. Use syntax 4
when pasting information from another application.

**Syntax**

**PASTE.SPECIAL**(format\_text, pastelink\_logical,
display\_icon\_logical, icon\_file, icon\_number, icon\_label)

**PASTE.SPECIAL**?(format\_text, pastelink\_logical,
display\_icon\_logical, icon\_file, icon\_number, icon\_label)

Format\_text&nbsp;&nbsp;&nbsp;&nbsp;is text specifying the type of data
you want to paste from the Clipboard.

  - > The valid data types vary depending on the application from which
    > the data was copied. For example, if you're copying data from
    > Microsoft Word, some of the data types are "Microsoft Document
    > Object", "Picture", and "Text".

  - > For more information about object classes, see your Microsoft
    > Windows or Apple system software documentation.


Pastelink\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying
whether to link the pasted data to its source application.

  - > If pastelink\_logical is TRUE, Microsoft Excel updates the pasted
    > information whenever it changes in the source application.

  - > If pastelink\_logical is FALSE or omitted, the information is
    > pasted without a link.

  - > If Microsoft Excel or the source application does not support
    > linking for the specified format\_text, then pastelink\_logical is
    > ignored.

Display\_icon\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that
specifies whether you want an application's icon to be displayed on the
worksheet instead of the linked data. Equivalent to the Display as Icon
check box on the Paste Special dialog box. If TRUE, the application's
icon will be displayed. If FALSE or omitted, the application's icon will
not be displayed.

Icon\_file&nbsp;&nbsp;&nbsp;&nbsp;is the name of the file (with an .EXE
or .DLL extension) that contains the icon. If display\_icon\_logical is
FALSE, this argument is ignored.

Icon\_number&nbsp;&nbsp;&nbsp;&nbsp;is the number associated with the
icon and corresponds to the icon's relative position within the Icon
Drop Down list box on the Change Icon Dialog box, which appears when you
click the Change Icon button in the Paste Special dialog box. If
display\_icon\_logical is FALSE, this argument is ignored.

Icon\_label&nbsp;&nbsp;&nbsp;&nbsp;is the caption that you want to
appear below the icon, and is equivalent to the Caption text box on the
Change Icon dialog box, which appears when you click the Change Icon
button in the Paste Special dialog box. If display\_icon\_logical is
FALSE, this argument is ignored.

**Related Functions**

[S](S.md)yntax 1&nbsp;&nbsp;&nbsp;Pasting into a sheet or macro sheet

[S](S.md)yntax 2&nbsp;&nbsp;&nbsp;Copying from a sheet and pasting into a chart

[S](S.md)yntax 3&nbsp;&nbsp;&nbsp;Copying and pasting between charts



Return to [README](README.md)

