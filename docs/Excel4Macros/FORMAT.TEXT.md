FORMAT.TEXT

Formats the selected worksheet text box or button or any text item on a
chart.

**Syntax**

**FORMAT.TEXT**(x\_align, y\_align, orient\_num, auto\_text, auto\_size,
show\_key, show\_value, add\_indent)

**FORMAT.TEXT**?(x\_align, y\_align, orient\_num, auto\_text,
auto\_size, show\_key, show\_value, add\_indent)

Arguments correspond to check boxes or options in the various tabs on
Format Object dialog box. Arguments that correspond to check boxes are
logical values. If an argument is TRUE, Microsoft Excel selects the
check box; if FALSE, Microsoft Excel clears the check box; if omitted,
the current setting is used.

X\_align    is a number from 1 to 4 specifying the horizontal alignment
of the text.

|              |                          |
| ------------ | ------------------------ |
| **X\_align** | **Horizontal alignment** |
| 1            | Left                     |
| 2            | Center                   |
| 3            | Right                    |
| 4            | Justify                  |

Y\_align    is a number from 1 to 4 specifying the vertical alignment of
the text.

|              |                        |
| ------------ | ---------------------- |
| **Y\_align** | **Vertical alignment** |
| 1            | Top                    |
| 2            | Center                 |
| 3            | Bottom                 |
| 4            | Justify                |

Orient\_num    is a number from 0 to 3 specifying the orientation of the
text.

|                 |                      |
| --------------- | -------------------- |
| **Orient\_num** | **Text orientation** |
| 0               | Horizontal           |
| 1               | Vertical             |
| 2               | Upward               |
| 3               | Downward             |

Auto\_text    corresponds to the Automatic Text check box. If the
selected text was created with the Data Labels command from the Insert
menu and later edited, this option restores the original text.
Auto\_text is ignored for text boxes on worksheets and macro sheets.

Auto\_size    corresponds to the Automatic Size check box. If you have
changed the size of the border around the selected text, this option
restores the border to automatic size. Automatic size makes the border
fit exactly around the text no matter how you change the text.

Show\_key    corresponds to the Show Legend Key Next to Label check box
in the Data Labels dialog box. This argument applies only if the
selected text is an attached data label on a chart.

Show\_value    corresponds to the Show Value option button in the Format
Data Labels dialog box. This argument applies only if the selected text
is an attached data label on a chart.

The following list summarizes which arguments apply to each type of text
item.

Add\_indent   This argument is for only Far East versions of Microsoft
Excel.

|                              |                                             |
| ---------------------------- | ------------------------------------------- |
| **Text item**                | **Arguments that apply**                    |
| Worksheet text box or button | X\_align, y\_align, orient\_num, auto\_size |
| Attached data label          | All arguments                               |
| Unattached text label        | X\_align, y\_align, orient\_num, auto\_size |
| Tickmark label               | Orient\_num                                 |

**Related Functions**

[CREATE.OBJECT](CREATE.OBJECT.md)   Creates an object

[FONT.PROPERTIES](FONT.PROPERTIES.md)   Applies a font to the selection

[FORMULA](FORMULA.md)   Enters values into a cell or range or onto a chart



Return to [README.md](README.md)

FORMAT.TEXT

Formats the selected worksheet text box or button or any text item on a
chart.

**Syntax**

**FORMAT.TEXT**(x\_align, y\_align, orient\_num, auto\_text, auto\_size,
show\_key, show\_value, add\_indent)

**FORMAT.TEXT**?(x\_align, y\_align, orient\_num, auto\_text,
auto\_size, show\_key, show\_value, add\_indent)

Arguments correspond to check boxes or options in the various tabs on
Format Object dialog box. Arguments that correspond to check boxes are
logical values. If an argument is TRUE, Microsoft Excel selects the
check box; if FALSE, Microsoft Excel clears the check box; if omitted,
the current setting is used.

X\_align    is a number from 1 to 4 specifying the horizontal alignment
of the text.

|              |                          |
| ------------ | ------------------------ |
| **X\_align** | **Horizontal alignment** |
| 1            | Left                     |
| 2            | Center                   |
| 3            | Right                    |
| 4            | Justify                  |

Y\_align    is a number from 1 to 4 specifying the vertical alignment of
the text.

|              |                        |
| ------------ | ---------------------- |
| **Y\_align** | **Vertical alignment** |
| 1            | Top                    |
| 2            | Center                 |
| 3            | Bottom                 |
| 4            | Justify                |

Orient\_num    is a number from 0 to 3 specifying the orientation of the
text.

|                 |                      |
| --------------- | -------------------- |
| **Orient\_num** | **Text orientation** |
| 0               | Horizontal           |
| 1               | Vertical             |
| 2               | Upward               |
| 3               | Downward             |

Auto\_text    corresponds to the Automatic Text check box. If the
selected text was created with the Data Labels command from the Insert
menu and later edited, this option restores the original text.
Auto\_text is ignored for text boxes on worksheets and macro sheets.

Auto\_size    corresponds to the Automatic Size check box. If you have
changed the size of the border around the selected text, this option
restores the border to automatic size. Automatic size makes the border
fit exactly around the text no matter how you change the text.

Show\_key    corresponds to the Show Legend Key Next to Label check box
in the Data Labels dialog box. This argument applies only if the
selected text is an attached data label on a chart.

Show\_value    corresponds to the Show Value option button in the Format
Data Labels dialog box. This argument applies only if the selected text
is an attached data label on a chart.

The following list summarizes which arguments apply to each type of text
item.

Add\_indent   This argument is for only Far East versions of Microsoft
Excel.

|                              |                                             |
| ---------------------------- | ------------------------------------------- |
| **Text item**                | **Arguments that apply**                    |
| Worksheet text box or button | X\_align, y\_align, orient\_num, auto\_size |
| Attached data label          | All arguments                               |
| Unattached text label        | X\_align, y\_align, orient\_num, auto\_size |
| Tickmark label               | Orient\_num                                 |

**Related Functions**

[CREATE.OBJECT](CREATE.OBJECT.md)   Creates an object

[FONT.PROPERTIES](FONT.PROPERTIES.md)   Applies a font to the selection

[FORMULA](FORMULA.md)   Enters values into a cell or range or onto a chart



Return to [README.md](README.md)

