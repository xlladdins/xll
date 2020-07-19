# FORMAT.FONT

Equivalent to choosing the Cells command from the Format menu, and then
selecting Font tab from the Format Cells dialog box. This function is
included for compatibility with Microsoft Excel version 4.0. Use
FONT.PROPERTIES to set various font properties. FORMAT.FONT has three
syntax forms. Syntax 1 is for cells; syntax 2 is for text boxes and
buttons; syntax 3 is used with all chart items (axes, labels, text, and
so on).

**Syntax 1**

Cells

**FORMAT.FONT**(name\_text, size\_num, bold, italic, underline, strike,
color, outline, shadow)

**FORMAT.FONT**?(name\_text, size\_num, bold, italic, underline, strike,
color, outline, shadow)

**Syntax 2**

Text boxes and buttons on worksheets and macro sheets

**FORMAT.FONT**(name\_text, size\_num, bold, italic, underline, strike,
color, outline, shadow, object\_id\_text, start\_num, char\_num)

**FORMAT.FONT**?(name\_text, size\_num, bold, italic, underline, strike,
color, outline, shadow, object\_id\_text, start\_num, char\_num)

**Syntax 3**

Chart items including unattached chart text

**FORMAT.FONT**(color, backgd, apply, name\_text, size\_num, bold,
italic, underline, strike, outline, shadow, object\_id\_text,
start\_num, char\_num)

**FORMAT.FONT**?(color, backgd, apply, name\_text, size\_num, bold,
italic, underline, strike, outline, shadow, object\_id\_text,
start\_num, char\_num)

Arguments correspond to check boxes and list box items in the Font tab
on the Format Cells dialog box. Arguments that correspond to check boxes
are logical values. If an argument is TRUE, Microsoft Excel selects the
check box; if FALSE, Microsoft Excel clears the check box. If an
argument is omitted, the format is not changed.

Name\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of the font as it appears
in the Font tab. For example, Courier is a font name.

Size\_num&nbsp;&nbsp;&nbsp;&nbsp;is the font size, in points.

Bold&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Bold item in the Font
Style list box. Makes the selection bold, if applicable.

Italic&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Italic item in the Font
Style list box. Makes the selection italic, if applicable.

Underline&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Underline check box.

Strike&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Strikethrough check
box.

Color&nbsp;&nbsp;&nbsp;&nbsp;is a number from 0 to 56 corresponding to
the colors in the Font tab; 0 corresponds to automatic color.

Outline&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Outline check box.
Outline fonts are available in Microsoft Excel for the Macintosh. For
macro compatibility, this argument is ignored by Microsoft Excel for
Windows.

Shadow&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Shadow check box.
Shadow fonts are available in Microsoft Excel for the Macintosh. For
macro compatibility, this argument is ignored by Microsoft Excel for
Windows.

**Note**&nbsp;&nbsp;&nbsp;For macro compatibility with Microsoft Excel
for the Macintosh, the presence of the outline and shadow arguments do
not prevent the macro from working on Microsoft Excel for Windows, nor
does their absence prevent it from working on the Macintosh.

Object\_id\_text&nbsp;&nbsp;&nbsp;&nbsp;identifies the text box you want
to format (for example, "Text 1", "Text 2", and so on). You can also use
the object number alone without the text identifier. For compatibility
with earlier versions of Microsoft Excel. This argument is ignored in
Microsoft Excel version 5.0 or later.

Start\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies the first character to be
formatted. If start\_num is omitted, it is assumed to be 1 (the first
character in the text box).

Char\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies how many characters to
format. If char\_num is omitted, Microsoft Excel formats all characters
in the text box starting at start\_num.

Backgd&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 3 specifying which
type of background to apply to text in a chart.

|            |                                |
| ---------- | ------------------------------ |
| **Backgd** | **Type of background applied** |
| 1          | Automatic                      |
| 2          | Transparent                    |
| 3          | Opaque                         |

Apply&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Apply To All check box.
This argument applies to data labels only.

**Remarks**

Some extended TrueType styles do not have corresponding arguments to
FORMAT.FONT. To access an extended TrueType font style, append the style
name to the font name in name\_text. For example, the font Taipei can be
formatted in an upside-down style by specifying "Taipei Upside-down" as
the name\_text argument. For more information about TrueType, see your
Microsoft Windows documentation.

**Related Functions**

[ALIGNMENT](ALIGNMENT.md)&nbsp;&nbsp;&nbsp;Aligns or wraps text in cells

[FONT.PROPERTIES](FONT.PROPERTIES.md)&nbsp;&nbsp;&nbsp;Sets various font attributes

[FORMAT.NUMBER](FORMAT.NUMBER.md)&nbsp;&nbsp;&nbsp;Applies a number format to the selection

[FORMAT.TEXT](FORMAT.TEXT.md)&nbsp;&nbsp;&nbsp;Formats a worksheet text box or a chart
text item



Return to [README](README.md)

