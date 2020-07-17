FONT.PROPERTIES

Equivalent to choosing the Cells command from the Format menu. Applies a
font and other attributes to the selection. Applies to cells, charts,
and text boxes and buttons on worksheets and macro sheets.

**Syntax**

**FONT.PROPERTIES**(font, font\_style, size, strikethrough, superscript,
subscript, outline, shadow, underline, color, normal, background,
start\_char, char\_count)

**FONT.PROPERTIES**?(font, font\_style, size, strikethrough,
superscript, subscript, outline, shadow, underline, color, normal,
background, start\_char, char\_count)

Arguments correspond to check boxes or options in the Font tab on the
Format Cells dialog box. Arguments that correspond to check boxes are
logical values. If an argument is TRUE, Microsoft Excel selects the
check box; if FALSE, Microsoft Excel clears the check box. If an
argument is omitted, the format is not changed.

Font    is the name of the font as it appears on the Font tab. For
example, Courier is a font name.

Font\_style    is the name of the font style as it appears on the Font
tab. For example, Bold Italic is a font style.

Size    is the font size, in points.

Strikethrough    corresponds to the Strikethrough check box.

Superscript    corresponds to the Superscript check box

Subscript    corresponds to the Subscript check box

Outline    corresponds to the Outline check box. Outline fonts are
available in Microsoft Excel for the Macintosh. For macro compatibility,
this argument is ignored by Microsoft Excel for Windows..

Shadow    corresponds to the Shadow check box. Shadow fonts are
available in Microsoft Excel for the Macintosh. For macro compatibility,
this argument is ignored by Microsoft Excel for Windows.

**Note   **For macro compatibility with Microsoft Excel for the
Macintosh, the presence of the outline and shadow arguments do not
prevent the macro from working on Microsoft Excel for Windows, nor does
their absence prevent it from working on the Macintosh.

Underline    corresponds to the Underline Drop-down box.

|               |                   |
| ------------- | ----------------- |
| **Underline** | **Type applied**  |
| 0             | None              |
| 1             | Single            |
| 2             | Double            |
| 3             | Single Accounting |
| 4             | Double Accounting |

Color    is a number from 0 to 56 corresponding to the colors listed in
the Color box; 0 corresponds to automatic color.

Normal    corresponds to the Normal Font check box. Applies the default
font for your system

Background    is a number from 1 to 3 specifying which type of
background to apply to text in a chart.

|                |                                |
| -------------- | ------------------------------ |
| **Background** | **Type of background applied** |
| 1              | Automatic                      |
| 2              | Transparent                    |
| 3              | Opaque                         |

Start\_char    specifies the first character to be formatted. If
start\_char is omitted, it is assumed to be 1 (the first character in
the cell or text box).

Char\_count    specifies how many characters to format. If char\_count
is omitted, Microsoft Excel formats all characters in the cell or text
box starting at start\_char.

**Remarks**

Some extended TrueType styles do not have corresponding arguments to
FONT.PROPERTIES. To access an extended TrueType font style, append the
style name to the font name in the font argument. For example, the font
Taipei can be formatted in an upside-down style by specifying "Taipei
Upside-down" as the font argument. For more information about TrueType,
see your Microsoft Windows documentation.

**Related Functions**

ALIGNMENT   Aligns or wraps text in cells

FORMAT.NUMBER   Applies a number format to the selection

FORMAT.TEXT   Formats a worksheet text box or a chart text item


