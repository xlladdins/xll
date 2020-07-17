PATTERNS

Equivalent to clicking the Patterns tab in the Format Cells dialog box,
which appears when you click the Cells command on the Format menu.
Changes the appearance of the selected cells or objects or the selected
chart item (you can select only one chart item at a time). The PATTERNS
function has eight syntax forms: syntax 1 is for cells on a sheet or
macro sheet. Syntax 2 is for lines or arrows on a worksheet, macro
sheet, or chart. Syntax 3 is for objects on a sheet or macro sheet.
Syntax 4 through syntax 8 are for chart items.

**Syntax 1**

Cells

**PATTERNS**(apattern, afore, aback, newui)

**PATTERNS**?(apattern, afore, aback, newui)

**Syntax 2**

Lines (arrows) on worksheets or charts

**PATTERNS**(lauto, lstyle, lcolor, lwt, hwidth, hlength, htype)

**PATTERNS**?(lauto, lstyle, lcolor, lwt, hwidth, hlength, htype)

**Syntax 3**

Text boxes, rectangles, ovals, arcs, and pictures on worksheets or macro
sheets

**PATTERNS**(bauto, bstyle, bcolor, bwt, shadow, aauto, apattern, afore,
aback, rounded, newui)

**PATTERNS**?(bauto, bstyle, bcolor, bwt, shadow, aauto, apattern,
afore, aback, rounded, newui)

**Syntax 4**

Chart plot areas, bars, columns, pie slices, and text labels

**PATTERNS**(bauto, bstyle, bcolor, bwt, shadow, aauto, apattern, afore,
aback, invert, apply, new\_fill)

**PATTERNS**?(bauto, bstyle, bcolor, bwt, shadow, aauto, apattern,
afore, aback, invert, apply, new\_fill)

**Syntax 5**

Chart axes

**PATTERNS**(lauto, lstyle, lcolor, lwt, tmajor, tminor, tlabel)

**PATTERNS**?(lauto, lstyle, lcolor, lwt, tmajor, tminor, tlabel)

**Syntax 6**

Chart gridlines, hi-lo lines, drop lines, lines on a picture line chart,
and picture charts of bar and column charts

**PATTERNS**(lauto, lstyle, lcolor, lwt, apply, smooth)

**PATTERNS**?(lauto, lstyle, lcolor, lwt, apply, smooth)

**Syntax 7**

Chart data lines

**PATTERNS**(lauto, lstyle, lcolor, lwt, mauto, mstyle, mfore, mback,
apply, smooth)

**PATTERNS**?(lauto, lstyle, lcolor, lwt, mauto, mstyle, mfore, mback,
apply, smooth)

**Syntax 8**

Picture chart markers

**PATTERNS**(type, picture\_units, apply)

**PATTERNS**?(type, picture\_units, apply)

The following argument descriptions are in alphabetic order. Arguments
correspond to check boxes, list boxes, and options in the Patterns tab
of the Format Cells dialog box for the selected item. The default for
each argument reflects the setting in the dialog box.

Aauto    is a number from 0 to 2 specifying area settings (that is, the
object's "surface area").

|                 |                                    |
| --------------- | ---------------------------------- |
| **If aauto is** | **Area settings are**              |
| 0               | Set by the user (custom)           |
| 1               | Automatic (set by Microsoft Excel) |
| 2               | None                               |

Aback    is a number from 1 to 56 corresponding to the 56 area
background colors in the Patterns tab of the Format Cells dialog box.

Afore    is a number from 1 to 56 corresponding to the 56 area
foreground colors in the Patterns tab of the Format Cells dialog box.

Apattern    is a number corresponding to the area patterns in the
Patterns tab of the Format Cells or Format Object dialog box. If an
object is selected, apattern can be from 1 to 18; if a cell is selected,
apattern can be from 0 to 18. If apattern is 0 and a cell is selected,
Microsoft Excel applies no pattern.

Apply    is a logical value corresponding to the Apply To All check box
in Microsoft Excel version 4.0. This argument is for compatibility with
previous versions of Microsoft Excel and applies only when a chart data
point or a data series is selected.

  - > If apply is TRUE, Microsoft Excel applies any formatting changes
    > to all items that are similar to the selected item on the chart.

  - > If apply is FALSE, Microsoft Excel applies formatting changes only
    > to the selected item on the chart.

>  

Bauto    is a number from 0 to 2 specifying border settings.

|                 |                                    |
| --------------- | ---------------------------------- |
| **If bauto is** | **Border settings are**            |
| 0               | Set by the user (custom)           |
| 1               | Automatic (set by Microsoft Excel) |
| 2               | None                               |

Bcolor    is a number from 1 to 56 corresponding to the 56 border colors
in the Border tab of the Format Object or Format (chart element) dialog
box.

Bstyle    is a number from 1 to 8 corresponding to the eight border
styles in the Border tab of the Format Object or Format (chart element)
dialog box.

Bwt    is a number from 1 to 4 corresponding to the four border weights
in the Border tab of the Format Object or Format (chart element) dialog
box.

|               |               |
| ------------- | ------------- |
| **If bwt is** | **Border is** |
| 1             | Hairline      |
| 2             | Thin          |
| 3             | Medium        |
| 4             | Thick         |

Hlength    is a number from 1 to 3 specifying the length of the
arrowhead.

|                   |                  |
| ----------------- | ---------------- |
| **If hlength is** | **Arrowhead is** |
| 1                 | Short            |
| 2                 | Medium           |
| 3                 | Long             |

Htype    is a number from 1 to 5 specifying the style of the arrowhead.

|                 |                           |
| --------------- | ------------------------- |
| **If htype is** | **Style of arrowhead is** |
| 1               | No head                   |
| 2               | Open head                 |
| 3               | Closed head               |
| 4               | Double open head          |
| 5               | Double closed head        |

Hwidth    is a number from 1 to 3 specifying the width of the arrowhead.

|                  |                  |
| ---------------- | ---------------- |
| **If hwidth is** | **Arrowhead is** |
| 1                | Narrow           |
| 2                | Medium           |
| 3                | Wide             |

Invert    is a logical value corresponding to the Invert If Negative
check box in the Patterns tab of the Format Data Series dialog box. This
argument applies only to data markers.

  - > If invert is TRUE, Microsoft Excel inverts the pattern in the
    > selected item if it corresponds to a negative number.

  - > If invert is FALSE, Microsoft Excel removes the inverted pattern,
    > if present, from the selected item corresponding to a negative
    > value.

>  

Lauto    is a number from 0 to 2 specifying line settings.

|                 |                                    |
| --------------- | ---------------------------------- |
| **If lauto is** | **Line settings are**              |
| 0               | Set by the user (custom)           |
| 1               | Automatic (set by Microsoft Excel) |
| 2               | None                               |

Lcolor    is a number from 1 to 56 corresponding to the 16 line colors
in the Patterns tab of the Format Object or Format (chart element)
dialog box.

Lstyle    is a number from 1 to 8 corresponding to the eight line styles
in the Patterns tab of the Format Object or Format (chart element)
dialog box.

Lwt    is a number from 1 to 4 corresponding to the four line weights in
the Patterns tab of the Format Object or Format (chart element) dialog
box.

|               |             |
| ------------- | ----------- |
| **If lwt is** | **Line is** |
| 1             | Hairline    |
| 2             | Thin        |
| 3             | Medium      |
| 4             | Thick       |

Mauto    is a number from 0 to 2 specifying marker settings.

|                 |                                    |
| --------------- | ---------------------------------- |
| **If mauto is** | **Marker settings are**            |
| 0               | Set by the user                    |
| 1               | Automatic (set by Microsoft Excel) |
| 2               | None                               |

Mback    is a number from 1 to 56corresponding to the 56 marker
background colors in the Patterns tab of the Format Data Series dialog
box.

Mfore    is a number from 1 to 56 corresponding to the 56 marker
foreground colors in the Patterns tab of the Format Data Series dialog
box.

Mstyle    is a number from 1 to 9 corresponding to the nine marker
styles in the Patterns tab of the Format Data Series dialog box.

Picture\_units    is the number of units you want each picture to
represent in a scaled, stacked picture chart. This argument applies only
to picture charts and only if type is 3.

Rounded    is a logical value corresponding to the Round Corners check
box and specifying whether to make the corners on text boxes and
rectangles rounded. If rounded is TRUE, the corners are rounded; if
FALSE, the corners are square. If the selection is an arc or an oval,
rounded is ignored.

Newui    is a logical value that specifies whether to use the
foreground, background, and patterns of Microsoft Excel version 5.0 or
later. If TRUE or omitted, the colors and patterns of Microsoft Excel
version 5.0 or later will be used. If FALSE, the colors and patterns of
Microsoft Excel version 4.0 will be used.

Newfill    is a logical value that specifies whether to use the chart
patterns of Microsoft Excel version 5.0 or later. If TRUE or omitted,
the chart patterns of Microsoft Excel version 5.0 or later will be used.
If FALSE, the chart patterns of Microsoft Excel version 4.0 will be
used.

Shadow    is a logical value corresponding to the Shadow check box.
Shadow does not apply to area charts or bars in bar charts. If shadow is
TRUE, Microsoft Excel adds a shadow to the selected item; if FALSE,
Microsoft Excel removes the shadow, if one is present, from the selected
item. If the selection is an arc, shadow is ignored.

Smooth    is a logical value that applies smoothing to picture markers
in line or xy (scatter) charts. The default is FALSE.

Tlabel    is a number from 1 to 4 specifying the position of tick
labels.

|                  |                            |
| ---------------- | -------------------------- |
| **If tlabel is** | **Tick label position is** |
| 1                | None                       |
| 2                | Low                        |
| 3                | High                       |
| 4                | Next to axis               |

Tmajor    is a number from 1 to 4 specifying the type of major tick
marks.

|                  |                                 |
| ---------------- | ------------------------------- |
| **If tmajor is** | **Type of major tick marks is** |
| 1                | None                            |
| 2                | Inside                          |
| 3                | Outside                         |
| 4                | Cross                           |

Tminor    is a number from 1 to 4 specifying the type of minor tick
marks.

|                  |                                 |
| ---------------- | ------------------------------- |
| **If tminor is** | **Type of minor tick marks is** |
| 1                | None                            |
| 2                | Inside                          |
| 3                | Outside                         |
| 4                | Cross                           |

Type    is a number from 1 to 3 specifying the type of pictures to use
in a picture chart.

|                |                                                                                           |
| -------------- | ----------------------------------------------------------------------------------------- |
| **If type is** | **Pictures should be**                                                                    |
| 1              | Stretched to reach a particular value                                                     |
| 2              | Stacked on top of each other to reach a particular value                                  |
| 3              | Stacked on top of each other, but you specify the number of units each picture represents |

**Remarks**

  - > You can select many graphic objects on a sheet or macro sheet and
    > apply formatting to them at the same time, but you can select only
    > one chart item at a time.

  - > If you select multiple objects and if one or more of the objects
    > requires a different form of the PATTERNS function, then choose
    > the syntax corresponding to the object with the most formatting
    > attributes—that is, the syntax with the most arguments. If you
    > specify an argument that does not apply to an item, the argument
    > has no effect on that item.

  - > To apply formatting to similar items on a chart, use the apply
    > argument described above.

>  

**Related Functions**

[FONT.PROPERTIES](FONT.PROPERTIES.md)   Applies a font to the selection

[FORMAT.TEXT](FORMAT.TEXT.md)   Formats a text box or a chart text item



Return to [README.md](README.md)

PATTERNS

Equivalent to clicking the Patterns tab in the Format Cells dialog box,
which appears when you click the Cells command on the Format menu.
Changes the appearance of the selected cells or objects or the selected
chart item (you can select only one chart item at a time). The PATTERNS
function has eight syntax forms: syntax 1 is for cells on a sheet or
macro sheet. Syntax 2 is for lines or arrows on a worksheet, macro
sheet, or chart. Syntax 3 is for objects on a sheet or macro sheet.
Syntax 4 through syntax 8 are for chart items.

**Syntax 1**

Cells

**PATTERNS**(apattern, afore, aback, newui)

**PATTERNS**?(apattern, afore, aback, newui)

**Syntax 2**

Lines (arrows) on worksheets or charts

**PATTERNS**(lauto, lstyle, lcolor, lwt, hwidth, hlength, htype)

**PATTERNS**?(lauto, lstyle, lcolor, lwt, hwidth, hlength, htype)

**Syntax 3**

Text boxes, rectangles, ovals, arcs, and pictures on worksheets or macro
sheets

**PATTERNS**(bauto, bstyle, bcolor, bwt, shadow, aauto, apattern, afore,
aback, rounded, newui)

**PATTERNS**?(bauto, bstyle, bcolor, bwt, shadow, aauto, apattern,
afore, aback, rounded, newui)

**Syntax 4**

Chart plot areas, bars, columns, pie slices, and text labels

**PATTERNS**(bauto, bstyle, bcolor, bwt, shadow, aauto, apattern, afore,
aback, invert, apply, new\_fill)

**PATTERNS**?(bauto, bstyle, bcolor, bwt, shadow, aauto, apattern,
afore, aback, invert, apply, new\_fill)

**Syntax 5**

Chart axes

**PATTERNS**(lauto, lstyle, lcolor, lwt, tmajor, tminor, tlabel)

**PATTERNS**?(lauto, lstyle, lcolor, lwt, tmajor, tminor, tlabel)

**Syntax 6**

Chart gridlines, hi-lo lines, drop lines, lines on a picture line chart,
and picture charts of bar and column charts

**PATTERNS**(lauto, lstyle, lcolor, lwt, apply, smooth)

**PATTERNS**?(lauto, lstyle, lcolor, lwt, apply, smooth)

**Syntax 7**

Chart data lines

**PATTERNS**(lauto, lstyle, lcolor, lwt, mauto, mstyle, mfore, mback,
apply, smooth)

**PATTERNS**?(lauto, lstyle, lcolor, lwt, mauto, mstyle, mfore, mback,
apply, smooth)

**Syntax 8**

Picture chart markers

**PATTERNS**(type, picture\_units, apply)

**PATTERNS**?(type, picture\_units, apply)

The following argument descriptions are in alphabetic order. Arguments
correspond to check boxes, list boxes, and options in the Patterns tab
of the Format Cells dialog box for the selected item. The default for
each argument reflects the setting in the dialog box.

Aauto    is a number from 0 to 2 specifying area settings (that is, the
object's "surface area").

|                 |                                    |
| --------------- | ---------------------------------- |
| **If aauto is** | **Area settings are**              |
| 0               | Set by the user (custom)           |
| 1               | Automatic (set by Microsoft Excel) |
| 2               | None                               |

Aback    is a number from 1 to 56 corresponding to the 56 area
background colors in the Patterns tab of the Format Cells dialog box.

Afore    is a number from 1 to 56 corresponding to the 56 area
foreground colors in the Patterns tab of the Format Cells dialog box.

Apattern    is a number corresponding to the area patterns in the
Patterns tab of the Format Cells or Format Object dialog box. If an
object is selected, apattern can be from 1 to 18; if a cell is selected,
apattern can be from 0 to 18. If apattern is 0 and a cell is selected,
Microsoft Excel applies no pattern.

Apply    is a logical value corresponding to the Apply To All check box
in Microsoft Excel version 4.0. This argument is for compatibility with
previous versions of Microsoft Excel and applies only when a chart data
point or a data series is selected.

  - > If apply is TRUE, Microsoft Excel applies any formatting changes
    > to all items that are similar to the selected item on the chart.

  - > If apply is FALSE, Microsoft Excel applies formatting changes only
    > to the selected item on the chart.

>  

Bauto    is a number from 0 to 2 specifying border settings.

|                 |                                    |
| --------------- | ---------------------------------- |
| **If bauto is** | **Border settings are**            |
| 0               | Set by the user (custom)           |
| 1               | Automatic (set by Microsoft Excel) |
| 2               | None                               |

Bcolor    is a number from 1 to 56 corresponding to the 56 border colors
in the Border tab of the Format Object or Format (chart element) dialog
box.

Bstyle    is a number from 1 to 8 corresponding to the eight border
styles in the Border tab of the Format Object or Format (chart element)
dialog box.

Bwt    is a number from 1 to 4 corresponding to the four border weights
in the Border tab of the Format Object or Format (chart element) dialog
box.

|               |               |
| ------------- | ------------- |
| **If bwt is** | **Border is** |
| 1             | Hairline      |
| 2             | Thin          |
| 3             | Medium        |
| 4             | Thick         |

Hlength    is a number from 1 to 3 specifying the length of the
arrowhead.

|                   |                  |
| ----------------- | ---------------- |
| **If hlength is** | **Arrowhead is** |
| 1                 | Short            |
| 2                 | Medium           |
| 3                 | Long             |

Htype    is a number from 1 to 5 specifying the style of the arrowhead.

|                 |                           |
| --------------- | ------------------------- |
| **If htype is** | **Style of arrowhead is** |
| 1               | No head                   |
| 2               | Open head                 |
| 3               | Closed head               |
| 4               | Double open head          |
| 5               | Double closed head        |

Hwidth    is a number from 1 to 3 specifying the width of the arrowhead.

|                  |                  |
| ---------------- | ---------------- |
| **If hwidth is** | **Arrowhead is** |
| 1                | Narrow           |
| 2                | Medium           |
| 3                | Wide             |

Invert    is a logical value corresponding to the Invert If Negative
check box in the Patterns tab of the Format Data Series dialog box. This
argument applies only to data markers.

  - > If invert is TRUE, Microsoft Excel inverts the pattern in the
    > selected item if it corresponds to a negative number.

  - > If invert is FALSE, Microsoft Excel removes the inverted pattern,
    > if present, from the selected item corresponding to a negative
    > value.

>  

Lauto    is a number from 0 to 2 specifying line settings.

|                 |                                    |
| --------------- | ---------------------------------- |
| **If lauto is** | **Line settings are**              |
| 0               | Set by the user (custom)           |
| 1               | Automatic (set by Microsoft Excel) |
| 2               | None                               |

Lcolor    is a number from 1 to 56 corresponding to the 16 line colors
in the Patterns tab of the Format Object or Format (chart element)
dialog box.

Lstyle    is a number from 1 to 8 corresponding to the eight line styles
in the Patterns tab of the Format Object or Format (chart element)
dialog box.

Lwt    is a number from 1 to 4 corresponding to the four line weights in
the Patterns tab of the Format Object or Format (chart element) dialog
box.

|               |             |
| ------------- | ----------- |
| **If lwt is** | **Line is** |
| 1             | Hairline    |
| 2             | Thin        |
| 3             | Medium      |
| 4             | Thick       |

Mauto    is a number from 0 to 2 specifying marker settings.

|                 |                                    |
| --------------- | ---------------------------------- |
| **If mauto is** | **Marker settings are**            |
| 0               | Set by the user                    |
| 1               | Automatic (set by Microsoft Excel) |
| 2               | None                               |

Mback    is a number from 1 to 56corresponding to the 56 marker
background colors in the Patterns tab of the Format Data Series dialog
box.

Mfore    is a number from 1 to 56 corresponding to the 56 marker
foreground colors in the Patterns tab of the Format Data Series dialog
box.

Mstyle    is a number from 1 to 9 corresponding to the nine marker
styles in the Patterns tab of the Format Data Series dialog box.

Picture\_units    is the number of units you want each picture to
represent in a scaled, stacked picture chart. This argument applies only
to picture charts and only if type is 3.

Rounded    is a logical value corresponding to the Round Corners check
box and specifying whether to make the corners on text boxes and
rectangles rounded. If rounded is TRUE, the corners are rounded; if
FALSE, the corners are square. If the selection is an arc or an oval,
rounded is ignored.

Newui    is a logical value that specifies whether to use the
foreground, background, and patterns of Microsoft Excel version 5.0 or
later. If TRUE or omitted, the colors and patterns of Microsoft Excel
version 5.0 or later will be used. If FALSE, the colors and patterns of
Microsoft Excel version 4.0 will be used.

Newfill    is a logical value that specifies whether to use the chart
patterns of Microsoft Excel version 5.0 or later. If TRUE or omitted,
the chart patterns of Microsoft Excel version 5.0 or later will be used.
If FALSE, the chart patterns of Microsoft Excel version 4.0 will be
used.

Shadow    is a logical value corresponding to the Shadow check box.
Shadow does not apply to area charts or bars in bar charts. If shadow is
TRUE, Microsoft Excel adds a shadow to the selected item; if FALSE,
Microsoft Excel removes the shadow, if one is present, from the selected
item. If the selection is an arc, shadow is ignored.

Smooth    is a logical value that applies smoothing to picture markers
in line or xy (scatter) charts. The default is FALSE.

Tlabel    is a number from 1 to 4 specifying the position of tick
labels.

|                  |                            |
| ---------------- | -------------------------- |
| **If tlabel is** | **Tick label position is** |
| 1                | None                       |
| 2                | Low                        |
| 3                | High                       |
| 4                | Next to axis               |

Tmajor    is a number from 1 to 4 specifying the type of major tick
marks.

|                  |                                 |
| ---------------- | ------------------------------- |
| **If tmajor is** | **Type of major tick marks is** |
| 1                | None                            |
| 2                | Inside                          |
| 3                | Outside                         |
| 4                | Cross                           |

Tminor    is a number from 1 to 4 specifying the type of minor tick
marks.

|                  |                                 |
| ---------------- | ------------------------------- |
| **If tminor is** | **Type of minor tick marks is** |
| 1                | None                            |
| 2                | Inside                          |
| 3                | Outside                         |
| 4                | Cross                           |

Type    is a number from 1 to 3 specifying the type of pictures to use
in a picture chart.

|                |                                                                                           |
| -------------- | ----------------------------------------------------------------------------------------- |
| **If type is** | **Pictures should be**                                                                    |
| 1              | Stretched to reach a particular value                                                     |
| 2              | Stacked on top of each other to reach a particular value                                  |
| 3              | Stacked on top of each other, but you specify the number of units each picture represents |

**Remarks**

  - > You can select many graphic objects on a sheet or macro sheet and
    > apply formatting to them at the same time, but you can select only
    > one chart item at a time.

  - > If you select multiple objects and if one or more of the objects
    > requires a different form of the PATTERNS function, then choose
    > the syntax corresponding to the object with the most formatting
    > attributes—that is, the syntax with the most arguments. If you
    > specify an argument that does not apply to an item, the argument
    > has no effect on that item.

  - > To apply formatting to similar items on a chart, use the apply
    > argument described above.

>  

**Related Functions**

[FONT.PROPERTIES](FONT.PROPERTIES.md)   Applies a font to the selection

[FORMAT.TEXT](FORMAT.TEXT.md)   Formats a text box or a chart text item



Return to [README.md](README.md)

