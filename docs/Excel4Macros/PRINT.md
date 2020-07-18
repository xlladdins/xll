PRINT

Equivalent to clicking the Print command on the File menu. Prints the
active workbook.

Arguments correspond to options, check boxes, and edit boxes in the
Print dialog box. Arguments corresponding to check boxes are logical
values. If an argument is TRUE, Microsoft Excel selects the check box;
if FALSE, Microsoft Excel clears the check box.

**Syntax**

**PRINT**(range\_num, from, to, copies, draft, preview, print\_what,
color, feed, quality, y\_resolution, selection)

**PRINT**?(range\_num, from, to, copies, draft, preview, print\_what,
color, feed, quality, y\_resolution, selection)

Range\_num**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a number specifying which pages to
print.

|                |                                                                                        |
| -------------- | -------------------------------------------------------------------------------------- |
| **Range\_num** | **Prints the following pages**                                                         |
| 1              | All the pages                                                                          |
| 2              | Prints a specified range. If range\_num is 2, then from and to are required arguments. |

From**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;specifies the first page to print. This
argument is ignored unless range\_num equals 2.

To**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;specifies the last page to print. This
argument is ignored unless range\_num equals 2.

Copies**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;specifies the number of copies to print.
If omitted, the default is 1.

Draft**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;This argument overrides the draft argument
from the PAGE.SETUP function. If omitted, the Draft Setting from the
Page.Setup function is used.

Preview**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a logical value corresponding to the
Print Preview button in the Print dialog box. If TRUE, the print preview
window will be displayed. If FALSE, the window will not be displayed

Print\_what**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a number from 1 to 3 that
specifies what parts of the sheet or macro sheet to print. If a chart is
active, print\_what is ignored. This argument will override the setting
in the Page Setup dialog box. If omitted, the note argument in the
Page.Setup function is used to determine whether to print notes or not.

|                 |                      |
| --------------- | -------------------- |
| **Print\_what** | **Prints**           |
| 1               | Sheet only           |
| 2               | Notes only           |
| 3               | Sheet and then notes |

Color**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;corresponds to the Print Using Color check
box. Color is available only in Microsoft Excel for the Macintosh. If
omitted, the setting is not changed.

Feed**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a number specifying the type of paper
feed. Feed is available only in Microsoft Excel for the Macintosh.

|              |                                   |
| ------------ | --------------------------------- |
| **Feed**     | **Type of paper feed**            |
| 1 or omitted | Continuous (paper cassette)       |
| 2            | Cut sheet or manual (manual feed) |

Quality**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;Specifies the DPI output quality you
want. If omitted, the corresponding settings in the Page Setup dialog
box will be used. If included, this argument overrides the quality
argument in the PAGE SETUP dialog box.

Y\_resolution**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;corresponds to the Print Quality
box in the Page Setup dialog box if you have specified a printer where
the horizontal and vertical resolution are not equal, such as a
dot-matrix printer. If omitted, the corresponding settings in the Page
Setup dialog box will be used. If included, this argument overrides the
print quality setting in the Page Setup dialog box.

Selection**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;specifies what portion of the sheet to
print.

|               |                                                                                                                                                                         |
| ------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Selection** | **Portion printed**                                                                                                                                                     |
| 1             | Prints the current selection from all selected sheets. For example, if A1:F40 is selected on the active sheet, A1:F40 will be printed from each of the selected sheets. |
| 2             | Prints the print area or entire sheet from all selected sheets.                                                                                                         |
| 3             | Prints print area or entire sheet from all sheets in the workbook.                                                                                                      |

**Related Functions**

[PAGE.SETUP](PAGE.SETUP.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Sets page printing specifications

[PRINT.PREVIEW](PRINT.PREVIEW.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Previews pages and page breaks before
printing

[PRINTER.SETUP](PRINTER.SETUP.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Identifies the printer

[SET.PRINT.AREA](SET.PRINT.AREA.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Defines the print area

[SET.PRINT.TITLES](SET.PRINT.TITLES.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Defines text to print as titles

[DEFINE.NAME](DEFINE.NAME.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Equivalent to clicking the Define command
on the Name submenu of the Insert menu



Return to [README](README.md)

