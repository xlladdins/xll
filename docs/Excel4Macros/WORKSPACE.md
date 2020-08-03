# WORKSPACE

Changes the workspace settings for a workbook. This function is provide
for compatibility with Microsoft Excel version 4.0 only. In Microsoft
Excel version 5.0 and later versions, you can change workbook settings
with OPTIONS.GENERAL function.

**Syntax**

**WORKSPACE**(fixed, decimals, r1c1, scroll, status, formula, menu\_key,
remote, entermove, underlines, tools, notes, nav\_keys,
menu\_key\_action, drag\_drop, show\_info)

**WORKSPACE**?(fixed, decimals, r1c1, scroll, status, formula,
menu\_key, remote, entermove, underlines, tools, notes, nav\_keys,
menu\_key\_action, drag\_drop, show\_info)

Arguments correspond to check boxes and text boxes in the Workspace
dialog box. Arguments corresponding to check boxes are logical values.
If an argument is TRUE, the check box is selected; if FALSE, the check
box is cleared; if omitted, the current setting is not changed.

Fixed&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Fixed Decimal check box.

Decimals&nbsp;&nbsp;&nbsp;&nbsp;specifies the number of decimal places.
Decimals is ignored if fixed is FALSE or omitted.

R1c1&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the R1C1 check box.

Scroll&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Scroll Bars check box.

Status&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Status Bar check box.

Formula&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Formula Bar check box.

Menu\_key&nbsp;&nbsp;&nbsp;&nbsp;is a text value indicating an alternate
menu key, and corresponds to the Alternate Menu Or Help Key box.

Remote&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Ignore Remote Requests
check box.

**Important**&nbsp;&nbsp;&nbsp;Microsoft Excel for the Macintosh
requires system software version 7.0 or later for this argument.

Entermove&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Move Selection After
Enter/Return check box.

Underlines&nbsp;&nbsp;&nbsp;&nbsp;is a number corresponding to the
Command Underline options as shown in the following table.

**Note**&nbsp;&nbsp;&nbsp;This argument is only available in Microsoft
Excel for the Macintosh.

|                      |                            |
| -------------------- | -------------------------- |
| **If underlines is** | **Command underlines are** |
| 1                    | On                         |
| 2                    | Off                        |
| 3                    | Automatic                  |

Tools&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If TRUE, the Standard
toolbar is displayed; if FALSE, all visible toolbars are hidden. If
omitted, the current toolbar display is not changed.

Notes&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Note Indicator check
box.

Nav\_keys&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Alternate Navigation
Keys check box. In Microsoft Excel for the Macintosh, nav\_keys is
ignored.

Menu\_key\_action&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 specifying
options for the alternate menu or Help key. In Microsoft Excel for the
Macintosh, menu\_key\_action is ignored.

|                       |                                          |
| --------------------- | ---------------------------------------- |
| **Menu\_key\_action** | **Alternate menu or Help key activates** |
| 1 or omitted          | Microsoft Excel menus                    |
| 2                     | Lotus 1-2-3 Help                         |

Drag\_drop&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Cell Drag And Drop
check box.

Show\_info&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Info Window check
box.

**Related Function**

[GET.WORKSPACE](GET.WORKSPACE.md)&nbsp;&nbsp;&nbsp;Returns information about the workspace



Return to [README](README.md#W)

# WORKSPACE

Changes the workspace settings for a workbook. This function is provide
for compatibility with Microsoft Excel version 4.0 only. In Microsoft
Excel version 5.0 and later versions, you can change workbook settings
with OPTIONS.GENERAL function.

**Syntax**

**WORKSPACE**(fixed, decimals, r1c1, scroll, status, formula, menu\_key,
remote, entermove, underlines, tools, notes, nav\_keys,
menu\_key\_action, drag\_drop, show\_info)

**WORKSPACE**?(fixed, decimals, r1c1, scroll, status, formula,
menu\_key, remote, entermove, underlines, tools, notes, nav\_keys,
menu\_key\_action, drag\_drop, show\_info)

Arguments correspond to check boxes and text boxes in the Workspace
dialog box. Arguments corresponding to check boxes are logical values.
If an argument is TRUE, the check box is selected; if FALSE, the check
box is cleared; if omitted, the current setting is not changed.

Fixed&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Fixed Decimal check box.

Decimals&nbsp;&nbsp;&nbsp;&nbsp;specifies the number of decimal places.
Decimals is ignored if fixed is FALSE or omitted.

R1c1&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the R1C1 check box.

Scroll&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Scroll Bars check box.

Status&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Status Bar check box.

Formula&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Formula Bar check box.

Menu\_key&nbsp;&nbsp;&nbsp;&nbsp;is a text value indicating an alternate
menu key, and corresponds to the Alternate Menu Or Help Key box.

Remote&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Ignore Remote Requests
check box.

**Important**&nbsp;&nbsp;&nbsp;Microsoft Excel for the Macintosh
requires system software version 7.0 or later for this argument.

Entermove&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Move Selection After
Enter/Return check box.

Underlines&nbsp;&nbsp;&nbsp;&nbsp;is a number corresponding to the
Command Underline options as shown in the following table.

**Note**&nbsp;&nbsp;&nbsp;This argument is only available in Microsoft
Excel for the Macintosh.

|                      |                            |
| -------------------- | -------------------------- |
| **If underlines is** | **Command underlines are** |
| 1                    | On                         |
| 2                    | Off                        |
| 3                    | Automatic                  |

Tools&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If TRUE, the Standard
toolbar is displayed; if FALSE, all visible toolbars are hidden. If
omitted, the current toolbar display is not changed.

Notes&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Note Indicator check
box.

Nav\_keys&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Alternate Navigation
Keys check box. In Microsoft Excel for the Macintosh, nav\_keys is
ignored.

Menu\_key\_action&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 specifying
options for the alternate menu or Help key. In Microsoft Excel for the
Macintosh, menu\_key\_action is ignored.

|                       |                                          |
| --------------------- | ---------------------------------------- |
| **Menu\_key\_action** | **Alternate menu or Help key activates** |
| 1 or omitted          | Microsoft Excel menus                    |
| 2                     | Lotus 1-2-3 Help                         |

Drag\_drop&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Cell Drag And Drop
check box.

Show\_info&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Info Window check
box.

**Related Function**

[GET.WORKSPACE](GET.WORKSPACE.md)&nbsp;&nbsp;&nbsp;Returns information about the workspace



Return to [README](README.md#W)

# WORKSPACE

Changes the workspace settings for a workbook. This function is provide
for compatibility with Microsoft Excel version 4.0 only. In Microsoft
Excel version 5.0 and later versions, you can change workbook settings
with OPTIONS.GENERAL function.

**Syntax**

**WORKSPACE**(fixed, decimals, r1c1, scroll, status, formula, menu\_key,
remote, entermove, underlines, tools, notes, nav\_keys,
menu\_key\_action, drag\_drop, show\_info)

**WORKSPACE**?(fixed, decimals, r1c1, scroll, status, formula,
menu\_key, remote, entermove, underlines, tools, notes, nav\_keys,
menu\_key\_action, drag\_drop, show\_info)

Arguments correspond to check boxes and text boxes in the Workspace
dialog box. Arguments corresponding to check boxes are logical values.
If an argument is TRUE, the check box is selected; if FALSE, the check
box is cleared; if omitted, the current setting is not changed.

Fixed&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Fixed Decimal check box.

Decimals&nbsp;&nbsp;&nbsp;&nbsp;specifies the number of decimal places.
Decimals is ignored if fixed is FALSE or omitted.

R1c1&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the R1C1 check box.

Scroll&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Scroll Bars check box.

Status&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Status Bar check box.

Formula&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Formula Bar check box.

Menu\_key&nbsp;&nbsp;&nbsp;&nbsp;is a text value indicating an alternate
menu key, and corresponds to the Alternate Menu Or Help Key box.

Remote&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Ignore Remote Requests
check box.

**Important**&nbsp;&nbsp;&nbsp;Microsoft Excel for the Macintosh
requires system software version 7.0 or later for this argument.

Entermove&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Move Selection After
Enter/Return check box.

Underlines&nbsp;&nbsp;&nbsp;&nbsp;is a number corresponding to the
Command Underline options as shown in the following table.

**Note**&nbsp;&nbsp;&nbsp;This argument is only available in Microsoft
Excel for the Macintosh.

|                      |                            |
| -------------------- | -------------------------- |
| **If underlines is** | **Command underlines are** |
| 1                    | On                         |
| 2                    | Off                        |
| 3                    | Automatic                  |

Tools&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If TRUE, the Standard
toolbar is displayed; if FALSE, all visible toolbars are hidden. If
omitted, the current toolbar display is not changed.

Notes&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Note Indicator check
box.

Nav\_keys&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Alternate Navigation
Keys check box. In Microsoft Excel for the Macintosh, nav\_keys is
ignored.

Menu\_key\_action&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 specifying
options for the alternate menu or Help key. In Microsoft Excel for the
Macintosh, menu\_key\_action is ignored.

|                       |                                          |
| --------------------- | ---------------------------------------- |
| **Menu\_key\_action** | **Alternate menu or Help key activates** |
| 1 or omitted          | Microsoft Excel menus                    |
| 2                     | Lotus 1-2-3 Help                         |

Drag\_drop&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Cell Drag And Drop
check box.

Show\_info&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Info Window check
box.

**Related Function**

[GET.WORKSPACE](GET.WORKSPACE.md)&nbsp;&nbsp;&nbsp;Returns information about the workspace



Return to [README](README.md#W)

