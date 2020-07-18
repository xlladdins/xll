OUTLINE

Creates an outline and defines settings for automatically creating
outlines.

The first three arguments are logical values corresponding to check
boxes in the Outline dialog box, which appears when you choose the
Settings command from the Group and Outline submenu on the Data menu. If
an argument is TRUE, Microsoft Excel selects the check box; if FALSE,
Microsoft Excel clears the check box. If an argument is omitted, the
check box is left in its current state.

**Syntax**

**OUTLINE**(auto\_styles, row\_dir, col\_dir, create\_apply)

**OUTLINE**?(auto\_styles, row\_dir, col\_dir, create\_apply)

Auto\_styles**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;corresponds to the Automatic Styles
check box.

Row\_dir**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;corresponds to the Summary Rows Below
Detail check box.

Col\_dir**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;corresponds to the Summary Columns To
Right Of Detail check box.

Create\_apply**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is the number 1 or 2 and
corresponds to the Create button and the Apply Styles button.

|                   |                                                                         |
| ----------------- | ----------------------------------------------------------------------- |
| **Create\_apply** | **Result**                                                              |
| 1                 | Creates an outline with the current settings                            |
| 2                 | Applies outlining styles to the selection based on outline levels       |
| Omitted           | Corresponds to choosing the OK button to set the other outline settings |

**Related Functions**

[DEMOTE](DEMOTE.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Demotes the selection in an outline

[PROMOTE](PROMOTE.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Promotes the selection in an outline



Return to [README](README.md)

