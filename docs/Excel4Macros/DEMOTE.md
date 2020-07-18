DEMOTE

Equivalent to clicking the Group tool. Demotes, or groups, the selected
rows or columns in an outline. Use DEMOTE to change the configuration of
an outline by grouping rows or columns of information.

**Syntax**

**DEMOTE**(**row\_col**)

**DEMOTE**?(row\_col)

Row\_col**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;specifies whether to group rows or
columns.

|              |             |
| ------------ | ----------- |
| **Row\_col** | **Demotes** |
| 1 or omitted | Rows        |
| 2            | Columns     |

**Remarks**

  - > If the selection consists of an entire row or rows, then rows are
    > demoted even if row\_col is 2. Similarly, selection of an entire
    > column overrides row\_col 1.

  - > If the selection is unambiguous (an entire row or column), then
    > DEMOTE? will not display the dialog box.


**Related Functions**

[PROMOTE](PROMOTE.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Promotes the selection in an outline

[SHOW.DETAIL](SHOW.DETAIL.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Expands or collapses a portion of an
outline

[SHOW.LEVELS](SHOW.LEVELS.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Displays a specific number of levels of an
outline



Return to [README](README.md)

