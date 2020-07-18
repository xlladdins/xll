MERGE.STYLES

Equivalent to clicking the Merge button in the Style dialog box, which
appears when you click the Style command on the Format menu. Merges all
the styles from another workbook into the active workbook. Use
MERGE.STYLES when you want to import styles from another sheet in
another workbook.

**Syntax**

**MERGE.STYLES**(document\_text)

Document\_text is the name of a sheet in a workbook from which you want
to merge styles into the active workbook.

**Remarks**

  - > If any styles from the workbook being merged have the same name as
    > styles in the active workbook, a dialog box appears asking if you
    > want to replace the existing definitions of the styles with the
    > "merged" definitions of the styles. If you click the Yes button,
    > all the definitions are replaced; if you click the No button, all
    > the original definitions in the active workbook are retained.

  - > When you move a sheet with styles to another workbook with styles,
    > any styles with identical names but conflicting definitions have
    > the sheet name added to the style name.


**Related Functions**

[DEFINE.STYLE](DEFINE.STYLE.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Creates or changes a cell style

[DELETE.STYLE](DELETE.STYLE.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Deletes a cell style



Return to [README](README.md)

