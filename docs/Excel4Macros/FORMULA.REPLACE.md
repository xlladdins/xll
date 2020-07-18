FORMULA.REPLACE

Equivalent to clicking the Replace command on the Edit menu. Finds and
replaces characters in cells on your worksheet.

**Syntax**

**FORMULA.REPLACE**(**find\_text, replace\_text**, look\_at, look\_by,
active\_cell, match\_case)

**FORMULA.REPLACE**?(find\_text, replace\_text, look\_at, look\_by,
active\_cell, match\_case)

Find\_text&nbsp;&nbsp;&nbsp;&nbsp;is the text you want to find. You can
use the wildcard characters, question mark (?) and asterisk (\*), in
find\_text. A question mark matches any single character; an asterisk
matches any sequence of characters. If you want to find an actual
question mark or asterisk, type a tilde (\~) before the character.

Replace\_text&nbsp;&nbsp;&nbsp;&nbsp;is the text you want to replace
find\_text with.

Look\_at&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying whether you want
find\_text to match the entire contents of a cell or any string of
matching characters.

|              |                                   |
| ------------ | --------------------------------- |
| **Look\_at** | **Looks for find\_text**          |
| 1 or omitted | As the entire contents of a cell  |
| 2            | As part of the contents of a cell |

Look\_by&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying whether to search
horizontally (through rows) or vertically (through columns).

|              |                          |
| ------------ | ------------------------ |
| **Look\_by** | **Looks for find\_text** |
| 1 or omitted | By rows                  |
| 2            | By columns               |

Active\_cell&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying the
cells in which find\_text is to be replaced.

  - > If active\_cell is TRUE, find\_text is replaced in the active cell
    > only.

  - > If active\_cell is FALSE, find\_text is replaced in the entire
    > selection, or, if the selection is a single cell, in the entire
    > sheet.

> &nbsp;

Match\_case&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to
the Match Case check box in the Replace dialog box. If match\_case is
TRUE, Microsoft Excel selects the check box; if FALSE, Microsoft Excel
clears the check box. If match\_case is omitted, the status of the check
box is unchanged.

**Remarks**

  - > In FORMULA.REPLACE?, the dialog-box form of the function, omitted
    > arguments are assumed to be the same arguments used in the
    > previous replace operation. If there was no previous replace
    > operation, omitted text arguments are assumed to be "" (empty
    > text).

  - > The result of FORMULA.REPLACE must be a valid cell entry. For
    > example, you cannot replace "=" with "= =" at the beginning of a
    > formula.

  - > If more than a single cell is selected before you use
    > FORMULA.REPLACE, only the selected cells are searched.

> &nbsp;

**Related Function**

[FORMULA.FIND](FORMULA.FIND.md)&nbsp;&nbsp;&nbsp;Finds text in a workbook



Return to [README](README.md)

