# FORMULA.FIND

Equivalent to clicking the Find command on the Edit menu. Selects the
next or previous cell containing the specified text and returns TRUE. If
a matching cell is not found, FORMULA.FIND returns FALSE and displays a
message.

**Syntax**

**FORMULA.FIND**(**text, in\_num, at\_num, by\_num**, dir\_num,
match\_case)

**FORMULA.FIND**?(text, in\_num, at\_num, by\_num, dir\_num,
match\_case)

Text&nbsp;&nbsp;&nbsp;&nbsp;is the text you want to find. Text
corresponds to the Find What box in the Find dialog box.

In\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 3 specifying where
to search.

|             |              |
| ----------- | ------------ |
| **In\_num** | **Searches** |
| 1           | Formulas     |
| 2           | Values       |
| 3           | Notes        |

At\_num&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 and specifies
whether to find cells containing only text or also cells containing text
within a longer string of characters.

|             |                                                  |
| ----------- | ------------------------------------------------ |
| **At\_num** | **Searches for text as**                         |
| 1           | A whole string (the only value in the cell)      |
| 2           | Either a whole string or part of a longer string |

By\_num&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 and specifies
whether to search by rows or by columns.

|             |                 |
| ----------- | --------------- |
| **By\_num** | **Searches by** |
| 1           | Rows            |
| 2           | Columns         |

Dir\_num&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 and specifies
whether to search for the next or previous occurrence of text.

|              |                                 |
| ------------ | ------------------------------- |
| **Dir\_num** | **Searches for**                |
| 1 or omitted | The next occurrence of text     |
| 2            | The previous occurrence of text |

Match\_case&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to
the Match Case check box in the Find dialog box. If match\_case is TRUE,
Microsoft Excel matches characters exactly, including uppercase and
lowercase; if FALSE or omitted, matching is not case-sensitive.

**Remarks**

  - > In Microsoft Excel for Windows, the dialog-box form of
    > FORMULA.FIND is equivalent to pressing SHIFT+F5.

  - > If more than one cell is selected when you use FORMULA.FIND,
    > Microsoft Excel searches only that selection.




Return to [README](README.md)

