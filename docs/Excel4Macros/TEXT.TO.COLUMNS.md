TEXT.TO.COLUMNS

Equivalent to clicking the Text To Columns command on the Data menu when
a column of data is to be separated into multiple columns. Parses text
into columns of data.

**Syntax**

**TEXT.TO.COLUMNS**(destination\_ref, data\_type, text\_delim,
consecutive\_delim, tab, semicolon, comma, space, other, other\_char,
field\_info)

The following arguments correspond to check boxes, option buttons and
text buttons in the Text To Columns Wizard, which is started with the
Text To Columns command on the Data menu.

Destination\_Ref    is a single cell reference to specifies where to
place the converted text.

Data\_Type    is a number specifying whether that data is delimited (1)
or fixed width (2)

Text\_Delim    denotes how text strings are represented, and can be one
of the following values:

|            |                 |
| ---------- | --------------- |
| **Number** | **Text\_Delim** |
| 1          | "               |
| 2          | '               |
| 3          | none            |

Consecutive\_delim    is a logical value corresponding to the Treat
Consecutive Delimiters as One check box which, if TRUE, allows
consecutive delimiters (such as ",,,") to be treated as a single
delimiter. If FALSE, all consecutive delimiters are considered separate
column breaks.

Tab    is a logical value which, if TRUE, specifies that the column has
tab delimiters. If FALSE, the column does not have tab delimiters. Tab
is ignored if data\_type is 2.

Semicolon    is a logical value which, if TRUE, specifies that the
column has semicolon delimiters. If FALSE, the column does not have
semicolon delimiters. Semicolon is ignored if data\_type is 2.

Comma    is a logical value which, if TRUE, specifies that the column
has comma delimiters. If FALSE, the column does not have comma
delimiters. Comma is ignored if data\_type is 2.

Space    is a logical value which, if TRUE, specifies that the column
has space delimiters. If FALSE, the column does not have space
delimiters. Space is ignored if data\_type is 2.

Other    is a logical value which, if TRUE, specifies that the column
has custom delimiters. If FALSE, the column does not have custom
delimiters. Other is ignored if data\_type is 2.

Other\_Char    is a single character that specifies a delimiter not
included in the list of delimiters. Other\_Char is ignored if data\_type
is 2 and if the argument other is FALSE.

Field\_Info    is an array which consists of the following elements:
"column number, data\_format", if data\_type is 1; or "start\_pos,
data\_format" if data\_type is 2. The second number defines the column's
data format, and can be one of the following.

|                           |                             |
| ------------------------- | --------------------------- |
| **2<sup>nd</sup> Number** | **Data Format**             |
| 1                         | General                     |
| 2                         | Text                        |
| 3                         | Date, in the form MDY       |
| 4                         | Date, in the form DMY       |
| 5                         | Date, in the form YMD       |
| 6                         | Date, in the form MYD       |
| 7                         | Date, in the form DYM       |
| 8                         | Date, in the form YDM       |
| 9                         | Do not import column (skip) |


