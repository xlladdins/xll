# GET.NOTE

Returns characters from a comment.

**Syntax**

**GET.NOTE**(cell\_ref, start\_char, num\_chars)

Cell\_ref&nbsp;&nbsp;&nbsp;&nbsp;is the cell to which the note is
attached. If cell\_ref is omitted, the comment attached to the active
cell is returned.

Start\_char&nbsp;&nbsp;&nbsp;&nbsp;is the number of the first character
in the comment to return. If start\_char is omitted, it is assumed to be
1, the first character in the comment.

Num\_chars&nbsp;&nbsp;&nbsp;&nbsp;is the number of characters to return.
Num\_chars must be less than or equal to 255. If num\_chars is omitted,
it is assumed to be the length of the comment attached to cell\_ref.

**Examples**

The following macro formula returns the first 200 characters in the
comment attached to cell A3 on the active sheet:

GET.NOTE(\!$A$3, 1, 200)

In Microsoft Excel for Windows, the following macro formula returns the
10th through the 39th characters of the comment attached to cell C2 on
SALES.XLS:

GET.NOTE("\[SALES.XLS\]Sheet1\!R2C3", 10, 30)

In Microsoft Excel for the Macintosh, the following macro formula
returns the 10th through the 39th characters of the comment attached to
cell C2 on SALES:

GET.NOTE("\[SALES\]Sheet1\!R2C3", 10, 30)

Use GET.NOTE with the NOTE function to move the contents of a comment to
a cell or text box or to another comment attached to a cell:

NOTE(GET.NOTE(\!$B$10),ACTIVE.CELL())

**Related Functions**

[GET.CELL](GET.CELL.md)&nbsp;&nbsp;&nbsp;Returns information about the specified cell

[NOTE](NOTE.md)&nbsp;&nbsp;&nbsp;Creates or changes a comment.



Return to [README](README.md#G)

# GET.NOTE

Returns characters from a comment.

**Syntax**

**GET.NOTE**(cell\_ref, start\_char, num\_chars)

Cell\_ref&nbsp;&nbsp;&nbsp;&nbsp;is the cell to which the note is
attached. If cell\_ref is omitted, the comment attached to the active
cell is returned.

Start\_char&nbsp;&nbsp;&nbsp;&nbsp;is the number of the first character
in the comment to return. If start\_char is omitted, it is assumed to be
1, the first character in the comment.

Num\_chars&nbsp;&nbsp;&nbsp;&nbsp;is the number of characters to return.
Num\_chars must be less than or equal to 255. If num\_chars is omitted,
it is assumed to be the length of the comment attached to cell\_ref.

**Examples**

The following macro formula returns the first 200 characters in the
comment attached to cell A3 on the active sheet:

GET.NOTE(\!$A$3, 1, 200)

In Microsoft Excel for Windows, the following macro formula returns the
10th through the 39th characters of the comment attached to cell C2 on
SALES.XLS:

GET.NOTE("\[SALES.XLS\]Sheet1\!R2C3", 10, 30)

In Microsoft Excel for the Macintosh, the following macro formula
returns the 10th through the 39th characters of the comment attached to
cell C2 on SALES:

GET.NOTE("\[SALES\]Sheet1\!R2C3", 10, 30)

Use GET.NOTE with the NOTE function to move the contents of a comment to
a cell or text box or to another comment attached to a cell:

NOTE(GET.NOTE(\!$B$10),ACTIVE.CELL())

**Related Functions**

[GET.CELL](GET.CELL.md)&nbsp;&nbsp;&nbsp;Returns information about the specified cell

[NOTE](NOTE.md)&nbsp;&nbsp;&nbsp;Creates or changes a comment.



Return to [README](README.md#G)

