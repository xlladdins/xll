# NOTE

Equivalent to choosing the Comment command from the Insert menu. Creates
a comment or replaces text characters in a comment.

**Syntax**

**NOTE**(add\_text, cell\_ref, start\_char, num\_chars)

**NOTE**?( )

Add\_text&nbsp;&nbsp;&nbsp;&nbsp;is text of up to 255 characters you
want to add to a comment. Add\_text must be enclosed in quotation marks.

  - > If add\_text is omitted, it is assumed to be "" (empty text).


Cell\_ref&nbsp;&nbsp;&nbsp;&nbsp;is the cell to which you want to add
the comment text. If cell\_ref is omitted, add\_text is added to the
active cell's comment.

Start\_char&nbsp;&nbsp;&nbsp;&nbsp;is the number of the character at
which you want add\_text to be added. If start\_char is omitted, it is
assumed to be 1. If there is no existing comment, start\_char is
ignored.

Num\_chars&nbsp;&nbsp;&nbsp;&nbsp;is the number of characters that you
want to replace in the comment. If num\_chars is omitted, it is assumed
to be equal to the length of the comment.

**Remarks**

  - > NOTE returns the number of the last character entered in the
    > comment. This is useful if you want to know how many characters
    > are in the text string.

  - > The dialog-box form of this function, NOTE?, takes no arguments.

  - > NOTE() deletes the comment attached to the active cell.


To find out if a cell has a comment attached to it, use GET.CELL.

**Related Function**

[GET.NOTE](GET.NOTE.md)&nbsp;&nbsp;&nbsp;Returns characters from a comment



Return to [README](README.md#N)

