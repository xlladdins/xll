# TEXT.BOX

Replaces characters in a text box or button with the text you specify.

**Syntax**

**TEXT.BOX**(**add\_text**, object\_id\_text, start\_num, num\_chars)

Add\_text&nbsp;&nbsp;&nbsp;&nbsp;is the text you want to add to the text
box or button.

Object\_id\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of the text box or
button to which you want to add text (for example, "Text 1" or "Button
2"). If object\_id\_text is omitted, it is assumed to be the selected
item.

Start\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the position of
the first character you want to replace (or the position at which you
want to insert characters if you do not want to replace any). If
start\_num is omitted, it is assumed to be 1.

Num\_chars&nbsp;&nbsp;&nbsp;&nbsp;is the number of characters you want
to replace. If num\_chars is 0, then no characters are replaced, and
add\_text is inserted starting at the position start\_num. If num\_chars
is omitted, all the characters are replaced.

**Examples**

The following macro formula replaces the first five characters in a text
box named "Text 5" with the text "Net Income":

TEXT.BOX("Net Income", "Text 5", 1, 5)

The following macro formula inserts the words "Account Summary for 1991"
at the beginning of a text box named "Text 6":

TEXT.BOX("Account Summary for 1991", "Text 6", 1, 0)

**Related Functions**

[CREATE.OBJECT](CREATE.OBJECT.md)&nbsp;&nbsp;&nbsp;Creates an object

[FONT.PROPERTIES](FONT.PROPERTIES.md)&nbsp;&nbsp;&nbsp;Applies a font to the selection

[GET.OBJECT](GET.OBJECT.md)&nbsp;&nbsp;&nbsp;Returns information about an object



Return to [README](README.md)

