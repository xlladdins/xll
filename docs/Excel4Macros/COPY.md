COPY

Equivalent to clicking the Copy command on the Edit menu. Copies and
pastes data or objects.

**Syntax**

**COPY**(from\_reference, to\_reference)

From\_reference    is a reference to the cell or range of cells you want
to copy. If from\_reference is omitted, it is assumed to be the current
selection.

To\_reference    is a reference to the cell or range of cells where you
want to paste what you have copied.

  - > To\_reference should be a single cell or an enlarged multiple of
    > from\_reference. For example, if from\_reference is a 2 by 4
    > rectangle, to\_reference can be a 4 by 8 rectangle.

  - > To\_reference can be omitted so that you can subsequently paste
    > using the PASTE, PASTE.LINK, or PASTE.SPECIAL functions.

>  

**Related Functions**

[CUT](CUT.md)   Cuts or moves data or objects

[PASTE](PASTE.md)   Pastes cut or copied data

[PASTE.LINK](PASTE.LINK.md)   Pastes copied data or objects and establishes a link to the
source of the data or object

[PASTE.SPECIAL](PASTE.SPECIAL.md)   Pastes specific components of copied data



Return to [README](README.md)

