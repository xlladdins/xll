REMOVE.LIST.ITEM

Removes an item in a list box or drop-down box.

**Syntax**

**REMOVE.LIST.ITEM**(**index\_num**, count\_num)

Index\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies the index of the item to
remove, from 1 to the number of items in the list. Specify zero to
remove all items in the list.

Count\_num&nbsp;&nbsp; Specifies the number of items to delete starting
from index\_num. If omitted, only one item is removed.

**Remarks**

If count\_num + index\_num is greater than the number of items in the
list, all items starting with index\_num to the end of the list are
removed.

**Examples**

REMOVE.LIST.ITEM(3,2) removes two items starting with the third item

REMOVE.LIST.ITEM(3) removes only the third item

**Related Function**

[LISTBOX.PROPERTIES](LISTBOX.PROPERTIES.md)&nbsp;&nbsp;&nbsp;Sets the properties of a list box
and drop-down controls on worksheet and dialog sheets



Return to [README](README.md)

