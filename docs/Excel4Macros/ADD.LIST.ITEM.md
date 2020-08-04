# ADD.LIST.ITEM

Adds an item in a list box or drop-down control on a worksheet or dialog
sheet control.

**Syntax**

**ADD.LIST.ITEM**(**text**, index\_num)

Text&nbsp;&nbsp;&nbsp;&nbsp;specifies the text of the item to be added.
Instead of text, an empty string may be inserted.

Index\_num&nbsp;&nbsp;&nbsp;&nbsp;is the list index to be used for the
new item. Blank entries are created from the end of the current list to
the new item index. If index\_num is omitted the new item is appended to
the list.

**Remarks**

If the list box or drop-down box was already filled using the
LISTBOX.PROPERTIES function, then adding an item with ADD.LIST.ITEM
causes the fillrange contents to be discarded in favor of the new list.

**Related Functions**

[REMOVE.LIST.ITEM](REMOVE.LIST.ITEM.md)&nbsp;&nbsp;&nbsp;Removes an item in a list box or
drop-down box

[SELECT.LIST.ITEM](SELECT.LIST.ITEM.md)&nbsp;&nbsp;&nbsp;Selects an item in a list box or in a
group box



Return to [README](README.md#A)

