# SET.LIST.ITEM

Sets the text of an item in a list box or drop-down box control.

**Syntax**

**SET.LIST.ITEM**(**text**, **index\_num**)

Text&nbsp;&nbsp;&nbsp;&nbsp;specifies the text of the item to be added.
Instead of text, an empty string may be inserted.

Index\_num&nbsp;&nbsp;&nbsp;&nbsp;is the list index of the item to be
changed, from 1 to the number of items in the list.

**Remarks**

If the list box or drop-down box was already filled using the
LISTBOX.PROPERTIES function, then changing an item with SET.LIST.ITEM
causes the fillrange contents to be discarded, leaving a list with one
non-blank element and index\_num entries.

**Related Functions**

[REMOVE.LIST.ITEM](REMOVE.LIST.ITEM.md)&nbsp;&nbsp;&nbsp;Removes an item in a list box or
drop-down box

[SELECT.LIST.ITEM](SELECT.LIST.ITEM.md)&nbsp;&nbsp;&nbsp;Selects an item in a list box or in a
group box



Return to [README](README.md)

