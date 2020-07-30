# OPTIONS.LISTS.ADD

This is the equivalent to clicking the Options command on the Tools menu
and then clicking the Custom Lists tab in the Options diralog box. Used
to add a new custom list.

**Syntax**

**OPTIONS.LISTS.ADD**(**string\_array**)

**OPTIONS.LISTS.ADD**(**import\_ref, by\_row**)

**OPTIONS.LISTS.ADD**?(**import\_ref, list\_num**)

String\_array&nbsp;&nbsp;&nbsp;&nbsp;is an array of strings or cell
reference that contains the custom items in the list, a named cell
reference, or an external reference containing the items of the custom
list to add.

Import\_ref&nbsp;&nbsp;&nbsp;&nbsp;is the reference to the cells that
contain the members of the custom list. If A1:A12 contains the twelve
signs of the Zodiac starting with Aquarius, then this function will add
the contents of these twelve cells as a custom list.

By\_row&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that if TRUE, and if
importing from cells, assumes that the list items are in sequential
rows. If FALSE, assumes that the list items are in columns. If omitted,
Microsoft Excel will try to determine the order of the custom lists
according to the layout of the sheet.

List\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying which list to
activate. If omitted, then New List will be activated.

**Remarks**

  - > To replace an existing custom list, you must first delete it and
    > then add the new list to the end.

  - > If the list already exists, then this function will do nothing.
    > The list is not case sensitive, so "Scorpio" and "scorpio" are
    > treated the same in custom lists.

**Related Functions**

[OPTIONS.VIEW](OPTIONS.VIEW.md)&nbsp;&nbsp;&nbsp;Sets various view settings

[OPTIONS.LISTS.GET](OPTIONS.LISTS.GET.md)&nbsp;&nbsp;&nbsp;Returns contents of custom AutoFill
lists

[OPTIONS.LISTS.DELETE](OPTIONS.LISTS.DELETE.md)&nbsp;&nbsp;&nbsp;Deletes a custom list



Return to [README](README.md#O)

