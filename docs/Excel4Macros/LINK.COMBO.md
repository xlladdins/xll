LINK.COMBO

Links an edit box and a list box control into a linked combination box
group. The resulting linked controls track each other's selection and
contents. Linked edit and list box combinations are similar to an
editable drop-down list box, except that the list box is permanently
visible and dropped down.

**Syntax**

**LINK.COMBO**(**link\_logical**)

Link\_logical    is a logical value that specifies whether the controls
are linked or unlinked. If TRUE, the controls will become linked. If
FALSE, the controls will be unlinked.

**Remarks**

To use this function, first select the list box and edit box to be
linked or unlinked. You can do this with SELECT("list box 1,Edit box
2").

**Examples**

LINK.COMBO(FALSE) will unlink a list box and an edit box.

**Related Functions**

[ADD.LIST.ITEM](ADD.LIST.ITEM.md)   Adds an item in a list box or drop-down control on a
worksheet or dialog sheet control

[SELECT.LIST.ITEM](SELECT.LIST.ITEM.md)   Selects an item in a list box or in a group box



Return to [README](README.md)

