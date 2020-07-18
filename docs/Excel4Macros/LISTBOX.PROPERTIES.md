LISTBOX.PROPERTIES

Sets the properties of a list box and drop-down controls on a worksheet
or dialog sheet.

**Syntax**

**LISTBOX.PROPERTIES**(range, link, drop\_size, multi\_select,
3d\_shading)

**LISTBOX.PROPERTIES**?(range, link, drop\_size, multi\_select,
3d\_shading)

Range&nbsp;&nbsp;&nbsp;&nbsp;is the cell range that the initial contents
of the list box are taken from. If blank (empty text), the list box is
initially unfilled.

Link&nbsp;&nbsp;&nbsp;&nbsp;is the cell on the sheet to which the list
box is linked, and indicates the numeric position of the currently
selected item in the list box. Whenever an item in the list box is
selected, its numeric position is entered into the linked cell on the
sheet.

Drop\_size&nbsp;&nbsp;&nbsp;&nbsp;is the number of lines shown when a
drop-down control is dropped. This value is ignored when applied to a
non-drop-down list box.

Multi\_select&nbsp;&nbsp;&nbsp;&nbsp;is a number that specifies the
selection mode of the list box. Zero is single selection. 1 is simple
multi-select. 2 is extended multi-select.

3d\_shading&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that specifies
whether the list box appears as 3-D. If TRUE, the list box will appear
as 3-D. If FALSE or omitted, the list box will not be 3-D. This argument
is available for only worksheets.

**Related Functions**

[ADD.LIST.ITEM](ADD.LIST.ITEM.md)&nbsp;&nbsp;&nbsp;Adds an item in a list box or drop-down
control on a worksheet or dialog sheet control

[SELECT.LIST.ITEM](SELECT.LIST.ITEM.md)&nbsp;&nbsp;&nbsp;Selects an item in a list box or in a
group box

[CHECKBOX.PROPERTIES](CHECKBOX.PROPERTIES.md)&nbsp;&nbsp;&nbsp;Sets various properties of check
box and option box controls

[SCROLLBAR.PROPERTIES](SCROLLBAR.PROPERTIES.md)&nbsp;&nbsp;&nbsp;Sets the properties of the scroll
bar and spinner controls

[PUSHBUTTON.PROPERTIES](PUSHBUTTON.PROPERTIES.md)&nbsp;&nbsp;&nbsp;Sets the properties of the push
button control



Return to [README](README.md)

