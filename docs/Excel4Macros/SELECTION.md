# SELECTION

Returns the reference or object identifier of the selection as an
external reference. Use SELECTION to return information about the
current selection for use in other macro formulas.

**Syntax**

**SELECTION**( )

If a cell or range of cells is selected, Microsoft Excel returns the
corresponding external reference. If an object is selected, Microsoft
Excel returns the object identifier listed in the following table.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Item selected</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Identifier returned</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Imported graphic</p>
</blockquote></td>
<td><blockquote>
<p>Picture n</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Linked graphic</p>
</blockquote></td>
<td><blockquote>
<p>Picture n</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Chart picture</p>
</blockquote></td>
<td><blockquote>
<p>Picture n</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Linked chart</p>
</blockquote></td>
<td><blockquote>
<p>Chart n</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Range</p>
</blockquote></td>
<td><blockquote>
<p>Picture n</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Linked range</p>
</blockquote></td>
<td><blockquote>
<p>Picture n</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Text box</p>
</blockquote></td>
<td><blockquote>
<p>Text n</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Button</p>
</blockquote></td>
<td><blockquote>
<p>Button n</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Rectangle</p>
</blockquote></td>
<td><blockquote>
<p>Rectangle n</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Oval</p>
</blockquote></td>
<td><blockquote>
<p>Oval n</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Line</p>
</blockquote></td>
<td><blockquote>
<p>Line n</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Arc</p>
</blockquote></td>
<td><blockquote>
<p>Arc n</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Group</p>
</blockquote></td>
<td><blockquote>
<p>Group n</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Freehand drawing or polygon</p>
</blockquote></td>
<td><blockquote>
<p>Drawing n</p>
</blockquote></td>
</tr>
</tbody>
</table>

SELECTION also returns the identifiers of chart items. The identifiers
returned are the same as the identifiers you specify when you use the
SELECT function. For a list of these identifiers, see the description of
item\_text in SELECT.

If you select cells and use the value returned by SELECTION in a
function or operation, you usually get the value contained in the
selection instead of its reference. References are automatically
converted to the contents of the reference. If you want to work with the
actual reference, use SET.NAME to assign a name to it, even if the
reference refers to objects. See the last example following. You can
also use the REFTEXT function to convert the reference to text, which
you can then store or manipulate.

**Remarks**

  - > If an object is selected, SELECTION returns the identifier of the
    > object. If multiple objects are selected, it returns the
    > identifiers of all the selected objects, as a string separated by
    > commas.

  - > If more than 1024 characters would be returned, SELECTION returns
    > the \#VALUE\! error value.


**Examples**

If the sheet in the active window is named SHEET1 in the workbook BOOK1,
and if A1:A3 is the selection, then:

SELECTION() equals \[BOOK1\]SHEET1\!A1:A3

The following macro formula moves the current selection one row down:

SELECT(OFFSET(SELECTION(), 1, 0))

The above formula is particularly useful for moving incrementally
through a database to add or modify records.

The following macro formula defines the name "EntryRange" on the active
sheet to refer to one row below the current selection on the active
sheet:

DEFINE.NAME("EntryRange", OFFSET(SELECTION(), 1, 0))

The following macro formula defines the name "Objects" on your macro
sheet to refer to the object names in the current multiple selection:

SET.NAME("Objects", SELECTION())

**Related Functions**

[ACTIVE.CELL](ACTIVE.CELL.md)&nbsp;&nbsp;&nbsp;Returns the reference of the active cell

[SELECT](SELECT.md)&nbsp;&nbsp;&nbsp;Selects a cell, graphic object, or chart



Return to [README](README.md#S)

# SELECTION

Returns the reference or object identifier of the selection as an
external reference. Use SELECTION to return information about the
current selection for use in other macro formulas.

**Syntax**

**SELECTION**( )

If a cell or range of cells is selected, Microsoft Excel returns the
corresponding external reference. If an object is selected, Microsoft
Excel returns the object identifier listed in the following table.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Item selected</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Identifier returned</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Imported graphic</p>
</blockquote></td>
<td><blockquote>
<p>Picture n</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Linked graphic</p>
</blockquote></td>
<td><blockquote>
<p>Picture n</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Chart picture</p>
</blockquote></td>
<td><blockquote>
<p>Picture n</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Linked chart</p>
</blockquote></td>
<td><blockquote>
<p>Chart n</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Range</p>
</blockquote></td>
<td><blockquote>
<p>Picture n</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Linked range</p>
</blockquote></td>
<td><blockquote>
<p>Picture n</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Text box</p>
</blockquote></td>
<td><blockquote>
<p>Text n</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Button</p>
</blockquote></td>
<td><blockquote>
<p>Button n</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Rectangle</p>
</blockquote></td>
<td><blockquote>
<p>Rectangle n</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Oval</p>
</blockquote></td>
<td><blockquote>
<p>Oval n</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Line</p>
</blockquote></td>
<td><blockquote>
<p>Line n</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Arc</p>
</blockquote></td>
<td><blockquote>
<p>Arc n</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Group</p>
</blockquote></td>
<td><blockquote>
<p>Group n</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Freehand drawing or polygon</p>
</blockquote></td>
<td><blockquote>
<p>Drawing n</p>
</blockquote></td>
</tr>
</tbody>
</table>

SELECTION also returns the identifiers of chart items. The identifiers
returned are the same as the identifiers you specify when you use the
SELECT function. For a list of these identifiers, see the description of
item\_text in SELECT.

If you select cells and use the value returned by SELECTION in a
function or operation, you usually get the value contained in the
selection instead of its reference. References are automatically
converted to the contents of the reference. If you want to work with the
actual reference, use SET.NAME to assign a name to it, even if the
reference refers to objects. See the last example following. You can
also use the REFTEXT function to convert the reference to text, which
you can then store or manipulate.

**Remarks**

  - > If an object is selected, SELECTION returns the identifier of the
    > object. If multiple objects are selected, it returns the
    > identifiers of all the selected objects, as a string separated by
    > commas.

  - > If more than 1024 characters would be returned, SELECTION returns
    > the \#VALUE\! error value.


**Examples**

If the sheet in the active window is named SHEET1 in the workbook BOOK1,
and if A1:A3 is the selection, then:

SELECTION() equals \[BOOK1\]SHEET1\!A1:A3

The following macro formula moves the current selection one row down:

SELECT(OFFSET(SELECTION(), 1, 0))

The above formula is particularly useful for moving incrementally
through a database to add or modify records.

The following macro formula defines the name "EntryRange" on the active
sheet to refer to one row below the current selection on the active
sheet:

DEFINE.NAME("EntryRange", OFFSET(SELECTION(), 1, 0))

The following macro formula defines the name "Objects" on your macro
sheet to refer to the object names in the current multiple selection:

SET.NAME("Objects", SELECTION())

**Related Functions**

[ACTIVE.CELL](ACTIVE.CELL.md)&nbsp;&nbsp;&nbsp;Returns the reference of the active cell

[SELECT](SELECT.md)&nbsp;&nbsp;&nbsp;Selects a cell, graphic object, or chart



Return to [README](README.md#S)

