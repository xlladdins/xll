SELECT

Equivalent to selecting cells or changing the active cell. There are
three syntax forms of SELECT. Use syntax 1 to select a cell on a
worksheet or macro sheet; use one of the other syntax forms to select
worksheet or macro sheet objects or chart items.

Syntax 1   Selects cells

Syntax 2   Selects objects on worksheets

Syntax 3   Selects chart objects


SELECT Syntax 1

Equivalent to selecting cells or changing the active cell. There are
three syntax forms of SELECT. Use syntax 1 to select a cell on a
worksheet or macro sheet; use one of the other syntax forms to select
worksheet or macro sheet objects or chart items.

**Syntax**

**SELECT**(selection, active\_cell)

Selection    is the cell or range of cells you want to select. Selection
can be a reference to the active worksheet, such as \!$A$1:$A$3 or
\!Sales, or an R1C1-style reference to a cell or range relative to the
active cell in the current selection, such as
"R\[-1\]C\[-1\]:R\[1\]C\[1\]". The reference must be in text form. If
selection is omitted, the current selection is used.

Active\_cell    is the cell in selection you want to make the active
cell. Active\_cell can be a reference to a single cell on the active
worksheet, such as \!$A$1, or an R1C1-style reference relative to the
active cell, such as "R\[-1\]C\[-1\]". The reference must be in text
form. If active\_cell is omitted, SELECT makes the cell in the
upper-left corner of selection the active cell.

**Remarks**

  - > Active\_cell must be within selection. If it is not, an error
    > message is displayed and SELECT returns the \#VALUE\! error value.

  - > If you are recording a macro using relative references, Microsoft
    > Excel records the action using R1C1-style relative references in
    > the form of text.

  - > If you are recording using absolute references, Microsoft Excel
    > records the action using R1C1-style absolute references in the
    > form of text.

  - > You cannot give an external reference to a specific sheet as the
    > selection argument. The sheet on which you want to make a
    > selection must be active when you use SELECT. Use FORMULA.GOTO to
    > make a selection on another sheet in the same workbook or in
    > another workbook.

>  

**Tip**   You can enter data in a cell without selecting the cell by
using the reference arguments to the CUT, COPY, or FORMULA functions.

**Examples**

The following macro formula selects cells C3:E5 on the active worksheet
and makes C5 the active cell:

SELECT(\!$C$3:$E$5, \!$C$5)

If the active cell is C3, the following macro formula selects cells
E5:G7 and makes cell F6 the active cell in the selection:

SELECT("R\[2\]C\[2\]:R\[4\]C\[4\]", "R\[1\]C\[1\]")

You can also make multiple nonadjacent selections with SELECT. The
following macro formula selects a number of nonadjacent ranges:

SELECT("R1C1, R3C2:R4C3, R8C4:R10C5")

The following sequence of macro formulas moves the active cell right,
left, down, and up within the selection, just as TAB, SHIFT+TAB, ENTER,
and SHIFT+ENTER do:

SELECT(, "RC\[1\]")

SELECT(, "RC\[-1\]")

SELECT(, "R\[1\]C")

SELECT(, "R\[-1\]C")

Use SELECT with the OFFSET function to select a new range a specified
distance away from the current range. For example, the following macro
formula selects a range that is the same size as the current range, one
column over:

SELECT(OFFSET(SELECTION(), 0, 1))

**Related Functions**

[ACTIVE.CELL](ACTIVE.CELL.md)   Returns the reference of the active cell

[SELECT.SPECIAL](SELECT.SPECIAL.md)   Selects a group of cells belonging to a category

[SELECTION](SELECTION.md)   Returns the reference of the selection

[SELECT](SELECT.md) Syntax 2   Selects objects on worksheets

[SELECT](SELECT.md) Syntax 3   Selects chart objects


SELECT Syntax 2

Equivalent to selecting objects on a chart, worksheet, or macro sheet.
There are three syntax forms of SELECT. Use syntax 2 to select an object
on which to perform an action; use one of the other syntax forms to
select cells on a worksheet or macro sheet or items on a chart.

**Syntax**

**SELECT**(object\_id\_text, replace)

Object\_id\_text    is text that identifies the object to select.
Object\_id\_text can be the name of more than one object. To give the
name of more than one object, use the following format:

SELECT("Oval 3, Arc 2, Line 4")

The last item in the object\_id\_text list will be the active object.
The active object is important when moving and sizing a group of
objects. A multiple selection of objects is moved and sized relative to
the upper-left corner of the active object.

Replace    is a logical value that specifies whether previously selected
objects are included in the selection. If replace is TRUE or omitted,
Microsoft Excel only selects the objects specified by object\_id\_text;
if FALSE, it includes any objects that were previously selected. For
example, if a button is selected and a SELECT formula selects an arc and
an oval, TRUE leaves only the arc and oval selected, and FALSE includes
the button with the arc and oval.

**Remarks**

Objects can be identified by their object type and number as described
in CREATE.OBJECT, or by the unique number that specifies the order of
their creation. For example, if the third object you create is an oval,
you could use either "oval 3" or "3" as object\_id\_text.

**Examples**

The following macro formulas each select a number of objects and specify
Arc 2 as the active object:

SELECT("Oval 3, Arc 1, Line 4, Arc 2")

SELECT("3, 1, 4, 2")

**Related Functions**

[FORMAT.MOVE](FORMAT.MOVE.md)   Moves the selected object

[FORMAT.SIZE](FORMAT.SIZE.md)   Changes the size of the selected objects

[GET.OBJECT](GET.OBJECT.md)   Returns information about an object

[SELECTION](SELECTION.md)   Returns the reference of the selection

[SELECT](SELECT.md) Syntax 1   Selects cells

[SELECT](SELECT.md) Syntax 3   Selects chart objects


SELECT Syntax 3

Selects a chart object as specified by the selection code item\_text.
There are three syntax forms of SELECT. Use syntax 3 to select a chart
item to which you want to apply formatting; use one of the other syntax
forms to select cells or objects on a worksheet or macro sheet.

**Syntax**

**SELECT**(item\_text, single\_point)

Item\_text    is a selection code from the following table which
specifies which chart object to select.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>To select</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Item_text</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Entire chart</p>
</blockquote></td>
<td><blockquote>
<p>"Chart"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Plot area</p>
</blockquote></td>
<td><blockquote>
<p>"Plot"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Legend</p>
</blockquote></td>
<td><blockquote>
<p>"Legend"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Primary chart value axis</p>
</blockquote></td>
<td><blockquote>
<p>"Axis 1"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Primary chart category axis</p>
</blockquote></td>
<td><blockquote>
<p>"Axis 2"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Secondary chart value axis or 3-D series axis</p>
</blockquote></td>
<td><blockquote>
<p>"Axis 3"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Secondary chart category axis</p>
</blockquote></td>
<td><blockquote>
<p>"Axis 4"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Chart title</p>
</blockquote></td>
<td><blockquote>
<p>"Title"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Label for the primary chart value axis</p>
</blockquote></td>
<td><blockquote>
<p>"Text Axis 1"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Label for the primary chart category axis</p>
</blockquote></td>
<td><blockquote>
<p>"Text Axis 2"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Label for the primary chart series axis</p>
</blockquote></td>
<td><blockquote>
<p>"Text Axis 3"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>nth floating text item</p>
</blockquote></td>
<td><blockquote>
<p>"Text n"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>nth arrow</p>
</blockquote></td>
<td><blockquote>
<p>"Arrow n"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Major gridlines of value axis</p>
</blockquote></td>
<td><blockquote>
<p>"Gridline 1"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Minor gridlines of value axis</p>
</blockquote></td>
<td><blockquote>
<p>"Gridline 2"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Major gridlines of category axis</p>
</blockquote></td>
<td><blockquote>
<p>"Gridline 3"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Minor gridlines of category axis</p>
</blockquote></td>
<td><blockquote>
<p>"Gridline 4"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Major gridlines of series axis</p>
</blockquote></td>
<td><blockquote>
<p>"Gridline 5"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Minor gridlines of series axis</p>
</blockquote></td>
<td><blockquote>
<p>"Gridline 6"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Primary chart droplines</p>
</blockquote></td>
<td><blockquote>
<p>"Dropline 1"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Secondary chart droplines</p>
</blockquote></td>
<td><blockquote>
<p>"Dropline 2"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Primary chart hi-lo lines</p>
</blockquote></td>
<td><blockquote>
<p>"Hiloline 1"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Secondary chart hi-lo lines</p>
</blockquote></td>
<td><blockquote>
<p>"Hiloline 2"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Primary chart up bar</p>
</blockquote></td>
<td><blockquote>
<p>"UpBar1"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Secondary chart up bar</p>
</blockquote></td>
<td><blockquote>
<p>"UpBar2"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Primary chart down bar</p>
</blockquote></td>
<td><blockquote>
<p>"DownBar1"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Secondary chart down bar</p>
</blockquote></td>
<td><blockquote>
<p>"DownBar2"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Primary chart series line</p>
</blockquote></td>
<td><blockquote>
<p>"Seriesline1"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Secondary chart series line</p>
</blockquote></td>
<td><blockquote>
<p>"Seriesline2"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Entire series</p>
</blockquote></td>
<td><blockquote>
<p>"Sn"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Data associated with point m in series n if single_point is TRUE</p>
</blockquote></td>
<td><blockquote>
<p>"SnPm"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Text attached to point m of series n</p>
</blockquote></td>
<td><blockquote>
<p>"Text SnPm"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Series title text of series n of an area chart</p>
</blockquote></td>
<td><blockquote>
<p>"Text Sn"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Base of a 3-D chart</p>
</blockquote></td>
<td><blockquote>
<p>"Floor"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Back of a 3-D chart</p>
</blockquote></td>
<td><blockquote>
<p>"Walls"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Corners of a 3-D chart</p>
</blockquote></td>
<td><blockquote>
<p>"Corners"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Trend line</p>
</blockquote></td>
<td><blockquote>
<p>"SnTm"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Error bars</p>
</blockquote></td>
<td><blockquote>
<p>"SnEm"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Legend Marker</p>
</blockquote></td>
<td><blockquote>
<p>"Legend Marker n"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Legend Entry</p>
</blockquote></td>
<td><blockquote>
<p>"Legend Entry n"</p>
</blockquote></td>
</tr>
</tbody>
</table>

For trend lines and error bars, the value m can be X or Y, depending on
which point you want to select. If m is blank, selects both.

Single\_point    is a logical value that determines whether to select a
single point. Single\_point is available only when item\_text is "SnPm".

  - > If single\_point is TRUE, Microsoft Excel selects a single point.

  - > If single\_point is FALSE or omitted, Microsoft Excel selects a
    > single point if there is only one series in the chart or selects
    > the entire series if there is more than one series in the chart.

  - > If you specify single\_point when item\_text is any value other
    > than "SnPm", SELECT returns an error value.

>  

**Examples**

SELECT("Chart") selects the entire chart.

SELECT("Dropline 2") selects the droplines of an overlay chart.

SELECT("S1P3", TRUE) selects the third point in the first series.

SELECT("Text S1") selects the series title text of the first series in
an area chart.

**Related Functions**

[SELECTION](SELECTION.md)   Returns the reference of the selection

[SELECT](SELECT.md) Syntax 1   Selects cells

[SELECT](SELECT.md) Syntax 2   Selects objects on worksheets


