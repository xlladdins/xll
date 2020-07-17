ROW.HEIGHT

Equivalent to choosing the Height command on the Row submenu of the
Format menu. Changes the height of the rows in a reference.

**Syntax**

**ROW.HEIGHT**(height\_num, reference, standard\_height, type\_num)

**ROW.HEIGHT**?(height\_num, reference, standard\_height, type\_num)

Height\_num    specifies how high you want the rows to be in points. If
standard\_height is TRUE, height\_num is ignored.

Reference    specifies the rows for which you want to change the height.

  - > If reference is omitted, the reference is assumed to be the
    > current selection.

  - > If reference is specified, it must be either an external reference
    > to the active worksheet, such as \!$2:$4 or \!Database, or an
    > R1C1-style reference in the form of text or a name, such as
    > "R1:R3", "R\[-4\]:R\[-2\]", or Database.

  - > If reference is a relative R1C1-style reference in the form of
    > text, it is assumed to be relative to the active cell.

>  

Standard\_height    is a logical value that sets the row height as
determined by the font in each row.

  - > If standard\_height is TRUE, Microsoft Excel sets the row height
    > to a standard height that may vary from row to row depending on
    > the fonts used in each row, ignoring height\_num.

  - > If standard\_height is FALSE or omitted, Microsoft Excel sets the
    > row height according to height\_num.

>  

Type\_num    is a number from 1 to 3 corresponding to selecting the
Hide, Unhide, or AutoFit commands from the Row submenu.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Action taken</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Hides the row selection by setting the row height to 0</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Unhides the row selection by setting the row height to the value set before the selection was hidden</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>Sets the row selection to an AutoFit height, which varies from row to row depending on how large the font is in any cell in each row or on how many lines of text are wrapped</p>
</blockquote></td>
</tr>
</tbody>
</table>

**Remarks**

  - > If any of the argument settings conflict, such as when
    > standard\_height is TRUE and type\_num is 3, Microsoft Excel uses
    > the type\_num argument and ignores any arguments that conflict
    > with type\_num.

  - > If you are recording a macro while using a mouse, and you change
    > row heights by dragging the row border, Microsoft Excel records
    > the reference of the rows using R1C1-style references in the form
    > of text. If Uses Relative References is selected, Microsoft Excel
    > uses R1C1-style relative references. If Uses Relative References
    > is not selected, Microsoft Excel uses R1C1-style absolute
    > references.

>  

**Related Function**

[COLUMN.WIDTH](COLUMN.WIDTH.md)   Sets the widths of the specified columns



Return to [README](README.md)

