# SUBSCRIBE.TO

Inserts the contents of the edition into the active sheet at the point
of the current selection. Use SUBSCRIBE.TO to incorporate editions
published from other workbooks into your Microsoft Excel worksheets and
macro sheets. SUBSCRIBE.TO returns TRUE if successful.

**Syntax**

**SUBSCRIBE.TO**(**file\_text, format\_num**)

**Important**&nbsp;&nbsp;&nbsp;**nbsp;&nbsp;&nbsp;&nbsp;**nbsp;&nbsp;&nbsp;&nbsp;**nbsp;This function is only available if you
are using Microsoft Excel for the Macintosh with system software version
7.0 or later.

File\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name, as a text string, of the
edition you want to insert into the active sheet. Unless file\_text is
in the current folder, supply the full path of the workbook. If
file\_text cannot be found, SUBSCRIBE.TO returns the \#VALUE\! error
value and interrupts the macro.

**Remarks**

  - > If a single cell is selected, the data from the edition file is
    > placed into as large a range of cells as is required by the data.
    > Data already present in those cells is replaced. If the data is a
    > picture, it is inserted from the upper-left corner of the selected
    > cell.

  - > If a range of cells is selected, and the range is not big enough
    > to contain the edition data, Microsoft Excel displays a dialog box
    > asking if you want to clip the data to fit the range.


Format\_num&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 and specifies
the format type of the file you are subscribing to.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Format_num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Format type</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1 or omitted</p>
</blockquote></td>
<td><blockquote>
<p>Picture</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Text (includes BIFF, VALU, TEXT, and CSV formats)</p>
</blockquote></td>
</tr>
</tbody>
</table>

**Related Functions**

[CREATE.PUBLISHER](CREATE.PUBLISHER.md)&nbsp;&nbsp;&nbsp;Creates a publisher from the selection

[EDITION.OPTIONS](EDITION.OPTIONS.md)&nbsp;&nbsp;&nbsp;Sets publisher and subscriber options

[GET.LINK.INFO](GET.LINK.INFO.md)&nbsp;&nbsp;&nbsp;Returns information about a link



Return to [README](README.md)

