RANKPERC

Returns a table that contains the ordinal and percent rank of each value
in a data set.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**RANKPERC**(**inprng**, outrng, grouped, labels)

**RANKPERC**?(inprng, outrng, grouped, labels)

Inprng&nbsp;&nbsp;&nbsp;&nbsp;is the input range.

Outrng&nbsp;&nbsp;&nbsp;&nbsp;is the first cell (the upper-left cell) in
the output table or the name, as text, of a new sheet to contain the
output table. If FALSE, blank, or omitted, places the output table in a
new workbook.

Grouped&nbsp;&nbsp;&nbsp;&nbsp;is a text character that indicates
whether the data in the input range is organized by row or column.

  - > If grouped is "C" or omitted, then the data is organized by
    > column.

  - > If grouped is "R", then the data is organized by row.

> &nbsp;

Labels&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that describes where
the labels are located in the input range, as shown in the following
table:

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Labels</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Grouped</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Labels are in</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>TRUE</p>
</blockquote></td>
<td><blockquote>
<p>"C"</p>
</blockquote></td>
<td><blockquote>
<p>First row of the input range.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>TRUE</p>
</blockquote></td>
<td><blockquote>
<p>"R"</p>
</blockquote></td>
<td><blockquote>
<p>First column of the input range.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>FALSE or omitted</p>
</blockquote></td>
<td><blockquote>
<p>(ignored)</p>
</blockquote></td>
<td><blockquote>
<p>No labels. All cells in the input range are data.</p>
</blockquote></td>
</tr>
</tbody>
</table>



Return to [README](README.md)

