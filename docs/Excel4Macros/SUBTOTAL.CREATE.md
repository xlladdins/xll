SUBTOTAL.CREATE

Equivalent to clicking the Subtotals command on the Data menu. Generates
a subtotal in a list or database.

**Syntax**

**SUBTOTAL.CREATE**(**at\_change\_in**, **function\_num**, **total**,
replace, pagebreaks, summary\_below)

**SUBTOTAL.CREATE**?(at\_change\_in, function\_num, total, replace,
pagebreaks, summary\_below)

At\_change\_in&nbsp;&nbsp;&nbsp;&nbsp;is a column offset corresponding
to the At Each Change In text box on the Subtotal dialog box.

Function\_Num&nbsp;&nbsp;&nbsp;&nbsp;is a number corresponding to the
Use Function list box specifying which function you want to use in
subtotaling your data.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Function</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Function_Num</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>SUM</p>
</blockquote></td>
<td><blockquote>
<p>1</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>COUNTA</p>
</blockquote></td>
<td><blockquote>
<p>2</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>AVERAGE</p>
</blockquote></td>
<td><blockquote>
<p>3</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>MAX</p>
</blockquote></td>
<td><blockquote>
<p>4</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>MIN</p>
</blockquote></td>
<td><blockquote>
<p>5</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>PRODUCT</p>
</blockquote></td>
<td><blockquote>
<p>6</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>COUNT</p>
</blockquote></td>
<td><blockquote>
<p>7</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>STDEV</p>
</blockquote></td>
<td><blockquote>
<p>8</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>STDEVP</p>
</blockquote></td>
<td><blockquote>
<p>9</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>VAR</p>
</blockquote></td>
<td><blockquote>
<p>10</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>VARP</p>
</blockquote></td>
<td><blockquote>
<p>11</p>
</blockquote></td>
</tr>
</tbody>
</table>

Total&nbsp;&nbsp;&nbsp;&nbsp;is an array of column offsets corresponding
to the Add Subtotal To list box. Indicates which columns you want
aggregated according to function\_num; for example, {4,5}

Replace&nbsp;&nbsp;&nbsp;&nbsp;is a logical value which, if TRUE, causes
any previous subtotals to be replaced by new subtotals. If FALSE or
omitted, subtotals will not be replaced.

PageBreaks&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to
the Page Break Between Groups check box which, if TRUE, creates a page
break below each subtotal. If FALSE or omitted, does not create a page
break below each subtotal.

Summary\_Below&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding
to the Summary Below Data check box which, if TRUE, places subtotal rows
below the data they refer to, and a grand total at the bottom of the
list. If FALSE, places subtotal rows above the data they refer to, and a
grand total just below the header.

**Related Function**

[SUBTOTAL.REMOVE](SUBTOTAL.REMOVE.md)&nbsp;&nbsp;&nbsp;Removes all previously existing
subtotals and grand totals in a list



Return to [README](README.md)

