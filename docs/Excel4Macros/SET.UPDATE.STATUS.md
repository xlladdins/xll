# SET.UPDATE.STATUS

Sets the update status of a link to automatic or manual. Use
SET.UPDATE.STATUS to change the way a link is updated.

**Syntax**

**SET.UPDATE.STATUS**(**link\_text, status**, type\_of\_link)

Link\_text&nbsp;&nbsp;&nbsp;&nbsp;is the path of the linked file for
which you want to change the update status.

Status&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 and describes how you
want the link to be updated.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Status</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Update method</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Automatic</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Manual</p>
</blockquote></td>
</tr>
</tbody>
</table>

Type\_of\_link&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 4 that
specifies what type of link you want to get information about.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_of_link</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Link document type</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Not available</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>DDE/OLE link</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>Not available</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>Not available</p>
</blockquote></td>
</tr>
</tbody>
</table>

**Example**

In Microsoft Excel for Windows, the following macro formula sets the
update status of the DDE link to Microsoft Word for Windows to manual:

SET.UPDATE.STATUS("WordDocument|'C:\\MEMO.DOC'\!DDE.LINK1", 2, 2)

**Related Functions**

[GET.LINK.INFO](GET.LINK.INFO.md)&nbsp;&nbsp;&nbsp;Returns information about a link

[UPDATE.LINK](UPDATE.LINK.md)&nbsp;&nbsp;&nbsp;Updates a link to another document



Return to [README](README.md#S)

