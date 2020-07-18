ROUTING.SLIP

Equivalent to clicking the Add Routing Slip command on the File menu.
Adds or Edits the routing slip attached to the current workbook.

**Syntax**

**ROUTING.SLIP**(recipients,subject, message, route\_num,
return\_logical, status\_logical)

**ROUTING.SLIP**?(recipients,subject, message, route\_num,
return\_logical, status\_logical)

Recipients&nbsp;&nbsp;&nbsp;&nbsp;is the name of the person to whom you
want to send the mail. The name should be given as text.

  - > To specify more than one name, give the list of names as an array.
    > For example, ROUTING.SLIP({"John", "Paul", "George", "Ringo"})
    > would send the active workbook to the four names in the array. You
    > can also refer to a range on a sheet or macro sheet that contains
    > a list of names to whom you want the mail to be sent.

  - > Specifying recipients while a routing is in progress only modifies
    > the non-grayed recipients (that is, those recipients who have not
    > received the message yet). Recipients who have already received,
    > reviewed and forwarded the routed workbook cannot be modified.

Subject&nbsp;&nbsp;&nbsp;&nbsp;is a text string containing the subject
text used for the mail messages used to route the workbook. If omitted,
the default subject line is "Routing: name", where name is the file name
or title as displayed in the Summary Info dialog box, if available.

Message&nbsp;&nbsp;&nbsp;&nbsp;is a text string containing the body text
used for the mail messages used to route the workbook.

Route\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number indicating the type of
routing method.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Route_num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Method</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1 or omitted</p>
</blockquote></td>
<td><blockquote>
<p>One after another routing</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>All at once routing</p>
</blockquote></td>
</tr>
</tbody>
</table>

Return\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value which, if TRUE
or omitted, indicates that the routing should be returned to the
originator when the routing is complete. If FALSE, the routing will end
with the last recipient in the To list box in the Routing Slip Dialog
box.

Status\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding
to the Track Status check box in the Routing Slip dialog box. If TRUE or
omitted, status tracking messages for the routing are sent. FALSE means
that no status tracking is performed.

**Remarks**

  - > If this function is used on a workbook that is already being
    > routed, the route\_num, status\_logical and return\_logical
    > arguments are ignored (they cannot be changed).

  - > When arguments are omitted and a routing slip already exists, the
    > omitted arguments are replaced by the current values of the
    > routing slip.

**Related Functions**

[ROUTE.DOCUMENT](ROUTE.DOCUMENT.md)&nbsp;&nbsp;&nbsp;Routes the workbook using the defined
routing slip information

[SEND.MAIL](SEND.MAIL.md)&nbsp;&nbsp;&nbsp;Sends the active workbook using email



Return to [README](README.md)

