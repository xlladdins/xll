SET.CONTROL.VALUE

Changes the value for the active control, such as a list box, drop-down
box, check box, option button, scroll bar, and spinner button.

**Syntax**

**SET.CONTROL.VALUE**(value)

Value    is the value you want to change. The control interprets this
value as follows:

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Control</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Value is</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>List box</p>
</blockquote></td>
<td><blockquote>
<p>The index of the selected item. If zero, then no item is selected.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Drop-down box</p>
</blockquote></td>
<td><blockquote>
<p>The index of the selected item. If zero, then no item is selected.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Check box</p>
</blockquote></td>
<td><blockquote>
<p>0 = Off<br />
1 = On<br />
2 = Mixed</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Option button</p>
</blockquote></td>
<td><blockquote>
<p>0= Off<br />
1 = On</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Scroll bar</p>
</blockquote></td>
<td><blockquote>
<p>The numeric value of the control, between the maximum and minimum values</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Spinner button</p>
</blockquote></td>
<td><blockquote>
<p>The numeric value of the control, between the maximum and minimum values</p>
</blockquote></td>
</tr>
</tbody>
</table>

**Related Functions**

[ADD.LIST.ITEM](ADD.LIST.ITEM.md)   Adds an item in a list box or drop-down control on a
worksheet or dialog sheet control

[REMOVE.LIST.ITEM](REMOVE.LIST.ITEM.md)   Removes an item in a list box or drop-down box

[SELECT.LIST.ITEM](SELECT.LIST.ITEM.md)   Selects an item in a list box or in a group box

[CHECKBOX.PROPERTIES](CHECKBOX.PROPERTIES.md)   Sets various properties of check box and option
box controls

[SCROLLBAR.PROPERTIES](SCROLLBAR.PROPERTIES.md)   Sets the properties of the scroll bar and spinner
controls


