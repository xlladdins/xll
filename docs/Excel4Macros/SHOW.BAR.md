# SHOW.BAR

Displays the specified menu bar. Use SHOW.BAR to display a menu bar you
have created with the ADD.BAR function or to display a built-in
Microsoft Excel 95 or earlier version menu bar.

**Syntax**

**SHOW.BAR**(bar\_num)

Bar\_num&nbsp;&nbsp;&nbsp;&nbsp;is the number of the menu bar you want
to display. It can be the number of one of the Microsoft Excel built-in
menu bars, the number returned by a previously executed ADD.BAR
function, or a reference to a cell containing a previously executed
ADD.BAR function.

If bar\_num is omitted, Microsoft Excel displays the appropriate menu
bar for the active workbook as shown in the following table.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Bar_num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Bar displayed</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>A sheet or macro sheet (Microsoft Excel version 4.0)</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>A chart (Microsoft Excel version 4.0)</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>No active window</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>The Info window (Microsoft Excel 95 or earlier versions)</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>A sheet or macro sheet (short menus)</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>6</p>
</blockquote></td>
<td><blockquote>
<p>A chart (short menus)</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>7</p>
</blockquote></td>
<td><blockquote>
<p>Shortcut menus 1 (for Cells, Workbook tabs, Toolbars, VB Windows)</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>Shortcut menus 2 (for objects)</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>9</p>
</blockquote></td>
<td><blockquote>
<p>Shortcut menus 3 (for chart elements)</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>10</p>
</blockquote></td>
<td><blockquote>
<p>A sheet or macro sheet</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>11</p>
</blockquote></td>
<td><blockquote>
<p>A chart</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>12</p>
</blockquote></td>
<td><blockquote>
<p>A Visual Basic module</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>13-35</p>
</blockquote></td>
<td><blockquote>
<p>Reserved for use by shortcut menus. These numbers will return an error if a macro tries to do anything with them.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>37-51</p>
</blockquote></td>
<td><blockquote>
<p>Custom menu bar for macro use</p>
</blockquote></td>
</tr>
</tbody>
</table>

**Remarks**

  - > When displaying a built-in menu bar, you can display only bars 1
    > or 5 if a sheet or macro sheet is active, bars 2 or 6 if a chart
    > is active, and so on. If you try to display a chart menu bar while
    > a sheet or macro sheet is active, SHOW.BAR returns an error and
    > interrupts the current macro.

  - > Displaying a custom menu bar disables automatic menu-bar switching
    > when different types of sheets are selected. For example, if a
    > custom menu bar is displayed and you switch to a chart, neither of
    > the two chart menus is automatically displayed as it would be when
    > you are using the built-in menu bars. Automatic menu-bar switching
    > is reenabled when a built-in bar is displayed using SHOW.BAR.


**Example**

The following macro formula displays short menus on a worksheet or macro
sheet:

SHOW.BAR(5)

**Related Functions**

[ADD.BAR](ADD.BAR.md)&nbsp;&nbsp;&nbsp;Adds a menu bar

[DELETE.BAR](DELETE.BAR.md)&nbsp;&nbsp;&nbsp;Deletes a menu bar

[SHOW.TOOLBAR](SHOW.TOOLBAR.md)&nbsp;&nbsp;&nbsp;Hides or displays a toolbar



Return to [README](README.md)

