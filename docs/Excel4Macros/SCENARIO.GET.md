# SCENARIO.GET

Returns the specified information about the scenarios defined on your
worksheet.

**Syntax**

**SCENARIO.GET**(**type\_num**, scen\_name)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 8 specifying the
type of information you want.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Information returned</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>A horizontal array of all scenario names in the form of text</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>A reference to the set of changing cells of scen_name (specified in the Changing Cells box of the Scenario Manager dialog box). If scen_name is omitted, the first scenario is used.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>A reference to the result cells (specified in the Result Cells box in the Scenario Summary dialog box)</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>An array of scenario values for the scenario scen_name . Each scenario is in a separate row. If scen_name is omitted, the first scenario is used.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>Comment, as text, for the scenario</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>6</p>
</blockquote></td>
<td><blockquote>
<p>Returns TRUE if the specified scenario is locked to prevent changes; FALSE, if unlocked. Scen_name is required.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>7</p>
</blockquote></td>
<td><blockquote>
<p>Returns TRUE if the specified scenario is hidden; FALSE, if visible to the user. Scen_name is required.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>Returns the user name of the person who last modified the scenario by either adding or editing a scenario. Scen_name is required.</p>
</blockquote></td>
</tr>
</tbody>
</table>

Scen\_name&nbsp;&nbsp;&nbsp;&nbsp;is the name of the scenario that you
want information about. Ignored if type\_num equals 1 or 3.

**Remarks**

In the returned array of scenario values, the number of rows is the
number of scenarios, and the number of columns is the number of changing
cells.



Return to [README](README.md#S)

# SCENARIO.GET

Returns the specified information about the scenarios defined on your
worksheet.

**Syntax**

**SCENARIO.GET**(**type\_num**, scen\_name)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 8 specifying the
type of information you want.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Information returned</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>A horizontal array of all scenario names in the form of text</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>A reference to the set of changing cells of scen_name (specified in the Changing Cells box of the Scenario Manager dialog box). If scen_name is omitted, the first scenario is used.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>A reference to the result cells (specified in the Result Cells box in the Scenario Summary dialog box)</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>An array of scenario values for the scenario scen_name . Each scenario is in a separate row. If scen_name is omitted, the first scenario is used.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>Comment, as text, for the scenario</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>6</p>
</blockquote></td>
<td><blockquote>
<p>Returns TRUE if the specified scenario is locked to prevent changes; FALSE, if unlocked. Scen_name is required.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>7</p>
</blockquote></td>
<td><blockquote>
<p>Returns TRUE if the specified scenario is hidden; FALSE, if visible to the user. Scen_name is required.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>Returns the user name of the person who last modified the scenario by either adding or editing a scenario. Scen_name is required.</p>
</blockquote></td>
</tr>
</tbody>
</table>

Scen\_name&nbsp;&nbsp;&nbsp;&nbsp;is the name of the scenario that you
want information about. Ignored if type\_num equals 1 or 3.

**Remarks**

In the returned array of scenario values, the number of rows is the
number of scenarios, and the number of columns is the number of changing
cells.



Return to [README](README.md#S)

