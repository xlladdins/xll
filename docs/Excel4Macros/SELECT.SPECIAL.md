# SELECT.SPECIAL

Equivalent to clicking the Go To command on the Edit menu and then
selecting the Special button. Use SELECT.SPECIAL to select groups of
similar cells in one of a variety of categories.

**Syntax**

**SELECT.SPECIAL**(**type\_num**, value\_type, levels)

**SELECT.SPECIAL**?(type\_num, value\_type, levels)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 13 corresponding
to options in the Go To Special dialog box and describes what to select.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Description</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Notes/comments</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Constants</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>Formulas</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>Blanks</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>Current region</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>6</p>
</blockquote></td>
<td><blockquote>
<p>Current array</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>7</p>
</blockquote></td>
<td><blockquote>
<p>Row differences</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>Column differences</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>9</p>
</blockquote></td>
<td><blockquote>
<p>Precedents</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>10</p>
</blockquote></td>
<td><blockquote>
<p>Dependents</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>11</p>
</blockquote></td>
<td><blockquote>
<p>Last cell</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>12</p>
</blockquote></td>
<td><blockquote>
<p>Visible cells only (outlining)</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>13</p>
</blockquote></td>
<td><blockquote>
<p>All objects</p>
</blockquote></td>
</tr>
</tbody>
</table>

Value\_type&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying which types of
constants or formulas you want to select. Value\_type is available only
when type\_num is 2 or 3.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Value_type</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Selects</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Numbers</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Text</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>Logical values</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>16</p>
</blockquote></td>
<td><blockquote>
<p>Error values</p>
</blockquote></td>
</tr>
</tbody>
</table>

These values can be added to select more than one type. The default for
value\_type is 23, which select all value types.

Levels&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying how precedents and
dependents are selected. Levels is available only when type\_num is 9 or
10. The default is 1.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Levels</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Selects</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Direct only</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>All levels</p>
</blockquote></td>
</tr>
</tbody>
</table>



Return to [README](README.md#S)

# SELECT.SPECIAL

Equivalent to clicking the Go To command on the Edit menu and then
selecting the Special button. Use SELECT.SPECIAL to select groups of
similar cells in one of a variety of categories.

**Syntax**

**SELECT.SPECIAL**(**type\_num**, value\_type, levels)

**SELECT.SPECIAL**?(type\_num, value\_type, levels)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 13 corresponding
to options in the Go To Special dialog box and describes what to select.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Description</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Notes/comments</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Constants</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>Formulas</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>Blanks</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>Current region</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>6</p>
</blockquote></td>
<td><blockquote>
<p>Current array</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>7</p>
</blockquote></td>
<td><blockquote>
<p>Row differences</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>Column differences</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>9</p>
</blockquote></td>
<td><blockquote>
<p>Precedents</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>10</p>
</blockquote></td>
<td><blockquote>
<p>Dependents</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>11</p>
</blockquote></td>
<td><blockquote>
<p>Last cell</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>12</p>
</blockquote></td>
<td><blockquote>
<p>Visible cells only (outlining)</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>13</p>
</blockquote></td>
<td><blockquote>
<p>All objects</p>
</blockquote></td>
</tr>
</tbody>
</table>

Value\_type&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying which types of
constants or formulas you want to select. Value\_type is available only
when type\_num is 2 or 3.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Value_type</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Selects</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Numbers</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Text</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>Logical values</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>16</p>
</blockquote></td>
<td><blockquote>
<p>Error values</p>
</blockquote></td>
</tr>
</tbody>
</table>

These values can be added to select more than one type. The default for
value\_type is 23, which select all value types.

Levels&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying how precedents and
dependents are selected. Levels is available only when type\_num is 9 or
10. The default is 1.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Levels</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Selects</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Direct only</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>All levels</p>
</blockquote></td>
</tr>
</tbody>
</table>



Return to [README](README.md#S)

# SELECT.SPECIAL

Equivalent to clicking the Go To command on the Edit menu and then
selecting the Special button. Use SELECT.SPECIAL to select groups of
similar cells in one of a variety of categories.

**Syntax**

**SELECT.SPECIAL**(**type\_num**, value\_type, levels)

**SELECT.SPECIAL**?(type\_num, value\_type, levels)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 13 corresponding
to options in the Go To Special dialog box and describes what to select.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Description</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Notes/comments</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Constants</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>Formulas</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>Blanks</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>Current region</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>6</p>
</blockquote></td>
<td><blockquote>
<p>Current array</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>7</p>
</blockquote></td>
<td><blockquote>
<p>Row differences</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>Column differences</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>9</p>
</blockquote></td>
<td><blockquote>
<p>Precedents</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>10</p>
</blockquote></td>
<td><blockquote>
<p>Dependents</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>11</p>
</blockquote></td>
<td><blockquote>
<p>Last cell</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>12</p>
</blockquote></td>
<td><blockquote>
<p>Visible cells only (outlining)</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>13</p>
</blockquote></td>
<td><blockquote>
<p>All objects</p>
</blockquote></td>
</tr>
</tbody>
</table>

Value\_type&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying which types of
constants or formulas you want to select. Value\_type is available only
when type\_num is 2 or 3.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Value_type</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Selects</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Numbers</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Text</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>Logical values</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>16</p>
</blockquote></td>
<td><blockquote>
<p>Error values</p>
</blockquote></td>
</tr>
</tbody>
</table>

These values can be added to select more than one type. The default for
value\_type is 23, which select all value types.

Levels&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying how precedents and
dependents are selected. Levels is available only when type\_num is 9 or
10. The default is 1.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Levels</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Selects</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Direct only</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>All levels</p>
</blockquote></td>
</tr>
</tbody>
</table>



Return to [README](README.md#S)

