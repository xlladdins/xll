# GET.PIVOT.FIELD

Returns information about a field in a PivotTable report.

**Syntax**

**GET.PIVOT.FIELD**(type\_num, pivot\_field\_name, pivot\_table\_name)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a value from 1 to 17 that returns
the following types of information:

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Value</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Returns an array of all the items which make up pivot_field_name. The array is made up of text constants, dates or numbers depending on the field.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>Returns an array of all items which are set to show with the pivot_field_name. The array is made up of text constants, dates or numbers depending on the field. The array is returned in the order that the items are displayed in the PivotTable report. If pivot_field_name is a page field, then the array contains only one element, the value corresponding to the active page (this could be all if the All item is showing).</td>
</tr>
<tr class="even">
<td>3</td>
<td>Returns an array of all items which are hidden in the pivot_field_name. The array is made up of text constants, dates or numbers depending on the field. If pivot_field_name is a data field or the data header name, this function returns the #N/A! error value.</td>
</tr>
<tr class="odd">
<td>4</td>
<td><p>Returns an integer describing where the field is displayed in the active PivotTable report (either row or column):</p>
<p>0 = Hidden</p>
<p>1 = Row</p>
<p>2 = Col</p>
<p>3 = Page</p>
<p>4 = Data</p></td>
</tr>
<tr class="even">
<td>5</td>
<td>Returns an array of all items in pivot_field_name that are group parents. The array is made up of text constants, dates or numbers depending on the field. The array is returned in the order which these items appear in the PivotTable report. Returns #N/A if there are no group parents and if the pivot_field_name is a data field or the data field header.</td>
</tr>
<tr class="odd">
<td>6</td>
<td>Returns a number between 0 and 4095 which describes the subtotals attached to the field. The number is the sum of the values associated with each subtotal function. See PIVOT.FIELD.PROPERTIES for a list of all the values associated with subtotal calculations. If the field is showing as a data field or data field header, #N/A! is returned.</td>
</tr>
<tr class="even">
<td>7</td>
<td><p>Returns an integer describing the type of data contained in the field:</p>
<p>0 = Text</p>
<p>1 = Number</p>
<p>2 = Date</p></td>
</tr>
<tr class="odd">
<td>8</td>
<td>Returns an array five columns wide and one row high describing the summary function's custom calculation shown with the specified field (Data field) in the PivotTable report. The array will look as follows: {function, calculation, base field, base item, number format}. If pivot_field_name is not showing in the active PivotTable report as a data field, #N/A! is returned.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns a reference to all of pivot_field_name's items currently showing in the active PivotTable report. If pivot_field_name is hidden, #N/A! is returned. If pivot_field_name is a page field, the reference to the currently showing page item is returned. If pivot_field_name is a data field, a reference to all the data for this field in the PivotTable report is returned. The references are returned as text.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Returns a reference to the header cell for pivot_field_name. If pivot_field_name is a data field, a reference to all the headers in the data row or column is returned. If pivot_field_name is hidden, #N/A! is returned. The reference is returned as text.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Returns the number of grouped fields in the grouped field set which includes pivot_field_name. If pivot_field_name is neither a parent field nor a child field, 1 is returned. If pivot_field_name is a data field or data header name, the function returns the #N/A! error value.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Returns the level of pivot_field_name in the grouped field set which includes pivot_field_name. Returns 1 for the highest level parent field, 2 for its child field, and so on. If pivot_field_name is neither a parent field nor a child field, 1 is returned. If pivot_field_name is a data field or data header name, the function returns the #N/A! error value.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Returns the name of the parent field for pivot_field_name as a text constant. If pivot_field_name is not a child field, #N/A! is returned.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Returns the name of the child field for pivot_field_name as a text constant. If pivot_field_name is not a parent field, #N/A! is returned.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Returns a text constant representing the original name of the field in the data source.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Returns the position of the field among all the other fields in its orientation. For instance, a 1 would be returned if the field was the first row field.</td>
</tr>
<tr class="even">
<td>17</td>
<td>Returns an array of all items in pivot_field_name that are group children. The array is made up of text constants, dates or numbers depending on the field. The array is returned in the order which these items appear in the PivotTable report. Returns #N/A if there are no group children, and if the pivot_field_name is a data field or the data field header.</td>
</tr>
</tbody>
</table>

Pivot\_field\_name&nbsp;&nbsp;&nbsp;&nbsp;is the name of the field that
you want information about. If there is no field named
pivot\_field\_name in the PivotTable report, returns \#VALUE\!.

Pivot\_table\_name&nbsp;&nbsp;&nbsp;&nbsp;is the name of a PivotTable
report containing the field that you want information about. If omitted,
the PivotTable report containing the active cell is used. If the active
cell is not in a PivotTable report, the \#VALUE\! error value is
returned.

**Related Functions**

[GET.PIVOT.ITEM](GET.PIVOT.ITEM.md)&nbsp;&nbsp;&nbsp;Returns information about an item in a
[P](P.md)ivotTable report.

[GET.PIVOT.TABLE](GET.PIVOT.TABLE.md)&nbsp;&nbsp;&nbsp;Returns information about a PivotTable
report.



Return to [README](README.md#G)

# GET.PIVOT.FIELD

Returns information about a field in a PivotTable report.

**Syntax**

**GET.PIVOT.FIELD**(type\_num, pivot\_field\_name, pivot\_table\_name)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a value from 1 to 17 that returns
the following types of information:

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Value</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Returns an array of all the items which make up pivot_field_name. The array is made up of text constants, dates or numbers depending on the field.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>Returns an array of all items which are set to show with the pivot_field_name. The array is made up of text constants, dates or numbers depending on the field. The array is returned in the order that the items are displayed in the PivotTable report. If pivot_field_name is a page field, then the array contains only one element, the value corresponding to the active page (this could be all if the All item is showing).</td>
</tr>
<tr class="even">
<td>3</td>
<td>Returns an array of all items which are hidden in the pivot_field_name. The array is made up of text constants, dates or numbers depending on the field. If pivot_field_name is a data field or the data header name, this function returns the #N/A! error value.</td>
</tr>
<tr class="odd">
<td>4</td>
<td><p>Returns an integer describing where the field is displayed in the active PivotTable report (either row or column):</p>
<p>0 = Hidden</p>
<p>1 = Row</p>
<p>2 = Col</p>
<p>3 = Page</p>
<p>4 = Data</p></td>
</tr>
<tr class="even">
<td>5</td>
<td>Returns an array of all items in pivot_field_name that are group parents. The array is made up of text constants, dates or numbers depending on the field. The array is returned in the order which these items appear in the PivotTable report. Returns #N/A if there are no group parents and if the pivot_field_name is a data field or the data field header.</td>
</tr>
<tr class="odd">
<td>6</td>
<td>Returns a number between 0 and 4095 which describes the subtotals attached to the field. The number is the sum of the values associated with each subtotal function. See PIVOT.FIELD.PROPERTIES for a list of all the values associated with subtotal calculations. If the field is showing as a data field or data field header, #N/A! is returned.</td>
</tr>
<tr class="even">
<td>7</td>
<td><p>Returns an integer describing the type of data contained in the field:</p>
<p>0 = Text</p>
<p>1 = Number</p>
<p>2 = Date</p></td>
</tr>
<tr class="odd">
<td>8</td>
<td>Returns an array five columns wide and one row high describing the summary function's custom calculation shown with the specified field (Data field) in the PivotTable report. The array will look as follows: {function, calculation, base field, base item, number format}. If pivot_field_name is not showing in the active PivotTable report as a data field, #N/A! is returned.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns a reference to all of pivot_field_name's items currently showing in the active PivotTable report. If pivot_field_name is hidden, #N/A! is returned. If pivot_field_name is a page field, the reference to the currently showing page item is returned. If pivot_field_name is a data field, a reference to all the data for this field in the PivotTable report is returned. The references are returned as text.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Returns a reference to the header cell for pivot_field_name. If pivot_field_name is a data field, a reference to all the headers in the data row or column is returned. If pivot_field_name is hidden, #N/A! is returned. The reference is returned as text.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Returns the number of grouped fields in the grouped field set which includes pivot_field_name. If pivot_field_name is neither a parent field nor a child field, 1 is returned. If pivot_field_name is a data field or data header name, the function returns the #N/A! error value.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Returns the level of pivot_field_name in the grouped field set which includes pivot_field_name. Returns 1 for the highest level parent field, 2 for its child field, and so on. If pivot_field_name is neither a parent field nor a child field, 1 is returned. If pivot_field_name is a data field or data header name, the function returns the #N/A! error value.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Returns the name of the parent field for pivot_field_name as a text constant. If pivot_field_name is not a child field, #N/A! is returned.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Returns the name of the child field for pivot_field_name as a text constant. If pivot_field_name is not a parent field, #N/A! is returned.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Returns a text constant representing the original name of the field in the data source.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Returns the position of the field among all the other fields in its orientation. For instance, a 1 would be returned if the field was the first row field.</td>
</tr>
<tr class="even">
<td>17</td>
<td>Returns an array of all items in pivot_field_name that are group children. The array is made up of text constants, dates or numbers depending on the field. The array is returned in the order which these items appear in the PivotTable report. Returns #N/A if there are no group children, and if the pivot_field_name is a data field or the data field header.</td>
</tr>
</tbody>
</table>

Pivot\_field\_name&nbsp;&nbsp;&nbsp;&nbsp;is the name of the field that
you want information about. If there is no field named
pivot\_field\_name in the PivotTable report, returns \#VALUE\!.

Pivot\_table\_name&nbsp;&nbsp;&nbsp;&nbsp;is the name of a PivotTable
report containing the field that you want information about. If omitted,
the PivotTable report containing the active cell is used. If the active
cell is not in a PivotTable report, the \#VALUE\! error value is
returned.

**Related Functions**

[GET.PIVOT.ITEM](GET.PIVOT.ITEM.md)&nbsp;&nbsp;&nbsp;Returns information about an item in a
[P](P.md)ivotTable report.

[GET.PIVOT.TABLE](GET.PIVOT.TABLE.md)&nbsp;&nbsp;&nbsp;Returns information about a PivotTable
report.



Return to [README](README.md#G)

# GET.PIVOT.FIELD

Returns information about a field in a PivotTable report.

**Syntax**

**GET.PIVOT.FIELD**(type\_num, pivot\_field\_name, pivot\_table\_name)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a value from 1 to 17 that returns
the following types of information:

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Value</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Returns an array of all the items which make up pivot_field_name. The array is made up of text constants, dates or numbers depending on the field.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>Returns an array of all items which are set to show with the pivot_field_name. The array is made up of text constants, dates or numbers depending on the field. The array is returned in the order that the items are displayed in the PivotTable report. If pivot_field_name is a page field, then the array contains only one element, the value corresponding to the active page (this could be all if the All item is showing).</td>
</tr>
<tr class="even">
<td>3</td>
<td>Returns an array of all items which are hidden in the pivot_field_name. The array is made up of text constants, dates or numbers depending on the field. If pivot_field_name is a data field or the data header name, this function returns the #N/A! error value.</td>
</tr>
<tr class="odd">
<td>4</td>
<td><p>Returns an integer describing where the field is displayed in the active PivotTable report (either row or column):</p>
<p>0 = Hidden</p>
<p>1 = Row</p>
<p>2 = Col</p>
<p>3 = Page</p>
<p>4 = Data</p></td>
</tr>
<tr class="even">
<td>5</td>
<td>Returns an array of all items in pivot_field_name that are group parents. The array is made up of text constants, dates or numbers depending on the field. The array is returned in the order which these items appear in the PivotTable report. Returns #N/A if there are no group parents and if the pivot_field_name is a data field or the data field header.</td>
</tr>
<tr class="odd">
<td>6</td>
<td>Returns a number between 0 and 4095 which describes the subtotals attached to the field. The number is the sum of the values associated with each subtotal function. See PIVOT.FIELD.PROPERTIES for a list of all the values associated with subtotal calculations. If the field is showing as a data field or data field header, #N/A! is returned.</td>
</tr>
<tr class="even">
<td>7</td>
<td><p>Returns an integer describing the type of data contained in the field:</p>
<p>0 = Text</p>
<p>1 = Number</p>
<p>2 = Date</p></td>
</tr>
<tr class="odd">
<td>8</td>
<td>Returns an array five columns wide and one row high describing the summary function's custom calculation shown with the specified field (Data field) in the PivotTable report. The array will look as follows: {function, calculation, base field, base item, number format}. If pivot_field_name is not showing in the active PivotTable report as a data field, #N/A! is returned.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns a reference to all of pivot_field_name's items currently showing in the active PivotTable report. If pivot_field_name is hidden, #N/A! is returned. If pivot_field_name is a page field, the reference to the currently showing page item is returned. If pivot_field_name is a data field, a reference to all the data for this field in the PivotTable report is returned. The references are returned as text.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Returns a reference to the header cell for pivot_field_name. If pivot_field_name is a data field, a reference to all the headers in the data row or column is returned. If pivot_field_name is hidden, #N/A! is returned. The reference is returned as text.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Returns the number of grouped fields in the grouped field set which includes pivot_field_name. If pivot_field_name is neither a parent field nor a child field, 1 is returned. If pivot_field_name is a data field or data header name, the function returns the #N/A! error value.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Returns the level of pivot_field_name in the grouped field set which includes pivot_field_name. Returns 1 for the highest level parent field, 2 for its child field, and so on. If pivot_field_name is neither a parent field nor a child field, 1 is returned. If pivot_field_name is a data field or data header name, the function returns the #N/A! error value.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Returns the name of the parent field for pivot_field_name as a text constant. If pivot_field_name is not a child field, #N/A! is returned.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Returns the name of the child field for pivot_field_name as a text constant. If pivot_field_name is not a parent field, #N/A! is returned.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Returns a text constant representing the original name of the field in the data source.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Returns the position of the field among all the other fields in its orientation. For instance, a 1 would be returned if the field was the first row field.</td>
</tr>
<tr class="even">
<td>17</td>
<td>Returns an array of all items in pivot_field_name that are group children. The array is made up of text constants, dates or numbers depending on the field. The array is returned in the order which these items appear in the PivotTable report. Returns #N/A if there are no group children, and if the pivot_field_name is a data field or the data field header.</td>
</tr>
</tbody>
</table>

Pivot\_field\_name&nbsp;&nbsp;&nbsp;&nbsp;is the name of the field that
you want information about. If there is no field named
pivot\_field\_name in the PivotTable report, returns \#VALUE\!.

Pivot\_table\_name&nbsp;&nbsp;&nbsp;&nbsp;is the name of a PivotTable
report containing the field that you want information about. If omitted,
the PivotTable report containing the active cell is used. If the active
cell is not in a PivotTable report, the \#VALUE\! error value is
returned.

**Related Functions**

[GET.PIVOT.ITEM](GET.PIVOT.ITEM.md)&nbsp;&nbsp;&nbsp;Returns information about an item in a
[P](P.md)ivotTable report.

[GET.PIVOT.TABLE](GET.PIVOT.TABLE.md)&nbsp;&nbsp;&nbsp;Returns information about a PivotTable
report.



Return to [README](README.md#G)

