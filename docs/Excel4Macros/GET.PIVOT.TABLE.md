GET.PIVOT.TABLE

Returns information about a PivotTable report.

**Syntax**

**GET.PIVOT.TABLE**(**type\_num**,pivot\_table\_name)

Type\_num is a value from 1 to 22 that represents a type of information
you want about a PivotTable report.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Information</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Returns the name of the person who last updated the PivotTable report, as a text constant.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>Returns the date the PivotTable report was last updated, as a serial number.</td>
</tr>
<tr class="even">
<td>3</td>
<td>Returns a horizontal array of text constants representing all the fields in the PivotTable report.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>Returns an integer representing the number of fields in the PivotTable report.</td>
</tr>
<tr class="even">
<td>5</td>
<td>Returns a horizontal array of text constants representing all the visible fields in the PivotTable report (rows, columns, pages or data)</td>
</tr>
<tr class="odd">
<td>6</td>
<td>Returns a horizontal array of text constants representing all the hidden fields in the PivotTable report. Return #N/A if no hidden fields.</td>
</tr>
<tr class="even">
<td>7</td>
<td>Returns a horizontal array of text constants representing the names of all the fields currently showing in the PivotTable report as row fields. Returns #N/A if there are no row fields.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as column fields. Returns #N/A if no column fields exist.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as page fields. Return #N/A if no page fields exist.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as data fields. Returns #N/A if there are no data fields.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (not including the page header). This reference is returned as text.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (including the page headers). This reference is returned as text.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Returns the reference to the row header area as text. The row header area includes each row field header along with all the items in each row field. Returns #N/A if there are no row headers.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Returns the reference to the column header area as text. The column header area includes each column field header along with all the items in each column field. Returns #N/A if there are no column headers.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Returns the reference to the data header area as text. The data header area includes the data field header along with all the headers in the data row/col. Returns #N/A if there is no data field.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Returns a reference to all the page headers as text.</td>
</tr>
<tr class="even">
<td>17</td>
<td>Returns the reference to the PivotTable report data area as text.</td>
</tr>
<tr class="odd">
<td>18</td>
<td>Returns TRUE if the PivotTable report is set to show row grand totals.</td>
</tr>
<tr class="even">
<td>19</td>
<td>Returns TRUE if the PivotTable report is set to show column grand totals.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Returns TRUE if the user is saving data with the PivotTable report.</td>
</tr>
<tr class="even">
<td>21</td>
<td>Returns TRUE if the PivotTable report is set up to Autoformat on pivoting.</td>
</tr>
<tr class="odd">
<td>22</td>
<td><p>Returns the data source of the PivotTable report. The kind of information returned depends on the data source:</p>
<p>If the data source is a Microsoft Excel list or database, the cell reference is returned as text.</p>
<p>If the data source is an external data source, then an array is returned. Each row consists of a SQL connection string with the remaining elements as the query string broken down into 200 character segments.</p>
<p>If the data source is Multiple Consolidation ranges, then a two dimensional array is returned, each row of which consists of a reference and associated page field items.</p>
<p>If the data source is another PivotTable report, then one of the above three kinds of information is returned.</p></td>
</tr>
</tbody>
</table>

Pivot\_table\_name    is the name of a PivotTable report containing the
field that you want information about. If omitted, uses the PivotTable
report containing the active cell.

**Remarks**

Returns \#VALUE\! error value when pivot\_table\_name is not a valid
PivotTable name on the active sheet and the active cell is not within a
PivotTable report.

**Related Functions**

GET.PIVOT.FIELD   Returns information about an item in a PivotTable
report.

GET.PIVOT.ITEM   Returns information about a PivotTable report.


GET.PIVOT.TABLE

Returns information about a PivotTable report.

**Syntax**

**GET.PIVOT.TABLE**(**type\_num**,pivot\_table\_name)

Type\_num is a value from 1 to 22 that represents a type of information
you want about a PivotTable report.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Information</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Returns the name of the person who last updated the PivotTable report, as a text constant.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>Returns the date the PivotTable report was last updated, as a serial number.</td>
</tr>
<tr class="even">
<td>3</td>
<td>Returns a horizontal array of text constants representing all the fields in the PivotTable report.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>Returns an integer representing the number of fields in the PivotTable report.</td>
</tr>
<tr class="even">
<td>5</td>
<td>Returns a horizontal array of text constants representing all the visible fields in the PivotTable report (rows, columns, pages or data)</td>
</tr>
<tr class="odd">
<td>6</td>
<td>Returns a horizontal array of text constants representing all the hidden fields in the PivotTable report. Return #N/A if no hidden fields.</td>
</tr>
<tr class="even">
<td>7</td>
<td>Returns a horizontal array of text constants representing the names of all the fields currently showing in the PivotTable report as row fields. Returns #N/A if there are no row fields.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as column fields. Returns #N/A if no column fields exist.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as page fields. Return #N/A if no page fields exist.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as data fields. Returns #N/A if there are no data fields.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (not including the page header). This reference is returned as text.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (including the page headers). This reference is returned as text.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Returns the reference to the row header area as text. The row header area includes each row field header along with all the items in each row field. Returns #N/A if there are no row headers.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Returns the reference to the column header area as text. The column header area includes each column field header along with all the items in each column field. Returns #N/A if there are no column headers.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Returns the reference to the data header area as text. The data header area includes the data field header along with all the headers in the data row/col. Returns #N/A if there is no data field.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Returns a reference to all the page headers as text.</td>
</tr>
<tr class="even">
<td>17</td>
<td>Returns the reference to the PivotTable report data area as text.</td>
</tr>
<tr class="odd">
<td>18</td>
<td>Returns TRUE if the PivotTable report is set to show row grand totals.</td>
</tr>
<tr class="even">
<td>19</td>
<td>Returns TRUE if the PivotTable report is set to show column grand totals.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Returns TRUE if the user is saving data with the PivotTable report.</td>
</tr>
<tr class="even">
<td>21</td>
<td>Returns TRUE if the PivotTable report is set up to Autoformat on pivoting.</td>
</tr>
<tr class="odd">
<td>22</td>
<td><p>Returns the data source of the PivotTable report. The kind of information returned depends on the data source:</p>
<p>If the data source is a Microsoft Excel list or database, the cell reference is returned as text.</p>
<p>If the data source is an external data source, then an array is returned. Each row consists of a SQL connection string with the remaining elements as the query string broken down into 200 character segments.</p>
<p>If the data source is Multiple Consolidation ranges, then a two dimensional array is returned, each row of which consists of a reference and associated page field items.</p>
<p>If the data source is another PivotTable report, then one of the above three kinds of information is returned.</p></td>
</tr>
</tbody>
</table>

Pivot\_table\_name    is the name of a PivotTable report containing the
field that you want information about. If omitted, uses the PivotTable
report containing the active cell.

**Remarks**

Returns \#VALUE\! error value when pivot\_table\_name is not a valid
PivotTable name on the active sheet and the active cell is not within a
PivotTable report.

**Related Functions**

GET.PIVOT.FIELD   Returns information about an item in a PivotTable
report.

GET.PIVOT.ITEM   Returns information about a PivotTable report.


GET.PIVOT.TABLE

Returns information about a PivotTable report.

**Syntax**

**GET.PIVOT.TABLE**(**type\_num**,pivot\_table\_name)

Type\_num is a value from 1 to 22 that represents a type of information
you want about a PivotTable report.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Information</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Returns the name of the person who last updated the PivotTable report, as a text constant.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>Returns the date the PivotTable report was last updated, as a serial number.</td>
</tr>
<tr class="even">
<td>3</td>
<td>Returns a horizontal array of text constants representing all the fields in the PivotTable report.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>Returns an integer representing the number of fields in the PivotTable report.</td>
</tr>
<tr class="even">
<td>5</td>
<td>Returns a horizontal array of text constants representing all the visible fields in the PivotTable report (rows, columns, pages or data)</td>
</tr>
<tr class="odd">
<td>6</td>
<td>Returns a horizontal array of text constants representing all the hidden fields in the PivotTable report. Return #N/A if no hidden fields.</td>
</tr>
<tr class="even">
<td>7</td>
<td>Returns a horizontal array of text constants representing the names of all the fields currently showing in the PivotTable report as row fields. Returns #N/A if there are no row fields.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as column fields. Returns #N/A if no column fields exist.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as page fields. Return #N/A if no page fields exist.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as data fields. Returns #N/A if there are no data fields.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (not including the page header). This reference is returned as text.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (including the page headers). This reference is returned as text.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Returns the reference to the row header area as text. The row header area includes each row field header along with all the items in each row field. Returns #N/A if there are no row headers.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Returns the reference to the column header area as text. The column header area includes each column field header along with all the items in each column field. Returns #N/A if there are no column headers.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Returns the reference to the data header area as text. The data header area includes the data field header along with all the headers in the data row/col. Returns #N/A if there is no data field.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Returns a reference to all the page headers as text.</td>
</tr>
<tr class="even">
<td>17</td>
<td>Returns the reference to the PivotTable report data area as text.</td>
</tr>
<tr class="odd">
<td>18</td>
<td>Returns TRUE if the PivotTable report is set to show row grand totals.</td>
</tr>
<tr class="even">
<td>19</td>
<td>Returns TRUE if the PivotTable report is set to show column grand totals.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Returns TRUE if the user is saving data with the PivotTable report.</td>
</tr>
<tr class="even">
<td>21</td>
<td>Returns TRUE if the PivotTable report is set up to Autoformat on pivoting.</td>
</tr>
<tr class="odd">
<td>22</td>
<td><p>Returns the data source of the PivotTable report. The kind of information returned depends on the data source:</p>
<p>If the data source is a Microsoft Excel list or database, the cell reference is returned as text.</p>
<p>If the data source is an external data source, then an array is returned. Each row consists of a SQL connection string with the remaining elements as the query string broken down into 200 character segments.</p>
<p>If the data source is Multiple Consolidation ranges, then a two dimensional array is returned, each row of which consists of a reference and associated page field items.</p>
<p>If the data source is another PivotTable report, then one of the above three kinds of information is returned.</p></td>
</tr>
</tbody>
</table>

Pivot\_table\_name    is the name of a PivotTable report containing the
field that you want information about. If omitted, uses the PivotTable
report containing the active cell.

**Remarks**

Returns \#VALUE\! error value when pivot\_table\_name is not a valid
PivotTable name on the active sheet and the active cell is not within a
PivotTable report.

**Related Functions**

GET.PIVOT.FIELD   Returns information about an item in a PivotTable
report.

GET.PIVOT.ITEM   Returns information about a PivotTable report.


GET.PIVOT.TABLE

Returns information about a PivotTable report.

**Syntax**

**GET.PIVOT.TABLE**(**type\_num**,pivot\_table\_name)

Type\_num is a value from 1 to 22 that represents a type of information
you want about a PivotTable report.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Information</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Returns the name of the person who last updated the PivotTable report, as a text constant.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>Returns the date the PivotTable report was last updated, as a serial number.</td>
</tr>
<tr class="even">
<td>3</td>
<td>Returns a horizontal array of text constants representing all the fields in the PivotTable report.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>Returns an integer representing the number of fields in the PivotTable report.</td>
</tr>
<tr class="even">
<td>5</td>
<td>Returns a horizontal array of text constants representing all the visible fields in the PivotTable report (rows, columns, pages or data)</td>
</tr>
<tr class="odd">
<td>6</td>
<td>Returns a horizontal array of text constants representing all the hidden fields in the PivotTable report. Return #N/A if no hidden fields.</td>
</tr>
<tr class="even">
<td>7</td>
<td>Returns a horizontal array of text constants representing the names of all the fields currently showing in the PivotTable report as row fields. Returns #N/A if there are no row fields.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as column fields. Returns #N/A if no column fields exist.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as page fields. Return #N/A if no page fields exist.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as data fields. Returns #N/A if there are no data fields.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (not including the page header). This reference is returned as text.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (including the page headers). This reference is returned as text.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Returns the reference to the row header area as text. The row header area includes each row field header along with all the items in each row field. Returns #N/A if there are no row headers.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Returns the reference to the column header area as text. The column header area includes each column field header along with all the items in each column field. Returns #N/A if there are no column headers.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Returns the reference to the data header area as text. The data header area includes the data field header along with all the headers in the data row/col. Returns #N/A if there is no data field.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Returns a reference to all the page headers as text.</td>
</tr>
<tr class="even">
<td>17</td>
<td>Returns the reference to the PivotTable report data area as text.</td>
</tr>
<tr class="odd">
<td>18</td>
<td>Returns TRUE if the PivotTable report is set to show row grand totals.</td>
</tr>
<tr class="even">
<td>19</td>
<td>Returns TRUE if the PivotTable report is set to show column grand totals.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Returns TRUE if the user is saving data with the PivotTable report.</td>
</tr>
<tr class="even">
<td>21</td>
<td>Returns TRUE if the PivotTable report is set up to Autoformat on pivoting.</td>
</tr>
<tr class="odd">
<td>22</td>
<td><p>Returns the data source of the PivotTable report. The kind of information returned depends on the data source:</p>
<p>If the data source is a Microsoft Excel list or database, the cell reference is returned as text.</p>
<p>If the data source is an external data source, then an array is returned. Each row consists of a SQL connection string with the remaining elements as the query string broken down into 200 character segments.</p>
<p>If the data source is Multiple Consolidation ranges, then a two dimensional array is returned, each row of which consists of a reference and associated page field items.</p>
<p>If the data source is another PivotTable report, then one of the above three kinds of information is returned.</p></td>
</tr>
</tbody>
</table>

Pivot\_table\_name    is the name of a PivotTable report containing the
field that you want information about. If omitted, uses the PivotTable
report containing the active cell.

**Remarks**

Returns \#VALUE\! error value when pivot\_table\_name is not a valid
PivotTable name on the active sheet and the active cell is not within a
PivotTable report.

**Related Functions**

GET.PIVOT.FIELD   Returns information about an item in a PivotTable
report.

GET.PIVOT.ITEM   Returns information about a PivotTable report.


GET.PIVOT.TABLE

Returns information about a PivotTable report.

**Syntax**

**GET.PIVOT.TABLE**(**type\_num**,pivot\_table\_name)

Type\_num is a value from 1 to 22 that represents a type of information
you want about a PivotTable report.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Information</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Returns the name of the person who last updated the PivotTable report, as a text constant.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>Returns the date the PivotTable report was last updated, as a serial number.</td>
</tr>
<tr class="even">
<td>3</td>
<td>Returns a horizontal array of text constants representing all the fields in the PivotTable report.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>Returns an integer representing the number of fields in the PivotTable report.</td>
</tr>
<tr class="even">
<td>5</td>
<td>Returns a horizontal array of text constants representing all the visible fields in the PivotTable report (rows, columns, pages or data)</td>
</tr>
<tr class="odd">
<td>6</td>
<td>Returns a horizontal array of text constants representing all the hidden fields in the PivotTable report. Return #N/A if no hidden fields.</td>
</tr>
<tr class="even">
<td>7</td>
<td>Returns a horizontal array of text constants representing the names of all the fields currently showing in the PivotTable report as row fields. Returns #N/A if there are no row fields.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as column fields. Returns #N/A if no column fields exist.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as page fields. Return #N/A if no page fields exist.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as data fields. Returns #N/A if there are no data fields.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (not including the page header). This reference is returned as text.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (including the page headers). This reference is returned as text.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Returns the reference to the row header area as text. The row header area includes each row field header along with all the items in each row field. Returns #N/A if there are no row headers.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Returns the reference to the column header area as text. The column header area includes each column field header along with all the items in each column field. Returns #N/A if there are no column headers.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Returns the reference to the data header area as text. The data header area includes the data field header along with all the headers in the data row/col. Returns #N/A if there is no data field.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Returns a reference to all the page headers as text.</td>
</tr>
<tr class="even">
<td>17</td>
<td>Returns the reference to the PivotTable report data area as text.</td>
</tr>
<tr class="odd">
<td>18</td>
<td>Returns TRUE if the PivotTable report is set to show row grand totals.</td>
</tr>
<tr class="even">
<td>19</td>
<td>Returns TRUE if the PivotTable report is set to show column grand totals.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Returns TRUE if the user is saving data with the PivotTable report.</td>
</tr>
<tr class="even">
<td>21</td>
<td>Returns TRUE if the PivotTable report is set up to Autoformat on pivoting.</td>
</tr>
<tr class="odd">
<td>22</td>
<td><p>Returns the data source of the PivotTable report. The kind of information returned depends on the data source:</p>
<p>If the data source is a Microsoft Excel list or database, the cell reference is returned as text.</p>
<p>If the data source is an external data source, then an array is returned. Each row consists of a SQL connection string with the remaining elements as the query string broken down into 200 character segments.</p>
<p>If the data source is Multiple Consolidation ranges, then a two dimensional array is returned, each row of which consists of a reference and associated page field items.</p>
<p>If the data source is another PivotTable report, then one of the above three kinds of information is returned.</p></td>
</tr>
</tbody>
</table>

Pivot\_table\_name    is the name of a PivotTable report containing the
field that you want information about. If omitted, uses the PivotTable
report containing the active cell.

**Remarks**

Returns \#VALUE\! error value when pivot\_table\_name is not a valid
PivotTable name on the active sheet and the active cell is not within a
PivotTable report.

**Related Functions**
**Related Functions**

GET.PIVOT.FIELD   Returns information about an item in a PivotTable
report.

GET.PIVOT.ITEM   Returns information about a PivotTable report.


GET.PIVOT.TABLE

Returns information about a PivotTable report.

**Syntax**

**GET.PIVOT.TABLE**(**type\_num**,pivot\_table\_name)

Type\_num is a value from 1 to 22 that represents a type of information
you want about a PivotTable report.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Information</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Returns the name of the person who last updated the PivotTable report, as a text constant.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>Returns the date the PivotTable report was last updated, as a serial number.</td>
</tr>
<tr class="even">
<td>3</td>
<td>Returns a horizontal array of text constants representing all the fields in the PivotTable report.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>Returns an integer representing the number of fields in the PivotTable report.</td>
</tr>
<tr class="even">
<td>5</td>
<td>Returns a horizontal array of text constants representing all the visible fields in the PivotTable report (rows, columns, pages or data)</td>
</tr>
<tr class="odd">
<td>6</td>
<td>Returns a horizontal array of text constants representing all the hidden fields in the PivotTable report. Return #N/A if no hidden fields.</td>
</tr>
<tr class="even">
<td>7</td>
<td>Returns a horizontal array of text constants representing the names of all the fields currently showing in the PivotTable report as row fields. Returns #N/A if there are no row fields.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as column fields. Returns #N/A if no column fields exist.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as page fields. Return #N/A if no page fields exist.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as data fields. Returns #N/A if there are no data fields.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (not including the page header). This reference is returned as text.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (including the page headers). This reference is returned as text.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Returns the reference to the row header area as text. The row header area includes each row field header along with all the items in each row field. Returns #N/A if there are no row headers.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Returns the reference to the column header area as text. The column header area includes each column field header along with all the items in each column field. Returns #N/A if there are no column headers.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Returns the reference to the data header area as text. The data header area includes the data field header along with all the headers in the data row/col. Returns #N/A if there is no data field.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Returns a reference to all the page headers as text.</td>
</tr>
<tr class="even">
<td>17</td>
<td>Returns the reference to the PivotTable report data area as text.</td>
</tr>
<tr class="odd">
<td>18</td>
<td>Returns TRUE if the PivotTable report is set to show row grand totals.</td>
</tr>
<tr class="even">
<td>19</td>
<td>Returns TRUE if the PivotTable report is set to show column grand totals.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Returns TRUE if the user is saving data with the PivotTable report.</td>
</tr>
<tr class="even">
<td>21</td>
<td>Returns TRUE if the PivotTable report is set up to Autoformat on pivoting.</td>
</tr>
<tr class="odd">
<td>22</td>
<td><p>Returns the data source of the PivotTable report. The kind of information returned depends on the data source:</p>
<p>If the data source is a Microsoft Excel list or database, the cell reference is returned as text.</p>
<p>If the data source is an external data source, then an array is returned. Each row consists of a SQL connection string with the remaining elements as the query string broken down into 200 character segments.</p>
<p>If the data source is Multiple Consolidation ranges, then a two dimensional array is returned, each row of which consists of a reference and associated page field items.</p>
<p>If the data source is another PivotTable report, then one of the above three kinds of information is returned.</p></td>
</tr>
</tbody>
</table>

Pivot\_table\_name    is the name of a PivotTable report containing the
field that you want information about. If omitted, uses the PivotTable
report containing the active cell.

**Remarks**

Returns \#VALUE\! error value when pivot\_table\_name is not a valid
PivotTable name on the active sheet and the active cell is not within a
PivotTable report.

**Related Functions**

GET.PIVOT.FIELD   Returns information about an item in a PivotTable
report.

GET.PIVOT.ITEM   Returns information about a PivotTable report.


GET.PIVOT.TABLE

Returns information about a PivotTable report.

**Syntax**

**GET.PIVOT.TABLE**(**type\_num**,pivot\_table\_name)

Type\_num is a value from 1 to 22 that represents a type of information
you want about a PivotTable report.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Information</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Returns the name of the person who last updated the PivotTable report, as a text constant.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>Returns the date the PivotTable report was last updated, as a serial number.</td>
</tr>
<tr class="even">
<td>3</td>
<td>Returns a horizontal array of text constants representing all the fields in the PivotTable report.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>Returns an integer representing the number of fields in the PivotTable report.</td>
</tr>
<tr class="even">
<td>5</td>
<td>Returns a horizontal array of text constants representing all the visible fields in the PivotTable report (rows, columns, pages or data)</td>
</tr>
<tr class="odd">
<td>6</td>
<td>Returns a horizontal array of text constants representing all the hidden fields in the PivotTable report. Return #N/A if no hidden fields.</td>
</tr>
<tr class="even">
<td>7</td>
<td>Returns a horizontal array of text constants representing the names of all the fields currently showing in the PivotTable report as row fields. Returns #N/A if there are no row fields.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as column fields. Returns #N/A if no column fields exist.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as page fields. Return #N/A if no page fields exist.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as data fields. Returns #N/A if there are no data fields.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (not including the page header). This reference is returned as text.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (including the page headers). This reference is returned as text.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Returns the reference to the row header area as text. The row header area includes each row field header along with all the items in each row field. Returns #N/A if there are no row headers.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Returns the reference to the column header area as text. The column header area includes each column field header along with all the items in each column field. Returns #N/A if there are no column headers.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Returns the reference to the data header area as text. The data header area includes the data field header along with all the headers in the data row/col. Returns #N/A if there is no data field.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Returns a reference to all the page headers as text.</td>
</tr>
<tr class="even">
<td>17</td>
<td>Returns the reference to the PivotTable report data area as text.</td>
</tr>
<tr class="odd">
<td>18</td>
<td>Returns TRUE if the PivotTable report is set to show row grand totals.</td>
</tr>
<tr class="even">
<td>19</td>
<td>Returns TRUE if the PivotTable report is set to show column grand totals.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Returns TRUE if the user is saving data with the PivotTable report.</td>
</tr>
<tr class="even">
<td>21</td>
<td>Returns TRUE if the PivotTable report is set up to Autoformat on pivoting.</td>
</tr>
<tr class="odd">
<td>22</td>
<td><p>Returns the data source of the PivotTable report. The kind of information returned depends on the data source:</p>
<p>If the data source is a Microsoft Excel list or database, the cell reference is returned as text.</p>
<p>If the data source is an external data source, then an array is returned. Each row consists of a SQL connection string with the remaining elements as the query string broken down into 200 character segments.</p>
<p>If the data source is Multiple Consolidation ranges, then a two dimensional array is returned, each row of which consists of a reference and associated page field items.</p>
<p>If the data source is another PivotTable report, then one of the above three kinds of information is returned.</p></td>
</tr>
</tbody>
</table>

Pivot\_table\_name    is the name of a PivotTable report containing the
field that you want information about. If omitted, uses the PivotTable
report containing the active cell.

**Remarks**

Returns \#VALUE\! error value when pivot\_table\_name is not a valid
PivotTable name on the active sheet and the active cell is not within a
PivotTable report.

**Related Functions**

GET.PIVOT.FIELD   Returns information about an item in a PivotTable
report.

GET.PIVOT.ITEM   Returns information about a PivotTable report.


GET.PIVOT.TABLE

Returns information about a PivotTable report.

**Syntax**

**GET.PIVOT.TABLE**(**type\_num**,pivot\_table\_name)

Type\_num is a value from 1 to 22 that represents a type of information
you want about a PivotTable report.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Information</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Returns the name of the person who last updated the PivotTable report, as a text constant.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>Returns the date the PivotTable report was last updated, as a serial number.</td>
</tr>
<tr class="even">
<td>3</td>
<td>Returns a horizontal array of text constants representing all the fields in the PivotTable report.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>Returns an integer representing the number of fields in the PivotTable report.</td>
</tr>
<tr class="even">
<td>5</td>
<td>Returns a horizontal array of text constants representing all the visible fields in the PivotTable report (rows, columns, pages or data)</td>
</tr>
<tr class="odd">
<td>6</td>
<td>Returns a horizontal array of text constants representing all the hidden fields in the PivotTable report. Return #N/A if no hidden fields.</td>
</tr>
<tr class="even">
<td>7</td>
<td>Returns a horizontal array of text constants representing the names of all the fields currently showing in the PivotTable report as row fields. Returns #N/A if there are no row fields.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as column fields. Returns #N/A if no column fields exist.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as page fields. Return #N/A if no page fields exist.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as data fields. Returns #N/A if there are no data fields.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (not including the page header). This reference is returned as text.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (including the page headers). This reference is returned as text.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Returns the reference to the row header area as text. The row header area includes each row field header along with all the items in each row field. Returns #N/A if there are no row headers.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Returns the reference to the column header area as text. The column header area includes each column field header along with all the items in each column field. Returns #N/A if there are no column headers.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Returns the reference to the data header area as text. The data header area includes the data field header along with all the headers in the data row/col. Returns #N/A if there is no data field.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Returns a reference to all the page headers as text.</td>
</tr>
<tr class="even">
<td>17</td>
<td>Returns the reference to the PivotTable report data area as text.</td>
</tr>
<tr class="odd">
<td>18</td>
<td>Returns TRUE if the PivotTable report is set to show row grand totals.</td>
</tr>
<tr class="even">
<td>19</td>
<td>Returns TRUE if the PivotTable report is set to show column grand totals.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Returns TRUE if the user is saving data with the PivotTable report.</td>
</tr>
<tr class="even">
<td>21</td>
<td>Returns TRUE if the PivotTable report is set up to Autoformat on pivoting.</td>
</tr>
<tr class="odd">
<td>22</td>
<td><p>Returns the data source of the PivotTable report. The kind of information returned depends on the data source:</p>
<p>If the data sourcGET.PIVOT.TABLE

Returns information about a PivotTable report.

**Syntax**

**GET.PIVOT.TABLE**(**type\_num**,pivot\_table\_name)

Type\_num is a value from 1 to 22 that represents a type of information
you want about a PivotTable report.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Information</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Returns the name of the person who last updated the PivotTable report, as a text constant.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>Returns the date the PivotTable report was last updated, as a serial number.</td>
</tr>
<tr class="even">
<td>3</td>
<td>Returns a horizontal array of text constants representing all the fields in the PivotTable report.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>Returns an integer representing the number of fields in the PivotTable report.</td>
</tr>
<tr class="even">
<td>5</td>
<td>Returns a horizontal array of text constants representing all the visible fields in the PivotTable report (rows, columns, pages or data)</td>
</tr>
<tr class="odd">
<td>6</td>
<td>Returns a horizontal array of text constants representing all the hidden fields in the PivotTable report. Return #N/A if no hidden fields.</td>
</tr>
<tr class="even">
<td>7</td>
<td>Returns a horizontal array of text constants representing the names of all the fields currently showing in the PivotTable report as row fields. Returns #N/A if there are no row fields.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as column fields. Returns #N/A if no column fields exist.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as page fields. Return #N/A if no page fields exist.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as data fields. Returns #N/A if there are no data fields.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (not including the page header). This reference is returned as text.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (including the page headers). This reference is returned as text.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Returns the reference to the row header area as text. The row header area includes each row field header along with all the items in each row field. Returns #N/A if there are no row headers.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Returns the reference to the column header area as text. The column header area includes each column field header along with all the items in each column field. Returns #N/A if there are no column headers.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Returns the reference to the data header area as text. The data header area includes the data field header along with all the headers in the data row/col. Returns #N/A if there is no data field.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Returns a reference to all the page headers as text.</td>
</tr>
<tr class="even">
<td>17</td>
<td>Returns the reference to the PivotTable report data area as text.</td>
</tr>
<tr class="odd">
<td>18</td>
<td>Returns TRUE if the PivotTable report is set to show row grand totals.</td>
</tr>
<tr class="even">
<td>19</td>
<td>Returns TRUE if the PivotTable report is set to show column grand totals.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Returns TRUE if the user is saving data with the PivotTable report.</td>
</tr>
<tr class="even">
<td>21</td>
<td>Returns TRUE if the PivotTable report is set up to Autoformat on pivoting.</td>
</tr>
<tr class="odd">
<td>22</td>
<td><p>Returns the data source of the PivotTable report. The kind of information returned depends on the data source:</p>
<p>If the data source is a Microsoft Excel list or database, the cell reference is returned as text.</p>
<p>If the data source is an external data source, then an array is returned. Each row consists of a SQL connection string with the remaining elements as the query string broken down into 200 character segments.</p>
<p>If the data source is Multiple Consolidation ranges, then a two dimensional array is returned, each row of which consists of a reference and associated page field items.</p>
<p>If the data source is another PivotTable report, then one of the above three kinds of information is returned.</p></td>
</tr>
</tbody>
</table>

Pivot\_table\_name    is the name of a PivotTable report containing the
field that you want information about. If omitted, uses the PivotTable
report containing the active cell.

**Remarks**

Returns \#VALUE\! error value when pivot\_table\_name is not a valid
PivotTable name on the active sheet and the active cell is not within a
PivotTable report.

**Related Functions**

GET.PIVOT.FIELD   Returns information about an item in a PivotTable
report.

GET.PIVOT.ITEM   Returns information about a PivotTable report.


GET.PIVOT.TABLE

Returns information about a PivotTable report.

**Syntax**

**GET.PIVOT.TABLE**(**type\_num**,pivot\_table\_name)

Type\_num is a value from 1 to 22 that represents a type of information
you want about a PivotTable report.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Information</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Returns the name of the person who last updated the PivotTable report, as a text constant.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>Returns the date the PivotTable report was last updated, as a serial number.</td>
</tr>
<tr class="even">
<td>3</td>
<td>Returns a horizontal array of text constants representing all the fields in the PivotTable report.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>Returns an integer representing the number of fields in the PivotTable report.</td>
</tr>
<tr class="even">
<td>5</td>
<td>Returns a horizontal array of text constants representing all the visible fields in the PivotTable report (rows, columns, pages or data)</td>
</tr>
<tr class="odd">
<td>6</td>
<td>Returns a horizontal array of text constants representing all the hidden fields in the PivotTable report. Return #N/A if no hidden fields.</td>
</tr>
<tr class="even">
<td>7</td>
<td>Returns a horizontal array of text constants representing the names of all the fields currently showing in the PivotTable report as row fields. Returns #N/A if there are no row fields.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as column fields. Returns #N/A if no column fields exist.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as page fields. Return #N/A if no page fields exist.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as data fields. Returns #N/A if there are no data fields.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (not including the page header). This reference is returned as text.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (including the page headers). This reference is returned as text.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Returns the reference to the row header area as text. The row header area includes each row field header along with all the items in each row field. Returns #N/A if there are no row headers.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Returns the reference to the column header area as text. The column header area includes each column field header along with all the items in each column field. Returns #N/A if there are no column headers.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Returns the reference to the data header area as text. The data header area includes the data field header along with all the headers in the data row/col. Returns #N/A if there is no data field.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Returns a reference to all the page headers as text.</td>
</tr>
<tr class="even">
<td>17</td>
<td>Returns the reference to the PivotTable report data area as text.</td>
</tr>
<tr class="odd">
<td>18</td>
<td>Returns TRUE if the PivotTable report is set to show row grand totals.</td>
</tr>
<tr class="even">
<td>19</td>
<td>Returns TRUE if the PivotTable report is set to show column grand totals.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Returns TRUE if the user is saving data with the PivotTable report.</td>
</tr>
<tr class="even">
<td>21</td>
<td>Returns TRUE if the PivotTable report is set up to Autoformat on pivoting.</td>
</tr>
<tr class="odd">
<td>22</td>
<td><p>Returns the data source of the PivotTable report. The kind of information returned depends on the data source:</p>
<p>If the data sourcGET.PIVOT.TABLE

Returns information about a PivotTable report.

**Syntax**

**GET.PIVOT.TABLE**(**type\_num**,pivot\_table\_name)

Type\_num is a value from 1 to 22 that represents a type of information
you want about a PivotTable report.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Information</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Returns the name of the person who last updated the PivotTable report, as a text constant.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>Returns the date the PivotTable report was last updated, as a serial number.</td>
</tr>
<tr class="even">
<td>3</td>
<td>Returns a horizontal array of text constants representing all the fields in the PivotTable report.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>Returns an integer representing the number of fields in the PivotTable report.</td>
</tr>
<tr class="even">
<td>5</td>
<td>Returns a horizontal array of text constants representing all the visible fields in the PivotTable report (rows, columns, pages or data)</td>
</tr>
<tr class="odd">
<td>6</td>
<td>Returns a horizontal array of text constants representing all the hidden fields in the PivotTable report. Return #N/A if no hidden fields.</td>
</tr>
<tr class="even">
<td>7</td>
<td>Returns a horizontal array of text constants representing the names of all the fields currently showing in the PivotTable report as row fields. Returns #N/A if there are no row fields.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as column fields. Returns #N/A if no column fields exist.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as page fields. Return #N/A if no page fields exist.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as data fields. Returns #N/A if there are no data fields.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (not including the page header). This reference is returned as text.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (including the page headers). This reference is returned as text.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Returns the reference to the row header area as text. The row header area includes each row field header along with all the items in each row field. Returns #N/A if there are no row headers.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Returns the reference to the column header area as text. The column header area includes each column field header along with all the items in each column field. Returns #N/A if there are no column headers.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Returns the reference to the data header area as text. The data header area includes the data field header along with all the headers in the data row/col. Returns #N/A if there is no data field.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Returns a reference to all the page headers as text.</td>
</tr>
<tr class="even">
<td>17</td>
<td>Returns the reference to the PivotTable report data area as text.</td>
</tr>
<tr class="odd">
<td>18</td>
<td>Returns TRUE if the PivotTable report is set to show row grand totals.</td>
</tr>
<tr class="even">
<td>19</td>
<td>Returns TRUE if the PivotTable report is set to show column grand totals.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Returns TRUE if the user is saving data with the PivotTable report.</td>
</tr>
<tr class="even">
<td>21</td>
<td>Returns TRUE if the PivotTable report is set up to Autoformat on pivoting.</td>
</tr>
<tr class="odd">
<td>22</td>
<td><p>Returns the data source of the PivotTable report. The kind of information returned depends on the data source:</p>
<p>If the data sourcGET.PIVOT.TABLE

Returns information about a PivotTable report.

**Syntax**

**GET.PIVOT.TABLE**(**type\_num**,pivot\_table\_name)

Type\_num is a value from 1 to 22 that represents a type of information
you want about a PivotTable report.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Information</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Returns the name of the person who last updated the PivotTable report, as a text constant.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>Returns the date the PivotTable report was last updated, as a serial number.</td>
</tr>
<tr class="even">
<td>3</td>
<td>Returns a horizontal array of text constants representing all the fields in the PivotTable report.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>Returns an integer representing the number of fields in the PivotTable report.</td>
</tr>
<tr class="even">
<td>5</td>
<td>Returns a horizontal array of text constants representing all the visible fields in the PivotTable report (rows, columns, pages or data)</td>
</tr>
<tr class="odd">
<td>6</td>
<td>Returns a horizontal array of text constants representing all the hidden fields in the PivotTable report. Return #N/A if no hidden fields.</td>
</tr>
<tr class="even">
<td>7</td>
<td>Returns a horizontal array of text constants representing the names of all the fields currently showing in the PivotTable report as row fields. Returns #N/A if there are no row fields.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as column fields. Returns #N/A if no column fields exist.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as page fields. Return #N/A if no page fields exist.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as data fields. Returns #N/A if there are no data fields.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (not including the page header). This reference is returned as text.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (including the page headers). This reference is returned as text.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Returns the reference to the row header area as text. The row header area includes each row field header along with all the items in each row field. Returns #N/A if there are no row headers.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Returns the reference to the column header area as text. The column header area includes each column field header along with all the items in each column field. Returns #N/A if there are no column headers.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Returns the reference to the data header area as text. The data header area includes the data field header along with all the headers in the data row/col. Returns #N/A if there is no data field.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Returns a reference to all the page headers as text.</td>
</tr>
<tr class="even">
<td>17</td>
<td>Returns the reference to the PivotTable report data area as text.</td>
</tr>
<tr class="odd">
<td>18</td>
<td>Returns TRUE if the PivotTable report is set to show row grand totals.</td>
</tr>
<tr class="even">
<td>19</td>
<td>Returns TRUE if the PivotTable report is set to show column grand totals.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Returns TRUE if the user is saving data with the PivotTable report.</td>
</tr>
<tr class="even">
<td>21</td>
<td>Returns TRUE if the PivotTable report is set up to Autoformat on pivoting.</td>
</tr>
<tr class="odd">
<td>22</td>
<td><p>Returns the data source of the PivotTable report. The kind of information returned depends on the data source:</p>
<p>If the data sourcGET.PIVOT.TABLE

Returns information about a PivotTable report.

**Syntax**

**GET.PIVOT.TABLE**(**type\_num**,pivot\_table\_name)

Type\_num is a value from 1 to 22 that represents a type of information
you want about a PivotTable report.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Information</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Returns the name of the person who last updated the PivotTable report, as a text constant.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>Returns the date the PivotTable report was last updated, as a serial number.</td>
</tr>
<tr class="even">
<td>3</td>
<td>Returns a horizontal array of text constants representing all the fields in the PivotTable report.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>Returns an integer representing the number of fields in the PivotTable report.</td>
</tr>
<tr class="even">
<td>5</td>
<td>Returns a horizontal array of text constants representing all the visible fields in the PivotTable report (rows, columns, pages or data)</td>
</tr>
<tr class="odd">
<td>6</td>
<td>Returns a horizontal array of text constants representing all the hidden fields in the PivotTable report. Return #N/A if no hidden fields.</td>
</tr>
<tr class="even">
<td>7</td>
<td>Returns a horizontal array of text constants representing the names of all the fields currently showing in the PivotTable report as row fields. Returns #N/A if there are no row fields.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as column fields. Returns #N/A if no column fields exist.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as page fields. Return #N/A if no page fields exist.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as data fields. Returns #N/A if there are no data fields.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (not including the page header). This reference is returned as text.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (including the page headers). This reference is returned as text.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Returns the reference to the row header area as text. The row header area includes each row field header along with all the items in each row field. Returns #N/A if there are no row headers.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Returns the reference to the column header area as text. The column header area includes each column field header along with all the items in each column field. Returns #N/A if there are no column headers.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Returns the reference to the data header area as text. The data header area includes the data field header along with all the headers in the data row/col. Returns #N/A if there is no data field.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Returns a reference to all the page headers as text.</td>
</tr>
<tr class="even">
<td>17</td>
<td>Returns the reference to the PivotTable report data area as text.</td>
</tr>
<tr class="odd">
<td>18</td>
<td>Returns TRUE if the PivotTable report is set to show row grand totals.</td>
</tr>
<tr class="even">
<td>19</td>
<td>Returns TRUE if the PivotTable report is set to show column grand totals.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Returns TRUE if the user is saving data with the PivotTable report.</td>
</tr>
<tr class="even">
<td>21</td>
<td>Returns TRUE if the PivotTable report is set up to Autoformat on pivoting.</td>
</tr>
<tr class="odd">
<td>22</td>
<td><p>Returns the data source of the PivotTable report. The kind of information returned depends on the data source:</p>
<p>If the data sourcGET.PIVOT.TABLE

Returns information about a PivotTable report.

**Syntax**

**GET.PIVOT.TABLE**(**type\_num**,pivot\_table\_name)

Type\_num is a value from 1 to 22 that represents a type of information
you want about a PivotTable report.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Information</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Returns the name of the person who last updated the PivotTable report, as a text constant.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>Returns the date the PivotTable report was last updated, as a serial number.</td>
</tr>
<tr class="even">
<td>3</td>
<td>Returns a horizontal array of text constants representing all the fields in the PivotTable report.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>Returns an integer representing the number of fields in the PivotTable report.</td>
</tr>
<tr class="even">
<td>5</td>
<td>Returns a horizontal array of text constants representing all the visible fields in the PivotTable report (rows, columns, pages or data)</td>
</tr>
<tr class="odd">
<td>6</td>
<td>Returns a horizontal array of text constants representing all the hidden fields in the PivotTable report. Return #N/A if no hidden fields.</td>
</tr>
<tr class="even">
<td>7</td>
<td>Returns a horizontal array of text constants representing the names of all the fields currently showing in the PivotTable report as row fields. Returns #N/A if there are no row fields.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as column fields. Returns #N/A if no column fields exist.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as page fields. Return #N/A if no page fields exist.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as data fields. Returns #N/A if there are no data fields.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (not including the page header). This reference is returned as text.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (including the page headers). This reference is returned as text.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Returns the reference to the row header area as text. The row header area includes each row field header along with all the items in each row field. Returns #N/A if there are no row headers.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Returns the reference to the column header area as text. The column header area includes each column field header along with all the items in each column field. Returns #N/A if there are no column headers.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Returns the reference to the data header area as text. The data header area includes the data field header along with all the headers in the data row/col. Returns #N/A if there is no data field.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Returns a reference to all the page headers as text.</td>
</tr>
<tr class="even">
<td>17</td>
<td>Returns the reference to the PivotTable report data area as text.</td>
</tr>
<tr class="odd">
<td>18</td>
<td>Returns TRUE if the PivotTable report is set to show row grand totals.</td>
</tr>
<tr class="even">
<td>19</td>
<td>Returns TRUE if the PivotTable report is set to show column grand totals.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Returns TRUE if the user is saving data with the PivotTable report.</td>
</tr>
<tr class="even">
<td>21</td>
<td>Returns TRUE if the PivotTable report is set up to Autoformat on pivoting.</td>
</tr>
<tr class="odd">
<td>22</td>
<td><p>Returns the data source of the PivotTable report. The kind of information returned depends on the data source:</p>
<p>If the data sourcGET.PIVOT.TABLE

Returns information about a PivotTable report.

**Syntax**

**GET.PIVOT.TABLE**(**type\_num**,pivot\_table\_name)

Type\_num is a value from 1 to 22 that represents a type of information
you want about a PivotTable report.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Information</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Returns the name of the person who last updated the PivotTable report, as a text constant.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>Returns the date the PivotTable report was last updated, as a serial number.</td>
</tr>
<tr class="even">
<td>3</td>
<td>Returns a horizontal array of text constants representing all the fields in the PivotTable report.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>Returns an integer representing the number of fields in the PivotTable report.</td>
</tr>
<tr class="even">
<td>5</td>
<td>Returns a horizontal array of text constants representing all the visible fields in the PivotTable report (rows, columns, pages or data)</td>
</tr>
<tr class="odd">
<td>6</td>
<td>Returns a horizontal array of text constants representing all the hidden fields in the PivotTable report. Return #N/A if no hidden fields.</td>
</tr>
<tr class="even">
<td>7</td>
<td>Returns a horizontal array of text constants representing the names of all the fields currently showing in the PivotTable report as row fields. Returns #N/A if there are no row fields.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as column fields. Returns #N/A if no column fields exist.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as page fields. Return #N/A if no page fields exist.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as data fields. Returns #N/A if there are no data fields.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (not including the page header). This reference is returned as text.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (including the page headers). This reference is returned as text.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Returns the reference to the row header area as text. The row header area includes each row field header along with all the items in each row field. Returns #N/A if there are no row headers.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Returns the reference to the column header area as text. The column header area includes each column field header along with all the items in each column field. Returns #N/A if there are no column headers.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Returns the reference to the data header area as text. The data header area includes the data field header along with all the headers in the data row/col. Returns #N/A if there is no data field.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Returns a reference to all the page headers as text.</td>
</tr>
<tr class="even">
<td>17</td>
<td>Returns the reference to the PivotTable report data area as text.</td>
</tr>
<tr class="odd">
<td>18</td>
<td>Returns TRUE if the PivotTable report is set to show row grand totals.</td>
</tr>
<tr class="even">
<td>19</td>
<td>Returns TRUE if the PivotTable report is set to show column grand totals.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Returns TRUE if the user is saving data with the PivotTable report.</td>
</tr>
<tr class="even">
<td>21</td>
<td>Returns TRUE if the PivotTable report is set up to Autoformat on pivoting.</td>
</tr>
<tr class="odd">
<td>22</td>
<td><p>Returns the data source of the PivotTable report. The kind of information returned depends on the data source:</p>
<p>If the data sourcGET.PIVOT.TABLE

Returns information about a PivotTable report.

**Syntax**

**GET.PIVOT.TABLE**(**type\_num**,pivot\_table\_name)

Type\_num is a value from 1 to 22 that represents a type of information
you want about a PivotTable report.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Information</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Returns the name of the person who last updated the PivotTable report, as a text constant.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>Returns the date the PivotTable report was last updated, as a serial number.</td>
</tr>
<tr class="even">
<td>3</td>
<td>Returns a horizontal array of text constants representing all the fields in the PivotTable report.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>Returns an integer representing the number of fields in the PivotTable report.</td>
</tr>
<tr class="even">
<td>5</td>
<td>Returns a horizontal array of text constants representing all the visible fields in the PivotTable report (rows, columns, pages or data)</td>
</tr>
<tr class="odd">
<td>6</td>
<td>Returns a horizontal array of text constants representing all the hidden fields in the PivotTable report. Return #N/A if no hidden fields.</td>
</tr>
<tr class="even">
<td>7</td>
<td>Returns a horizontal array of text constants representing the names of all the fields currently showing in the PivotTable report as row fields. Returns #N/A if there are no row fields.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as column fields. Returns #N/A if no column fields exist.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as page fields. Return #N/A if no page fields exist.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as data fields. Returns #N/A if there are no data fields.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (not including the page header). This reference is returned as text.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (including the page headers). This reference is returned as text.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Returns the reference to the row header area as text. The row header area includes each row field header along with all the items in each row field. Returns #N/A if there are no row headers.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Returns the reference to the column header area as text. The column header area includes each column field header along with all the items in each column field. Returns #N/A if there are no column headers.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Returns the reference to the data header area as text. The data header area includes the data field header along with all the headers in the data row/col. Returns #N/A if there is no data field.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Returns a reference to all the page headers as text.</td>
</tr>
<tr class="even">
<td>17</td>
<td>Returns the reference to the PivotTable report data area as text.</td>
</tr>
<tr class="odd">
<td>18</td>
<td>Returns TRUE if the PivotTable report is set to show row grand totals.</td>
</tr>
<tr class="even">
<td>19</td>
<td>Returns TRUE if the PivotTable report is set to show column grand totals.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Returns TRUE if the user is saving data with the PivotTable report.</td>
</tr>
<tr class="even">
<td>21</td>
<td>Returns TRUE if the PivotTable report is set up to Autoformat on pivoting.</td>
</tr>
<tr class="odd">
<td>22</td>
<td><p>Returns the data source of the PivotTable report. The kind of information returned depends on the data source:</p>
<p>If the data source is a Microsoft Excel list or database, the cell reference is returned as text.</p>
<p>If the data source is an external data source, then an array is returned. Each row consists of a SQL connection string with the remaining elements as the query string broken down into 200 character segments.</p>
<p>If the data source is Multiple Consolidation ranges, then a two dimensional array is returned, each row of which consists of a reference and associated page field items.</p>
<p>If the data source is another PivotTable report, then one of the above three kinds of information is returned.</p></td>
</tr>
</tbody>
</table>

Pivot\_table\_name    is the name of a PivotTable report containing the
field that you want information about. If omitted, uses the PivotTable
report containing the active cell.

**Remarks**

Returns \#VALUE\! error value when pivot\_table\_name is not a valid
PivotTable name on the active sheet and the active cell is not within a
PivotTable report.

**Related Functions**

[GET.PIVOT.FIELD](GET.PIVOT.FIELD.md)   Returns information about an item in a PivotTable
report.

[GET.PIVOT.ITEM](GET.PIVOT.ITEM.md)   Returns information about a PivotTable report.


