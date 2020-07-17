GET.PIVOT.TABLE
===============

Returns information about a PivotTable report.

**Syntax**

**GET.PIVOT.TABLE**(**type\_num**,pivot\_table\_name)

Type\_num is a value from 1 to 22 that represents a type of information
you want about a PivotTable report.

+---------------+-----------------------------------------------------+
| **Type\_num** | **Information**                                     |
+---------------+-----------------------------------------------------+
| 1             | Returns the name of the person who last updated the |
|               | PivotTable report, as a text constant.              |
+---------------+-----------------------------------------------------+
| 2             | Returns the date the PivotTable report was last     |
|               | updated, as a serial number.                        |
+---------------+-----------------------------------------------------+
| 3             | Returns a horizontal array of text constants        |
|               | representing all the fields in the PivotTable       |
|               | report.                                             |
+---------------+-----------------------------------------------------+
| 4             | Returns an integer representing the number of       |
|               | fields in the PivotTable report.                    |
+---------------+-----------------------------------------------------+
| 5             | Returns a horizontal array of text constants        |
|               | representing all the visible fields in the          |
|               | PivotTable report (rows, columns, pages or data)    |
+---------------+-----------------------------------------------------+
| 6             | Returns a horizontal array of text constants        |
|               | representing all the hidden fields in the           |
|               | PivotTable report. Return \#N/A if no hidden        |
|               | fields.                                             |
+---------------+-----------------------------------------------------+
| 7             | Returns a horizontal array of text constants        |
|               | representing the names of all the fields currently  |
|               | showing in the PivotTable report as row fields.     |
|               | Returns \#N/A if there are no row fields.           |
+---------------+-----------------------------------------------------+
| 8             | Returns a horizontal array of text constants        |
|               | representing all the fields currently showing in    |
|               | the PivotTable report as column fields. Returns     |
|               | \#N/A if no column fields exist.                    |
+---------------+-----------------------------------------------------+
| 9             | Returns a horizontal array of text constants        |
|               | representing all the fields currently showing in    |
|               | the PivotTable report as page fields. Return \#N/A  |
|               | if no page fields exist.                            |
+---------------+-----------------------------------------------------+
| 10            | Returns a horizontal array of text constants        |
|               | representing all the fields currently showing in    |
|               | the PivotTable report as data fields. Returns \#N/A |
|               | if there are no data fields.                        |
+---------------+-----------------------------------------------------+
| 11            | Returns the smallest rectangular reference which    |
|               | bounds the PivotTable report and all headers (not   |
|               | including the page header). This reference is       |
|               | returned as text.                                   |
+---------------+-----------------------------------------------------+
| 12            | Returns the smallest rectangular reference which    |
|               | bounds the PivotTable report and all headers        |
|               | (including the page headers). This reference is     |
|               | returned as text.                                   |
+---------------+-----------------------------------------------------+
| 13            | Returns the reference to the row header area as     |
|               | text. The row header area includes each row field   |
|               | header along with all the items in each row field.  |
|               | Returns \#N/A if there are no row headers.          |
+---------------+-----------------------------------------------------+
| 14            | Returns the reference to the column header area as  |
|               | text. The column header area includes each column   |
|               | field header along with all the items in each       |
|               | column field. Returns \#N/A if there are no column  |
|               | headers.                                            |
+---------------+-----------------------------------------------------+
| 15            | Returns the reference to the data header area as    |
|               | text. The data header area includes the data field  |
|               | header along with all the headers in the data       |
|               | row/col. Returns \#N/A if there is no data field.   |
+---------------+-----------------------------------------------------+
| 16            | Returns a reference to all the page headers as      |
|               | text.                                               |
+---------------+-----------------------------------------------------+
| 17            | Returns the reference to the PivotTable report data |
|               | area as text.                                       |
+---------------+-----------------------------------------------------+
| 18            | Returns TRUE if the PivotTable report is set to     |
|               | show row grand totals.                              |
+---------------+-----------------------------------------------------+
| 19            | Returns TRUE if the PivotTable report is set to     |
|               | show column grand totals.                           |
+---------------+-----------------------------------------------------+
| 20            | Returns TRUE if the user is saving data with the    |
|               | PivotTable report.                                  |
+---------------+-----------------------------------------------------+
| 21            | Returns TRUE if the PivotTable report is set up to  |
|               | Autoformat on pivoting.                             |
+---------------+-----------------------------------------------------+
| 22            | Returns the data source of the PivotTable report.   |
|               | The kind of information returned depends on the     |
|               | data source:                                        |
|               |                                                     |
|               | If the data source is a Microsoft Excel list or     |
|               | database, the cell reference is returned as text.   |
|               |                                                     |
|               | If the data source is an external data source, then |
|               | an array is returned. Each row consists of a SQL    |
|               | connection string with the remaining elements as    |
|               | the query string broken down into 200 character     |
|               | segments.                                           |
|               |                                                     |
|               | If the data source is Multiple Consolidation        |
|               | ranges, then a two dimensional array is returned,   |
|               | each row of which consists of a reference and       |
|               | associated page field items.                        |
|               |                                                     |
|               | If the data source is another PivotTable report,    |
|               | then one of the above three kinds of information is |
|               | returned.                                           |
+---------------+-----------------------------------------------------+

Pivot\_table\_name    is the name of a PivotTable report containing the
field that you want information about. If omitted, uses the PivotTable
report containing the active cell.

**Remarks**

Returns \#VALUE! error value when pivot\_table\_name is not a valid
PivotTable name on the active sheet and the active cell is not within a
PivotTable report.

**Related Functions**

GET.PIVOT.FIELD   Returns information about an item in a PivotTable
report.

GET.PIVOT.ITEM   Returns information about a PivotTable report.

Return to [top](#E)

GET.TOOL
