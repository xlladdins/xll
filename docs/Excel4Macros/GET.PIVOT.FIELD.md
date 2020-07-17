GET.PIVOT.FIELD
===============

Returns information about a field in a PivotTable report.

**Syntax**

**GET.PIVOT.FIELD**(type\_num, pivot\_field\_name, pivot\_table\_name)

Type\_num    is a value from 1 to 17 that returns the following types of
information:

+---------------+-----------------------------------------------------+
| **Type\_num** | **Value**                                           |
+---------------+-----------------------------------------------------+
| 1             | Returns an array of all the items which make up     |
|               | pivot\_field\_name. The array is made up of text    |
|               | constants, dates or numbers depending on the field. |
+---------------+-----------------------------------------------------+
| 2             | Returns an array of all items which are set to show |
|               | with the pivot\_field\_name. The array is made up   |
|               | of text constants, dates or numbers depending on    |
|               | the field. The array is returned in the order that  |
|               | the items are displayed in the PivotTable report.   |
|               | If pivot\_field\_name is a page field, then the     |
|               | array contains only one element, the value          |
|               | corresponding to the active page (this could be all |
|               | if the All item is showing).                        |
+---------------+-----------------------------------------------------+
| 3             | Returns an array of all items which are hidden in   |
|               | the pivot\_field\_name. The array is made up of     |
|               | text constants, dates or numbers depending on the   |
|               | field. If pivot\_field\_name is a data field or the |
|               | data header name, this function returns the \#N/A!  |
|               | error value.                                        |
+---------------+-----------------------------------------------------+
| 4             | Returns an integer describing where the field is    |
|               | displayed in the active PivotTable report (either   |
|               | row or column):                                     |
|               |                                                     |
|               | 0 = Hidden                                          |
|               |                                                     |
|               | 1 = Row                                             |
|               |                                                     |
|               | 2 = Col                                             |
|               |                                                     |
|               | 3 = Page                                            |
|               |                                                     |
|               | 4 = Data                                            |
+---------------+-----------------------------------------------------+
| 5             | Returns an array of all items in pivot\_field\_name |
|               | that are group parents. The array is made up of     |
|               | text constants, dates or numbers depending on the   |
|               | field. The array is returned in the order which     |
|               | these items appear in the PivotTable report.        |
|               | Returns \#N/A if there are no group parents and if  |
|               | the pivot\_field\_name is a data field or the data  |
|               | field header.                                       |
+---------------+-----------------------------------------------------+
| 6             | Returns a number between 0 and 4095 which describes |
|               | the subtotals attached to the field. The number is  |
|               | the sum of the values associated with each subtotal |
|               | function. See PIVOT.FIELD.PROPERTIES for a list of  |
|               | all the values associated with subtotal             |
|               | calculations. If the field is showing as a data     |
|               | field or data field header, \#N/A! is returned.     |
+---------------+-----------------------------------------------------+
| 7             | Returns an integer describing the type of data      |
|               | contained in the field:                             |
|               |                                                     |
|               | 0 = Text                                            |
|               |                                                     |
|               | 1 = Number                                          |
|               |                                                     |
|               | 2 = Date                                            |
+---------------+-----------------------------------------------------+
| 8             | Returns an array five columns wide and one row high |
|               | describing the summary function\'s custom           |
|               | calculation shown with the specified field (Data    |
|               | field) in the PivotTable report. The array will     |
|               | look as follows: {function, calculation, base       |
|               | field, base item, number format}. If                |
|               | pivot\_field\_name is not showing in the active     |
|               | PivotTable report as a data field, \#N/A! is        |
|               | returned.                                           |
+---------------+-----------------------------------------------------+
| 9             | Returns a reference to all of pivot\_field\_name\'s |
|               | items currently showing in the active PivotTable    |
|               | report. If pivot\_field\_name is hidden, \#N/A! is  |
|               | returned. If pivot\_field\_name is a page field,    |
|               | the reference to the currently showing page item is |
|               | returned. If pivot\_field\_name is a data field, a  |
|               | reference to all the data for this field in the     |
|               | PivotTable report is returned. The references are   |
|               | returned as text.                                   |
+---------------+-----------------------------------------------------+
| 10            | Returns a reference to the header cell for          |
|               | pivot\_field\_name. If pivot\_field\_name is a data |
|               | field, a reference to all the headers in the data   |
|               | row or column is returned. If pivot\_field\_name is |
|               | hidden, \#N/A! is returned. The reference is        |
|               | returned as text.                                   |
+---------------+-----------------------------------------------------+
| 11            | Returns the number of grouped fields in the grouped |
|               | field set which includes pivot\_field\_name. If     |
|               | pivot\_field\_name is neither a parent field nor a  |
|               | child field, 1 is returned. If pivot\_field\_name   |
|               | is a data field or data header name, the function   |
|               | returns the \#N/A! error value.                     |
+---------------+-----------------------------------------------------+
| 12            | Returns the level of pivot\_field\_name in the      |
|               | grouped field set which includes                    |
|               | pivot\_field\_name. Returns 1 for the highest level |
|               | parent field, 2 for its child field, and so on. If  |
|               | pivot\_field\_name is neither a parent field nor a  |
|               | child field, 1 is returned. If pivot\_field\_name   |
|               | is a data field or data header name, the function   |
|               | returns the \#N/A! error value.                     |
+---------------+-----------------------------------------------------+
| 13            | Returns the name of the parent field for            |
|               | pivot\_field\_name as a text constant. If           |
|               | pivot\_field\_name is not a child field, \#N/A! is  |
|               | returned.                                           |
+---------------+-----------------------------------------------------+
| 14            | Returns the name of the child field for             |
|               | pivot\_field\_name as a text constant. If           |
|               | pivot\_field\_name is not a parent field, \#N/A! is |
|               | returned.                                           |
+---------------+-----------------------------------------------------+
| 15            | Returns a text constant representing the original   |
|               | name of the field in the data source.               |
+---------------+-----------------------------------------------------+
| 16            | Returns the position of the field among all the     |
|               | other fields in its orientation. For instance, a 1  |
|               | would be returned if the field was the first row    |
|               | field.                                              |
+---------------+-----------------------------------------------------+
| 17            | Returns an array of all items in pivot\_field\_name |
|               | that are group children. The array is made up of    |
|               | text constants, dates or numbers depending on the   |
|               | field. The array is returned in the order which     |
|               | these items appear in the PivotTable report.        |
|               | Returns \#N/A if there are no group children, and   |
|               | if the pivot\_field\_name is a data field or the    |
|               | data field header.                                  |
+---------------+-----------------------------------------------------+

Pivot\_field\_name    is the name of the field that you want information
about. If there is no field named pivot\_field\_name in the PivotTable
report, returns \#VALUE!.

Pivot\_table\_name    is the name of a PivotTable report containing the
field that you want information about. If omitted, the PivotTable report
containing the active cell is used. If the active cell is not in a
PivotTable report, the \#VALUE! error value is returned.

**Related Functions**

GET.PIVOT.ITEM   Returns information about an item in a PivotTable
report.

GET.PIVOT.TABLE   Returns information about a PivotTable report.

Return to [top](#E)

GET.PIVOT.ITEM
