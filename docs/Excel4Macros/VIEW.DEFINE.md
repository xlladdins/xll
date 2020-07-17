VIEW.DEFINE

Equivalent to clicking the Add button in the Custom Views dialog box in
Microsoft Excel 97 or later, which appears when you click the Custom
Views command on the View menu. In Microsoft Excel 97, the Custom Views
command replaced the View Manager command that was available in
Microsoft Excel 95 and earlier versions. Creates or replaces a view.

If this function is not available in Microsoft Excel 95 or in earlier
versions, you must install the View Manager add-in.

**Syntax**

**VIEW.DEFINE**(**view\_name**, print\_settings\_log, row\_col\_log)

View\_name    is text enclosed in quotation marks and specifies the name
of the view you want to define for the sheet. If the workbook already
contains a view with view\_name, the new view replaces the existing one.

Print\_settings\_log    is a logical value that, if TRUE or omitted,
includes current print settings in the view or, if FALSE, does not
include current print settings in the view.

Row\_col\_log    is a logical value that, if TRUE or omitted, includes
current row and column settings in the view or, if FALSE, does not
include current row and column settings in the view.

**Related Functions**

[VIEW.DELETE](VIEW.DELETE.md)   Removes a view from the active workbook

[VIEW.SHOW](VIEW.SHOW.md)   Shows a view



Return to [README](README.md)

