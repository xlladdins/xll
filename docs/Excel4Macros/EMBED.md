EMBED
=====

Displayed in the formula bar when an embedded object is selected. EMBED
cannot be entered on a sheet or used in a macro.

**Syntax**

**EMBED**(**object\_class**, item)

Object\_class    is the name of the application and document type that
created the embedded object. For example, the object\_class arguments
used when Microsoft Excel sheets are embedded in other applications are
\"Excel.sheet.5\" and \"Excel.Chart.5\".

Item    is the area selected to copy, and determines the view on the
embedded document. When item is empty text (\"\"), EMBED creates a view
on the entire document.

**Remarks**

If you delete the EMBED formula, the embedded object remains on the
sheet as a graphic, and the link to the creating application is deleted.
Double-clicking the object no longer starts the creating application.

Return to [top](#E)

ENABLE.COMMAND
