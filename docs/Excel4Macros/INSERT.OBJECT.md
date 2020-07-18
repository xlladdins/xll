INSERT.OBJECT

Equivalent to choosing the Object command from the Insert menu, and then
selecting an object type and choosing the OK button. Creates an embedded
object whose source data is supplied by another application. Also starts
an application of the appropriate class for the specified object type.

**Syntax**

**INSERT.OBJECT**(**object\_class**, file\_name, link\_logical,
display\_icon\_logical, icon\_file, icon\_number, icon\_label)

**INSERT.OBJECT**?(object\_class, file\_name, link\_logical,
display\_icon\_logical, icon\_file, icon\_number, icon\_label)

Object\_class&nbsp;&nbsp;&nbsp;&nbsp;is a text string containing the
classname for the object you want to create.

  - > Object\_class is the classname corresponding to the Object Type
    > selection in the Insert Object dialog box.

  - > For more information about object classnames, consult the
    > documentation for your source application to see how it supports
    > object linking and embedding (OLE).


File\_name&nbsp;&nbsp;&nbsp;&nbsp;is a text string specifying the file
from which to create an OLE object.

Link\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value indicating
whether the new object based on file\_name should be linked to
file\_name. If it is not linked, the object is created as a copy or the
file. Link\_logical is ignored if file\_name is not specified. If
link\_logical is FALSE or omitted, then no link is established.

Display\_icon\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value
corresponding to the Display as Icon checkbox. If it is FALSE or
omitted, then the regular picture for the object is displayed. If it is
TRUE, then the icon icon\_number found in icon\_file is displayed with
the label icon\_label. If display\_icon\_logical is not TRUE, then
icon\_file, icon\_number, and icon\_label are ignored.

Icon\_file&nbsp;&nbsp;&nbsp;&nbsp;is the name of the file where the icon
to display is located.

Icon\_number&nbsp;&nbsp;&nbsp;&nbsp;is the number of the icon within
icon\_file that should be used.

Icon\_label&nbsp;&nbsp;&nbsp;&nbsp;is a text string indicating a label
to display beneath the icon. If the parameter is an empty string ("") or
is omitted, no label is displayed.

**Remarks**

  - > If INSERT.OBJECT starts another application, your macro pauses.
    > Your macro resumes when you return to Microsoft Excel.

  - > Although you will not normally use Microsoft Excel class names in
    > a Microsoft Excel macro, you may need them in macros written for
    > other applications. Microsoft Excel uses classnames
    > "Excel.Sheet.5" and "Excel.Chart.5".


**Related Function**

[EDIT.OBJECT](EDIT.OBJECT.md)&nbsp;&nbsp;&nbsp;Edits an object



Return to [README](README.md)

