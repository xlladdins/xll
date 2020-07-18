EDIT.OBJECT

Equivalent to clicking the Edit command on the (selected object) Object
submenu of the Edit menu. Starts the application associated with the
selected object and makes the object available for editing or other
actions.

**Syntax**

**EDIT.OBJECT**(verb\_num)

Verb\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying which verb to
use while working with the object, that is, what you want to do with the
object.

  - > The available verbs are determined by the object's source
    > application. 1 often specifies "edit, " and 2 often specifies
    > "play" (for sound, animation, and so on). For more information,
    > consult the documentation for the object's application to see how
    > it supports object linking and embedding (OLE).

  - > If the object does not support multiple verbs, verb\_num is
    > ignored.

  - > If verb\_num is omitted, it is assumed to be 1.

> &nbsp;

**Remarks**

Your macro pauses while you're editing the object and resumes when you
return to Microsoft Excel.

**Related Function**

[INSERT.OBJECT](INSERT.OBJECT.md)&nbsp;&nbsp;&nbsp;Creates an object of a specified type



Return to [README](README.md)

