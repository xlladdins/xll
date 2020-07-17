SELECT Syntax 2
===============

Equivalent to selecting objects on a chart, worksheet, or macro sheet.
There are three syntax forms of SELECT. Use syntax 2 to select an object
on which to perform an action; use one of the other syntax forms to
select cells on a worksheet or macro sheet or items on a chart.

**Syntax**

**SELECT**(object\_id\_text, replace)

Object\_id\_text    is text that identifies the object to select.
Object\_id\_text can be the name of more than one object. To give the
name of more than one object, use the following format:

SELECT(\"Oval 3, Arc 2, Line 4\")

The last item in the object\_id\_text list will be the active object.
The active object is important when moving and sizing a group of
objects. A multiple selection of objects is moved and sized relative to
the upper-left corner of the active object.

Replace    is a logical value that specifies whether previously selected
objects are included in the selection. If replace is TRUE or omitted,
Microsoft Excel only selects the objects specified by object\_id\_text;
if FALSE, it includes any objects that were previously selected. For
example, if a button is selected and a SELECT formula selects an arc and
an oval, TRUE leaves only the arc and oval selected, and FALSE includes
the button with the arc and oval.

**Remarks**

Objects can be identified by their object type and number as described
in CREATE.OBJECT, or by the unique number that specifies the order of
their creation. For example, if the third object you create is an oval,
you could use either \"oval 3\" or \"3\" as object\_id\_text.

**Examples**

The following macro formulas each select a number of objects and specify
Arc 2 as the active object:

SELECT(\"Oval 3, Arc 1, Line 4, Arc 2\")

SELECT(\"3, 1, 4, 2\")

**Related Functions**

FORMAT.MOVE   Moves the selected object

FORMAT.SIZE   Changes the size of the selected objects

GET.OBJECT   Returns information about an object

SELECTION   Returns the reference of the selection

SELECT Syntax 1   Selects cells

SELECT Syntax 3   Selects chart objects

Return to [top](#Q)

SELECT Syntax 3
