OBJECT.PROPERTIES

Determines how the selected object or objects are attached to the cells
beneath them and whether they are printed. The way an object is attached
to the cells beneath it affects how the object is moved or sized
whenever you move or size the cells.

**OBJECT.PROPERTIES**(placement\_type, print\_object)

**OBJECT.PROPERTIES**?(placement\_type, print\_object)

Placement\_type    is a number from 1 to 3 specifying how to attach the
selected object or objects. If placement\_type is omitted, the current
status is unchanged.

|                           |                                                              |
| ------------------------- | ------------------------------------------------------------ |
| **If placement\_type is** | **The selected object is**                                   |
| 1                         | Moved and sized with cells.                                  |
| 2                         | Moved but not sized with cells.                              |
| 3                         | Free-floating—it is not affected by moving and sizing cells. |

Print\_object    is a logical value specifying whether to print the
selected object or objects. If TRUE or omitted, the objects are printed;
if FALSE, they are not printed.

**Remarks**

If an object is not selected, OBJECT.PROPERTIES interrupts the macro and
returns the \#VALUE\! error value.

**Related Functions**

[CREATE.OBJECT](CREATE.OBJECT.md)   Creates an object

[FORMAT.MOVE](FORMAT.MOVE.md)   Moves the selected object

[FORMAT.SIZE](FORMAT.SIZE.md)   Changes the size of the selected object


