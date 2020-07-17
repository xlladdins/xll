ASSIGN.TO.OBJECT

Assigns a macro to the currently select object.

**Syntax**

**ASSIGN.TO.OBJECT**(**macro\_ref**)

**ASSIGN.TO.OBJECT**?(macro\_ref)

Macro\_ref    is the name of, or a reference to, the macro you want to
run when the object is clicked. If macro\_ref is omitted, Microsoft
Excel no longer runs the previously specified macro (ASSIGN.TO.OBJECT is
turned off).

**Remarks**

  - > If an object is not selected, ASSIGN.TO.OBJECT returns the
    > \#VALUE\! error value and interrupts the macro.

  - > To change the macro assigned to an object, select the object and
    > use ASSIGN.TO.OBJECT again, using the reference to the new macro
    > as macro\_ref. The previous macro is replaced with the new macro.

>  

**Related Functions**

CREATE.OBJECT   Creates an object

RUN   Runs a macro


