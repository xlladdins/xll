# OBJECT.PROTECTION

Changes the protection status of the selected object.

**Syntax**

**OBJECT.PROTECTION**(locked, lock\_text)

**OBJECT.PROTECTION**?(locked, lock\_text)

Locked&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that determines whether
the selected object is locked or unlocked. If locked is TRUE, Microsoft
Excel locks the object; if FALSE, Microsoft Excel unlocks the object.

Lock\_text&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that determines
whether text in a text box or button can be changed. Lock\_text applies
only if the object is a text box, button, or worksheet control. If
lock\_text is TRUE or omitted, text cannot be changed; if FALSE, text
can be changed.

**Remarks**

  - > You cannot lock or unlock an individual object with
    > OBJECT.PROTECTION when protection is selected for objects in the
    > Protect Sheet dialog box.

  - > If an object is not selected, the function returns the \#VALUE\!
    > error value and halts the macro.

  - > In order for an object to be protected, you must use the
    > PROTECT.DOCUMENT(, , , TRUE) function after changing the object's
    > status with OBJECT.PROTECTION.


**Related Functions**

[PROTECT.DOCUMENT](PROTECT.DOCUMENT.md)&nbsp;&nbsp;&nbsp;Controls protection for the active
worksheet

[WORKBOOK.PROTECT](WORKBOOK.PROTECT.md)&nbsp;&nbsp;&nbsp;Controls protection for the active
workbook



Return to [README](README.md)

