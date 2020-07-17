SET.CONTROL.VALUE
=================

Changes the value for the active control, such as a list box, drop-down
box, check box, option button, scroll bar, and spinner button.

**Syntax**

**SET.CONTROL.VALUE**(value)

Value    is the value you want to change. The control interprets this
value as follows:

+------------------+--------------------------------------------------+
| > **Control**    | > **Value is**                                   |
+------------------+--------------------------------------------------+
| > List box       | > The index of the selected item. If zero, then  |
|                  | > no item is selected.                           |
+------------------+--------------------------------------------------+
| > Drop-down box  | > The index of the selected item. If zero, then  |
|                  | > no item is selected.                           |
+------------------+--------------------------------------------------+
| > Check box      | > 0 = Off\                                       |
|                  | > 1 = On\                                        |
|                  | > 2 = Mixed                                      |
+------------------+--------------------------------------------------+
| > Option button  | > 0= Off\                                        |
|                  | > 1 = On                                         |
+------------------+--------------------------------------------------+
| > Scroll bar     | > The numeric value of the control, between the  |
|                  | > maximum and minimum values                     |
+------------------+--------------------------------------------------+
| > Spinner button | > The numeric value of the control, between the  |
|                  | > maximum and minimum values                     |
+------------------+--------------------------------------------------+

**Related Functions**

ADD.LIST.ITEM   Adds an item in a list box or drop-down control on a
worksheet or dialog sheet control

REMOVE.LIST.ITEM   Removes an item in a list box or drop-down box

SELECT.LIST.ITEM   Selects an item in a list box or in a group box

CHECKBOX.PROPERTIES   Sets various properties of check box and option
box controls

SCROLLBAR.PROPERTIES   Sets the properties of the scroll bar and spinner
controls

Return to [top](#Q)

SET.CRITERIA
