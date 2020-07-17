DELETE.BAR
==========

Deletes a custom menu bar.

**Syntax**

**DELETE.BAR**(**bar\_num**)

Bar\_num    is the ID number of the custom menu bar you want to delete.

**Tip**   Rather than trying to discover the ID number of the menu bar
you want to delete, use a reference to the ADD.BAR function that created
the bar. For example, the following macro formula deletes the menu bar
created by the ADD.BAR function in the cell named ReportsBar:

DELETE.BAR(ReportsBar)

**Related Functions**

ADD.BAR   Adds a menu bar

SHOW.BAR   Displays a menu bar

Return to [top](#A)

DELETE.CHART.AUTOFORMAT
