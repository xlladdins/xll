AXES

Controls whether the axes on a chart are visible. There are two syntax
forms of this function. Syntax 1 is for 2-D charts; syntax 2 is for 3-D
charts.

**Syntax 1**

For 2-D charts

**AXES**(x\_primary, y\_primary, x\_secondary, y\_secondary)

**AXES**?(x\_primary, y\_primary, x\_secondary, y\_secondary)

**Syntax 2**

For 3-D charts

**AXES**(x\_primary, y\_primary, z\_primary)

**AXES**?(x\_primary, y\_primary, z\_primary)

Arguments are logical values corresponding to the check boxes in the
Axes dialog box.

  - > If an argument is TRUE, Microsoft Excel selects the check box and
    > displays the corresponding axis.

  - > If an argument is FALSE, Microsoft Excel clears the check box and
    > hides the corresponding axis.

  - > If an argument is omitted, the display of that axis is unchanged.

>  

X\_primary    corresponds to the primary category (x) axis.

Y\_primary    corresponds to the primary value (y) axis.

Z\_primary    corresponds to the value (z) axis on the primary 3-D
chart.

X\_secondary    corresponds to the secondary category (x) axis on a 2-D
chart only.

Y\_secondary    corresponds to the secondary value (y) axis on a 2-D
chart only.

If a 2-D chart has no secondary axis, only the first two arguments are
used.

**Related Function**

GRIDLINES   Controls whether chart gridlines are visible


