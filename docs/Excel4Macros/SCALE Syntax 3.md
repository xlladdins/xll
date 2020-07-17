SCALE Syntax 3
==============

Equivalent to clicking the Selected Axes command on the Format menu when
a chart\'s category (x) axis is selected, and then click the Scale tab.
There are five syntax forms of this function. Syntax 3 of SCALE applies
if the selected axis is a category (x) axis on a 3-D chart. Use this
syntax of SCALE to change the position, formatting, and scaling of the
category axis.

**Syntax 3**

**SCALE**(cat\_labels, cat\_marks, reverse, between)

**SCALE**?(cat\_labels, cat\_marks, reverse, between)

Cat\_labels    is a number corresponding to the Number Of Categories
Between Tick-Mark Labels box. The default is 1. Cat\_labels can also be
a logical value. If TRUE, an automatic setting will be used. If FALSE,
or omitted, the number will be used.

Cat\_marks    is a number corresponding to the Number Of Categories
Between Tick Marks text box. The default is 1. Cat\_marks can also be a
logical value. If TRUE, an automatic setting will be used. If FALSE, or
omitted, the number will be used.

Reverse    corresponds to the Categories In Reverse Order check box. If
reverse is TRUE, Microsoft Excel selects the check box; if FALSE,
Microsoft Excel clears the check box.

Between    corresponds to the Value (Z) Axis Crosses Between Categories
check box. If between is TRUE, Microsoft Excel selects the check box and
the data points appear between categories. If between is FALSE or
omitted, Microsoft Excel clears the check box.

**Related Functions**

Syntax 1   Changes the position, formatting, and scaling of the category
axis in 2-D charts

Syntax 2   Changes the position, formatting, and scaling of the value
axis in 2-D charts

Syntax 4   Changes the position, formatting, and scaling of the series
axis in 3-D charts

Syntax 5   Changes the position, formatting, and scaling of the value
axis in 3-D charts

Return to [top](#Q)

SCALE Syntax 4
