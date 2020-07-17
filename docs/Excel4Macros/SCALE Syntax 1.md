SCALE Syntax 1
==============

Equivalent to clicking the Selected Axis command on the Format menu when
a chart\'s category (x) axis is selected, and then clicking the Scale
tab. There are five syntax forms of this function. Syntax 1 of SCALE
applies if the selected axis is a category (x) axis on a 2-D chart and
the chart is not an xy (scatter) chart. Use this syntax of SCALE to
change the position, formatting, and scaling of the category axis.

**Syntax 1**

**SCALE**(cross, cat\_labels, cat\_marks, between, max, reverse)

**SCALE**?(cross, cat\_labels, cat\_marks, between, max, reverse)

Arguments correspond to text boxes and check boxes in the Scale tab on
the Format Axis dialog box. Arguments corresponding to check boxes are
logical values. If an argument is TRUE, Microsoft Excel selects the
check box; if FALSE, Microsoft Excel clears the check box.

Cross    is a number corresponding to the Value (Y) Axis Crosses At
Category number text box. The default is 1. Cross is ignored if max is
set to TRUE.

Cat\_labels    is a number corresponding to the Number Of Categories
Between Tick Mark Labels text box. The default is 1.

Cat\_marks    is a number corresponding to the Number Of Categories
Between Tick Marks text box. The default is 1.

Between    corresponds to the Value (Y) Axis Crosses Between Categories
check box. This argument only applies if cat\_labels is set to a number
other than 1.

Max    corresponds to the Value (Y) Axis Crosses At Maximum Category
check box. If max is TRUE, it overrides any setting for cross.

Reverse    corresponds to the Categories In Reverse Order check box.

**Related Functions**

AXES   Controls whether axes on a chart are visible

GRIDLINES   Controls whether chart gridlines are visible

Syntax 2   Changes the position, formatting, and scaling of the value
axis in 2-D charts

Syntax 3   Changes the position, formatting, and scaling of the category
axis in 3-D charts

Syntax 4   Changes the position, formatting, and scaling of the series
axis in 3-D charts

Syntax 5   Changes the position, formatting, and scaling of the value
axis in 3-D charts

Return to [top](#Q)

SCALE Syntax 2
