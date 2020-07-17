SCALE
==============

Equivalent to clicking the Selected Axes command on the Format menu when
a chart\'s value (z) axis is selected, and then clicking the Scale tab.
There are five syntax forms of this function. Syntax 5 of SCALE applies
if the selected axis is a value (z) axis on a 3-D chart. Use this syntax
of SCALE to change the position, formatting, and scaling of the value
axis.

**Syntax 5**

**SCALE**(min\_num, max\_num, major, minor, cross, logarithmic, reverse,
min)

**SCALE**?(min\_num, max\_num, major, minor, cross, logarithmic,
reverse, min)

The first five arguments correspond to the five range variables in the
Format Axis dialog box, as shown in the following list. Each argument
can be either the logical value TRUE or a number.

-   If TRUE or omitted, the Auto check box is selected.

-   If a number, that number is used.

Min\_num    corresponds to the Minimum check box and is the minimum
value for the value axis.

Max\_num    corresponds to the Maximum check box and is the maximum
value for the value axis.

Major    corresponds to the Major Unit check box and is the major unit
of measure.

Minor    corresponds to the Minor Unit check box and is the minor unit
of measure.

Cross    corresponds to the Floor (XY Plane) Crosses At check box.

The last three arguments are logical values corresponding to check boxes
on the Scale tab. If an argument is TRUE, Microsoft Excel selects the
check box; if FALSE, Microsoft Excel clears the check box.

Logarithmic    corresponds to the Logarithmic Scale check box.

Reverse    corresponds to the Values In Reverse Order check box.

Min    corresponds to the Floor (XY Plane) Crosses At Minimum Value
check box.

**Related Functions**

Syntax 1   Changes the position, formatting, and scaling of the category
axis in 2-D charts

Syntax 2   Changes the position, formatting, and scaling of the value
axis in 2-D charts

Syntax 3   Changes the position, formatting, and scaling of the category
axis in 3-D charts

Syntax 4   Changes the position, formatting, and scaling of the series
axis in 3-D charts

Return to [top](#Q)

SCENARIO.ADD
