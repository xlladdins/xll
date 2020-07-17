SCALE Syntax 2
==============

Equivalent to clicking the Selected Axes command on the Format menu when
a chart\'s value (y) axis is selected, and then clicking the Scale tab.
There are five syntax forms of this function. Syntax 2 of SCALE applies
if the selected axis is a value (y) axis on a 2-D chart, or either axis
on an xy (scatter) chart. Use this syntax of SCALE to change the
position, formatting, and scaling of the value axis.

**Syntax 2**

**SCALE**(min\_num, max\_num, major, minor, cross, logarithmic, reverse,
max)

**SCALE**?(min\_num, max\_num, major, minor, cross, logarithmic,
reverse, max)

The first five arguments correspond to the five range variables on the
Scale tab. Each argument can be either the logical value TRUE or a
number:

-   If an argument is TRUE, Microsoft Excel selects the Auto check box.

-   If an argument is a number, that number is used for the variable.

>  

Min\_num    corresponds to the Minimum check box and is the minimum
value for the value axis.

Max\_num    corresponds to the Maximum check box and is the maximum
value for the value axis.

Major    corresponds to the Major Unit check box and is the major unit
of measure.

Minor    corresponds to the Minor Unit check box and is the minor unit
of measure.

Cross    corresponds to the Category (X) Axis Crosses At text box for
the value (y) axis of a 2-D chart or the Value (Y) Axis Crosses At text
box for the category (x) axis of an xy (scatter) chart.

The last three arguments are logical values corresponding to check boxes
on the Scale tab . If an argument is TRUE, Microsoft Excel selects the
check box; if FALSE, Microsoft Excel clears the check box.

Logarithmic    corresponds to the Logarithmic Scale check box.

Reverse    corresponds to the Values In Reverse Order check box.

Max    corresponds to the Category (X) Axis Crosses At Maximum Value
check box.

**Related Functions**

Syntax 1   Changes the position, formatting, and scaling of the category
axis in 2-D charts

Syntax 3   Changes the position, formatting, and scaling of the category
axis in 3-D charts

Syntax 4   Changes the position, formatting, and scaling of the series
axis in 3-D charts

Syntax 5   Changes the position, formatting, and scaling of the value
axis in 3-D charts

Return to [top](#Q)

SCALE Syntax 3
