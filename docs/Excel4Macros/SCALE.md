SCALE

Changes the position, formatting, and scaling of axes in a chart. There
are five syntax forms of this function.

Syntax 1   Changes the position, formatting, and scaling of the category
axis in 2-D charts

Syntax 2   Changes the position, formatting, and scaling of the value
axis in 2-D charts

Syntax 3   Changes the position, formatting, and scaling of the category
axis in 3-D charts

Syntax 4   Changes the position, formatting, and scaling of the series
axis in 3-D charts

Syntax 5   Changes the position, formatting, and scaling of the value
axis in 3-D charts


SCALE Syntax 1

Equivalent to clicking the Selected Axis command on the Format menu when
a chart's category (x) axis is selected, and then clicking the Scale
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


SCALE Syntax 2

Equivalent to clicking the Selected Axes command on the Format menu when
a chart's value (y) axis is selected, and then clicking the Scale tab.
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

  - > If an argument is TRUE, Microsoft Excel selects the Auto check
    > box.

  - > If an argument is a number, that number is used for the variable.

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


SCALE Syntax 3

Equivalent to clicking the Selected Axes command on the Format menu when
a chart's category (x) axis is selected, and then click the Scale tab.
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


SCALE Syntax 4

Equivalent to clicking the Selected Axes command on the Format menu when
a chart's value (y) axis is selected, and then clicking the Scale tab.
There are five syntax forms of this function. Syntax 4 of SCALE applies
if the selected axis is a series (y) axis on a 3-D chart. Use this
syntax of SCALE to change the position, formatting, and scaling of the
series axis.

**Syntax 4**

Series (y) axis, 3-D chart

**SCALE**(series\_labels, series\_marks, reverse)

**SCALE**?(series\_labels, series\_marks, reverse)

Series\_labels    is a number corresponding to the Number Of Series
Between Tick Labels text box. The default is 1. Series\_labels can also
be a logical value. If TRUE, and automatic setting will be used. If
FALSE, or omitted, the number will be used.

Series\_marks    is a number corresponding to the Number Of Series
Between Tick Marks text box. The default is 1. Series\_marks can also be
a logical value. If TRUE, and automatic setting will be used. If FALSE,
or omitted, the number will be used.

Reverse    is a logical value that corresponds to the Series In Reverse
Order check box on the Scale tab. If reverse is TRUE, Microsoft Excel
displays the series in reverse order; if FALSE or omitted, Microsoft
Excel displays the series normally.

**Related Functions**

Syntax 1   Changes the position, formatting, and scaling of the category
axis in 2-D charts

Syntax 2   Changes the position, formatting, and scaling of the value
axis in 2-D charts

Syntax 3   Changes the position, formatting, and scaling of the category
axis in 3-D charts

Syntax 5   Changes the position, formatting, and scaling of the value
axis in 3-D charts


SCALE Syntax 5

Equivalent to clicking the Selected Axes command on the Format menu when
a chart's value (z) axis is selected, and then clicking the Scale tab.
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

  - > If TRUE or omitted, the Auto check box is selected.

  - > If a number, that number is used.

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


