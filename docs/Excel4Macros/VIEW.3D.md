VIEW.3D

Equivalent to clicking the 3-D View command on the Format menu in
Microsoft Excel version 4.0, available when a chart sheet is the active
sheet. Adjusts the view of the active 3-D chart. Use VIEW.3D to
emphasize different parts of your chart by viewing it from different
angles and perspectives.

**Syntax**

**VIEW.3D**(elevation, perspective, rotation, axes, height%, autoscale)

**VIEW.3D**?(elevation, perspective, rotation, axes, height%, autoscale)

Elevation    is a number from -90 to 90 specifying the viewing elevation
of the chart and is measured in degrees. Elevation corresponds to the
Elevation box in the 3-D View dialog box in Microsoft Excel version 4.0.

  - > If elevation is 0, you view the chart straight on. If elevation is
    > 90, you view the chart from above (a "bird's eye view"). If
    > elevation is -90, you view the chart from below.

  - > If elevation is omitted, the current value is used..

  - > Elevation is limited to 0 to 44 for 3-D bar charts and 0 to 80 for
    > 3-D pie charts.

>  

Perspective    is a number from 0 to 100% specifying the perspective of
the chart. Perspective corresponds to the Perspective box in the 3-D
View dialog box in Microsoft Excel version 4.0.

  - > A higher perspective value simulates a closer view.

  - > If perspective is omitted, the current value is used..

  - > Perspective is ignored on 3-D bar and pie charts.

>  

Rotation    is a number from 0 to 360 specifying the rotation of the
chart around the value (z) axis and is measured in degrees. Rotation
corresponds to the Rotation box in the 3-D View dialog box in Microsoft
Excel version 4.0. As you rotate the chart, the back and side walls are
moved so that they do not block the chart.

  - > If rotation is omitted, the current value is used..

  - > Rotation is limited to 0 to 44 for 3-D bar charts.

>  

Axes    is a logical value specifying whether axes are fixed in the
plane of the screen or can rotate with the chart. Axes corresponds to
the Right Angle Axes check box in the 3-D View dialog box in Microsoft
Excel version 4.0.

  - > If axes is TRUE, Microsoft Excel locks the axes.

  - > If axes is FALSE, Microsoft Excel allows the axes to rotate.

  - > If axes is omitted and the chart view is 3-D layout, axes is
    > assumed to be FALSE.

  - > Axes is TRUE for 3-D bar charts and ignored for 3-D pie charts.

>  

Height%    is a number from 5 to 500 specifying the height of the chart
as a percentage of the length of the base. Height% corresponds to the
Height box in the 3-D View dialog box in Microsoft Excel version 4.0.
Height% is useful for controlling the appearance of charts with many
series or data points. If height% is omitted, the current value is
used..

Autoscale    is a logical value corresponding to the Auto Scaling check
box in the 3-D View dialog box in Microsoft Excel version 4.0. If TRUE,
automatic scaling is used; if FALSE, it is not; if omitted, the current
setting is not changed.

**Related Function**

[FORMAT.MAIN](FORMAT.MAIN.md)   Formats a main chart



Return to [README](README.md)

