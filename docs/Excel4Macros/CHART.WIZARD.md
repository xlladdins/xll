CHART.WIZARD

Equivalent to clicking the ChartWizard button on the standard or chart
toolbar. Creates a chart. It is generally easier to use the macro
recorder to enter this function on your macro sheet.

**Syntax**

**CHART.WIZARD**(long, **ref**, gallery\_num, type\_num, plot\_by,
categories, ser\_titles, legend, title, x\_title, y\_title, z\_title,
number\_cats, number\_titles)

**CHART.WIZARD**?(long, ref, gallery\_num, type\_num, plot\_by,
categories, ser\_titles, legend, title, x\_title, y\_title, z\_title,
number\_cats, number\_titles)

Long&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that determines which
type of ChartWizard button CHART.WIZARD is equivalent to.

  - > If long is TRUE, CHART.WIZARD is equivalent to using the five-step
    > ChartWizard button.

  - > If long is FALSE or omitted, CHART.WIZARD is equivalent to using
    > the two-step ChartWizard button, and gallery\_num, type\_num,
    > legend, and the title arguments are ignored.


Ref&nbsp;&nbsp;&nbsp;&nbsp;is a reference to the range of cells on the
active worksheet that contains the source data for the chart, or the
object identifier of the chart if it has already been created.

Gallery\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 15 specifying
the type of chart you want to create.

|                  |              |
| ---------------- | ------------ |
| **Gallery\_num** | **Chart**    |
| 1                | Area         |
| 2                | Bar          |
| 3                | Column       |
| 4                | Line         |
| 5                | Pie          |
| 6                | Radar        |
| 7                | XY (scatter) |
| 8                | Combination  |
| 9                | 3-D area     |
| 10               | 3-D bar      |
| 11               | 3-D column   |
| 12               | 3-D line     |
| 13               | 3-D pie      |
| 14               | 3-D surface  |
| 15               | Doughnut     |

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number identifying a formatting
option. The first formatting option in any gallery is 1.

Plot\_by&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 and specifies
whether the data for each data series is in rows or columns. 1 specifies
rows; 2 specifies columns. If plot\_by is omitted, Microsoft Excel uses
the appropriate value for the chart you're creating.

Categories&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 and specifies
whether the first row or column contains a list of x-axis labels, or
data for the first data series. 1 specifies x-axis labels; 2 specifies
the first data series. If categories is omitted, Microsoft Excel uses
the appropriate value for the chart you're creating. If number\_cats is
specified, this argument is ignored.

Ser\_titles&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 and specifies
whether the first column or row contains series titles, or data for the
first data point in each series. 1 specifies series titles; 2 specifies
the first data point. If ser\_titles is omitted, Microsoft Excel uses
the appropriate value for the chart you're creating. If number\_titles
is specified, this argument is ignored.

Legend&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 and specifies whether
to include a legend. 1 specifies a legend; 2 specifies no legend. If
legend is omitted, Microsoft Excel does not include a legend.

For the following arguments, if an argument is omitted or is empty text
(""), no title is specified.

Title&nbsp;&nbsp;&nbsp;&nbsp;is text that you want to use as a chart
title.

X\_title&nbsp;&nbsp;&nbsp;&nbsp;is text that you want to use as an
x-axis title.

Y\_title&nbsp;&nbsp;&nbsp;&nbsp;is text that you want to use as a y-axis
title.

Z\_title&nbsp;&nbsp;&nbsp;&nbsp;is text that you want to use as a z-axis
title.

Number\_cats&nbsp;&nbsp;&nbsp;&nbsp;specifies the number of rows or
columns (depending on the value of plot\_by) to use for the category
labels in the chart. This argument overrides the categories argument.

Number\_titles&nbsp;&nbsp;&nbsp;&nbsp;specifies the number of rows or
columns (depending on the value of plot\_by) to use for the series
labels in the chart. This argument overrides the ser\_title argument.

**Remarks**

If you are using the macro recorder, Microsoft Excel records a
CREATE.OBJECT and a COPY function when the chart is created and a
CHART.WIZARD function when the chart is formatted. You must precede this
function with a CREATE.OBJECT function if you are not using the macro
recorder.

**Related Function**

[CREATE.OBJECT](CREATE.OBJECT.md)&nbsp;&nbsp;&nbsp;Creates an object



Return to [README](README.md)

