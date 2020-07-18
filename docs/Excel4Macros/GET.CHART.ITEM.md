GET.CHART.ITEM

Returns the vertical or horizontal position of a point on a chart item.
Use these position numbers with FORMAT.MOVE and FORMAT.SIZE to change
the position and size of chart items. Position is measured in points; a
point is 1/72nd of an inch.

**Syntax**

**GET.CHART.ITEM**(**x\_y\_index**, point\_index, item\_text)

X\_y\_index&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying which of the
coordinates you want returned.

|                 |                         |
| --------------- | ----------------------- |
| **X\_y\_index** | **Coordinate returned** |
| 1               | Horizontal coordinate   |
| 2               | Vertical coordinate     |

Point\_index&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the point on
the chart item. These indexes are described below. If point\_index is
omitted, it is assumed to be 1.

  - > If the specified item is a point, point\_index must be 1.

  - > If the specified item is any line other than a data line, use the
    > following values for point\_index.

|                  |                         |
| ---------------- | ----------------------- |
| **Point\_index** | **Chart item position** |
| 1                | Lower or left           |
| 2                | Upper or right          |

&nbsp;

  - > If the selected item is a legend, plot area, chart area, or an
    > area in an area chart, use the following values for point\_index.

|                  |                         |
| ---------------- | ----------------------- |
| **Point\_index** | **Chart item position** |
| 1                | Upper left              |
| 2                | Upper middle            |
| 3                | Upper right             |
| 4                | Right middle            |
| 5                | Lower right             |
| 6                | Lower middle            |
| 7                | Lower left              |
| 8                | Left middle             |

&nbsp;

  - > If the selected item is an arrow in Microsoft Excel 4.0, use the
    > following values for point\_index. In Microsoft Excel version 5.0
    > or later, arrows are named lines, and the arrowhead position
    > returned is equivalent to the end of a line where the arrowhead
    > begins.

|                  |                         |
| ---------------- | ----------------------- |
| **Point\_index** | **Chart item position** |
| 1                | Arrow shaft             |
| 2                | Arrowhead               |

&nbsp;

  - > If the selected item is a pie slice, use the following values for
    > point\_index.

|                  |                                              |
| ---------------- | -------------------------------------------- |
| **Point\_index** | **Chart item position**                      |
| 1                | Outermost counterclockwise point             |
| 2                | Outer center point                           |
| 3                | Outermost clockwise point                    |
| 4                | Midpoint of the most clockwise radius        |
| 5                | Center point                                 |
| 6                | Midpoint of the most counterclockwise radius |

Item\_text&nbsp;&nbsp;&nbsp;&nbsp;is a selection code that specifies
which item of a chart to select. See the chart form of SELECT for the
item\_text codes to use for each item of a chart.

  - > If item\_text is omitted, it is assumed to be the currently
    > selected item.

  - > If item\_text is omitted and no item is selected, GET.CHART.ITEM
    > returns the \#VALUE\! error value.


**Remarks**

If the specified item does not exist, or if a chart is not active when
the function is carried out, the \#VALUE\! error value is returned.

**Examples**

The following macro formulas return the horizontal and vertical
locations, respectively, of the top of the main-chart value axis:

GET.CHART.ITEM(1, 2, "Axis 1")

GET.CHART.ITEM(2, 2, "Axis 1")

You could then use FORMAT.MOVE to move a floating text item to the
position returned by these two formulas.

**Related Functions**

[GET.DOCUMENT](GET.DOCUMENT.md)&nbsp;&nbsp;&nbsp;Returns information about a workbook

[GET.FORMULA](GET.FORMULA.md)&nbsp;&nbsp;&nbsp;Returns the contents of a cell



Return to [README](README.md)

