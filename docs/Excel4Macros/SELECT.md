SELECT
===============

Selects a chart object as specified by the selection code item\_text.
There are three syntax forms of SELECT. Use syntax 3 to select a chart
item to which you want to apply formatting; use one of the other syntax
forms to select cells or objects on a worksheet or macro sheet.

**Syntax**

**SELECT**(item\_text, single\_point)

Item\_text    is a selection code from the following table which
specifies which chart object to select.

+---------------------------------------------+-----------------------+
| > **To select**                             | > **Item\_text**      |
+---------------------------------------------+-----------------------+
| > Entire chart                              | > \"Chart\"           |
+---------------------------------------------+-----------------------+
| > Plot area                                 | > \"Plot\"            |
+---------------------------------------------+-----------------------+
| > Legend                                    | > \"Legend\"          |
+---------------------------------------------+-----------------------+
| > Primary chart value axis                  | > \"Axis 1\"          |
+---------------------------------------------+-----------------------+
| > Primary chart category axis               | > \"Axis 2\"          |
+---------------------------------------------+-----------------------+
| > Secondary chart value axis or 3-D series  | > \"Axis 3\"          |
| > axis                                      |                       |
+---------------------------------------------+-----------------------+
| > Secondary chart category axis             | > \"Axis 4\"          |
+---------------------------------------------+-----------------------+
| > Chart title                               | > \"Title\"           |
+---------------------------------------------+-----------------------+
| > Label for the primary chart value axis    | > \"Text Axis 1\"     |
+---------------------------------------------+-----------------------+
| > Label for the primary chart category axis | > \"Text Axis 2\"     |
+---------------------------------------------+-----------------------+
| > Label for the primary chart series axis   | > \"Text Axis 3\"     |
+---------------------------------------------+-----------------------+
| > nth floating text item                    | > \"Text n\"          |
+---------------------------------------------+-----------------------+
| > nth arrow                                 | > \"Arrow n\"         |
+---------------------------------------------+-----------------------+
| > Major gridlines of value axis             | > \"Gridline 1\"      |
+---------------------------------------------+-----------------------+
| > Minor gridlines of value axis             | > \"Gridline 2\"      |
+---------------------------------------------+-----------------------+
| > Major gridlines of category axis          | > \"Gridline 3\"      |
+---------------------------------------------+-----------------------+
| > Minor gridlines of category axis          | > \"Gridline 4\"      |
+---------------------------------------------+-----------------------+
| > Major gridlines of series axis            | > \"Gridline 5\"      |
+---------------------------------------------+-----------------------+
| > Minor gridlines of series axis            | > \"Gridline 6\"      |
+---------------------------------------------+-----------------------+
| > Primary chart droplines                   | > \"Dropline 1\"      |
+---------------------------------------------+-----------------------+
| > Secondary chart droplines                 | > \"Dropline 2\"      |
+---------------------------------------------+-----------------------+
| > Primary chart hi-lo lines                 | > \"Hiloline 1\"      |
+---------------------------------------------+-----------------------+
| > Secondary chart hi-lo lines               | > \"Hiloline 2\"      |
+---------------------------------------------+-----------------------+
| > Primary chart up bar                      | > \"UpBar1\"          |
+---------------------------------------------+-----------------------+
| > Secondary chart up bar                    | > \"UpBar2\"          |
+---------------------------------------------+-----------------------+
| > Primary chart down bar                    | > \"DownBar1\"        |
+---------------------------------------------+-----------------------+
| > Secondary chart down bar                  | > \"DownBar2\"        |
+---------------------------------------------+-----------------------+
| > Primary chart series line                 | > \"Seriesline1\"     |
+---------------------------------------------+-----------------------+
| > Secondary chart series line               | > \"Seriesline2\"     |
+---------------------------------------------+-----------------------+
| > Entire series                             | > \"Sn\"              |
+---------------------------------------------+-----------------------+
| > Data associated with point m in series n  | > \"SnPm\"            |
| > if single\_point is TRUE                  |                       |
+---------------------------------------------+-----------------------+
| > Text attached to point m of series n      | > \"Text SnPm\"       |
+---------------------------------------------+-----------------------+
| > Series title text of series n of an area  | > \"Text Sn\"         |
| > chart                                     |                       |
+---------------------------------------------+-----------------------+
| > Base of a 3-D chart                       | > \"Floor\"           |
+---------------------------------------------+-----------------------+
| > Back of a 3-D chart                       | > \"Walls\"           |
+---------------------------------------------+-----------------------+
| > Corners of a 3-D chart                    | > \"Corners\"         |
+---------------------------------------------+-----------------------+
| > Trend line                                | > \"SnTm\"            |
+---------------------------------------------+-----------------------+
| > Error bars                                | > \"SnEm\"            |
+---------------------------------------------+-----------------------+
| > Legend Marker                             | > \"Legend Marker n\" |
+---------------------------------------------+-----------------------+
| > Legend Entry                              | > \"Legend Entry n\"  |
+---------------------------------------------+-----------------------+

For trend lines and error bars, the value m can be X or Y, depending on
which point you want to select. If m is blank, selects both.

Single\_point    is a logical value that determines whether to select a
single point. Single\_point is available only when item\_text is
\"SnPm\".

-   If single\_point is TRUE, Microsoft Excel selects a single point.

-   If single\_point is FALSE or omitted, Microsoft Excel selects a
    > single point if there is only one series in the chart or selects
    > the entire series if there is more than one series in the chart.

-   If you specify single\_point when item\_text is any value other than
    > \"SnPm\", SELECT returns an error value.

>  

**Examples**

SELECT(\"Chart\") selects the entire chart.

SELECT(\"Dropline 2\") selects the droplines of an overlay chart.

SELECT(\"S1P3\", TRUE) selects the third point in the first series.

SELECT(\"Text S1\") selects the series title text of the first series in
an area chart.

**Related Functions**

SELECTION   Returns the reference of the selection

SELECT Syntax 1   Selects cells

SELECT Syntax 2   Selects objects on worksheets

Return to [top](#Q)

SELECT.ALL
