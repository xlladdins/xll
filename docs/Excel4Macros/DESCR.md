# DESCR

Generates descriptive statistics for data in the input range.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**DESCR**(**inprng**, outrng, grouped, labels, summary, ds\_large,
ds\_small, confid)

**DESCR**?(inprng, outrng, grouped, labels, summary, ds\_large,
ds\_small, confid)

Inprng&nbsp;&nbsp;&nbsp;&nbsp;is the input range.

Outrng&nbsp;&nbsp;&nbsp;&nbsp;is the first cell (the upper-left cell) in
the output table or the name, as text, of a new sheet to contain the
output table. If FALSE, blank, or omitted, places the output table in a
new workbook.

Grouped&nbsp;&nbsp;&nbsp;&nbsp;is a text character that indicates
whether the data in the input range is organized by row or column.

  - > If grouped is "C" or omitted, then the data is organized by
    > column.

  - > If grouped is "R" then the data is organized by row.


Labels&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that describes where
the labels are located in the input range, as shown in the following
table:

|                  |             |                                                   |
| ---------------- | ----------- | ------------------------------------------------- |
| **Labels**       | **Grouped** | **Labels are in**                                 |
| TRUE             | "C"         | First row of the input range.                     |
| TRUE             | "R"         | First column of the input range.                  |
| FALSE or omitted | (ignored)   | No labels. All cells in the input range are data. |

Summary&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If TRUE, DESCR
reports the summary statistics. If FALSE or omitted, no summary
statistics are reported.

Ds\_large&nbsp;&nbsp;&nbsp;&nbsp;is an integer k. If ds\_large is
present, DESCR reports the k-th largest data point. If ds\_large is
omitted, the value is not reported.

Ds\_small&nbsp;&nbsp;&nbsp;&nbsp;is an integer k. If ds\_small is
present, DESCR reports the k-th smallest data point. If ds\_small is
omitted, the value is not reported.

Confid&nbsp;&nbsp;&nbsp;&nbsp;is the confidence level of the mean. If
confid is given, DESCR reports the confidence interval for the input
range. If confid is omitted, the confidence interval is 95%.



Return to [README](README.md#D)

# DESCR

Generates descriptive statistics for data in the input range.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**DESCR**(**inprng**, outrng, grouped, labels, summary, ds\_large,
ds\_small, confid)

**DESCR**?(inprng, outrng, grouped, labels, summary, ds\_large,
ds\_small, confid)

Inprng&nbsp;&nbsp;&nbsp;&nbsp;is the input range.

Outrng&nbsp;&nbsp;&nbsp;&nbsp;is the first cell (the upper-left cell) in
the output table or the name, as text, of a new sheet to contain the
output table. If FALSE, blank, or omitted, places the output table in a
new workbook.

Grouped&nbsp;&nbsp;&nbsp;&nbsp;is a text character that indicates
whether the data in the input range is organized by row or column.

  - > If grouped is "C" or omitted, then the data is organized by
    > column.

  - > If grouped is "R" then the data is organized by row.


Labels&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that describes where
the labels are located in the input range, as shown in the following
table:

|                  |             |                                                   |
| ---------------- | ----------- | ------------------------------------------------- |
| **Labels**       | **Grouped** | **Labels are in**                                 |
| TRUE             | "C"         | First row of the input range.                     |
| TRUE             | "R"         | First column of the input range.                  |
| FALSE or omitted | (ignored)   | No labels. All cells in the input range are data. |

Summary&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If TRUE, DESCR
reports the summary statistics. If FALSE or omitted, no summary
statistics are reported.

Ds\_large&nbsp;&nbsp;&nbsp;&nbsp;is an integer k. If ds\_large is
present, DESCR reports the k-th largest data point. If ds\_large is
omitted, the value is not reported.

Ds\_small&nbsp;&nbsp;&nbsp;&nbsp;is an integer k. If ds\_small is
present, DESCR reports the k-th smallest data point. If ds\_small is
omitted, the value is not reported.

Confid&nbsp;&nbsp;&nbsp;&nbsp;is the confidence level of the mean. If
confid is given, DESCR reports the confidence interval for the input
range. If confid is omitted, the confidence interval is 95%.



Return to [README](README.md#D)

