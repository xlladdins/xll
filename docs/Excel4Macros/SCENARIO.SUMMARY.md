SCENARIO.SUMMARY
================

Equivalent to clicking the Scenarios command on the Tools menu and then
clicking the Summary button. Generates a table summarizing the results
of all the scenarios for the model on your worksheet.

**Syntax**

**SCENARIO.SUMMARY**(result\_ref, report\_type)

**SCENARIO.SUMMARY**?(result\_ref, report\_type)

Result\_ref    is a reference to the result cells you want to include in
the summary report. Normally, result\_ref refers to one or more cells
containing the formulas that depend on the changing cell values for your
model---that is, the cells that show the results of a particular
scenario.

-   If result\_ref is omitted, no result cells are included in the
    > report.

-   If result\_ref contains nonadjacent references, you must separate
    > the reference areas by commas and enclose result\_ref in an extra
    > set of parentheses.

Report\_type    is a number specifying the type of report desired.

+--------------------+------------------------------------------------+
| > **Report\_type** | > **Type of Report**                           |
+--------------------+------------------------------------------------+
| > 1 or omitted     | > A scenario summary report (Microsoft Excel   |
|                    | > version 4.0)                                 |
+--------------------+------------------------------------------------+
| > 2                | > A scenario PivotTable report. Requires       |
|                    | > result\_ref.                                 |
+--------------------+------------------------------------------------+

**Remarks**

-   SCENARIO.SUMMARY generates a summary table of the changing cell and
    > result cell values for each scenario.

-   The table is generated on a new sheet in the current workbook. The
    > sheet becomes active after SCENARIO.SUMMARY runs.

Return to [top](#Q)

SCROLLBAR.PROPERTIES
