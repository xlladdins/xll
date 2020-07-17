SUBTOTAL.CREATE
===============

Equivalent to clicking the Subtotals command on the Data menu. Generates
a subtotal in a list or database.

**Syntax**

**SUBTOTAL.CREATE**(**at\_change\_in**, **function\_num**, **total**,
replace, pagebreaks, summary\_below)

**SUBTOTAL.CREATE**?(at\_change\_in, function\_num, total, replace,
pagebreaks, summary\_below)

At\_change\_in    is a column offset corresponding to the At Each Change
In text box on the Subtotal dialog box.

Function\_Num    is a number corresponding to the Use Function list box
specifying which function you want to use in subtotaling your data.

+----------------+---------------------+
| > **Function** | > **Function\_Num** |
+----------------+---------------------+
| > SUM          | > 1                 |
+----------------+---------------------+
| > COUNTA       | > 2                 |
+----------------+---------------------+
| > AVERAGE      | > 3                 |
+----------------+---------------------+
| > MAX          | > 4                 |
+----------------+---------------------+
| > MIN          | > 5                 |
+----------------+---------------------+
| > PRODUCT      | > 6                 |
+----------------+---------------------+
| > COUNT        | > 7                 |
+----------------+---------------------+
| > STDEV        | > 8                 |
+----------------+---------------------+
| > STDEVP       | > 9                 |
+----------------+---------------------+
| > VAR          | > 10                |
+----------------+---------------------+
| > VARP         | > 11                |
+----------------+---------------------+

Total    is an array of column offsets corresponding to the Add Subtotal
To list box. Indicates which columns you want aggregated according to
function\_num; for example, {4,5}

Replace    is a logical value which, if TRUE, causes any previous
subtotals to be replaced by new subtotals. If FALSE or omitted,
subtotals will not be replaced.

PageBreaks    is a logical value corresponding to the Page Break Between
Groups check box which, if TRUE, creates a page break below each
subtotal. If FALSE or omitted, does not create a page break below each
subtotal.

Summary\_Below    is a logical value corresponding to the Summary Below
Data check box which, if TRUE, places subtotal rows below the data they
refer to, and a grand total at the bottom of the list. If FALSE, places
subtotal rows above the data they refer to, and a grand total just below
the header.

**Related Function**

SUBTOTAL.REMOVE   Removes all previously existing subtotals and grand
totals in a list

Return to [top](#Q)

SUBTOTAL.REMOVE
