# REGRESS

Performs multiple linear regression analysis.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**REGRESS**(**inpyrng, inpxrng**, constant, labels, confid, soutrng,
residuals, sresiduals, rplots, lplots, routrng, nplots, poutrng)

**REGRESS**?(inpyrng, inpxrng, constant, labels, confid, soutrng,
residuals, sresiduals, rplots, lplots, routrng, nplots, poutrng)

Inpyrng&nbsp;&nbsp;&nbsp;&nbsp;is the input range for the y-values
(dependent variable).

Inpxrng&nbsp;&nbsp;&nbsp;&nbsp;is the input range for the x-values
(independent variable).

Constant&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If constant is TRUE,
the y-intercept is assumed to be zero (the regression line passes
through the origin). If constant is FALSE or omitted, the y-intercept is
assumed to be a non-zero number.

Labels&nbsp;&nbsp;&nbsp;&nbsp;is a logical value.

  - > If labels is TRUE, then the first row or column of the input
    > ranges contain labels.

  - > If labels is FALSE or omitted, all cells in inpyrng and inpxrng
    > are considered data. Microsoft Excel will then generate the
    > appropriate data labels for the output table.


Confid&nbsp;&nbsp;&nbsp;&nbsp;is an additional confidence level to apply
to the regression. If omitted, confid is 95%.

Soutrng&nbsp;&nbsp;&nbsp;&nbsp;is the first cell (the upper-left cell)
in the output table or the name, as text, of the new sheet to contain
the summary output table. If FALSE, blank, or omitted, places the
summary output table in a new workbook. Microsoft Excel version 5.0 uses
a single output table for REGRESS; Microsoft Excel version 4.0 used
three different output tables for summary, residual, and probability
data.

Residuals&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If residuals is
TRUE, REGRESS includes residuals in the output table. If residuals is
FALSE or omitted, residuals are not included.

Sresiduals&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If sresiduals is
TRUE, REGRESS includes standardized residuals in the output table. If
sresiduals is FALSE or omitted, standardized residuals are not included.

Rplots&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If rplots is TRUE,
REGRESS generates separate charts for each x versus the residual. If
rplots is FALSE or omitted, separate charts are not generated.

Lplots&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If lplots is TRUE,
REGRESS generates a chart showing the regression line fitted to the
observed values. If lplots is FALSE or omitted, the chart is not
generated.

Routrng&nbsp;&nbsp;&nbsp;&nbsp;is the first cell (the upper-left cell)
in the residuals output table or the name, as text, of the new sheet to
contain the residuals output table. If FALSE, blank, or omitted, places
the residuals output table in a new worksheet. This argument is for
compatibility with Microsoft Excel version 4.0 only and is ignored in
later versions.

Nplots&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If nplots is TRUE,
REGRESS generates a chart of normal probabilities. If nplots is FALSE or
omitted, the chart is not generated.

Poutrng&nbsp;&nbsp;&nbsp;&nbsp;is the first cell (the upper-left cell)
in the probability data output table or the name, as text, of the new
sheet to contain the probability output table. If FALSE, blank, or
omitted, places the probability output table in a new worksheet. This
argument is for compatibility with Microsoft Excel version 4.0 only and
is ignored in later versions.



Return to [README](README.md#R)

# REGRESS

Performs multiple linear regression analysis.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**REGRESS**(**inpyrng, inpxrng**, constant, labels, confid, soutrng,
residuals, sresiduals, rplots, lplots, routrng, nplots, poutrng)

**REGRESS**?(inpyrng, inpxrng, constant, labels, confid, soutrng,
residuals, sresiduals, rplots, lplots, routrng, nplots, poutrng)

Inpyrng&nbsp;&nbsp;&nbsp;&nbsp;is the input range for the y-values
(dependent variable).

Inpxrng&nbsp;&nbsp;&nbsp;&nbsp;is the input range for the x-values
(independent variable).

Constant&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If constant is TRUE,
the y-intercept is assumed to be zero (the regression line passes
through the origin). If constant is FALSE or omitted, the y-intercept is
assumed to be a non-zero number.

Labels&nbsp;&nbsp;&nbsp;&nbsp;is a logical value.

  - > If labels is TRUE, then the first row or column of the input
    > ranges contain labels.

  - > If labels is FALSE or omitted, all cells in inpyrng and inpxrng
    > are considered data. Microsoft Excel will then generate the
    > appropriate data labels for the output table.


Confid&nbsp;&nbsp;&nbsp;&nbsp;is an additional confidence level to apply
to the regression. If omitted, confid is 95%.

Soutrng&nbsp;&nbsp;&nbsp;&nbsp;is the first cell (the upper-left cell)
in the output table or the name, as text, of the new sheet to contain
the summary output table. If FALSE, blank, or omitted, places the
summary output table in a new workbook. Microsoft Excel version 5.0 uses
a single output table for REGRESS; Microsoft Excel version 4.0 used
three different output tables for summary, residual, and probability
data.

Residuals&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If residuals is
TRUE, REGRESS includes residuals in the output table. If residuals is
FALSE or omitted, residuals are not included.

Sresiduals&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If sresiduals is
TRUE, REGRESS includes standardized residuals in the output table. If
sresiduals is FALSE or omitted, standardized residuals are not included.

Rplots&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If rplots is TRUE,
REGRESS generates separate charts for each x versus the residual. If
rplots is FALSE or omitted, separate charts are not generated.

Lplots&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If lplots is TRUE,
REGRESS generates a chart showing the regression line fitted to the
observed values. If lplots is FALSE or omitted, the chart is not
generated.

Routrng&nbsp;&nbsp;&nbsp;&nbsp;is the first cell (the upper-left cell)
in the residuals output table or the name, as text, of the new sheet to
contain the residuals output table. If FALSE, blank, or omitted, places
the residuals output table in a new worksheet. This argument is for
compatibility with Microsoft Excel version 4.0 only and is ignored in
later versions.

Nplots&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If nplots is TRUE,
REGRESS generates a chart of normal probabilities. If nplots is FALSE or
omitted, the chart is not generated.

Poutrng&nbsp;&nbsp;&nbsp;&nbsp;is the first cell (the upper-left cell)
in the probability data output table or the name, as text, of the new
sheet to contain the probability output table. If FALSE, blank, or
omitted, places the probability output table in a new worksheet. This
argument is for compatibility with Microsoft Excel version 4.0 only and
is ignored in later versions.



Return to [README](README.md#R)

# REGRESS

Performs multiple linear regression analysis.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**REGRESS**(**inpyrng, inpxrng**, constant, labels, confid, soutrng,
residuals, sresiduals, rplots, lplots, routrng, nplots, poutrng)

**REGRESS**?(inpyrng, inpxrng, constant, labels, confid, soutrng,
residuals, sresiduals, rplots, lplots, routrng, nplots, poutrng)

Inpyrng&nbsp;&nbsp;&nbsp;&nbsp;is the input range for the y-values
(dependent variable).

Inpxrng&nbsp;&nbsp;&nbsp;&nbsp;is the input range for the x-values
(independent variable).

Constant&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If constant is TRUE,
the y-intercept is assumed to be zero (the regression line passes
through the origin). If constant is FALSE or omitted, the y-intercept is
assumed to be a non-zero number.

Labels&nbsp;&nbsp;&nbsp;&nbsp;is a logical value.

  - > If labels is TRUE, then the first row or column of the input
    > ranges contain labels.

  - > If labels is FALSE or omitted, all cells in inpyrng and inpxrng
    > are considered data. Microsoft Excel will then generate the
    > appropriate data labels for the output table.


Confid&nbsp;&nbsp;&nbsp;&nbsp;is an additional confidence level to apply
to the regression. If omitted, confid is 95%.

Soutrng&nbsp;&nbsp;&nbsp;&nbsp;is the first cell (the upper-left cell)
in the output table or the name, as text, of the new sheet to contain
the summary output table. If FALSE, blank, or omitted, places the
summary output table in a new workbook. Microsoft Excel version 5.0 uses
a single output table for REGRESS; Microsoft Excel version 4.0 used
three different output tables for summary, residual, and probability
data.

Residuals&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If residuals is
TRUE, REGRESS includes residuals in the output table. If residuals is
FALSE or omitted, residuals are not included.

Sresiduals&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If sresiduals is
TRUE, REGRESS includes standardized residuals in the output table. If
sresiduals is FALSE or omitted, standardized residuals are not included.

Rplots&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If rplots is TRUE,
REGRESS generates separate charts for each x versus the residual. If
rplots is FALSE or omitted, separate charts are not generated.

Lplots&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If lplots is TRUE,
REGRESS generates a chart showing the regression line fitted to the
observed values. If lplots is FALSE or omitted, the chart is not
generated.

Routrng&nbsp;&nbsp;&nbsp;&nbsp;is the first cell (the upper-left cell)
in the residuals output table or the name, as text, of the new sheet to
contain the residuals output table. If FALSE, blank, or omitted, places
the residuals output table in a new worksheet. This argument is for
compatibility with Microsoft Excel version 4.0 only and is ignored in
later versions.

Nplots&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If nplots is TRUE,
REGRESS generates a chart of normal probabilities. If nplots is FALSE or
omitted, the chart is not generated.

Poutrng&nbsp;&nbsp;&nbsp;&nbsp;is the first cell (the upper-left cell)
in the probability data output table or the name, as text, of the new
sheet to contain the probability output table. If FALSE, blank, or
omitted, places the probability output table in a new worksheet. This
argument is for compatibility with Microsoft Excel version 4.0 only and
is ignored in later versions.



Return to [README](README.md#R)

