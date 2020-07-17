CALCULATION

Controls when and how formulas in open workbooks are calculated. This
function is included for compatibility with Microsoft Excel version 4.0.
For controlling calculation in Microsoft Excel version 5.0 or later, see
OPTIONS.CALCULATION.

**Syntax**

**CALCULATION**(**type\_num**, iter, max\_num, max\_change, update,
precision, date\_1904, calc\_save, save\_values, alt\_exp, alt\_form)  
**CALCULATION**?(type\_num, iter, max\_num, max\_change, update,
precision, date\_1904, calc\_save, save\_values, alt\_exp, alt\_form)

Arguments correspond to check boxes and options in the Calculation
dialog box. Arguments that correspond to check boxes are logical values.
If an argument is TRUE, Microsoft Excel selects the check box; if FALSE,
Microsoft Excel clears the check box.

Type\_num    is a number from 1 to 3 indicating the type of calculation.

|               |                         |
| ------------- | ----------------------- |
| **Type\_num** | **Type of calculation** |
| 1             | Automatic               |
| 2             | Automatic except tables |
| 3             | Manual                  |

Iter    corresponds to the Iteration check box. The default is FALSE.

Max\_num    is the maximum number of iterations. The default is 100.

Max\_change    is the maximum change of each iteration. The default is
0.001.

Update    corresponds to the Update Remote References check box. The
default is TRUE.

Precision    corresponds to the Precision As Displayed check box. The
default is FALSE.

Date\_1904    corresponds to the 1904 Date System check box. The default
is FALSE in Microsoft Excel for Windows and TRUE in Microsoft Excel for
the Macintosh.

Calc\_save    corresponds to the Recalculate Before Save check box. If
calc\_save is FALSE, the workbook is not recalculated before saving when
in manual calculation mode. The default is TRUE.

Save\_values    corresponds to the Save External Link Values check box.
The default is TRUE.

Alt\_exp    corresponds to the Transition Formula Evaluation check box
in the Transition tab of the Options dialog box.

  - > If alt\_exp is TRUE, Microsoft Excel uses a set of rules
    > compatible with that of Lotus 1-2-3 when calculating formulas.
    > Text is treated as 0; TRUE and FALSE are treated as 1 and 0; and
    > certain characters in database criteria ranges are interpreted the
    > same way Lotus 1-2-3 interprets them.

  - > If alt\_exp is FALSE or omitted, Microsoft Excel calculates
    > normally.

>  

Alt\_form    corresponds to the Transition Formula Entry check box in
the Transition tab of the Options dialog box.

  - > This argument is available only in Microsoft Excel for Windows.

  - > If alt\_form is TRUE, Microsoft Excel accepts formulas entered in
    > Lotus 1-2-3 style.

  - > If alt\_form is FALSE or omitted, Microsoft Excel only accepts
    > formulas entered in Microsoft Excel style.

>  

**Note   **Microsoft Excel for Windows and Microsoft Excel for the
Macintosh use different date systems as their default. Excel for Windows
uses the 1900 date system, in which serial numbers correspond to the
dates January 1, 1900, through December 31, 9999. Excel for the
Macintosh uses the 1904 date system, in which serial numbers correspond
to the dates January 1, 1904, through December 31, 9999.

**Remarks**

Use GET.DOCUMENT to return the current calculation settings for your
workbook. For more information, see GET.DOCUMENT.

**Related Functions**

[CALCULATE.DOCUMENT](CALCULATE.DOCUMENT.md)   Calculates the active sheet only

[CALCULATE.NOW](CALCULATE.NOW.md)   Calculates all open workbooks immediately

[GET.DOCUMENT](GET.DOCUMENT.md)   Returns information about a workbook

[OPTIONS.CALCULATION](OPTIONS.CALCULATION.md)   Controls calculation

[OPTIONS.TRANSITION](OPTIONS.TRANSITION.md)   Controls transition options



Return to [README.md](README.md)

CALCULATION

Controls when and how formulas in open workbooks are calculated. This
function is included for compatibility with Microsoft Excel version 4.0.
For controlling calculation in Microsoft Excel version 5.0 or later, see
OPTIONS.CALCULATION.

**Syntax**

**CALCULATION**(**type\_num**, iter, max\_num, max\_change, update,
precision, date\_1904, calc\_save, save\_values, alt\_exp, alt\_form)  
**CALCULATION**?(type\_num, iter, max\_num, max\_change, update,
precision, date\_1904, calc\_save, save\_values, alt\_exp, alt\_form)

Arguments correspond to check boxes and options in the Calculation
dialog box. Arguments that correspond to check boxes are logical values.
If an argument is TRUE, Microsoft Excel selects the check box; if FALSE,
Microsoft Excel clears the check box.

Type\_num    is a number from 1 to 3 indicating the type of calculation.

|               |                         |
| ------------- | ----------------------- |
| **Type\_num** | **Type of calculation** |
| 1             | Automatic               |
| 2             | Automatic except tables |
| 3             | Manual                  |

Iter    corresponds to the Iteration check box. The default is FALSE.

Max\_num    is the maximum number of iterations. The default is 100.

Max\_change    is the maximum change of each iteration. The default is
0.001.

Update    corresponds to the Update Remote References check box. The
default is TRUE.

Precision    corresponds to the Precision As Displayed check box. The
default is FALSE.

Date\_1904    corresponds to the 1904 Date System check box. The default
is FALSE in Microsoft Excel for Windows and TRUE in Microsoft Excel for
the Macintosh.

Calc\_save    corresponds to the Recalculate Before Save check box. If
calc\_save is FALSE, the workbook is not recalculated before saving when
in manual calculation mode. The default is TRUE.

Save\_values    corresponds to the Save External Link Values check box.
The default is TRUE.

Alt\_exp    corresponds to the Transition Formula Evaluation check box
in the Transition tab of the Options dialog box.

  - > If alt\_exp is TRUE, Microsoft Excel uses a set of rules
    > compatible with that of Lotus 1-2-3 when calculating formulas.
    > Text is treated as 0; TRUE and FALSE are treated as 1 and 0; and
    > certain characters in database criteria ranges are interpreted the
    > same way Lotus 1-2-3 interprets them.

  - > If alt\_exp is FALSE or omitted, Microsoft Excel calculates
    > normally.

>  

Alt\_form    corresponds to the Transition Formula Entry check box in
the Transition tab of the Options dialog box.

  - > This argument is available only in Microsoft Excel for Windows.

  - > If alt\_form is TRUE, Microsoft Excel accepts formulas entered in
    > Lotus 1-2-3 style.

  - > If alt\_form is FALSE or omitted, Microsoft Excel only accepts
    > formulas entered in Microsoft Excel style.

>  

**Note   **Microsoft Excel for Windows and Microsoft Excel for the
Macintosh use different date systems as their default. Excel for Windows
uses the 1900 date system, in which serial numbers correspond to the
dates January 1, 1900, through December 31, 9999. Excel for the
Macintosh uses the 1904 date system, in which serial numbers correspond
to the dates January 1, 1904, through December 31, 9999.

**Remarks**

Use GET.DOCUMENT to return the current calculation settings for your
workbook. For more information, see GET.DOCUMENT.

**Related Functions**

[CALCULATE.DOCUMENT](CALCULATE.DOCUMENT.md)   Calculates the active sheet only

[CALCULATE.NOW](CALCULATE.NOW.md)   Calculates all open workbooks immediately

[GET.DOCUMENT](GET.DOCUMENT.md)   Returns information about a workbook

[OPTIONS.CALCULATION](OPTIONS.CALCULATION.md)   Controls calculation

[OPTIONS.TRANSITION](OPTIONS.TRANSITION.md)   Controls transition options



Return to [README.md](README.md)

