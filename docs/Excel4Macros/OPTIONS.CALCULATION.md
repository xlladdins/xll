# OPTIONS.CALCULATION

Equivalent to clicking the Options command on the Tools menu, and
clicking the Calculation tab in the Options dialog box. Sets various
worksheet calculation settings.

**Syntax**

**OPTIONS.CALCULATION**(type\_num, iter, max\_num, max\_change, update,
precision, date\_1904, calc\_save, save\_values)

**OPTIONS.CALCULATION**?(type\_num, iter, max\_num, max\_change, update,
precision, date\_1904, calc\_save, save\_values)

Arguments correspond to check boxes and options in the Calculation tab
in the Options dialog box. Arguments that correspond to check boxes are
logical values. If an argument is TRUE, Microsoft Excel selects the
check box; if FALSE, Microsoft Excel clears the check box.

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 3 indicating the
type of calculation.

|               |                         |
| ------------- | ----------------------- |
| **Type\_num** | **Type of calculation** |
| 1             | Automatic               |
| 2             | Automatic except tables |
| 3             | Manual                  |

Iter&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Iteration check box. The
default is FALSE.

Max\_num&nbsp;&nbsp;&nbsp;&nbsp;is the maximum number of iterations. The
default is 100.

Max\_change&nbsp;&nbsp;&nbsp;&nbsp;is the maximum change of each
iteration. The default is 0.001.

Update&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Update Remote
References check box. The default is TRUE.

Precision&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Precision As
Displayed check box. The default is FALSE.

Date\_1904&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the 1904 Date System
check box. The default is FALSE in Microsoft Excel for Windows and TRUE
in Microsoft Excel for the Macintosh.

Calc\_save&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Recalculate Before
Save check box. If calc\_save is FALSE, the workbook is not recalculated
before saving when in manual calculation mode. The default is TRUE.

Save\_values&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Save External
Link Values check box. The default is TRUE.

**Note**&nbsp;&nbsp;&nbsp;Microsoft Excel for Windows and Microsoft
Excel for the Macintosh use different date systems as their default. For
more information, see NOW.



Return to [README](README.md)

