# SET.PRINT.AREA

Defines the print area for the workbook&mdash;the area that prints when
you click the Print command on the File menu. Equivalent to entering a
range in the Print Area edit box on the Sheet tab in the Page Setup
dialog box, which appears when you click the Page Setup command on the
File menu.

**Syntax**

**SET.PRINT.AREA**(range)

Range&nbsp;&nbsp;&nbsp;&nbsp;is the reference to the range that you want
to be printed. If you specify no range by using a set of empty quotation
marks (""), deletes the print area.

**Remarks**

  - > If you use SET.PRINT.AREA with a multiple selection and then use
    > the PRINT function, the individual selections are printed one
    > after the other in the order they were selected.

  - > To resume printing the entire worksheet, click the Page Setup
    > command on the File menu and click the Sheet tab. Then delete the
    > range in the Print Area edit box.


**Related Functions**

[PRINT](PRINT.md)&nbsp;&nbsp;&nbsp;Prints the active sheet

[SET.PRINT.TITLES](SET.PRINT.TITLES.md)&nbsp;&nbsp;&nbsp;Identifies text to print as titles



Return to [README](README.md)

