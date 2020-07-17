EDIT.COLOR
==========

Equivalent to clicking the Modify button on the Color tab, which appears
when you click the Options command on the Tools menu. Defines the color
for one of the 56 color palette boxes.

Use EDIT.COLOR if you want to use a color that is not currently on the
palette and if your system hardware has more than 56 colors available.
After you set the color for the color box, any items previously
formatted with that color are displayed in the new color.

**Syntax**

**EDIT.COLOR**(**color\_num**, red\_value, green\_value, blue\_value)

**EDIT.COLOR**?(**color\_num**, red\_value, green\_value, blue\_value)

Color\_num    is a number from 1 to 56 specifying one of the 56 color
palette boxes for which you want to set the color.

Red\_value, green\_value, and blue\_value    are numbers that specify
how much red, green, and blue are in each color.

-   In Microsoft Excel for Windows, red\_value, green\_value, and
    > blue\_value are numbers from 0 to 255.

-   In Microsoft Excel for the Macintosh, red\_value, green\_value, and
    > blue\_value are also numbers from 0 to 255. However, the color
    > editing dialog box displays numbers from 0 to 65, 535. Microsoft
    > Excel automatically converts the numbers between the two ranges.
    > This allows you to display similar colors in all operating
    > environments without modifying your macros.

-   If red\_value, green\_value, and blue\_value are all set to 255, the
    > resulting color is white. If they are all set to zero, the
    > resulting color is black.

-   If red\_value, green\_value, or blue\_value is omitted, Microsoft
    > Excel assumes it to be the appropriate value for that color\_num.

>  

**Remarks**

-   Your system hardware determines the number of unique colors that you
    > can choose from and the number of colors that can be displayed on
    > the screen at the same time.

-   EDIT.COLOR does not use hue, saturation, or brightness values. If
    > you are using the macro recorder and set the color of a color
    > palette box using hue, saturation, and luminance, Microsoft Excel
    > records the corresponding red, green, and blue values instead.

-   The dialog-box form of this function, EDIT.COLOR?(color\_num),
    > displays your system\'s color editing dialog box. The default
    > red\_value, green\_value, and blue\_value are determined by the
    > current settings for the color\_num you specify. Color\_num is a
    > required argument for the dialog-box form of this function.

>  

**Related Function**

COLOR.PALETTE   Copies a color palette from one workbook to another

Return to [top](#E)

EDIT.DELETE
