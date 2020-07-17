SUBSCRIBE.TO
============

Inserts the contents of the edition into the active sheet at the point
of the current selection. Use SUBSCRIBE.TO to incorporate editions
published from other workbooks into your Microsoft Excel worksheets and
macro sheets. SUBSCRIBE.TO returns TRUE if successful.

**Syntax**

**SUBSCRIBE.TO**(**file\_text, format\_num**)

**Important   **This function is only available if you are using
Microsoft Excel for the Macintosh with system software version 7.0 or
later.

File\_text    is the name, as a text string, of the edition you want to
insert into the active sheet. Unless file\_text is in the current
folder, supply the full path of the workbook. If file\_text cannot be
found, SUBSCRIBE.TO returns the \#VALUE! error value and interrupts the
macro.

**Remarks**

-   If a single cell is selected, the data from the edition file is
    > placed into as large a range of cells as is required by the data.
    > Data already present in those cells is replaced. If the data is a
    > picture, it is inserted from the upper-left corner of the selected
    > cell.

-   If a range of cells is selected, and the range is not big enough to
    > contain the edition data, Microsoft Excel displays a dialog box
    > asking if you want to clip the data to fit the range.

>  

Format\_num    is the number 1 or 2 and specifies the format type of the
file you are subscribing to.

+-------------------+-----------------------------------------------------+
| > **Format\_num** | > **Format type**                                   |
+-------------------+-----------------------------------------------------+
| > 1 or omitted    | > Picture                                           |
+-------------------+-----------------------------------------------------+
| > 2               | > Text (includes BIFF, VALU, TEXT, and CSV formats) |
+-------------------+-----------------------------------------------------+

**Related Functions**

CREATE.PUBLISHER   Creates a publisher from the selection

EDITION.OPTIONS   Sets publisher and subscriber options

GET.LINK.INFO   Returns information about a link

Return to [top](#Q)

SUBTOTAL.CREATE
