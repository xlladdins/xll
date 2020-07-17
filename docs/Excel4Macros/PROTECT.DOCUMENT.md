PROTECT.DOCUMENT

Adds protection to or removes protection from the active sheet, macro
sheet, chart, dialog sheet, module, or scenario. Use PROTECT.DOCUMENT to
prevent yourself or others from changing cell contents, or objects in a
workbook. To protect workbooks in Microsoft Excel version 5.0 or later,
see WORKBOOK.PROTECT.

**Syntax**

**PROTECT.DOCUMENT**(contents, windows, password, objects, scenarios)

**PROTECT.DOCUMENT**?( contents, windows, password, objects, scenarios)

Contents    is a logical value corresponding to the Contents check box
in the Protect Sheet dialog box.

  - > If contents is TRUE or omitted, Microsoft Excel selects the check
    > box and protects cells and chart elements on the sheet or macro
    > sheet.

  - > If contents is FALSE, Microsoft Excel clears the check box (and
    > removes protection if the correct password is supplied).

>  

Windows    is provided for compatibility with Microsoft Excel version
4.0. To protect the window placement and structure of workbooks in
Microsoft Excel version 5.0 or later, see WORKBOOK.PROTECT.

  - > If windows is TRUE, Microsoft Excel prevents a workbook's windows
    > from being moved or sized.

  - > If windows is FALSE or omitted, Microsoft Excel removes protection
    > if the correct password is supplied.

>  

Password    is the password you specify in the form of text to protect
or unprotect the file. Password is case-sensitive.

  - > If password is omitted when you protect a sheet, then you will be
    > able to remove protection without a password. This is useful if
    > you want only to protect the sheet from accidental changes.

  - > If password is omitted when you try to remove protection from a
    > sheet that was protected with a password, the normal password
    > dialog box is displayed.

  - > Passwords are not recorded into the PROTECT.DOCUMENT function when
    > you use the macro recorder.

>  

Objects    is a logical value. This argument applies only to charts,
worksheets and macro sheets. Objects corresponds to the Objects check
box in the Protect Sheet dialog box.

  - > If objects is TRUE or omitted, Microsoft Excel selects the check
    > box and protects all locked objects on the chart, worksheet or
    > macro sheet.

  - > If objects is FALSE, Microsoft Excel clears the check box.

Scenarios    is a logical value that corresponds to the Scenarios check
box on the Protect Sheet dialog box. If TRUE, Microsoft Excel protects
all the scenarios. If FALSE, the scenarios are not protected.

**Remarks**

  - > If contents and objects are FALSE, PROTECT.DOCUMENT carries out
    > the Unprotect Sheet command. If contents, or objects is TRUE, it
    > carries out the Protect Sheet command.

  - > Make sure that you hide macro sheets that protect or unprotect
    > worksheets. If you type a password directly into the function on
    > an unhidden macro sheet, then someone could see the password
    > needed to unprotect the worksheet. For example,
    > PROTECT.DOCUMENT(TRUE, TRUE, "XD1411C", TRUE).

>  

**Warning**   If you forget the password of a workbook that was
previously protected with a password, you cannot unprotect the workbook.

**Related Functions**

[CELL.PROTECTION](CELL.PROTECTION.md)   Controls protection for the selected cells

[ENTER.DATA](ENTER.DATA.md)   Turns Data Entry mode on and off

[OBJECT.PROTECTION](OBJECT.PROTECTION.md)   Controls how an object is protected

[SAVE.AS](SAVE.AS.md)   Saves a workbook and allows you to specify the name, file
type, password, backup file, and location of the workbook

[WORKBOOK.PROTECT](WORKBOOK.PROTECT.md)   Protects a workbook



Return to [README](README.md)

