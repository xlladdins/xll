FILE.DELETE
===========

Deletes a file from the disk. Although you will normally delete files
manually, you can, for example, use FILE.DELETE in a macro to delete
temporary files created by the macro.

**Syntax**

**FILE.DELETE**(**file\_text**)

**FILE.DELETE**?(file\_text)

File\_text    is the name of the file to delete.

**Remarks**

-   If Microsoft Excel can\'t find file\_text, it displays a message
    > saying that it cannot delete the file. To avoid this, include the
    > entire path in file\_text. See the following second and fifth
    > examples. You can also use FILES to generate an array of filenames
    > and then check if the file you want to delete is in the array.

-   If a file is open when you delete it, the file is removed from the
    > disk but remains open in Microsoft Excel.

-   In the dialog-box form, FILE.DELETE?, you can use an asterisk (\*)
    > to represent any series of characters and a question mark (?) to
    > represent any single character. See the following third and sixth
    > examples.

>  

**Examples**

In Microsoft Excel for Windows, the following macro formula deletes a
file called CHART1.XLS from the current directory:

FILE.DELETE(\"CHART1.XLS\")

The following macro formula deletes a file called 92INFO.XLS kept in the
EXCEL\\SALES subdirectory:

FILE.DELETE(\"C:\\EXCEL\\SALES\\92INFO.XLS\")

The following macro formula displays the Delete dialog box listing all
documents whose extensions begin with the letters \"XL\":

FILE.DELETE?(\"\*.XL?\")

In Microsoft Excel for the Macintosh, the following macro formula
deletes a file called CHART1 from the current folder:

FILE.DELETE(\"CHART1\")

The following macro formula deletes a file called 1992 INFO kept in a
series of nested folders:

FILE.DELETE(\"HARD DISK:EXCEL 5:SALES WORKSHEETS:1992 INFO\")

The following macro formula displays the Delete dialog box listing all
documents beginning with the word \"Clients\":

FILE.DELETE?(\"Clients\*\")

**Related Functions**

FILE.CLOSE   Closes the active workbook

FILES   Returns the filenames in the specified directory or folder

Return to [top](#E)

FILES
