# FILES

Returns a horizontal text array of the names of all files in the
specified directory or folder. Use FILES to build a list of filenames
upon which you want your macro to operate.

**Syntax**

**FILES**(directory\_text)

Directory\_text&nbsp;&nbsp;&nbsp;&nbsp;specifies which directories or
folders to return filenames from.

  - > Directory\_text accepts an asterisk (\*) to represent a series of
    > characters and a question mark (?) to represent a single character
    > in filenames.

  - > If directory\_text is not specified, FILES returns filenames from
    > the current directory.


**Remarks**

If you enter FILES in a single cell, only one filename is returned. You
will normally use FILES with SET.NAME to assign the returned array to a
name. See the last example below.

**Tips**&nbsp;&nbsp;&nbsp;You can use COLUMNS to count the number of
entries in the returned array. You can use TRANSPOSE to change a
horizontal array to a vertical one.

**Examples**

In Microsoft Excel for Windows, the following macro formula returns the
names of all files starting with the letter F in the current directory
or folder:

FILES("F\*.\*")

When entered as an array formula in several cells, the following macro
formula returns the filenames in the current directory to those cells.
If the directory contains fewer files than can fit in the selected
cells, the \#N/A error value appears in the extra cells.

FILES()

In Microsoft Excel for Windows, the following macro formula returns all
files starting with "SALE" and ending with the .XLS extension in the
\\EXCEL\\CHARTS subdirectory:

FILES("C:\\EXCEL\\CHARTS\\SALE\*.XLS")

In Microsoft Excel for the Macintosh, the following macro formula
returns all files starting with "SALE" in the nested CHART folder:

FILES("DISK:EXCEL:CHART:SALE\*")

The following macro stores the names of the files in the current
directory in the named array FileArray

SET.NAME("FileArray",FILES())

**Related Functions**

[DOCUMENTS](DOCUMENTS.md)&nbsp;&nbsp;&nbsp;Returns the names of the specified open
workbooks

[FILE.DELETE](FILE.DELETE.md)&nbsp;&nbsp;&nbsp;Deletes a file

[OPEN](OPEN.md)&nbsp;&nbsp;&nbsp;Opens a workbook

[SET.NAME](SET.NAME.md)&nbsp;&nbsp;&nbsp;Defines a name as a value



Return to [README](README.md#F)

