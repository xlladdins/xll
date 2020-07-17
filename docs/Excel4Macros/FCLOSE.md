FCLOSE

Closes the specified file.

**Syntax**

**FCLOSE**(**file\_num**)

File\_num    is the number of the file you want to close. File\_num is
returned by the FOPEN function that originally opened the file. If
file\_num is not a valid file number, FCLOSE halts the macro and returns
the \#VALUE\! error value.

**Examples**

The following function closes the file identified by FileNumber:

FCLOSE(FileNumber)

**Related Functions**

CLOSE   Closes the active window

FILE.CLOSE   Closes the active workbook

FOPEN   Opens a file with the type of permission specified


