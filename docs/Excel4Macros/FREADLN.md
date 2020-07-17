FREADLN

Reads characters from a file, starting at the current position in the
file and continuing to the end of the line, placing the characters in
the cell containing FREADLN. (For more information about a file's
position, see FPOS.) If FREADLN is successful, it returns the text it
read, up to but not including the carriage-return and linefeed
characters at the end of the line (in Microsoft Excel for Windows) or
the carriage-return character at the end of the line (in Microsoft Excel
for the Macintosh). If the current file position is the end of the file
or if FREADLN can't read the file, it returns the \#N/A error value.

**Syntax**

**FREADLN**(**file\_num**)

File\_num    is the unique ID number of the file you want to read data
from. File\_num is returned by a previously executed FOPEN function. If
file\_num is not valid, FREADLN returns the \#VALUE\! error value.

**Example**

The following function reads the next line from the open file identified
as FileNumber:

FREADLN(FileNumber)

**Related Functions**

[FOPEN](FOPEN.md)   Opens a file with the type of permission specified

[FPOS](FPOS.md)   Sets the position in a text file

[FREAD](FREAD.md)   Reads characters from a text file

[FWRITE](FWRITE.md)   Writes characters to a text file

[FWRITELN](FWRITELN.md)   Writes a line to a text file


