FREAD

Reads characters from a file, starting at the current position in the
file. (For more information about a file's position, see FPOS.) If FREAD
is successful, it returns the text to the cell containing FREAD and
set's the file's position to the start of the following line. If the end
of the file is reached or if FREAD can't read the file, it returns the
\#N/A error value. Use FREAD instead of FREADLN when you need to read a
specific number of characters from a text file.

**Syntax**

**FREAD**(**file\_num, num\_chars**)

File\_num    is the unique ID number of the file you want to read data
from. File\_num is returned by a previously executed FOPEN function. If
file\_num is not valid, FREAD returns the \#VALUE\! error value.

Num\_chars    specifies how many bytes to read from the file. FREAD can
read up to 255 bytes at a time.

**Example**

The following function reads the next 200 bytes from the open file
identified as FileNumber:

FREAD(FileNumber, 200)

**Related Functions**

[FOPEN](FOPEN.md)   Opens a file with the type of permission specified

[FPOS](FPOS.md)   Sets the position in a text file

[FREADLN](FREADLN.md)   Reads a line from a text file

[FWRITE](FWRITE.md)   Writes characters to a text file


