FWRITE

Writes text to a file, starting at the current position in that file.
(For more information about a file's position, see FPOS.) If FWRITE
can't write to the file, it returns the \#N/A error value.

**Syntax**

**FWRITE**(**file\_num, text**)

File\_num    is the unique ID number of the file you want to write data
to. File\_num is returned by a previously executed FOPEN function. If
file\_num is not valid, FWRITE returns the \#VALUE\! error value.

Text    is the text you want to write to the file.

**Example**

The following function writes the current month to the open file
identified as FileNumber:

FWRITE(FileNumber, TEXT(MONTH(NOW()),"mmmm"))

**Related Functions**

[FOPEN](FOPEN.md)   Opens a file with the type of permission specified

[FPOS](FPOS.md)   Sets the position in a text file

[FREAD](FREAD.md)   Reads characters from a text file

[FWRITELN](FWRITELN.md)   Writes a line to a text file



Return to [README.md](README.md)

FWRITE

Writes text to a file, starting at the current position in that file.
(For more information about a file's position, see FPOS.) If FWRITE
can't write to the file, it returns the \#N/A error value.

**Syntax**

**FWRITE**(**file\_num, text**)

File\_num    is the unique ID number of the file you want to write data
to. File\_num is returned by a previously executed FOPEN function. If
file\_num is not valid, FWRITE returns the \#VALUE\! error value.

Text    is the text you want to write to the file.

**Example**

The following function writes the current month to the open file
identified as FileNumber:

FWRITE(FileNumber, TEXT(MONTH(NOW()),"mmmm"))

**Related Functions**

[FOPEN](FOPEN.md)   Opens a file with the type of permission specified

[FPOS](FPOS.md)   Sets the position in a text file

[FREAD](FREAD.md)   Reads characters from a text file

[FWRITELN](FWRITELN.md)   Writes a line to a text file



Return to [README.md](README.md)

