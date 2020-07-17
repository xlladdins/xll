FOPEN

Opens a file with the type of permission specified. Unlike OPEN, FOPEN
does not load the file into memory and display it; instead, FOPEN
establishes a channel with the file so that you can exchange information
with it. If the file is opened successfully, FOPEN returns a file ID
number. If it can't open the file, FOPEN returns the \#N/A error value.
Use the file ID number with other file functions (such as FREAD, FWRITE,
and FSIZE) when you want to get information from or send information to
the file.

**Syntax**

**FOPEN**(**file\_text**, access\_num)

File\_text    is the name of the file you want to open.

Access\_num    is a number from 1 to 3 specifying what type of
permission to allow to the file:

|                 |                                                                       |
| --------------- | --------------------------------------------------------------------- |
| **Access\_num** | **Type of permission**                                                |
| 1 or omitted    | Can read and write to the file (read/write permission)                |
| 2               | Can read the file, but can't write to the file (read-only permission) |
| 3               | Creates a new file with read/write permission                         |

 

  - > If the file doesn't exist and access\_num is 3, FOPEN creates a
    > new file.

  - > If the file does exist and access\_num is 3, FOPEN replaces the
    > contents of the file with any information you supply using the
    > FWRITE or FWRITELN functions.

  - > If the file doesn't exist and access\_num is 1 or 2, FOPEN returns
    > the \#N/A error value.

>  

**Remarks**

Use FCLOSE to close a file after you finish using it.

**Example**

The following function opens a file identified as FileName using
read-only permission:

FOPEN(FileName, 2)

**Related Functions**

FCLOSE   Closes a text file

FREAD   Reads characters from a text file

FWRITE   Writes characters to a text file

OPEN   Opens a workbook


