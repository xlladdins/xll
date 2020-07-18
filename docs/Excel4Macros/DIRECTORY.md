# DIRECTORY

Sets the current drive and directory or folder to the specified path and
returns the name of the new directory or folder as text. Use DIRECTORY
to get the name of the current directory or folder for use with the OPEN
and SAVE.AS functions or to specify a directory or folder from which to
return a list of files with the FILES function.

**Syntax**

**DIRECTORY**(**path\_text**)

Path\_text&nbsp;&nbsp;&nbsp;&nbsp;is the drive and directory or folder
you want to change to.

  - > If path\_text is not specified, DIRECTORY returns the name of the
    > current directory or folder as text.

  - > If path\_text does not specify a drive, the current drive is
    > assumed.


**Examples**

In Microsoft Excel for Windows, the following macro formula sets the
directory to \\EXCEL\\MODELS on the current drive and returns the value
"drive:\\EXCEL\\MODELS":

DIRECTORY("\\EXCEL\\MODELS")

The following macro formula sets the current drive to E and sets the
directory to \\EXCEL\\MODELS on E. It returns the value
"E:\\EXCEL\\MODELS":

DIRECTORY("E:\\EXCEL\\MODELS")

In Microsoft Excel for the Macintosh, the following macro formula sets
the folder to HARD DISK: APPS:EXCEL:FINANCIALS and returns the value
"HARD DISK:APPS:EXCEL:FINANCIALS":

DIRECTORY("HARD DISK:APPS:EXCEL:FINANCIALS")

**Related Function**

[FILES](FILES.md)&nbsp;&nbsp;&nbsp;Returns the filenames in the specified directory
or folder



Return to [README](README.md)

