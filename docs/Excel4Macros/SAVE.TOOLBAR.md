SAVE.TOOLBAR

Saves one or more toolbar definitions to a specified file.

**Syntax**

**SAVE.TOOLBAR**(bar\_id, filename)

Bar\_id    is either the name or number of a toolbar whose definition
you want to save or an array of toolbar names or numbers whose
definitions you want to save. Use an array to save several toolbar
definitions at the same time. For detailed information about bar\_id,
see ADD.TOOL. If bar\_id is omitted, all toolbar definitions are saved.

Filename    is text specifying the name of the destination file. If
filename does not exist, Microsoft Excel creates a new file. If filename
exists, Microsoft Excel overwrites the file. If filename is omitted,
Microsoft Excel saves the toolbar or toolbars in Username8.xlb, where
"username" is your Windows or network logon name. With Microsoft
Windows, Username8.xlb is stored in the directory where Windows is
installed; with Apple Macintosh, EXCEL TOOLBARS is stored in the
System:Preferences folder

**Examples**

In Microsoft Excel for Windows, the following macro formula saves
Toolbar6 as \\EXCDT\\TOOLFILE.XLB.

SAVE.TOOLBAR("Toolbar6", "\\EXCDT\\TOOLFILE.XLB")

In Microsoft Excel for the Macintosh, the following macro formula saves
Toolbar6 as TOOLFILE.

SAVE.TOOLBAR("Toolbar6", "TOOLFILE")

**Related Functions**

ADD.TOOL   Adds one or more tools to a toolbar

ADD.TOOLBAR   Creates a new toolbar with the specified tools

OPEN   Opens a workbook


