HELP
====

Starts or switches to Help and displays the specified custom Help topic.
Use HELP with custom Help files to create your own Help system, which
can be used just like the built-in Microsoft Excel Help.

**Syntax**

**HELP**(help\_ref)

Help\_ref    is a reference to a topic in a Help file, in the form
\"filename!topic\_number\".

-   Help\_ref must be given as text.

>  

**Remarks**

-   Microsoft Excel for Windows does not support the use of Help files
    > in the text file format for custom Help.

-   In Microsoft Excel for the Macintosh, custom Help files are plain
    > text files or text files with line breaks.

>  

**Tips**

-   In Microsoft Excel for Windows, the following macro formula switches
    > back to Microsoft Excel when Help is active:

APP.ACTIVATE()

-   The following macro formula closes Help when Help is active:

SEND.KEYS(\"%{F4}\")

**Examples**

In Microsoft Excel for Windows, the following macro formula displays the
Help topic numbered 101 in the file CUSTHELP.DOC. The Help window
remains open if the user switches to another window or application.

HELP(\"CUSTHELP.DOC!101\")

If the custom Help file is not in the current directory, specify the
full path along with the name of the file. For example:

HELP(\"C:\\EXCEL\\CUSTHELP.DOC!101\")

In Microsoft Excel for the Macintosh, the following macro formula
displays the Help topic numbered 101 in the file CUSTOM HELP:

HELP(\"CUSTOM HELP!101\")

If the custom Help file is not in the current folder, specify the full
path along with the name of the file. For example:

HELP(\"HARD DISK:EXCEL:HELP:CUSTOM HELP!101\")

Return to [top](#H)

HIDE
