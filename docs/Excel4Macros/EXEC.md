EXEC

Starts a separate program. Use EXEC to start other programs with which
you want to communicate. Use EXEC with Microsoft Excel's other DDE
functions (INITIATE, EXECUTE, and SEND.KEYS) to create a channel to
another program and to send keystrokes and commands to the program.
(SEND.KEYS is available only in Microsoft Excel for Windows.)

Syntax 1 is for Microsoft Excel for Windows. Syntax 2 is for Microsoft
Excel for the Macintosh.

**Syntax 1**

For Microsoft Excel for Windows

**EXEC**(program\_text, window\_num)

**Syntax 2**

For Microsoft Excel for the Macintosh

**EXEC**(program\_text, , background, preferred\_size\_only)

**Important**   Microsoft Excel for the Macintosh requires system
software version 7.0 or later for the last two arguments of this
function.

Program\_text    is the name, as a text string, of any executable file
or, in Microsoft Excel for Windows, any data file that is associated
with an executable file.

  - > Use paths when the file or program to be started is not in the
    > current directory or folder.

  - > In Microsoft Excel for Windows, program\_text can include any
    > arguments and switches that are accepted by the program to be
    > started. Also, if program\_text is the name of a file associated
    > with a specific installed program, EXEC starts the program and
    > loads the specified file.

>  

**Note   **In Microsoft Excel for the Macintosh, you must use an extra
comma after the program\_text argument. This skips the window\_num
argument that does not apply to the Macintosh.

Window\_num    is a number from 1 to 3 that specifies how the window
containing the program should appear. Window\_num is only available for
use with Microsoft Excel for Windows. The window\_num argument is
allowed on the Macintosh, but it is ignored.

|                 |                    |
| --------------- | ------------------ |
| **Window\_num** | **Window appears** |
| 1               | Normal size        |
| 2 or omitted    | Minimized size     |
| 3               | Maximized size     |

Background    is a logical value that determines whether the program
specified by program\_text is opened as the active program or in the
background, leaving Microsoft Excel as the active program. If background
is TRUE, the program is started in the background; if FALSE or omitted,
the program is started in the foreground. Background is only available
for use with Microsoft Excel for the Macintosh and system software
version 7.0 or later.

Preferred\_size\_only    is a logical value that determines the amount
of memory allocated to the program. If preferred\_size\_only is TRUE,
the program is opened with its preferred memory allocation; if FALSE or
omitted, it opens with the available memory if greater than its minimum
requirement. Preferred\_size\_only is only available for use with
Microsoft Excel for the Macintosh and system software version 7.0 or
later. For information about changing the preferred memory size, see
your Macintosh documentation.

**Remarks**

In Microsoft Excel for Windows and in Microsoft Excel for the Macintosh
with system software version 7.0, if the EXEC function is successful, it
returns the task ID number of the started program. The task ID number is
a unique number that identifies a program. Use the task ID number in
other macro functions, such as APP.ACTIVATE, to refer to the program. In
Microsoft Excel for the Macintosh with system software version 6.0, if
EXEC is successful, it returns TRUE. If EXEC is unsuccessful, it returns
the \#VALUE\! error value.

**Examples**

In Microsoft Excel for Windows, the following macro formula starts the
program SEARCH.EXE. Use paths when the file or program to be started is
not in the current directory:

EXEC("C:\\WINDOWS\\SEARCH.EXE")

The following macro formula starts Microsoft Word for Windows and loads
the document SALES.DOC:

EXEC("C:\\WINWORD\\WINWORD.EXE C:\\MYFILES\\SALES.DOC")

In Microsoft Excel for the Macintosh, the following macro formula starts
Microsoft Word:

EXEC("HARD DISK:APPS:WORD")

**Related Functions**

APP.ACTIVATE   Switches to another application

EXECUTE   Carries out a command in another application

INITIATE   Opens a channel to another application

SEND.KEYS   Sends a key sequence to an application

TERMINATE   Closes a channel to another application

REQUEST   Requests an array of a specific type of information from an
application with which you have a dynamic data exchange (DDE) link

POKE   Sends data to another application with which you have a dynamic
data exchange (DDE) link


EXEC

Starts a separate program. Use EXEC to start other programs with which
you want to communicate. Use EXEC with Microsoft Excel's other DDE
functions (INITIATE, EXECUTE, and SEND.KEYS) to create a channel to
another program and to send keystrokes and commands to the program.
(SEND.KEYS is available only in Microsoft Excel for Windows.)

Syntax 1 is for Microsoft Excel for Windows. Syntax 2 is for Microsoft
Excel for the Macintosh.

**Syntax 1**

For Microsoft Excel for Windows

**EXEC**(program\_text, window\_num)

**Syntax 2**

For Microsoft Excel for the Macintosh

**EXEC**(program\_text, , background, preferred\_size\_only)

**Important**   Microsoft Excel for the Macintosh requires system
software version 7.0 or later for the last two arguments of this
function.

Program\_text    is the name, as a text string, of any executable file
or, in Microsoft Excel for Windows, any data file that is associated
with an executable file.

  - > Use paths when the file or program to be started is not in the
    > current directory or folder.

  - > In Microsoft Excel for Windows, program\_text can include any
    > arguments and switches that are accepted by the program to be
    > started. Also, if program\_text is the name of a file associated
    > with a specific installed program, EXEC starts the program and
    > loads the specified file.

>  

**Note   **In Microsoft Excel for the Macintosh, you must use an extra
comma after the program\_text argument. This skips the window\_num
argument that does not apply to the Macintosh.

Window\_num    is a number from 1 to 3 that specifies how the window
containing the program should appear. Window\_num is only available for
use with Microsoft Excel for Windows. The window\_num argument is
allowed on the Macintosh, but it is ignored.

|                 |                    |
| --------------- | ------------------ |
| **Window\_num** | **Window appears** |
| 1               | Normal size        |
| 2 or omitted    | Minimized size     |
| 3               | Maximized size     |

Background    is a logical value that determines whether the program
specified by program\_text is opened as the active program or in the
background, leaving Microsoft Excel as the active program. If background
is TRUE, the program is started in the background; if FALSE or omitted,
the program is started in the foreground. Background is only available
for use with Microsoft Excel for the Macintosh and system software
version 7.0 or later.

Preferred\_size\_only    is a logical value that determines the amount
of memory allocated to the program. If preferred\_size\_only is TRUE,
the program is opened with its preferred memory allocation; if FALSE or
omitted, it opens with the available memory if greater than its minimum
requirement. Preferred\_size\_only is only available for use with
Microsoft Excel for the Macintosh and system software version 7.0 or
later. For information about changing the preferred memory size, see
your Macintosh documentation.

**Remarks**

In Microsoft Excel for Windows and in Microsoft Excel for the Macintosh
with system software version 7.0, if the EXEC function is successful, it
returns the task ID number of the started program. The task ID number is
a unique number that identifies a program. Use the task ID number in
other macro functions, such as APP.ACTIVATE, to refer to the program. In
Microsoft Excel for the Macintosh with system software version 6.0, if
EXEC is successful, it returns TRUE. If EXEC is unsuccessful, it returns
the \#VALUE\! error value.

**Examples**

In Microsoft Excel for Windows, the following macro formula starts the
program SEARCH.EXE. Use paths when the file or program to be started is
not in the current directory:

EXEC("C:\\WINDOWS\\SEARCH.EXE")

The following macro formula starts Microsoft Word for Windows and loads
the document SALES.DOC:

EXEC("C:\\WINWORD\\WINWORD.EXE C:\\MYFILES\\SALES.DOC")

In Microsoft Excel for the Macintosh, the following macro formula starts
Microsoft Word:

EXEC("HARD DISK:APPS:WORD")

**Related Functions**

APP.ACTIVATE   Switches to another application

EXECUTE   Carries out a command in another application

INITIATE   Opens a channel to another application

SEND.KEYS   Sends a key sequence to an application

TERMINATE   Closes a channel to another application

REQUEST   Requests an array of a specific type of information from an
application with which you have a dynamic data exchange (DDE) link

POKE   Sends data to another application with which you have a dynamic
data exchange (DDE) link


EXEC

Starts a separate program. Use EXEC to start other programs with which
you want to communicate. Use EXEC with Microsoft Excel's other DDE
functions (INITIATE, EXECUTE, and SEND.KEYS) to create a channel to
another program and to send keystrokes and commands to the program.
(SEND.KEYS is available only in Microsoft Excel for Windows.)

Syntax 1 is for Microsoft Excel for Windows. Syntax 2 is for Microsoft
Excel for the Macintosh.

**Syntax 1**

For Microsoft Excel for Windows

**EXEC**(program\_text, window\_num)

**Syntax 2**

For Microsoft Excel for the Macintosh

**EXEC**(program\_text, , background, preferred\_size\_only)

**Important**   Microsoft Excel for the Macintosh requires system
software version 7.0 or later for the last two arguments of this
function.

Program\_text    is the name, as a text string, of any executable file
or, in Microsoft Excel for Windows, any data file that is associated
with an executable file.

  - > Use paths when the file or program to be started is not in the
    > current directory or folder.

  - > In Microsoft Excel for Windows, program\_text can include any
    > arguments and switches that are accepted by the program to be
    > started. Also, if program\_text is the name of a file associated
    > with a specific installed program, EXEC starts the program and
    > loads the specified file.

>  

**Note   **In Microsoft Excel for the Macintosh, you must use an extra
comma after the program\_text argument. This skips the window\_num
argument that does not apply to the Macintosh.

Window\_num    is a number from 1 to 3 that specifies how the window
containing the program should appear. Window\_num is only available for
use with Microsoft Excel for Windows. The window\_num argument is
allowed on the Macintosh, but it is ignored.

|                 |                    |
| --------------- | ------------------ |
| **Window\_num** | **Window appears** |
| 1               | Normal size        |
| 2 or omitted    | Minimized size     |
| 3               | Maximized size     |

Background    is a logical value that determines whether the program
specified by program\_text is opened as the active program or in the
background, leaving Microsoft Excel as the active program. If background
is TRUE, the program is started in the background; if FALSE or omitted,
the program is started in the foreground. Background is only available
for use with Microsoft Excel for the Macintosh and system software
version 7.0 or later.

Preferred\_size\_only    is a logical value that determines the amount
of memory allocated to the program. If preferred\_size\_only is TRUE,
the program is opened with its preferred memory allocation; if FALSE or
omitted, it opens with the available memory if greater than its minimum
requirement. Preferred\_size\_only is only available for use with
Microsoft Excel for the Macintosh and system software version 7.0 or
later. For information about changing the preferred memory size, see
your Macintosh documentation.

**Remarks**

In Microsoft Excel for Windows and in Microsoft Excel for the Macintosh
with system software version 7.0, if the EXEC function is successful, it
returns the task ID number of the started program. The task ID number is
a unique number that identifies a program. Use the task ID number in
other macro functions, such as APP.ACTIVATE, to refer to the program. In
Microsoft Excel for the Macintosh with system software version 6.0, if
EXEC is successful, it returns TRUE. If EXEC is unsuccessful, it returns
the \#VALUE\! error value.

**Examples**

In Microsoft Excel for Windows, the following macro formula starts the
program SEARCH.EXE. Use paths when the file or program to be started is
not in the current directory:

EXEC("C:\\WINDOWS\\SEARCH.EXE")

The following macro formula starts Microsoft Word for Windows and loads
the document SALES.DOC:

EXEC("C:\\WINWORD\\WINWORD.EXE C:\\MYFILES\\SALES.DOC")

In Microsoft Excel for the Macintosh, the following macro formula starts
Microsoft Word:

EXEC("HARD DISK:APPS:WORD")

**Related Functions**

APP.ACTIVATE   Switches to another application

EXECUTE   Carries out a command in another application

INITIATE   Opens a channel to another application

SEND.KEYS   Sends a key sequence to an application

TERMINATE   Closes a channel to another application

REQUEST   Requests an array of a specific type of information from an
application with which you have a dynamic data exchange (DDE) link

POKE   Sends data to another application with which you have a dynamic
data exchange (DDE) link


EXEC

Starts a separate program. Use EXEC to start other programs with which
you want to communicate. Use EXEC with Microsoft Excel's other DDE
functions (INITIATE, EXECUTE, and SEND.KEYS) to create a channel to
another program and to send keystrokes and commands to the program.
(SEND.KEYS is available only in Microsoft Excel for Windows.)

Syntax 1 is for Microsoft Excel for Windows. Syntax 2 is for Microsoft
Excel for the Macintosh.

**Syntax 1**

For Microsoft Excel for Windows

**EXEC**(program\_text, window\_num)

**Syntax 2**

For Microsoft Excel for the Macintosh

**EXEC**(program\_text, , background, preferred\_size\_only)

**Important**   Microsoft Excel for the Macintosh requires system
software version 7.0 or later for the last two arguments of this
function.

Program\_text    is the name, as a text string, of any executable file
or, in Microsoft Excel for Windows, any data file that is associated
with an executable file.

  - > Use paths when the file or program to be started is not in the
    > current directory or folder.

  - > In Microsoft Excel for Windows, program\_text can include any
    > arguments and switches that are accepted by the program to be
    > started. Also, if program\_text is the name of a file associated
    > with a specific installed program, EXEC starts the program and
    > loads the specified file.

>  

**Note   **In Microsoft Excel for the Macintosh, you must use an extra
comma after the program\_text argument. This skips the window\_num
argument that does not apply to the Macintosh.

Window\_num    is a number from 1 to 3 that specifies how the window
containing the program should appear. Window\_num is only available for
use with Microsoft Excel for Windows. The window\_num argument is
allowed on the Macintosh, but it is ignored.

|                 |                    |
| --------------- | ------------------ |
| **Window\_num** | **Window appears** |
| 1               | Normal size        |
| 2 or omitted    | Minimized size     |
| 3               | Maximized size     |

Background    is a logical value that determines whether the program
specified by program\_text is opened as the active program or in the
background, leaving Microsoft Excel as the active program. If background
is TRUE, the program is started in the background; if FALSE or omitted,
the program is started in the foreground. Background is only available
for use with Microsoft Excel for the Macintosh and system software
version 7.0 or later.

Preferred\_size\_only    is a logical value that determines the amount
of memory allocated to the program. If preferred\_size\_only is TRUE,
the program is opened with its preferred memory allocation; if FALSE or
omitted, it opens with the available memory if greater than its minimum
requirement. Preferred\_size\_only is only available for use with
Microsoft Excel for the Macintosh and system software version 7.0 or
later. For information about changing the preferred memory size, see
your Macintosh documentation.

**Remarks**

In Microsoft Excel for Windows and in Microsoft Excel for the Macintosh
with system software version 7.0, if the EXEC function is successful, it
returns the task ID number of the started program. The task ID number is
a unique number that identifies a program. Use the task ID number in
other macro functions, such as APP.ACTIVATE, to refer to the program. In
Microsoft Excel for the Macintosh with system software version 6.0, if
EXEC is successful, it returns TRUE. If EXEC is unsuccessful, it returns
the \#VALUE\! error value.

**Examples**

In Microsoft Excel for Windows, the following macro formula starts the
program SEARCH.EXE. Use paths when the file or program to be started is
not in the current directory:

EXEC("C:\\WINDOWS\\SEARCH.EXE")

The following macro formula starts Microsoft Word for Windows and loads
the document SALES.DOC:

EXEC("C:\\WINWORD\\WINWORD.EXE C:\\MYFILES\\SALES.DOC")

In Microsoft Excel for the Macintosh, the following macro formula starts
Microsoft Word:

EXEC("HARD DISK:APPS:WORD")

**Related Functions**

APP.ACTIVATE   Switches to another application

EXECUTE   Carries out a command in another application

INITIATE   Opens a channel to another application

SEND.KEYS   Sends a key sequence to an application

TERMINATE   Closes a channel to another application

REQUEST   Requests an array of a specific type of information from an
application with which you have a dynamic data exchange (DDE) link

POKE   Sends data to another application with which you have a dynamic
data exchange (DDE) link


EXEC

Starts a separate program. Use EXEC to start other programs with which
you want to communicate. Use EXEC with Microsoft Excel's other DDE
functions (INITIATE, EXECUTE, and SEND.KEYS) to create a channel to
another program and to send keystrokes and commands to the program.
(SEND.KEYS is available only in Microsoft Excel for Windows.)

Syntax 1 is for Microsoft Excel for Windows. Syntax 2 is for Microsoft
Excel for the Macintosh.

**Syntax 1**

For Microsoft Excel for Windows

**EXEC**(program\_text, window\_num)

**Syntax 2**

For Microsoft Excel for the Macintosh

**EXEC**(program\_text, , background, preferred\_size\_only)

**Important**   Microsoft Excel for the Macintosh requires system
software version 7.0 or later for the last two arguments of this
function.

Program\_text    is the name, as a text string, of any executable file
or, in Microsoft Excel for Windows, any data file that is associated
with an executable file.

  - > Use paths when the file or program to be started is not in the
    > current directory or folder.

  - > In Microsoft Excel for Windows, program\_text can include any
    > arguments and switches that are accepted by the program to be
    > started. Also, if program\_text is the name of a file associated
    > with a specific installed program, EXEC starts the program and
    > loads the specified file.

>  

**Note   **In Microsoft Excel for the Macintosh, you must use an extra
comma after the program\_text argument. This skips the window\_num
argument that does not apply to the Macintosh.

Window\_num    is a number from 1 to 3 that specifies how the window
containing the program should appear. Window\_num is only available for
use with Microsoft Excel for Windows. The window\_num argument is
allowed on the Macintosh, but it is ignored.

|                 |                    |
| --------------- | ------------------ |
| **Window\_num** | **Window appears** |
| 1               | Normal size        |
| 2 or omitted    | Minimized size     |
| 3               | Maximized size     |

Background    is a logical value that determines whether the program
specified by program\_text is opened as the active program or in the
background, leaving Microsoft Excel as the active program. If background
is TRUE, the program is started in the background; if FALSE or omitted,
the program is started in the foreground. Background is only available
for use with Microsoft Excel for the Macintosh and system software
version 7.0 or later.

Preferred\_size\_only    is a logical value that determines the amount
of memory allocated to the program. If preferred\_size\_only is TRUE,
the program is opened with its preferred memory allocation; if FALSE or
omitted, it opens with the available memory if greater than its minimum
requirement. Preferred\_size\_only is only available for use with
Microsoft Excel for the Macintosh and system software version 7.0 or
later. For information about changing the preferred memory size, see
your Macintosh documentation.

**Remarks**

In Microsoft Excel for Windows and in Microsoft Excel for the Macintosh
with system software version 7.0, if the EXEC function is successful, it
returns the task ID number of the started program. The task ID number is
a unique number that identifies a program. Use the task ID number in
other macro functions, such as APP.ACTIVATE, to refer to the program. In
Microsoft Excel for the Macintosh with system software version 6.0, if
EXEC is successful, it returns TRUE. If EXEC is unsuccessful, it returns
the \#VALUE\! error value.

**Examples**

In Microsoft Excel for Windows, the following macro formula starts the
program SEARCH.EXE. Use paths when the file or program to be started is
not in the current directory:

EXEC("C:\\WINDOWS\\SEARCH.EXE")

The following macro formula starts Microsoft Word for Windows and loads
the document SALES.DOC:

EXEC("C:\\WINWORD\\WINWORD.EXE C:\\MYFILES\\SALES.DOC")

In Microsoft Excel for the Macintosh, the following macro formula starts
Microsoft Word:

EXEC("HARD DISK:APPS:WORD")

**Related Functions**
**Related Functions**

APP.ACTIVATE   Switches to another application

EXECUTE   Carries out a command in another application

INITIATE   Opens a channel to another application

SEND.KEYS   Sends a key sequence to an application

TERMINATE   Closes a channel to another application

REQUEST   Requests an array of a specific type of information from an
application with which you have a dynamic data exchange (DDE) link

POKE   Sends data to another application with which you have a dynamic
data exchange (DDE) link


EXEC

Starts a separate program. Use EXEC to start other programs with which
you want to communicate. Use EXEC with Microsoft Excel's other DDE
functions (INITIATE, EXECUTE, and SEND.KEYS) to create a channel to
another program and to send keystrokes and commands to the program.
(SEND.KEYS is available only in Microsoft Excel for Windows.)

Syntax 1 is for Microsoft Excel for Windows. Syntax 2 is for Microsoft
Excel for the Macintosh.

**Syntax 1**

For Microsoft Excel for Windows

**EXEC**(program\_text, window\_num)

**Syntax 2**

For Microsoft Excel for the Macintosh

**EXEC**(program\_text, , background, preferred\_size\_only)

**Important**   Microsoft Excel for the Macintosh requires system
software version 7.0 or later for the last two arguments of this
function.

Program\_text    is the name, as a text string, of any executable file
or, in Microsoft Excel for Windows, any data file that is associated
with an executable file.

  - > Use paths when the file or program to be started is not in the
    > current directory or folder.

  - > In Microsoft Excel for Windows, program\_text can include any
    > arguments and switches that are accepted by the program to be
    > started. Also, if program\_text is the name of a file associated
    > with a specific installed program, EXEC starts the program and
    > loads the specified file.

>  

**Note   **In Microsoft Excel for the Macintosh, you must use an extra
comma after the program\_text argument. This skips the window\_num
argument that does not apply to the Macintosh.

Window\_num    is a number from 1 to 3 that specifies how the window
containing the program should appear. Window\_num is only available for
use with Microsoft Excel for Windows. The window\_num argument is
allowed on the Macintosh, but it is ignored.

|                 |                    |
| --------------- | ------------------ |
| **Window\_num** | **Window appears** |
| 1               | Normal size        |
| 2 or omitted    | Minimized size     |
| 3               | Maximized size     |

Background    is a logical value that determines whether the program
specified by program\_text is opened as the active program or in the
background, leaving Microsoft Excel as the active program. If background
is TRUE, the program is started in the background; if FALSE or omitted,
the program is started in the foreground. Background is only available
for use with Microsoft Excel for the Macintosh and system software
version 7.0 or later.

Preferred\_size\_only    is a logical value that determines the amount
of memory allocated to the program. If preferred\_size\_only is TRUE,
the program is opened with its preferred memory allocation; if FALSE or
omitted, it opens with the available memory if greater than its minimum
requirement. Preferred\_size\_only is only available for use with
Microsoft Excel for the Macintosh and system software version 7.0 or
later. For information about changing the preferred memory size, see
your Macintosh documentation.

**Remarks**

In Microsoft Excel for Windows and in Microsoft Excel for the Macintosh
with system software version 7.0, if the EXEC function is successful, it
returns the task ID number of the started program. The task ID number is
a unique number that identifies a program. Use the task ID number in
other macro functions, such as APP.ACTIVATE, to refer to the program. In
Microsoft Excel for the Macintosh with system software version 6.0, if
EXEC is successful, it returns TRUE. If EXEC is unsuccessful, it returns
the \#VALUE\! error value.

**Examples**

In Microsoft Excel for Windows, the following macro formula starts the
program SEARCH.EXE. Use paths when the file or program to be started is
not in the current directory:

EXEC("C:\\WINDOWS\\SEARCH.EXE")

The following macro formula starts Microsoft Word for Windows and loads
the document SALES.DOC:

EXEC("C:\\WINWORD\\WINWORD.EXE C:\\MYFILES\\SALES.DOC")

In Microsoft Excel for the Macintosh, the following macro formula starts
Microsoft Word:

EXEC("HARD DISK:APPS:WORD")

**Related Functions**

APP.ACTIVATE   Switches to another application

EXECUTE   Carries out a command in another application

INITIATE   Opens a channel to another application

SEND.KEYS   Sends a key sequence to an application

TERMINATE   Closes a channel to another application

REQUEST   Requests an array of a specific type of information from an
application with which you have a dynamic data exchange (DDE) link

POKE   Sends data to another application with which you have a dynamic
data exchange (DDE) link


EXEC

Starts a separate program. Use EXEC to start other programs with which
you want to communicate. Use EXEC with Microsoft Excel's other DDE
functions (INITIATE, EXECUTE, and SEND.KEYS) to create a channel to
another program and to send keystrokes and commands to the program.
(SEND.KEYS is available only in Microsoft Excel for Windows.)

Syntax 1 is for Microsoft Excel for Windows. Syntax 2 is for Microsoft
Excel for the Macintosh.

**Syntax 1**

For Microsoft Excel for Windows

**EXEC**(program\_text, window\_num)

**Syntax 2**

For Microsoft Excel for the Macintosh

**EXEC**(program\_text, , background, preferred\_size\_only)

**Important**   Microsoft Excel for the Macintosh requires system
software version 7.0 or later for the last two arguments of this
function.

Program\_text    is the name, as a text string, of any executable file
or, in Microsoft Excel for Windows, any data file that is associated
with an executable file.

  - > Use paths when the file or program to be started is not in the
    > current directory or folder.

  - > In Microsoft Excel for Windows, program\_text can include any
    > arguments and switches that are accepted by the program to be
    > started. Also, if program\_text is the name of a file associated
    > with a specific installed program, EXEC starts the program and
    > loads the specified file.

>  

**Note   **In Microsoft Excel for the Macintosh, you must use an extra
comma after the program\_text argument. This skips the window\_num
argument that does not apply to the Macintosh.

Window\_num    is a number from 1 to 3 that specifies how the window
containing the program should appear. Window\_num is only available for
use with Microsoft Excel for Windows. The window\_num argument is
allowed on the Macintosh, but it is ignored.

|                 |                    |
| --------------- | ------------------ |
| **Window\_num** | **Window appears** |
| 1               | Normal size        |
| 2 or omitted    | Minimized size     |
| 3               | Maximized size     |

Background    is a logical value that determines whether the program
specified by program\_text is opened as the active program or in the
background, leaving Microsoft Excel as the active program. If background
is TRUE, the program is started in the background; if FALSE or omitted,
the program is started in the foreground. Background is only available
for use with Microsoft Excel for the Macintosh and system software
version 7.0 or later.

Preferred\_size\_only    is a logical value that determines the amount
of memory allocated to the program. If preferred\_size\_only is TRUE,
the program is opened with its preferred memory allocation; if FALSE or
omitted, it opens with the available memory if greater than its minimum
requirement. Preferred\_size\_only is only available for use with
Microsoft Excel for the Macintosh and system software version 7.0 or
later. For information about changing the preferred memory size, see
your Macintosh documentation.

**Remarks**

In Microsoft Excel for Windows and in Microsoft Excel for the Macintosh
with system software version 7.0, if the EXEC function is successful, it
returns the task ID number of the started program. The task ID number is
a unique number that identifies a program. Use the task ID number in
other macro functions, such as APP.ACTIVATE, to refer to the program. In
Microsoft Excel for the Macintosh with system software version 6.0, if
EXEC is successful, it returns TRUE. If EXEC is unsuccessful, it returns
the \#VALUE\! error value.

**Examples**

In Microsoft Excel for Windows, the following macro formula starts the
program SEARCH.EXE. Use paths when the file or program to be started is
not in the current directory:

EXEC("C:\\WINDOWS\\SEARCH.EXE")

The following macro formula starts Microsoft Word for Windows and loads
the document SALES.DOC:

EXEC("C:\\WINWORD\\WINWORD.EXE C:\\MYFILES\\SALES.DOC")

In Microsoft Excel for the Macintosh, the following macro formula starts
Microsoft Word:

EXEC("HARD DISK:APPS:WORD")

**Related Functions**

APP.ACTIVATE   Switches to another application

EXECUTE   Carries out a command in another application

INITIATE   Opens a channel to another application

SEND.KEYS   Sends a key sequence to an application

TERMINATE   Closes a channel to another application

REQUEST   Requests an array of a specific type of information from an
application with which you have a dynamic data exchange (DDE) link

POKE   Sends data to another application with which you have a dynamic
data exchange (DDE) link


EXEC

Starts a separate program. Use EXEC to start other programs with which
you want to communicate. Use EXEC with Microsoft Excel's other DDE
functions (INITIATE, EXECUTE, and SEND.KEYS) to create a channel to
another program and to send keystrokes and commands to the program.
(SEND.KEYS is available only in Microsoft Excel for Windows.)

Syntax 1 is for Microsoft Excel for Windows. Syntax 2 is for Microsoft
Excel for the Macintosh.

**Syntax 1**

For Microsoft Excel for Windows

**EXEC**(program\_text, window\_num)

**Syntax 2**

For Microsoft Excel for the Macintosh

**EXEC**(program\_text, , background, preferred\_size\_only)

**Important**   Microsoft Excel for the Macintosh requires system
software version 7.0 or later for the last two arguments of this
function.

Program\_text    is the name, as a text string, of any executable file
or, in Microsoft Excel for Windows, any data file that is associated
with an executable file.

  - > Use paths when the file or program to be started is not in the
    > current directory or folder.

  - > In Microsoft Excel for Windows, program\_text can include any
    > arguments and switches that are accepted by the program to be
    > started. Also, if program\_text is the name of a file associated
    > with a specific installed program, EXEC starts the program and
    > loads the specified file.

>  

**Note   **In Microsoft Excel for the Macintosh, you must use an extra
comma after the program\_text argument. This skips the window\_num
argument that does not apply to the Macintosh.

Window\_num    is a number from 1 to 3 that specifies how the window
containing the program should appear. Window\_num is only available for
use with Microsoft Excel for Windows. The window\_num argument is
allowed on the Macintosh, but it is ignored.

|                 |                    |
| --------------- | ------------------ |
| **Window\_num** | **Window appears** |
| 1               | Normal size        |
| 2 or omitted    | Minimized size     |
| 3               | Maximized size     |

Background    is a logical value that determines whether the program
specified by program\_text is opened as the active program or in the
background, leaving Microsoft Excel as the active program. If background
is TRUE, the program is started in the background; if FALSE or omitted,
the program is started in the foreground. Background is only available
for use with Microsoft Excel for the Macintosh and system software
version 7.0 or later.

Preferred\_size\_only    is a logical value that determines the amount
of memory allocated to the program. If preferred\_size\_only is TRUE,
the program is opened with its preferred memory allocation; if FALSE or
omitted, it opens with the available memory if greater than its minimum
requirement. Preferred\_size\_only is only available for use with
Microsoft Excel for the Macintosh and system software version 7.0 or
later. For information about changing the preferred memory size, see
your Macintosh documentation.

**Remarks**

In Microsoft Excel for Windows and in Microsoft Excel for the Macintosh
with system software version 7.0, if the EXEC function is successful, it
returns the task ID number of the started program. The task ID number is
a unique number that identifies a program. Use the task ID number in
other macro functions, such as APP.ACTIVATE, to refer to the program. In
Microsoft Excel for the Macintosh with system software version 6.0, if
EXEC is successful, it returns TRUE. If EXEC is unsuccessful, it returns
the \#VALUE\! error value.

**Examples**

In Microsoft Excel for Windows, the following macro formula starts the
program SEARCH.EXE. Use paths when the file or program to be started is
not in the current directory:

EXEC("C:\\WINDOWS\\SEARCH.EXE")

The following macro formula starts Microsoft Word for Windows and loads
the document SALES.DOC:

EXEC("C:\\WINWORD\\WINWORD.EXE C:\\MYFILES\\SALES.DOC")

In Microsoft Excel for the Macintosh, the following macro formula starts
Microsoft Word:

EXEC("HARD DISK:APPS:WORD")

**Related FunctioEXEC

Starts a separate program. Use EXEC to start other programs with which
you want to communicate. Use EXEC with Microsoft Excel's other DDE
functions (INITIATE, EXECUTE, and SEND.KEYS) to create a channel to
another program and to send keystrokes and commands to the program.
(SEND.KEYS is available only in Microsoft Excel for Windows.)

Syntax 1 is for Microsoft Excel for Windows. Syntax 2 is for Microsoft
Excel for the Macintosh.

**Syntax 1**

For Microsoft Excel for Windows

**EXEC**(program\_text, window\_num)

**Syntax 2**

For Microsoft Excel for the Macintosh

**EXEC**(program\_text, , background, preferred\_size\_only)

**Important**   Microsoft Excel for the Macintosh requires system
software version 7.0 or later for the last two arguments of this
function.

Program\_text    is the name, as a text string, of any executable file
or, in Microsoft Excel for Windows, any data file that is associated
with an executable file.

  - > Use paths when the file or program to be started is not in the
    > current directory or folder.

  - > In Microsoft Excel for Windows, program\_text can include any
    > arguments and switches that are accepted by the program to be
    > started. Also, if program\_text is the name of a file associated
    > with a specific installed program, EXEC starts the program and
    > loads the specified file.

>  

**Note   **In Microsoft Excel for the Macintosh, you must use an extra
comma after the program\_text argument. This skips the window\_num
argument that does not apply to the Macintosh.

Window\_num    is a number from 1 to 3 that specifies how the window
containing the program should appear. Window\_num is only available for
use with Microsoft Excel for Windows. The window\_num argument is
allowed on the Macintosh, but it is ignored.

|                 |                    |
| --------------- | ------------------ |
| **Window\_num** | **Window appears** |
| 1               | Normal size        |
| 2 or omitted    | Minimized size     |
| 3               | Maximized size     |

Background    is a logical value that determines whether the program
specified by program\_text is opened as the active program or in the
background, leaving Microsoft Excel as the active program. If background
is TRUE, the program is started in the background; if FALSE or omitted,
the program is started in the foreground. Background is only available
for use with Microsoft Excel for the Macintosh and system software
version 7.0 or later.

Preferred\_size\_only    is a logical value that determines the amount
of memory allocated to the program. If preferred\_size\_only is TRUE,
the program is opened with its preferred memory allocation; if FALSE or
omitted, it opens with the available memory if greater than its minimum
requirement. Preferred\_size\_only is only available for use with
Microsoft Excel for the Macintosh and system software version 7.0 or
later. For information about changing the preferred memory size, see
your Macintosh documentation.

**Remarks**

In Microsoft Excel for Windows and in Microsoft Excel for the Macintosh
with system software version 7.0, if the EXEC function is successful, it
returns the task ID number of the started program. The task ID number is
a unique number that identifies a program. Use the task ID number in
other macro functions, such as APP.ACTIVATE, to refer to the program. In
Microsoft Excel for the Macintosh with system software version 6.0, if
EXEC is successful, it returns TRUE. If EXEC is unsuccessful, it returns
the \#VALUE\! error value.

**Examples**

In Microsoft Excel for Windows, the following macro formula starts the
program SEARCH.EXE. Use paths when the file or program to be started is
not in the current directory:

EXEC("C:\\WINDOWS\\SEARCH.EXE")

The following macro formula starts Microsoft Word for Windows and loads
the document SALES.DOC:

EXEC("C:\\WINWORD\\WINWORD.EXE C:\\MYFILES\\SALES.DOC")

In Microsoft Excel for the Macintosh, the following macro formula starts
Microsoft Word:

EXEC("HARD DISK:APPS:WORD")

**Related Functions**

APP.ACTIVATE   Switches to another application

EXECUTE   Carries out a command in another application

INITIATE   Opens a channel to another application

SEND.KEYS   Sends a key sequence to an application

TERMINATE   Closes a channel to another application

REQUEST   Requests an array of a specific type of information from an
application with which you have a dynamic data exchange (DDE) link

POKE   Sends data to another application with which you have a dynamic
data exchange (DDE) link


EXEC

Starts a separate program. Use EXEC to start other programs with which
you want to communicate. Use EXEC with Microsoft Excel's other DDE
functions (INITIATE, EXECUTE, and SEND.KEYS) to create a channel to
another program and to send keystrokes and commands to the program.
(SEND.KEYS is available only in Microsoft Excel for Windows.)

Syntax 1 is for Microsoft Excel for Windows. Syntax 2 is for Microsoft
Excel for the Macintosh.

**Syntax 1**

For Microsoft Excel for Windows

**EXEC**(program\_text, window\_num)

**Syntax 2**

For Microsoft Excel for the Macintosh

**EXEC**(program\_text, , background, preferred\_size\_only)

**Important**   Microsoft Excel for the Macintosh requires system
software version 7.0 or later for the last two arguments of this
function.

Program\_text    is the name, as a text string, of any executable file
or, in Microsoft Excel for Windows, any data file that is associated
with an executable file.

  - > Use paths when the file or program to be started is not in the
    > current directory or folder.

  - > In Microsoft Excel for Windows, program\_text can include any
    > arguments and switches that are accepted by the program to be
    > started. Also, if program\_text is the name of a file associated
    > with a specific installed program, EXEC starts the program and
    > loads the specified file.

>  

**Note   **In Microsoft Excel for the Macintosh, you must use an extra
comma after the program\_text argument. This skips the window\_num
argument that does not apply to the Macintosh.

Window\_num    is a number from 1 to 3 that specifies how the window
containing the program should appear. Window\_num is only available for
use with Microsoft Excel for Windows. The window\_num argument is
allowed on the Macintosh, but it is ignored.

|                 |                    |
| --------------- | ------------------ |
| **Window\_num** | **Window appears** |
| 1               | Normal size        |
| 2 or omitted    | Minimized size     |
| 3               | Maximized size     |

Background    is a logical value that determines whether the program
specified by program\_text is opened as the active program or in the
background, leaving Microsoft Excel as the active program. If background
is TRUE, the program is started in the background; if FALSE or omitted,
the program is started in the foreground. Background is only available
for use with Microsoft Excel for the Macintosh and system software
version 7.0 or later.

Preferred\_size\_only    is a logical value that determines the amount
of memory allocated to the program. If preferred\_size\_only is TRUE,
the program is opened with its preferred memory allocation; if FALSE or
omitted, it opens with the available memory if greater than its minimum
requirement. Preferred\_size\_only is only available for use with
Microsoft Excel for the Macintosh and system software version 7.0 or
later. For information about changing the preferred memory size, see
your Macintosh documentation.

**Remarks**

In Microsoft Excel for Windows and in Microsoft Excel for the Macintosh
with system software version 7.0, if the EXEC function is successful, it
returns the task ID number of the started program. The task ID number is
a unique number that identifies a program. Use the task ID number in
other macro functions, such as APP.ACTIVATE, to refer to the program. In
Microsoft Excel for the Macintosh with system software version 6.0, if
EXEC is successful, it returns TRUE. If EXEC is unsuccessful, it returns
the \#VALUE\! error value.

**Examples**

In Microsoft Excel for Windows, the following macro formula starts the
program SEARCH.EXE. Use paths when the file or program to be started is
not in the current directory:

EXEC("C:\\WINDOWS\\SEARCH.EXE")

The following macro formula starts Microsoft Word for Windows and loads
the document SALES.DOC:

EXEC("C:\\WINWORD\\WINWORD.EXE C:\\MYFILES\\SALES.DOC")

In Microsoft Excel for the Macintosh, the following macro formula starts
Microsoft Word:

EXEC("HARD DISK:APPS:WORD")

**Related FunctioEXEC

Starts a separate program. Use EXEC to start other programs with which
you want to communicate. Use EXEC with Microsoft Excel's other DDE
functions (INITIATE, EXECUTE, and SEND.KEYS) to create a channel to
another program and to send keystrokes and commands to the program.
(SEND.KEYS is available only in Microsoft Excel for Windows.)

Syntax 1 is for Microsoft Excel for Windows. Syntax 2 is for Microsoft
Excel for the Macintosh.

**Syntax 1**

For Microsoft Excel for Windows

**EXEC**(program\_text, window\_num)

**Syntax 2**

For Microsoft Excel for the Macintosh

**EXEC**(program\_text, , background, preferred\_size\_only)

**Important**   Microsoft Excel for the Macintosh requires system
software version 7.0 or later for the last two arguments of this
function.

Program\_text    is the name, as a text string, of any executable file
or, in Microsoft Excel for Windows, any data file that is associated
with an executable file.

  - > Use paths when the file or program to be started is not in the
    > current directory or folder.

  - > In Microsoft Excel for Windows, program\_text can include any
    > arguments and switches that are accepted by the program to be
    > started. Also, if program\_text is the name of a file associated
    > with a specific installed program, EXEC starts the program and
    > loads the specified file.

>  

**Note   **In Microsoft Excel for the Macintosh, you must use an extra
comma after the program\_text argument. This skips the window\_num
argument that does not apply to the Macintosh.

Window\_num    is a number from 1 to 3 that specifies how the window
containing the program should appear. Window\_num is only available for
use with Microsoft Excel for Windows. The window\_num argument is
allowed on the Macintosh, but it is ignored.

|                 |                    |
| --------------- | ------------------ |
| **Window\_num** | **Window appears** |
| 1               | Normal size        |
| 2 or omitted    | Minimized size     |
| 3               | Maximized size     |

Background    is a logical value that determines whether the program
specified by program\_text is opened as the active program or in the
background, leaving Microsoft Excel as the active program. If background
is TRUE, the program is started in the background; if FALSE or omitted,
the program is started in the foreground. Background is only available
for use with Microsoft Excel for the Macintosh and system software
version 7.0 or later.

Preferred\_size\_only    is a logical value that determines the amount
of memory allocated to the program. If preferred\_size\_only is TRUE,
the program is opened with its preferred memory allocation; if FALSE or
omitted, it opens with the available memory if greater than its minimum
requirement. Preferred\_size\_only is only available for use with
Microsoft Excel for the Macintosh and system software version 7.0 or
later. For information about changing the preferred memory size, see
your Macintosh documentation.

**Remarks**

In Microsoft Excel for Windows and in Microsoft Excel for the Macintosh
with system software version 7.0, if the EXEC function is successful, it
returns the task ID number of the started program. The task ID number is
a unique number that identifies a program. Use the task ID number in
other macro functions, such as APP.ACTIVATE, to refer to the program. In
Microsoft Excel for the Macintosh with system software version 6.0, if
EXEC is successful, it returns TRUE. If EXEC is unsuccessful, it returns
the \#VALUE\! error value.

**Examples**

In Microsoft Excel for Windows, the following macro formula starts the
program SEARCH.EXE. Use paths when the file or program to be started is
not in the current directory:

EXEC("C:\\WINDOWS\\SEARCH.EXE")

The following macro formula starts Microsoft Word for Windows and loads
the document SALES.DOC:

EXEC("C:\\WINWORD\\WINWORD.EXE C:\\MYFILES\\SALES.DOC")

In Microsoft Excel for the Macintosh, the following macro formula starts
Microsoft Word:

EXEC("HARD DISK:APPS:WORD")

**Related FunctioEXEC

Starts a separate program. Use EXEC to start other programs with which
you want to communicate. Use EXEC with Microsoft Excel's other DDE
functions (INITIATE, EXECUTE, and SEND.KEYS) to create a channel to
another program and to send keystrokes and commands to the program.
(SEND.KEYS is available only in Microsoft Excel for Windows.)

Syntax 1 is for Microsoft Excel for Windows. Syntax 2 is for Microsoft
Excel for the Macintosh.

**Syntax 1**

For Microsoft Excel for Windows

**EXEC**(program\_text, window\_num)

**Syntax 2**

For Microsoft Excel for the Macintosh

**EXEC**(program\_text, , background, preferred\_size\_only)

**Important**   Microsoft Excel for the Macintosh requires system
software version 7.0 or later for the last two arguments of this
function.

Program\_text    is the name, as a text string, of any executable file
or, in Microsoft Excel for Windows, any data file that is associated
with an executable file.

  - > Use paths when the file or program to be started is not in the
    > current directory or folder.

  - > In Microsoft Excel for Windows, program\_text can include any
    > arguments and switches that are accepted by the program to be
    > started. Also, if program\_text is the name of a file associated
    > with a specific installed program, EXEC starts the program and
    > loads the specified file.

>  

**Note   **In Microsoft Excel for the Macintosh, you must use an extra
comma after the program\_text argument. This skips the window\_num
argument that does not apply to the Macintosh.

Window\_num    is a number from 1 to 3 that specifies how the window
containing the program should appear. Window\_num is only available for
use with Microsoft Excel for Windows. The window\_num argument is
allowed on the Macintosh, but it is ignored.

|                 |                    |
| --------------- | ------------------ |
| **Window\_num** | **Window appears** |
| 1               | Normal size        |
| 2 or omitted    | Minimized size     |
| 3               | Maximized size     |

Background    is a logical value that determines whether the program
specified by program\_text is opened as the active program or in the
background, leaving Microsoft Excel as the active program. If background
is TRUE, the program is started in the background; if FALSE or omitted,
the program is started in the foreground. Background is only available
for use with Microsoft Excel for the Macintosh and system software
version 7.0 or later.

Preferred\_size\_only    is a logical value that determines the amount
of memory allocated to the program. If preferred\_size\_only is TRUE,
the program is opened with its preferred memory allocation; if FALSE or
omitted, it opens with the available memory if greater than its minimum
requirement. Preferred\_size\_only is only available for use with
Microsoft Excel for the Macintosh and system software version 7.0 or
later. For information about changing the preferred memory size, see
your Macintosh documentation.

**Remarks**

In Microsoft Excel for Windows and in Microsoft Excel for the Macintosh
with system software version 7.0, if the EXEC function is successful, it
returns the task ID number of the started program. The task ID number is
a unique number that identifies a program. Use the task ID number in
other macro functions, such as APP.ACTIVATE, to refer to the program. In
Microsoft Excel for the Macintosh with system software version 6.0, if
EXEC is successful, it returns TRUE. If EXEC is unsuccessful, it returns
the \#VALUE\! error value.

**Examples**

In Microsoft Excel for Windows, the following macro formula starts the
program SEARCH.EXE. Use paths when the file or program to be started is
not in the current directory:

EXEC("C:\\WINDOWS\\SEARCH.EXE")

The following macro formula starts Microsoft Word for Windows and loads
the document SALES.DOC:

EXEC("C:\\WINWORD\\WINWORD.EXE C:\\MYFILES\\SALES.DOC")

In Microsoft Excel for the Macintosh, the following macro formula starts
Microsoft Word:

EXEC("HARD DISK:APPS:WORD")

**Related FunctioEXEC

Starts a separate program. Use EXEC to start other programs with which
you want to communicate. Use EXEC with Microsoft Excel's other DDE
functions (INITIATE, EXECUTE, and SEND.KEYS) to create a channel to
another program and to send keystrokes and commands to the program.
(SEND.KEYS is available only in Microsoft Excel for Windows.)

Syntax 1 is for Microsoft Excel for Windows. Syntax 2 is for Microsoft
Excel for the Macintosh.

**Syntax 1**

For Microsoft Excel for Windows

**EXEC**(program\_text, window\_num)

**Syntax 2**

For Microsoft Excel for the Macintosh

**EXEC**(program\_text, , background, preferred\_size\_only)

**Important**   Microsoft Excel for the Macintosh requires system
software version 7.0 or later for the last two arguments of this
function.

Program\_text    is the name, as a text string, of any executable file
or, in Microsoft Excel for Windows, any data file that is associated
with an executable file.

  - > Use paths when the file or program to be started is not in the
    > current directory or folder.

  - > In Microsoft Excel for Windows, program\_text can include any
    > arguments and switches that are accepted by the program to be
    > started. Also, if program\_text is the name of a file associated
    > with a specific installed program, EXEC starts the program and
    > loads the specified file.

>  

**Note   **In Microsoft Excel for the Macintosh, you must use an extra
comma after the program\_text argument. This skips the window\_num
argument that does not apply to the Macintosh.

Window\_num    is a number from 1 to 3 that specifies how the window
containing the program should appear. Window\_num is only available for
use with Microsoft Excel for Windows. The window\_num argument is
allowed on the Macintosh, but it is ignored.

|                 |                    |
| --------------- | ------------------ |
| **Window\_num** | **Window appears** |
| 1               | Normal size        |
| 2 or omitted    | Minimized size     |
| 3               | Maximized size     |

Background    is a logical value that determines whether the program
specified by program\_text is opened as the active program or in the
background, leaving Microsoft Excel as the active program. If background
is TRUE, the program is started in the background; if FALSE or omitted,
the program is started in the foreground. Background is only available
for use with Microsoft Excel for the Macintosh and system software
version 7.0 or later.

Preferred\_size\_only    is a logical value that determines the amount
of memory allocated to the program. If preferred\_size\_only is TRUE,
the program is opened with its preferred memory allocation; if FALSE or
omitted, it opens with the available memory if greater than its minimum
requirement. Preferred\_size\_only is only available for use with
Microsoft Excel for the Macintosh and system software version 7.0 or
later. For information about changing the preferred memory size, see
your Macintosh documentation.

**Remarks**

In Microsoft Excel for Windows and in Microsoft Excel for the Macintosh
with system software version 7.0, if the EXEC function is successful, it
returns the task ID number of the started program. The task ID number is
a unique number that identifies a program. Use the task ID number in
other macro functions, such as APP.ACTIVATE, to refer to the program. In
Microsoft Excel for the Macintosh with system software version 6.0, if
EXEC is successful, it returns TRUE. If EXEC is unsuccessful, it returns
the \#VALUE\! error value.

**Examples**

In Microsoft Excel for Windows, the following macro formula starts the
program SEARCH.EXE. Use paths when the file or program to be started is
not in the current directory:

EXEC("C:\\WINDOWS\\SEARCH.EXE")

The following macro formula starts Microsoft Word for Windows and loads
the document SALES.DOC:

EXEC("C:\\WINWORD\\WINWORD.EXE C:\\MYFILES\\SALES.DOC")

In Microsoft Excel for the Macintosh, the following macro formula starts
Microsoft Word:

EXEC("HARD DISK:APPS:WORD")

**Related FunctioEXEC

Starts a separate program. Use EXEC to start other programs with which
you want to communicate. Use EXEC with Microsoft Excel's other DDE
functions (INITIATE, EXECUTE, and SEND.KEYS) to create a channel to
another program and to send keystrokes and commands to the program.
(SEND.KEYS is available only in Microsoft Excel for Windows.)

Syntax 1 is for Microsoft Excel for Windows. Syntax 2 is for Microsoft
Excel for the Macintosh.

**Syntax 1**

For Microsoft Excel for Windows

**EXEC**(program\_text, window\_num)

**Syntax 2**

For Microsoft Excel for the Macintosh

**EXEC**(program\_text, , background, preferred\_size\_only)

**Important**   Microsoft Excel for the Macintosh requires system
software version 7.0 or later for the last two arguments of this
function.

Program\_text    is the name, as a text string, of any executable file
or, in Microsoft Excel for Windows, any data file that is associated
with an executable file.

  - > Use paths when the file or program to be started is not in the
    > current directory or folder.

  - > In Microsoft Excel for Windows, program\_text can include any
    > arguments and switches that are accepted by the program to be
    > started. Also, if program\_text is the name of a file associated
    > with a specific installed program, EXEC starts the program and
    > loads the specified file.

>  

**Note   **In Microsoft Excel for the Macintosh, you must use an extra
comma after the program\_text argument. This skips the window\_num
argument that does not apply to the Macintosh.

Window\_num    is a number from 1 to 3 that specifies how the window
containing the program should appear. Window\_num is only available for
use with Microsoft Excel for Windows. The window\_num argument is
allowed on the Macintosh, but it is ignored.

|                 |                    |
| --------------- | ------------------ |
| **Window\_num** | **Window appears** |
| 1               | Normal size        |
| 2 or omitted    | Minimized size     |
| 3               | Maximized size     |

Background    is a logical value that determines whether the program
specified by program\_text is opened as the active program or in the
background, leaving Microsoft Excel as the active program. If background
is TRUE, the program is started in the background; if FALSE or omitted,
the program is started in the foreground. Background is only available
for use with Microsoft Excel for the Macintosh and system software
version 7.0 or later.

Preferred\_size\_only    is a logical value that determines the amount
of memory allocated to the program. If preferred\_size\_only is TRUE,
the program is opened with its preferred memory allocation; if FALSE or
omitted, it opens with the available memory if greater than its minimum
requirement. Preferred\_size\_only is only available for use with
Microsoft Excel for the Macintosh and system software version 7.0 or
later. For information about changing the preferred memory size, see
your Macintosh documentation.

**Remarks**

In Microsoft Excel for Windows and in Microsoft Excel for the Macintosh
with system software version 7.0, if the EXEC function is successful, it
returns the task ID number of the started program. The task ID number is
a unique number that identifies a program. Use the task ID number in
other macro functions, such as APP.ACTIVATE, to refer to the program. In
Microsoft Excel for the Macintosh with system software version 6.0, if
EXEC is successful, it returns TRUE. If EXEC is unsuccessful, it returns
the \#VALUE\! error value.

**Examples**

In Microsoft Excel for Windows, the following macro formula starts the
program SEARCH.EXE. Use paths when the file or program to be started is
not in the current directory:

EXEC("C:\\WINDOWS\\SEARCH.EXE")

The following macro formula starts Microsoft Word for Windows and loads
the document SALES.DOC:

EXEC("C:\\WINWORD\\WINWORD.EXE C:\\MYFILES\\SALES.DOC")

In Microsoft Excel for the Macintosh, the following macro formula starts
Microsoft Word:

EXEC("HARD DISK:APPS:WORD")

**Related FunctioEXEC

Starts a separate program. Use EXEC to start other programs with which
you want to communicate. Use EXEC with Microsoft Excel's other DDE
functions (INITIATE, EXECUTE, and SEND.KEYS) to create a channel to
another program and to send keystrokes and commands to the program.
(SEND.KEYS is available only in Microsoft Excel for Windows.)

Syntax 1 is for Microsoft Excel for Windows. Syntax 2 is for Microsoft
Excel for the Macintosh.

**Syntax 1**

For Microsoft Excel for Windows

**EXEC**(program\_text, window\_num)

**Syntax 2**

For Microsoft Excel for the Macintosh

**EXEC**(program\_text, , background, preferred\_size\_only)

**Important**   Microsoft Excel for the Macintosh requires system
software version 7.0 or later for the last two arguments of this
function.

Program\_text    is the name, as a text string, of any executable file
or, in Microsoft Excel for Windows, any data file that is associated
with an executable file.

  - > Use paths when the file or program to be started is not in the
    > current directory or folder.

  - > In Microsoft Excel for Windows, program\_text can include any
    > arguments and switches that are accepted by the program to be
    > started. Also, if program\_text is the name of a file associated
    > with a specific installed program, EXEC starts the program and
    > loads the specified file.

>  

**Note   **In Microsoft Excel for the Macintosh, you must use an extra
comma after the program\_text argument. This skips the window\_num
argument that does not apply to the Macintosh.

Window\_num    is a number from 1 to 3 that specifies how the window
containing the program should appear. Window\_num is only available for
use with Microsoft Excel for Windows. The window\_num argument is
allowed on the Macintosh, but it is ignored.

|                 |                    |
| --------------- | ------------------ |
| **Window\_num** | **Window appears** |
| 1               | Normal size        |
| 2 or omitted    | Minimized size     |
| 3               | Maximized size     |

Background    is a logical value that determines whether the program
specified by program\_text is opened as the active program or in the
background, leaving Microsoft Excel as the active program. If background
is TRUE, the program is started in the background; if FALSE or omitted,
the program is started in the foreground. Background is only available
for use with Microsoft Excel for the Macintosh and system software
version 7.0 or later.

Preferred\_size\_only    is a logical value that determines the amount
of memory allocated to the program. If preferred\_size\_only is TRUE,
the program is opened with its preferred memory allocation; if FALSE or
omitted, it opens with the available memory if greater than its minimum
requirement. Preferred\_size\_only is only available for use with
Microsoft Excel for the Macintosh and system software version 7.0 or
later. For information about changing the preferred memory size, see
your Macintosh documentation.

**Remarks**

In Microsoft Excel for Windows and in Microsoft Excel for the Macintosh
with system software version 7.0, if the EXEC function is successful, it
returns the task ID number of the started program. The task ID number is
a unique number that identifies a program. Use the task ID number in
other macro functions, such as APP.ACTIVATE, to refer to the program. In
Microsoft Excel for the Macintosh with system software version 6.0, if
EXEC is successful, it returns TRUE. If EXEC is unsuccessful, it returns
the \#VALUE\! error value.

**Examples**

In Microsoft Excel for Windows, the following macro formula starts the
program SEARCH.EXE. Use paths when the file or program to be started is
not in the current directory:

EXEC("C:\\WINDOWS\\SEARCH.EXE")

The following macro formula starts Microsoft Word for Windows and loads
the document SALES.DOC:

EXEC("C:\\WINWORD\\WINWORD.EXE C:\\MYFILES\\SALES.DOC")

In Microsoft Excel for the Macintosh, the following macro formula starts
Microsoft Word:

EXEC("HARD DISK:APPS:WORD")

**Related FunctioEXEC

Starts a separate program. Use EXEC to start other programs with which
you want to communicate. Use EXEC with Microsoft Excel's other DDE
functions (INITIATE, EXECUTE, and SEND.KEYS) to create a channel to
another program and to send keystrokes and commands to the program.
(SEND.KEYS is available only in Microsoft Excel for Windows.)

Syntax 1 is for Microsoft Excel for Windows. Syntax 2 is for Microsoft
Excel for the Macintosh.

**Syntax 1**

For Microsoft Excel for Windows

**EXEC**(program\_text, window\_num)

**Syntax 2**

For Microsoft Excel for the Macintosh

**EXEC**(program\_text, , background, preferred\_size\_only)

**Important**   Microsoft Excel for the Macintosh requires system
software version 7.0 or later for the last two arguments of this
function.

Program\_text    is the name, as a text string, of any executable file
or, in Microsoft Excel for Windows, any data file that is associated
with an executable file.

  - > Use paths when the file or program to be started is not in the
    > current directory or folder.

  - > In Microsoft Excel for Windows, program\_text can include any
    > arguments and switches that are accepted by the program to be
    > started. Also, if program\_text is the name of a file associated
    > with a specific installed program, EXEC starts the program and
    > loads the specified file.

>  

**Note   **In Microsoft Excel for the Macintosh, you must use an extra
comma after the program\_text argument. This skips the window\_num
argument that does not apply to the Macintosh.

Window\_num    is a number from 1 to 3 that specifies how the window
containing the program should appear. Window\_num is only available for
use with Microsoft Excel for Windows. The window\_num argument is
allowed on the Macintosh, but it is ignored.

|                 |                    |
| --------------- | ------------------ |
| **Window\_num** | **Window appears** |
| 1               | Normal size        |
| 2 or omitted    | Minimized size     |
| 3               | Maximized size     |

Background    is a logical value that determines whether the program
specified by program\_text is opened as the active program or in the
background, leaving Microsoft Excel as the active program. If background
is TRUE, the program is started in the background; if FALSE or omitted,
the program is started in the foreground. Background is only available
for use with Microsoft Excel for the Macintosh and system software
version 7.0 or later.

Preferred\_size\_only    is a logical value that determines the amount
of memory allocated to the program. If preferred\_size\_only is TRUE,
the program is opened with its preferred memory allocation; if FALSE or
omitted, it opens with the available memory if greater than its minimum
requirement. Preferred\_size\_only is only available for use with
Microsoft Excel for the Macintosh and system software version 7.0 or
later. For information about changing the preferred memory size, see
your Macintosh documentation.

**Remarks**

In Microsoft Excel for Windows and in Microsoft Excel for the Macintosh
with system software version 7.0, if the EXEC function is successful, it
returns the task ID number of the started program. The task ID number is
a unique number that identifies a program. Use the task ID number in
other macro functions, such as APP.ACTIVATE, to refer to the program. In
Microsoft Excel for the Macintosh with system software version 6.0, if
EXEC is successful, it returns TRUE. If EXEC is unsuccessful, it returns
the \#VALUE\! error value.

**Examples**

In Microsoft Excel for Windows, the following macro formula starts the
program SEARCH.EXE. Use paths when the file or program to be started is
not in the current directory:

EXEC("C:\\WINDOWS\\SEARCH.EXE")

The following macro formula starts Microsoft Word for Windows and loads
the document SALES.DOC:

EXEC("C:\\WINWORD\\WINWORD.EXE C:\\MYFILES\\SALES.DOC")

In Microsoft Excel for the Macintosh, the following macro formula starts
Microsoft Word:

EXEC("HARD DISK:APPS:WORD")

**Related FunctioEXEC

Starts a separate program. Use EXEC to start other programs with which
you want to communicate. Use EXEC with Microsoft Excel's other DDE
functions (INITIATE, EXECUTE, and SEND.KEYS) to create a channel to
another program and to send keystrokes and commands to the program.
(SEND.KEYS is available only in Microsoft Excel for Windows.)

Syntax 1 is for Microsoft Excel for Windows. Syntax 2 is for Microsoft
Excel for the Macintosh.

**Syntax 1**

For Microsoft Excel for Windows

**EXEC**(program\_text, window\_num)

**Syntax 2**

For Microsoft Excel for the Macintosh

**EXEC**(program\_text, , background, preferred\_size\_only)

**Important**   Microsoft Excel for the Macintosh requires system
software version 7.0 or later for the last two arguments of this
function.

Program\_text    is the name, as a text string, of any executable file
or, in Microsoft Excel for Windows, any data file that is associated
with an executable file.

  - > Use paths when the file or program to be started is not in the
    > current directory or folder.

  - > In Microsoft Excel for Windows, program\_text can include any
    > arguments and switches that are accepted by the program to be
    > started. Also, if program\_text is the name of a file associated
    > with a specific installed program, EXEC starts the program and
    > loads the specified file.

>  

**Note   **In Microsoft Excel for the Macintosh, you must use an extra
comma after the program\_text argument. This skips the window\_num
argument that does not apply to the Macintosh.

Window\_num    is a number from 1 to 3 that specifies how the window
containing the program should appear. Window\_num is only available for
use with Microsoft Excel for Windows. The window\_num argument is
allowed on the Macintosh, but it is ignored.

|                 |                    |
| --------------- | ------------------ |
| **Window\_num** | **Window appears** |
| 1               | Normal size        |
| 2 or omitted    | Minimized size     |
| 3               | Maximized size     |

Background    is a logical value that determines whether the program
specified by program\_text is opened as the active program or in the
background, leaving Microsoft Excel as the active program. If background
is TRUE, the program is started in the background; if FALSE or omitted,
the program is started in the foreground. Background is only available
for use with Microsoft Excel for the Macintosh and system software
version 7.0 or later.

Preferred\_size\_only    is a logical value that determines the amount
of memory allocated to the program. If preferred\_size\_only is TRUE,
the program is opened with its preferred memory allocation; if FALSE or
omitted, it opens with the available memory if greater than its minimum
requirement. Preferred\_size\_only is only available for use with
Microsoft Excel for the Macintosh and system software version 7.0 or
later. For information about changing the preferred memory size, see
your Macintosh documentation.

**Remarks**

In Microsoft Excel for Windows and in Microsoft Excel for the Macintosh
with system software version 7.0, if the EXEC function is successful, it
returns the task ID number of the started program. The task ID number is
a unique number that identifies a program. Use the task ID number in
other macro functions, such as APP.ACTIVATE, to refer to the program. In
Microsoft Excel for the Macintosh with system software version 6.0, if
EXEC is successful, it returns TRUE. If EXEC is unsuccessful, it returns
the \#VALUE\! error value.

**Examples**

In Microsoft Excel for Windows, the following macro formula starts the
program SEARCH.EXE. Use paths when the file or program to be started is
not in the current directory:

EXEC("C:\\WINDOWS\\SEARCH.EXE")

The following macro formula starts Microsoft Word for Windows and loads
the document SALES.DOC:

EXEC("C:\\WINWORD\\WINWORD.EXE C:\\MYFILES\\SALES.DOC")

In Microsoft Excel for the Macintosh, the following macro formula starts
Microsoft Word:

EXEC("HARD DISK:APPS:WORD")

**Related FunctioEXEC

Starts a separate program. Use EXEC to start other programs with which
you want to communicate. Use EXEC with Microsoft Excel's other DDE
functions (INITIATE, EXECUTE, and SEND.KEYS) to create a channel to
another program and to send keystrokes and commands to the program.
(SEND.KEYS is available only in Microsoft Excel for Windows.)

Syntax 1 is for Microsoft Excel for Windows. Syntax 2 is for Microsoft
Excel for the Macintosh.

**Syntax 1**

For Microsoft Excel for Windows

**EXEC**(program\_text, window\_num)

**Syntax 2**

For Microsoft Excel for the Macintosh

**EXEC**(program\_text, , background, preferred\_size\_only)

**Important**   Microsoft Excel for the Macintosh requires system
software version 7.0 or later for the last two arguments of this
function.

Program\_text    is the name, as a text string, of any executable file
or, in Microsoft Excel for Windows, any data file that is associated
with an executable file.

  - > Use paths when the file or program to be started is not in the
    > current directory or folder.

  - > In Microsoft Excel for Windows, program\_text can include any
    > arguments and switches that are accepted by the program to be
    > started. Also, if program\_text is the name of a file associated
    > with a specific installed program, EXEC starts the program and
    > loads the specified file.

>  

**Note   **In Microsoft Excel for the Macintosh, you must use an extra
comma after the program\_text argument. This skips the window\_num
argument that does not apply to the Macintosh.

Window\_num    is a number from 1 to 3 that specifies how the window
containing the program should appear. Window\_num is only available for
use with Microsoft Excel for Windows. The window\_num argument is
allowed on the Macintosh, but it is ignored.

|                 |                    |
| --------------- | ------------------ |
| **Window\_num** | **Window appears** |
| 1               | Normal size        |
| 2 or omitted    | Minimized size     |
| 3               | Maximized size     |

Background    is a logical value that determines whether the program
specified by program\_text is opened as the active program or in the
background, leaving Microsoft Excel as the active program. If background
is TRUE, the program is started in the background; if FALSE or omitted,
the program is started in the foreground. Background is only available
for use with Microsoft Excel for the Macintosh and system software
version 7.0 or later.

Preferred\_size\_only    is a logical value that determines the amount
of memory allocated to the program. If preferred\_size\_only is TRUE,
the program is opened with its preferred memory allocation; if FALSE or
omitted, it opens with the available memory if greater than its minimum
requirement. Preferred\_size\_only is only available for use with
Microsoft Excel for the Macintosh and system software version 7.0 or
later. For information about changing the preferred memory size, see
your Macintosh documentation.

**Remarks**

In Microsoft Excel for Windows and in Microsoft Excel for the Macintosh
with system software version 7.0, if the EXEC function is successful, it
returns the task ID number of the started program. The task ID number is
a unique number that identifies a program. Use the task ID number in
other macro functions, such as APP.ACTIVATE, to refer to the program. In
Microsoft Excel for the Macintosh with system software version 6.0, if
EXEC is successful, it returns TRUE. If EXEC is unsuccessful, it returns
the \#VALUE\! error value.

**Examples**

In Microsoft Excel for Windows, the following macro formula starts the
program SEARCH.EXE. Use paths when the file or program to be started is
not in the current directory:

EXEC("C:\\WINDOWS\\SEARCH.EXE")

The following macro formula starts Microsoft Word for Windows and loads
the document SALES.DOC:

EXEC("C:\\WINWORD\\WINWORD.EXE C:\\MYFILES\\SALES.DOC")

In Microsoft Excel for the Macintosh, the following macro formula starts
Microsoft Word:

EXEC("HARD DISK:APPS:WORD")

**Related FunctioEXEC

Starts a separate program. Use EXEC to start other programs with which
you want to communicate. Use EXEC with Microsoft Excel's other DDE
functions (INITIATE, EXECUTE, and SEND.KEYS) to create a channel to
another program and to send keystrokes and commands to the program.
(SEND.KEYS is available only in Microsoft Excel for Windows.)

Syntax 1 is for Microsoft Excel for Windows. Syntax 2 is for Microsoft
Excel for the Macintosh.

**Syntax 1**

For Microsoft Excel for Windows

**EXEC**(program\_text, window\_num)

**Syntax 2**

For Microsoft Excel for the Macintosh

**EXEC**(program\_text, , background, preferred\_size\_only)

**Important**   Microsoft Excel for the Macintosh requires system
software version 7.0 or later for the last two arguments of this
function.

Program\_text    is the name, as a text string, of any executable file
or, in Microsoft Excel for Windows, any data file that is associated
with an executable file.

  - > Use paths when the file or program to be started is not in the
    > current directory or folder.

  - > In Microsoft Excel for Windows, program\_text can include any
    > arguments and switches that are accepted by the program to be
    > started. Also, if program\_text is the name of a file associated
    > with a specific installed program, EXEC starts the program and
    > loads the specified file.

>  

**Note   **In Microsoft Excel for the Macintosh, you must use an extra
comma after the program\_text argument. This skips the window\_num
argument that does not apply to the Macintosh.

Window\_num    is a number from 1 to 3 that specifies how the window
containing the program should appear. Window\_num is only available for
use with Microsoft Excel for Windows. The window\_num argument is
allowed on the Macintosh, but it is ignored.

|                 |                    |
| --------------- | ------------------ |
| **Window\_num** | **Window appears** |
| 1               | Normal size        |
| 2 or omitted    | Minimized size     |
| 3               | Maximized size     |

Background    is a logical value that determines whether the program
specified by program\_text is opened as the active program or in the
background, leaving Microsoft Excel as the active program. If background
is TRUE, the program is started in the background; if FALSE or omitted,
the program is started in the foreground. Background is only available
for use with Microsoft Excel for the Macintosh and system software
version 7.0 or later.

Preferred\_size\_only    is a logical value that determines the amount
of memory allocated to the program. If preferred\_size\_only is TRUE,
the program is opened with its preferred memory allocation; if FALSE or
omitted, it opens with the available memory if greater than its minimum
requirement. Preferred\_size\_only is only available for use with
Microsoft Excel for the Macintosh and system software version 7.0 or
later. For information about changing the preferred memory size, see
your Macintosh documentation.

**Remarks**

In Microsoft Excel for Windows and in Microsoft Excel for the Macintosh
with system software version 7.0, if the EXEC function is successful, it
returns the task ID number of the started program. The task ID number is
a unique number that identifies a program. Use the task ID number in
other macro functions, such as APP.ACTIVATE, to refer to the program. In
Microsoft Excel for the Macintosh with system software version 6.0, if
EXEC is successful, it returns TRUE. If EXEC is unsuccessful, it returns
the \#VALUE\! error value.

**Examples**

In Microsoft Excel for Windows, the following macro formula starts the
program SEARCH.EXE. Use paths when the file or program to be started is
not in the current directory:

EXEC("C:\\WINDOWS\\SEARCH.EXE")

The following macro formula starts Microsoft Word for Windows and loads
the document SALES.DOC:

EXEC("C:\\WINWORD\\WINWORD.EXE C:\\MYFILES\\SALES.DOC")

In Microsoft Excel for the Macintosh, the following macro formula starts
Microsoft Word:

EXEC("HARD DISK:APPS:WORD")

**Related Functions**

[APP.ACTIVATE](APP.ACTIVATE.md)   Switches to another application

[EXECUTE](EXECUTE.md)   Carries out a command in another application

[INITIATE](INITIATE.md)   Opens a channel to another application

[SEND.KEYS](SEND.KEYS.md)   Sends a key sequence to an application

[TERMINATE](TERMINATE.md)   Closes a channel to another application

[REQUEST](REQUEST.md)   Requests an array of a specific type of information from an
application with which you have a dynamic data exchange (DDE) link

[POKE](POKE.md)   Sends data to another application with which you have a dynamic
data exchange (DDE) link


