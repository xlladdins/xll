WINDOWS

Returns the names of the specified open Microsoft Excel windows,
including hidden windows. Use WINDOWS to get a list of active windows
for use by other macro functions that return information about or
manipulate windows, such as GET.WINDOW and ACTIVATE. The names are
returned as a horizontal array of text values, in order of their
appearance on your screen. The first name is the active window, the next
name is the window directly under the active window, and so on.

**Syntax**

**WINDOWS**(type\_num, match\_text)

Type\_num    is a number that specifies which types of workbooks are
returned by WINDOWS, according to the following table.

|               |                                                        |
| ------------- | ------------------------------------------------------ |
| **Type\_num** | **Returns window names from these types of documents** |
| 1 or omitted  | All windows except those belonging to add-in workbooks |
| 2             | Add-in workbooks only                                  |
| 3             | All types of workbooks                                 |

Match\_text    specifies the windows whose names you want returned and
can include wildcard characters. If match\_text is omitted, WINDOWS
returns the names of all open windows.

**Tips**

  - > You can change the output of a horizontal array to vertical with
    > the TRANSPOSE function.

  - > You can use WINDOWS with the INDEX function to select individual
    > window names from the array for use in other functions that take
    > window names as arguments.

  - > You can use the COLUMNS functions to count the number of entries
    > in the array, which is the number of windows.

>  

**Examples**

If the active window, named BOOK1, is on top of a window named MACROS:2,
which is on top of a window named MACROS:1, then:

WINDOWS() equals {"BOOK1", "MACROS:2", "MACROS:1"}

**Related Functions**

ACTIVATE   Switches to a window

DOCUMENTS   Returns the names of the specified open workbooks

GET.WINDOW   Returns information about a window

NEW.WINDOW   Creates a new window for an existing sheet or macro sheet

ON.WINDOW   Runs a macro when you switch to a window


