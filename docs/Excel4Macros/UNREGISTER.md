# UNREGISTER

Unregisters a previously registered dynamic link library (DLL) or code
resource. You can use UNREGISTER to free memory that was allocated to a
DLL or code resource when it was registered. There are two syntax forms
of this function. Use syntax 1 when you want Microsoft Excel to
unregister a function or code resource according to its use count. Use
syntax 2 when you want Microsoft Excel to unregister a function or code
resource regardless of the use count.

**Syntax 1**

**UNREGISTER**(**register\_id**)

Register\_id&nbsp;&nbsp;&nbsp;&nbsp;is the register ID returned by the
REGISTER or REGISTER.ID function, which corresponds to the function or
code resource to be removed from memory.

Microsoft Excel counts the number of times you register a function or
code resource. This number is called the use count. Each time you
unregister a function or code resource, its use count is decremented by
1. When the use count equals 0, Microsoft Excel frees the allocated
memory. Therefore, if you register a function or code resource more than
once, you must use a corresponding number of UNREGISTER functions to
ensure that it is completely unregistered.

**Note**&nbsp;&nbsp;&nbsp;**nbsp;&nbsp;&nbsp;&nbsp;**nbsp;&nbsp;&nbsp;&nbsp;**nbsp;Because Microsoft Excel for Windows and
Microsoft Excel for the Macintosh use different types of code resources,
UNREGISTER has a slightly different syntax form when used in each
operating environment.

**Syntax 2a**

For Microsoft Excel for Windows

**UNREGISTER**(**module\_text**)

**Syntax 2b**

For Microsoft Excel for the Macintosh

**UNREGISTER**(**file\_text**)

Module\_text or file\_text&nbsp;&nbsp;&nbsp;&nbsp;is text specifying the
name of the dynamic link library (DLL) that contains the function (in
Microsoft Excel for Windows) or the name of the file that contains the
code resource (in Microsoft Excel for the Macintosh).

If you use this syntax of UNREGISTER, all functions in the DLL (or all
code resources in the file) are immediately unregistered, regardless of
the use count.

**Examples**

Assuming that a REGISTER function in cell A5 of a macro sheet has
already run (and has run only once), the following macro formula
unregisters the corresponding function or code resource:

UNREGISTER(A5)

You could also use REGISTER.ID to return the register ID, instead of
specifying a cell reference:

UNREGISTER(REGISTER.ID("User", "GetTickCount")

Assuming that you have registered several different functions from the
USER.EXE DLL of Microsoft Windows, the following macro formula
unregisters all functions in that DLL:

UNREGISTER("User")

**Tip**&nbsp;&nbsp;&nbsp;**nbsp;&nbsp;&nbsp;&nbsp;**nbsp;&nbsp;&nbsp;&nbsp;**nbsp;If you register a function or code resource,
and use the optional function\_text argument to specify a custom name
that will appear in the Paste Function dialog box, this custom name will
not be removed by the UNREGISTER function. To remove the custom name,
use the SET.NAME function without its second argument.

**Related Function**

[REGISTER](REGISTER.md)&nbsp;&nbsp;&nbsp;Registers a code resource



Return to [README](README.md)

