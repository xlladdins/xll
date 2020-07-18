# REGISTER

Registers the specified dynamic link library (DLL) or code resource and
returns the register ID. You can also specify a custom function name and
argument names that will appear in the Paste Function dialog box. If you
register a command (macro\_type = 2), you can also specify a shortcut
key. Because Microsoft Excel for Windows and Microsoft Excel for the
Macintosh use different types of code resources, REGISTER has a slightly
different syntax form when used in each operating environment.

**Important**&nbsp;&nbsp; This function is provided for advanced users
only. If you use the CALL function incorrectly, you could cause errors
that will require you to restart your computer.

**Syntax 1**

For Microsoft Excel for Windows

**REGISTER**(**module\_text**, procedure, type\_text, function\_text,
argument\_text, macro\_type, category, shortcut\_text, help\_topic,
function\_help, argument\_help1, argument\_help2,...)

**Syntax 2**

For Microsoft Excel for the Macintosh

**REGISTER**(**file\_text**, resource, type\_text, function\_text,
argument\_text, macro\_type, category, shortcut\_text, help\_topic,
function\_help, argument\_help1, argument\_help2,...)

Module\_text or file\_text&nbsp;&nbsp;&nbsp;&nbsp;is text specifying the
name of the DLL that contains the function (in Microsoft Excel for
Windows) or the name of the file that contains the code resource (in
Microsoft Excel for the Macintosh).

Procedure or resource&nbsp;&nbsp;&nbsp;&nbsp;is text specifying the name
of the function in the DLL (in Microsoft Excel for Windows) or the name
of the code resource (in Microsoft Excel for the Macintosh). In
Microsoft Excel for Windows, you can also use the ordinal value of the
function from the EXPORTS statement in the module-definition file
(.DEF). In Microsoft Excel for the Macintosh, you can also use the
resource ID number. The ordinal value or resource ID number should not
be in text form.

This argument may be omitted for stand-alone DLLs or code resources. In
this case, REGISTER will register all functions or code resources and
then return module\_text or file\_text.

Type\_text&nbsp;&nbsp;&nbsp;&nbsp;is text specifying the data type of
the return value and the data types of all arguments to the DLL or code
resource. The first letter of type\_text specifies the return value.

Function\_text&nbsp;&nbsp;&nbsp;&nbsp;is text specifying the name of the
function as you want it to appear in the Paste Function dialog box. If
you omit this argument, the function will not appear in the Paste
Function dialog box.

Argument\_text&nbsp;&nbsp;&nbsp;&nbsp;is text specifying the names of
the arguments you want to appear in the Paste Function dialog box.
Argument names should be separated by commas.

Macro\_type&nbsp;&nbsp;&nbsp;&nbsp;specifies the macro type: 1 for a
function or 2 for a command. If macro\_type is omitted, it is assumed to
be 1 (function).

Category&nbsp;&nbsp;&nbsp;&nbsp;specifies the function category in the
Paste Function dialog box in which you want the registered function to
appear. You can use the category number or the category name for
category. If you use the category name, be sure to enclose it in double
quotation marks. If category is omitted, it is assumed to be 14 (User
Defined).

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Category number</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Category name</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Financial</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Date &amp; Time</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>Math &amp; Trig</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>Text</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>Logical</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>6</p>
</blockquote></td>
<td><blockquote>
<p>Lookup &amp; Matrix</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>7</p>
</blockquote></td>
<td><blockquote>
<p>Database</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>Statistical</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>9</p>
</blockquote></td>
<td><blockquote>
<p>Information</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>10</p>
</blockquote></td>
<td><blockquote>
<p>Commands (macro sheets only)</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>11</p>
</blockquote></td>
<td><blockquote>
<p>Actions (macro sheets only)</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>12</p>
</blockquote></td>
<td><blockquote>
<p>Customizing (macro sheets only)</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>13</p>
</blockquote></td>
<td><blockquote>
<p>Macro Control (macro sheets only)</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>14</p>
</blockquote></td>
<td><blockquote>
<p>User Defined</p>
</blockquote></td>
</tr>
</tbody>
</table>

Shortcut\_text&nbsp;&nbsp;&nbsp;&nbsp;is a character specifying the
shortcut key for the registered command. The shortcut key is
case-sensitive. This argument is used only if macro\_type = 2 (command).
If shortcut\_text is omitted, the command will not have a shortcut key.

Help\_topic&nbsp;&nbsp;&nbsp;&nbsp;is the reference (including path) to
the help file that you want displayed when the user clicks the Help
button when your custom function is displayed.

Function\_help&nbsp;&nbsp;&nbsp;&nbsp;is a text string describing your
custom function when it is selected in the Paste Function dialog box.
The maximum number of characters is 255.

Argument\_help1, argument\_help2&nbsp;&nbsp;&nbsp;&nbsp;are 1 to 21 text
strings that describes you custom function's arguments when the function
is selected in the Paste Function dialog box.

Example&nbsp;&nbsp;&nbsp;&nbsp;  
Syntax 1

In Microsoft Excel for Windows, the following macro formula registers
the GetTickCount function from Microsoft Windows. This function returns
the number of milliseconds that have elapsed since Microsoft Windows was
started.

REGISTER("User", "GetTickCount", "J")

Assuming that the REGISTER function is in cell A5, after your macro
registers GetTickCount, you can use the CALL function to return the
number of milliseconds that have elapsed:

CALL(A5)

Example&nbsp;&nbsp;&nbsp;&nbsp;  
Syntax 1 with optional function\_text

You can use the following macro formula to register the GetTickCount
function from Microsoft Windows and assign the custom name GetTicks to
it. To do this, include "GetTicks" as the optional function\_text
argument to the REGISTER function.

REGISTER("User", "GetTickCount", "J", "GetTicks", , 1, 9)

After the function is registered, the custom name GetTicks will appear
in the Information function category (category = 9) in the Paste
Function dialog box.

You can call the function from the same macro sheet on which it was
registered using the following formula:

GetTicks()

You can call the function from another sheet or macro sheet by including
the name of the original macro sheet in the formula. For example,
assuming the macro sheet on which GetTicks was registered is named
MACRO1.XLS, the following formula calls the function from another sheet:

MACRO1.XLS\!GetTicks()

**Tip**&nbsp;&nbsp;&nbsp;**nbsp;&nbsp;&nbsp;&nbsp;**nbsp;&nbsp;&nbsp;&nbsp;**nbsp;You can use functions in a DLL or code resource
directly on a sheet without first registering them from a macro sheet.
Use syntax 2a or 2b of the CALL function. For more information, see
CALL.

**Related Function**

[UNREGISTER](UNREGISTER.md)&nbsp;&nbsp;&nbsp;Removes a registered code resource from
memory



Return to [README](README.md)

