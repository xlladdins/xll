# ADDIN.MANAGER

Equivalent to clicking the Add-Ins command on the Tools menu. Adds or
removes an installed add-in from the working set. The add-in file must
already be installed.

**Syntax**

**ADDIN.MANAGER**(operation\_num, addinname\_text, copy\_logical)

**ADDIN.MANAGER**?(operation\_num, addinname\_text, copy\_logical)

Operation\_num&nbsp;&nbsp;&nbsp;&nbsp;determines the operation that the
add-in manager will perform.

|                    |                                                                                                                                                                       |
| ------------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Operation\_num** | **Operation**                                                                                                                                                         |
| 1                  | Adds an add-in to the working set, using the descriptive name in the Add-Ins dialog box.                                                                              |
| 2                  | Removes an add-in from the working set, using the descriptive name in the Add-Ins dialog box.                                                                         |
| 3                  | Adds a new add-in to the list of add-ins that Microsoft Excel knows about. Equivalent to clicking on the Browse button in the Add-Ins dialog box and clicking a file. |

Addinname\_text&nbsp;&nbsp;&nbsp;&nbsp;specifies the name of the add-in.
If operation\_num is 1 or 2, use the descriptive name of the add-in,
such as "SOLVER". If operation\_num is 3, this contains the filename of
the add-in.

Copy\_logical&nbsp;&nbsp;&nbsp;&nbsp;specifies whether the add-in should
be copied to the library directory. This argument is only used if
operation\_num is 3. If omitted, and the file is on removable media, the
user will be asked if they want to copy it to removable media.



Return to [README](README.md#A)

