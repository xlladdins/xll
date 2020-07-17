WORKBOOK.PROTECT

Equivalent to clicking the Protect Workbook command on the Protection
submenu of the Tools menu. Controls protection of workbooks.

**Syntax**

**WORKBOOK.PROTECT**(structure, windows, password)

**WORKBOOK.PROTECT**?(structure, windows, password)

Structure    specifies whether the structure of the workbook is
protected. If TRUE, the structure is protected. If FALSE or omitted, the
structure is not protected.

Windows    specifies whether the windows of the workbook are protected.
If TRUE, the windows are protected. If FALSE or omitted, the windows are
not protected.

Password    specifies whether to protect the workbook with a password.
If omitted no password is used. When recording a macro, this argument is
not recorded. In the dialog box form of this function, you can specify a
password.

**Remarks**

To protect a sheet in a workbook, use the PROTECT.DOCUMENT function.

**Related Function**

[PROTECT.DOCUMENT](PROTECT.DOCUMENT.md)   Protects a sheet in a workbook



Return to [README](README.md)

