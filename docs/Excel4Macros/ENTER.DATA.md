# ENTER.DATA

Turns on Data Entry mode and allows you to select and to enter data into
the unlocked cells in the current selection only (the data entry area).
Use ENTER.DATA when you want to enter data only in a specific part of
your sheet. You can then use that part of the sheet as a simple data
form.

**Syntax**

**ENTER.DATA**(logical)

Logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that turns Data Entry
mode on or off.

  - > If logical is TRUE, Data Entry mode is turned on; if FALSE, Data
    > Entry mode is turned off and data entry, cell movement, and cell
    > selection return to normal. If logical is omitted, ENTER.DATA
    > toggles Data Entry mode.

  - > Logical can also be the number 2. This setting turns on Data Entry
    > mode and prevents the ESC key from turning it off.

  - > Logical can also be a reference. Using a reference for this
    > argument turns on Data Entry mode for the supplied reference.


**Remarks**

  - > In Data Entry mode, you can move the active cell and select cell
    > ranges only in the data entry area. The arrow keys and the TAB and
    > SHIFT+TAB keys move from one unlocked cell to the next. The HOME
    > and END keys move to the first and last cell in the data entry
    > area, respectively. You cannot select entire rows or columns, and
    > clicking a cell outside the data entry area does not select it.

  - > The only commands available in Data Entry mode are commands
    > normally available to protected workbooks.

  - > To turn off Data Entry mode, press ESC (unless logical is 2),
    > activate another sheet in the active workbook window, or use
    > another ENTER.DATA function. If you use another ENTER.DATA
    > function, you will usually design your macros in one of two ways:
    
      - > The macro turns on Data Entry mode, pauses while you enter
        > data, resumes, and then turns off Data Entry mode.
    
      - > The macro turns on Data Entry mode and ends. After entering
        > data, another macro turns off Data Entry mode; this latter
        > macro could be assigned to a "Finished" button, for example.

> With either method, you can use Microsoft Excel's ON functions to
> resume or run other macros based on an event, such as pressing the
> CONTROL+D keys.
> 

**Tips**

  - > Normally you use Data Entry mode to enter data, but you can also
    > prevent someone from entering data or moving the active cell by
    > locking all the cells in the current selection before turning on
    > Data Entry mode. This is useful if you want a user to view a range
    > of cells but not change it or move the active cell. Similarly, if
    > you unlock certain cells, you can restrict the user's movement to
    > the Data Entry area only.

  - > To prevent someone from activating another workbook, which would
    > turn off Data Entry mode, use the ON.WINDOW function or an
    > Auto\_Deactivate macro.

**Related Functions**

[DISABLE.INPUT](DISABLE.INPUT.md)&nbsp;&nbsp;&nbsp;Blocks all input to Microsoft Excel

[FORMULA](FORMULA.md)&nbsp;&nbsp;&nbsp;Enters values into a cell or range or onto a
chart



Return to [README](README.md#E)

# ENTER.DATA

Turns on Data Entry mode and allows you to select and to enter data into
the unlocked cells in the current selection only (the data entry area).
Use ENTER.DATA when you want to enter data only in a specific part of
your sheet. You can then use that part of the sheet as a simple data
form.

**Syntax**

**ENTER.DATA**(logical)

Logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that turns Data Entry
mode on or off.

  - > If logical is TRUE, Data Entry mode is turned on; if FALSE, Data
    > Entry mode is turned off and data entry, cell movement, and cell
    > selection return to normal. If logical is omitted, ENTER.DATA
    > toggles Data Entry mode.

  - > Logical can also be the number 2. This setting turns on Data Entry
    > mode and prevents the ESC key from turning it off.

  - > Logical can also be a reference. Using a reference for this
    > argument turns on Data Entry mode for the supplied reference.


**Remarks**

  - > In Data Entry mode, you can move the active cell and select cell
    > ranges only in the data entry area. The arrow keys and the TAB and
    > SHIFT+TAB keys move from one unlocked cell to the next. The HOME
    > and END keys move to the first and last cell in the data entry
    > area, respectively. You cannot select entire rows or columns, and
    > clicking a cell outside the data entry area does not select it.

  - > The only commands available in Data Entry mode are commands
    > normally available to protected workbooks.

  - > To turn off Data Entry mode, press ESC (unless logical is 2),
    > activate another sheet in the active workbook window, or use
    > another ENTER.DATA function. If you use another ENTER.DATA
    > function, you will usually design your macros in one of two ways:
    
      - > The macro turns on Data Entry mode, pauses while you enter
        > data, resumes, and then turns off Data Entry mode.
    
      - > The macro turns on Data Entry mode and ends. After entering
        > data, another macro turns off Data Entry mode; this latter
        > macro could be assigned to a "Finished" button, for example.

> With either method, you can use Microsoft Excel's ON functions to
> resume or run other macros based on an event, such as pressing the
> CONTROL+D keys.
> 

**Tips**

  - > Normally you use Data Entry mode to enter data, but you can also
    > prevent someone from entering data or moving the active cell by
    > locking all the cells in the current selection before turning on
    > Data Entry mode. This is useful if you want a user to view a range
    > of cells but not change it or move the active cell. Similarly, if
    > you unlock certain cells, you can restrict the user's movement to
    > the Data Entry area only.

  - > To prevent someone from activating another workbook, which would
    > turn off Data Entry mode, use the ON.WINDOW function or an
    > Auto\_Deactivate macro.

**Related Functions**

[DISABLE.INPUT](DISABLE.INPUT.md)&nbsp;&nbsp;&nbsp;Blocks all input to Microsoft Excel

[FORMULA](FORMULA.md)&nbsp;&nbsp;&nbsp;Enters values into a cell or range or onto a
chart



Return to [README](README.md#E)

# ENTER.DATA

Turns on Data Entry mode and allows you to select and to enter data into
the unlocked cells in the current selection only (the data entry area).
Use ENTER.DATA when you want to enter data only in a specific part of
your sheet. You can then use that part of the sheet as a simple data
form.

**Syntax**

**ENTER.DATA**(logical)

Logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that turns Data Entry
mode on or off.

  - > If logical is TRUE, Data Entry mode is turned on; if FALSE, Data
    > Entry mode is turned off and data entry, cell movement, and cell
    > selection return to normal. If logical is omitted, ENTER.DATA
    > toggles Data Entry mode.

  - > Logical can also be the number 2. This setting turns on Data Entry
    > mode and prevents the ESC key from turning it off.

  - > Logical can also be a reference. Using a reference for this
    > argument turns on Data Entry mode for the supplied reference.


**Remarks**

  - > In Data Entry mode, you can move the active cell and select cell
    > ranges only in the data entry area. The arrow keys and the TAB and
    > SHIFT+TAB keys move from one unlocked cell to the next. The HOME
    > and END keys move to the first and last cell in the data entry
    > area, respectively. You cannot select entire rows or columns, and
    > clicking a cell outside the data entry area does not select it.

  - > The only commands available in Data Entry mode are commands
    > normally available to protected workbooks.

  - > To turn off Data Entry mode, press ESC (unless logical is 2),
    > activate another sheet in the active workbook window, or use
    > another ENTER.DATA function. If you use another ENTER.DATA
    > function, you will usually design your macros in one of two ways:
    
      - > The macro turns on Data Entry mode, pauses while you enter
        > data, resumes, and then turns off Data Entry mode.
    
      - > The macro turns on Data Entry mode and ends. After entering
        > data, another macro turns off Data Entry mode; this latter
        > macro could be assigned to a "Finished" button, for example.

> With either method, you can use Microsoft Excel's ON functions to
> resume or run other macros based on an event, such as pressing the
> CONTROL+D keys.
> 

**Tips**

  - > Normally you use Data Entry mode to enter data, but you can also
    > prevent someone from entering data or moving the active cell by
    > locking all the cells in the current selection before turning on
    > Data Entry mode. This is useful if you want a user to view a range
    > of cells but not change it or move the active cell. Similarly, if
    > you unlock certain cells, you can restrict the user's movement to
    > the Data Entry area only.

  - > To prevent someone from activating another workbook, which would
    > turn off Data Entry mode, use the ON.WINDOW function or an
    > Auto\_Deactivate macro.

**Related Functions**

[DISABLE.INPUT](DISABLE.INPUT.md)&nbsp;&nbsp;&nbsp;Blocks all input to Microsoft Excel

[FORMULA](FORMULA.md)&nbsp;&nbsp;&nbsp;Enters values into a cell or range or onto a
chart



Return to [README](README.md#E)

