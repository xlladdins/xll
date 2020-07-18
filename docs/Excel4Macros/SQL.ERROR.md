SQL.ERROR

Returns detailed error information when it is called after a previous
XLODBC.XLA function call has failed. If this function is not available,
you must install the Microsoft ODBC add-in (XLODBC.XLA).

**Syntax**

**SQL.ERROR**()

Calling SQL.ERROR returns detailed error information in a two
dimensional array. Each row in the array describes exactly one error. If
a function call generates multiple errors, a row will be created for
each error. When SQL.ERROR is processed successfully, all SQL.ERROR
information is cleared. Also, all SQL.ERROR information is automatically
removed whenever an ODBC function completes successfully.

Each row will have exactly three fields. The information in these three
fields is obtained through the SQLERROR API function call. These fields
are:

  - > A textual message describing the error.

  - > The ODBC error class and subclass as a character string.

  - > The data source native error code as a numeric value.


If one or more of these fields is not available for the type of error
that was encountered, the field will be left blank. For more information
on the meaning of these three fields, refer to Chapter 24, "ODBC
Function Reference", in the Microsoft Open Database Connectivity
Programmer's Reference for the SQLError API function. See also Appendix
A, "ODBC Error Codes" in the same manual.

**Remarks**

  - > SQL.ERROR cannot provide information on Excel errors.

  - > If no error information is available when SQL.ERROR is called,
    > then it well return the error value \#N/A but does not post any
    > error information to SQL.ERROR.

  - > SQL.ERROR stores and returns error information by processing
    > SQL.ERROR (in the ODBC API reference) in a loop until
    > SQL\_NO\_DATA\_FOUND is encountered. In the SQL.ERROR function,
    > the error information is automatically defined and stored in
    > memory whenever an XLODBC.XLA function fails. If the call is
    > successful then SQL.ERROR will return the error information
    > available. If SQL.ERROR fails, it will return the error value
    > \#N/A.

**Example**

When entered as an array formula, the following example will return
error information about each argument in, for example, SQL.GET.QUERY.

SQL.ERROR()

**Related Functions**

[SQL.OPEN](SQL.OPEN.md)&nbsp;&nbsp;&nbsp;Establishes a connection with a data source

[SQL.EXEC.QUERY](SQL.EXEC.QUERY.md)&nbsp;&nbsp;&nbsp;Sends a query to a data source

[SQL.BIND](SQL.BIND.md)&nbsp;&nbsp;&nbsp;Specifies storage for a result column

[SQL.RETRIEVE.TO.FILE](SQL.RETRIEVE.TO.FILE.md)&nbsp;&nbsp;&nbsp;Retrieves query results and places
them in a file

[SQL.RETRIEVE](SQL.RETRIEVE.md)&nbsp;&nbsp;&nbsp;Retrieves query results

[SQL.CLOSE](SQL.CLOSE.md)&nbsp;&nbsp;&nbsp;Closes a data source connection



Return to [README](README.md)

