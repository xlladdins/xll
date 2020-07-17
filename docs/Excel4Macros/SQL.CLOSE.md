SQL.CLOSE

Terminates a connection to an external data source. If this function is
not available, you must install the Microsoft ODBC add-in (XLODBC.XLA).

**Syntax**

**SQL.CLOSE**(**connection\_num**)

Connection\_num    is the unique connection ID of the data source from
which you wish to disconnect.

  - > Connection\_num is returned by a previously executed SQL.OPEN
    > function.

  - > If connection\_num is not valid, SQL.CLOSE returns the \#VALUE\!
    > error value.

>  

**Remarks**

  - > If the connection is successfully terminated SQL.CLOSE will return
    > zero and the connection ID number is then no longer valid.

  - > If SQL.CLOSE is unable to disconnect with the data source then it
    > will return the error value the \#N/A error value. In such a case
    > SQL.CLOSE will place error information in memory for the SQL.ERROR
    > function, if such information is available.

  - > SQL.CLOSE works with data sources in much the same manner as
    > FCLOSE works with files. If the call is successful then SQL.CLOSE
    > will terminate the specified data source connection.

**Example**

SQL.CLOSE(conn1) will close the connection with connection\_num conn1

**Related Functions**

[SQL.OPEN](SQL.OPEN.md)   Establishes a connection with a data source

[SQL.EXEC.QUERY](SQL.EXEC.QUERY.md)   Sends a query to a data source

[SQL.BIND](SQL.BIND.md)   Specifies storage for a result column

[SQL.RETRIEVE.TO.FILE](SQL.RETRIEVE.TO.FILE.md)   Retrieves query results and places them in a file

[SQL.RETRIEVE](SQL.RETRIEVE.md)   Retrieves query results

[SQL.GET.SCHEMA](SQL.GET.SCHEMA.md)   Gets information about a connected data source.

[SQL.ERROR](SQL.ERROR.md)   Returns detailed error information


