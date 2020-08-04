# SQL.BIND

Specifies where results from a SQL query are placed when they are
retrieved with SQL.RETRIEVE. If this function is not available, you must
install the Microsoft ODBC add-in (XLODBC.XLA).

**Syntax**

**SQL.BIND**(**connection\_num**, column, reference)

Connection\_num&nbsp;&nbsp;&nbsp;&nbsp;is the unique connection ID of
the data source for which you want to define storage.

  - > Connection\_num was returned by a previously executed SQL.OPEN
    > function.

  - > If connection\_num is not valid, then SQL.BIND returns the
    > \#VALUE\! error value.


Column&nbsp;&nbsp;&nbsp;&nbsp;is the number of the result column that is
to be bound. Result columns are numbered from left to right starting
with 1. If column is omitted then all bindings for connection\_num are
removed. Column number 0 contains row numbers for the result set. If
column number 0 is bound then SQL.RETRIEVE will return row numbers in
the bound location.

Reference&nbsp;&nbsp;&nbsp;&nbsp; is a single cell reference on the
worksheet where the results should be placed. If reference is omitted,
the binding is removed for the column.

When SQL.RETRIEVE is called, the result rows in this column will be
placed in the reference cell and the cells immediately below reference.
The number of rows that will be retrieved is one of the SQL.RETRIEVE
arguments.

**Remarks**

  - > If SQL.BIND is completed successfully then it will return a
    > vertical array listing the bound columns on the current
    > connection. If SQL.BIND is unable to bind the result column then
    > it will return the error value \#N/A. In such a case SQL.BIND will
    > place error information in memory for the SQL.ERROR function, if
    > such information is available.

  - > SQL.BIND tells the ODBC interface where to place results when they
    > are retrieved using SQL.RETRIEVE. Binding is not necessary but can
    > be useful if you want the results from different columns to be
    > placed in disjoint worksheet locations.

  - > If bindings are used, SQL.BIND must be called once for each column
    > in the result set. If a result column is not bound then it will
    > not be returned. A binding remains valid for as long as
    > connection\_num is open.

  - > Call SQL.BIND after calling SQL.OPEN and SQL.EXEC.QUERY, but
    > before calling SQL.RETRIEVE or SQL.RETRIEVE.TO.FILE. Calls to
    > SQL.BIND will not affect results that have already been retrieved.

**Example**

SQL.BIND(conn1,1,"\[Book1\]Sheet1\!C1") stores data obtained from the
data source conn1 on Sheet1 from left to right in cell C1, starting with
column1.

**Related Functions**

[SQL.OPEN](SQL.OPEN.md)&nbsp;&nbsp;&nbsp;Establishes a connection with a data source

[SQL.EXEC.QUERY](SQL.EXEC.QUERY.md)&nbsp;&nbsp;&nbsp;Sends a query to a data source

[SQL.RETRIEVE.TO.FILE](SQL.RETRIEVE.TO.FILE.md)&nbsp;&nbsp;&nbsp;Retrieves query results and places
them in a file

[SQL.RETRIEVE](SQL.RETRIEVE.md)&nbsp;&nbsp;&nbsp;Retrieves query results

[SQL.GET.SCHEMA](SQL.GET.SCHEMA.md)&nbsp;&nbsp;&nbsp;Gets information about a connected data
source.

[SQL.CLOSE](SQL.CLOSE.md)&nbsp;&nbsp;&nbsp;Closes a connection to a data source.

[SQL.ERROR](SQL.ERROR.md)&nbsp;&nbsp;&nbsp;Returns detailed error information



Return to [README](README.md#S)

