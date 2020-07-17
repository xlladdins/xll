SQL.EXEC.QUERY
==============

Sends a query to a data source using an existing connection. If this
function is not available, you must install the Microsoft ODBC add-in
(XLODBC.XLA).

**Syntax**

**SQL.EXEC.QUERY**(**connection\_num**, **query\_text**)

Connection\_num    is the unique connection ID of the data source you
want to query.

-   Connection\_num is returned by a previously executed SQL.OPEN
    > function.

-   If connection\_num is not valid, SQL.EXEC.QUERY returns the \#VALUE!
    > error value.

Query\_text is the SQL language query that is to be executed on the data
source. The query must follow the SQL syntax guidelines in the Appendix
of the Microsoft Excel ODBC Developers Guide.

-   If SQL.EXEC.QUERY is unable to execute query\_text on the specified
    > data source, SQL.EXEC.QUERY returns the \#N/A error value.

-   Excel limits strings to a length of 255 characters. If query\_text
    > needs to be longer than 255 characters then query\_text should be
    > a vertical array or vertical range of cells. The values in the
    > array will be joined together to form the complete SQL query.

>  

**Remarks**

-   Before calling SQL.EXEC.QUERY a connection must be established with
    > a data source using SQL.OPEN. A successful call to SQL.OPEN
    > returns a unique connection ID number. SQL.EXEC.QUERY uses that
    > connection ID number to send SQL language queries to the data
    > source.

-   Any results generated from the query will not be returned
    > immediately\-- SQL.EXEC.QUERY only executes the query. Retrieving
    > results is handled by the functions SQL.RETRIEVE and
    > SQL.RETRIEVE.TO.FILE.

-   If SQL.EXEC.QUERY is called using a previously used connection ID
    > number, all pending results on that connection will automatically
    > be discarded. The connection ID will then refer to the new query
    > and its results.

-   If SQL.EXEC.QUERY is unable to successfully execute the query on the
    > specified data source then an error value will be returned. In
    > such a case SQL.EXEC.QUERY will place error information in memory
    > for the SQL.ERROR function, if such information is available. If
    > SQL.EXEC.QUERY is able to successfully execute the query on the
    > specified connection it will return one of three values depending
    > on the type of SQL statement that was executed.

    -   If it was a SELECT statement then SQL.EXEC.QUERY will return the
        > number of result columns available.

    -   If it was an UPDATE, INSERT, or DELETE statement then
        > SQL.EXEC.QUERY will return the number of rows affected by the
        > statement.

    -   If it was a legal SQL query that is not in one of the categories
        > above, SQL.EXEC.QUERY will return 0 (zero).

**Example**

SQL.EXEC.QUERY(conn1, \"SELECT Custmr\_ID, Due\_Date FROM Orders WHERE
Order\_Amt \> 100\") executes a SQL query from a SQL table named
\"Orders\"

**Related Functions**

SQL.OPEN   Establishes a connection with a data source

SQL.BIND   Specifies storage for a result column

SQL.RETRIEVE.TO.FILE   Retrieves query results and places them in a file

SQL.RETRIEVE   Retrieves query results

SQL.GET.SCHEMA   Gets information about a connected data source.

SQL.CLOSE   Closes a data source connection

SQL.ERROR   Returns detailed error information

Return to [top](#Q)

SQL.GET.SCHEMA
