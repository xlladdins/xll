SQL.GET.SCHEMA
==============

Returns information about the structure of the data source on a
particular connection. The return value from a successful call to
SQL.GET.SCHEMA depends on the type of information that was requested. A
list of the accepted requests and their respective return values is
listed in the syntax section below.

If this function is not available, you must install the Microsoft ODBC
add-in (XLODBC.XLA).

**Syntax**

**SQL.GET.SCHEMA**(**connection\_num**, **type\_num**, qualifier\_text)

Connection\_num    is the unique connection ID of the data source you
want information about.

-   Connection\_num is returned by a previously executed SQL.OPEN
    > function.

-   If connection\_num is not valid, SQL.GET.SCHEMA returns the \#VALUE!
    > error value.

>  

Type\_num specifies the type of information you want returned. The
following is a list of valid type\_num values.

+-----------------+---------------------------------------------------+
| > **Type\_num** | > **Returns**                                     |
+-----------------+---------------------------------------------------+
| > 1             | > A list of available data sources, as a vertical |
|                 | > array.                                          |
+-----------------+---------------------------------------------------+
| > 2             | > A list of databases on the current connection,  |
|                 | > as a vertical array .                           |
+-----------------+---------------------------------------------------+
| > 3             | > A list of owners in a database on the current   |
|                 | > connection, as a vertical array.                |
+-----------------+---------------------------------------------------+
| > 4             | > A list of tables for a given owner and database |
|                 | > on the current connection, as a vertical array. |
+-----------------+---------------------------------------------------+
| > 5             | > A list of columns in a particular table and     |
|                 | > their data types, as a two-dimensional array.   |
|                 | > The returned array will have two fields and     |
|                 | > will have a row for each column in the table.   |
|                 | > The first field will be the name of the column. |
|                 | > The second field is the data type of the        |
|                 | > column. The data type will be a number that     |
|                 | > corresponds to the ODBC C header file data      |
|                 | > types. These \#define numbers are found in      |
|                 | > Microsoft Excel ODBC Developer Guide.           |
+-----------------+---------------------------------------------------+
| > 6             | > User ID of the current user                     |
+-----------------+---------------------------------------------------+
| > 7             | > Name of the current database.                   |
+-----------------+---------------------------------------------------+
| > 8             | > The name of the data source as given in the     |
|                 | > ODBC.INI file.                                  |
+-----------------+---------------------------------------------------+
| > 9             | > The name of the data source DBMS (i.e. Oracle,  |
|                 | > SQL Server, etc.).                              |
+-----------------+---------------------------------------------------+
| > 10            | > The server name for the data source.            |
+-----------------+---------------------------------------------------+
| > 11            | > The terminology used by the data source to      |
|                 | > refer to owners ( i.e. \"owner\",               |
|                 | > \"Authorization ID\", \"Schema\", etc.).        |
+-----------------+---------------------------------------------------+
| > 12            | > The terminology used by the data source to      |
|                 | > refer to tables ( i.e. \"table\", \"file\",     |
|                 | > etc.).                                          |
+-----------------+---------------------------------------------------+
| > 13            | > The terminology used by the data source to      |
|                 | > refer to qualifiers (i.e. \"database\" or       |
|                 | > \"directory\").                                 |
+-----------------+---------------------------------------------------+
| > 14            | > The terminology used by the data source to      |
|                 | > refer to procedures (i.e. \"database            |
|                 | > procedure\", \"stored procedure\", or           |
|                 | > \"procedure\").                                 |
+-----------------+---------------------------------------------------+

Qualifier\_text    is only included for type\_num values of 3, 4 and 5.
It is a text string used to qualify the search for the requested
information and should be enclosed by quotation marks.

+-----------------+---------------------------------------------------+
| > **Type\_num** | > **Qualifier\_text**                             |
+-----------------+---------------------------------------------------+
| > 3             | > The value of qualifier\_text should be the name |
|                 | > of a database in the current data source.       |
|                 | > SQL.GET.SCHEMA will then only return the names  |
|                 | > of table owners in that database.               |
+-----------------+---------------------------------------------------+
| > 4             | > The value of qualifier\_text should be both a   |
|                 | > database name and an owner name. The syntax of  |
|                 | > qualifier\_text is \"DatabaseName.OwnerName\".  |
|                 | > A period is used to separate the two names.     |
|                 | > SQL.GET.SCHEMA will then return an array of     |
|                 | > table names that are located in the given       |
|                 | > database and owned by the given owner.          |
+-----------------+---------------------------------------------------+
| > 5             | > The value of qualifier\_text should be the name |
|                 | > of a table. Information about the columns in    |
|                 | > that table will be returned.                    |
+-----------------+---------------------------------------------------+

**Remarks**

-   If SQL.GET.SCHEMA is unable to find the requested information then
    > it will return the error value \#N/A. In such a case
    > SQL.GET.SCHEMA will place error information in memory for the
    > SQL.ERROR function, if such information is available.

-   SQL.GET.SCHEMA works with the ODBC functions SQLGetInfo and
    > SQLTables to find the requested information. Refer to the
    > Microsoft Excel ODBC Programmer Guide for more information on
    > these two functions.

**Example**

SQL.GET.SCHEMA(conn1,7) returns the name of the current database.

SQL.GET.SCHEMA(conn1,9) returns the name of the DBMS used by the data
source.

**Related Functions**

SQL.OPEN   Establishes a connection with a data source

SQL.EXEC.QUERY   Sends a query to a data source

SQL.BIND   Specifies storage for a result column

SQL.RETRIEVE.TO.FILE   Retrieves query results and places them in a file

SQL.RETRIEVE   Retrieves query results

SQL.CLOSE   Closes a data source connection

SQL.ERROR   Returns detailed error information

Return to [top](#Q)

SQL.OPEN
