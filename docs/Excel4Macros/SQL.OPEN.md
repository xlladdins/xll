SQL.OPEN

Establishes a connection with a data source. If the connection is
successfully established SQL.OPEN will return a connection ID number.
Use the connection ID number with other ODBC macro functions to identify
a connection.

If this function is not available, you must install the Microsoft ODBC
add-in (XLODBC.XLA).

**Syntax**

**SQL.OPEN**(**connection\_string**, output\_ref, driver\_prompt)

Connection\_string    is a text string that contains the information
necessary to establish a connection to a data source. Any
data-source-name that is used in connection\_string must be an existing
data source name defined with ODBC Setup or the ODBC Administration
Utility.

  - > Connection\_string must follow the format described in Chapter 24,
    > "ODBC Function Reference", of the Microsoft Open Database
    > Connectivity Programmer's Reference for SQLDriverConnect. In this
    > string the user supplies the data source name, one or more user
    > ID's, one or more passwords, and any other information necessary
    > to successfully connect to a DBMS. An example of a SQL.OPEN
    > connection\_string entered would be: "DSN=MyServer; UID=dbayer;
    > PWD=123; Database=pubs"

  - > Enter the connection\_string as an array when the length exceeds
    > 255 characters. Or enter connection\_string as an array of cells
    > containing the same information. The connection string should be
    > horizontal array.

Output\_ref    is a single cell reference where you want the completed
connection string to be placed. Use output\_ref when you want SQL.OPEN
to return the completed connection string. If output\_ref is omitted, a
completed connection string will not be returned.

Driver\_prompt     is a number from 1 to 4 specifying if and how you
want to be prompted by the driver. This sets the fDriverCompletion flag
in ODBC's SQLDriverConnect.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Number</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Description</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Always brings up a dialog box.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Bring up dialog only if there is not enough information to connect. The driver uses information from the connection string and from the data source specification as defaults.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>Same as 2, but the driver grays and disables any prompts for information not needed.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>If the connection string is unsuccessful, do not bring up a dialog box.</p>
</blockquote></td>
</tr>
</tbody>
</table>

**Remarks**

  - > If SQL.OPEN is unable to connect with the information provided
    > then it will return the error value \#N/A. In such a case,
    > SQL.OPEN will place error information in memory for the SQL.ERROR
    > function, if more information is available.

  - > If the call is successful then SQL.OPEN will return a unique
    > connection ID number that can be used in future function calls to
    > identify the connection.

  - > If connection\_array does not allow SQL.OPEN to connect to a data
    > source, then the error value \#N/A will be returned.

**Example**

conn1=SQL.OPEN('DSN=NWind;DBQ=C:\\MSQUERY;FIL=dBASE4',C15, 2) sets the
name conn1 to the return value of SQL.OPEN, which connects to the NWind
data source, specifies where to place the connection string, and
displays the driver dialog box only if additional information is needed.

**Related Functions**

[SQL.EXEC.QUERY](SQL.EXEC.QUERY.md)   Sends a query to a data source

[SQL.BIND](SQL.BIND.md)   Specifies storage for a result column

[SQL.RETRIEVE.TO.FILE](SQL.RETRIEVE.TO.FILE.md)   Retrieves query results and places them in a file

[SQL.RETRIEVE](SQL.RETRIEVE.md)   Retrieves query results

[SQL.GET.SCHEMA](SQL.GET.SCHEMA.md)   Gets information about a connected data source.

[SQL.CLOSE](SQL.CLOSE.md)   Closes a data source connection

[SQL.ERROR](SQL.ERROR.md)   Returns detailed error information



Return to [README](README.md)

