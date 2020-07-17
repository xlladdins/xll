QUERY.GET.DATA

Builds a new query using the supplied information. The application
Microsoft Query nor any dialog boxes are displayed.

**Syntax**

**QUERY.GET.DATA**(**connection\_string**, **query\_text**,
keep\_query\_def, field\_names, row\_numbers, destination)

**QUERY.GET.DATA**?(connection\_string, query\_text, keep\_query\_def,
field\_names, row\_numbers, destination)

Connection\_string    supplies information, such as the data source
name, user ID, and passwords, necessary to making a SQL connection to an
external data source. For example: "DSN=Myserver; Server=server1;
UID=dbayer; PWD=buyer1; Database=nwind".

 

  - > You must define the data source name (DSN) used in
    > connection\_string before you try to connect to it.

  - > You can enter connection\_string as an array or a string. If
    > connection\_string exceeds 250 characters, you must enter it as an
    > array.

  - > If QUERY.GET.DATA is unable to access the data source using
    > connection\_string, it returns the \#N/A error value.

>  

Query\_text    is the SQL language query to be executed on the data
source.

Keep\_query\_def    is a logical value that, if TRUE or omitted,
preserves the query definition. If FALSE, the query definition is lost
and the data from the query no longer constitutes a data range.

Field\_names    is a logical value that, if TRUE or omitted, places
field names from Microsoft Query into the first row of the data range.
If FALSE, the field names are discarded.

Row\_numbers    is a logical value that, if TRUE, places row numbers
from Microsoft Query into the first column in the data range. If FALSE
or omitted, the row numbers are discarded.

Destination    is the location as a cell reference where you want the
data placed. If destination is in a current data range then that data
range is changed to reflect the new SQL statement. The default
destination is the currently selected cell or range.

**Remarks**

  - > If the information provided is not sufficient to create the query
    > then the error value \#REF\! is returned.

  - > If Microsoft Query is unavailable or can not be found, \#N/A is
    > returned.

  - > If connection string is longer than 255 characters, the string
    > will be truncated at the last semi-colon.

**Related Function**

QUERY.REFRESH   Refreshes the data in a data range returned by Microsoft
Query


QUERY.GET.DATA

Builds a new query using the supplied information. The application
Microsoft Query nor any dialog boxes are displayed.

**Syntax**

**QUERY.GET.DATA**(**connection\_string**, **query\_text**,
keep\_query\_def, field\_names, row\_numbers, destination)

**QUERY.GET.DATA**?(connection\_string, query\_text, keep\_query\_def,
field\_names, row\_numbers, destination)

Connection\_string    supplies information, such as the data source
name, user ID, and passwords, necessary to making a SQL connection to an
external data source. For example: "DSN=Myserver; Server=server1;
UID=dbayer; PWD=buyer1; Database=nwind".

 

  - > You must define the data source name (DSN) used in
    > connection\_string before you try to connect to it.

  - > You can enter connection\_string as an array or a string. If
    > connection\_string exceeds 250 characters, you must enter it as an
    > array.

  - > If QUERY.GET.DATA is unable to access the data source using
    > connection\_string, it returns the \#N/A error value.

>  

Query\_text    is the SQL language query to be executed on the data
source.

Keep\_query\_def    is a logical value that, if TRUE or omitted,
preserves the query definition. If FALSE, the query definition is lost
and the data from the query no longer constitutes a data range.

Field\_names    is a logical value that, if TRUE or omitted, places
field names from Microsoft Query into the first row of the data range.
If FALSE, the field names are discarded.

Row\_numbers    is a logical value that, if TRUE, places row numbers
from Microsoft Query into the first column in the data range. If FALSE
or omitted, the row numbers are discarded.

Destination    is the location as a cell reference where you want the
data placed. If destination is in a current data range then that data
range is changed to reflect the new SQL statement. The default
destination is the currently selected cell or range.

**Remarks**

  - > If the information provided is not sufficient to create the query
    > then the error value \#REF\! is returned.

  - > If Microsoft Query is unavailable or can not be found, \#N/A is
    > returned.

  - > If connection string is longer than 255 characters, the string
    > will be truncated at the last semi-colon.

**Related Function**

QUERY.REFRESH   Refreshes the data in a data range returned by Microsoft
Query


QUERY.GET.DATA

Builds a new query using the supplied information. The application
Microsoft Query nor any dialog boxes are displayed.

**Syntax**

**QUERY.GET.DATA**(**connection\_string**, **query\_text**,
keep\_query\_def, field\_names, row\_numbers, destination)

**QUERY.GET.DATA**?(connection\_string, query\_text, keep\_query\_def,
field\_names, row\_numbers, destination)

Connection\_string    supplies information, such as the data source
name, user ID, and passwords, necessary to making a SQL connection to an
external data source. For example: "DSN=Myserver; Server=server1;
UID=dbayer; PWD=buyer1; Database=nwind".

 

  - > You must define the data source name (DSN) used in
    > connection\_string before you try to connect to it.

  - > You can enter connection\_string as an array or a string. If
    > connection\_string exceeds 250 characters, you must enter it as an
    > array.

  - > If QUERY.GET.DATA is unable to access the data source using
    > connection\_string, it returns the \#N/A error value.

>  

Query\_text    is the SQL language query to be executed on the data
source.

Keep\_query\_def    is a logical value that, if TRUE or omitted,
preserves the query definition. If FALSE, the query definition is lost
and the data from the query no longer constitutes a data range.

Field\_names    is a logical value that, if TRUE or omitted, places
field names from Microsoft Query into the first row of the data range.
If FALSE, the field names are discarded.

Row\_numbers    is a logical value that, if TRUE, places row numbers
from Microsoft Query into the first column in the data range. If FALSE
or omitted, the row numbers are discarded.

Destination    is the location as a cell reference where you want the
data placed. If destination is in a current data range then that data
range is changed to reflect the new SQL statement. The default
destination is the currently selected cell or range.

**Remarks**

  - > If the information provided is not sufficient to create the query
    > then the error value \#REF\! is returned.

  - > If Microsoft Query is unavailable or can not be found, \#N/A is
    > returned.

  - > If connection string is longer than 255 characters, the string
    > will be truncated at the last semi-colon.

**Related Function**

QUERY.REFRESH   Refreshes the data in a data range returned by Microsoft
Query


QUERY.GET.DATA

Builds a new query using the supplied information. The application
Microsoft Query nor any dialog boxes are displayed.

**Syntax**

**QUERY.GET.DATA**(**connection\_string**, **query\_text**,
keep\_query\_def, field\_names, row\_numbers, destination)

**QUERY.GET.DATA**?(connection\_string, query\_text, keep\_query\_def,
field\_names, row\_numbers, destination)

Connection\_string    supplies information, such as the data source
name, user ID, and passwords, necessary to making a SQL connection to an
external data source. For example: "DSN=Myserver; Server=server1;
UID=dbayer; PWD=buyer1; Database=nwind".

 

  - > You must define the data source name (DSN) used in
    > connection\_string before you try to connect to it.

  - > You can enter connection\_string as an array or a string. If
    > connection\_string exceeds 250 characters, you must enter it as an
    > array.

  - > If QUERY.GET.DATA is unable to access the data source using
    > connection\_string, it returns the \#N/A error value.

>  

Query\_text    is the SQL language query to be executed on the data
source.

Keep\_query\_def    is a logical value that, if TRUE or omitted,
preserves the query definition. If FALSE, the query definition is lost
and the data from the query no longer constitutes a data range.

Field\_names    is a logical value that, if TRUE or omitted, places
field names from Microsoft Query into the first row of the data range.
If FALSE, the field names are discarded.

Row\_numbers    is a logical value that, if TRUE, places row numbers
from Microsoft Query into the first column in the data range. If FALSE
or omitted, the row numbers are discarded.

Destination    is the location as a cell reference where you want the
data placed. If destination is in a current data range then that data
range is changed to reflect the new SQL statement. The default
destination is the currently selected cell or range.

**Remarks**

  - > If the information provided is not sufficient to create the query
    > then the error value \#REF\! is returned.

  - > If Microsoft Query is unavailable or can not be found, \#N/A is
    > returned.

  - > If connection string is longer than 255 characters, the string
    > will be truncated at the last semi-colon.

**Related Function**

QUERY.REFRESH   Refreshes the data in a data range returned by Microsoft
Query


QUERY.GET.DATA

Builds a new query using the supplied information. The application
Microsoft Query nor any dialog boxes are displayed.

**Syntax**

**QUERY.GET.DATA**(**connection\_string**, **query\_text**,
keep\_query\_def, field\_names, row\_numbers, destination)

**QUERY.GET.DATA**?(connection\_string, query\_text, keep\_query\_def,
field\_names, row\_numbers, destination)

Connection\_string    supplies information, such as the data source
name, user ID, and passwords, necessary to making a SQL connection to an
external data source. For example: "DSN=Myserver; Server=server1;
UID=dbayer; PWD=buyer1; Database=nwind".

 

  - > You must define the data source name (DSN) used in
    > connection\_string before you try to connect to it.

  - > You can enter connection\_string as an array or a string. If
    > connection\_string exceeds 250 characters, you must enter it as an
    > array.

  - > If QUERY.GET.DATA is unable to access the data source using
    > connection\_string, it returns the \#N/A error value.

>  

Query\_text    is the SQL language query to be executed on the data
source.

Keep\_query\_def    is a logical value that, if TRUE or omitted,
preserves the query definition. If FALSE, the query definition is lost
and the data from the query no longer constitutes a data range.

Field\_names    is a logical value that, if TRUE or omitted, places
field names from Microsoft Query into the first row of the data range.
If FALSE, the field names are discarded.

Row\_numbers    is a logical value that, if TRUE, places row numbers
from Microsoft Query into the first column in the data range. If FALSE
or omitted, the row numbers are discarded.

Destination    is the location as a cell reference where you want the
data placed. If destination is in a current data range then that data
range is changed to reflect the new SQL statement. The default
destination is the currently selected cell or range.

**Remarks**

  - > If the information provided is not sufficient to create the query
    > then the error value \#REF\! is returned.

  - > If Microsoft Query is unavailable or can not be found, \#N/A is
    > returned.

  - > If connection string is longer than 255 characters, the string
    > will be truncated at the last semi-colon.

**Related Function**
**Related Function**

QUERY.REFRESH   Refreshes the data in a data range returned by Microsoft
Query


QUERY.GET.DATA

Builds a new query using the supplied information. The application
Microsoft Query nor any dialog boxes are displayed.

**Syntax**

**QUERY.GET.DATA**(**connection\_string**, **query\_text**,
keep\_query\_def, field\_names, row\_numbers, destination)

**QUERY.GET.DATA**?(connection\_string, query\_text, keep\_query\_def,
field\_names, row\_numbers, destination)

Connection\_string    supplies information, such as the data source
name, user ID, and passwords, necessary to making a SQL connection to an
external data source. For example: "DSN=Myserver; Server=server1;
UID=dbayer; PWD=buyer1; Database=nwind".

 

  - > You must define the data source name (DSN) used in
    > connection\_string before you try to connect to it.

  - > You can enter connection\_string as an array or a string. If
    > connection\_string exceeds 250 characters, you must enter it as an
    > array.

  - > If QUERY.GET.DATA is unable to access the data source using
    > connection\_string, it returns the \#N/A error value.

>  

Query\_text    is the SQL language query to be executed on the data
source.

Keep\_query\_def    is a logical value that, if TRUE or omitted,
preserves the query definition. If FALSE, the query definition is lost
and the data from the query no longer constitutes a data range.

Field\_names    is a logical value that, if TRUE or omitted, places
field names from Microsoft Query into the first row of the data range.
If FALSE, the field names are discarded.

Row\_numbers    is a logical value that, if TRUE, places row numbers
from Microsoft Query into the first column in the data range. If FALSE
or omitted, the row numbers are discarded.

Destination    is the location as a cell reference where you want the
data placed. If destination is in a current data range then that data
range is changed to reflect the new SQL statement. The default
destination is the currently selected cell or range.

**Remarks**

  - > If the information provided is not sufficient to create the query
    > then the error value \#REF\! is returned.

  - > If Microsoft Query is unavailable or can not be found, \#N/A is
    > returned.

  - > If connection string is longer than 255 characters, the string
    > will be truncated at the last semi-colon.

**Related Function**

QUERY.REFRESH   Refreshes the data in a data range returned by Microsoft
Query


QUERY.GET.DATA

Builds a new query using the supplied information. The application
Microsoft Query nor any dialog boxes are displayed.

**Syntax**

**QUERY.GET.DATA**(**connection\_string**, **query\_text**,
keep\_query\_def, field\_names, row\_numbers, destination)

**QUERY.GET.DATA**?(connection\_string, query\_text, keep\_query\_def,
field\_names, row\_numbers, destination)

Connection\_string    supplies information, such as the data source
name, user ID, and passwords, necessary to making a SQL connection to an
external data source. For example: "DSN=Myserver; Server=server1;
UID=dbayer; PWD=buyer1; Database=nwind".

 

  - > You must define the data source name (DSN) used in
    > connection\_string before you try to connect to it.

  - > You can enter connection\_string as an array or a string. If
    > connection\_string exceeds 250 characters, you must enter it as an
    > array.

  - > If QUERY.GET.DATA is unable to access the data source using
    > connection\_string, it returns the \#N/A error value.

>  

Query\_text    is the SQL language query to be executed on the data
source.

Keep\_query\_def    is a logical value that, if TRUE or omitted,
preserves the query definition. If FALSE, the query definition is lost
and the data from the query no longer constitutes a data range.

Field\_names    is a logical value that, if TRUE or omitted, places
field names from Microsoft Query into the first row of the data range.
If FALSE, the field names are discarded.

Row\_numbers    is a logical value that, if TRUE, places row numbers
from Microsoft Query into the first column in the data range. If FALSE
or omitted, the row numbers are discarded.

Destination    is the location as a cell reference where you want the
data placed. If destination is in a current data range then that data
range is changed to reflect the new SQL statement. The default
destination is the currently selected cell or range.

**Remarks**

  - > If the information provided is not sufficient to create the query
    > then the error value \#REF\! is returned.

  - > If Microsoft Query is unavailable or can not be found, \#N/A is
    > returned.

  - > If connection string is longer than 255 characters, the string
    > will be truncated at the last semi-colon.

**Related Function**

QUERY.REFRESH   Refreshes the data in a data range returned by Microsoft
Query


QUERY.GET.DATA

Builds a new query using the supplied information. The application
Microsoft Query nor any dialog boxes are displayed.

**Syntax**

**QUERY.GET.DATA**(**connection\_string**, **query\_text**,
keep\_query\_def, field\_names, row\_numbers, destination)

**QUERY.GET.DATA**?(connection\_string, query\_text, keep\_query\_def,
field\_names, row\_numbers, destination)

Connection\_string    supplies information, such as the data source
name, user ID, and passwords, necessary to making a SQL connection to an
external data source. For example: "DSN=Myserver; Server=server1;
UID=dbayer; PWD=buyer1; Database=nwind".

 

  - > You must define the data source name (DSN) used in
    > connection\_string before you try to connect to it.

  - > You can enter connection\_string as an array or a string. If
    > connection\_string exceeds 250 characters, you must enter it as an
    > array.

  - > If QUERY.GET.DATA is unable to access the data source using
    > connection\_string, it returns the \#N/A error value.

>  

Query\_text    is the SQL language query to be executed on the data
source.

Keep\_query\_def    is a logical value that, if TRUE or omitted,
preserves the query definition. If FALSE, the query definition is lost
and the data from the query no longer constitutes a data range.

Field\_names    is a logical value that, if TRUE or omitted, places
field names from Microsoft Query into the first row of the data range.
If FALSE, the field names are discarded.

Row\_numbers    is a logical value that, if TRUE, places row numbers
from Microsoft Query into the first column in the data range. If FALSE
or omitted, the row numbers are discarded.

Destination    is the location as a cell reference where you want the
data placed. If destination is in a current data range then that data
range is changed to reflect the new SQL statement. The default
destination is the currently selected cell or range.

**Remarks**

  - > If the information provided is not sufficient to create the query
    > then the error value \#REF\! is returned.

  - > If Microsoft Query is unavailable or can not be found, \#N/A is
    > returned.

  - > If connection string is longer than 255 characters, the string
    > will be truncated at the last semi-colon.

**Related Function**

QUERY.REFRESH   Refreshes the data in a data range returned by Microsoft
Query


QUERY.GET.DATA

Builds a new query using the supplied information. The application
Microsoft Query nor any dialog boxes are displayed.

**Syntax**

**QUERY.GET.DATA**(**connection\_string**, **query\_text**,
keep\_query\_def, field\_names, row\_numbers, destination)

**QUERY.GET.DATA**?(connection\_string, query\_text, keep\_query\_def,
field\_names, row\_numbers, destination)

Connection\_string    supplies information, such as the data source
name, user ID, and passwords, necessary to making a SQL connection to an
external data source. For example: "DSN=Myserver; Server=server1;
UID=dbayer; PWD=buyer1; Database=nwind".

 

  - > You must define the data source name (DSN) used in
    > connection\_string before you try to connect to it.

  - > You can enter connection\_string as an array or a string. If
    > connection\_string exceeds 250 characters, you must enter it as an
    > array.

  - > If QUERY.GET.DATA is unable to access the data source using
    > connection\_string, it returns the \#N/A error value.

>  

Query\_text    is the SQL language query to be executed on the data
source.

Keep\_query\_def    is a logical value that, if TRUE or omitted,
preserves the query definition. If FALSE, the query definition is lost
and the data from the query no longer constitutes a data range.

Field\_names    is a logical value that, if TRUE or omitted, places
field names from Microsoft Query into the first row of the data range.
If FALSE, the field names are discarded.

Row\_numbers    is a logical value that, if TRUE, places row numbers
from Microsoft Query into the first column in the data range. If FALSE
or omitted, the row numbers are discarded.

Destination    is the location as a cell reference where you want the
data placed. If destination is in a current data range then that data
range is changed to reflect the new SQL statement. The default
destination is the currently selected cell or range.

**Remarks**

  - > If the information provided is not sufficient to create the query
    > then the error value \#REF\! is returned.

  - > If Microsoft Query is unavailable or can not be found, \#N/A is
    > returned.

  - > If connection string is longer than 255 characters, the string
    > will be truncated at the last semi-colon.

**Related Function**

[QUERY.REFRESH](QUERY.REFRESH.md)   Refreshes the data in a data range returned by Microsoft
[Q](Q.md)uery


