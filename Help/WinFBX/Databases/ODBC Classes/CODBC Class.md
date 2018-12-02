# ODBC Classes

**CODBC** is a wrapper class on top of ODBC. 

The main purpose of the class if to ease the use of the ODBC API, that is notoriously difficult to use.

**Include files**: CODBC.inc, CODBCDb.inc, CODBCStmt.inc.

You create an instance of the class using:

```
' // Create a connection object and connect with the database
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr
```

You create an statement object creating an instance of the CODBCStmt object passing the connection object pointer as the parameter:

```
' // Allocate an statement object
DIM pStmt AS COdbcStmt = pDbc
```

When the class is destroyed, its Destructor method closes the cursor and the connection and frees the allocated connection handle. Therefore, it is not needed to explicitly close the database and free handles. The class keeps track of the number of active connections and frees the environment handle if there are no active connections.

#### Example

```
#include once "Afx/COdbc/COdbc.inc"
USING Afx

' // Create a connection object and connect with the database
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr

' // Allocate an statement object
DIM pStmt AS COdbcStmt = pDbc

' // Generate a result set
pStmt.ExecDirect ("SELECT * FROM Authors ORDER BY Author")

' // Parse the result set
DIM cwsOutput AS CWSTR
DO
   ' // Fetch the record
   IF pStmt.Fetch = FALSE THEN EXIT DO
   ' // Get the values of the columns and display them
   cwsOutput = ""
   cwsOutput += pStmt.GetData(1) & " "
   cwsOutput += pStmt.GetData(2) & " "
   cwsOutput += pStmt.GetData(3)
   PRINT cwsOutput
LOOP
```

# CODBCBase Class

Base class for all the ODBC classes. Implements some common methods that all the other interfaces inherit. You don't have to instantiate this class.

**Include file**: CODBC.inc

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [Error](#Error) | Returns TRUE if there has been an error; FALSE, otherwise. |
| [GetCPMatch](#GetCPMatch) | Returns a 32-bit SQLUINTEGER value that determines how a connection is chosen from a connection pool. |
| [GetDataSources](#GetDataSources) | Lists available DSN / Drivers installed. |
| [GetDrivers](#GetDrivers) | Lists driver descriptions and driver attribute keywords. |
| [GetEnvAttr](#GetEnvAttr) | Returns the current setting of an environment attribute. |
| [GetErrorInfo](#GetErrorInfo) | Returns a verbose description of the last error(s). |
| [GetLastResult](#GetLastResult) | Returns the last result code. |
| [GetOutputNTS](#GetOutputNTS) | Returns a 32-bit integer that determines how the driver returns string data. |
| [GetSqlState](#GetSqlState) | Returns the SqlState for the specified handle. |
| [ODBCVersion](#ODBCVersion) | Returns a 32-bit integer that determines whether certain functionality exhibits ODBC 2.x behavior or ODBC 3.x behavior. |
| [SetCPMatch](#SetCPMatch) | Returns a 32-bit SQLUINTEGER value that determines how a connection is chosen from a connection pool. |
| [SetEnvAttr](#SetEnvAttr) | Sets attributes that govern aspects of environments. |
| [SetErrorProc](#SetErrorProc) | Sets the address of an application defined error callback. |
| [SetOutputNTS](#SetOutputNTS) | Returns a 32-bit integer that determines how the driver returns string data. |

# CODBC Class

Implements methods to create and manage connection objects. Inherits from COdbcBase.

**Include file**: CODBCDbc.inc

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorsDb) | Requests a commit operation for all active operations on all statements associated with an environment. |
| [CommitTran](#CommitTran) | Requests a commit operation for all active operations on all statements associated with an environment. |
| [EnvHandle](#EnvHandle) | Returns the environment handle. |
| [Functions](#Functions) | Returns information about whether a driver supports a specific ODBC function. |
| [GetConnectAttr](#GetConnectAttr) | Returns the current setting of a connection attribute. |
| [GetDiagField](#GetDiagField) | Returns the current value of a field of a record of the diagnostic data structure (associated with an environment handle) that contains error, warning, and status information. |
| [GetDiagRec](#GetDiagRec) | Returns the current values of multiple fields of a diagnostic record that contains error, warning, and status information. |
| [GetErrorInfo](#GetErrorInfo) | Returns a verbose description of the last error. |
| [GetInfo](#GetInfo) | Returns general information about the driver and data source associated with a connection. |
| [GetSqlState](#GetSqlState) | Returns the SqlState for the connection handle. |
| [Handle](#Handle) | Returns the connection handle. |
| [NativeSql](#NativeSql) | Returns the SQL string as modified by the driver. NativeSql does not execute the SQL statement. |
| [RollbackTran](#RollbackTran) | RollbackTran requests a rollback operation for all active operations on all statements associated with an environment. |
| [SetConnectAttr](#SetConnectAttr) | Sets attributes that govern aspects of connections. |
| [Supports](#Supports) | Returns information about whether a driver supports a specific ODBC function. It is an alias for **Functions**. |

# <a name="Error"></a>Error (CODBCBase)

Returns TRUE if there has been an error; FALSE, otherwise.

```
FUNCTION Error () AS BOOLEAN
```

#### Return value

Returns TRUE if the last result code is SQL_ERROR or SQL_INVALID_HANDLE.

# <a name="GetCPMatch"></a>GetCPMatch (CODBCBase)

Returns a 32-bit SQLUINTEGER value that determines how a connection is chosen from a connection pool.

```
FUNCTION GetCPMatch () AS SQLUINTEGER
```

#### Return value

The current value of the attribute.

#### Result code (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Remarks

The Driver Manager determines which connection is reused from the pool and attempts to match the connection options in the call and the connection attributes set by the application to the keywords and connection attributes of the connections in the pool. The value of this attribute determines the level of precision of the matching criteria. The following values are used to set the value of this attribute:

**SQL_CP_STRICT_MATCH**

Only connections that exactly match the connection options in the call and the connection attributes set by the application are reused. This is the default.

**SQL_CP_RELAXED_MATCH**

Connections with matching connection string keywords can be used. Keywords must match, but not all connection attributes must  match.

# <a name="GetDataSources"></a>GetDataSources (CODBCBase)

Lists available DSN / Drivers installed.

```
FUNCTION GetDataSources (BYVAL Direction AS SQLUSMALLINT, BYREF cwsServerName AS CWSTR, _
   BYREF cwsDescription AS CWSTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *Direction* | Determines which data source the Driver Manager returns information on. Can be:<br>SQL_FETCH_NEXT (to fetch the next data source name in the list), SQL_FETCH_FIRST (to fetch from the beginning of the list), SQL_FETCH_FIRST_USER (to fetch the first user DSN), or SQL_FETCH_FIRST_SYSTEM (to fetch the first system DSN).<br>When *Direction* is set to SQL_FETCH_FIRST, subsequent calls to **GetDataSources** with *Direction* set to SQL_FETCH_NEXT return both user and system DSNs. When *Direction* is set to SQL_FETCH_FIRST_USER, all subsequent calls to **GetDataSources** with *Direction* set to SQL_FETCH_NEXT return only user DSNs. When *Direction* is set to SQL_FETCH_FIRST_SYSTEM, all subsequent calls to **GetDataSources** with *Direction* set to SQL_FETCH_NEXT return only system DSNs. |
| *cwsServerName* | An string variable in which to return the data source name. |
| *cwsDescription* | An string variable in which to return the description. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Diagnostics

When **GetDataSources** returns either SQL_ERROR or SQL_SUCCESS_WITH_INFO, an associated SQLSTATE value can be obtained by calling the **SqlState** property.

# <a name="GetDrivers"></a>GetDrivers (CODBCBase)

Lists driver descriptions and driver attribute keywords. This function is implemented only by the Driver Manager.

```
FUNCTION GetDrivers (BYVAL Direction AS SQLUSMALLINT, BYREF cwsDriverDesc AS CWSTR, _
   BYREF cwsDriverAttributes AS CWSTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *Direction* | Determines whether the Driver Manager fetches the next driver description in the list (SQL_FETCH_NEXT) or whether the search starts from the beginning of the list (SQL_FETCH_FIRST). |
| *cwsDriverDesc* | A CWSTR variable in which to return the driver description. |
| *cwsDriverAttributes* | A CWSTR variable in which to return the list of driver attribute value pairs (see "Comments"). |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Diagnostics

When **GetDrivers** returns either SQL_ERROR or SQL_SUCCESS_WITH_INFO, an associated SQLSTATE value can be obtained by calling the **SqlState** property.

#### Comments

**GetDrivers** returns the driver description in the *cwsDriverDesc* variable. It returns additional information about the driver in the *cwsDriverAttributes* buffer as a list of keyword-value pairs. All keywords listed in the system information for drivers will be returned for all drivers, except for **CreateDSN**, which is used to prompt creation of data sources and therefore is optional. Each pair is terminated with a semicolon.

Driver attribute keywords are added from the system information when the driver is installed.

An application can call GetDrivers multiple times to retrieve all driver descriptions. The Driver Manager retrieves this information from the system information. When there are no more driver descriptions, **GetDrivers** returns SQL_NO_DATA. If **GetDrivers** is called with SQL_FETCH_NEXT immediately after it returns SQL_NO_DATA, it returns the first driver description.

If SQL_FETCH_NEXT is passed to **GetDrivers** the very first time it is called, **GetDrivers** returns the first data source name.

Because **GetDrivers** is implemented in the Driver Manager, it is supported for all drivers regardless of a particular driver's standards compliance.

#### Example

```
#include once "Afx/COdbc/COdbc.inc"
USING Afx

' // Create a connection object and connect with the database
' // We need to call it to create the environment handle
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr
IF pDbc.Handle = NULL THEN PRINT "Unable to create the connection handle" : SLEEP : END

DIM wDirection AS WORD
DIM cwsDriverAttributes AS CWSTR
DIM cwsDriverDesc AS CWSTR

wDirection = SQL_FETCH_FIRST
DO
   cwsDriverDesc = ""
   cwsDriverAttributes = ""
   IF SQL_SUCCEEDED(pDbc.GetDrivers(wDirection, cwsDriverDesc, cwsDriverAttributes)) = 0 THEN EXIT DO
   ? "Driver description: " & cwsDriverDesc
   ? "Driver attributes: " & cwsDriverAttributes
   wDirection = SQL_FETCH_NEXT
LOOP
```

# <a name="GetEnvAttr"></a>GetEnvAttr (CODBCBase)

Returns the current setting of an environment attribute.

```
FUNCTION GetEnvAttr (BYVAL Attribute AS SQLINTEGER, BYVAL ValuePtr AS SQLPOINTER, _
   BYVAL BufferLength AS SQLINTEGER, BYVAL StringLength AS SQLINTEGER PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *Attribute* | Attribute to retrieve. |
| *ValuePtr* |  Pointer to a buffer in which to return the current value of the attribute specified by **Attribute**. |
| *BufferLength* | If *ValuePtr* points to a character string, this argument should be the length of *ValuePtr*. If *ValuePtr* is an integer, *BufferLength* is ignored. If the attribute value is not a character string, *BufferLength* is unused. |
| *StringLength* | A pointer to a buffer in which to return the total number of bytes (excluding the null-termination character) available to return in *ValuePtr*. If *ValuePtr* is a null pointer, no length is returned. If the attribute value is a character string and the number of bytes available to return is greater than or equal to *BufferLength*, the data in *ValuePtr* is truncated to *BufferLength* minus the length of a null-termination character and is null-terminated by the driver. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetErrorInfo"></a>GetErrorInfo (CODBCBase)

Returns a verbose description of the last error(s).

```
FUNCTION GetErrorInfo (BYVAL HandleType AS SQLSMALLINT, BYVAL Handle AS SQLHANDLE, _
   BYVAL iErrorCode AS SQLRETURN = 0) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *HandleType* | The handle type:<br>SQL_HANDLE_DBC (connection handle)<br>SQL_HANDLE_STMT (statement handle) |
| *Handle* | The handle value. |
| *iErrorCode* | Optional. The error code returned by **GetLastResult**. |

#### Return value

The description of the error or errors.

# <a name="GetLastResult"></a>GetLastResult (CODBCBase)

Returns the last result code.

```
FUNCTION GetLastResult () AS SQLRETURN
```

#### Return value

The result code returned by the last executed ODBC method.

GetDiagRec or GetDiagField returns SQLSTATE values as defined by X/Open Data Management: Structured Query Language (SQL), Version 2 (March 1995). SQLSTATE values are strings that contain five characters. The following table lists SQLSTATE values that a driver can return for **GetDiagRec**.

The character string value returned for an SQLSTATE consists of a two-character class value followed by a three-character subclass value. A class value of "01" indicates a warning and is accompanied by a return code of SQL_SUCCESS_WITH_INFO. Class values other than "01," except for the class "IM," indicate an error and are accompanied by a return value of SQL_ERROR. The class "IM" is specific to warnings and errors that derive from the implementation of ODBC itself. The subclass value "000" in any class indicates that there is no subclass for that SQLSTATE. The assignment of class and subclass values is defined by SQL-92.

Note Although successful execution of a function is normally indicated by a return value of SQL_SUCCESS, the SQLSTATE 00000 also indicates success.

# <a name="GetOutputNTS"></a>GetOutputNTS (CODBCBase)

Returns a 32-bit integer that determines how the driver returns string data. If SQL_TRUE, the driver returns string data null-terminated. If SQL_FALSE, the driver does not return string data null-terminated. This attribute defaults to SQL_TRUE. A call to SetEnvAttr to set it to SQL_TRUE returns SQL_SUCCESS. A call to SetEnvAttr to set it to SQL_FALSE returns SQL_ERROR and SQLSTATE HYC00.

**Note**: Optional feature not implemented in Microsoft Access Driver.

```
FUNCTION GetOutputNTS () AS SQLUINTEGER
```

#### Return value

The current value of the attribute.

#### Result code (GetLastStatus)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_ERROR, or SQL_INVALID_HANDLE.


# <a name="GetSqlState"></a>GetSqlState (CODBCBase)

Returns the SqlState for the specified handle.

**Note**: Optional feature not implemented in Microsoft Access Driver.

```
FUNCTION GetSqlState (BYVAL HandleType AS SQLSMALLINT, BYVAL Handle AS SQLHANDLE) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *HandleType* | The handle type:<br>SQL_HANDLE_DBC (connection handle)<br>SQL_HANDLE_STMT (statement handle) |
| *Handle* | The handle value. |

#### Result code (GetLastStatus)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Return value

The SqlState.

#### ODBC Error Codes

**GetDiagRec** or **GetDiagField** returns SQLSTATE values as defined by X/Open Data Management: Structured Query Language (SQL), Version 2 (March 1995). SQLSTATE values are strings that contain five characters. The following table lists SQLSTATE values that a driver can return for GetDiagRec.

The character string value returned for an SQLSTATE consists of a two-character class value followed by a three-character subclass value. A class value of "01" indicates a warning and is accompanied by a return code of SQL_SUCCESS_WITH_INFO. Class values other than "01," except for the class "IM," indicate an error and are accompanied by a return value of SQL_ERROR. The class "IM" is specific to warnings and errors that derive from the implementation of ODBC itself. The subclass value "000" in any class indicates that there is no subclass for that SQLSTATE. The assignment of class and subclass values is defined by SQL-92.

**Note**: Although successful execution of a function is normally indicated by a return value of SQL_SUCCESS, the SQLSTATE 00000 also indicates success.

| SQLSTATE | Error |
| -------- | ----- |
| 01000 | General warning |
| 01001 | Cursor operation conflict |
| 01002 | Disconnect error |
| 01003 | NULL value eliminated in set function |
| 01004 | String data, right-truncated |
| 01006 | Privilege not revoked |
| 01007 | Privilege not granted |
| 01S00 | Invalid connection string attribute |
| 01S01 | Error in row |
| 01S02 | Option value changed |
| 01S06 | Attempt to fetch before the result set returned the first rowset |
| 01S08 | Error saving File DSN |
| 01S09 | Invalid keyword |
| 07001 | Wrong number of parameters |
| 07002 | COUNT field incorrect |
| 07005 | Prepared statement not a cursor-specification |
| 07006 | Restricted data type attribute violation |
| 07009 | Invalid descriptor index |
| 07S01 | Invalid use of default parameter |
| 08001 | Client unable to establish connection |
| 08002 | Connection name in use |
| 08003 | Connection does not exist |
| 08004 | Server rejected the connection |
| 08007 | Connection failure during transaction |
| 08S01 | Communication link failure |
| 21S01 | Insert value list does not match column list |
| 21S02 | Degree of derived table does not match column list |
| 22001 | String data, right-truncated |
| 22002 | Indicator variable required but not supplied |
| 22003 | Numeric value out of range |
| 22007 | Invalid datetime format |
| 22008 | Datetime field overflow |
| 22012 | Division by zero |
| 22015 | Interval field overflow |
| 22018 | Invalid character value for cast specification |
| 22019 | Invalid escape character |
| 22025 | Invalid escape sequence |
| 22026 | String data, length mismatch |
| 23000 | Integrity constraint violation |
| 24000 | Invalid cursor state |
| 25000 | Invalid transaction state |
| 25S01 | Transaction state |
| 25S02 | Transaction is still active |
| 25S03 | Transaction is rolled back |
| 28000 | Invalid authorization specification |
| 34000 | Invalid cursor name |
| 3C000 | Duplicate cursor name |
| 3D000 | Invalid catalog name |
| 3F000 | Invalid schema name |
| 40001 | Serialization failure |
| 40002 | Integrity constraint violation |
| 40003 | Statement completion unknown |
| 42000 | Syntax error or access violation |
| 42S01 | Base table or view already exists |
| 42S02 | Base table or view not found |
| 42S11 | Index already exists |
| 42S12 | Index not found |
| 42S21 | Column already exists |
| 42S22 | Column not found |
| 44000 | WITH CHECK OPTION violation |
| HY000 | General error |
| HY001 | Memory allocation error |
| HY003 | Invalid application buffer type |
| HY004 | Invalid SQL data type |
| HY007 | Associated statement is not prepared |
| HY008 | Operation canceled |
| HY009 | Invalid use of null pointer |
| HY010 | Function sequence error |
| HY011 | Attribute cannot be set now |
| HY012 | Invalid transaction operation code |
| HY013 | Memory management error |
| HY014 | Limit on the number of handles exceeded |
| HY015 | No cursor name available |
| HY016 | Cannot modify an implementation row descriptor |
| HY017 | Invalid use of an automatically allocated descriptor handle |
| HY018 | Server declined cancel request |
| HY019 | Non-character and non-binary data sent in pieces |
| HY020 | Attempt to concatenate a null value |
| HY021 | Inconsistent descriptor information |
| HY024 | Invalid attribute value |
| HY090 | Invalid string or buffer length |
| HY091 | Invalid descriptor field identifier |
| HY092 | Invalid attribute/option identifier |
| HY095 | Function type out of range |
| HY096 | Invalid information type |
| HY097 | Column type out of range |
| HY098 | Scope type out of range |
| HY099 | Nullable type out of range |
| HY100 | Uniqueness option type out of range |
| HY101 | Accuracy option type out of range |
| HY103 | Invalid retrieval code |
| HY104 | Invalid precision or scale value |
| HY105 | Invalid parameter type |
| HY106 | Fetch type out of range |
| HY107 | Row value out of range |
| HY109 | Invalid cursor position |
| HY110 | Invalid driver completion |
| HY111 | Invalid bookmark value |
| HYC00 | Optional feature not implemented |
| HYT00 | Timeout expired |
| HYT01 | Connection timeout expired |
| IM001 | Driver does not support this function |
| IM002 | Data source name not found and no default driver specified |
| IM003 | Specified driver could not be loaded |
| IM004 | Driver's SQLAllocHandle on SQL_HANDLE_ENV failed |
| IM005 | Driver's SQLAllocHandle on SQL_HANDLE_DBC failed |
| IM006 | Driver's SQLSetConnectAttr failed |
| IM007 | No data source or driver specified; dialog prohibited |
| IM008 | Dialog failed |
| IM009 | Unable to load translation DLL |
| IM010 | Data source name too long |
| IM011 | Driver name too long |
| IM012 | DRIVER keyword syntax error |
| IM013 | Trace file error |
| IM014 | Invalid name of File DSN |
| IM015 | Corrupt file data source |


# <a name="ODBCVersion"></a>ODBCVersion (CODBCBase)

Returns a 32-bit integer that determines whether certain functionality exhibits ODBC 2.x behavior or ODBC 3.x behavior.

```
FUNCTION ODBCVersion () AS SQLUINTEGER
```

#### Return value

* **SQL_OV_ODBC3_80** = The Driver Manager and driver exhibit the following ODBC 3.8 behavior:

   The driver returns and expects ODBC 3.x codes for date, time, and timestamp.
   The driver returns ODBC 3.x SQLSTATE codes when Error, GetDiagField, or GetDiagRec is called.
   The CatalogName argument in a call to SQLTables accepts a search pattern.
   The Driver Manager supports C data type extensibility. For more information about C data type extensibility, see C Data Types in ODBC.

* **SQL_OV_ODBC3** = The Driver Manager and driver exhibit the following ODBC 3.x behavior:

   The driver returns and expects ODBC 3.x codes for date, time, and timestamp.
   The driver returns ODBC 3.x SQLSTATE codes when Error,  GetDiagField, or GetDiagRec is called.
   The CatalogName argument in a call to SqlTables accepts a search pattern.

* **SQL_OV_ODBC2** = The Driver Manager and driver exhibit the following ODBC 2.x behavior. This is especially useful for an ODBC 2.x application working with an ODBC 3.x driver.

   The driver returns and expects ODBC 2.x codes for date, time, and timestamp.
   The driver returns ODBC 2.x SQLSTATE codes when Error, GetDiagField, or GetDiagRec is called.
   The CatalogName argument in a call to SqlTables does not accept a search pattern.

An application must set this environment attribute before calling any function that has an SQLHENV argument, or the call will return SQLSTATE HY010 (Function sequence error). It is driver-specific whether or not additional behaviors exist for these environmental flags.

#### Remarks

To set the ODBC version, use the optional parameters of the **CODBC** class constructors.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetCPMatch"></a>SetCPMatch (CODBCBase)

Sets a 32-bit SQLUINTEGER value that determines how a connection is chosen from a connection pool.

```
FUNCTION SetCPMatch (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwAttr* | Value of the attribute. |

#### Remarks

The Driver Manager determines which connection is reused from the pool and attempts to match the connection options in the call and the connection attributes set by the application to the keywords and connection  attributes of the connections in the pool. The value of this attribute determines the level of precision of the matching criteria. The following values are used to set the value of this attribute:

**SQL_CP_STRICT_MATCH**

Only connections that exactly match the connection options in the call and the connection attributes set by the application are reused. This is the default.

**SQL_CP_RELAXED_MATCH**

Connections with matching connection string keywords can be used. Keywords must match, but not all connection attributes must  match.

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetEnvAttr"></a>SetEnvAttr (CODBCBase)

Sets attributes that govern aspects of environments.

```
FUNCTION SetEnvAttr (BYVAL Attribute AS SQLINTEGER, BYVAL ValuePtr AS SQLPOINTER, _
   BYVAL StringLength AS SQLINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *Attribute* | Attribute to set. |
| *ValuePtr* | Pointer to the value to be associated with *Attribute*. Depending on the value of *Attribute*, *ValuePtr* will be a 32-bit integer value or point to a null-terminated character string. |
| *StringLength* | If *ValuePtr* points to a character string or a binary buffer, this argument should be the length of *ValuePtr*. If *ValuePtr* is an integer, *StringLength* is ignored. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="SetErrorProc"></a>SetErrorProc (CODBCBase)

Sets the address of an application defined error callback.

```
SUB SetErrorProc (BYVAL pProc AS ANY PTR, BYVAL reportWarnings AS BOOLEAN = FALSE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pProc* | Address of the application defined callback. |
| *reportWarnings* | Optional. Report also warnings. |

#### Example of an application defined callback:

```
SUB ODBC_ErrorCallback (BYVAL nResult AS SQLRETURN, BYREF wszSrc AS WSTRING, BYREF wszErrorMsg AS WSTRING)
   PRINT "Error: " & STR(nResult) & " - Source: " & wszSrc
   IF LEN(wszErrorMsg) THEN PRINT wszErrorMsg
END SUB

pDbc.SetErrorProc(@ODBC_ErrorCallback)    ' // Sets the error callback for the connection object
pStmt.SetErrorProc(@ODBC_ErrorCallback)   ' // Sets the error callback for the statement object
```

# <a name="SetOutputNTS"></a>SetOutputNTS (CODBCBase)

Returns a 32-bit integer that determines how the driver returns string data. If SQL_TRUE, the driver returns string data null-terminated. If SQL_FALSE, the driver does not return string data null-terminated. This attribute defaults to SQL_TRUE. A call to SetEnvAttr to set it to SQL_TRUE returns SQL_SUCCESS. A call to SetEnvAttr to set it to SQL_FALSE returns SQL_ERROR and SQLSTATE HYC00.

**Note**: Optional feature not implemented in Microsoft Access Driver.

```
FUNCTION SetOutputNTS (BYVAL dwAttr AS SQLUINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwAttr* | Value of the attribute. SQL_TRUE or SQL_FALSE. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="ConstructorsDb"></a>Constructors (CODBC)

Allocates a connection handle and, if needed, an environment handle, and opens the database.

```
CONSTRUCTOR CODBC (BYREF wszConnectionString AS WSTRING, BYVAL nODbcVersion AS SQLINTEGER = SQL_OV_ODBC3, _
   BYVAL ConnectionPoolingAttr AS SQLUINTEGER = 0)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszConnectionString* | The connection string. |
| *nOdbcVersion* | Optional. ODBC version number: SQL_OV_ODBC2, SQL_OV_ODBC3 or SQL_OV_ODBC3_80. |
| *ConnectionPoolingAttr* | Optional. SQL_CP_ONE_PER_DRIVER or SQL_CP_ONE_PER_HENV. |

Establishes connections to a driver and a data source. The connection handle references storage of all information about the connection to the data source, including status, transaction state, and error information. 

```
CONSTRUCTOR CODBC (BYREF wszServerName AS WSTRING, BYREF wszUserName AS WSTRING, _
   BYREF wszAuthentication AS WSTRING, BYVAL nODbcVersion AS SQLINTEGER = SQL_OV_ODBC3, _
   BYVAL ConnectionPoolingAttr AS SQLUINTEGER = 0)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszServerName* | Data source name. The data might be located on the same computer as the program, or on another computer somewhere on a network. |
| *wszUserName* | User identifier. |
| *wszAuthentication* | Authentication string (typically the password). |
| *nOdbcVersion* | Optional. ODBC version number: SQL_OV_ODBC2, SQL_OV_ODBC3 or SQL_OV_ODBC3_80. |
| *ConnectionPoolingAttr* | Optional. SQL_CP_ONE_PER_DRIVER or SQL_CP_ONE_PER_HENV. |

#### GetLastResult

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_ERROR, SQL_INVALID_HANDLE.

# <a name="CommitTran"></a>CommitTran (CODBC)

Requests a commit operation for all active operations on all statements associated with an environment. 

```
FUNCTION CommitTran () AS SQLRETURN
```
#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Example

```
#include once "Afx/COdbc/COdbc.inc"
USING Afx

' // Create a connection object and connect with the database
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr
IF pDbc.Handle = NULL THEN PRINT "Unable to create the connection handle" : SLEEP : END

' // Allocate an statement object
DIM pStmt AS COdbcStmt = @pDbc
IF pStmt.Handle = NULL THEN PRINT "Unable to create the statement handle" : SLEEP : END

' // Cursor type
pStmt.SetMultiuserKeysetCursor
' // Bind the columns
DIM AS LONG lAuId, cbAuId
pStmt.BindCol(1, @lAuId, @cbAuId)
DIM wszAuthor AS WSTRING * 260, cbAuthor AS LONG
pStmt.BindCol(2, @wszAuthor, SIZEOF(wszAuthor), @cbAuthor)
DIM iYearBorn AS SHORT, cbYearBorn AS LONG
pStmt.BindCol(3, @iYearBorn, @cbYearBorn)

' // Generate a result set
pStmt.ExecDirect ("SELECT * FROM Authors ORDER BY Author")

' // Fill the values of the bounded application variables and its sizes
lAuId     = 998               : cbAuID     = SIZEOF(lAuId)
wszAuthor = "Edgar Allan Poe" : cbAuthor   = LEN(wszAuthor)
iYearBorn = 1809              : cbYearBorn = SIZEOF(iYearBorn)
' // Add the record
pStmt.AddRecord
IF pStmt.Error = FALSE THEN PRINT "Record added" ELSE PRINT pStmt.GetErrorInfo

' // Commit the transaction
' pDbc.CommitTran
' IF pDbc.Error = FALSE THEN PRINT "Commit succeeded"
' or Rollbacks it because this is a test
pDbc.RollbackTran 
IF pDbc.Error = FALSE THEN PRINT "Rollback succeeded" ELSE PRINT pDbc.GetErrorInfo
```

# <a name="EnvHandle"></a>EnvHandle (CODBC)

Returns the environment handle.

```
FUNCTION EnvHandle () AS SQLHANDLE
```

# <a name="Functions"></a>Functions (CODBC)

Returns information about whether a driver supports a specific ODBC function.

```
FUNCTION Functions (BYVAL FunctionId AS SQLUSMALLINT) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *FunctionId* | A value that identifies the ODBC function of interest. |

#### Return value

TRUE if the function is supported, FALSE if it is not.

**Result value** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

**Values to identify ODBC APIs**:

```
SQL_API_SQLALLOCCONNECT
SQL_API_SQLALLOCENV
SQL_API_SQLALLOCHANDLE
SQL_API_SQLALLOCSTMT
SQL_API_SQLBINDCOL
SQL_API_SQLBINDPARAM
SQL_API_SQLCANCEL
SQL_API_SQLCLOSECURSOR
SQL_API_SQLCOLATTRIBUTE
SQL_API_SQLCOLUMNS
SQL_API_SQLCONNECT
SQL_API_SQLCOPYDESC
SQL_API_SQLDATASOURCES
SQL_API_SQLDESCRIBECOL
SQL_API_SQLDISCONNECT
SQL_API_SQLENDTRAN
SQL_API_SQLERROR
SQL_API_SQLEXECDIRECT
SQL_API_SQLEXECUTE
SQL_API_SQLFETCH
SQL_API_SQLFETCHSCROLL
SQL_API_SQLFREECONNECT
SQL_API_SQLFREEENV
SQL_API_SQLFREEHANDLE
SQL_API_SQLFREESTMT
SQL_API_SQLGETCONNECTATTR
SQL_API_SQLGETCONNECTOPTION
SQL_API_SQLGETCURSORNAME
SQL_API_SQLGETDATA
SQL_API_SQLGETDESCFIELD
SQL_API_SQLGETDESCREC
SQL_API_SQLGETDIAGFIELD
SQL_API_SQLGETDIAGREC
SQL_API_SQLGETENVATTR
SQL_API_SQLGETFUNCTIONS
SQL_API_SQLGETINFO
SQL_API_SQLGETSTMTATTR
SQL_API_SQLGETSTMTOPTION
SQL_API_SQLGETTYPEINFO
SQL_API_SQLNUMRESULTCOLS
SQL_API_SQLPARAMDATA
SQL_API_SQLPREPARE
SQL_API_SQLPUTDATA
SQL_API_SQLROWCOUNT
SQL_API_SQLSETCONNECTATTR
SQL_API_SQLSETCONNECTOPTION
SQL_API_SQLSETCURSORNAME
SQL_API_SQLSETDESCFIELD
SQL_API_SQLSETDESCREC
SQL_API_SQLSETENVATTR
SQL_API_SQLSETPARAM
SQL_API_SQLSETSTMTATTR
SQL_API_SQLSETSTMTOPTION
SQL_API_SQLSPECIALCOLUMNS
SQL_API_SQLSTATISTICS
SQL_API_SQLTABLES
SQL_API_SQLTRANSACT
SQL_API_SQLALLOCHANDLESTD
SQL_API_SQLBULKOPERATIONS
SQL_API_SQLBINDPARAMETER
SQL_API_SQLBROWSECONNECT
SQL_API_SQLCOLATTRIBUTES
SQL_API_SQLCOLUMNPRIVILEGES
SQL_API_SQLDESCRIBEPARAM
SQL_API_SQLDRIVERCONNECT
SQL_API_SQLDRIVERS
SQL_API_SQLEXTENDEDFETCH
SQL_API_SQLFOREIGNKEYS
SQL_API_SQLMORERESULTS
SQL_API_SQLNATIVESQL
SQL_API_SQLNUMPARAMS
SQL_API_SQLPARAMOPTIONS
SQL_API_SQLPRIMARYKEYS
SQL_API_SQLPROCEDURECOLUMNS
SQL_API_SQLPROCEDURES
SQL_API_SQLSETPOS
SQL_API_SQLSETSCROLLOPTIONS
SQL_API_SQLTABLEPRIVILEGES
```

# <a name="GetConnectAttr"></a>GetConnectAttr (CODBC)

Returns the current setting of a connection attribute.

```
FUNCTION GetConnectAttr (BYVAL Attribute AS SQLINTEGER, BYVAL ValuePtr AS SQLPOINTER, _
   BYVAL BufferLength AS SQLINTEGER, BYVAL StringLength AS SQLINTEGER PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *Attribute* | Attribute to retrieve.  |
| *ValuePtr* | A pointer to memory in which to return the current value of the attribute specified by *Attribute*. |
| *BufferLength* |  If *Attribute* is an ODBC-defined attribute and *ValuePtr* points to a character string or a binary buffer, this argument should be the length of *ValuePtr*. If *Attribute* is an ODBC-defined attribute and *ValuePtr* is an integer, *BufferLength* is ignored.<br>If *Attribute* is a driver-defined attribute, the application indicates the nature of the attribute to the Driver Manager by setting the *BufferLength* argument. *BufferLength* can have the following values:<br>If *ValuePtr* is a pointer to a character string, then *BufferLength* is the length of the string or SQL_NTS.<br>If *ValuePtr* is a pointer to a binary buffer, then the application places the result of the SQL_LEN_BINARY_ATTR(length) macro in *BufferLength*. This places a negative value in *BufferLength*.<br>If *ValuePtr* is a pointer to a value other than a character string or binary string, then *BufferLength* should have the value SQL_IS_POINTER.<br>If *ValuePtr* contains a fixed-length data type, then *BufferLength* is either SQL_IS_INTEGER or SQL_IS_UINTEGER, as appropriate. |
| *StringLength* | A pointer to a buffer in which to return the total number of bytes (excluding the null-termination character) available to return in *ValuePtr*. If *ValuePtr* is a null pointer, no length is returned. If the attribute value is a character string and the number of bytes available to return is greater than *BufferLength* minus the length of the null-termination character, the data in *ValuePtr* is truncated to *BufferLength* minus the length of the null-termination character and is null-terminated by the driver. |

#### Result code

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_ERROR, or SQL_INVALID_HANDLE.

```
FUNCTION GetConnectAttr (BYVAL Attribute AS SQLINTEGER, BYVAL ValuePtr AS SQLPOINTER, _
FUNCTION GetConnectAttrStr (BYREF wszAttribute AS WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszAttribute* | The attribute to retrieve. |

**"AccessMode"**

Returns an SQLUINTEGER value. SQL_MODE_READ_ONLY is used by the driver or data source as an indicator that the connection is not required to support SQL statements that cause updates to occur. This mode can be used to optimize locking strategies, transaction management, or other areas as appropriate to the driver or data source. The driver is not required to prevent such statements from being submitted to  the data source. The behavior of the driver and data source when asked to process SQL statements that are not read-only during a read-only connection is implementation-defined. SQL_MODE_READ_WRITE is the default.

**"AsyncEnable"**

SQL_ASYNC_ENABLE_OFF = Off (the default)<br>
SQL_ASYNC_ENABLE_ON = On

**"AutoIPD"**

Returns a read-only SQLUINTEGER value that specifies whether automatic population of the IPD after a call to Prepare is supported. Optional feature not implemented by the Microsoft Access Driver.

**"AutoCommit"**

SQL_AUTOCOMMIT_OFF = The driver uses manual-commit mode, and theapplication must explicitly commit or roll back transactions with OdbcEndTran.

SQL_AUTOCOMMIT_ON = The driver uses autocommit mode. Each statement is committed immediately after it is executed. This is the default. Any open transactions on the connection are committed when SQL_ATTR_AUTOCOMMIT is set to SQL_AUTOCOMMIT_ON to change from manual-commit mode to autocommit mode.

**"ConnectionDead"**

SQL_TRUE (1) or SQL_FALSE (0).

**"ConnectionTimeout"**

Returns an SQLUINTEGER value corresponding to the number of seconds to wait for any request on the connection to complete before returning to the application. The driver should return SQLSTATE HYT00 (Timeout expired) anytime that it is possible to time out in a situation not associated with query execution or login. If the value is equal to 0 (the default), there is no timeout. Optional feature not implemented by the Microsoft Access Driver.

**"CurrentCatalog"**

Returnss a character string containing the name of the catalog to be used by the data source.

**"Cursors"**

An SQLULEN value specifying how the Driver Manager uses the ODBC cursor library:

SQL_CUR_USE_IF_NEEDED = The Driver Manager uses the ODBC cursor library only if it is needed. If the driver supports the SQL_FETCH_PRIOR option in SQLFetchScroll, the Driver Manager uses the scrolling capabilities of the driver. Otherwise, it uses the ODBC cursor library.

SQL_CUR_USE_ODBC = The Driver Manager uses the ODBC cursor library.

SQL_CUR_USE_DRIVER = The Driver Manager uses the scrolling capabilities of the driver. This is the default setting.

Warning: The cursor library will be removed in a future version of Windows. Avoid using this feature in new development work and plan to modify applications that currently use this feature. Microsoft recommends using the driver's cursor functionality.

**"LoginTimeout"**

Returns an SQLUINTEGER value corresponding to the number of seconds to wait for a login request to complete before returning to the application. The default is driver-dependent. If *ValuePtr* is 0, the timeout is disabled and a connection attempt will wait indefinitely. If the specified timeout exceeds the maximum login timeout in the data source, the driver substitutes that value and returns SQLSTATE 01S02 (Option value changed).

**"MetadataID"**

Returns an SQLUINTEGER value that determines how the string arguments of catalog functions are treated. Optional feature not implemented by the Microsoft Access Driver.

If SQL_TRUE, the string argument of catalog functions are treated as identifiers. The case is not significant. For nondelimited strings, the driver removes any trailing spaces and the string is folded to uppercase. For delimited strings, the driver removes any leading or trailing spaces and takes literally whatever is between the delimiters. If one of these arguments is set to a null pointer, the function returns SQL_ERROR and SQLSTATE HY009 (Invalid use of null pointer). If SQL_FALSE, the string arguments of catalog functions are not treated as identifiers. The case is significant. They can either contain a string search pattern or not, depending on the argument. The default value is SQL_FALSE. The *TableType* argument of **SQLTables**, which takes a list of values, is not affected by this attribute. SQL_ATTR_METADATA_ID can also be set on the statement level. (It is the only connection attribute that is also a statement attribute.)

**"PacketSize"**

Returns an SQLUINTEGER value specifying the network packet size in bytes. Optional feature not implemented by the Microsoft Access Driver.

Note Many data sources either do not support this option or only can return but not set the network packet size. If the specified size exceeds the maximum packet size or is smaller than the minimum packet size, the driver substitutes that value and returns SQLSTATE 01S02 (Option value changed). If the application sets packet size after a connection has already been made, the driver will return SQLSTATE HY011 (Attribute cannot be set now).

**"QuietMode"**

Returns a 32-bit window handle.

If the window handle is a null pointer, the driver does not display any dialog boxes. If the window handle is not a null pointer, it should be the parent window handle of the application. This is the default. The driver uses this handle to display dialog boxes.

Note The SQL_ATTR_QUIET_MODE connection attribute does not apply to dialog boxes displayed by **DriverConnect**.

**"Trace"**

Returns an SQLUINTEGER value telling the Driver Manager whether to perform tracing.

SQL_OPT_TRACE_OFF = Tracing off (the default)<br>
SQL_OPT_TRACE_ON = Tracing on

When tracing is on, the Driver Manager writes each ODBC function call to the trace file. Note When tracing is on, the Driver Manager can return SQLSTATE IM013 (Trace file error) from any function.

An application specifies a trace file with the **TraceFile** property. If the file already exists, the Driver Manager appends to the file. Otherwise, it creates the file. If tracing is on and no trace file has been specified, the Driver Manager writes to the file SQL.LOG in the root directory.

**"TraceFile"**

A string containing the name of the trace file.

**"TranslateLib"**

A null-terminated character string containing the name of a library containing the functions **SQLDriverToDataSource** and **SQLDataSourceToDriver** that the driver accesses to perform tasks such as character set translation. This option may be specified only if the driver has connected to the data source. The setting of this attribute will persist across connections. Optional feature not implemented by the Microsoft Access Driver.

**"TxnIsolation"**

An SQLUINTEGER bitmask.

The following bitmasks are used in conjunction with the flag to determine which options are supported:

SQL_TXN_READ_UNCOMMITTED<br>
SQL_TXN_READ_COMMITTED<br>
SQL_TXN_REPEATABLE_READ<br>
SQL_TXN_SERIALIZABLE

For descriptions of these isolation levels, see the description of SQL_DEFAULT_TXN_ISOLATION.

To set the transaction isolation level, an application calls **SetConnectAttr** to set the SQL_ATTR_TXN_ISOLATION attribute.

An SQL-92 Entry level-conformant driver will always return SQL_TXN_SERIALIZABLE as supported. A FIPS Transitional level-conformant driver will always return all of these options as supported.

## Return value

The current value of the attribute.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetDiagField"></a>GetDiagField (CODBC)

Returns the current value of a field of a record of the diagnostic data structure (associated with an environment handle) that contains error, warning, and status information.

```
FUNCTION GetDiagField (BYVAL RecNumber AS SQLSMALLINT, BYVAL DiagIdentifier AS SQLSMALLINT, _
   BYVAL DiagInfoPtr AS SQLPOINTER, BYVAL BufferLength AS SQLSMALLINT, _
   BYVAL StringLength AS SQLSMALLINT PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the status record from which the application seeks information. Status records are numbered from 1. If the *DiagIdentifier* argument indicates any field of the diagnostics header, *RecNumber* is ignored. If not, it should be greater than 0. |
| *DiagIdentifier* | Indicates the field of the diagnostic whose value is to be returned. |
| *DiagInfoPtr* | Pointer to a buffer in which to return the diagnostic information. The data type depends on the value of *DiagIdentifier*. |
| *BufferLength* | If *DiagIdentifier* is an ODBC-defined diagnostic and *DiagInfoPtr* points to a character string or a binary buffer, this argument should be the length of *DiagInfoPtr*. If *DiagIdentifier* is an ODBC-defined field and *DiagInfoPtr* is an integer, *BufferLength* is ignored.<br>If *DiagIdentifier* is a driver-defined field, the application indicates the nature of the field to the Driver Manager by setting the *BufferLength* argument. *BufferLength* can have the following values:<br>If *DiagInfoPtr* is a pointer to a character string, then *BufferLength* is the length of the string or SQL_NTS.<br>If *DiagInfoPtr* is a pointer to a binary buffer, then the application places the result of the SQL_LEN_BINARY_ATTR(length) macro in BufferLength. This places a negative value in *BufferLength*.<br>If *DiagInfoPtr* is a pointer to a value other than a character string or binary string, then *BufferLength* should have the value SQL_IS_POINTER.<br>If *DiagInfoPtr* is contains a fixed-length data type, then *BufferLength* is SQL_IS_INTEGER, SQL_IS_UINTEGER, SQL_IS_SMALLINT, or SQL_IS_USMALLINT, as appropriate. |
| *StringLength* | Pointer to a buffer in which to return the total number of bytes (excluding the number of bytes required for the null-termination character) available to return in *DiagInfoPtr*, for character data. If the number of bytes available to return is greater than or equal to *BufferLength*, the text in *DiagInfoPtr* is truncated to *BufferLength* minus the length of a null-termination character. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, SQL_INVALID_HANDLE, or SQL_NO_DATA.

# <a name="GetDiagRec"></a>GetDiagRec (CODBC)

Returns the current values of multiple fields of a diagnostic record that contains error, warning, and status information. Unlike **GetDiagField**, which returns one diagnostic field per call, **GetDiagRec** returns several commonly used fields of a diagnostic record, including the SQLSTATE, the native error code, and the diagnostic message text.

```
FUNCTION GetDiagRec (BYVAL RecNumber AS SQLSMALLINT, BYVAL Sqlstate AS WSTRING PTR, _
   BYVAL NativeError AS SQLINTEGER PTR, BYVAL MessageText AS WSTRING PTR, _
   BYVAL BufferLength AS SQLSMALLINT, BYVAL TextLength AS SQLSMALLINT PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecNumber* | Indicates the status record from which the application seeks information. Status records are numbered from 1. |
| *SQlState* | Pointer to a buffer in which to return a five-character SQLSTATE code pertaining to the diagnostic record *RecNumber*. The first two characters indicate the class; the next three indicate the subclass. This information is contained in the SQL_DIAG_SQLSTATE diagnostic field. |
| *NativeError* | Pointer to a buffer in which to return the native error code, specific to the data source. This information is contained in the SQL_DIAG_NATIVE diagnostic field. |
| *MessageText* | Pointer to a buffer in which to return the diagnostic message text string. This information is contained in the SQL_DIAG_MESSAGE_TEXT diagnostic field. |
| *BufferLength* | Length of the *MessageText* buffer in characters. There is no maximum length of the diagnostic message text. |
| *TextLength* |  Pointer to a buffer in which to return the total number of bytes (excluding the number of bytes required for the null-termination character) available to return in *MessageText*. If the number of bytes available to return is greater than *BufferLength*, the diagnostic message text in *MessageText* is truncated to *BufferLength* minus the length of a null-termination character. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetErrorInfo"></a>GetErrorInfo (CODBC)

Returns a verbose description of the last error(s).

```
FUNCTION GetErrorInfo (BYVAL iErrorCode AS SQLRETURN = 0) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *iErrorCode* | Optional. The error code returned by **GetLastResult**. |

# <a name="GetInfo"></a>GetInfo (CODBC)

Returns general information about the driver and data source associated with a connection.

```
FUNCTION GetInfo (BYVAL InfoType AS SQLUSMALLINT, BYVAL InfoValuePtr AS SQLPOINTER, _
   BYVAL BufferLength AS SQLSMALLINT, BYVAL StringLength AS SQLSMALLINT PTR) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *InfoType* | The type of information to retrieve. |
| *InfoValuePtr* | Pointer to a buffer in which to return the information. Depending on the *InfoType* requested, the information returned will be one of the following: a null-terminated character string, an SQLUSMALLINT value, an SQLUINTEGER bitmask, an SQLUINTEGER flag, or a SQLUINTEGER binary value.<br>If the *InfoType* argument is SQL_DRIVER_HDESC or SQL_DRIVER_HSTMT, the *InfoValuePtr* argument is both input and output. |
| *BufferLength* | Length of the *InfoValuePtr* buffer. If the value in *InfoValuePtr* is not a character string or if *InfoValuePtr* is a null pointer, the *BufferLength* argument is ignored. The driver assumes that the size of *InfoValuePtr* is SQLUSMALLINT or SQLUINTEGER, based on the *InfoType*. since this method works with Unicode, the *BufferLength* argument must be an even number; if not, SQLSTATE HY090 (Invalid string or buffer length) is returned. |
| *StringLength* | Pointer to a buffer in which to return the total number of bytes (excluding the null-termination character for character data) available to return in *InfoValuePtr*. For character data, if the number of bytes available to return is greater than or equal to *BufferLength*, the information in *InfoValuePtr* is truncated to *BufferLength* bytes minus the length of a null-termination character and is null-terminated by the driver. For all other types of data, the value of *BufferLength* is ignored and the driver assumes the size of *InfoValuePtr* is SQLUSMALLINT or SQLUINTEGER, depending on the *InfoType*.<br>Important note: With the ODBC version that comes installed with Windows 7, you have to specify 2 (SQLUSMALLINT) or 4 (SQLUINTEGER) or it will fail. |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

Returns general information about the driver and data source associated with a connection.

```
FUNCTION GetInfo (BYREF wszInfoType AS WSTRING) AS SQLINTEGER
FUNCTION GetInfoStr (BYREF wszInfoType AS WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszInfoType* | The type of information to retrieve. |

**"AccessibleProcedures"**

Returns a character string: "Y" if the user can execute all procedures returned by Procedures; "N" if there may be procedures returned that the user cannot execute.

**"AccessibleTables"**

Returns a character string: "Y" if the user is guaranteed SELECT privileges to all tables returned by Tables; "N" if there may be tables returned that the user cannot access.

**"ActiveEnvironments"**

Returns an SQLUSMALLINT value specifying the maximum number of active environments that the driver can support. If there is no specified limit or the limit is unknown, this value is set to zero.

**"AggregateFunctions"**

Returns an SQLUINTEGER bitmask enumerating support for aggregation functions.

SQL_AF_ALL<br>
SQL_AF_AVG<br>
SQL_AF_COUNT<br>
SQL_AF_DISTINCT<br>
SQL_AF_MAX<br>
SQL_AF_MIN<br>
SQL_AF_SUM

An SQL-92 Entry level-conformant driver will always return all of these options as supported.

**"AlterDomain"**

Returns an SQLUINTEGER bitmask enumerating the clauses in the ALTER DOMAIN statement, as defined in SQL-92, supported by the data source. An SQL-92 Full level-compliant driver will always return all of the bitmasks. A return value of "0" means that the ALTER DOMAIN statement is not supported.

The SQL-92 or FIPS conformance level at which this feature needs to be supported is shown in parentheses next to each bitmask.

**"AlterTable"**

An SQLUINTEGER bitmask enumerating the clauses in the ALTER TABLE statement supported by the data source.

The SQL-92 or FIPS conformance level at which this feature needs to be supported is shown in parentheses next to each bitmask.

The following bitmasks are used to determine which clauses are supported:

SQL_AT_ADD_COLUMN_COLLATION = <add column> clause is supported, with facility to specify column collation (Full level) (ODBC 3.0)

SQL_AT_ADD_COLUMN_DEFAULT = <add column> clause is supported, with facility to specify column defaults (FIPS Transitional level) (ODBC 3.0)

SQL_AT_ADD_COLUMN_SINGLE = <add column> is supported (FIPS Transitional level) (ODBC 3.0)

SQL_AT_ADD_CONSTRAINT = <add column> clause is supported, with facility to specify column constraints (FIPS Transitional level) (ODBC 3.0)

SQL_AT_ADD_TABLE_CONSTRAINT = <add table constraint> clause is supported (FIPS Transitional level) (ODBC 3.0)

SQL_AT_CONSTRAINT_NAME_DEFINITION = <constraint name definition> is supported for naming column and table constraints (Intermediate level) (ODBC 3.0)

SQL_AT_DROP_COLUMN_CASCADE = <drop column> CASCADE is supported (FIPS Transitional level) (ODBC 3.0)

SQL_AT_DROP_COLUMN_DEFAULT = <alter column> <drop column default clause> is supported (Intermediate level) (ODBC 3.0)

SQL_AT_DROP_COLUMN_RESTRICT = <drop column> RESTRICT is supported (FIPS Transitional level) (ODBC 3.0)

SQL_AT_DROP_TABLE_CONSTRAINT_CASCADE (ODBC 3.0)

SQL_AT_DROP_TABLE_CONSTRAINT_RESTRICT = <drop column> RESTRICT is supported (FIPS Transitional level) (ODBC 3.0)

SQL_AT_SET_COLUMN_DEFAULT = <alter column> <set column default clause> is supported (Intermediate level) (ODBC 3.0)

The following bits specify the support <constraint attributes> if specifying column or table constraints is supported (the SQL_AT_ADD_CONSTRAINT bit is set):

SQL_AT_CONSTRAINT_INITIALLY_DEFERRED (Full level) (ODBC 3.0)

SQL_AT_CONSTRAINT_INITIALLY_IMMEDIATE (Full level) (ODBC 3.0)

SQL_AT_CONSTRAINT_DEFERRABLE (Full level) (ODBC 3.0)

SQL_AT_CONSTRAINT_NON_DEFERRABLE (Full level) (ODBC 3.0)

**"AsyncMode"**

Returns an SQLUINTEGER value indicating the level of asynchronous support in the driver.

The level of asynchronous support in the driver:

SQL_AM_CONNECTION = Connection level asynchronous execution is supported. Either all statement handles associated with a given connection handle are in asynchronous mode or all are in synchronous mode. A statement handle on a connection cannot be in asynchronous mode while another statement handle on the same connection is in synchronous mode, and vice versa.

SQL_AM_STATEMENT = Statement level asynchronous execution is supported. Some statement handles associated with a connection handle can be in asynchronous mode, while other statement handles on the same connection are in synchronous mode.

SQL_AM_NONE = Asynchronous mode is not supported.

**"BatchRowCount"**

Returns an SQLUINTEGER bitmask enumerating the behavior of the driver with respect to the availability of row counts.

The following bitmasks are used in conjunction with the information type:

SQL_BRC_ROLLED_UP = Row counts for consecutive INSERT, DELETE, or UPDATE statements are rolled up into one. If this bit is not set, then row counts are available for each individual statement.

SQL_BRC_PROCEDURES = Row counts, if any, are available when a batch is executed in a stored procedure. If row counts are available, they can be rolled up or individually available, depending on the SQL_BRC_ROLLED_UP bit.

SQL_BRC_EXPLICIT = Row counts, if any, are available when a batch is executed directly by calling OdbcExecute or OdbcExecDirect. If row counts are available, they can be rolled up or individually available, depending on the SQL_BRC_ROLLED_UP bit.

**"BatchSupport"**

Returns an SQLUINTEGER bitmask enumerating the driver's support for batches.

The following bitmasks are used to determine which level is supported:

SQL_BS_SELECT_EXPLICIT = The driver supports explicit batches that can have result-set generating statements.

SQL_BS_ROW_COUNT_EXPLICIT = The driver supports explicit batches that can have row-count generating statements.

SQL_BS_SELECT_PROC = The driver supports explicit procedures that can have result-set generating statements.

SQL_BS_ROW_COUNT_PROC = The driver supports explicit procedures that can have row-count generating statements.

**"BookmarkPersistence"**

Returns an SQLUINTEGER bitmask enumerating the operations through which bookmarks persist.

The following bitmasks are used in conjunction with the flag to determine through which options bookmarks persist:

SQL_BP_CLOSE = Bookmarks are valid after an application calls FreeStmt with the SQL_CLOSE option, or SQLCloseCursor to close the cursor associated with a statement.

SQL_BP_DELETE = The bookmark for a row is valid after that row has been deleted.

SQL_BP_DROP = Bookmarks are valid after an application calls OdbcFreeHandle with a HandleType of SQL_HANDLE_STMT to drop a statement.

SQL_BP_TRANSACTION = Bookmarks are valid after an application commits or rolls back a transaction.

SQL_BP_UPDATE = The bookmark for a row is valid after any column in that  row has been updated, including key columns.

SQL_BP_OTHER_HSTMT = A bookmark associated with one statement can be used with another statement. Unless SQL_BP_CLOSE or SQL_BP_DROP is specified, the cursor on the first statement must be open.

**"CatalogLocation"**

Returns an SQLUSMALLINT value indicating the position of the catalog in a qualified table name.

The position of the catalog in a qualified table name:

SQL_CL_START<br>
SQL_CL_END

For example, an Xbase driver returns SQL_CL_START because the directory (catalog) name is at the start of the table name, as in \EMPDATA\EMP.DBF.

An ORACLE Server driver returns SQL_CL_END because the catalog is at the end of the table name, as in ADMIN.EMP@EMPDATA.

An SQL-92 Full level-conformant driver will always return SQL_CL_START.

A value of 0 is returned if catalogs are not supported by the data source.

To find out whether catalogs are supported, an application calls OdbcGetInfo with the SQL_CATALOG_NAME information type.

This InfoType has been renamed for ODBC 3.0 from the ODBC 2.0 InfoType SQL_QUALIFIER_LOCATION.

**"CatalogName"**

Returns a character string: "Y" if the server supports catalog names, or "N" if it does not.

An SQL-92 Full level-conformant driver will always return "Y".

**"CatalogNameSeparator"**

Returns a character string: the character or characters that the data source defines as the separator between a catalog name and the qualified name element that follows or precedes it.

An empty string is returned if catalogs are not supported by the data source. To find out whether catalogs are supported, an application calls OdbcGetInfo with the SQL_CATALOG_NAME information type. An SQL-92 Full level-conformant driver will always return ".".

This InfoType has been renamed for ODBC 3.0 from the ODBC 2.0  InfoTypeSQL_QUALIFIER_NAME_SEPARATOR.

**"CatalogTerm"**

Returns a character string with the data source vendor's name for a catalog; for example, "database" or "directory". This string can be in upper, lower, or mixed case.

An empty string is returned if catalogs are not supported by the data source. To find out whether catalogs are supported, an application calls GetInfo with the SQL_CATALOG_NAME information type. An SQL-92 Full level-conformant driver will always return "catalog".

This InfoType has been renamed for ODBC 3.0 from the ODBC 2.0 InfoType SQL_QUALIFIER_TERM.

**"CatalogUsage"**

Returns an SQLUINTEGER bitmask enumerating the statements in which catalogs can be used.

SQL_CU_DML_STATEMENTS = Catalogs are supported in all Data Manipulation Language statements: SELECT, INSERT, UPDATE, DELETE, and if supported, SELECT FOR UPDATE and positioned update and delete statements.

SQL_CU_PROCEDURE_INVOCATION = Catalogs are supported in the ODBC procedure invocation statement.

SQL_CU_TABLE_DEFINITION = Catalogs are supported in all table definition statements: CREATE TABLE, CREATE VIEW, ALTER TABLE, DROP TABLE, and DROP VIEW.

SQL_CU_INDEX_DEFINITION = Catalogs are supported in all index definition statements: CREATE INDEX and DROP INDEX.

SQL_CU_PRIVILEGE_DEFINITION = Catalogs are supported in all privilege definition statements: GRANT and REVOKE.

A value of 0 is returned if catalogs are not supported by the data source.

To find out whether catalogs are supported, an application calls SQLGetInfo with the SQL_CATALOG_NAME information type. An SQL-92 Full level-conformant driver will always return a bitmask with all of these bits set.

This InfoType has been renamed for ODBC 3.0 from the ODBC 2.0 InfoType SQL_QUALIFIER_USAGE.

**"CollationSeq"**

Returns the name of the collation sequence. This is a character string that indicates the name of the default collation for the default character set for this server (for example, 'ISO 8859-1' or EBCDIC). If this is unknown, an empty string will be returned. An SQL-92 Full level-conformant driver will always return a non-empty string.

**"ColumnAlias"**

Returns a character string: "Y" if the data source supports column aliases; otherwise, "N".

A column alias is an alternate name that can be specified for a column in the select list by using an AS clause. An SQL-92 Entry level-conformant driver will always return "Y".

**"ConcatNullBehavior"**

Returns an SQLUSMALLINT value indicating how the data source handles the concatenation of NULL valued character data type columns with non-NULL valued character data type columns.

SQL_CB_NULL = Result is NULL valued.

SQL_CB_NON_NULL = Result is concatenation of non-NULL valued column or columns.

An SQL-92 Entry level-conformant driver will always return SQL_CB_NULL.

**"ConvertBigInt"**

Returns an SQLUINTEGER bitmask. If the bitmask equals zero, the data source does not support any conversions from data of the named type, including conversion to the same data type.

To find out if a data source supports the conversion of SQL_BIGINT data to other data type, an application calls ConvertBigInt and then performs an AND operation with the returned bitmask and the bitmask of the type to be converted. If the resulting value is nonzero, the conversion is supported.

**"ConvertBinary"**

Returns an SQLUINTEGER bitmask. If the bitmask equals zero, the data source does not support any conversions from data of the named type, including conversion to the same data type.

To find out if a data source supports the conversion of SQL_BINARY data to other data type, an application calls ConvertBinary and then performs an AND operation with the returned bitmask and the bitmask of the type to be converted. If the resulting value is nonzero, the conversion is supported.

**"ConvertBit"**

Returns an SQLUINTEGER bitmask. If the bitmask equals zero, the data source does not support any conversions from data of the named type, including conversion to the same data type.

To find out if a data source supports the conversion of SQL_BIT data to other data type, an application calls ConvertBit and then performs an AND operation with the returned bitmask and the bitmask of the type to be converted. If the resulting value is nonzero, the conversion is supported.

**"ConvertChar"**

Returns an SQLUINTEGER bitmask. If the bitmask equals zero, the data source does not support any conversions from data of the named type, including conversion to the same data type.

To find out if a data source supports the conversion of SQL_CHAR data to other data type, an application calls ConvertChar and then performs an AND operation with the returned bitmask and the bitmask of the type to be converted. If the resulting value is nonzero, the conversion is supported.

**"ConvertDate"**

Returns an SQLUINTEGER bitmask. If the bitmask equals zero, the data source does not support any conversions from data of the named type, including conversion to the same data type.

To find out if a data source supports the conversion of SQL_DATE data to other data type, an application calls ConvertDate and then performs an AND operation with the returned bitmask and the bitmask of the type to be converted. If the resulting value is nonzero, the conversion is supported.

**"ConvertDecimal"**

Returns an SQLUINTEGER bitmask. If the bitmask equals zero, the data source does not support any conversions from data of the named type, including conversion to the same data type.

To find out if a data source supports the conversion of SQL_DECIMAL data to other data type, an application calls ConvertDecimal and then performs an AND operation with the returned bitmask and the bitmask of the type to be converted. If the resulting value is nonzero, the conversion is supported.

**"ConvertDouble"**

Returns an SQLUINTEGER bitmask. If the bitmask equals zero, the data source does not support any conversions from data of the named type, including conversion to the same data type.

To find out if a data source supports the conversion of SQL_DOUBLE data to other data type, an application calls ConvertDouble and then performs an AND operation with the returned bitmask and the bitmask of the type to be converted. If the resulting value is nonzero, the conversion is supported.

**"ConvertFloat"**

Returns an SQLUINTEGER bitmask. If the bitmask equals zero, the data source does not support any conversions from data of the named type, including conversion to the same data type.

To find out if a data source supports the conversion of SQL_FLOAT data to other data type, an application calls ConvertFloat and then performs an AND operation with the returned bitmask and the bitmask of the type to be converted. If the resulting value is nonzero, the conversion is supported.

**"ConvertFunctions"**

Returns an SQLUINTEGER bitmask enumerating the scalar conversion functions supported by the driver and associated data source.

**"ConvertInteger"**

Returns an SQLUINTEGER bitmask. If the bitmask equals zero, the data source does not support any conversions from data of the named type, including conversion to the same data type.

To find out if a data source supports the conversion of SQL_INTEGER data to other data type, an application calls ConvertInteger and then performs an AND operation with the returned bitmask and the bitmask of the type to be converted. If the resulting value is nonzero, the conversion is supported.

**"ConvertIntervalDayTime"**

Returns an SQLUINTEGER bitmask. If the bitmask equals zero, the data source does not support any conversions from data of the named type, including conversion to the same data type.

To find out if a data source supports the conversion of SQL_INTERVAL_DAY_TIME data to other data type, an application calls ConvertIntervalDayTime and then performs an AND operation with the returned bitmask and the bitmask of the type to be converted. If the resulting value is nonzero, the conversion is supported.

**"ConvertIntervalYearMonth"**

Returns an SQLUINTEGER bitmask. If the bitmask equals zero, the data source does not support any conversions from data of the named type, including conversion to the same data type.

To find out if a data source supports the conversion of SQL_INTERVAL_YEAR_MONTH data to other data type, an application calls ConvertIntervalTearMonth and then performs an AND operation with the returned bitmask and the bitmask of the type to be converted. If the resulting value is nonzero, the conversion is supported.

**"ConvertLongVarBinary"**

Returns an SQLUINTEGER bitmask. If the bitmask equals zero, the data source does not support any conversions from data of the named type, including conversion to the same data type.

To find out if a data source supports the conversion of SQL_LONGVARBINARY data to other data type, an application calls ConvertLongVarBinary and then performs an AND operation with the returned bitmask and the bitmask of the type to be converted. If the resulting value is nonzero, the conversion is supported.

**"ConvertLongVarChar"**

Returns an SQLUINTEGER bitmask. If the bitmask equals zero, the data source does not support any conversions from data of the named type, including conversion to the same data type.

To find out if a data source supports the conversion of SQL_LONGVARCHAR data to other data type, an application calls ConvertLongVarChar and then performs an AND operation with the returned bitmask and the bitmask of the type to be converted. If the resulting value is nonzero, the conversion is supported.

**"ConvertNumeric"**

Returns an SQLUINTEGER bitmask. If the bitmask equals zero, the data source does not support any conversions from data of the named type, including conversion to the same data type.

To find out if a data source supports the conversion of SQL_NUMERIC data to other data type, an application calls ConvertNumeric and then performs an AND operation with the returned bitmask and the bitmask of the type to be converted. If the resulting value is nonzero, the conversion is supported.

**"ConvertReal"**

Returns an SQLUINTEGER bitmask. If the bitmask equals zero, the data source does not support any conversions from data of the named type, including conversion to the same data type.

To find out if a data source supports the conversion of SQL_REAL data to other data type, an application calls ConvertReal and then performs an AND operation with the returned bitmask and the bitmask of the type to be converted. If the resulting value is nonzero, the conversion is supported.

**"ConvertSmallInt"**

Returns an SQLUINTEGER bitmask. If the bitmask equals zero, the data source does not support any conversions from data of the named type, including conversion to the same data type.

To find out if a data source supports the conversion of SQL_SMALLINT data to other data type, an application calls ConvertSmallInt and then performs an AND operation with the returned bitmask and the bitmask of the type to be converted. If the resulting value is nonzero, the conversion is supported.

**"ConvertTime"**

Returns an SQLUINTEGER bitmask. If the bitmask equals zero, the data source does not support any conversions from data of the named type, including conversion to the same data type.

To find out if a data source supports the conversion of SQL_TIME data to other data type, an application calls ConvertTime and then performs an AND operation with the returned bitmask and the bitmask of the type to be converted. If the resulting value is nonzero, the conversion is supported.

**"ConvertTimeStamp"**

Returns an SQLUINTEGER bitmask. If the bitmask equals zero, the data source does not support any conversions from data of the named type, including conversion to the same data type.

To find out if a data source supports the conversion of SQL_TIMESTAMP data to other data type, an application calls ConvertTimeStamp and then performs an AND operation with the returned bitmask and the bitmask of the type to be converted. If the resulting value is nonzero, the conversion is supported.

**"ConvertTinyInt"**

Returns an SQLUINTEGER bitmask. If the bitmask equals zero, the data source does not support any conversions from data of the named type, including conversion to the same data type.

To find out if a data source supports the conversion of SQL_TINYINT data to other data type, an application calls ConvertTinyInt and then performs an AND operation with the returned bitmask and the bitmask of the type to be converted. If the resulting value is nonzero, the conversion is supported.

**ConvertVarBinary"**

Returns an SQLUINTEGER bitmask. If the bitmask equals zero, the data source does not support any conversions from data of the named type, including conversion to the same data type.

To find out if a data source supports the conversion of SQL_VARBINARY data to other data type, an application calls ConvertVarBinary and then performs an AND operation with the returned bitmask and the bitmask of the type to be converted. If the resulting value is nonzero, the conversion is supported.

**"ConvertVarChar"**

Returns an SQLUINTEGER bitmask. If the bitmask equals zero, the data source does not support any conversions from data of the named type, including conversion to the same data type.

To find out if a data source supports the conversion of SQL_VARCHAR data to other data type, an application calls ConvertVarChar and then performs an AND operation with the returned bitmask and the bitmask of the type to be converted. If the resulting value is nonzero, the conversion is supported.

**"CorrelationName"**

Returns an SQLUSMALLINT value indicating whether table correlation names are supported.

SQL_CN_NONE = Correlation names are not supported.

SQL_CN_DIFFERENT = Correlation names are supported but must differ from the names of the tables they represent.

SQL_CN_ANY = Correlation names are supported and can be any valid user-defined name.

An SQL-92 Entry level-conformant driver will always return SQL_CN_ANY.

**"CreateAssertion"**

Returns an SQLUINTEGER bitmask enumerating the clauses in the CREATE ASSERTION statement, as defined in SQL-92, supported by the data source.

**"CreateCharacterSet"**

Returns an SQLUINTEGER bitmask enumerating the clauses in the CREATE CHARACTER SET statement, as defined in SQL-92, supported by the data source.

**"CreateCollation"**

Returns an SQLUINTEGER bitmask enumerating the clauses in the CREATE COLLATION statement, as defined in SQL-92, supported by the data source.

**"CreateDomain"**

Returns an SQLUINTEGER bitmask enumerating the clauses in the CREATE DOMAIN statement, as defined in SQL-92, supported by the data source.

**"CreateSchema"**

Returns an SQLUINTEGER bitmask enumerating the clauses in the CREATE SCHEMA statement, as defined in SQL-92, supported by the data source.

**"CreateTable"**

Returns an SQLUINTEGER bitmask enumerating the clauses in the CREATE TABLE statement, as defined in SQL-92, supported by the data source.

**"CreateTranslation"**

Returns an SQLUINTEGER bitmask enumerating the clauses in the CREATE TRANSLATION statement, as defined in SQL-92, supported by the data source.

**"CreateView"**

Returns an SQLUINTEGER bitmask enumerating the clauses in the CREATE VIEW statement, as defined in SQL-92, supported by the data source.

**"CursorCommitBehavior"**

Returns an SQLUSMALLINT value indicating how a COMMIT operation affects cursors and prepared statements in the data source.

SQL_CB_DELETE = Close cursors and delete prepared statements. To use the cursor again, the application must reprepare and reexecute the statement.

SQL_CB_CLOSE = Close cursors. For prepared statements, the application can call OdbcExecute on the statement without calling Prepare again. 

SQL_CB_PRESERVE = Preserve cursors in the same position as before the COMMIT operation. The application can continue to fetch data, or it can close the cursor and reexecute the statement without repreparing it.

**"CursorRollbackBehavior"**

Returns an SQLUSMALLINT value indicating how a ROLLBACK operation affects cursors and prepared statements in the data source.
SQL_CB_DELETE = Close cursors and delete prepared statements. To use the  cursor again, the application must reprepare and reexecute the statement.

SQL_CB_CLOSE = Close cursors. For prepared statements, the application can call Execute on the statement without calling Prepare again.

SQL_CB_PRESERVE = Preserve cursors in the same position as before the ROLLBACK operation. The application can continue to fetch data, or it can close the cursor and reexecute the statement without repreparing it.

**"CursorSensitivitySupport"**

Returns an SQLUINTEGER value indicating the support for cursor sensitivity.

SQL_INSENSITIVE = All cursors on the statement handle show the result set without reflecting any changes made to it by any other cursor within the same transaction.

SQL_UNSPECIFIED = It is unspecified whether cursors on the statement handle make visible the changes made to a result set by another cursor within the same transaction. Cursors on the statement handle may make visible none, some, or all such changes.

SQL_SENSITIVE = Cursors are sensitive to changes made by other cursors within the same transaction.

An SQL-92 Entry level-conformant driver will always return the SQL_UNSPECIFIED option as supported.

An SQL-92 Full level-conformant driver will always return the SQL_INSENSITIVE option as supported.

**"DatabaseName"**

Returns a character string with the name of the current database in use, if the data source defines a named object called "database".

Note In ODBC 3.x, the value returned for this InfoType can also be returned by calling GetConnectAttr with an Attribute argument of SQL_ATTR_CURRENT_CATALOG.

**"DataExtensions"**

Returns an SQLUINTEGER bitmask enumerating extensions to GetData.

The following bitmasks are used in conjunction with the flag to determine what common extensions the driver supports for GetData:

SQL_GD_ANY_COLUMN = GetData can be called for any unbound column, including those before the last bound column. Note that the columns must be called in order of ascending column number unless SQL_GD_ANY_ORDER is also returned.

SQL_GD_ANY_ORDER = GetData can be called for unbound columns in any order. Note that OdbcGetData can be called only for columns after the last bound column unless SQL_GD_ANY_COLUMN is also returned.

SQL_GD_BLOCK = GetData can be called for an unbound column in any row in a block (where the rowset size is greater than 1) of data after positioning to that row with SetPos.

SQL_GD_BOUND = GetData can be called for bound columns as well as unbound columns. A driver cannot return this value unless it also returns SQL_GD_ANY_COLUMN.

GetData is required to return data only from unbound columns that occur after the last bound column, are called in order of increasing column number, and are not in a row in a block of rows.

If a driver supports bookmarks (either fixed-length or variable-length), it must support calling GetData on column 0. This support is required regardless of what the driver returns for a call to GetInfo with the SQL_GETDATA_EXTENSIONS InfoType.

**"DataSourceName"**

Returns a character string with the data source name used during connection. If the application called Connect, this is the value of the szDSN argument. If the application called DriverConnect or BrowseConnect, this is the value of the DSN keyword in the connection string passed to the driver. If the connection string did not contain the DSN keyword (such as when it contains the DRIVER keyword), this is an empty string.

**"DataSourceReadOnly"**

Returns a character string. "Y" if the data source is set to READ ONLY mode, "N" if it is otherwise.

This characteristic pertains only to the data source itself; it is not a characteristic of the driver that enables access to the data source. A driver that is read/write can be used with a data source that is read-only. If a driver is read-only, all of its data sources must be read-only and must return SQL_DATA_SOURCE_READ_ONLY.

**"DateTimeLiterals"**

An SQLUINTEGER bitmask enumerating the SQL-92 datetime literals supported by the data source. Note that these are the datetime literals listed in the SQL-92 specification and are separate from the datetime literal escape clauses defined by ODBC.

A FIPS Transitional level-conformant driver will always return the "1" value in the bitmask for the bits listed below. A value of "0" means that SQL-92 datetime literals are not supported.

**"DBMSName"**

Returns a character string with the name of the DBMS product accessed by the driver.

**"DBMSVer"**

Returns a character string indicating the version of the DBMS product accessed by the driver. The version is of the form ##.##.####, where the first two digits are the major version, the next two digits are the minor version, and the last four digits are the release version. The driver must render the DBMS product version in this form but can also append the DBMS product-specific version as well. For example, "04.01.0000 Rdb 4.1".

**"DDLIndex"**

Returns an SQLUINTEGER value that indicates support for creation and dropping of indexes.

**"DefaultTxnIsolation"**

Returns an SQLUINTEGER value that indicates the default transaction isolation level supported by the driver or data source, or zero if the data source does not support transactions.

**"DescribeParameter"**

Returns a character string: "Y" if parameters can be described; "N", if not. An SQL-92 Full level-conformant driver will usually return "Y" because it will support the DESCRIBE INPUT statement. Because this does not directly specify the underlying SQL support, however, describing parameters might not be supported, even in a SQL-92 Full level-conformant driver.

**"DMVer"**

A character string with the version of the Driver Manager.

The version is in the form ##.##.####.####, where:

The first set of two digits is the major ODBC version, as given by theconstant SQL_SPEC_MAJOR.

The second set of two digits is the minor ODBC version, as given by the constant SQL_SPEC_MINOR.

The third set of four digits is the Driver Manager major build number.

The last set of four digits is the Driver Manager minor build number.

**"DriverhDbc"**

Returns the driver's connection handle.

**"DriverhDesc"**

Returns the driver's descriptor handle.

**"DriverhEnv"**

Returns the driver's environment handle.

**"DriverhLib"**

Returns an SQLUINTEGER value, the hinst from the load library returned to the Driver Manager when it loaded the driver DLL (on a Microsoft® Windows® platform) or equivalent on a non-Windows platform. The handle is valid only for the connection handle specified in the call to GetInfo. This information type is implemented by the Driver Manager alone.

**"DriverhStmt"**

Returns the driver's statement handle.

**"DriverName"**

Returns a character string with the file name of the driver used to access the data source.

**"DriverOdbcVer"**

Returns a character string with the version of ODBC that the driver supports.

The version is of the form ##.##, where the first two digits are the major version and the next two digits are the minor version.

SQL_SPEC_MAJOR and SQL_SPEC_MINOR define the major and minor version numbers.

For the version of ODBC described in this manual, these are 3 and 0, and the driver should return "03.00".

**"DriverVer"**

Returns a character string with the version of the driver and optionally, a description of the driver.

At a minimum, the version is of the form ##.##.####, where the first two digits are the major version, the next two digits are the minor version, and the last four digits are the release version.

**"DropAssertion"**

Returns an SQLUINTEGER bitmask enumerating the clauses in the DROP ASSERTION statement, as defined in SQL-92, supported by the data source.

**"DropCharacterSet"**

Returns an SQLUINTEGER bitmask enumerating the clauses in the DROP CHARACTER SET statement, as defined in SQL-92, supported by the data source.

**"DropCollation"**

Returns an SQLUINTEGER bitmask enumerating the clauses in the DROP COLLATION statement, as defined in SQL-92, supported by the data source.

**"DropDomain"**

Returns an SQLUINTEGER bitmask enumerating the clauses in the DROP DOMAIN statement, as defined in SQL-92, supported by the data source.

The following bitmasks are used to determine which clauses are supported:

SQL_DD_DROP_DOMAIN<br>
SQL_DD_CASCADE<br>
SQL_DD_RESTRICT

An SQL-92 Intermediate level-conformant driver will always return all of these options as supported.

**"DropSchema"**

Returns an SQLUINTEGER bitmask enumerating the clauses in the DROP SCHEMA statement, as defined in SQL-92, supported by the data source.

The following bitmasks are used to determine which clauses are supported:

SQL_DS_DROP_SCHEMA<br>
SQL_DS_CASCADE<br>
SQL_DS_RESTRICT

An SQL-92 Intermediate level-conformant driver will always return all of these options as supported.

**"DropTable"**

Returns an SQLUINTEGER bitmask enumerating the clauses in the DROP TABLE statement, as defined in SQL-92, supported by the data source.

**"DropTranslation"**

Returns an SQLUINTEGER bitmask enumerating the clauses in the DROP TRANSLATION statement, as defined in SQL-92, supported by the data source.

**"DropView"**

Returns an SQLUINTEGER bitmask enumerating the clauses in the DROP VIEW statement, as defined in SQL-92, supported by the data source.

**"DynamicCursorAttributes1"**

Returns an SQLUINTEGER bitmask that describes the attributes of a dynamic cursor that are supported by the driver. This bitmask contains the first subset of attributes; for the second subset, see DynamicCursorAttributes2.

**"DynamicCursorAttributes2"**

Returns an SQLUINTEGER bitmask that describes the attributes of a dynamic cursor that are supported by the driver. This bitmask contains the second subset of attributes; for the first subset, see DynamicCursorAttributes1.

**"ExpressionsInOrderBy"**

Returns a character string: "Y" if the data source supports expressions in the ORDER BY list; "N" if it does not.

**"FileUsage"**

Returns an SQLUSMALLINT value indicating how a single-tier driver directly treats files in a data source:

SQL_FILE_NOT_SUPPORTED = The driver is not a single-tier driver. For example, an ORACLE driver is a two-tier driver.

SQL_FILE_TABLE = A single-tier driver treats files in a data source as tables. For example, an Xbase driver treats each Xbase file as a table.

SQL_FILE_CATALOG = A single-tier driver treats files in a data source as a catalog. For example, a Microsoft® Access driver treats each Microsoft Access file as a complete database.

An application might use this to determine how users will select data. For example, Xbase users often think of data as stored in files, while ORACLE and MicrosoftAccess users generally think of data as stored in tables. When a user selects an Xbase data source, the application could display the Windows File Open common dialog box; when the user selects a Microsoft Access or ORACLE data source, the application could display a custom Select Table dialog box.

**"ForwardOnlyCursorAttributes1"**

Returns an SQLUINTEGER bitmask that describes the attributes of a forward-only cursor that are supported by the driver. This bitmask contains the first subset of attributes; for the second subset, see ForwardOnlyCursorAttributes2.

**"ForwardOnlyCursorAttributes2"**

Returns an SQLUINTEGER bitmask that describes the attributes of a forward-only cursor that are supported by the driver. This bitmask contains the second subset of attributes; for the first subset, see ForwardOnlyCursorAttributes1.

**"GroupBy"**

Returns an SQLUSMALLINT value specifying the relationship between the columns in the GROUP BY clause and the nonaggregated columns in the select list.

SQL_GB_COLLATE = A COLLATE clause can be specified at the end of each grouping column. (ODBC 3.0)

SQL_GB_NOT_SUPPORTED = GROUP BY clauses are not supported. (ODBC 2.0)

SQL_GB_GROUP_BY_EQUALS_SELECT = The GROUP BY clause must contain all nonaggregated columns in the select list. It cannot contain any other columns. For example, SELECT DEPT, MAX(SALARY) FROM EMPLOYEE GROUP BY DEPT. (ODBC 2.0)

SQL_GB_GROUP_BY_CONTAINS_SELECT = The GROUP BY clause must contain all nonaggregated columns in the select list. It can contain columns that are not in the select list. For example, SELECT DEPT, MAX(SALARY) FROM EMPLOYEE GROUP BY DEPT, AGE. (ODBC 2.0)

SQL_GB_NO_RELATION = The columns in the GROUP BY clause and the select list are not related. The meaning of nongrouped, nonaggregated columns in the select list is data source-dependent. For example, SELECT DEPT, SALARY FROM EMPLOYEE GROUP BY DEPT, AGE. (ODBC 2.0)

An SQL-92 Entry level-conformant driver will always return the SQL_GB_GROUP_BY_EQUALS_SELECT option as supported. An SQL-92 Full level-conformant driver will always return the SQL_GB_COLLATE option as supported. If none of the options is supported, the GROUP BY clause is not supported by the data source.

**"IdentifierCase"**

Returns an SQLUSMALLINT value with the supported identifier case.

SQL_IC_UPPER = Identifiers in SQL are not case-sensitive and are stored in uppercase in system catalog.

SQL_IC_LOWER = Identifiers in SQL are not case-sensitive and are stored in lowercase in system catalog.

SQL_IC_SENSITIVE = Identifiers in SQL are case-sensitive and are stored in mixed case in system catalog.

SQL_IC_MIXED = Identifiers in SQL are not case-sensitive and are stored in mixed case in system catalog.

Because identifiers in SQL-92 are never case-sensitive, a driver that conforms strictly to SQL-92 (any level) will never return the SQL_IC_SENSITIVE option as supported.

**"IdentifierQuoteChar"**

The character string used as the starting and ending delimiter of a quoted (delimited) identifier in SQL statements. (Identifiers passed as arguments to ODBC functions do not need to be quoted.) If the data source does not support quoted identifiers, a blank is returned.

This character string can also be used for quoting catalog function arguments when the connection attribute SQL_ATTR_METADATA_ID is set to SQL_TRUE.

Because the identifier quote character in SQL-92 is the double quotation mark ("), a driver that conforms strictly to SQL-92 will always return the double quotation mark character.

**"IndexKeywords"**

Returns an SQLUINTEGER bitmask that enumerates keywords in the CREATE INDEX statement that are supported by the driver.

SQL_IK_NONE = None of the keywords is supported.<br>
SQL_IK_ASC = ASC keyword is supported.<br>
SQL_IK_DESC = DESC keyword is supported.<br>
SQL_IK_ALL = All keywords are supported.

To see if the CREATE INDEX statement is supported, an application calls GetInfo with the SQL_DLL_INDEX information type.

**"InfoSchemaViews"**

Returns an SQLUINTEGER bitmask enumerating the views in the INFORMATION_SCHEMA that are supported by the driver. The views in, and the contents of, INFORMATION_SCHEMA are as defined in SQL-92. The SQL-92 or FIPS conformance level at which this feature needs to be supported is shown in parentheses next to each bitmask.

**"InsertStatement"**

Returns an SQLUINTEGER bitmask that indicates support for INSERT statements.

SQL_IS_INSERT_LITERALS<br>
SQL_IS_INSERT_SEARCHED<br>
SQL_IS_SELECT_INTO

An SQL-92 Entry level-conformant driver will always return all of these options as supported.

**"Integrity"**

Returns a character string: "Y" if the data source supports the Integrity Enhancement Facility; "N" if it does not.

**"KeysetCursorAttributes1"**

Returns an SQLUINTEGER bitmask that describes the attributes of a keyset cursor that are supported by the driver. This bitmask contains the first  subset of attributes; for the second subset, see KeysetCursorAttributes2.
The following bitmasks are used to determine which attributes are supported:

SQL_CA1_NEXT = A FetchOrientation argument of SQL_FETCH_NEXT is supported in a call to FetchScroll when the cursor is a dynamic cursor.

SQL_CA1_ABSOLUTE = FetchOrientation arguments of SQL_FETCH_FIRST, SQL_FETCH_LAST, and SQL_FETCH_ABSOLUTE are supported in a call to FetchScroll when the cursor is a dynamic cursor. (The rowset that will be fetched is independent of the current cursor position.)

SQL_CA1_RELATIVE = FetchOrientation arguments of SQL_FETCH_PRIOR and SQL_FETCH_RELATIVE are supported in a call to FetchScroll when the cursor is a dynamic cursor. (The rowset that will be fetched is dependent on the current cursor position. Note that this is separated from SQL_FETCH_NEXT because in a forward-only cursor, only SQL_FETCH_NEXT is supported.)

SQL_CA1_BOOKMARK = A FetchOrientation argument of SQL_FETCH_BOOKMARK is supported in a call to FetchScroll when the cursor is a dynamic cursor.

SQL_CA1_LOCK_EXCLUSIVE = A LockType argument of SQL_LOCK_EXCLUSIVE is supported in a call to SetPos when the cursor is a dynamic cursor.

SQL_CA1_LOCK_NO_CHANGE = A LockType argument of SQL_LOCK_NO_CHANGE is supported in a call to SetPos when the cursor is a dynamic cursor.

SQL_CA1_LOCK_UNLOCK = A LockType argument of SQL_LOCK_UNLOCK is supported in a call to SetPos when the cursor is a dynamic cursor.

SQL_CA1_POS_POSITION = An Operation argument of SQL_POSITION is supported  in a call to SetPos when the cursor is a dynamic cursor.

SQL_CA1_POS_UPDATE = An Operation argument of SQL_UPDATE is supported in a call to SetPos when the cursor is a dynamic cursor.

SQL_CA1_POS_DELETE = An Operation argument of SQL_DELETE is supported in a call to SetPos when the cursor is a dynamic cursor.

SQL_CA1_POS_REFRESH = An Operation argument of SQL_REFRESH is supported in a call to SetPos when the cursor is a dynamic cursor.

SQL_CA1_POSITIONED_UPDATE = An UPDATE WHERE CURRENT OF SQL statement is supported when the cursor is a dynamic cursor. (An SQL-92 Entry level-conformant driver will always return this option as supported.)

SQL_CA1_POSITIONED_DELETE = A DELETE WHERE CURRENT OF SQL statement is supported when the cursor is a dynamic cursor. (An SQL-92 Entry level-conformant driver will always return this option as supported.)

SQL_CA1_SELECT_FOR_UPDATE = A SELECT FOR UPDATE SQL statement is supported when the cursor is a dynamic cursor. (An SQL-92 Entry level-conformant driver will always return this option as supported.)

SQL_CA1_BULK_ADD = An Operation argument of SQL_ADD is supported in a call to BulkOperations when the cursor is a dynamic cursor.

SQL_CA1_BULK_UPDATE_BY_BOOKMARK = An Operation argument of SQL_UPDATE_BY_BOOKMARK is supported in a call to BulkOperations when the cursor is a dynamic cursor.

SQL_CA1_BULK_DELETE_BY_BOOKMARK = An Operation argument of SQL_DELETE_BY_BOOKMARK is supported in a call to BulkOperations when the cursor is a dynamic cursor.

SQL_CA1_BULK_FETCH_BY_BOOKMARK = An Operation argument of SQL_FETCH_BY_BOOKMARK is supported in a call to BulkOperations when the cursor is a dynamic cursor.

An SQL-92 Intermediate level-conformant driver will usually return the SQL_CA1_NEXT, SQL_CA1_ABSOLUTE, and SQL_CA1_RELATIVE options as supported,  because the driver supports scrollable cursors through the embedded SQL FETCH statement. Because this does not directly determine the underlying SQL support, however, scrollable cursors may not be supported, even for an SQL-92 Intermediate level-conformant driver.

**"KeysetCursorAttributes2"**

Returns an SQLUINTEGER bitmask that describes the attributes of a keyset cursor that are supported by the driver. This bitmask contains the second subset of attributes; for the first subset, see KeysetCursorAttributes1.

The following bitmasks are used to determine which attributes are supported:

SQL_CA2_READ_ONLY_CONCURRENCY = A read-only dynamic cursor, in which no updates are allowed, is supported. (The SQL_ATTR_CONCURRENCY statement attribute can be SQL_CONCUR_READ_ONLY for a dynamic cursor).

SQL_CA2_LOCK_CONCURRENCY = A dynamic cursor that uses the lowest level of locking sufficient to ensure that the row can be updated is supported. (The SQL_ATTR_CONCURRENCY statement attribute can be SQL_CONCUR_LOCK for a dynamic cursor.) These locks must be consistent with the transaction isolation level set by the SQL_ATTR_TXN_ISOLATION connection attribute.

SQL_CA2_OPT_ROWVER_CONCURRENCY = A dynamic cursor that uses the optimistic  concurrency control comparing row versions is supported. (The SQL_ATTR_CONCURRENCY statement attribute can be SQL_CONCUR_ROWVER for a dynamic cursor.)

SQL_CA2_OPT_VALUES_CONCURRENCY = A dynamic cursor that uses the optimistic concurrency control comparing values is supported. (The SQL_ATTR_CONCURRENCY statement attribute can be SQL_CONCUR_VALUES for a dynamic cursor.)

SQL_CA2_SENSITIVITY_ADDITIONS = Added rows are visible to a dynamic cursor; the cursor can scroll to those rows. (Where these rows are added to the cursor is driver-dependent.)

SQL_CA2_SENSITIVITY_DELETIONS = Deleted rows are no longer available to a dynamic cursor, and do not leave a "hole" in the result set; after the dynamic cursor scrolls from a deleted row, it cannot return to that row.

SQL_CA2_SENSITIVITY_UPDATES = Updates to rows are visible to a dynamic cursor; if the dynamic cursor scrolls from and returns to an updated row, the data returned by the cursor is the updated data, not the original data.

SQL_CA2_MAX_ROWS_SELECT = The SQL_ATTR_MAX_ROWS statement attribute affects SELECT statements when the cursor is a dynamic cursor.

SQL_CA2_MAX_ROWS_INSERT = The SQL_ATTR_MAX_ROWS statement attribute affects INSERT statements when the cursor is a dynamic cursor.

SQL_CA2_MAX_ROWS_DELETE = The SQL_ATTR_MAX_ROWS statement attribute affects DELETE statements when the cursor is a dynamic cursor.

SQL_CA2_MAX_ROWS_UPDATE = The SQL_ATTR_MAX_ROWS statement attribute affects UPDATE statements when the cursor is a dynamic cursor.

SQL_CA2_MAX_ROWS_CATALOG = The SQL_ATTR_MAX_ROWS statement attribute affects CATALOG result sets when the cursor is a dynamic cursor.

SQL_CA2_MAX_ROWS_AFFECTS_ALL = The SQL_ATTR_MAX_ROWS statement attribute affects SELECT, INSERT, DELETE, and UPDATE statements, and CATALOG result sets, when the cursor is a dynamic cursor.

SQL_CA2_CRC_EXACT = The exact row count is available in the SQL_DIAG_CURSOR_ROW_COUNT diagnostic field when the cursor is a dynamic cursor.

SQL_CA2_CRC_APPROXIMATE = An approximate row count is available in the SQL_DIAG_CURSOR_ROW_COUNT diagnostic field when the cursor is a dynamic cursor.

SQL_CA2_SIMULATE_NON_UNIQUE = The driver does not guarantee that simulated positioned update or delete statements will affect only one row when the cursor is a dynamic cursor; it is the application's responsibility to guarantee this. (If a statement affects more than one row, Execute or ExecDirect returns SQLSTATE 01001 [Cursor operation conflict].) To set this behavior, the application calls SetStmtAttr with the SQL_ATTR_SIMULATE_CURSOR attribute set to SQL_SC_NON_UNIQUE.

SQL_CA2_SIMULATE_TRY_UNIQUE = The driver attempts to guarantee that simulated positioned update or delete statements will affect only one row when the cursor is a dynamic cursor. The driver always executes such statements, even if they might affect more than one row, such as when there is no unique key. (If a statement affects more than one row, Execute or ExecDirect returns SQLSTATE 01001 [Cursor operation conflict].) To set this behavior, the application calls SetStmtAttr with the SQL_ATTR_SIMULATE_CURSOR attribute set to SQL_SC_TRY_UNIQUE.

SQL_CA2_SIMULATE_UNIQUE = The driver guarantees that simulated positioned update or delete statements will affect only one row when the cursor is a dynamic cursor. If the driver cannot guarantee this for a given statement, ExecDirect or Prepare return SQLSTATE 01001 (Cursor operation conflict). To set this behavior, the application calls SetStmtAttr with the SQL_ATTR_SIMULATE_CURSOR attribute set to SQL_SC_UNIQUE.

**"Keywords"**

Returns a character string containing a comma-separated list of all data source-specific keywords. This list does not contain keywords specific to ODBC or keywords used by both the data source and ODBC. This list represents all the reserved keywords; interoperable applications should not use these words in object names.

**"LikeEscapeClause"**

Returns a character string: "Y" if the data source supports an escape character for the percent character (%) and underscore character (_) in a LIKE predicate and the driver supports the ODBC syntax for defining a LIKE predicate escape character; "N" otherwise.

**"MaxAsyncConcurrentStatements"**

Returns an SQLUINTEGER value specifying the maximum number of active concurrent statements in asynchronous mode that the driver can support on a given connection. If there is no specific limit or the limit is unknown, this value is zero.

**"MaxBinaryLiteralLen"**

Returns an SQLUINTEGER value specifying the maximum length (number of hexadecimal characters, excluding the literal prefix and suffix returned by GetTypeInfo) of a binary literal in an SQL statement. For example, the binary literal 0xFFAA has a length of 4. If there is no maximum length or the length is unknown, this value is set to zero.

**"MaxCatalogNameLen"**

Returns an SQLUSMALLINT value specifying the maximum length of a catalog name in the data source. If there is no maximum length or the length is unknown, this value is set to zero.

**"MaxCharLiteralLen"**

Returns an SQLUINTEGER value specifying the maximum length (number of characters, excluding the literal prefix and suffix returned by GetTypeInfo) of a character literal in an SQL statement. If there is no maximum length or the length is unknown, this value is set to zero.

**"MaxColumnNameLen"**

Returns an SQLUSMALLINT value specifying the maximum length of a column name in the data source. If there is no maximum length or the length is unknown, this value is set to zero.

**"MaxColumnsInGroupBy"**

Returns an SQLUSMALLINT value specifying the maximum number of columns allowed in a GROUP BY clause. If there is no specified limit or the limit is unknown, this value is set to zero.

**"MaxColumnsInIndex"**

Returns an SQLUSMALLINT value specifying the maximum number of columns allowed in an index. If there is no specified limit or the limit is unknown, this value is set to zero.

**"MaxColumnsInOrderBy"**

Returns an SQLUSMALLINT value specifying the maximum number of columns allowed in an ORDER BY clause. If there is no specified limit or the limit is unknown, this value is set to zero.

**"MaxColumnsInSelect"**

Returns an SQLUSMALLINT value specifying the maximum number of columns allowed in a select list. If there is no specified limit or the limit is unknown, this value is set to zero.

**"MaxColumnsInTable"**

Returns an SQLUSMALLINT value specifying the maximum number of columns allowed in a table. If there is no specified limit or the limit is unknown, this value is set to zero.

**"MaxConcurrentActivities"**

Returns an SQLUSMALLINT value specifying the maximum number of active statements that the driver can support for a connection.

**"MaxCursorNameLen"**

Returns an SQLUSMALLINT value specifying the maximum length of a cursor name in the data source. If there is no maximum length or the length is unknown, this value is set to zero.

**"MaxDriverConnections"**

Returns an SQLUSMALLINT value specifying the maximum number of active connections that the driver can support for an environment. This value can reflect a limitation imposed by either the driver or the data source. If there is no specified limit or the limit is unknown, this value is set to zero.

**"MaxIdentifierLen"**

Returns an SQLUSMALLINT that indicates the maximum size in characters that the data source supports for user-defined names.

**"MaxIndexSize"**

Returns an SQLUINTEGER value specifying the maximum number of bytes allowed in the combined fields of an index. If there is no specified limit or the limit is unknown, this value is set to zero.

**"MaxProcedureNameLen"**

Returns an SQLUSMALLINT value specifying the maximum length of a procedure  name in the data source. If there is no maximum length or the length is unknown, this value is set to zero.

**"MaxRowSize"**

Returns an SQLUINTEGER value specifying the maximum length of a single row in a table. If there is no specified limit or the limit is unknown, this value is set to zero.

**"MaxRowSizeIncludesLong"**

Returns a character string: "Y" if the maximum row size returned for the SQL_MAX_ROW_SIZE information type includes the length of all SQL_LONGVARCHAR and SQL_LONGVARBINARY columns in the row; "N" otherwise.

**"MaxSchemaNameLen"**

Returns an SQLUSMALLINT value specifying the maximum length of a schema name in the data source. If there is no maximum length or the length is unknown, this value is set to zero.

**"MaxStatementLen"**

Returns an SQLUINTEGER value specifying the maximum length (number of characters, including white space) of an SQL statement. If there is no maximum length or the length is unknown, this value is set to zero.

**"MaxTablenameLen"**

Returns an SQLUSMALLINT value specifying the maximum length of a table name in the data source. If there is no maximum length or the length is unknown, this value is set to zero.

**"MaxTablesInSelect"**

Returns an SQLUSMALLINT value specifying the maximum number of tables allowed in the FROM clause of a SELECT statement. If there is no specified limit or the limit is unknown, this value is set to zero.

**"MaxUserNameLen"**

Returns an SQLUSMALLINT value specifying the maximum length of a user name in the data source. If there is no maximum length or the length is unknown, this value is set to zero.

**"MultipleActiveTxn"**

Returns a character string: "Y" if the driver supports more than one active transaction at the same time, "N" if only one transaction can be active at any time.

**"MultResultSets"**

Returns a character string: "Y" if the data source supports multiple result sets, "N" if it does not.

**"NeedLongDataLen"**

Returns a character string: "Y" if the data source needs the length of a long data value (the data type is SQL_LONGVARCHAR, SQL_LONGVARBINARY, or  a long data source-specific data type) before that value is sent to the data source, "N" if it does not. For more information, see BindParameter and SetPos.

**"NonNullableColumns"**

Returns an SQLUSMALLINT value specifying whether the data source supports NOT NULL in columns:

SQL_NNC_NULL = All columns must be nullable.
SQL_NNC_NON_NULL = Columns cannot be nullable. (The data source supports the NOT NULL column constraint in CREATE TABLE statements.)

An SQL-92 Entry level-conformant driver will return SQL_NNC_NON_NULL.

**"NullCollation"**

Returns an SQLUSMALLINT value specifying where NULLs are sorted in a result set.

SQL_NC_END = NULLs are sorted at the end of the result set, regardless of the ASC or DESC keywords.

SQL_NC_HIGH = NULLs are sorted at the high end of the result set, depending on the ASC or DESC keywords.

SQL_NC_LOW = NULLs are sorted at the low end of the result set, depending on the ASC or DESC keywords.

SQL_NC_START = NULLs are sorted at the start of the result set, regardless of the ASC or DESC keywords.

**"NumericFunctions"**

Returns an SQLUINTEGER bitmask enumerating the scalar numeric functions supported by the driver and associated data source.

**"ODBCInterfaceConformance"**

Returns an SQLUINTEGER value indicating the level of the ODBC 3.x interface that the driver conforms to.

SQL_OIC_CORE: The minimum level that all ODBC drivers are expected to conform to. This level includes basic interface elements such as connection functions, functions for preparing and executing an SQL statement, basic result set metadata functions, basic catalog functions, and so on.

SQL_OIC_LEVEL1: A level including the core standards compliance level functionality, plus scrollable cursors, bookmarks, positioned updates and deletes, and so on.

SQL_OIC_LEVEL2: A level including level 1 standards compliance level functionality, plus advanced features such as sensitive cursors; update, delete, and refresh by bookmarks; stored procedure support; catalog functions for primary and foreign keys; multicatalog support; and so on.

**"ODBCVer"**

Returns a character string with the version of ODBC to which the Driver Manager conforms. The version is of the form ##.##.0000, where the first two digits are the major version and the next two digits are the minor version. This is implemented solely in the Driver Manager.

**"OJCapabilities"**

Returns an SQLUINTEGER bitmask enumerating the types of outer joins supported by the driver and data source.

**"OrderByColumnsInSelect"**

Returns a character string: "Y" if the columns in the ORDER BY clause must be in the select list; otherwise, "N".

**"OuterJoins"**

Returns a character string indicating the level of outer joins support.

"N" = No. The data source does not support outer joins. (ODBC 1.0)

"Y" = Yes. The data source supports two-table outer joins, and the driver supports the ODBC outer join syntax except for nested outer joins. However, columns on the left side of the comparison operator in the ON clause must come from the left-hand table in the outer join, and columns on the right side of the comparison operator must come from the right-hand table. (ODBC 1.0)

"P" = Partial. The data source partially supports nested outer joins, and the driver supports the ODBC outer join syntax. However, columns on the left side of the comparison operator in the ON clause must come from the left-hand table in the outer join and columns on the right side of the comparison operator must come from the right-hand table. Also, the right-hand table of an outer join cannot be included in an inner join. (ODBC 2.0)

"F" = Full. The data source fully supports nested outer joins, and the driver supports the ODBC outer join syntax. (ODBC 2.0)

**"ParamArrayRowCounts"**

Returns an SQLUINTEGER enumerating the driver's properties regarding the availability of row counts in a parameterized execution.

SQL_PARC_BATCH = Individual row counts are available for each set of parameters. This is conceptually equivalent to the driver generating a batch of SQL statements, one for each parameter set in the array. Extended error information can be retrieved by using the SQL_PARAM_STATUS_PTR descriptor field.

SQL_PARC_NO_BATCH = There is only one row count available, which is the cumulative row count resulting from the execution of the statement for the entire array of parameters. This is conceptually equivalent to treating the statement along with the entire parameter array as one atomic unit. Errors  are handled the same as if one statement were executed.

**"ParamArraySelects"**

Returns an SQLUINTEGER enumerating the driver's properties regarding the availability of result sets in a parameterized execution.

SQL_PAS_BATCH = There is one result set available per set of parameters. This is conceptually equivalent to the driver generating a batch of SQL statements, one for each parameter set in the array.

SQL_PAS_NO_BATCH = There is only one result set available, which represents the cumulative result set resulting from the execution of the statement for the entire array of parameters. This is conceptually equivalent to treating the statement along with the entire parameter array as one atomic unit.

SQL_PAS_NO_SELECT = A driver does not allow a result-set generating statement to be executed with an array of parameters.

**"PosOperations"**

Returns an SQLINTEGER bitmask enumerating the support operations in SetPos.

The following bitmasks are used in conjunction with the flag to determine which options are supported.

SQL_POS_POSITION (ODBC 2.0) SQL_POS_REFRESH (ODBC 2.0)

SQL_POS_UPDATE (ODBC 2.0) SQL_POS_DELETE (ODBC 2.0) SQL_POS_ADD (ODBC 2.0)

**"Procedures"**

A character string: "Y" if the data source supports procedures and the driver supports the ODBC procedure invocation syntax; "N" otherwise.

**"ProcedureTerm"**

Returns a character string with the data source vendor's name for a procedure; for example, "database procedure", "stored procedure", "procedure", "package", or "stored query".

**"QuotedIdentifierCase"**

Returns an SQLUSMALLINT value with the supported quoted identifier case.

SQL_IC_UPPER = Quoted identifiers in SQL are not case-sensitive and are stored in uppercase in the system catalog.

SQL_IC_LOWER = Quoted identifiers in SQL are not case-sensitive and are stored in lowercase in the system catalog.

SQL_IC_SENSITIVE = Quoted identifiers in SQL are case-sensitive and are stored in mixed case in the system catalog. (In an SQL-92-compliant database, quoted identifiers are always case-sensitive.)

SQL_IC_MIXED = Quoted identifiers in SQL are not case-sensitive and are stored in mixed case in the system catalog.

An SQL-92 Entry level-conformant driver will always return SQL_IC_SENSITIVE.

**"RowUpdates"**

Returns a character string: "Y" if a keyset-driven or mixed cursor maintains row versions or values for all fetched rows and therefore can detect any updates made to a row by any user since the row was last fetched. (This applies only to updates, not to deletions or insertions.) The driver can return the SQL_ROW_UPDATED flag to the row status array when FetchScroll is called. Otherwise, "N".

**"SchemaTerm"**

Returns a character string with the data source vendor's name for an schema; for example, "owner", "Authorization ID", or "Schema". The character string can be returned in upper, lower, or mixed case.

An SQL-92 Entry level–conformant driver will always return "schema". This InfoType has been renamed for ODBC 3.0 from the ODBC 2.0 InfoType SQL_OWNER_TERM.

**"SchemaUsage"**

Returns an SQLUINTEGER bitmask enumerating the statements in which schemas can be used.

SQL_SU_DML_STATEMENTS = Schemas are supported in all Data Manipulation Language statements: SELECT, INSERT, UPDATE, DELETE, and if supported, SELECT FOR UPDATE and positioned update and delete statements.

SQL_SU_PROCEDURE_INVOCATION = Schemas are supported in the ODBC procedure invocation statement.

SQL_SU_TABLE_DEFINITION = Schemas are supported in all table definition statements: CREATE TABLE, CREATE VIEW, ALTER TABLE, DROP TABLE, and DROP VIEW.

SQL_SU_INDEX_DEFINITION = Schemas are supported in all index definition statements: CREATE INDEX and DROP INDEX.

SQL_SU_PRIVILEGE_DEFINITION = Schemas are supported in all privilege definition statements: GRANT and REVOKE. An SQL-92 Entry level-conformant driver will always return the SQL_SU_DML_STATEMENTS, SQL_SU_TABLE_DEFINITION, and SQL_SU_PRIVILEGE_DEFINITION options, as supported.

This InfoType has been renamed for ODBC 3.0 from the ODBC 2.0 InfoType SQL_OWNER_USAGE.

**"ScrollOptions"**

Returns an SQLUINTEGER bitmask enumerating the scroll options supported for scrollable cursors.

The following bitmasks are used to determine which options are supported:

SQL_SO_FORWARD_ONLY = The cursor only scrolls forward. (ODBC 1.0)

SQL_SO_STATIC = The data in the result set is static. (ODBC 2.0)

SQL_SO_KEYSET_DRIVEN = The driver saves and uses the keys for every row in the result set. (ODBC 1.0)

SQL_SO_DYNAMIC = The driver keeps the keys for every row in the rowset (the keyset size is the same as the rowset size). (ODBC 1.0)

SQL_SO_MIXED = The driver keeps the keys for every row in the keyset, and  the keyset size is greater than the rowset size. The cursor is keyset-driven inside the keyset and dynamic outside the keyset. (ODBC 1.0)

**"SearchPatternEscape"**

Returns a character string specifying what the driver supports as an escape character that permits the use of the pattern match metacharacters underscore (\_) and percent sign (%) as valid characters in search patterns. This escape character applies only for those catalog function arguments that support search strings. If this string is empty, the driver does not support a search-pattern escape character. Because this information type does not indicate general support of the escape character in the LIKE predicate, SQL-92 does not include requirements for this character string.

This info type is limited to catalog functions.

**"ServerName"**

Returns a character string with the actual data source-specific server name; useful when a data source name is used during Connect, DriverConnect, and BrowseConnect.

**"SpecialCharacters"**

Returns a character string containing all special characters (that is, all characters except a through z, A through Z, 0 through 9, and underscore) that can be used in an identifier name, such as a table name, column column name, or index name, on the data source. For example, "#$^". If an identifier contains one or more of these characters, the identifier must be a delimited identifier.

**"SQL92DateTimeFunctions"**

Returns an SQLUINTEGER bitmask enumerating the datetime scalar functions that are supported by the driver and the associated data source, as defined in SQL-92.

The following bitmasks are used to determine which datetime functions are supported:

SQL_SDF_CURRENT_DATE<br>
SQL_SDF_CURRENT_TIME<br>
SQL_SDF_CURRENT_TIMESTAMP

**"SQL92ForeignKeyDeleteRule"**

Returns an SQLUINTEGER bitmask enumerating the rules supported for a foreign key in a DELETE statement, as defined in SQL-92.

The following bitmasks are used to determine which clauses are supported by the data source:

SQL_SFKD_CASCADE<br>
SQL_SFKD_NO_ACTION<br>
SQL_SFKD_SET_DEFAULT<br>
SQL_SFKD_SET_NULL

An FIPS Transitional level-conformant driver will always return all of these options as supported.

**"SQL92ForeignKeyUpdateRule"**

Returns an SQLUINTEGER bitmask enumerating the rules supported for a foreign key in an UPDATE statement, as defined in SQL-92.
The following bitmasks are used to determine which clauses are supported by the data source:

SQL_SFKU_CASCADE<br>
SQL_SFKU_NO_ACTION<br>
SQL_SFKU_SET_DEFAULT<br>
SQL_SFKU_SET_NULL

An SQL-92 Full level-conformant driver will always return all of these options as supported.

**"SQL92Grant"**

Returns an SQLUINTEGER bitmask enumerating the clauses supported in the GRANT statement, as defined in SQL-92.

The SQL-92 or FIPS conformance level at which this feature needs to be supported is shown in parentheses next to each bitmask.

The following bitmasks are used to determine which clauses are supported by the data source:

SQL_SG_DELETE_TABLE (Entry level)<br>
SQL_SG_INSERT_COLUMN (Intermediate level)<br>
SQL_SG_INSERT_TABLE (Entry level)<br>
SQL_SG_REFERENCES_TABLE (Entry level)<br>
SQL_SG_REFERENCES_COLUMN (Entry level)<br>
SQL_SG_SELECT_TABLE (Entry level)<br>
SQL_SG_UPDATE_COLUMN (Entry level)<br>
SQL_SG_UPDATE_TABLE (Entry level)<br>
SQL_SG_USAGE_ON_DOMAIN (FIPS Transitional level)<br>
SQL_SG_USAGE_ON_CHARACTER_SET (FIPS Transitional level)<br>
SQL_SG_USAGE_ON_COLLATION (FIPS Transitional level)<br>
SQL_SG_USAGE_ON_TRANSLATION (FIPS Transitional level)<br>
SQL_SG_WITH_GRANT_OPTION (Entry level)

**"SQL92NumericValueFunctions"**

Returns an SQLUINTEGER bitmask enumerating the numeric value scalar functions that are supported by the driver and the associated data source, as defined in SQL-92.
The following bitmasks are used to determine which numeric functions are supported:

SQL_SNVF_BIT_LENGTH<br>
SQL_SNVF_CHAR_LENGTH<br>
SQL_SNVF_CHARACTER_LENGTH<br>
SQL_SNVF_EXTRACT<br>
SQL_SNVF_OCTET_LENGTH<br>
SQL_SNVF_POSITION

**"SQL92Predicates"**

Returns an SQLUINTEGER bitmask enumerating the predicates supported in a SELECT statement, as defined in SQL-92.

The SQL-92 or FIPS conformance level at which this feature needs to be supported is shown in parentheses next to each bitmask.

The following bitmasks are used to determine which options are supported by the data source:

SQL_SP_BETWEEN (Entry level)<br>
SQL_SP_COMPARISON (Entry level)<br>
SQL_SP_EXISTS (Entry level)<br>
SQL_SP_IN (Entry level)<br>
SQL_SP_ISNOTNULL (Entry level)<br>
SQL_SP_ISNULL (Entry level)<br>
SQL_SP_LIKE (Entry level)<br>
SQL_SP_MATCH_FULL (Full level)<br>
SQL_SP_MATCH_PARTIAL(Full level)<br>
SQL_SP_MATCH_UNIQUE_FULL (Full level)<br>
SQL_SP_MATCH_UNIQUE_PARTIAL (Full level)<br>
SQL_SP_OVERLAPS (FIPS Transitional level)<br>
SQL_SP_QUANTIFIED_COMPARISON (Entry level)<br>
SQL_SP_UNIQUE (Entry level)

**"SQL92RelationalJoinOperators"**

Returns an SQLUINTEGER bitmask enumerating the relational join operators supported in a SELECT statement, as defined in SQL-92.

The SQL-92 or FIPS conformance level at which this feature needs to be supported is shown in parentheses next to each bitmask.

The following bitmasks are used to determine which options are supported by the data source:

SQL_SRJO_CORRESPONDING_CLAUSE (Intermediate level)<br>
SQL_SRJO_CROSS_JOIN (Full level)<br>
SQL_SRJO_EXCEPT_JOIN (Intermediate level)<br>
SQL_SRJO_FULL_OUTER_JOIN (Intermediate level)<br>
SQL_SRJO_INNER_JOIN (FIPS Transitional level)<br>
SQL_SRJO_INTERSECT_JOIN (Intermediate level)<br>
SQL_SRJO_LEFT_OUTER_JOIN (FIPS Transitional level)<br>
SQL_SRJO_NATURAL_JOIN (FIPS Transitional level)<br>
SQL_SRJO_RIGHT_OUTER_JOIN (FIPS Transitional level)<br>
SQL_SRJO_UNION_JOIN (Full level)

SQL_SRJO_INNER_JOIN indicates support for the INNER JOIN syntax, not for the inner join capability. Support for the INNER JOIN syntax is FIPS TRANSITIONAL, while support for the inner join capability is ENTRY.

**"SQL92Revoke"**

Returns an SQLUINTEGER bitmask enumerating the clauses supported in the REVOKE statement, as defined in SQL-92, supported by the data source.

The SQL-92 or FIPS conformance level at which this feature needs to be supported is shown in parentheses next to each bitmask.

The following bitmasks are used to determine which clauses are supported by the data source:

SQL_SR_CASCADE (FIPS Transitional level)<br>
SQL_SR_DELETE_TABLE (Entry level)<br>
SQL_SR_GRANT_OPTION_FOR (Intermediate level)<br>
SQL_SR_INSERT_COLUMN (Intermediate level)<br>
SQL_SR_INSERT_TABLE (Entry level)<br>
SQL_SR_REFERENCES_COLUMN (Entry level)<br>
SQL_SR_REFERENCES_TABLE (Entry level)<br>
SQL_SR_RESTRICT (FIPS Transitional level)<br>
SQL_SR_SELECT_TABLE (Entry level)<br>
SQL_SR_UPDATE_COLUMN (Entry level)<br>
SQL_SR_UPDATE_TABLE (Entry level)<br>
SQL_SR_USAGE_ON_DOMAIN (FIPS Transitional level)<br>
SQL_SR_USAGE_ON_CHARACTER_SET (FIPS Transitional level)<br>
SQL_SR_USAGE_ON_COLLATION (FIPS Transitional level)<br>
SQL_SR_USAGE_ON_TRANSLATION (FIPS Transitional level)

**"SQL92RowValueConstructor"**

Returns an SQLUINTEGER bitmask enumerating the row value constructor expressions supported in a SELECT statement, as defined in SQL-92.

The following bitmasks are used to determine which options are supported by the data source:

SQL_SRVC_VALUE_EXPRESSION<br>
SQL_SRVC_NULL<br>
SQL_SRVC_DEFAULT<br>
SQL_SRVC_ROW_SUBQUERY

**"SQL92StringFunctions"**

Returns an SQLUINTEGER bitmask enumerating the string scalar functions that are supported by the driver and the associated data source, as defined in SQL-92.

The following bitmasks are used to determine which string functions are supported:

SQL_SSF_CONVERT<br>
SQL_SSF_LOWER<br>
SQL_SSF_UPPER<br>
SQL_SSF_SUBSTRING<br>
SQL_SSF_TRANSLATE<br>
SQL_SSF_TRIM_BOTH<br>
SQL_SSF_TRIM_LEADING<br>
SQL_SSF_TRIM_TRAILING

**"SQL92ValueExpressions"**

Returns an SQLUINTEGER bitmask enumerating the value expressions supported, as defined in SQL-92.

The SQL-92 or FIPS conformance level at which this feature needs to be supported is shown in parentheses next to each bitmask.

The following bitmasks are used to determine which options are supported by the data source:

SQL_SVE_CASE (Intermediate level)<br>
SQL_SVE_CAST (FIPS Transitional level)<br>
SQL_SVE_COALESCE (Intermediate level)<br>
SQL_SVE_NULLIF (Intermediate level)

**"SqlConformance"**

Returns an SQLUINTEGER value indicating the level of SQL-92 supported by the driver.

SQL_SC_SQL92_ENTRY = Entry level SQL-92 compliant.<br>
SQL_SC_FIPS127_2_TRANSITIONAL = FIPS 127-2 transitional level compliant.<br>
SQL_SC_SQL92_FULL = Full level SQL-92 compliant.<br>
SQL_SC_ SQL92_INTERMEDIATE = Intermediate level SQL-92 compliant.

**"StandardCliConformance"**

Returns an SQLUINTEGER bitmask enumerating the CLI standard or standards to which the driver conforms. The following bitmasks are used to determine which levels the driver conforms to:

SQL_SCC_XOPEN_CLI_VERSION1: The driver conforms to the X/Open CLI version 1.

SQL_SCC_ISO92_CLI: The driver conforms to the ISO 92 CLI.

**"StaticCursorAttributes1"**

Returns an SQLUINTEGER bitmask that describes the attributes of a static cursor that are supported by the driver. This bitmask contains the first subset of attributes; for the second subset, see StaticCursorAttributes2.

The following bitmasks are used to determine which attributes are supported:

SQL_CA1_NEXT = A FetchOrientation argument of SQL_FETCH_NEXT is supported in a call to FetchScroll when the cursor is a dynamic cursor.

SQL_CA1_ABSOLUTE = FetchOrientation arguments of SQL_FETCH_FIRST, SQL_FETCH_LAST, and SQL_FETCH_ABSOLUTE are supported in a call to FetchScroll when the cursor is a dynamic cursor. (The rowset that will be fetched is independent of the current cursor position.)

SQL_CA1_RELATIVE = FetchOrientation arguments of SQL_FETCH_PRIOR and SQL_FETCH_RELATIVE are supported in a call to FetchScroll when the cursor is a dynamic cursor. (The rowset that will be fetched is dependent on the current cursor position. Note that this is separated from SQL_FETCH_NEXT because in a forward-only cursor, only SQL_FETCH_NEXT is supported.)

SQL_CA1_BOOKMARK = A FetchOrientation argument of SQL_FETCH_BOOKMARK is supported in a call to FetchScroll when the cursor is a dynamic cursor.

SQL_CA1_LOCK_EXCLUSIVE = A LockType argument of SQL_LOCK_EXCLUSIVE is supported in a call to SetPos when the cursor is a dynamic cursor.

SQL_CA1_LOCK_NO_CHANGE = A LockType argument of SQL_LOCK_NO_CHANGE is supported in a call to SetPos when the cursor is a dynamic cursor.

SQL_CA1_LOCK_UNLOCK = A LockType argument of SQL_LOCK_UNLOCK is supported in a call to SetPos when the cursor is a dynamic cursor.

SQL_CA1_POS_POSITION = An Operation argument of SQL_POSITION is supported  in a call to SetPos when the cursor is a dynamic cursor.

SQL_CA1_POS_UPDATE = An Operation argument of SQL_UPDATE is supported in a call to SetPos when the cursor is a dynamic cursor.

SQL_CA1_POS_DELETE = An Operation argument of SQL_DELETE is supported in a call to SetPos when the cursor is a dynamic cursor.

SQL_CA1_POS_REFRESH = An Operation argument of SQL_REFRESH is supported in a call to SetPos when the cursor is a dynamic cursor.

SQL_CA1_POSITIONED_UPDATE = An UPDATE WHERE CURRENT OF SQL statement is supported when the cursor is a dynamic cursor. (An SQL-92 Entry level-conformant driver will always return this option as supported.)

SQL_CA1_POSITIONED_DELETE = A DELETE WHERE CURRENT OF SQL statement is supported when the cursor is a dynamic cursor. (An SQL-92 Entry level-conformant driver will always return this option as supported.)

SQL_CA1_SELECT_FOR_UPDATE = A SELECT FOR UPDATE SQL statement is supported when the cursor is a dynamic cursor. (An SQL-92 Entry level-conformant driver will always return this option as supported.)

SQL_CA1_BULK_ADD = An Operation argument of SQL_ADD is supported in a call to BulkOperations when the cursor is a dynamic cursor.

SQL_CA1_BULK_UPDATE_BY_BOOKMARK = An Operation argument of SQL_UPDATE_BY_BOOKMARK is supported in a call to BulkOperations when the cursor is a dynamic cursor.

SQL_CA1_BULK_DELETE_BY_BOOKMARK = An Operation argument of SQL_DELETE_BY_BOOKMARK is supported in a call to BulkOperations when the cursor is a dynamic cursor.

SQL_CA1_BULK_FETCH_BY_BOOKMARK = An Operation argument of SQL_FETCH_BY_BOOKMARK is supported in a call to BulkOperations when the cursor is a dynamic cursor.

An SQL-92 Intermediate level-conformant driver will usually return the SQL_CA1_NEXT, SQL_CA1_ABSOLUTE, and SQL_CA1_RELATIVE options as supported, because it supports scrollable cursors through the embedded %SQL FETCH statement. Because this does not directly determine the underlying SQL support, however, scrollable cursors may not be supported, even for an SQL-92 Intermediate level-conformant driver.

**"StaticCursorAttributes2"**

Returns an SQLUINTEGER bitmask that describes the attributes of a static cursor that are supported by the driver. This bitmask contains the second subset of attributes; for the first subset, see StaticCursorAttributes1.

The following bitmasks are used to determine which attributes are supported:

SQL_CA2_READ_ONLY_CONCURRENCY = A read-only dynamic cursor, in which no updates are allowed, is supported. (The SQL_ATTR_CONCURRENCY statement attribute can be SQL_CONCUR_READ_ONLY for a dynamic cursor).

SQL_CA2_LOCK_CONCURRENCY = A dynamic cursor that uses the lowest level of locking sufficient to ensure that the row can be updated is supported. (The SQL_ATTR_CONCURRENCY statement attribute can be SQL_CONCUR_LOCK for a dynamic cursor.) These locks must be consistent with the transaction isolation level set by the SQL_ATTR_TXN_ISOLATION connection attribute.

SQL_CA2_OPT_ROWVER_CONCURRENCY = A dynamic cursor that uses the optimistic  concurrency control comparing row versions is supported. (The SQL_ATTR_CONCURRENCY statement attribute can be SQL_CONCUR_ROWVER for a dynamic cursor.)

SQL_CA2_OPT_VALUES_CONCURRENCY = A dynamic cursor that uses the optimistic concurrency control comparing values is supported. (The SQL_ATTR_CONCURRENCY statement attribute can be SQL_CONCUR_VALUES for a dynamic cursor.)

SQL_CA2_SENSITIVITY_ADDITIONS = Added rows are visible to a dynamic cursor; the cursor can scroll to those rows. (Where these rows are added to the cursor is driver-dependent.)

SQL_CA2_SENSITIVITY_DELETIONS = Deleted rows are no longer available to a dynamic cursor, and do not leave a "hole" in the result set; after the dynamic cursor scrolls from a deleted row, it cannot return to that row.

SQL_CA2_SENSITIVITY_UPDATES = Updates to rows are visible to a dynamic cursor; if the dynamic cursor scrolls from and returns to an updated row, the data returned by the cursor is the updated data, not the original data.

SQL_CA2_MAX_ROWS_SELECT = The SQL_ATTR_MAX_ROWS statement attribute affects SELECT statements when the cursor is a dynamic cursor.

SQL_CA2_MAX_ROWS_INSERT = The SQL_ATTR_MAX_ROWS statement attribute affects INSERT statements when the cursor is a dynamic cursor.

SQL_CA2_MAX_ROWS_DELETE = The SQL_ATTR_MAX_ROWS statement attribute affects DELETE statements when the cursor is a dynamic cursor.

SQL_CA2_MAX_ROWS_UPDATE = The SQL_ATTR_MAX_ROWS statement attribute affects UPDATE statements when the cursor is a dynamic cursor.

SQL_CA2_MAX_ROWS_CATALOG = The SQL_ATTR_MAX_ROWS statement attribute affects CATALOG result sets when the cursor is a dynamic cursor.

SQL_CA2_MAX_ROWS_AFFECTS_ALL = The SQL_ATTR_MAX_ROWS statement attribute affects SELECT, INSERT, DELETE, and UPDATE statements, and CATALOG result sets, when the cursor is a dynamic cursor.

SQL_CA2_CRC_EXACT = The exact row count is available in the SQL_DIAG_CURSOR_ROW_COUNT diagnostic field when the cursor is a dynamic cursor.

SQL_CA2_CRC_APPROXIMATE = An approximate row count is available in the SQL_DIAG_CURSOR_ROW_COUNT diagnostic field when the cursor is a dynamic cursor.

SQL_CA2_SIMULATE_NON_UNIQUE = The driver does not guarantee that simulated positioned update or delete statements will affect only one row when the cursor is a dynamic cursor; it is the application's responsibility to guarantee this. (If a statement affects more than one row, Execute or ExecDirect returns SQLSTATE 01001 [Cursor operation conflict].) To set this behavior, the application calls SetStmtAttr with the SQL_ATTR_SIMULATE_CURSOR attribute set to SQL_SC_NON_UNIQUE.

SQL_CA2_SIMULATE_TRY_UNIQUE = The driver attempts to guarantee that simulated positioned update or delete statements will affect only one row when the cursor is a dynamic cursor. The driver always executes such statements, even if they might affect more than one row, such as when there is no unique key. (If a statement affects more than one row, Execute or ExecDirect returns SQLSTATE 01001 [Cursor operation conflict].) To set this behavior, the application calls SetStmtAttr with the SQL_ATTR_SIMULATE_CURSOR attribute set to SQL_SC_TRY_UNIQUE.

SQL_CA2_SIMULATE_UNIQUE = The driver guarantees that simulated positioned update or delete statements will affect only one row when the cursor is a dynamic cursor. If the driver cannot guarantee this for a given statement, ExecDirect or Prepare return SQLSTATE 01001 (Cursor operation conflict). To set this behavior, the application calls SetStmtAttr with the SQL_ATTR_SIMULATE_CURSOR attribute set to SQL_SC_UNIQUE.

**"StringFunctions"**

Returns an SQLUINTEGER bitmask enumerating the scalar string functions supported by the driver and associated data source.

The following bitmasks are used to determine which string functions are supported:

SQL_FN_STR_ASCII (ODBC 1.0)<br>
SQL_FN_STR_BIT_LENGTH (ODBC 3.0)<br>
SQL_FN_STR_CHAR (ODBC 1.0)<br>
SQL_FN_STR_CHAR_LENGTH (ODBC 3.0)<br>
SQL_FN_STR_CHARACTER_LENGTH (ODBC 3.0)<br>
SQL_FN_STR_CONCAT (ODBC 1.0)<br>
SQL_FN_STR_DIFFERENCE (ODBC 2.0)<br>
SQL_FN_STR_INSERT (ODBC 1.0)<br>
SQL_FN_STR_LCASE (ODBC 1.0)<br>
SQL_FN_STR_LEFT (ODBC 1.0)<br>
SQL_FN_STR_LENGTH (ODBC 1.0)<br>
SQL_FN_STR_LOCATE (ODBC 1.0)<br>
SQL_FN_STR_LTRIM (ODBC 1.0)<br>
SQL_FN_STR_OCTET_LENGTH (ODBC 3.0)<br>
SQL_FN_STR_POSITION (ODBC 3.0)<br>
SQL_FN_STR_REPEAT (ODBC 1.0)<br>
SQL_FN_STR_REPLACE (ODBC 1.0)<br>
SQL_FN_STR_RIGHT (ODBC 1.0)<br>
SQL_FN_STR_RTRIM (ODBC 1.0)<br>
SQL_FN_STR_SOUNDEX (ODBC 2.0)<br>
SQL_FN_STR_SPACE (ODBC 2.0)<br>
SQL_FN_STR_SUBSTRING (ODBC 1.0)<br>
SQL_FN_STR_UCASE (ODBC 1.0)

If an application can call the LOCATE scalar function with the string_exp1, string_exp2, and start arguments, the driver returns the SQL_FN_STR_LOCATE bitmask. If an application can call the LOCATE scalar function with only the string_exp1 and string_exp2 arguments, the driver returns the SQL_FN_STR_LOCATE_2 bitmask. Drivers that fully support the LOCATE scalar function return both bitmasks.

**"Subqueries"**

Returns an SQLUINTEGER bitmask enumerating the predicates that support subqueries:

SQL_SQ_CORRELATED_SUBQUERIES<br>
SQL_SQ_COMPARISON<br>
SQL_SQ_EXISTS<br>
SQL_SQ_IN<br>
SQL_SQ_QUANTIFIED

The SQL_SQ_CORRELATED_SUBQUERIES bitmask indicates that all predicates that  support subqueries support correlated subqueries.

An SQL-92 Entry level-conformant driver will always return a bitmask in which all of these bits are set.

**"SystemFunctions"**

Returns an SQLUINTEGER bitmask enumerating the scalar system functions supported by the driver and associated data source.

The following bitmasks are used to determine which system functions are supported:

SQL_FN_SYS_DBNAME<br>
SQL_FN_SYS_IFNULL<br>
SQL_FN_SYS_USERNAME<br>

**"TableTerm"**

Returns a character string with the data source vendor's name for a table; for example, "table" or "file".

This character string can be in upper, lower, or mixed case.

An SQL-92 Entry level–conformant driver will always return "table".

**"TimeDateAddIntervals"**

Returns an SQLUINTEGER bitmask enumerating the timestamp intervals supported by the driver and associated data source for the TIMESTAMPADD scalar function.

The following bitmasks are used to determine which intervals are supported:

SQL_FN_TSI_FRAC_SECOND<br>
SQL_FN_TSI_SECOND<br>
SQL_FN_TSI_MINUTE<br>
SQL_FN_TSI_HOUR<br>
SQL_FN_TSI_DAY<br>
SQL_FN_TSI_WEEK<br>
SQL_FN_TSI_MONTH<br>
SQL_FN_TSI_QUARTER<br>
SQL_FN_TSI_YEAR

An FIPS Transitional level-conformant driver will always return a bitmask in which all of these bits are set.

**"TimeDateDiffIntervals"**

Returns an SQLUINTEGER bitmask enumerating the timestamp intervals supported by the driver and associated data source for the TIMESTAMPDIFF scalar function.

The following bitmasks are used to determine which intervals are supported:

SQL_FN_TSI_FRAC_SECOND<br>
SQL_FN_TSI_SECOND<br>
SQL_FN_TSI_MINUTE<br>
SQL_FN_TSI_HOUR<br>
SQL_FN_TSI_DAY<br>
SQL_FN_TSI_WEEK<br>
SQL_FN_TSI_MONTH<br>
SQL_FN_TSI_QUARTER<br>
SQL_FN_TSI_YEAR

An FIPS Transitional level-conformant driver will always return a bitmask in which all of these bits are set.

**"TimeDateFunctions"**

Returns an SQLUINTEGER bitmask enumerating the scalar date and time functions supported by the driver and associated data source.

The following bitmasks are used to determine which date and time functions are supported:

SQL_FN_TD_CURRENT_DATE ODBC 3.0)<br>
SQL_FN_TD_CURRENT_TIME (ODBC 3.0)<br>
SQL_FN_TD_CURRENT_TIMESTAMP (ODBC 3.0)<br>
SQL_FN_TD_CURDATE (ODBC 1.0)<br>
SQL_FN_TD_CURTIME (ODBC 1.0)<br>
SQL_FN_TD_DAYNAME (ODBC 2.0)<br>
SQL_FN_TD_DAYOFMONTH (ODBC 1.0)<br>
SQL_FN_TD_DAYOFWEEK (ODBC 1.0)<br>
SQL_FN_TD_DAYOFYEAR (ODBC 1.0)<br>
SQL_FN_TD_EXTRACT (ODBC 3.0)<br>
SQL_FN_TD_HOUR (ODBC 1.0)<br>
SQL_FN_TD_MINUTE (ODBC 1.0)<br>
SQL_FN_TD_MONTH (ODBC 1.0)<br>
SQL_FN_TD_MONTHNAME (ODBC 2.0)<br>
SQL_FN_TD_NOW (ODBC 1.0)<br>
SQL_FN_TD_QUARTER (ODBC 1.0)<br>
SQL_FN_TD_SECOND (ODBC 1.0)<br>
SQL_FN_TD_TIMESTAMPADD (ODBC 2.0)<br>
SQL_FN_TD_TIMESTAMPDIFF (ODBC 2.0)<br>
SQL_FN_TD_WEEK (ODBC 1.0)<br>
SQL_FN_TD_YEAR (ODBC 1.0)

**"TxnCapable"**

Returns an SQLUSMALLINT value describing the transaction support in the driver or data source.

SQL_TC_NONE = Transactions not supported. (ODBC 1.0)

SQL_TC_DML = Transactions can contain only Data Manipulation Language (DML) statements (SELECT, INSERT, UPDATE, DELETE). Data Definition Language (DDL) statements encountered in a transaction cause an error. (ODBC 1.0)

SQL_TC_DDL_COMMIT = Transactions can contain only DML statements. DDL statements (CREATE TABLE, DROP INDEX, and so on) encountered in a transaction cause the transaction to be committed. (ODBC 2.0)

SQL_TC_DDL_IGNORE = Transactions can contain only DML statements. DDL statements encountered in a transaction are ignored. (ODBC 2.0)

SQL_TC_ALL = Transactions can contain DDL statements and DML statements in any order. (ODBC 1.0) (Because support of transactions is mandatory in SQL-92, an SQL-92 conformant driver [any level] will never return SQL_TC_NONE.)

**"TxnIsolationOption"**

Returns an SQLUINTEGER bitmask enumerating the transaction isolation levels available from the driver or data source.

The following bitmasks are used in conjunction with the flag to determine which options are supported:

SQL_TXN_READ_UNCOMMITTED<br>
SQL_TXN_READ_COMMITTED<br>
SQL_TXN_REPEATABLE_READ<br>
SQL_TXN_SERIALIZABLE

For descriptions of these isolation levels, see the description of SQL_DEFAULT_TXN_ISOLATION.

To set the transaction isolation level, an application calls SetConnectAttr to set the SQL_ATTR_TXN_ISOLATION attribute.

An SQL-92 Entry level-conformant driver will always return SQL_TXN_SERIALIZABLE as supported. A FIPS Transitional level-conformant driver will always return all of these options as supported.

**"Union"**

Returns an SQLUINTEGER bitmask enumerating the support for the UNION clause.

SQL_U_UNION = The data source supports the UNION clause.

SQL_U_UNION_ALL = The data source supports the ALL keyword in the UNION clause. (OdbcGetInfo returns both SQL_U_UNION and SQL_U_UNION_ALL in this case.)

An SQL-92 Entry level-conformant driver will always return both of these options as supported.

**"UserName"**

Returns a character string with the name used in a particular database, which can be different from the login name.

**"XOpenCliYear"**

Returns a character string that indicates the year of publication of the X/Open sepecification with which the version of the Driver Manager fully complies.

#### Result code

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="GetSqlState"></a>GetSqlState (CODBC)

Lists driver descriptions and driver attribute keywords. This function is implemented only by the Driver Manager.

```
FUNCTION GetSqlState () AS CWSTR
```

# <a name="Handle"></a>Handle (CODBC)

Returns the connection handle.

```
FUNCTION Handle () AS SQLHANDLE
```

# <a name="NativeSql"></a>NativeSql (CODBC)

Returns the SQL string as modified by the driver. NativeSql does not execute the SQL statement.

```
FUNCTION NativeSql (BYREF wszInText AS WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszInText* | SQL text string to be translated. |

#### Return value

The SQL string as modified by the driver.

**Result code** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="RollbackTran"></a>RollbackTran (CODBC)

Requests a rollback operation for all active operations on all statements associated with an environment. 

```
FUNCTION RollbackTran () AS SQLRETURN
```

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

#### Example

```
#include once "Afx/COdbc/COdbc.inc"
USING Afx

' // Create a connection object and connect with the database
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr
IF pDbc.Handle = NULL THEN PRINT "Unable to create the connection handle" : SLEEP : END

' // Allocate an statement object
DIM pStmt AS COdbcStmt = @pDbc
IF pStmt.Handle = NULL THEN PRINT "Unable to create the statement handle" : SLEEP : END

' // Cursor type
pStmt.SetMultiuserKeysetCursor
' // Bind the columns
DIM AS LONG lAuId, cbAuId
pStmt.BindCol(1, @lAuId, @cbAuId)
DIM wszAuthor AS WSTRING * 260, cbAuthor AS LONG
pStmt.BindCol(2, @wszAuthor, SIZEOF(wszAuthor), @cbAuthor)
DIM iYearBorn AS SHORT, cbYearBorn AS LONG
pStmt.BindCol(3, @iYearBorn, @cbYearBorn)

' // Generate a result set
pStmt.ExecDirect ("SELECT * FROM Authors ORDER BY Author")

' // Fill the values of the bounded application variables and its sizes
lAuId     = 998               : cbAuID     = SIZEOF(lAuId)
wszAuthor = "Edgar Allan Poe" : cbAuthor   = LEN(wszAuthor)
iYearBorn = 1809              : cbYearBorn = SIZEOF(iYearBorn)
' // Add the record
pStmt.AddRecord
IF pStmt.Error = FALSE THEN PRINT "Record added" ELSE PRINT pStmt.GetErrorInfo

' // Commit the transaction
' pDbc.CommitTran
' IF pDbc.Error = FALSE THEN PRINT "Commit succeeded"
' or Rollbacks it because this is a test
pDbc.RollbackTran 
IF pDbc.Error = FALSE THEN PRINT "Rollback succeeded" ELSE PRINT pDbc.GetErrorInfo
```

# <a name="SetConnectAttr"></a>SetConnectAttr (CODBC)

Sets attributes that govern aspects of connections.

```
FUNCTION SetConnectAttr (BYVAL Attribute AS SQLINTEGER, BYVAL ValuePtr AS SQLPOINTER, _
   BYVAL StringLength AS SQLINTEGER) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *Attribute* | Attribute to set, listed in "Comments." |
| *ValuePtr* | Pointer to the value to be associated with *Attribute*. Depending on the value of *Attribute*, *ValuePtr* will be a 32-bit unsigned integer value or will point to a null-terminated character string. Note that if the *Attribute* argument is a driver-specific value, the value in *ValuePtr* may be a signed integer. |
| *StringLength* | If *Attribute* is an ODBC-defined attribute and *ValuePtr* points to a character string or a binary buffer, this argument should be the length of *ValuePtr*. If Attribute is an ODBC-defined attribute and ValuePtr is an integer, *StringLength* is ignored.<br>If *Attribute* is a driver-defined attribute, the application indicates the nature of the attribute to the Driver Manager by setting the *StringLength* argument. *StringLength* can have the following values:<ul><li>If *ValuePtr* is a pointer to a character string, then *StringLength* is the length of the string or SQL_NTS.</li><li>If *ValuePtr* is a pointer to a binary buffer, then the application places the result of the SQL_LEN_BINARY_ATTR(length) macro in *StringLength*. This places a negative value in *StringLength*.</li><li>If *ValuePtr* is a pointer to a value other than a character string or a binary string, then *StringLength* should have the value SQL_IS_POINTER.</li><li>If *ValuePtr* contains a fixed-length value, then *StringLength* is either SQL_IS_INTEGER or SQL_IS_UINTEGER, as appropriate.</li></ul> |

#### Return value

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, SQL_INVALID_HANDLE, or SQL_STILL_EXECUTING.

### Overloaded Functions

Sets the current setting of the specified connection attribute.

```
FUNCTION SetConnectAttr (BYREF wszAttribute AS WSTRING, BYVAL dwAttribute AS SQLUINTEGER) AS SQLRETURN
FUNCTION SetConnectAttr (BYREF wszAttribute AS WSTRING, BYREF wszAttrValue AS WSTRING) AS SQLRETURN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszAttribute* | Attribute to set. |
| *dwAttribute* | The value to set. |
| *wszAttrValue* | The value to set. |

**"AccessMode"**

Sets an SQLUINTEGER value. SQL_MODE_READ_ONLY is used by the driver or data source as an indicator that the connection is not required to support SQL statements that cause updates to occur. This mode can be used to optimize locking strategies, transaction management, or other areas as appropriate to the driver or data source. The driver is not required to prevent such statements from being submitted to  the data source. The behavior of the driver and data source when asked to process SQL statements that are not read-only during a read-only connection is implementation-defined. SQL_MODE_READ_WRITE is the default.

**"AsyncEnable**

SQL_ASYNC_ENABLE_OFF = Off (the default)<br>
SQL_ASYNC_ENABLE_ON = On

Setting SQL_ASYNC_ENABLE_ON enables asynchronous execution for all future statement handles allocated on this connection. It is driver-defined whether this enables asynchronous execution for existing statement handles associated with this connection. An error is returned if asynchronous execution is enabled while there is an active statement on the connection.

In general, applications should execute functions asynchronously only on single-thread operating systems. On multithread operating systems, applications should execute functions on separate threads rather than executing them asynchronously on the same thread. Drivers that operate only on multithread operating systems do not need to support asynchronous execution.

**"AutoCommit"**

SQL_AUTOCOMMIT_OFF = The driver uses manual-commit mode, and theapplication must explicitly commit or roll back transactions with OdbcEndTran.

SQL_AUTOCOMMIT_ON = The driver uses autocommit mode. Each statement is committed immediately after it is executed. This is the default. Any open transactions on the connection are committed when SQL_ATTR_AUTOCOMMIT is set to SQL_AUTOCOMMIT_ON to change from manual-commit mode to autocommit mode.

Important: Some data sources delete the access plans and close the cursors for all statements on a connection each time a statement is committed; autocommit mode can cause this to happen after each nonquery statement is executed or when the cursor is closed for a query. When a batch is executed in autocommit mode, two things are possible. The entire batch can be treated as an autocommitable unit, or each statement in a batch is treated as an autocommitable unit. Certain data sources can support both these behaviors and may provide a way of choosing one or the other. It is driver-defined whether a batch is treated as an autocommitable unit or whether each individual statement within the batch is autocommitable.

**"ConnectionTimeout"**

Sets an SQLUINTEGER value corresponding to the number of seconds to wait for any request on the connection to complete before returning to the application. The driver should return SQLSTATE HYT00 (Timeout expired) anytime that it is possible to time out in a situation not associated with query execution or login. If the value is equal to 0 (the default), there is no timeout. Optional feature not implemented by the Microsoft Access Driver.

**"CurrentCatalog"**

Sets a character string containing the name of the catalog to be used by the data source. For example, in SQL Server, the catalog is a database, so the driver sends a USE database statement to the data source, where database is the database specified in wszAttrValue. For a single-tier driver, the catalog might be a directory, so the driver changes its current directory to the directory specified in wszAttrValue.

**"Cursors"**

An SQLULEN value specifying how the Driver Manager uses the ODBC cursor library:

SQL_CUR_USE_IF_NEEDED = The Driver Manager uses the ODBC cursor library only if it is needed. If the driver supports the SQL_FETCH_PRIOR option in SQLFetchScroll, the Driver Manager uses the scrolling capabilities of the driver. Otherwise, it uses the ODBC cursor library.

SQL_CUR_USE_ODBC = The Driver Manager uses the ODBC cursor library.

SQL_CUR_USE_DRIVER = The Driver Manager uses the scrolling capabilities of the driver. This is the default setting.

Warning: The cursor library will be removed in a future version of Windows. Avoid using this feature in new development work and plan to modify applications that currently use this feature. Microsoft recommends using the driver's cursor functionality.

**"LoginTimeout"**

Sets an SQLUINTEGER value corresponding to the number of seconds to wait for a login request to complete before returning to the application. The default is driver-dependent. If ValuePtr is 0, the timeout is disabled and a connection attempt will wait indefinitely. If the specified timeout exceeds the maximum login timeout in the data source, the driver substitutes that value and returns SQLSTATE 01S02 (Option value changed).

**"MetadataID"**

Returns an SQLUINTEGER value that determines how the string arguments of catalog functions are treated. Optional feature not implemented by the Microsoft Access Driver.

If SQL_TRUE, the string argument of catalog functions are treated as identifiers. The case is not significant. For nondelimited strings, the driver removes any trailing spaces and the string is folded to uppercase. For delimited strings, the driver removes any leading or trailing spaces and takes literally whatever is between the delimiters. If one of these arguments is set to a null pointer, the function returns SQL_ERROR and SQLSTATE HY009 (Invalid use of null pointer). If SQL_FALSE, the string arguments of catalog functions are not treated as identifiers. The case is significant. They can either contain a string search pattern or not, depending on the argument. The default value is SQL_FALSE. The TableType argument of SQLTables, which takes a list of values, is not affected by this attribute. SQL_ATTR_METADATA_ID can also be set on the statement level. (It is the only connection attribute that is also a statement attribute.)

**"PacketSize"**

Sets an SQLUINTEGER value specifying the network packet size in bytes. Optional feature not implemented by the Microsoft Access Driver.

Note Many data sources either do not support this option or only can return but not set the network packet size. If the specified size exceeds the maximum packet size or is smaller than the minimum packet size, the driver substitutes that value and returns SQLSTATE 01S02 (Option value changed). If the application sets packet size after a connection has already been made, the driver will return SQLSTATE HY011 (Attribute cannot be set now).

**"QuietMode"**

Sets a 32-bit window handle.

If the window handle is a null pointer, the driver does not display any dialog boxes. If the window handle is not a null pointer, it should be the parent window handle of the application. This is the default. The driver uses this handle to display dialog boxes.

Note The SQL_ATTR_QUIET_MODE connection attribute does not apply to dialog boxes displayed by DriverConnect.

**"Trace"**

Sets an SQLUINTEGER value telling the Driver Manager whether to perform tracing.

SQL_OPT_TRACE_OFF = Tracing off (the default)<br>
SQL_OPT_TRACE_ON = Tracing on

When tracing is on, the Driver Manager writes each ODBC function call to the trace file. Note When tracing is on, the Driver Manager can return SQLSTATE IM013 (Trace file error) from any function.

An application specifies a trace file with the TraceFile property. If the file already exists, the Driver Manager appends to the file. Otherwise, it creates the file. If tracing is on and no trace file has been specified, the Driver Manager writes to the file SQL.LOG in the root directory.

**"TxnIsolation"**

An SQLUINTEGER bitmask.

The following bitmasks are used in conjunction with the flag to determine which options are supported:

SQL_TXN_READ_UNCOMMITTED<br>
SQL_TXN_READ_COMMITTED<br>
SQL_TXN_REPEATABLE_READ<br>
SQL_TXN_SERIALIZABLE

For descriptions of these isolation levels, see the description of SQL_DEFAULT_TXN_ISOLATION.

To set the transaction isolation level, an application calls SetConnectAttr to set the SQL_ATTR_TXN_ISOLATION attribute.

An SQL-92 Entry level-conformant driver will always return SQL_TXN_SERIALIZABLE as supported. A FIPS Transitional level-conformant driver will always return all of these options as supported.

#### Return value

The current value of the attribute.

#### Result code

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_NO_DATA, SQL_ERROR, or SQL_INVALID_HANDLE.

# <a name="Supports"></a>Supports (CODBC)

Returns information about whether a driver supports a specific ODBC function. It is an alias for **Functions**.

```
FUNCTION Supports (BYVAL FunctionId AS SQLUSMALLINT) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *FunctionId* | A value that identifies the ODBC function of interest. |

#### Return value

TRUE if the function is supported, FALSE if it is not.

**Result value** (GetLastResult)

SQL_SUCCESS, SQL_SUCCESS_WITH_INFO, SQL_ERROR, or SQL_INVALID_HANDLE.

**Values to identify ODBC APIs**:

```
SQL_API_SQLALLOCCONNECT
SQL_API_SQLALLOCENV
SQL_API_SQLALLOCHANDLE
SQL_API_SQLALLOCSTMT
SQL_API_SQLBINDCOL
SQL_API_SQLBINDPARAM
SQL_API_SQLCANCEL
SQL_API_SQLCLOSECURSOR
SQL_API_SQLCOLATTRIBUTE
SQL_API_SQLCOLUMNS
SQL_API_SQLCONNECT
SQL_API_SQLCOPYDESC
SQL_API_SQLDATASOURCES
SQL_API_SQLDESCRIBECOL
SQL_API_SQLDISCONNECT
SQL_API_SQLENDTRAN
SQL_API_SQLERROR
SQL_API_SQLEXECDIRECT
SQL_API_SQLEXECUTE
SQL_API_SQLFETCH
SQL_API_SQLFETCHSCROLL
SQL_API_SQLFREECONNECT
SQL_API_SQLFREEENV
SQL_API_SQLFREEHANDLE
SQL_API_SQLFREESTMT
SQL_API_SQLGETCONNECTATTR
SQL_API_SQLGETCONNECTOPTION
SQL_API_SQLGETCURSORNAME
SQL_API_SQLGETDATA
SQL_API_SQLGETDESCFIELD
SQL_API_SQLGETDESCREC
SQL_API_SQLGETDIAGFIELD
SQL_API_SQLGETDIAGREC
SQL_API_SQLGETENVATTR
SQL_API_SQLGETFUNCTIONS
SQL_API_SQLGETINFO
SQL_API_SQLGETSTMTATTR
SQL_API_SQLGETSTMTOPTION
SQL_API_SQLGETTYPEINFO
SQL_API_SQLNUMRESULTCOLS
SQL_API_SQLPARAMDATA
SQL_API_SQLPREPARE
SQL_API_SQLPUTDATA
SQL_API_SQLROWCOUNT
SQL_API_SQLSETCONNECTATTR
SQL_API_SQLSETCONNECTOPTION
SQL_API_SQLSETCURSORNAME
SQL_API_SQLSETDESCFIELD
SQL_API_SQLSETDESCREC
SQL_API_SQLSETENVATTR
SQL_API_SQLSETPARAM
SQL_API_SQLSETSTMTATTR
SQL_API_SQLSETSTMTOPTION
SQL_API_SQLSPECIALCOLUMNS
SQL_API_SQLSTATISTICS
SQL_API_SQLTABLES
SQL_API_SQLTRANSACT
SQL_API_SQLALLOCHANDLESTD
SQL_API_SQLBULKOPERATIONS
SQL_API_SQLBINDPARAMETER
SQL_API_SQLBROWSECONNECT
SQL_API_SQLCOLATTRIBUTES
SQL_API_SQLCOLUMNPRIVILEGES
SQL_API_SQLDESCRIBEPARAM
SQL_API_SQLDRIVERCONNECT
SQL_API_SQLDRIVERS
SQL_API_SQLEXTENDEDFETCH
SQL_API_SQLFOREIGNKEYS
SQL_API_SQLMORERESULTS
SQL_API_SQLNATIVESQL
SQL_API_SQLNUMPARAMS
SQL_API_SQLPARAMOPTIONS
SQL_API_SQLPRIMARYKEYS
SQL_API_SQLPROCEDURECOLUMNS
SQL_API_SQLPROCEDURES
SQL_API_SQLSETPOS
SQL_API_SQLSETSCROLLOPTIONS
SQL_API_SQLTABLEPRIVILEGES
```
