# CADOConnection Class

Represents an open connection to a data source.

**Include file**: CADOConnection.inc (include CADODB.inc)

A **Connection** object represents a unique session with a data source. In the case of a client/server database system, it may be equivalent to an actual network connection to the server. Depending on the functionality supported by the provider, some collections, methods, or properties of a **Connection** object may not be available.

With the collections, methods, and properties of a **Connection** object, you can do the following:

* Configure the connection before opening it with the **ConnectionString**, **ConnectionTimeout**, and **Mode** properties. **ConnectionString** is the default property of the **Connection** object.
* Set the **CursorLocation** property to client to invoke the Microsoft Cursor Service for OLE DB, which supports batch updates.
* Set the default database for the connection with the **DefaultDatabase** property.
* Set the level of isolation for the transactions opened on the connection with the **IsolationLevel** property.
* Specify an OLE DB provider with the **Provider** property.
* Establish, and later break, the physical connection to the data source with the **Open** and **Close** methods.
* Execute a command on the connection with the **Execute** method and configure the execution with the **CommandTimeout** property.
* Manage transactions on the open connection, including nested transactions if the provider supports them, with the **BeginTrans**, **CommitTrans**, and **RollbackTrans** methods and the **Attributes** property.
* Examine errors returned from the data source with the **Errors** collection.
* Read the version from the ADO implementation used with the **Version** property.
* Obtain schema information about your database with the **OpenSchema** method.

You can create **Connection** objects independently of any other previously defined object.

To execute a query without using a **Command** object, pass a query string to the **Execute** method of a **Connection** object. However, a **Command** object is required when you want to persist the command text and re-execute it, or use query parameters.

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [Attributes](#Attributes) | Indicates one or more characteristics of an object. |
| [BeginTrans](#BeginTrans) | Begins a new transaction. |
| [Cancel](#Cancel) | Cancels execution of a pending, asynchronous method call. |
| [Close](#Close) | Closes a **Connection** object and any dependent objects. |
| [CommandTimeout](#CommandTimeout) | Sets or returns a Long value that indicates, in seconds, how long to wait for a command to execute. Default is 30. |
| [CommitTrans](#CommitTrans) | Saves any changes and ends the current transaction. It may also start a new transaction. |
| [ConnectionString](#ConnectionString) | Indicates the information used to establish a connection to a data source. |
| [ConnectionTimeout](#ConnectionTimeout) | Indicates how long to wait while establishing a connection before terminating the attempt and generating an error. |
| [CursorLocation](#CursorLocation) | Indicates the location of the cursor service. |
| [DefaultDatabase](#DefaultDatabase) | Indicates the default database for a **Connection** object. |
| [Execute](#Execute) | Executes the specified query, SQL statement, stored procedure, or provider-specific text. |
| [GetErrorInfo](#GetErrorInfo) | Returns information about ADO errors. |
| [IsolationLevel](#IsolationLevel) | Indicates the level of isolation for a **Connection** object. |
| [Mode](#Mode) | Indicates the available permissions for modifying data in a **Connection** object. |
| [Open](#Open) | Opens a connection to a data source. |
| [OpenSchema](#OpenSchema) | Obtains database schema information from the provider. |
| [Properties](#Properties) | Returns a reference to the **Properties** collection. |
| [Provider](#Provider) | Indicates the name of the provider for a **Connection** object. |
| [RollbackTrans](#RollbackTrans) | Saves any changes and ends the current transaction. It may also start a new transaction. |
| [State](#State) | Indicates if a **Connection** is open or closed. |
| [Version](#Version) | Indicates the ADO version number. |

### Connection Dynamic Properties

The table below is a cross-index of the ADO and OLE DB names for each standard OLE DB provider dynamic property. Your providers may add more properties than listed here. For the specific information about provider-specific dynamic properties, see your provider documentation.

| ADO Property Name | OLE DB Property Name |
| ----------------- | ------------------- |
| Active Sessions | DBPROP_ACTIVESESSIONS |
| Autocommit Isolation Levels | DBPROP_SESS_AUTOCOMMITISOLEVELS |
| Catalog Location | DBPROP_CATALOGLOCATION |
| Catalog Term | DBPROP_CATALOGTERM |
| Column Definition | DBPROP_COLUMNDEFINITION |
| Connect Timeout | DBPROP_INIT_TIMEOUT |
| Current Catalog | DBPROP_CURRENTCATALOG |
| Data Source | DBPROP_INIT_DATASOURCE |
| Data Source Name | DBPROP_DATASOURCENAME |
| Data Source Object Threading Model | DBPROP_DSOTHREADMODEL |
| DBMS Name | DBPROP_DBMSNAME |
| DBMS Version | DBPROP_DBMSVER |
| Extended Properties | DBPROP_INIT_PROVIDERSTRING |
| GROUP BY Support | DBPROP_GROUPBY |
| Heterogeneous Table Support | DBPROP_HETEROGENEOUSTABLES |
| Identifier Case Sensitivity | DBPROP_IDENTIFIERCASE |
| Initial Catalog | DBPROP_INIT_CATALOG |
| Isolation Levels | DBPROP_SUPPORTEDTXNISOLEVELS |
| Isolation Retention | DBPROP_SUPPORTEDTXNISORETAIN |
| Locale Identifier | DBPROP_INIT_LCID |
| Location | DBPROP_INIT_LOCATION |
| Maximum Index Size | DBPROP_MAXINDEXSIZE |
| Maximum Row Size | DBPROP_MAXROWSIZE |
| Maximum Row Size Includes BLOB | BPROP_MAXROWSIZEINCLUDESBLOB |
| Maximum Tables in SELECT | DBPROP_MAXTABLESINSELECT |
| Mode | DBPROP_INIT_MODE |
| Multiple Parameter Sets | DBPROP_MULTIPLEPARAMSETS |
| Multiple Results | DBPROP_MULTIPLERESULTS |
| Multiple Storage Objects | DBPROP_MULTIPLESTORAGEOBJECTS |
| Multi-Table Update | DBPROP_MULTITABLEUPDATE |
| NULL Collation Order | DBPROP_NULLCOLLATION |
| NULL Concatenation Behavior | DBPROP_CONCATNULLBEHAVIOR |
| OLE DB Services | DBPROP_INIT_OLEDBSERVICES |
| OLE DB Version | DBPROP_PROVIDEROLEDBVER |
| OLE Object Support | DBPROP_OLEOBJECTS |
| Open Rowset Support | DBPROP_OPENROWSETSUPPORT |
| ORDER BY Columns in Select List | DBPROP_ORDERBYCOLUMNSINSELECT |
| Output Parameter Availability | DBPROP_OUTPUTPARAMETERAVAILABILITY |
| Pass By Ref Accessors | DBPROP_BYREFACCESSORS |
| Password | DBPROP_AUTH_PASSWORD |
| Persist Security Info | DBPROP_AUTH_PERSIST_SENSITIVE_AUTHINFO |
| Persistent ID Type | DBPROP_PERSISTENTIDTYPE |
| Prepare Abort Behavior | DBPROP_PREPAREABORTBEHAVIOR |
| Prepare Commit Behavior | DBPROP_PREPARECOMMITBEHAVIOR |
| Procedure Term | DBPROP_PROCEDURETERM |
| Prompt | DBPROP_INIT_PROMPT |
| Provider Friendly Name | DBPROP_PROVIDERFRIENDLYNAME |
| Provider Name | DBPROP_PROVIDERFILENAME |
| Provider Version | DBPROP_PROVIDERVER |
| Read-Only Data Source | DBPROP_DATASOURCEREADONLY |
| Rowset Conversions on Command | DBPROP_ROWSETCONVERSIONSONCOMMAND |
| Schema Term | DBPROP_SCHEMATERM |
| Schema Usage | DBPROP_SCHEMAUSAGE |
| SQL Support | DBPROP_SQLSUPPORT |
| Structured Storage | DBPROP_STRUCTUREDSTORAGE |
| Subquery Support | DBPROP_SUBQUERIES |
| Table Term | DBPROP_TABLETERM |
| Transaction DDL | DBPROP_SUPPORTEDTXNDDL |
| User ID | DBPROP_AUTH_USERID |
| User Name | DBPROP_USERNAME |
| Window Handle | DBPROP_INIT_HWND |

# <a name="Attributes"></a>Attributes

Indicates one or more characteristics of an object.

```
PROPERTY Attributes () AS LONG
PROPERTY Attributes (BYVAL lAttr AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *lAttr* | The value can be the sum of one or more **XactAttributeEnum** values. The default is zero (0). |

#### Return value

#### XactAttributeEnum

Specifies the transaction attributes of a Connection object.

| Constant   | Description |
| ---------- | ----------- |
| **adXactAbortRetaining** | Performs retaining aborts—that is, calling **RollbackTrans** automatically starts a new transaction. Not all providers support this. |
| **adXactCommitRetaining** | Performs retaining commits—that is, calling **CommitTrans** automatically starts a new transaction. Not all providers support this. |

#### Remarks

Use the **Attributes** property to set or return characteristics of **Connection** objects.

When you set multiple attributes, you can sum the appropriate constants. If you set the property value to a sum including incompatible constants, an error occurs.

#### Remote Data Service Usage

This property is not available on a client-side **Connection** object.


# <a name="BeginTrans"></a>BeginTrans

Begins a new transaction.

```
FUNCTION BeginTrans () AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *lAttr* | The value can be the sum of one or more **XactAttributeEnum** values. The default is zero (0). |

#### Return value

The nesting level of the transaction.

#### Remarks

Use **BeginTrans**, **CommitTrans** and **RollbackTrans** with a **Connection** object when you want to save or cancel a series of changes made to the source data as a single unit. For example, to transfer money between accounts, you subtract an amount from one and add the same amount to the other. If either update fails, the accounts no longer balance. Making these changes within an open transaction ensures that either all or none of the changes go through.

Not all providers support transactions. Verify that the provider-defined property "Transaction DDL" appears in the **Connectio** object's **Properties** collection, indicating that the provider supports transactions. If the provider does not support transactions, calling one of these methods will return an error.

After you call the **BeginTrans** method, the provider will no longer instantaneously commit changes you make until you call **CommitTrans** or **RollbackTrans** to end the transaction.

For providers that support nested transactions, calling the **BeginTrans** method within an open transaction starts a new, nested transaction. The return value indicates the level of nesting: a return value of "1" indicates you have opened a top-level transaction (that is, the transaction is not nested within another transaction), "2" indicates that you have opened a second-level transaction (a transaction nested within a top-level transaction), and so forth. Calling **CommitTrans** or **RollbackTrans** affects only the most recently opened transaction; you must close or roll back the current transaction before you can resolve any higher-level transactions.

Calling the **CommitTrans** method saves changes made within an open transaction on the connection and ends the transaction. Calling the **RollbackTrans** method reverses any changes made within an open transaction and ends the transaction. Calling either method when there is no open transaction generates an error.

Depending on the **Connection** object's **Attributes** property, calling either the **CommitTrans** or **RollbackTrans** methods may automatically start a new transaction. If the **Attributes** property is set to **adXactCommitRetaining**, the provider automatically starts a new transaction after a **CommitTrans** call. If the **Attributes** property is set to **adXactAbortRetaining**, the provider automatically starts a new transaction after a **RollbackTrans** call.

#### Remote Data Service

The **BeginTrans**, **CommitTrans**, and **RollbackTrans** methods are not available on a client-side **Connection** object.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Open the recordset
DIM pRecordset AS CAdoRecordset
DIM cvSource AS CVAR = "SELECT * FROM Authors"
pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

' // Begin a transaction
pConnection.BeginTrans

' // Parse the recordset
DO
   ' // While not at the end of the recordset...
   IF pRecordset.EOF THEN EXIT DO
   ' // Get the content of the "Author" column
   DIM cvRes AS CVAR = pRecordset.Collect("Year Born")
   IF cvRes.ValInt = 1947 THEN pRecordset.Collect("Year Born") = 1900
   ' // Fetch the next row
   IF pRecordset.MoveNext <> S_OK THEN EXIT DO
LOOP
' // Commit the transaction
'pConnection.CommitTrans
' // Rollback the transaction because this is a demo
pConnection.RollbackTrans
```

# <a name="Cancel"></a>Cancel

Cancels execution of a pending, asynchronous method call.

```
FUNCTION Cancel () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use the **Cancel** method to terminate execution of an asynchronous method call (that is, a method invoked with the **adAsyncConnect**, **adAsyncExecute**, or **adAsyncFetch** option).

For a **Connection** object, the last asynchronous call to the **Execute** or **Open** methods is terminated.

# <a name="Close"></a>Close

Closes a **Connection** object and any dependent objects.

```
FUNCTION Close () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use the **Close** method to close a **Connection** to free any associated system resources. Closing an object does not remove it from memory; you can change its property settings and open it again later. To completely eliminate an object from memory, release the connection calling the **Release** method of the interface.

#### Usage example

```
IF pConnection.State = adStateOpen THEN pConnection.Close
```

# <a name="CommandTimeout"></a>CommandTimeout

Sets or returns a Long value that indicates, in seconds, how long to wait for a command to execute. Default is 30.

```
PROPERTY CommandTimeout () AS LONG
PROPERTY CommandTimeout (BYVAL lTimeout AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *lTimeout* | Time to wait for a command to execute. |

#### Return value

The timeout value.

#### Remarks

Use the **CommandTimeout** property on a **Connection** object to allow the cancellation of an **Execute** method call, due to delays from network traffic or heavy server use. If the interval set in the **CommandTimeout** property elapses before the command completes execution, an error occurs and ADO cancels the command. If you set the property to zero, ADO will wait indefinitely until the execution is complete. Make sure the provider and data source to which you are writing code support the **CommandTimeout** functionality.

The **CommandTimeout** setting on a **Connection** object has no effect on the **CommandTimeout** setting on a **Command** object on the same **Connection**; that is, the **Command** object's **CommandTimeout** property does not inherit the value of the **Connection** object's **CommandTimeout** value.

On a **Connection** object, the **CommandTimeout** property remains read/write after the **Connection** is opened.

#### Examples

```
' // Sets the timeout to 35 seconds
pConnection.CommandTimeout = 35
```
```
' // Gets the timeout value
DIM lTimeout AS LONG = pConnection.CommandTimeout
```

# <a name="CommitTrans"></a>CommitTrans

Saves any changes and ends the current transaction. It may also start a new transaction.

```
FUNCTION CommitTrans () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use **BeginTrans**, **CommitTrans** and **RollbackTrans** with a **Connection** object when you want to save or cancel a series of changes made to the source data as a single unit. For example, to transfer money between accounts, you subtract an amount from one and add the same amount to the other. If either update fails, the accounts no longer balance. Making these changes within an open transaction ensures that either all or none of the changes go through.

Not all providers support transactions. Verify that the provider-defined property "Transaction DDL" appears in the **Connection** object's **Properties** collection, indicating that the provider supports transactions. If the provider does not support transactions, calling one of these methods will return an error.

After you call the **BeginTrans method**, the provider will no longer instantaneously commit changes you make until you call **CommitTrans** or **RollbackTrans** to end the transaction.

For providers that support nested transactions, calling the **BeginTrans** method within an open transaction starts a new, nested transaction. The return value indicates the level of nesting: a return value of "1" indicates you have opened a top-level transaction (that is, the transaction is not nested within another transaction), "2" indicates that you have opened a second-level transaction (a transaction nested within a top-level transaction), and so forth. Calling **CommitTrans** or **RollbackTrans** affects only the most recently opened transaction; you must close or roll back the current transaction before you can resolve any higher-level transactions.

Calling the **CommitTrans** method saves changes made within an open transaction on the connection and ends the transaction. Calling the **RollbackTrans** method reverses any changes made within an open transaction and ends the transaction. Calling either method when there is no open transaction generates an error.

Depending on the **Connection** object's **Attributes** property, calling either the **CommitTrans** or **RollbackTrans** methods may automatically start a new transaction. If the **Attributes** property is set to **adXactCommitRetaining**, the provider automatically starts a new transaction after a **CommitTrans** call. If the **Attributes** property is set to **adXactAbortRetaining**, the provider automatically starts a new transaction after a **RollbackTrans** call.

#### Remote Data Service

The **BeginTrans**, **CommitTrans**, and **RollbackTrans** methods are not available on a client-side **Connection** object.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Open the recordset
DIM pRecordset AS CAdoRecordset
DIM cvSource AS CVAR = "SELECT * FROM Authors"
pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

' // Begin a transaction
pConnection.BeginTrans

' // Parse the recordset
DO
   ' // While not at the end of the recordset...
   IF pRecordset.EOF THEN EXIT DO
   ' // Get the content of the "Author" column
   DIM cvRes AS CVAR = pRecordset.Collect("Year Born")
   IF cvRes.ValInr = 1947 THEN pRecordset.Collect("Year Born") = 1900
   ' // Fetch the next row
   IF pRecordset.MoveNext <> S_OK THEN EXIT DO
LOOP
' // Commit the transaction
'pConnection.CommitTrans
' // Rollback the transaction because this is a demo
pConnection.RollbackTrans
```

# <a name="ConnectionString"></a>ConnectionString

Indicates the information used to establish a connection to a data source.

```
PROPERTY ConnectionString () AS CBSTR
PROPERTY ConnectionString (BYREF cbsConConStr AS CBSTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsConConStr* | The connection string. |

#### Return value

The connection string.

#### Remarks

Use the **ConnectionString** property to specify a data source by passing a detailed connection string containing a series of argument = value statements separated by semicolons.

ADO supports five arguments for the **ConnectionString** property; any other arguments pass directly to the provider without any processing by ADO. The arguments ADO supports are as follows.

| Argument   | Description |
| ---------- | ----------- |
| **Provider=** | Specifies the name of a provider to use for the connection. |
| **File Name=** | Specifies the name of a provider-specific file (for example, a persisted data source object) containing preset connection information. |
| **Remote Provider=** | Specifies the name of a provider to use when opening a client-side connection. (Remote Data Service only.) |
| **Remote Server=** | Specifies the path name of the server to use when opening a client-side connection. (Remote Data Service only.) |
| **URL=** | Specifies the connection string as an absolute URL identifying a resource, such as a file or directory. |

After you set the **ConnectionString** property and open the **Connection** object, the provider may alter the contents of the property, for example, by mapping the ADO-defined argument names to their provider equivalents.

The **ConnectionString** property automatically inherits the value used for the **ConnectionString** argument of the Open method, so you can override the current **ConnectionString** property during the **Open** method call.

Because the **File Name** argument causes ADO to load the associated provider, you cannot pass both the **Provider** and **File Name** arguments.

The **ConnectionString** property is read/write when the connection is closed and read-only when it is open.

Duplicates of an argument in the **ConnectionString** property are ignored. The last instance of any argument is used.

#### Remote Data Service Usage

When used on a client-side **Connection** object, the **ConnectionString** property can include only the **Remote Provider** and **Remote Server** parameters.

#### Usage examples

```
pConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"
pConnection.Open
```
```
DIM cbsConStr AS CBSTR = pConnection.ConnectionString
```

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"
pConnection.Open

' // Open the recordset
DIM pRecordset AS CAdoRecordset
DIM cbsSource AS CBSTR = "SELECT * FROM Authors"
pRecordset.Open(cbsSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

' // Parse the recordset
DO
   ' // While not at the end of the recordset...
   IF pRecordset.EOF THEN EXIT DO
   ' // Get the content of the "Author" column
   DIM cvRes AS CVAR = pRecordset.Collect("Author")
   PRINT cvRes
   ' // Fetch the next row
   IF pRecordset.MoveNext <> S_OK THEN EXIT DO
LOOP
```

# <a name="ConnectionTimeout"></a>ConnectionTimeout

Indicates how long to wait while establishing a connection before terminating the attempt and generating an error.

```
PROPERTY ConnectionTimeout () AS LONG
PROPERTY ConnectionTimeout (BYVAL lTimeout AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *lTimeout* | Time to wait for a command to execute. |

#### Return value

The timeout value.

#### Remarks

Use the **ConnectionTimeout** property on a **Connection** object if delays from network traffic or heavy server use make it necessary to abandon a connection attempt. If the time from the **ConnectionTimeout** property setting elapses prior to the opening of the connection, an error occurs and ADO cancels the attempt. If you set the property to zero, ADO will wait indefinitely until the connection is opened. Make sure the provider to which you are writing code supports the **ConnectionTimeout** functionality.

The **ConnectionTimeout** property is read/write when the connection is closed and read-only when it is open.

On a **Connection** object, the **ConnectionTimeout** property remains read/write after the **Connection** is opened.

# <a name="CursorLocation"></a>CursorLocation

Indicates the location of the cursor service.

```
PROPERTY CursorLocation () AS CursorLocationEnum
PROPERTY CursorLocation (BYVAL lCursorLoc AS CursorLocationEnum)
```

| Parameter  | Description |
| ---------- | ----------- |
| *lCursorLoc* | One of the **CursorLocationEnum** values. |

#### Return value

One of the **CursorLocationEnum** values.

#### CursorLocationEnum

Specifies the location of the cursor service.

| Constant   | Description |
| ---------- | ----------- |
| **adUseClient** | Uses client-side cursors supplied by a local cursor library. Local cursor services often will allow many features that driver-supplied cursors may not, so using this setting may provide an advantage with respect to features that will be enabled. For backward compatibility, the synonym **adUseClientBatch** is also supported. |
| **adUseNone** | Does not use cursor services. (This constant is obsolete and appears solely for the sake of backward compatibility.) |
| **adUseServer** | Default. Uses data-provider or driver-supplied cursors. These cursors are sometimes very flexible and allow for additional sensitivity to changes others make to the data source. However, some features of the Microsoft Cursor Service for OLE DB (such as disassociated Recordset objects) cannot be simulated with server-side cursors and these features will be unavailable with this setting. |

#### Remarks

This property allows you to choose between various cursor libraries accessible to the provider. Usually, you can choose between using a client-side cursor library or one that is located on the server.

This property setting affects connections established only after the property has been set. Changing the **CursorLocation** property has no effect on existing connections.

Cursors returned by the **Execute** method inherit this setting. **Recordset** objects will automatically inherit this setting from their associated connections.

This property is read/write on a **Connection** or a closed **Recordset**, and read-only on an open **Recordset**.

#### Remote Data Service Usage

When used on a client-side **Recordset** or **Connection** object, the **CursorLocation** property can only be set to **adUseClient**.

#### Examples

```
' // Sets the cursor location
pConnection.CursorLocation = adUseClient
```
```
' // Gets the timeout value
DIM lCursorLoc AS LONG = pConnection.CursorLocation
```

# <a name="DefaultDatabase"></a>DefaultDatabase

Indicates the default database for a **Connection** object.

```
PROPERTY DefaultDatabase () AS CBSTR
PROPERTY DefaultDatabase (BYREF cbsDatabase AS CBSTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsDatabase* | Name of the database available from the provider. |

#### Return value

The name of the database.

#### Remarks

Use the **DefaultDatabase** property to set or return the name of the default database on a specific **Connection** object.

If there is a default database, SQL strings may use an unqualified syntax to access objects in that database. To access objects in a database other than the one specified in the **DefaultDatabase** property, you must qualify object names with the desired database name. Upon connection, the provider will write default database information to the **DefaultDatabase** property. Some providers allow only one database per connection, in which case you cannot change the **DefaultDatabase** property.

Some data sources and providers may not support this feature, and may return an error or an empty string.

#### Remote Data Service Usage

This property is not available on a client-side **Connection** object.

# <a name="Execute"></a>Execute

Executes the specified query, SQL statement, stored procedure, or provider-specific text.

```
FUNCTION Execute (BYREF cbsCommandText AS CBSTR, BYVAL RecordsAffected AS LONG PTR = NULL, _
   BYVAL Options AS LONG = -1) AS Afx_ADORecordset PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsCommandText* | A string value that contains the SQL statement, stored procedure, a URL, or provider-specific text to execute. Optionally, table names can be used but only if the provider is SQL aware. For example if a table name of "Customers" is used, ADO will automatically prepend the standard SQL Select syntax to form and pass "SELECT * FROM Customers" as a T-SQL statement to the provider. |
| *RecordsAffected* | Optional. Pointer to a LONG to which the provider returns the number of records that the operation affected. |
| *Options* | Optional. A Long value that indicates how the provider should evaluate the **CommandText** argument. Can be a bitmask of one or more **CommandTypeENum** or **ExecuteOptionEnum** values. |

Use the **ExecuteOptionEnum** value **adExecuteNoRecords** to improve performance by minimizing internal processing.

Do not use **adExecuteStream** with the **Execute** method of a **Connection** object.

Do not use the **CommandTypeEnum** values of **adCmdFile** or **adCmdTableDirect** with **Execute**. These values can only be used as options with the **Open** and **Requery** methods of a **Recordset**.

#### CommandTypeEnum

Specifies how a command argument should be interpreted.

| Constant   | Value       |
| ---------- | ----------- |
| **adCmdUnspecified** | Does not specify the command type argument. |
| **adCmdText** | Evaluates **CommandText** as a textual definition of a command or stored procedure call. |
| **adCmdTable** | Evaluates **CommandText** as a table name whose columns are all returned by an internally generated SQL query. |
| **adCmdStoredProc** | Evaluates **CommandText** as a stored procedure name. |
| **adCmdUnknown** | Default. Indicates that the type of command in the **CommandText** property is not known. |
| **adCmdFile** | Evaluates **CommandText** as the file name of a persistently stored **Recordset**. Used with **Recordset** **Open** or **Requery** only. |
| **adCmdTableDirect** | Evaluates **CommandText** as a table name whose columns are all returned. Used with **Recordset** **Open** or **Requery** only. To use the **Seek** method, the **Recordset** must be opened with **adCmdTableDirect**. This value cannot be combined with the **ExecuteOptionEnum** value **adAsyncExecute**. |

#### ExecuteOptionEnum

Specifies how a provider should execute a command.

| Constant   | Value       |
| ---------- | ----------- |
| **adAsyncExecute** | Indicates that the command should execute asynchronously. This value cannot be combined with the **CommandTypeEnum** value **adCmdTableDirect**. |
| **adAsyncFetch** | Indicates that the remaining rows after the initial quantity specified in the CacheSize property should be retrieved asynchronously. |
| **adAsyncFetchNonBlocking** | Indicates that the main thread never blocks while retrieving. If the requested row has not been retrieved, the current row automatically moves to the end of the file. If you open a **Recordset** from a **Stream** containing a persistently stored **Recordset**, **adAsyncFetchNonBlocking** will not have an effect; the operation will be synchronous and blocking. **adAsynchFetchNonBlocking** has no effect when the **adCmdTableDirect** option is used to open the **Recordset**. |
| **adExecuteNoRecords** | Indicates that the command text is a command or stored procedure that does not return rows (for example, a command that only inserts data). If any rows are retrieved, they are discarded and not returned. **adExecuteNoRecords** can only be passed as an optional parameter to the **Command** or **Connection** **Execute** method. |
| **adExecuteStream** | Indicates that the results of a command execution should be returned as a stream. |
| **adExecuteStream** | Indicates that the results of a command execution should be returned as a stream. **adExecuteStream** can only be passed as an optional parameter to the **Command** **Execute** method. |
| **adExecuteRecord** | Indicates that the **CommandText** is a command or stored procedure that returns a single row which should be returned as a **Record** object. |
| **adOptionUnspecified** | Indicates that the command is unspecified. |

#### Return value

An **Afx_ADORecordset** object reference on success, or NULL on failure.

#### Remarks

Using the **Execute** method on a **Connection** object executes whatever query you pass to the method in the **CommandText** argument on the specified connection. If the **CommandText** argument specifies a row-returning query, any results that the execution generates are stored in a new **Recordset** object. If the command is not intended to return results (for example, an SQL UPDATE query) the provider returns Nothing as long as the option **adExecuteNoRecords** is specified; otherwise **Execute** returns a closed **Recordset**.

The returned **Recordset** object is always a read-only, forward-only cursor. If you need a **Recordset** object with more functionality, first create a **Recordset** object with the desired property settings, then use the **Recordset** object's **Open** method to execute the query and return the desired cursor type.

The contents of the **CommandText** argument are specific to the provider and can be standard SQL syntax or any special command format that the provider supports.

An **ExecuteComplete** event will be issued when this operation concludes.

**Note**: URLs using the http scheme will automatically invoke the Microsoft OLE DB Provider for Internet Publishing.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Create the recordset by executing a query
DIM pRecordset AS CAdoRecordset
pRecordset.Attach(pConnection.Execute("SELECT TOP 20 * FROM Authors ORDER BY Author"))

' // Parse the recordset
DO
   ' // While not at the end of the recordset...
   IF pRecordset.EOF THEN EXIT DO
   ' // Get the content of the "Author" column
   DIM cvRes AS CVAR = pRecordset.Collect("Author")
   PRINT cvRes
   ' // Fetch the next row
   IF pRecordset.MoveNext <> S_OK THEN EXIT DO
LOOP
```

# <a name="GetErrorInfo"></a>GetErrorInfo

Returns information about ADO errors.

```
FUNCTION GetErrorInfo (BYVAL nError AS HRESULT = 0) AS CBSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nError* | Optional. The error code. |

#### Return value

A description of the error(s).

# <a name="IsolationLevel"></a>IsolationLevel

Indicates the level of isolation for a **Connection** object.

```
PROPERTY IsolationLevel () AS IsolationLevelEnum
PROPERTY IsolationLevel (BYVAL Level AS IsolationLevelEnum)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nError* | Optional. The error code. |

#### Return value

An **IsolationLevelEnum** value.

#### IsolationLevelEnum

Specifies the level of transaction isolation for a Connection object.

| Constant   | Description |
| ---------- | ----------- |
| **adXactUnspecified** | Indicates that the provider is using a different isolation level than specified, but that the level cannot be determined. |
| **adXactChaos** | Indicates that pending changes from more highly isolated transactions cannot be overwritten. |
| **adXactBrowse** | Indicates that from one transaction you can view uncommitted changes in other transactions. |
| **adXactReadUncommitted** | Same as **adXactBrowse**. |
| **adXactCursorStability** | Indicates that from one transaction you can view changes in other transactions only after they have been committed. |
| **adXactReadCommitted** | Same as **adXactCursorStability**. |
| **adXactRepeatableRead** | Indicates that from one transaction you cannot see changes made in other transactions, but that requerying can retrieve new Recordset objects. |
| **adXactIsolated** | Indicates that transactions are conducted in isolation of other transactions. |
| **adXactSerializable** | Same as **adXactIsolated**. |

#### Remarks

Use the **IsolationLevel** property to set the isolation level of a **Connection** object. The setting does not take effect until the next time you call the **BeginTrans** method. If the level of isolation you request is unavailable, the provider may return the next greater level of isolation without updating the **IsolationLevel** property.

#### Remote Data Service Usage

When used on a client-side **Connection** object, the **IsolationLevel** property can be set only to **adXactUnspecified**.

Because users are working with disconnected **Recordset** objects on a client-side cache, there may be multiuser issues. For instance, when two different users try to update the same record, Remote Data Service simply allows the user who updates the record first to "win." The second user's update request will fail with an error.

#### Examples

```
' // Sets the isolation level
pConnection.IsolationLevel = adXactUnspecified
```
```
' // Gets the isolation level
DIM level AS LONG = pConnection.IsolationLevel
```

# <a name="Mode"></a>Mode

Indicates the available permissions for modifying data in a **Connection** object.

```
PROPERTY Mode () AS ConnectModeEnum
PROPERTY Mode (BYVAL lMode AS ConnectModeEnum)
```

| Parameter  | Description |
| ---------- | ----------- |
| *lMode* | ConnectModeEnum value. The default value for a **Connection** is **adModeUnknown**. |

#### Return value

A **ConnectModeEnum** value.

#### ConnectModeEnum

Specifies the available permissions for modifying data in a **Connection**, opening a **Record**, or specifying values for the Mode property of the **Record** and **Stream** objects.

| Constant   | Description |
| ---------- | ----------- |
| **adModeRead** | Indicates read-only permissions. |
| **adModeReadWrite** | Indicates read/write permissions. |
| **adModeRecursive** | Used in conjunction with the other *ShareDeny* values (**adModeShareDenyNone**, **adModeShareDenyWrite**, or **adModeShareDenyRead**) to propagate sharing restrictions to all sub-records of the current **Record**. It has no affect if the **Record** does not have any children. A run-time error is generated if it is used with **adModeShareDenyNone** only. However, it can be used with **adModeShareDenyNone** when combined with other values. For example, you can use "**adModeRead OR adModeShareDenyNone OR adModeRecursive**". |
| **adModeShareDenyNone** | Allows others to open a connection with any permissions. Neither read nor write access can be denied to others. |
| **adModeShareDenyRead** | Prevents others from opening a connection with read permissions. |
| **adModeShareDenyWrite** | Prevents others from opening a connection with write permissions. |
| **adModeShareExclusive** | Prevents others from opening a connection. |
| **adModeUnknown** | Default. Indicates that the permissions have not yet been set or cannot be determined. |
| **adModeWrite** | Indicates write-only permissions. |

# <a name="Open"></a>Open

Opens a connection to a data source.

```
FUNCTION Open (BYREF cbsConStr AS CBSTR = "", BYREF cbsUserID AS CBSTR = "", _
   BYREF cbsPassword AS CBSTR = "", BYVAL Options AS LONG = adOptionUnspecified) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsConStr* | Optional. An string value that contains connection information. |
| *cbsUserID* | Optional. An string value that contains a user name to use when establishing the connection. |
| *cbsPassword* | Optional. An string value that contains a password to use when establishing the connection. |
| *Options* | Optional. A **ConnectOptionEnum** value that determines whether this method should return after (synchronously) or before (asynchronously) the connection is established. |

#### ConnectOptionEnum

Specifies whether the **Open** method of a **Connection** object should return after (synchronously) or before (asynchronously) the connection is established.

| Constant   | Description |
| ---------- | ----------- |
| **adAsyncConnect** | Opens the connection asynchronously. The **ConnectComplete** event may be used to determine when the connection is available. |
| **adConnectUnspecified** | Default. Opens the connection synchronously. |

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Using the **Open** method on a **Connection** object establishes the physical connection to a data source. After this method successfully completes, the connection is live and you can issue commands against it and process the results.

Use the optional **ConnectionString** argument to specify either a connection string containing a series of argument = value statements separated by semicolons, or a file or directory resource identified with a URL. The **ConnectionString** property automatically inherits the value used for the **ConnectionString** argument. Therefore, you can either set the **ConnectionString** property of the **Connection** object before opening it, or use the **ConnectionString** argument to set or override the current connection parameters during the **Open** method call.

If you pass user and password information both in the **ConnectionString** argument and in the optional UserID and Password arguments, the UserID and **Password** arguments will override the values specified in **ConnectionString**.

When you have concluded your operations over an open **Connection**, use the **Close** method to free any associated system resources. Closing an object does not remove it from memory; you can change its property settings and use the **Open** method to open it again later. To completely eliminate an object from memory, release the object variable.

#### Remote Data Service Usage

When used on a client-side **Connection** object, the **Open** method doesn't actually establish a connection to the server until a Recordset is opened on the **Connection** object.

**Note**: URLs using the http scheme will automatically invoke the Microsoft OLE DB Provider for Internet Publishing.

#### Example

```
' // Set the connection string
pConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"
' // Open the connection
pConnection.Open
```
' --> Alternate way <--
```
DIM cbsConStr AS CBSTR = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"
pConnection.Open cbsConStr
```
' --> Alternate way <--
```
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"
```

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Create a Connection object
DIM pConnection AS CAdoConnection
' // Create a Recordset object
DIM pRecordset AS CAdoRecordset

' // Open the connection
DIM cbsConStr AS CBSTR = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"
pConnection.Open cbsConStr

' // Open the recordset
DIM cvSource AS CVAR = "SELECT TOP 20 * FROM Authors ORDER BY Author"
DIM hr AS HRESULT = pRecordset.Open(cvource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

' // Parse the recordset
DO
   ' // While not at the end of the recordset...
   IF pRecordset.EOF THEN EXIT DO
   ' // Get the content of the "Author" column
   SCOPE
      DIM cvRes AS CVAR = pRecordset.Collect("Author")
      PRINT cvRes
   END SCOPE
   ' // Fetch the next row
   pRecordset.MoveNext
LOOP
```

# <a name="OpenSchema"></a>OpenSchema

Obtains database schema information from the provider.

```
FUNCTION OpenSchema (BYVAL Schema AS SchemaEnum, _
   BYVAL Restrictions AS VARIANT = TYPE(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
   BYVAL SchemaID AS VARIANT = TYPE(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS Afx_ADORecordset PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *Schema* | Any **SchemaEnum** value that represents the type of schema query to run. |
| *Restrictions* | Optional. An array of query constraints for each *QueryType* option, as listed in **SchemaEnum**. |
| *SchemaID* | Optional. The GUID for a provider-schema query not defined by the OLE DB specification. This parameter is required if *QueryType* is set to **adSchemaProviderSpecific**; otherwise, it is not used. |

#### Return value

An ADO **Recordset** object reference.

#### Remarks

The **OpenSchema** method returns self-descriptive information about the data source, such as what tables are in the data source, the columns in the tables, and the data types supported.

The *QueryType* argument is a GUID that indicates the columns (schemas) returned. The OLE DB specification has a full list of schemas.

The **Criteria** argument limits the results of a schema query. **Criteria** specifies an array of values that must occur in a corresponding subset of columns, called constraint columns, in the resulting **Recordset**.

The constant **adSchemaProviderSpecific** is used for the *QueryType* argument if the provider defines its own nonstandard schema queries outside those listed above. When this constant is used, the **SchemaID** argument is required to pass the GUID of the schema query to execute. If *QueryType* is set to **adSchemaProviderSpecific** but *SchemaID* is not provided, an error will result.

Providers are not required to support all of the OLE DB standard schema queries. Specifically, only **adSchemaTables**, **adSchemaColumns**, and **adSchemaProviderTypes** are required by the OLE DB specification. However, the provider is not required to support the **Criteria** constraints listed above for those schema queries.

#### Remote Data Service Usage

The **OpenSchema** method is not available on a client-side **Connection** object.

When not using client side cursors, retrieving the "ORDINAL_POSITION" of a column schema in ADO returns a variant of type VT_R8 in MDAC 2.7 and later while the type of used in MDAC 2.6 was VT_I4. Programs written for MDAC 2.6 that only look for a variant returned of type VT_I4 would get a zero for every ordinal if run under MDAC 2.7 and later without modification. This change was made because the data type that OLE DB returns is DBTYPE_UI4, and in the signed VT_I4 type there is not enough room to contain all possible values without possibly truncation occurring and thereby causing a loss of data.

# <a name="Properties"></a>Properties

Returns a reference to the **Properties** collection.

```
PROPERTY Properties () AS Afx_ADOProperties PTR
```

#### Return value

An Afx_ADOProperties object reference.

#### Usage example

```
DIM pProperties AS Afx_ADOProperties
pProperties = pConnection.Properties
```

# <a name="Provider"></a>Provider

Indicates the name of the provider for a Connection object.

```
PROPERTY Provider () AS CBSTR
PROPERTY Provider (BYREF cbsProvider AS CBSTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsProvider* | The provider name. |

#### Return value

The provider name.

#### Remarks

Use the **Provider** property to set or return the name of the provider for a connection. This property can also be set by the contents of the **ConnectionString** property or the **ConnectionString** argument of the Open method; however, specifying a provider in more than one place while calling the **Open** method can have unpredictable results. If no provider is specified, the property will default to MSDASQL (Microsoft OLE DB Provider for ODBC).

The **Provider** property is read/write when the connection is closed and read-only when it is open. The setting does not take effect until you either open the **Connection** object or access the **Properties** collection of the **Connection** object. If the setting is not valid, an error occurs.

#### Usage example

```
DIM cbsProvider AS CBSTR = pConnection.Provider
```

# <a name="RollbackTrans"></a>RollbackTrans

Saves any changes and ends the current transaction. It may also start a new transaction.

```
FUNCTION RollbackTrans () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

Use **BeginTrans**, **CommitTrans** and **RollbackTrans** with a **Connection** object when you want to save or cancel a series of changes made to the source data as a single unit. For example, to transfer money between accounts, you subtract an amount from one and add the same amount to the other. If either update fails, the accounts no longer balance. Making these changes within an open transaction ensures that either all or none of the changes go through.

Not all providers support transactions. Verify that the provider-defined property "Transaction DDL" appears in the **Connection** object's **Properties** collection, indicating that the provider supports transactions. If the provider does not support transactions, calling one of these methods will return an error.

After you call the **BeginTrans method**, the provider will no longer instantaneously commit changes you make until you call **CommitTrans** or **RollbackTrans** to end the transaction.

For providers that support nested transactions, calling the **BeginTrans** method within an open transaction starts a new, nested transaction. The return value indicates the level of nesting: a return value of "1" indicates you have opened a top-level transaction (that is, the transaction is not nested within another transaction), "2" indicates that you have opened a second-level transaction (a transaction nested within a top-level transaction), and so forth. Calling **CommitTrans** or **RollbackTrans** affects only the most recently opened transaction; you must close or roll back the current transaction before you can resolve any higher-level transactions.

Calling the **CommitTrans** method saves changes made within an open transaction on the connection and ends the transaction. Calling the **RollbackTrans** method reverses any changes made within an open transaction and ends the transaction. Calling either method when there is no open transaction generates an error.

Depending on the **Connection** object's **Attributes** property, calling either the **CommitTrans** or **RollbackTrans** methods may automatically start a new transaction. If the **Attributes** property is set to **adXactCommitRetaining**, the provider automatically starts a new transaction after a **CommitTrans** call. If the **Attributes** property is set to **adXactAbortRetaining**, the provider automatically starts a new transaction after a **RollbackTrans** call.

#### Remote Data Service

The **BeginTrans**, **CommitTrans**, and **RollbackTrans** methods are not available on a client-side **Connection** object.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Open the recordset
DIM pRecordset AS CAdoRecordset
DIM cvSource AS CVAR = "SELECT * FROM Authors"
pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

' // Begin a transaction
pConnection.BeginTrans

' // Parse the recordset
DO
   ' // While not at the end of the recordset...
   IF pRecordset.EOF THEN EXIT DO
   ' // Get the content of the "Author" column
   DIM cvRes AS CVAR = pRecordset.Collect("Year Born")
   IF cvRes.ValInt = 1947 THEN pRecordset.Collect("Year Born") = 1900
   ' // Fetch the next row
   IF pRecordset.MoveNext <> S_OK THEN EXIT DO
LOOP
' // Commit the transaction
'pConnection.CommitTrans
' // Rollback the transaction because this is a demo
pConnection.RollbackTrans
```

# <a name="State"></a>State

Indicates if a **Connection** is open or closed.

```
PROPERTY State () AS LONG
```

#### Return value

The current connection state.

#### Remarks

You can use the **State** property to determine the current state of a given object at any time.

The object's **State** property can have a combination of values. For example, if a statement is executing, this property will have a combined value of **adStateOpen** and **adStateExecuting**.

#### ObjectStateEnum

Specifies whether the Open method of a Connection object should return after (synchronously) or before (asynchronously) the connection is established.

| Constant   | Description |
| ---------- | ----------- |
| **adStateClosed** | Indicates that the object is closed. |
| **adStateOpen** | Indicates that the object is open. |
| **adStateConnecting** | Indicates that the object is connecting. |
| **adStateExecuting** | Indicates that the object is executing a command. |
| adStateFetching | Indicates that the rows of the object are being retrieved. |

#### Usage example

```
IF pConnection.State = adStateOpen THEN pConnection.Close
```

# <a name="Version"></a>Version

Indicates the ADO version number.

```
PROPERTY Version () AS CBSTR
```

#### Return value

The ADO version number.

#### Remarks

Use the **Version** property to return the version number of the ADO implementation.

The version of the provider will be available as a dynamic property in the **Properties** collection.

#### Example

```
DIM cbsVersion AS CBSTR = pConnection.Version
```

