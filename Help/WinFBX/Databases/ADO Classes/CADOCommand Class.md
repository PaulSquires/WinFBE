# CADOCommand Class

Defines a specific command that you intend to execute against a data source.

**Include file**: CADOCommand.inc (include CADODB.inc)

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [ActiveConnection](#ActiveConnection) | Determines the **Connection** object over which the specified **Command** object will execute. |
| [Cancel](#Cancel) | Cancels execution of a pending, asynchronous method call. |
| [CommandStream](#CommandStream) | Sets or returns the stream used as the input for a **Command** object. |
| [CommandText](#CommandText) | Sets or returns an string value that contains a provider command, such as an SQL statement, a table name, a relative URL, or a stored procedure call. |
| [CommandTimeout](#CommandTimeout) | Sets or returns a Long value that indicates, in seconds, how long to wait for a command to execute. |
| [CommandType](#CommandType) | Sets or returns one or more **CommandTypeEnum** values. |
| [CreateParameter](#CreateParameter) | Creates a new **Paramete**r object with the specified properties. |
| [Dialect](#Dialect) | Indicates the dialect of the **CommandText** or **CommandStream** properties. |
| [Execute](#Execute) | Executes the query, SQL statement, or stored procedure specified in the **CommandText** or **CommandStream** property. |
| [GetErrorInfo](#GetErrorInfo) | Returns information about ADO errors. |
| [Name](#Name) | Sets or returns an string value that indicates the name of a **Command** object. |
| [NamedParameters](#NamedParameters) | Indicates whether parameter names should be passed to the provider. |
| [Parameters](#Parameters) | Returns a reference to the **Parameters** collection. |
| [Prepared](#Prepared) | Sets or returns a Boolean value that, if set to True, indicates that the command should be prepared. |
| [Properties](#Properties) | Returns a reference to the **Properties** collection. |
| [State](#State) | Indicates if the **Command** object is open or closed. |

#### Remarks

Use a **Command** object to query a database and return records in a **Recordset** object, to execute a bulk operation, or to manipulate the structure of a database. Depending on the functionality of the provider, some **Command** collections, methods, or properties may generate an error when referenced.

With the collections, methods, and properties of a Command object, you can do the following:

* Define the executable text of the command (for example, an SQL statement) with the **CommandText** property. Alternatively, for command or query structures other than simple strings (for example, XML template queries) define the command with the **CommandStream** property.
* Optionally, indicate the command dialect used in the **CommandText** or **CommandStream** with the **Dialect** property.
* Define parameterized queries or stored-procedure arguments with **Parameter** objects and the **Parameters** collection.
* Indicate whether parameter names should be passed to the provider with the **NamedParameters** property.
* Execute a command and return a **Recordset** object if appropriate with the **Execute** method.
* Specify the type of command with the **CommandType** property prior to execution to optimize performance.
* Control whether the provider saves a prepared (or compiled) version of the command prior to execution with the **Prepared** property.
* Set the number of seconds that a provider will wait for a command to execute with the **CommandTimeout** property.
* Associate an open connection with a **Command** object by setting its **ActiveConnection** property.
* Set the Name property to identify the **Command** object as a method on the associated **Connection** object.
* Pass a **Command** object to the **Source** property of a **Recordset** in order to obtain data.
* Access provider-specific attributes with the **Properties** collection.

**Note**: To execute a query without using a **Command** object, pass a query string to the Execute method of a **Connection** object or to the **Open** method of a **Recordset** object. However, a **Command** object is required when you want to persist the command text and re-execute it, or use query parameters.

To create a **Command** object independently of a previously defined **Connection** object, set its **ActiveConnection** property to a valid connection string. ADO still creates a **Connection** object, but it doesn't assign that object to an object variable. However, if you are associating multiple **Command** objects with the same connection, you should explicitly create and open a **Connection** object; this assigns the **Connection** object to an object variable. Make sure the **Connection** object was opened successfully before you assign it to the **Command** object's **ActiveConnection** property, because assigning a closed **Connection** object causes an error. If you do not set the **Command** object's **ActiveConnection** property to this object variable, ADO creates a new **Connection** object for each **Command** object, even if you use the same connection string.

To execute a **Command**, simply call it by its **Name** property on the associated **Connection** object. The **Command** must have its **ActiveConnection** property set to the **Connection** object. If the **Command** has parameters, pass their values as arguments to the method.

If two or more **Command** objects are executed on the same connection and either **Command** object is a stored procedure with output parameters, an error occurs. To execute each **Command** object, use separate connections or disconnect all other **Command** objects from the connection.

# CADOParameters Class

Contains all the **Parameter** objects of a **Command** object.

**Include file**: CADOParameters.inc (include CADODB.inc)

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [Append](#Append) | Appends an object to the **Parameters** collection. |
| [Count](#Count) | Retrieves the number of objects of the **Parameters** collection. |
| [Delete_](#Delete_) | Deletes an object from the **Parameters** collection. |
| [Item](#Item) | Indicates a specific member of the **Parameters** collection, by name or ordinal number. |
| [Refresh](#Refresh) | Refreshes the contents of the **Parameters** collection. |

#### Remarks

A **Command** object has a **Parameters** collection made up of **Parameter** objects.

Using the **Refresh** method on a **Command** object's **Parameters** collection retrieves provider parameter information for the stored procedure or parameterized query specified in the **Command** object. Some providers do not support stored procedure calls or parameterized queries; calling the **Refresh** method on the **Parameters** collection when using such a provider will return an error.

If you have not defined your own **Parameter** objects and you access the **Parameters** collection before calling the **Refresh** method, ADO will automatically call the method and populate the collection for you.

You can minimize calls to the provider to improve performance if you know the properties of the parameters associated with the stored procedure or parameterized query you wish to call. Use the **CreateParameter** method to create **Parameter** objects with the appropriate property settings and use the Append method to add them to the **Parameters** collection. This lets you set and return parameter values without having to call the provider for the parameter information. If you are writing to a provider that does not supply parameter information, you must manually populate the **Parameters** collection using this method to be able to use parameters at all. Use the **Delete_** method to remove **Parameter** objects from the **Parameters** collection if necessary.

The objects in the **Parameters** collection of a **Recordset** go out of scope (therefore becoming unavailable) when the **Recordset** is closed.

# CADOParameter Class

Represents a parameter or argument associated with a **Command** object based on a parameterized query or stored procedure.

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [AppendChunk](#AppendChunk) | Appends data to a **Parameter** object. |
| [Attach](#Attach) | Attaches a reference to an ADO **Parameter** object to the class. |
| [Attributes](#Attributes) | Indicates one or more characteristics of an object. |
| [Direction](#Direction) | Indicates whether the **Parameter** represents an input parameter, an output parameter, an input and an output parameter, or if the parameter is the return value from a stored procedure. |
| [Name](#Name2) | Sets or returns an string value that indicates the name of a **Parameter** object. |
| [NumericScale](#NumericScale) | Indicates the scale of numeric values in a **Parameter** object. |
| [Precision](#Precision) | Indicates the degree of precision for numeric values in a **Parameter** object. |
| [Properties](#Properties2) | Returns a reference to the **Properties** collection. |
| [Size](#Size) | Sets or returns a Long value that indicates the maximum size in bytes or characters of a value in a **Parameter** object |
| [Type_](#Type_) | Indicates the operational type or data type of a **Parameter** object. |
| [Value](#Value) | Indicates the value assigned to a **Parameter** object. |

#### Remarks

Many providers support parameterized commands. These are commands in which the desired action is defined once, but variables (or parameters) are used to alter some details of the command. For example, an SQL SELECT statement could use a parameter to define the matching criteria of a WHERE clause, and another to define the column name for a SORT BY clause.

**Parameter** objects represent parameters associated with parameterized queries, or the in/out arguments and the Return values of stored procedures. Depending on the functionality of the provider, some collections, methods, or properties of a Parameter object may not be available.

With the collections, methods, and properties of a **Parameter** object, you can do the following:

* Set or return the name of a parameter with the **Name** property.
+ Set or return the value of a parameter with the **Value** property. **Value** is the default property of the **Parameter** object.
* Set or return parameter characteristics with the **Attributes**, **Direction**, **Precision**, **NumericScale**, **Size**, and **Type_** properties.
* Pass long binary or character data to a parameter with the **AppendChunk** method.
* Access provider-specific attributes with the **Properties** collection.

If you know the names and properties of the parameters associated with the stored procedure or parameterized query you wish to call, you can use the **CreateParameter** method to create **Parameter** objects with the appropriate property settings and use the **Append** method to add them to the **Parameters** collection. This lets you set and return parameter values without having to call the **Refresh** method on the **Parameters** collection to retrieve the parameter information from the provider, a potentially resource-intensive operation.

The **Parameter** object is not safe for scripting.

# <a name="ActiveConnection"></a>ActiveConnection (CADOCommand)

Determines the **Connection** object over which the specified **Command** object will execute.

```
PROPERTY ActiveConnection (BYREF vConn AS CVAR)
PROPERTY ActiveConnection (BYVAL pconn AS Afx_ADOConnection PTR)
PROPERTY ActiveConnection (BYREF pconn AS CAdoConnection)
PROPERTY ActiveConnection () AS Afx_ADOCOnnection PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *pconn* | A reference to the **Connection** object or to the **CADOConnection** class. |
| vConn | An string containing a definition for a connection if the connection is closed, or a VARIANT of type VT_DISPATCH containing the current **Connection** object if the connection is open. |

#### Return value

An ADO **Connection** object reference. You must release it calling the **Release** method of the interface when no longer needed.

#### Remarks

If you attempt to call the **Execute** method on a **Command** object before setting this property to an open **Connection** object or valid connection string, an error occurs.

If a **Connection** object is assigned to the **ActiveConnection** property, the object must be opened. Assigning a closed **Connection** object causes an error.

Setting the **ActiveConnection** property to a null reference disassociates the **Command** object from the current **Connection** and causes the provider to release any associated resources on the data source. You can then associate the **Command** object with the same or another **Connection** object. Some providers allow you to change the property setting from one **Connection** to another, without having to first set the property to null.

If the **Parameters** collection of the **Command** object contains parameters supplied by the provider, the collection is cleared if you set the **ActiveConnection** property to null or to another **Connection** object. If you manually create **Parameter** objects and use them to fill the **Parameters** collection of the **Command** object, setting the **ActiveConnection** property to null or to another **Connection** object leaves the **Parameters** collection intact.

Closing the **Connection** object with which a **Command** object is associated sets the **ActiveConnection** property to null. Setting this property to a closed **Connection** object generates an error.

#### Example

```
' // Create a Connection object
DIM pConnection AS CAdoConnection
' // Create a Command object
DIM pCommand AS CAdoCommand
' // Open the connection
DIM cvConStr AS CVAR = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"
pConnection.Open cvConStr
' // Set the active connection
pCommand.ActiveConnection = pConnection
```

# <a name="Cancel"></a>Cancel (CADOCommand)

Cancels execution of a pending, asynchronous method call.

```
FUNCTION Cancel() AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use the **Cancel** method to terminate execution of an asynchronous method call (that is, a method invoked with the **adAsyncConnect**, **adAsyncExecute**, or **adAsyncFetch** option).

For a **Command** object, the last asynchronous call to the **Execute** method is terminated.

# <a name="CommandStream"></a>CommandStream (CADOCommand)

Sets or returns the stream used as the input for a **Command** object. The format for this stream is provider-specific, see your provider's documentation for details. This property is similar to the **CommandText** property, used to specify a string for the input of a **Command**.

```
POPERTY CommandStream () AS Afx_AdoStream
POPERTY CommandStream (BYVAL pStream AS Afx_AdoStream)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pStream* | Object variable containing a reference to the stream. |

#### Return value

An Stream object reference.

#### Remarks

**CommandStream** and **CommandText** are mutually exclusive. When the user sets the **CommandStream** property, the **CommandText** property will be set to the empty string (""). If the user sets the **CommandText** property, the **CommandStream** property will be set to null.

The behaviors of the **Command.Parameters.Refresh** and **Command.Prepare** methods are defined by the provider. The values of parameters in a stream may not be refreshed.

The input stream is not available to other ADO objects that return the source of a **Command**. For example, if the **Source** property of a **Recordset** is set to a **Command** object that has a stream as its input, **Recordset.Source** continues to return the **CommandText** property, which contains an empty string (""), instead of the stream contents of the **CommandStream** property.

When using a command stream (as specified by **CommandStream**), the only valid **CommandTypeEnum** values for the **CommandType** property are **adCmdText** and **adCmdUnknown**. Any other value causes an error.

# <a name="CommandText"></a>CommandText (CADOCommand)

Sets or returns an string value that contains a provider command, such as an SQL statement, a table name, a relative URL, or a stored procedure call. Default is "" (zero-length string).

```
PROPERTY CommandText () AS CBSTR
PROPERTY CommandText (BYVAL cbsText AS CBSTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsText* | A provider command, a table name, a relative URL, or a stored procedure call. |

#### Return value

A provider command, a table name, a relative URL, or a stored procedure call.

#### Remarks

Use the **CommandText** property to set or return the text of a command represented by a **Command** object. Usually this will be an SQL statement, but can also be any other type of command statement recognized by the provider, such as a stored procedure call. An SQL statement must be of the particular dialect or version supported by the provider's query processor.

If the **Prepared** property of the **Command** object is set to True and the **Command** object is bound to an open connection when you set the **CommandText** property, ADO prepares the query (that is, a compiled form of the query that is stored by the provider) when you call the **Execute** or **Open** methods.

Depending on the **CommandType** property setting, ADO may alter the **CommandText** property. You can read the **CommandText** property at any time to see the actual command text that ADO will use during execution.

Use the **CommandText** property to set or return a relative URL that specifies a resource, such as a file or directory. The resource is relative to a location specified explicitly by an absolute URL, or implicitly by an open **Connection** object.

**Note**: URLs using the http scheme will automatically invoke the Microsoft OLE DB Provider for Internet Publishing. 

#### Example

```
' // Create a Connection object
DIM pConnection AS CAdoConnection
' // Create a Command object
DIM pCommand AS CAdoCommand
' // Open the connection
DIM cvConStr AS CVAR = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"
pConnection.Open cvConStr
' // Set the active connection
pCommand.ActiveConnection = pConnection
' // Set the CommandText property
DIM cbsSqlStr AS CBSTR = "SELECT TOP 20 * FROM Authors ORDER BY Author"
pCommand.CommandText = cbsSqlStr
```

# <a name="CommandTimeout"></a>CommandTimeout (CADOCommand)

Sets or returns a Long value that indicates, in seconds, how long to wait for a command to execute. Default is 30.

```
PROPERTY CommandTimeout () AS LONG
PROPERTY CommandTimeout (BYVAL lTimeout AS LONG) 
```

| Parameter  | Description |
| ---------- | ----------- |
| *lTimeout* | Timeout value, in seconds. |

#### Return value

Timeout value, in seconds.

#### Remarks

Use the **CommandTimeout** property on a **Connection** object or **Command** object to allow the cancellation of an **Execute** method call, due to delays from network traffic or heavy server use. If the interval set in the **CommandTimeout** property elapses before the command completes execution, an error occurs and ADO cancels the command. If you set the property to zero, ADO will wait indefinitely until the execution is complete. Make sure the provider and data source to which you are writing code support the **CommandTimeout** functionality.

The **CommandTimeout** setting on a **Connection** object has no effect on the **CommandTimeout** setting on a **Command** object on the same **Connection**; that is, the **Command** object's **CommandTimeout** property does not inherit the value of the **Connection** object's **CommandTimeout** value.

On a **Connection** object, the **CommandTimeout** property remains read/write after the **Connection** is opened.

# <a name="CommandType"></a>CommandType (CADOCommand)

Sets or returns one or more **CommandTypeEnum** values.

**Note**: Do not use the **CommandTypeEnum** values of **adCmdFile** or **adCmdTableDirect** with **CommandType**. These values can only be used as options with the **Open** and **Requery** methods of a **Recordset**.

```
PROPERTY CommandType (BYVAL lCmdType AS CommandTypeEnum)
PROPERTY CommandType () AS CommandTypeEnum
```

| Parameter  | Description |
| ---------- | ----------- |
| *lCmdType* | One of more **CommandTypeEnum** values. |

#### Return value

A **CommandTypeEnum** value.

#### CommandTypeEnum

| Constant   | Description |
| ---------- | ----------- |
| **adCmdUnspecified** | Does not specify the command type argument. |
| **adCmdText** | Evaluates **CommandText** as a textual definition of a command or stored procedure call. |
| **adCmdTable** | Evaluates **CommandText** as a table name whose columns are all returned by an internally generated SQL query. |
| **adCmdStoredProc** | Evaluates **CommandText** as a stored procedure name. |
| **adCmdUnknown** | Default. Indicates that the type of command in the **CommandText** property is not known. |
| **adCmdFile** | Evaluates CommandText as the file name of a persistently stored **Recordset**. Used with **Recordset** **Open** or **Requery** only. |
| **adCmdTableDirect** | Evaluates **CommandText** as a table name whose columns are all returned. Used with **Recordset** **Open** or **Requery** only. To use the **Seek** method, the **Recordset** must be opened with **adCmdTableDirect**. This value cannot be combined with the **ExecuteOptionEnum** value **adAsyncExecute**. |

#### Remarks

Use the **CommandType** property to optimize evaluation of the **CommandText** property.

If the **CommandType** property value equals **adCmdUnknown** (the default value), you may experience diminished performance because ADO must make calls to the provider to determine if the **CommandText** property is an SQL statement, a stored procedure, or a table name. If you know what type of command you're using, setting the **CommandType** property instructs ADO to go directly to the relevant code. If the **CommandType** property does not match the type of command in the **CommandText** property, an error occurs when you call the **Execute** method.

# <a name="CreateParameter"></a>CreateParameter (CADOCommand)

Creates a new **Parameter** object with the specified properties.

```
FUNCTION CreateParameter (BYREF cbsName AS CBSTR = "", BYVAL nType AS DataTypeEnum = 0, _
   BYVAL Direction AS ParameterDirectionEnum = adParamInput, BYVAL Size AS LONG = 0, _
   BYREF cvValue AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS Afx_ADOParameter PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsName* | Optional. The name of the **Parameter** object. |
| *nType* | Optional. A **DataTypeEnum** value that specifies the data type of the **Parameter** object. |
| *Direction* | Optional. A **ParameterDirectionEnum** value that specifies the type of **Parameter** object. |
| *Size* | Optional. The maximum length for the parameter value in characters or bytes. |
| *cvValue* | Optional. The value for the **Parameter** object. |

#### Return value

Reference to an ADO **Parameter** object.

#### ParameterDirectionEnum

Specifies whether the **Parameter** represents an input parameter, an output parameter, both an input and an output parameter, or the Return value from a stored procedure.

| Constant   | Description |
| ---------- | ----------- |
| **adParamInput** | Default. Indicates that the parameter represents an input parameter. |
| **adParamInputOutput** | Indicates that the parameter represents both an input and output parameter. |
| **adParamOutput** | Indicates that the parameter represents an output parameter. |
| **adParamReturnValue** | Indicates that the parameter represents a Return value. |
| **adParamUnknown** | Indicates that the parameter direction is unknown. |

#### DataTypeEnum

Specifies the data type of a **Field**, **Parameter**, or **Property**. The corresponding OLE DB type indicator is shown in parentheses in the description column of the following table.

| Constant   | Description |
| ---------- | ----------- |
| **AdArray** | A flag value, always combined with another data type constant, that indicates an array of that other data type. (Does not apply to ADOX.) |
| **adBigInt** | Indicates an eight-byte signed integer (DBTYPE_I8). |
| **adBinary** | Indicates a binary value (DBTYPE_BYTES). |
| **adBoolean** | Indicates a boolean value (DBTYPE_BOOL). |
| **adBSTR** | Indicates a null-terminated character string (Unicode) (DBTYPE_BSTR). |
| **adChapter** | Indicates a four-byte chapter value that identifies rows in a child rowset (DBTYPE_HCHAPTER). |
| **adChar** | Indicates a string value (DBTYPE_STR). |
| **adCurrency** | Indicates a currency value (DBTYPE_CY). Currency is a fixed-point number with four digits to the right of the decimal point. It is stored in an eight-byte signed integer scaled by 10,000. |
| **adDate** | Indicates a date value (DBTYPE_DATE). A date is stored as a double, the whole part of which is the number of days since December 30, 1899, and the fractional part of which is the fraction of a day. |
| **adDBDate** | Indicates a date value (yyyymmdd) (DBTYPE_DBDATE). |
| **adDBTime** | Indicates a time value (hhmmss) (DBTYPE_DBTIME). |
| **adDBTimeStamp** | Indicates a date/time stamp (yyyymmddhhmmss plus a fraction in billionths) (DBTYPE_DBTIMESTAMP). |
| **adDecimal** | Indicates an exact numeric value with a fixed precision and scale (DBTYPE_DECIMAL). |
| **adDouble** | Indicates a double-precision floating-point value (DBTYPE_R8). |
| **adEmpty** | Specifies no value (DBTYPE_EMPTY). |
| **adError** | Indicates a 32-bit error code (DBTYPE_ERROR). |
| **adFileTime** | Indicates a 64-bit value representing the number of 100-nanosecond intervals since January 1, 1601 (DBTYPE_FILETIME). |
| **adGUID** | Indicates a globally unique identifier (GUID) (DBTYPE_GUID). |
| **adIDispatch** | Indicates a pointer to an IDispatch interface on a COM object (DBTYPE_IDISPATCH). Note: This data type is currently not supported by ADO. Usage may cause unpredictable results. |
| **adInteger** | Indicates a four-byte signed integer (DBTYPE_I4). |
| **adIUnknown** | Indicates a pointer to an IUnknown interface on a COM object (DBTYPE_IUNKNOWN). Note: This data type is currently not supported by ADO. Usage may cause unpredictable results. |
| **adLongVarBinary** | Indicates a long binary value. |
| **adLongVarChar** | Indicates a long string value. |
| **adLongVarWChar** | Indicates a long null-terminated Unicode string value. |
| **adNumeric** | Indicates an exact numeric value with a fixed precision and scale (DBTYPE_NUMERIC). |
| **adPropVariant** | Indicates an Automation PROPVARIANT (DBTYPE_PROP_VARIANT). |
| **adSingle** | Indicates a single-precision floating-point value (DBTYPE_R4). |
| **adSmallInt** | Indicates a two-byte signed integer (DBTYPE_I2). |
| **adTinyInt** | Indicates a one-byte signed integer (DBTYPE_I1). |
| **adUnsignedBigInt** | Indicates an eight-byte unsigned integer (DBTYPE_UI8). |
| **adUnsignedInt** | Indicates a four-byte unsigned integer (DBTYPE_UI4). |
| **adUnsignedSmallInt** | Indicates a two-byte unsigned integer (DBTYPE_UI2). |
| **adUnsignedTinyInt** | Indicates a one-byte unsigned integer (DBTYPE_UI1). |
| **adVarBinary** | Indicates a binary value. |
| **adVarChar** | Indicates a string value. |
| **adVariant** | Indicates an Automation Variant (DBTYPE_VARIANT). Note: This data type is currently not supported by ADO. Usage may cause unpredictable results. |
| **adVarNumeric** | Indicates a numeric value. |
| **adVarWChar** | Indicates a null-terminated Unicode character string. |
| **adWChar** | Indicates a null-terminated Unicode character string (DBTYPE_WSTR). |

#### Remarks

Use the **CreateParameter** method to create a new **Parameter** object with a specified name, type, direction, size, and value. Any values you pass in the arguments are written to the corresponding **Parameter** properties.

This method does not automatically append the **Parameter** object to the **Parameters** collection of a **Command** object. This lets you set additional properties whose values ADO will validate when you append the **Parameter** object to the collection.

If you specify a variable-length data type in the *nType* argument, you must either pass a *nSize* argument or set the **Size** property of the **Parameter** object before appending it to the **Parameters** collection; otherwise, an error occurs.

If you specify a numeric data type (**adNumeric** or **adDecimal**) in the *nType* argument, then you must also set the **NumericScale** and **Precision** properties.

# <a name="Dialect"></a>Dialect (CADOCommand)

Indicates the dialect of the **CommandText** or **CommandStream** properties. The dialect defines the syntax and general rules that the provider uses to parse the string or stream.

The **Dialect** property contains a valid GUID that represents the dialect of the command text or stream. The default value for this property is {C8B521FB-5CF3-11CE-ADE5-00AA0044773D}, which indicates that the provider should choose how to interpret the command text or stream.

```
PROPERTY Dialect (BYREF cbsDialect AS CBSTR)
PROPERTY Dialect () AS CBSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsDialect* | The string representation of a valid GUID that represents the dialect of the command text or stream. |

#### Return value

The string representation of a valid GUID that represents the dialect of the command text or stream.

#### Remarks

ADO does not query the provider when the user reads the value of this property; it returns a string representation of the value currently stored in the **Command** object.

When the user sets the **Dialect** property, ADO validates the GUID and raises an error if the value supplied is not a valid GUID. See the documentation for your provider to determine the GUID values supported by the **Dialect** property.

# <a name="Execute"></a>Execute (CADOCommand)

Executes the query, SQL statement, or stored procedure specified in the **CommandText** or **CommandStream** property.

```
FUNCTION Execute (BYVAL RecordsAffected AS LONG PTR = NULL, _
   BYREF cvParameters AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
   BYVAL Options AS LONG = adCmdUnspecified) AS Afx_ADORecordset PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecordsAffected* | Optional. A LONG to which the provider returns the number of records that the operation affected. The **RecordsAffected** parameter applies only for action queries or stored procedures. **RecordsAffected** does not return the number of records returned by a result-returning query or stored procedure. To obtain this information, use the **RecordCount** property. The **Execute** method will not return the correct information when used with **adAsyncExecute**, simply because when a command is executed asynchronously, the number of records affected may not yet be known at the time the method returns. |
| *cvParameters* | Optional. A Variant array of parameter values used in conjunction with the input string or stream specified in **CommandText** or **CommandStream**. (Output parameters will not return correct values when passed in this argument.) |
| *Options* | Optional. A Long value that indicates how the provider should evaluate the **CommandText** or the **CommandStream** property of the **Command** object. Can be a bitmask value made using **CommandTypeEnum** and/or **ExecuteOptionEnum** values. For example, you could use **adCmdText** and **adExecuteNoRecords** in combination if you want to have ADO evaluate the value of the **CommandText** property as text, and indicate that the command should discard and not return any records that might be generated when the command text executes. |

**Note**: Use the **ExecuteOptionEnum** value **adExecuteNoRecords** to improve performance by minimizing internal processing. If **adExecuteStream** was specified, the options **adAsyncFetch** and **adAsynchFetchNonBlocking** are ignored. Do not use the **CommandTypeEnum** values of **adCmdFile** or **adCmdTableDirect** with **Execute**. These values can only be used as options with the **Open** and **Requery** methods of a **Recordset**.

#### CommandTypeEnum

| Constant   | Description |
| ---------- | ----------- |
| **adCmdUnspecified** | Does not specify the command type argument. |
| **adCmdText** | Evaluates **CommandText** as a textual definition of a command or stored procedure call. |
| **adCmdTable** | Evaluates **CommandText** as a table name whose columns are all returned by an internally generated SQL query. |
| **adCmdStoredProc** | Evaluates **CommandText** as a stored procedure name. |
| **adCmdUnknown** | Default. Indicates that the type of command in the **CommandText** property is not known. |
| **adCmdFile** | Evaluates CommandText as the file name of a persistently stored **Recordset**. Used with **Recordset** **Open** or **Requery** only. |
| **adCmdTableDirect** | Evaluates **CommandText** as a table name whose columns are all returned. Used with **Recordset** **Open** or **Requery** only. To use the **Seek** method, the **Recordset** must be opened with **adCmdTableDirect**. This value cannot be combined with the **ExecuteOptionEnum** value **adAsyncExecute**. |

#### ExecuteOptionEnum

Specifies how a command argument should be interpreted.

| Constant   | Description |
| ---------- | ----------- |
| **adAsyncExecute** | Indicates that the command should execute asynchronously. This value cannot be combined with the **CommandTypeEnum** value **adCmdTableDirect**. |
| **adAsyncFetch** | Indicates that the remaining rows after the initial quantity specified in the **CacheSize** property should be retrieved asynchronously. |
| **adAsyncFetchNonBlocking** | Indicates that the main thread never blocks while retrieving. If the requested row has not been retrieved, the current row automatically moves to the end of the file. If you open a **Recordset** from a **Stream** containing a persistently stored **Recordset**, **adAsyncFetchNonBlocking** will not have an effect; the operation will be synchronous and blocking. **adAsynchFetchNonBlocking** has no effect when the **CmdTableDirect** option is used to open the **Recordset**. |
| **adExecuteNoRecords** | Indicates that the command text is a command or stored procedure that does not return rows (for example, a command that only inserts data). If any rows are retrieved, they are discarded and not returned. **adExecuteNoRecords** can only be passed as an optional parameter to the Command or **Connection** **Execute** method. |
| **adExecuteStream** | Indicates that the results of a command execution should be returned as a stream. **adExecuteStream** can only be passed as an optional parameter to the **Command** **Execute** method. |
| **adExecuteRecord** | Indicates that the **CommandText** is a command or stored procedure that returns a single row which should be returned as a **Record** object. |
| **adOptionUnspecified** | Indicates that the command is unspecified. |

#### Return value

An ADO **Recordset** object reference.

#### Remarks

Using the **Execute** method on a **Command** object executes the query specified in the **CommandText** property or **CommandStream** property of the object.

Results are returned in a **Recordset** (by default) or as a stream of binary information. To obtain a binary stream, specify **adExecuteStream** in **Options**, then supply a stream by setting the "Output Stream" property. An ADO **Stream** object can be specified to receive the results, or another stream object such as the IIS Response object can be specified. If no stream was specified before calling **Execute** with **adExecuteStream**, an error occurs. The position of the stream on return from **Execute** is provider specific.

If the command is not intended to return results (for example, an SQL UPDATE query) the provider returns NULL as long as the option **adExecuteNoRecords** is specified; otherwise **Execute** returns a closed **Recordset**. Some application languages allow you to ignore this return value if no **Recordset** is desired.

**Execute** raises an error if the user specifies a value for **CommandStream** when the **CommandType** is **adCmdStoredProc**, **adCmdTable**, or **adCmdTableDirect**.

If the query has parameters, the current values for the **Command** object's parameters are used unless you override these with parameter values passed with the **Execute** call. You can override a subset of the parameters by omitting new values for some of the parameters when calling the **Execute method**. The order in which you specify the parameters is the same order in which the method passes them. For example, if there were four (or more) parameters and you wanted to pass new values for only the first and fourth parameters, you would pass an array of variants with the first and fourth elements holding the values and the second and third elements set to EMPTY as the **Parameters** argument.

**Note**: Output parameters will not return correct values when passed in the **Parameters** argument.

An **ExecuteComplete** event will be issued when this operation concludes.

**Note**: When issuing commands containing URLs, those using the http scheme will automatically invoke the Microsoft OLE DB Provider for Internet Publishing. 

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Create a Command object
DIM pCommand AS CAdoCommand

' // Set the active connection
pCommand.ActiveConnection = pConnection

' // Set the CommandText property
pCommand.CommandText = "SELECT TOP 20 * FROM Authors ORDER BY Author"

' // Create the recordset by executing a query and attaching
' // the resulting recordset to an instance of the CAdoRecordset class.
DIM pRecordset AS CAdoRecordset
pRecordset.Attach(pCommand.Execute)

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

# <a name="GetErrorInfo"></a>GetErrorInfo (CADOCommand)

Returns information about ADO errors.

```
FUNCTION GetErrorInfo (BYVAL nError AS HRESULT = 0) AS CBSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nError* | Optional. The error code. |

#### Return value

A description of the error(s).

# <a name="Name"></a>Name (CADOCommand)

Sets or returns an string value that indicates the name of a **Command** object. 

```
PROPERTY Name () AS CBSTR
PROPERTY Name (BYREF cbsName AS CBSTR) 
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsName* | The name of the **Command** object. |

#### Return value

The name of the **Command** object.

# <a name="NamedParameters"></a>NamedParameters (CADOCommand)

Indicates whether parameter names should be passed to the provider.

```
PROPERTY NamedParameters () AS BOOLEAN
PROPERTY NamedParameters (BYVAL fNamedParameters AS BOOLEAN)
```

| Parameter  | Description |
| ---------- | ----------- |
| *fNamedParameters* | True or False. |

#### Return value

True or False.

#### Remarks

When this property is true, ADO passes the value of the **Name** property of each parameter in the **Command**'s **Parameter** collection. The provider uses a parameter name to match parameters in the **CommandText** or **CommandStream** properties. If this property is false (the default), parameter names are ignored and the provider uses the order of parameters to match values to parameters in the **CommandText** or **CommandStream** properties.

# <a name="Parameters"></a>Parameters (CADOCommand)

Returns a reference to the **Parameters** collection.

```
PROPERTY Parameters () AS Afx_ADOParameters
```

#### Return value

A reference to the **Parameters** collection.

# <a name="Prepared"></a>Prepared (CADOCommand)

Sets or returns a Boolean value that, if set to True, indicates that the command should be prepared.

```
PROPERTY Prepared () AS BOOLEAN
PROPERTY Prepared (BYVAL fPrepared AS BOOLEAN)
```

| Parameter  | Description |
| ---------- | ----------- |
| *fPrepared* | True or False. |

#### Return value

True or False.

#### Remarks

Use the **Prepared** property to have the provider save a prepared (or compiled) version of the query specified in the **CommandText** property before a **Command** object's first execution. This may slow a command's first execution, but once the provider compiles a command, the provider will use the compiled version of the command for any subsequent executions, which will result in improved performance.

If the property is False, the provider will execute the **Command** object directly without creating a compiled version.

If the provider does not support command preparation, it may return an error as soon as this property is set to True. If it does not return an error, it simply ignores the request to prepare the command and sets the Prepared property to False.

# <a name="Properties"></a>Properties (CADOCommand)

Returns a reference to the **Properties** collection.

```
PROPERTY Properties () AS Afx_ADOProperties
```

#### Return value

A reference to the **Properties** collection.

# <a name="State"></a>State (CADOCommand)

Indicates if the **Command** object is open or closed.

```
PROPERTY State () AS LONG
```

#### Return value

Returns a Long value that can be an **ObjectStateEnum** value. The default value is **adStateClosed**.

#### ObjectStateEnum

Specifies whether the **Open** method of a **Connection** object should return after (synchronously) or before (asynchronously) the connection is established.

| Constant   | Description |
| ---------- | ----------- |
| **adStateClosed** | Indicates that the object is closed. |
| **adStateOpen** | Indicates that the object is open. |
| **adStateConnecting** | Indicates that the object is connecting. |
| **adStateExecuting** | Indicates that the object is executing a command. |
| **adStateFetching** | Indicates that the rows of the object are being retrieved. |

# <a name="Append"></a>Append (CADOParameters)

Appends an object to the **Parameters** collection.

```
FUNCTION Append (BYVAL pObject AS IDispatch PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pObject* | A reference to the object to be appended. |

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

You must set the **Type_** property of a **Parameter** object before appending it to the **Parameters** collection. If you select a variable-length data type, you must also set the **Size** property to a value greater than zero.

Describing parameters yourself minimizes calls to the provider and consequently improves performance when using stored procedures or parameterized queries. However, you must know the properties of the parameters associated with the stored procedure or parameterized query that you want to call.

Use the **CreateParameter** method to create **Parameter** objects with the appropriate property settings and use the **Append** method to add them to the **Parameters** collection. This lets you set and return parameter values without having to call the provider for the parameter information. If you are writing to a provider that does not supply parameter information, you must use this method to manually populate the **Parameters** collection in order to use parameters at all.

# <a name="Count"></a>Count (CADOParameters)

Retrieves the number of objects of the **Parameters** collection.

```
PROPERTY Count () AS LONG
```

#### Return value

The number of objects in the **Parameters** collection.

#### Remarks

Because numbering for members of a collection begins with zero, you should always code loops starting with the zero member and ending with the value of the **Count** property minus 1.

If the **Count** property is zero, there are no objects in the collection.

# <a name="Delete_"></a>Delete_ (CADOParameters)

Deletes an object from the **Parameters** collection.

```
FUNCTION Delete_ (BYREF cvIndex AS CVAR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cvIndex* | The name of the object you want to delete, or the object’s ordinal position (index) in the collection. |

#### Return value

S_OK (0) or am HRESULT code.

#### Remarks

Using the **Delete_** method on a collection lets you remove one of the objects in the collection. This method is available only on the **Parameters** collection of a **Command** object. You must use the **Parameter** object's **Name** property or its collection index when calling the **Delete_** method—an object variable is not a valid argument.

# <a name="Item"></a>Item (CADOParameters)

Indicates a specific member of the **Parameters** collection, by name or ordinal number.

```
PROPERTY Item (BYREF cvIndex AS CVAR) AS Afx_ADOParameter PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *cvIndex* | A Variant that evaluates to the ordinal number of an object in the **Parameters** collection. |

#### Return value

An ADO **Parameter** object reference.

#### Remarks

If **Item** cannot find an object in the collection corresponding to the **Index** argument, an error occurs.

# <a name="Refresh"></a>Refresh (CADOParameters)

Refreshes the contents of the **Parameters** collection.

```
FUNCTION Refresh () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Using the **Refresh** method on a **Command** object's **Parameters** collection retrieves provider-side parameter information for the stored procedure or parameterized query specified in the **Command** object. The collection will be empty for providers that do not support stored procedure calls or parameterized queries.

You should set the **ActiveConnection** property of the **Command** object to a valid **Connection** object, the **CommandText** property to a valid command, and the **CommandType** property to **adCmdStoredProc** before calling the **Refresh** method.

If you access the **Parameters** collection before calling the **Refresh** method, ADO will automatically call the method and populate the collection for you.

**Note**: If you use the **Refresh** method to obtain parameter information from the provider and it returns one or more variable-length data type Parameter objects, ADO may allocate memory for the parameters based on their maximum potential size, which will cause an error during execution. You should explicitly set the **Size** property for these parameters before calling the **Execute** method to prevent errors.

# <a name="AppendChunk"></a>AppendChunk (CADOParameter)

Appends data to a **Parameter** object.

```
FUNCTION AppendChunk (BYREF cvData AS CVAR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cvData* | A Variant that contains the data to append to the object. |

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use the **AppendChunk** method on a **Parameter** object to fill it with long binary or character data. In situations where system memory is limited, you can use the **AppendChunk** method to manipulate long values in portions rather than in their entirety.

If the **adParamLong** bit in the **Attributes** property of a **Parameter** object is set to true, you can use the **AppendChunk** method for that parameter.

The first **AppendChunk** call on a **Parameter** object writes data to the parameter, overwriting any existing data. Subsequent AppendChunk calls on a **Parameter** object add to existing parameter data. An **AppendChunk** call that passes a null value discards all of the parameter data.

# <a name="Attach"></a>Attach (CADOParameter)

Attaches a reference to an **Parameter** object to the class.

```
SUB Attach (BYVAL pParameter AS Afx_ADOParameter PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pParameter* | A reference to a **Parameter** object. |
| *fAddRef* | TRUE = increase the reference count; FALSE = don't increase the reference count. |

# <a name="Attributes"></a>Attributes (CADOParameter)

Indicates one or more characteristics of an object.

For a **Parameter** object, the **Attributes** property is read/write, and its value can be the sum of any one or more **ParameterAttributesEnum** values. The default is **adParamSigned**.

```
PROPERTY Attributes () AS LONG
PROPERTY Attributes (BYVAL lAttr AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *lAttr* | Can be the sum of any one or more **ParameterAttributesEnum** values. |

#### Return value

One or more **ParameterAttributesEnum** values.

#### ParameterAttributesEnum

Specifies the attributes of a Parameter object.

| Constant   | Description |
| ---------- | ----------- |
| **adParamSigned** | Indicates that the parameter accepts signed values. |
| **adParamNullable** | Indicates that the parameter accepts null values. |
| **adParamLong** | Indicates that the parameter accepts long binary data. |

#### Remarks

For a **Parameter** object, the **Attributes** property is read/write, and its value can be the sum of any one or more **ParameterAttributesEnum** values. The default is **adParamSigned**.

# <a name="Direction"></a>Direction (CADOParameter)

Indicates whether the **Parameter** represents an input parameter, an output parameter, an input and an output parameter, or if the parameter is the return value from a stored procedure.

Sets or returns a **ParameterDirectionEnum** value.

```
PROPERTY Direction () AS ParameterDirectionEnum
PROPERTY Direction (BYVAL lParmDirection AS ParameterDirectionEnum)
```

| Parameter  | Description |
| ---------- | ----------- |
| *lParmDirection* | A **ParameterDirectionEnum** value. |

#### Return value

One or more **ParameterDirectionEnum** values.

#### ParameterDirectionEnum

Specifies whether the Parameter represents an input parameter, an output parameter, both an input and an output parameter, or the return value from a stored procedure.

| Constant   | Description |
| ---------- | ----------- |
| **adParamInput** | Default. Indicates that the parameter represents an input parameter. |
| **adParamInputOutput** | Indicates that the parameter represents both an input and output parameter. |
| **adParamOutput** | Indicates that the parameter represents an output parameter. |
| **adParamReturnValue** | Indicates that the parameter represents a Return value. |
| **adParamUnknown** | Indicates that the parameter direction is unknown. |

#### Remarks

Use the **Direction** property to specify how a parameter is passed to or from a procedure. The **Direction** property is read/write; this allows you to work with providers that don't return this information or to set this information when you don't want ADO to make an extra call to the provider to retrieve parameter information.

Not all providers can determine the direction of parameters in their stored procedures. In these cases, you must set the **Direction** property before you execute the query.

# <a name="Name2"></a>Name (CADOParameter)

Sets or returns an string value that indicates the name of a **Parameter** object. 

```
PROPERTY Name () AS CBSTR
PROPERTY Name (BYVAL cbsName AS CBSTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsName* | The name of the **Parameter** object. |

#### Return value

The name of the **Parameter" object.

#### Remarks

For **Parameter** objects not yet appended to the **Parameters** collection, the **Name** property is read/write. For appended **Parameter** objects and all other objects, the **Name** property is read-only. Names do not have to be unique within a collection.

# <a name="NumericScale"></a>NumericScale (CADOParameter)

Indicates the scale of numeric values in a **Parameter** object.

Sets or returns a Byte value that indicates the number of decimal places to which numeric values will be resolved.

```
PROPERTY NumericScale () AS BYTE
PROPERTY NumericScale (BYVAL bScale AS BYTE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *bScale* | The number of decimal places to which numeric values will be resolved. |

#### Return value

The number of decimal places.

#### Remarks

Use the **umericScale** property to determine how many digits to the right of the decimal point will be used to represent values for a numeric **Parameter** object.

For **Parameter** objects, the **NumericScale** property is read/write.

# <a name="Precision"></a>Precision (CADOParameter)

Indicates the degree of precision for numeric values in a **Parameter** object.

Sets or returns a Byte value that indicates the maximum number of digits used to represent values.

```
PROPERTY Precision () AS BYTE
PROPERTY Precision (BYVAL bPrecision AS BYTE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *bPrecision* | The maximum number of digits used to represent values. |

#### Return value

The degress of precision.

#### Remarks

Use the **Precision** property to determine the maximum number of digits used to represent values for a numeric **Parameter** object.

# <a name="Properties2"></a>Properties (CADOParameter)

Returns a reference to the **Properties** collection.

```
PROPERTY Properties () AS Afx_ADOProperties
```

#### Return value

A **Properties** object reference.

# <a name="Size"></a>Size (CADOParameter)

Sets or returns a Long value that indicates the maximum size in bytes or characters of a value in a **Parameter** object.

```
PROPERTY Size () AS LONG
PROPERTY Size (BYVAL lSize AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *lSize* | The maximum size in bytes or characters of a value in a **Parameter** object. |

#### Return value

The maximum size.

#### Remarks

Use the **Size** property to determine the maximum size for values written to or read from the **Value** property of a **Parameter** object.

If you specify a variable-length data type for a **Parameter** object (for example, any string type, such as **adVarChar**), you must set the object's **Size** property before appending it to the **Parameters** collection; otherwise, an error occurs.

If you have already appended the **Parameter** object to the **Parameters** collection of a **Command** object and you change its type to a variable-length data type, you must set the **Parameter** object's **Size** property before executing the **Command** object; otherwise, an error occurs.

If you use the **Refresh** method to obtain parameter information from the provider and it returns one or more variable-length data type Parameter objects, ADO may allocate memory for the parameters based on their maximum potential size, which could cause an error during execution. To prevent an error, you should explicitly set the **Size** property for these parameters before executing the command.

The **Size** property is read/write.

# <a name="Type_"></a>Type_ (CADOParameter)

Indicates the operational type or data type of a **Parameter** object.

```
PROPERTY Type_ () AS DataTypeEnum
```

#### Return value

A **DataTypeEnum** value.

#### DataTypeEnum

Specifies the data type of a **Field**, **Parameter**, or **Property**. The corresponding OLE DB type indicator is shown in parentheses in the description column of the following table.

| Constant   | Description |
| ---------- | ----------- |
| **AdArray** | A flag value, always combined with another data type constant, that indicates an array of that other data type. (Does not apply to ADOX.) |
| **adBigInt** | Indicates an eight-byte signed integer (DBTYPE_I8). |
| **adBinary** | Indicates a binary value (DBTYPE_BYTES). |
| **adBoolean** | Indicates a boolean value (DBTYPE_BOOL). |
| **adBSTR** | Indicates a null-terminated character string (Unicode) (DBTYPE_BSTR). |
| **adChapter** | Indicates a four-byte chapter value that identifies rows in a child rowset (DBTYPE_HCHAPTER). |
| **adChar** | Indicates a string value (DBTYPE_STR). |
| **adCurrency** | Indicates a currency value (DBTYPE_CY). Currency is a fixed-point number with four digits to the right of the decimal point. It is stored in an eight-byte signed integer scaled by 10,000. |
| **adDate** | Indicates a date value (DBTYPE_DATE). A date is stored as a double, the whole part of which is the number of days since December 30, 1899, and the fractional part of which is the fraction of a day. |
| **adDBDate** | Indicates a date value (yyyymmdd) (DBTYPE_DBDATE). |
| **adDBTime** | Indicates a time value (hhmmss) (DBTYPE_DBTIME). |
| **adDBTimeStamp** | Indicates a date/time stamp (yyyymmddhhmmss plus a fraction in billionths) (DBTYPE_DBTIMESTAMP). |
| **adDecimal** | Indicates an exact numeric value with a fixed precision and scale (DBTYPE_DECIMAL). |
| **adDouble** | Indicates a double-precision floating-point value (DBTYPE_R8). |
| **adEmpty** | Specifies no value (DBTYPE_EMPTY). |
| **adError** | Indicates a 32-bit error code (DBTYPE_ERROR). |
| **adFileTime** | Indicates a 64-bit value representing the number of 100-nanosecond intervals since January 1, 1601 (DBTYPE_FILETIME). |
| **adGUID** | Indicates a globally unique identifier (GUID) (DBTYPE_GUID). |
| **adIDispatch** | Indicates a pointer to an IDispatch interface on a COM object (DBTYPE_IDISPATCH). Note: This data type is currently not supported by ADO. Usage may cause unpredictable results. |
| **adInteger** | Indicates a four-byte signed integer (DBTYPE_I4). |
| **adIUnknown** | Indicates a pointer to an IUnknown interface on a COM object (DBTYPE_IUNKNOWN). Note: This data type is currently not supported by ADO. Usage may cause unpredictable results. |
| **adLongVarBinary** | Indicates a long binary value. |
| **adLongVarChar** | Indicates a long string value. |
| **adLongVarWChar** | Indicates a long null-terminated Unicode string value. |
| **adNumeric** | Indicates an exact numeric value with a fixed precision and scale (DBTYPE_NUMERIC). |
| **adPropVariant** | Indicates an Automation PROPVARIANT (DBTYPE_PROP_VARIANT). |
| **adSingle** | Indicates a single-precision floating-point value (DBTYPE_R4). |
| **adSmallInt** | Indicates a two-byte signed integer (DBTYPE_I2). |
| **adTinyInt** | Indicates a one-byte signed integer (DBTYPE_I1). |
| **adUnsignedBigInt** | Indicates an eight-byte unsigned integer (DBTYPE_UI8). |
| **adUnsignedInt** | Indicates a four-byte unsigned integer (DBTYPE_UI4). |
| **adUnsignedSmallInt** | Indicates a two-byte unsigned integer (DBTYPE_UI2). |
| **adUnsignedTinyInt** | Indicates a one-byte unsigned integer (DBTYPE_UI1). |
| **adVarBinary** | Indicates a binary value. |
| **adVarChar** | Indicates a string value. |
| **adVariant** | Indicates an Automation Variant (DBTYPE_VARIANT). Note: This data type is currently not supported by ADO. Usage may cause unpredictable results. |
| **adVarNumeric** | Indicates a numeric value. |
| **adVarWChar** | Indicates a null-terminated Unicode character string. |
| **adWChar** | Indicates a null-terminated Unicode character string (DBTYPE_WSTR). |

# <a name="Value"></a>Value (CADOParameter)

Indicates the value assigned to a **Parameter** object.

```
PROPERTY Value () AS CVAR
PROPERTY Value (BYREF cvValue AS CVAR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cvValue* | The value of the **Parameter** object. |

#### Return value

The value of the **Parameter** object.

#### Remarks

Use the **Value** property to set or return parameter values with **Parameter** objects.

ADO allows setting and returning long binary data with the **Value** property.

For **Parameter** objects, ADO reads the **Value** property only once from the provider. If a command contains a **Paramete**r whose **Value** property is empty, and you create a **Recordset** from the command, ensure that you first close the **Recordset** before retrieving the **Value** property. Otherwise, for some providers, the **Value** property may be empty, and won't contain the correct value.
