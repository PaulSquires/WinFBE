# CADORecord Class

Represents a row from a **Recordset** or the data provider, or an object returned by a semi-structured data provider, such as a file or directory.

**Include file**: CADORecord.inc (include CADODB.inc)

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [ActiveConnection](#ActiveConnection) | Indicates to which **Connection** object the specified **Record** object currently belongs. |
| [Attach](#Attach) | Attaches a reference to an ADO **Record** object to the class. |
| [Cancel](#Cancel) | Cancels execution of a pending, asynchronous method call. |
| [Close](#Close) | Closes a **Record** object and any dependent objects. |
| [CopyRecord](#CopyRecord) | Copies a entity represented by a Record to another location. |
| [DeleteRecord](#DeleteRecord) | Deletes a the entity represented by a **Record**. |
| [Fields](#Fields) | Gets a reference to the Fields collection of a **Record** object. |
| [GetChildren](#GetChildren) | Returns a **Recordset** whose rows represent the children of a collection **Record**. |
| [GetErrorInfo](#GetErrorInfo) | Returns information about ADO errors. |
| [Mode](#Mode) | Sets or returns a **ConnectModeEnum** value. |
| [MoveRecord](#MoveRecord) | Moves a entity represented by a **Record** to another location. |
| [Open](#Open) | Opens an existing **Record** object, or creates a new item represented by the **Record** (such as a file or directory). |
| [ParentURL](#ParentURL) | Indicates an absolute URL string that points to the parent **Record** of the current **Record** object. |
| [Properties](#Properties) | Returns a reference to the **Properties** collection. |
| [RecordType](#RecordType) | Indicates the type of **Record** object. |
| [Source](#Source) | Indicates the data source or object represented by the **Record**. |
| [State](#State) | Indicates for all applicable objects whether the state of the object is open or closed. |

#### Remarks

A **Record** object represents one row of data, and has some conceptual similarities with a one-row **Recordset**. Depending upon the capabilities of your provider, **Record** objects may be returned directly from your provider instead of a one-row **Recordset**, for example when an SQL query that selects only one row is executed. Or, a **Record** object can be obtained directly from a **Recordset** object. Or, a **Record** can be returned directly from a provider to semi-structured data, such as the Microsoft Exchange OLE DB provider.

You can view the fields associated with the **Record** object by way of the **Fields** collection on the **Record** object. ADO allows object-valued columns including **Recordset**, **SafeArray**, and scalar values in the **Fields** collection of **Record** objects.

If the **Record** object represents a row in a **Recordset**, then it is possible to return to that original **Recordset** with the **Source** property.

The **Record** object can also be used by semi-structured data providers such as the Microsoft OLE DB Provider for Internet Publishing, to model tree-structured namespaces. Each node in the tree is a **Record** object with associated columns. The columns can represent the attributes of that node and other relevant information. The **Record** object can represent both a leaf node and a non-leaf node in the tree structure. Non-leaf nodes have other nodes as their contents while leaf nodes do not have such contents. Leaf nodes typically contain binary streams of data while non-leaf nodes may also have a default binary stream associated with them. **Properties** on the **Record** object identify the type of node.

The **Record** object also represents an alternative way for navigating hierarchically organized data. A **Record** object may be created to represent the root of a specific sub-tree in a large tree structure and new **Record** objects may be opened to represent child nodes.

A resource (for example, a file or directory) can be uniquely identified by an absolute URL. A **Connection** object is implicitly created and set to the **Record** object when the **Record** is opened with an absolute URL. A **Connection** object may explicitly be set to the **Record** object via the **ActiveConnection** property. The files and directories accessible via the **Connection** object define the context in which **Record** operations may occur.

Data modification and navigation methods on the **Record** object also accept a relative URL, which locates a resource using an absolute URL or the **Connection** object context as a starting point.

**Note**: URLs using the http scheme will automatically invoke the Microsoft OLE DB Provider for Internet Publishing.

A **Connection** object is associated with each **Record** object. Therefore, **Record** object operations can be part of a transaction by invoking **Connection** object transaction methods.

The **Record** object does not support ADO events, and therefore will not respond to notifications.

With the methods and properties of a **Record** object, you can do the following:

* Set or return the associated **Connection** object with the **ActiveConnection** property.
* Indicate access permissions with the **Mode** property.
* Return the URL of the directory, if any, that contains the resource represented by the **Record** with the **ParentURL** property.
* Indicate the absolute URL, relative URL, or **Recordset** from which the **Record** is derived with the **Source** property.
* Indicate the current status of the **Record** with the **State** property.
* Indicate the type of **Record**—simple, collection, or structured document—with the **RecordType** property.
* Halt execution of an asynchronous operation with the **Cancel** method.
* Disassociate the **Record** from a data source with the **Close** method.
* Copy the file or directory represented by a **Record** to another location with the **CopyRecord** method.
* Delete the file, or directory and subdirectories, represented by a **Record** with the **DeleteRecord** method.
* Open a **Recordset** containing rows that represent the subdirectories and files of the entity represented by the **Record** with the GetChildren method.
* Move (rename) the file, or directory and subdirectories, represented by a **Record** to another location with the **MoveRecord** method.
* Associate the **Record** with an existing data source, or create a new file or directory with the **Open** method.

#### Fields

A **Record** object has two special fields that can be indexed with FieldEnum constants. One constant accesses a field containing the default stream for the **Record**, and the other accesses a field containing the absolute URL string for the **Record**.

Certain providers (for example, the Microsoft OLE DB Provider for Internet Publishing) may populate the **Fields** collection with a subset of available fields for the **Record** or **Recordset**. Other fields will not be added to the collection until they are first referenced by name or indexed by your code.

Each **Field** object corresponds to a column in the **Recordset**. You use the **Value** property of **Field** objects to set or return data for the current record. Depending on the functionality the provider exposes, some collections, methods, or properties of a **Field** object may not be available.

With the collections, methods, and properties of a Field object, you can do the following:

* Return the name of a field with the **Name** property.
* View or change the data in the field with the **Value** property. **Valu**e is the default property of the **Field** object.
* Return the basic characteristics of a field with the **Type_**, **Precision**, and **NumericScale** properties.
* Return the declared size of a field with the **DefinedSize** property.
* Return the actual size of the data in a given field with the **ActualSize** property.
* Determine what types of functionality are supported for a given field with the **Attributes** property and **Properties** collection.
* Manipulate the values of fields containing long binary or long character data with the **AppendChunk** and **GetChunk** methods.
* If the provider supports batch updates, resolve discrepancies in field values during batch updating with the **OriginalValue** and **UnderlyingValue** properties.

All of the metadata properties (**Name**, **Type_**, **DefinedSize**, **Precision**, and **NumericScale**) are available before opening the **Field** object's **Recordset**. Setting them at that time is useful for dynamically constructing forms.

# <a name="ActiveConnection"></a>ActiveConnection

Sets or returns an string value that contains a definition for a connection if the connection is closed, or a Variant containing the current **Connection** object if the connection is open. Default is a null object reference.

```
PROPERTY ActiveConnection (BYREF vConn AS CVAR)
PROPERTY ActiveConnection (BYVAL pconn AS Afx_ADOConnection PTR)
PROPERTY ActiveConnection (BYREF pconn AS CAdoConnection)
PROPERTY ActiveConnection () AS Afx_ADOConnection
```

| Parameter  | Description |
| ---------- | ----------- |
| *vConn* | An string value that contains a definition for a connection if the connection is closed, or a Variant containing the current **Connection** object if the connection is open. |
| *pconn* | A reference to the **Connection** object or to the **CAdoConnection** class. |

#### Return value

A reference to the **Afx_ADOConnection** object. You must release it calling the **Release** method when no longer needed.

#### Remarks

Sets or returns a string value that contains a definition for a connection if the connection is closed, or a Variant containing the current **Connection** object if the connection is open. Default is a null object reference. See the **ConnectionString** property in the documentation for the **CADOConnection** class.

This property is read/write when the **Record** object is closed, and may contain a connection string or reference to an open **Connection** object. This property is read-only when the **Record** object is open, and contains a reference to an open **Connection** object.

A **Connection** object is created implicitly when the **Record** object is opened from a URL. Open the **Record** with an existing, open **Connection** object by assigning the Connection object to this property, or using the **Connection** object as a parameter in the **Open** method call. If the **Record** is opened from an existing **Record** or **Recordset**, then it is automatically associated with that **Record** or **Recordset** object's **Connection** object.

**Note**: URLs using the http scheme will automatically invoke the Microsoft OLE DB Provider for Internet Publishing.

# <a name="Attach"></a>Attach

Attaches a record to the class.

```
FUNCTION Attach (BYVAL pRecordset AS Afx_ADORecord PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pRecordset* | A reference to an **Afx_ADORecord** object. |
| *fAddRef* | TRUE = increase the reference count; FALSE = don't increase the reference count. |

# <a name="Cancel"></a>Cancel

Cancels execution of a pending, asynchronous method call.

```
FUNCTION Cancel () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use the **Cancel** method to terminate execution of an asynchronous method call (that is, a method invoked with the **adAsyncConnect**, **adAsyncExecute**, or **adAsyncFetch** option).

For a **Record** object, the last asynchronous call to the **CopyRecord**, **DeleteRecord**, **MoveRecord** or **Open** methods is terminated.

# <a name="Close"></a>Close

Closes a **Record** object and any dependent objects.

```
FUNCTION Close () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use the **Close** method to close a **Record** to free any associated system resources. Closing an object does not remove it from memory; you can change its property settings and open it again later. To completely eliminate an object from memory, release the connection calling the **Release** method of the interface.

# <a name="CopyRecord"></a>CopyRecord

Copies a entity represented by a **Record** to another location.

```
FUNCTION CopyRecord (BYREF Source AS CBSTR = "", BYREF Destination AS CBSTR = "", _
   BYREF UserName AS CBSTR = "", BYREF Password AS CBSTR = "", _
   BYVAL Options AS MoveRecordOptionsEnum = adCopyUnspecified, _
   BYVAL Async AS BOOLEAN = FALSE) AS CBSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *Source* | Optional. An string value that contains a URL specifying the entity to be copied (for example, a file or directory). If *Source* is omitted or specifies an empty string, the file or directory represented by the current **Record** will be copied. |
| *Destination* | Optional. An string value that contains a URL specifying the location where *Source* will be copied. |
| *UserName* | Optional. An string value that contains the user ID that, if needed, authorizes access to *Destination*. |
| *Password* | Optional. An string value that contains the password that, if needed, verifies *UserName*. |
| *Options* | Optional. A **CopyRecordOptionsEnum** value that has a default value of **adCopyUnspecified**. Specifies the behavior of this method. |
| *Async* | Optional. A Boolean value that, when True, specifies that this operation should be asynchronous. |

#### CopyRecordOptionsEnum

Specifies the behavior of the **CopyRecord** method.

| Constant   | Description |
| ---------- | ----------- |
| **adCopyAllowEmulation** | Indicates that the *Source* provider attempts to simulate the copy using download and upload operations if this method fails due to *Destination* being on a different server or is serviced by a different provider than Source. Note that differing provider capabilities may hamper performance or lose data. |
| **adCopyNonRecursive** | Copies the current directory, but none of its subdirectories, to the destination. The copy operation is not recursive. |
| **adCopyOverWrite** | Overwrites the file or directory if the *Destination* points to an existing file or directory. |
| **adCopyUnspecified** | Default. Performs the default copy operation: The operation fails if the destination file or directory already exists, and the operation copies recursively. |

#### Return value

Typically returns the value of *Destination*. However, the exact value returned is provider-dependent.

#### Remarks

The values of *Source* and *Destination* must not be identical; otherwise, a run-time error occurs. At least one of the server, path, or resource names must differ.

All children (for example, subdirectories) of *Source* are copied recursively, unless **adCopyNonRecursive** is specified. In a recursive operation, *Destination* must not be a subdirectory of Source; otherwise, the operation will not complete.

This method fails if *Destination* identifies an existing entity (for example, a file or directory), unless **adCopyOverWrite** is specified.

**Important**: Use the **adCopyOverWrite** option judiciously. For example, specifying this option when copying a file to a directory will delete the directory and replace it with the file.

**Note**: URLs using the http scheme will automatically invoke the Microsoft OLE DB Provider for Internet Publishing. 

# <a name="DeleteRecord"></a>DeleteRecord

Deletes a the entity represented by a **Record**.

```
FUNCTION DeleteRecord (BYREF cbsSource CBSTR = "", BYVAL Async BOOLEAN = FALSE) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsSource* | Optional. An string value that contains a URL identifying the entity (for example, the file or directory) to be deleted. If Source is omitted or specifies an empty string, the entity represented by the current **Record** is deleted. If the **Record** is a collection record (**RecordType** of **adCollectionRecord**, such as a directory) all children (for example, subdirectories) will also be deleted. |
| *Async* | Optional. A Boolean value that, when True, specifies that the delete operation is asychronous. |

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Operations on the object represented by this **Record** may fail after this method completes. After calling **DeleteRecord**, the **Record** should be closed because the behavior of the **Record** may become unpredictable depending upon when the provider updates the **Record** with the data source.

If this **Record** was obtained from a **Recordset**, then the results of this operation will not be reflected immediately in the **Recordset**. Refresh the **Recordset** by closing and re-opening it, or by executing the **Recordset** **Requery**, or **Update** and **Resync** methods.

**Note**: URLs using the http scheme will automatically invoke the Microsoft OLE DB Provider for Internet Publishing.

# <a name="Fields"></a>Fields

Gets a reference to the **Fields** collection of a **Record** object.

```
PROPERTY Fields () AS Afx_ADOFields PTR
```

# <a name="GetChildren"></a>GetChildren

Returns a **Recordset** whose rows represent the children of a collection **Record**.

```
FUNCTION GetChildren () AS Afx_ADORecordset
```

#### Return value

A **Recordset** object reference.

#### Remarks

The provider determines what columns exist in the returned **Recordset**. For example, a document source provider always returns a resource **Recordset**.

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

# <a name="Mode"></a>Mode

Sets or returns a **ConnectModeEnum** value. The default value for a **Record** object is **adModeRead**.

```
PROPERTY Mode () AS ConnectModeEnum
PROPERTY Mode (BYVAL Mode AS ConnectModeEnum)
```

| Parameter  | Description |
| ---------- | ----------- |
| *Mode* | A **ConnectModeEnum** value. |

#### Return value

A **ConnectModeEnum** value.

#### ConnectModeEnum

Specifies the available permissions for modifying data in a **Connection**, opening a **Record**, or specifying values for the **Mode** property of the **Record** and **Stream** objects.

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

# <a name="MoveRecord"></a>MoveRecord

Moves a entity represented by a **Record** to another location.

```
FUNCTION MoveRecord (BYREF Source AS CBSTR = "", BYREF Destination AS CBSTR = "", _
   BYREF UserName AS CBSTR = "", BYREF Password AS CBSTR = "", _
   BYVAL Options AS MoveRecordOptionsEnum = adMoveUnspecified, _
   BYVAL Async AS BOOLEAN = FALSE) AS CBSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *Source* | Optional. An string value that contains a URL identifying the **Record to be moved**. If *Source* is omitted or specifies an empty string, the object represented by this **Record** is moved. For example, if the **Record** represents a file, the contents of the file are moved to the location specified by *Destination*. |
| *Destination* | Optional. An string value that contains a URL specifying the location where *Source* will be moved. |
| *UserName* | Optional. An string value that contains the user ID that, if needed, authorizes access to *Destination*. |
| *Password* | Optional. An string value that contains the password that, if needed, verifies *UserName*. |
| *Options* | Optional. A **MoveRecordOptionsEnum** value whose default value is **adMoveUnspecified**. Specifies the behavior of this method. |
| *Async* | Optional. A Boolean value that, when True, specifies that this operation should be asynchronous. |

#### MoveRecordOptionsEnum

Specifies the behavior of the **Record** object **MoveRecord** method.

| Parameter  | Description |
| ---------- | ----------- |
| **adMoveUnspecified** | Default. Performs the default move operation: The operation fails if the destination file or directory already exists, and the operation updates hypertext links. |
| **adMoveOverWrite** | Overwrites the destination file or directory, even if it already exists. |
| **adMoveDontUpdateLinks** | Modifies the default behavior of **MoveRecord** method by not updating the hypertext links of the source Record. The default behavior depends on the capabilities of the provider. Move operation updates links if the provider is capable. If the provider cannot fix links or if this value is not specified, then the move succeeds even when links have not been fixed. |
| **adMoveAllowEmulation** | Requests that the provider attempt to simulate the move (using download, upload, and delete operations). If the attempt to move the Record fails because the destination URL is on a different server or serviced by a different provider than the source, this may cause increased latency or data loss, due to different provider capabilities when moving resources between providers. |

#### Remarks

The values of *Source* and *Destination* must not be identical; otherwise, a run-time error occurs. At least one of the server, path, or resource names must differ.

For files moved using the Internet Publishing Provider, this method updates all hypertext links in files being moved unless otherwise specified by Options. This method fails if *Destination* identifies an existing object (for example, a file or directory), unless **adMoveOverWrite** is specified.

**Note**: Use the **adMoveOverWrite** option judiciously. For example, specifying this option when moving a file to a directory will delete the directory and replace it with the file.

Certain attributes of the **Record** object, such as the **ParentURL** property, will not be updated after this operation completes. **Refresh** the **Record** object's properties by closing the **Record**, then re-opening it with the URL of the location where the file or directory was moved.

If this **Record** was obtained from a **Recordset**, the new location of the moved file or directory will not be reflected immediately in the **Recordset**. **Refresh** the **Recordset** by closing and re-opening it.

**Note**: URLs using the http scheme will automatically invoke the Microsoft OLE DB Provider for Internet Publishing.

# <a name="Open"></a>Open

Opens an existing **Record** object, or creates a new item represented by the **Record** (such as a file or directory).

```
FUNCTION Open (BYREF cvSource AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
   BYREF cvActiveConnection AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
   BYVAL nMode AS ConnectModeEnum = adModeUnknown, _
   BYVAL CreateOptions AS RecordCreateOptionsEnum = adFailIfNotExists, _
   BYREF cbsUserName AS CBSTR = "", BYREF cbsPassword AS CBSTR = "") AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cvSource* | Optional. A Variant that may represent the URL of the entity to be represented by this **Record** object, a **Command**, an open **Recordset** or another **Record** object, a string containing a SQL SELECT statement or a table name. |
| *cvActiveConnection* | Optional. A Variant that represents the connect string or open **Connection** object. |
| *nMode* | Optional. A **ConnectModeEnum** value, whose default value is **adModeUnknown**, that specifies the access mode for the resultant **Record** object. |
| *CreateOptions* | Optional. A **RecordCreateOptionsEnum** value, whose default value is **adFailIfNotExists**, that specifies whether an existing file or directory should be opened, or a new file or directory should be created. If set to the default value, the access mode is obtained from the **Mode** property. This parameter is ignored when the *cbsSource* parameter doesn't contain a URL. |
| *cbsUserName* | Optional. An string value that contains the user ID that, if needed, authorizes access to *cbsSource*. |
| *cbsPassword* | Optional. An string value that contains the password that, if needed, verifies *cbsUserName*. |

#### Return value

S_OK (0) or an HRESULT value.

#### Remarks

Source may be:

* A URL. If the protocol for the URL is http, then the Internet Provider will be invoked by default. If the URL points to a node that contains an executable script (such as an .ASP page), then a **Record** containing the source rather than the executed contents is opened by default. Use the *Options* argument to modify this behavior.

* A **Record** object. A **Record** object opened from another **Record** will clone the original **Record** object.

* A **Command** object. The opened **Record** object represents the single row returned by executing the **Command**. If the results contain more than a single row, the contents of the first row are placed in the record and an error may be added to the **Errors** collection.

* A SQL SELECT statement. The opened **Record** object represents the single row returned by executing the contents of the string. If the results contain more than a single row, the contents of the first row are placed in the record and an error may be added to the **Errors** collection.

* A table name.<br>If the **Record** object represents an entity that cannot be accessed with a URL (for example, a row of a **Recordset** derived from a database), then the values of both the **ParentURL** property and the field accessed with the **adRecordURL** constant are null.

**Note**: URLs using the http scheme will automatically invoke the Microsoft OLE DB Provider for Internet Publishing.

# <a name="ParentURL"></a>ParentURL

Indicates an absolute URL string that points to the parent **Record** of the current **Record** object.

```
PROPERTY ParentURL () AS CBSTR
```

#### Return value

The absolute URL string that points to the parent **Record** of the current **Record** object.

#### Remarks

The **ParentURL** property depends upon the source used to open the **Record** object. For example, the **Record** may be opened with a source containing a relative path name of a directory referenced by the **ActiveConnection** property.

Suppose "second" is a folder contained under "first". Open the **Record** object with the following:
```
AdoRecord.Open "second", "http://first, Mode, CreateOptions, Options, UserName, Password
```
Now, the value of the **ParentURL** property is "http://first", the same as **ActiveConnection**.

The source may also be an absolute URL such as, "http://first/second". The **ParentURL** property is then "http://first", the level above "second".

This property may be a null value if:

* There is no parent for the current object; for example, if the **Record** object represents the root of a directory.

* The **Record** object represents an entity that cannot be specified with a URL.

This property is read-only.

**Note**: This property is only supported by document source providers, such as the Microsoft OLE DB Provider for Internet Publishing.

**Note**: URLs using the http scheme will automatically invoke the Microsoft OLE DB Provider for Internet Publishing.

**Note**: If the current record contains a data record from an ADO Recordset, accessing the **ParentURL** property causes a run-time error, indicating that no URL is possible.

# <a name="Properties"></a>Properties

Returns a reference to the **Properties** collection.

```
PROPERTY Properties () AS Afx_ADOProperties
```

#### Return value

A **Properties** object reference.

# <a name="RecordType"></a>RecordType

Indicates the type of **Record** object.

```
PROPERTY RecordType () AS RecordTypeEnum
```

#### Return value

A **RecordTypeEnum** value.

#### RecordTypeEnum

Specifies the type of **Record** object.

| Constant   | Description |
| ---------- | ----------- |
| **adSimpleRecord** | Indicates a simple record (does not contain child nodes). |
| **adCollectionRecord** | Indicates a collection record (contains child nodes). |
| **adRecordUnknown** | Indicates that the type of this **Record** is unknown. |
| **adStructDoc** | Indicates a special kind of collection record that represents COM structured documents. |

# <a name="Source"></a>Source

Indicates the data source or object represented by the **Record**.

```
PROPERTY Source () AS CVAR
PROPERTY Source (BYREF cbsSource AS CBSTR)
PROPERTY PutRefSource (BYVAL pSource AS IDispatch PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsSource* | An absolute or relative URL string. |
| *pSource* | A **Recordset** or **Command** object. |

**Note**: I have used **PutRefSource**, instead of **Source**, for the overloaded property because, otherwise, the compiler fails with an ambiguous call error.

#### Return value

The absolute or relative URL string or a reference to an already open **Record**.

#### Remarks

The **Source** property returns the **Source** argument of the **Record** object **Open** method. It can contain an absolute or relative URL string. An absolute URL can be used without setting the **ActiveConnection** property to directly open the **Record** object. An implicit **Connection** object is created in this case.

The **Source** property can also contain a reference to an already open **Recordset**, which opens a **Record** object representing the current row in the **Recordset**.

# <a name="State"></a>State

Indicates for all applicable objects whether the state of the object is open or closed.

```
PROPERTY State () AS LONG
```

#### Return value

The current record state.

#### ObjectStateEnum

Specifies whether the **Open** method of a **Connection** object should return after (synchronously) or before (asynchronously) the connection is established.

| Constant   | Description |
| ---------- | ----------- |
| **adStateClosed** | Indicates that the object is closed. |
| **adStateOpen** | Indicates that the object is open. |
| **adStateConnecting** | Indicates that the object is connecting. |
| **adStateExecuting** | Indicates that the object is executing a command. |
| **adStateFetching** | Indicates that the rows of the object are being retrieved. |

#### Remarks

You can use the **State** property to determine the current state of a given object at any time.

The object's **State** property can have a combination of values. For example, if a statement is executing, this property will have a combined value of **adStateOpen** and **adStateExecuting**.
