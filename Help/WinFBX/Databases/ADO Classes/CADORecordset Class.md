# CADORecordset Class

Represents the entire set of records from a base table or the results of an executed command. At any time, the **Recordset** object refers to only a single record within the set as the current record.

**Include file**: CADORecordset.inc (include CADODB.inc)

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [AbsolutePage](#AbsolutePage) | Sets or returns a Long value from 1 to the number of pages in the **Recordset** object (**PageCount**), or returns one of the **PositionEnum** values. |
| [AbsolutePosition](#AbsolutePosition) | A value from 1 to the number of records in the **Recordset** object (**RecordCount**). |
| [ActiveCommand](#ActiveCommand) | Indicates the **Command** object that created the associated **Recordset** object. |
| [ActiveConnection](#ActiveConnection) | Sets or returns a BSTR value that contains a definition for a connection if the connection is closed, or a Variant containing the current **Connection** object if the connection is open. Default is a null object reference. |
| [AddNew](#AddNew) | Creates a new record for an updatable **Recordset** object. |
| [Attach](#Attach) | Attaches a recordset to the class. |
| [BOF](#BOF) | Indicates that the current record position is before the first record in a **Recordset** object. |
| [Bookmark](#Bookmark) | Sets or returns a Variant expression that evaluates to a valid bookmark. |
| [CacheSize](#CacheSize) | Sets or returns a Long value that must be greater than 0. Default is 1. |
| [Cancel](#Cancel) | Cancels execution of a pending, asynchronous method call. |
| [CancelBatch](#CancelBatch) | Cancels a pending batch update. |
| [CancelUpdate](#CancelUpdate) | Cancels any changes made to the current or new row of a **Recordset** object before calling the **Update** method. |
| [Clone](#Clone) | Creates a duplicate **Recordset** object from an existing **Recordset** object. Optionally, specifies that the clone be read-only. |
| [Close](#Close) | Closes a **Recordset** object and any dependent objects. |
| [Collect](#Collect) | Sets or returns a Variant value that indicates the value of the object. |
| [CompareBookmarks](#CompareBookmarks) | Compares two bookmarks and returns an indication of their relative values. |
| [CursorLocation](#CursorLocation) | Indicates the location of the cursor service. |
| [CursorType](#CursorType) | Sets or returns a **CursorTypeEnum** value. The default value is **adOpenForwardOnly**. |
| [DataMember](#DataMember) | Indicates the name of the data member that will be retrieved from the object referenced by the **DataSource** property. |
| [DataSource](#DataSource) | Indicates an object that contains data to be represented as a **Recordset** object. |
| [Delete_](#Delete_) | Deletes the current record or a group of records. |
| [EditMode](#EditMode) | Indicates the editing status of the current record. |
| [EOF](#EOF) | Indicates that the current record position is after the last record in a **Recordset** object. |
| [Fields](#Fields) | Gets a reference to the **Fields** collection of a **Record** object. |
| [Filter](#Filter) | Indicates a filter for data in a **Recordset**. |
| [Find](#Find) | Searches a **Recordset** for the row that satisfies the specified criteria. |
| [GetErrorInfo](#GetErrorInfo) | Returns information about ADO errors. |
| [GetRows](#GetRows) | Retrieves multiple records of a **Recordset** object into an array. |
| [GetString](#GetString) | Returns the **Recordset** as a string. |
| [Index](#Index) | Sets or returns a string value, which is the name of the index. |
| [LockType](#LockType) | Sets or returns the lock type, a Long value that must be greater than 0. Default is 1. |
| [MarshalOptions](#MarshalOptions) | Indicates which records are to be marshaled back to the server. |
| [MaxRecords](#MaxRecords) | Sets or returns a Long value that indicates the maximum number of records to return. Default is zero (no limit). |
| [Move](#Move) | Moves the position of the current record in a **Recordset** object. |
| [MoveFirst](#MoveFirst) | Moves to the first record in a specified **Recordset** object and makes that record the current record. |
| [MoveLast](#MoveLast) | Moves to the last record in a specified **Recordset** object and makes that record the current record. |
| [MoveNext](#MoveNext) | Moves to the next record in a specified **Recordset** object and makes that record the current record. |
| [MovePrevious](#MovePrevious) | Moves to the previous record in a specified **Recordset** object and makes that record the current record. |
| [NextRecordset](#NextRecordset) | Moves to the previous record in a specified **Recordset** object and makes that record the current record. |
| [Open](#Open) | Opens a connection to a data source. |
| [PageCount](#PageCount) | Indicates how many pages of data the **Recordset** object contains. |
| [PageSize](#PageSize) | Sets or returns a Long value that indicates how many records are on a page. The default is 10. |
| [Properties](#Properties) | Returns a reference to the **Properties** collection. |
| [RecordCount](#RecordCount) | Indicates the number of records in a **Recordset** object. |
| [Requery](#Requery) | Updates the data in a **Recordset** object by re-executing the query on which the object is based. |
| [Resync](#Resync) | Refreshes the data in the current **Recordset** object from the underlying database. |
| [Save](#Save) | Saves the **Recordset** in a file or **Stream** object. |
| [Seek](#Seek) | Searches the index of a **Recordset** to quickly locate the row that matches the specified values, and changes the current row position to that row. |
| [Sort](#Sort) | Sets or returns a string value that indicates the field names in the **Recordset** on which to sort. |
| [Source](#Source) | Indicates the data source for a **Recordset** object. |
| [State](#State) | Indicates for whether the state of the **Recordset** object is open or closed. |
| [Status](#Status) | Indicates for whether the state of the **Recordset** object is open or closed. |
| [StayInSync](#StayInSync) | Indicates, in a hierarchical **Recordset** object, whether the reference to the underlying child records (that is, the chapter) changes when the parent row position changes. |
| [Supports](#Supports) | Determines whether a specified **Recordset** object supports a particular type of functionality. |
| [Update](#Update) | Saves any changes you make to the current row of a **Recordset** object. |
| [UpdateBatch](#UpdateBatch) | Writes all pending batch updates to disk. |

#### Remarks

You use **Recordset** objects to manipulate data from a provider. When you use ADO, you manipulate data almost entirely using **Recordset** objects. All **Recordset** objects consist of records (rows) and fields (columns). Depending on the functionality supported by the provider, some **Recordset** methods or properties may not be available.

There are four different cursor types defined in ADO:

* **Dynamic cursor** -- allows you to view additions, changes, and deletions by other users; allows all types of movement through the Recordset that doesn't rely on bookmarks; and allows bookmarks if the provider supports them.

* **Keyset cursor** -- behaves like a dynamic cursor, except that it prevents you from seeing records that other users add, and prevents access to records that other users delete. Data changes by other users will still be visible. It always supports bookmarks and therefore allows all types of movement through the **Recordset**.

* **Static cursor** -- provides a static copy of a set of records for you to use to find data or generate reports; always allows bookmarks and therefore allows all types of movement through the Recordset. Additions, changes, or deletions by other users will not be visible. This is the only type of cursor allowed when you open a client-side **Recordset** object.

* **Forward-only cursor** -- allows you to only scroll forward through the **Recordset++. Additions, changes, or deletions by other users will not be visible. This improves performance in situations where you need to make only a single pass through a **Recordset**.

Set the **CursorType** property prior to opening the **Recordset** to choose the cursor type, or pass a **CursorType** argument with the **Open** method. Some providers don't support all cursor types. Check the documentation for the provider. If you don't specify a cursor type, ADO opens a forward-only cursor by default.

If the **CursorLocation** property is set to **adUseClient** to open a **Recordset**, the **UnderlyingValue** property on **Field** objects is not available in the returned Recordset object. When used with some providers (such as the Microsoft ODBC Provider for OLE DB in conjunction with Microsoft SQL Server), you can create **Recordset** objects independently of a previously defined **Connection** object by passing a connection string with the **Open** method. ADO still creates a **Connection** object, but it doesn't assign that object to an object variable. However, if you are opening multiple Recordset objects over the same connection, you should explicitly create and open a **Connection** object; this assigns the **Connection** object to an object variable. If you do not use this object variable when opening your **Recordset** objects, ADO creates a new **Connection** object for each new **Recordset**, even if you pass the same connection string.

You can create as many **Recordset** objects as needed.

When you open a **Recordset**, the current record is positioned to the first record (if any) and the **BOF** and **EOF** properties are set to False. If there are no records, the **BOF** and **EOF** property settings are True.

You can use the **MoveFirst**, **MoveLast**, **MoveNext**, and **MovePrevious** methods; the **Move** method; and the **AbsolutePosition**, **AbsolutePage**, and **Filter** properties to reposition the current record, assuming the provider supports the relevant functionality. Forward-only **Recordset** objects support only the **MoveNext** method. When you use the **Move** methods to visit each record (or enumerate the **Recordset**), you can use the **BOF** and **EOF** properties to determine if you've moved beyond the beginning or end of the **Recordset**.

Before using any functionality of a **Recordset** object, you must call the **Supports** method on the object to verify that the functionality is supported or available. You must not use the functionality when the **Supports** method returns false. For example, you can use the **MovePrevious** method only if **Recordset.Supports**(*adMovePrevious*) returns true. Otherwise, you will get an error, because the **Recordset** object might have been closed and the functionality rendered unavailable on the instance. If a feature you are interested in is not supported, **Supports** will return false as well. In this case, you should avoid calling the corresponding property or method on the **Recordset** object.

Recordset objects can support two types of updating: immediate and batched. In immediate updating, all changes to data are written immediately to the underlying data source once you call the **Update** method. You can also pass arrays of values as parameters with the **AddNew** and **Update** methods and simultaneously update several fields in a record.

If a provider supports batch updating, you can have the provider cache changes to more than one record and then transmit them in a single call to the database with the **UpdateBatch** method. This applies to changes made with the **AddNew**, **Update**, and **Delete_** methods. After you call the **UpdateBatch** method, you can use the **Status** property to check for any data conflicts in order to resolve them.

**Note**: To execute a query without using a **Command** object, pass a query string to the **Open** method of a **Recordset** object. However, a **Command** object is required when you want to persist the command text and re-execute it, or use query parameters.

The **Mode** property governs access permissions.

A **Recordset** object has a **Fields** collection made up of **Field** objects. Each **Field** object corresponds to a column in the **Recordset**. You can populate the **Fields** collection before opening the **Recordset** by calling the **Refresh** method on the collection.

The **Fields** collection has an **Append** method, which provisionally creates and adds a **Field** object to the collection, and an Update method, which finalizes any additions or deletions.

Certain providers (for example, the Microsoft OLE DB Provider for Internet Publishing) may populate the **Fields** collection with a subset of available fields for the **Record** or **Recordset**. Other fields will not be added to the collection until they are first referenced by name or indexed by your code.

If you attempt to reference a nonexistent field by name, a new **Field** object will be appended to the **Fields** collection with a **Status** of **adFieldPendingInsert**. When you call **Update**, ADO will create a new field in your data source if allowed by your provider.

Each **Field** object corresponds to a column in the **Recordset**. You use the **Value** property of **Field** objects to set or return data for the current record. Depending on the functionality the provider exposes, some collections, methods, or properties of a **Field** object may not be available.

With the collections, methods, and properties of a **Field** object, you can do the following:

* Return the name of a field with the **Name** property.
* View or change the data in the field with the **Value** property. **Value** is the default property of the **Field** object.
* Return the basic characteristics of a field with the **Type_**, **Precision**, and **NumericScale** properties.
* Return the declared size of a field with the **DefinedSize** property.
* Return the actual size of the data in a given field with the **ActualSize** property.
* Determine what types of functionality are supported for a given field with the **Attributes** property and **Properties** collection.
* Manipulate the values of fields containing long binary or long character data with the **AppendChunk** and **GetChunk** methods.
If the provider supports batch updates, resolve discrepancies in field values during batch updating with the **OriginalValue** and **UnderlyingValue** properties.

All of the metadata properties (**Name**, **Type_**, **DefinedSize**, **Precision**, and **NumericScale**) are available before opening the **Field** object's **Recordset**. Setting them at that time is useful for dynamically constructing forms.

When a **Recordset** object is passed across processes, only the rowset values are marshalled, and the properties of the **Recordset** object are ignored. During unmarshalling, the rowset is unpacked into a newly created **Recordset** object, which also sets its properties to the default values.

### Recordset Dynamic Properties

The table below is a cross-index of the ADO and OLE DB names for each standard OLE DB provider dynamic property. Your providers may add more properties than listed here. For the specific information about provider-specific dynamic properties, see your provider documentation.

| ADO Property Name | OLE DB Property Name |
| ----------------- | ----------- |
| IAccessor | DBPROP_IACCESSOR (1) 
| IChapteredRowset (1) | |
| ColumnsInfo | DBPROP_ICOLUMNSINFO (1) |
| IColumnsRowset | DBPROP_ICOLUMNSROWSET (1) |
| IConnectionPointContainer | DBPROP_ICONNECTIONPOINTCONTAINER (1) | 
| IConvertType (1) | |
| ILockBytes | DBPROP_ILOCKBYTES (1) |
| IRowset | DBPROP_IROWSET (1) |
| IDBAsynchStatus | DBPROP_IDBASYNCHSTATUS (1) |
| IParentRowset (1) | |
| IRowsetChange | DBPROP_IROWSETCHANGE (1) |
| IRowsetExactScroll (1) | |
| IRowsetFind | DBPROP_IROWSETFIND (1) |
| IRowsetIdentity | DBPROP_IROWSETIDENTITY (1) |
| IRowsetInfo | DBPROP_IROWSETINFO (1) |
| IRowsetLocate | DBPROP_IROWSETLOCATE (1) |
| IRowsetRefresh | DBPROP_IROWSETREFRESH (1) |
| IRowsetResynch (1) | |
| IRowsetScroll | DBPROP_IROWSETSCROLL (1) |
| IRowsetUpdate | DBPROP_IROWSETUPDATE (1) |
| IRowsetView | DBPROP_IROWSETVIEW (1) |
| IRowsetIndex | DBPROP_IROWSETINDEX (1) |
| ISequentialStream | DBPROP_ISEQUENTIALSTREAM (1) |
| IStorage | DBPROP_ISTORAGE (1) |
| IStream | DBPROP_ISTREAM (1) |
| ISupportErrorInfo | DBPROP_ISUPPORTERRORINFO (1) |
| Access Order | DBPROP_ACCESSORDER |
| Append-Only Rowset | DBPROP_APPENDONLY |
| Asynchronous Rowset Processing | DBPROP_ROWSET_ASYNCH |
| Auto Recalc | DBPROP_ADC_AUTORECALC |
| Background Fetch Size | DBPROP_ASYNCHFETCHSIZE |
| Background Thread Priority | DBPROP_ASYNCHTHREADPRIORITY |
| Batch Size | DBPROP_ADC_BATCHSIZE |
| Blocking Storage Objects | DBPROP_BLOCKINGSTORAGEOBJECTS |
| Bookmark Type | DBPROP_BOOKMARKTYPE |
| Bookmarkable | DBPROP_IROWSETLOCATE (2) |
| Bookmarks Ordered | DBPROP_ORDEREDBOOKMARKS |
| Cache Child Rows | DBPROP_ADC_CACHECHILDROWS |
| Cache Deferred Columns | DBPROP_CACHEDEFERRED |
| Change Inserted Rows | DBPROP_CHANGEINSERTEDROWS |
| Column Privileges | DBPROP_COLUMNRESTRICT |
| Column Set Notification | DBPROP_NOTIFYCOLUMNSET |
| Column Writable | DBPROP_MAYWRITECOLUMN |
| Command Time Out | DBPROP_COMMANDTIMEOUT |
| Cursor Engine Version | DBPROP_ADC_CEVER |
| Defer Column | DBPROP_DEFERRED |
| Delay Storage Object Updates | DBPROP_DELAYSTORAGEOBJECTS |
| Fetch Backwards | DBPROP_CANFETCHBACKWARDS |
| Filter Operations | DBPROP_FILTERCOMPAREOPS |
| Find Operations | DBPROP_FINDCOMPAREOPS |
| Hidden Columns (Count) | DBPROP_HIDDENCOLUMNS (3) |
| Hold Rows | DBPROP_CANHOLDROWS |
| Immobile Rows | DBPROP_IMMOBILEROWS |
| Initial Fetch Size | DBPROP_ASYNCHPREFETCHSIZE |
| Literal Bookmarks | DBPROP_LITERALBOOKMARKS |
| Literal Row Identity | DBPROP_LITERALIDENTITY |
| Maintain Change Status | DBPROP_ADC_MAINTAINCHANGESTATUS |
| Maximum Open Rows | DBPROP_MAXOPENROWS |
| Maximum Pending Rows | DBPROP_MAXPENDINGROWS |
| Maximum Rows | DBPROP_MAXROWS (4) |
| Memory Usage | DBPROP_MEMORYUSAGE |
| Notification Granularity | DBPROP_NOTIFICATIONGRANULARITY |
| Notification Phases | DBPROP_NOTIFICATIONPHASES |
| Objects Transacted | DBPROP_TRANSACTEDOBJECT |
| Others' Changes Visible | DBPROP_OTHERUPDATEDELETE |
| Others' Inserts Visible | DBPROP_OTHERINSERT |
| Own Changes Visible | DBPROP_OWNUPDATEDELETE |
| Own Inserts Visible | DBPROP_OWNINSERT |
| Preserve on Abort | DBPROP_ABORTPRESERVE |
| Preserve on Commit | DBPROP_COMMITPRESERVE |
| Private1 (5) | |
| Quick Restart | DBPROP_QUICKRESTART |
| Reentrant Events | DBPROP_REENTRANTEVENTS |
| Remove Deleted Rows | DBPROP_REMOVEDELETED |
| Report Multiple Changes | DBPROP_REPORTMULTIPLECHANGES |
| Reshape Name | DBPROP_ADC_RESHAPENAME |
| Resync Command | DBPROP_ADC_CUSTOMRESYNCH |
| Return Pending Inserts | DBPROP_RETURNPENDINGINSERTS |
| Row Delete Notification | DBPROP_NOTIFYROWDELETE |
| Row First Change Notification | DBPROP_NOTIFYROWFIRSTCHANGE |
| Row Insert Notification | DBPROP_NOTIFYROWINSERT |
| Row Privileges | DBPROP_ROWRESTRICT |
| Row Resynchronization Notification | DBPROP_NOTIFYROWRESYNCH |
| Row Threading Model | DBPROP_ROWTHREADMODEL |
| Row Undo Change Notification | DBPROP_NOTIFYROWUNDOCHANGE |
| Row Undo Delete Notification | DBPROP_NOTIFYROWUNDODELETE |
| Row Undo Insert Notification | DBPROP_NOTIFYROWUNDOINSERT |
| Row Update Notification | DBPROP_NOTIFYROWUPDATE |
| Rowset Fetch Position Change Notification | DBPROP_NOTIFYROWSETFETCHPOSITIONCHANGE |
| Rowset Release Notification | DBPROP_NOTIFYROWSETRELEASE |
| Scroll Backwards | DBPROP_CANSCROLLBACKWARDS |
| Server Cursor | DBPROP_SERVERCURSOR |
| Skip Deleted Bookmarks | DBPROP_BOOKMARKSKIPPED |
| Strong Row Identity | DBPROP_STRONGIDENTITY |
| Unique Catalog | DBPROP_ADC_UNIQUECATALOG |
| Unique Rows | DBPROP_UNIQUEROWS |
| Unique Schema | DBPROP_ADC_UNIQUESCHEMA |
| Unique Table | DBPROP_ADC_UNIQUETABLE |
| Updatability | DBPROP_UPDATABILITY |
| Update Criteria | DBPROP_ADC_UPDATECRITERIA |
| Update Resync | DBPROP_ADC_UPDATERESYNC |
| Use Bookmarks | DBPROP_BOOKMARKS |

Note numbers used in the cross-index:

(1) This property is a Boolean flag indicating whether the named interface should be used. The equivalent OLE DB property name is listed if it exists.

(2) The "Bookmarkable" ADO property is generated internally for backwards compatibility, and is mapped to the OLE DB property, DBPROP_IROWSETLOCATE. This is the same property that corresponds to the ADO property, ^^IRowsetLocate**.

(3) The ADO property name, "Hidden Columns", is named differently than the OLE DB property name Description, "Hidden Columns Count".

(4) For hierarchical recordsets, the "Maximum Rows" ADO property gets applied across all children. Depending on the order in which the rows are returned, you might have all, some or no children for each parent or orphaned children in the result set. Therefore, when reshaping hierarchical recordsets, the identifier for every child should be unique. In general, the Microsoft Data Shaping Service for OLE DB (MSDATASHAPE) provider does not allow for distinction between properties that can be inherited from the parent and those that cannot be inherited.

(5) Does not apply.

### In-memory recordsets

As an alternative to arrays and/or dictionary objects, we can use ADO to create in-memory recordsets. Some advantages of using ADO over SQLite for this task is that we don't need to use a third party DLL and that we can save a recordset to disk as a stream or as XML just calling **Save** or **SaveAsXml**.

**Seek** does not work (Seek only works with server-side cursors), but we can use **Find** (to search by name) or **AbsolutePosition** (to search by ordinal). We can also use other ADO methods such **Delete_**, **Sort** and **Filter**.

We can also get all the rows as a two-dimensional safe array, calling the **GetRows** method, or as an string calling the **GetString** method, that allows to specify the number of rows to read, a separator (default = tab) and a row delimiter (default = CRLF).

This is a good option for things that we can't do with normal arrays, such having a multi-dimensional array in which each dimension can be of any type.

To get a localized description of the last ADO error, call the wrapper function **AfxGetOleErrorInfo**. The wrapper function **AfxAdoGetErrorInfo** or the **GetErrorInfo** method won't work because they need an active connection to a database.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
USING Afx

' // Create an instance of the CAdoRecorset class
DIM pRecordset AS CAdoRecordset
' // Get a reference to the Fields collection
DIM pFields AS CAdoFields = pRecordset.Fields

pFields.Append("Key", adVarChar, 10)
pFields.Append("Item", adVarChar, 20)

'pRecordset.CursorType = adOpenKeyset
'pRecordset.LockType = adLockOptimistic
pRecordset.CursorLocation = adUseClient
pRecordset.Open(adOpenKeyset, adLockOptimistic)

print "Record count: ", pRecordset.Recordcount

pRecordset.AddNew
   pRecordset.Collect("Key") = "One"
   pRecordset.Collect("Item") = "Item one"
'pRecordset.Update

pRecordset.AddNew
   pRecordset.Collect("Key") = "Two"
   pRecordset.Collect("Item") = "Item two"
'pRecordset.Update

pRecordset.AddNew
   pRecordset.Collect("Key") = "Three"
   pRecordset.Collect("Item") = "Item three"
'pRecordset.Update

print "New record count: ", pRecordset.Recordcount

'pRecordset.MoveFirst
pRecordset.AbsolutePosition = 3
DO
   IF pRecordset.EOF THEN EXIT DO
   PRINT pRecordset.Collect("Key")
   PRINT pRecordset.Collect("Item")
   IF pRecordset.MoveNext <> S_OK THEN EXIT DO
LOOP

pRecordset.MoveFirst
PRINT pRecordset.GetString
```

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Create an instance of the CAdoRecorset class
DIM pRecordset AS CAdoRecordset
' // Get a reference to the Fields collection
DIM pFields AS CAdoFields = pRecordset.Fields

pFields.Append("Key", adVarChar, 10)
pFields.Append("Item", adVarChar, 20)

pRecordset.CursorLocation = adUseClient
pRecordset.Open(adOpenKeyset, adLockOptimistic)

pRecordset.AddNew
   pRecordset.Collect("Key") = "One"
   pRecordset.Collect("Item") = "Item one"

pRecordset.AddNew
   pRecordset.Collect("Key") = "Two"
   pRecordset.Collect("Item") = "Item two"

pRecordset.AddNew
   pRecordset.Collect("Key") = "Three"
   pRecordset.Collect("Item") = "Item three"

' // Start searching from the first record (adBookmarkFirst)
' // You can also start from the current record (the default value, adBookmarkCurrent)
' // or from the last record (adBookmarkLast). The default search direction is
' // adSearchForward. If you want to perform a backwards search, you will have to
' // specify adSearchBackward in the search direction parameter.
IF pRecordset.Find("Key = 'Two'", adBookmarkFirst) = S_OK THEN
   PRINT pRecordset.Collect("Item").ToStr
ELSE
   PRINT "Record not found"
END IF
```

# <a name="AbsolutePage"></a>AbsolutePage

Sets or returns a Long value from 1 to the number of pages in the **Recordset** object (**PageCount**), or returns one of the **PositionEnum** values.

```
PROPERTY AbsolutePage () AS PositionEnum
PROPERTY AbsolutePage (BYVAL Page AS PositionEnum)
```

#### PositionEnum

Specifies the current position of the record pointer within a Recordset.

| Constant   | Description |
| ---------- | ----------- |
| **adPosBOF** | Indicates that the current record pointer is at **BOF** (that is, the **BOF** property is True). |
| **adPosEOF** | Indicates that the current record pointer is at **EOF** (that is, the **EOF** property is True). |
| **adPosUnknown** | Indicates that the **Recordset** is empty, the current position is unknown, or the provider does not support the **AbsolutePage** or **AbsolutePosition** property. |

#### Return value

The page number.

#### Remarks

This property can be used to identify the page number on which the current record is located. It uses the **PageSize** property to logically divide the total rowset count of the **Recordset** object into a series of pages, each of which has the number of records equal to **PageSize** (except for the last page, which may have fewer records). The provider must support the appropriate functionality for this property to be available.

When getting or setting the **AbsolutePage** property, ADO uses the **AbsolutePosition** property and the **PageSize** property together as follows:

* To get the **AbsolutePage**, ADO first retrieves the **AbsolutePosition**, and then divides it by the **PageSize**.
* To set the **AbsolutePage**, ADO moves the **AbsolutePosition** as follows: it multiplies the **PageSize** by the new **AbsolutePage** value and then adds 1 to the value. As a result, the current position in the **Recordset** after successfully setting **AbsolutePage** is, the first record in that page.

Like the **AbsolutePosition** property, **AbsolutePage** is 1-based and equals 1 when the current record is the first record in the **Recordset**. Set this property to move to the first record of a particular page. Obtain the total number of pages from the **PageCount** property.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Set the cursor location
DIM pRecordset AS CAdoRecordset
pRecordset.CursorLocation = adUseClient

' // Open the recordset
DIM cvSource AS CVAR = "SELECT * FROM Publishers"
pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

' // Display five records at a time
DIM nPageSize AS LONG = 5
pRecordset.PageSize = nPageSize
' // Retrieve the number of pages
DIM nPageCount AS LONG = pRecordset.PageCount

' // Parse the recordset
FOR i AS LONG = 1 TO nPageCount
   ' // Set the cursor at the beginning of the page
   pRecordset.AbsolutePage = i
   FOR x AS LONG = 1 TO nPageSize
      ' // Get the content of the "Name" column
      DIM cvRes AS CVAR = pRecordset.Collect("Name")
      PRINT cvRes
      ' // Fetch the next row
      pRecordset.MoveNext
      IF pRecordset.EOF THEN EXIT FOR
   NEXT
   PRINT
NEXT
```

# <a name="AbsolutePosition"></a>AbsolutePosition

Sets or returns a Long value from 1 to the number of records in the **Recordset** object (**RecordCount**), or returns one of the **PositionEnum** values.

```
PROPERTY AbsolutePosition () AS PositionEnum
PROPERTY AbsolutePosition (BYVAL Position AS PositionEnum)
```

#### PositionEnum

Specifies the current position of the record pointer within a Recordset.

| Constant   | Description |
| ---------- | ----------- |
| **adPosBOF** | Indicates that the current record pointer is at **BOF** (that is, the **BOF** property is True). |
| **adPosEOF** | Indicates that the current record pointer is at **EOF** (that is, the **EOF** property is True). |
| **adPosUnknown** | Indicates that the **Recordset** is empty, the current position is unknown, or the provider does not support the **AbsolutePage** or **AbsolutePosition** property. |

#### Return value

The absolute position.

#### Remarks

In order to set the **AbsolutePosition** property, ADO requires that the OLE DB provider you are using implement the **IRowsetLocate** interface.

Accessing the **AbsolutePosition** property of a **Recordset** that was opened with either a forward-only or dynamic cursor raises the error **adErrFeatureNotAvailable**. With other cursor types, the correct position will be returned as long as the provider supports the **IRowsetScroll** interface. If the provider does not support the **IRowsetScroll** interface, the property is set to **adPosUnknown**. See the documentation for your provider to determine whether it supports **IRowsetScroll**.

Use the **AbsolutePosition** property to move to a record based on its ordinal position in the **Recordset** object, or to determine the ordinal position of the current record. The provider must support the appropriate functionality for this property to be available.

Like the **AbsolutePage** property, **AbsolutePosition** is 1-based and equals 1 when the current record is the first record in the **Recordset**. You can obtain the total number of records in the **Recordset** object from the **RecordCount** property.

When you set the **AbsolutePosition** property, even if it is to a record in the current cache, ADO reloads the cache with a new group of records starting with the record you specified. The **CacheSize** property determines the size of this group.

**Note**: You should not use the **AbsolutePosition** property as a surrogate record number. The position of a given record changes when you delete a preceding record. There is also no assurance that a given record will have the same **AbsolutePosition** if the **Recordset** object is requeried or reopened. Bookmarks are still the recommended way of retaining and returning to a given position and are the only way of positioning across all types of **Recordset** objects.

#### Example

This example demonstrates how the **AbsolutePosition** property can track the progress of a loop that enumerates all the records of a **Recordset**. It uses the **CursorLocation** property to enable the **AbsolutePosition** property by setting the cursor to a client cursor.

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Set the cursor location
DIM pRecordset AS CAdoRecordset
pRecordset.CursorLocation = adUseClient

' // Open the recordset
DIM cvSource AS CVAR = "Publishers"
pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdTable)

' // Parse the recordset
DO
   ' // While not at the end of the recordset...
   IF pRecordset.EOF THEN EXIT DO
   ' // Get the contents of the "City" and "Name" columns
   DIM cvRes AS CVAR = pRecordset.Collect("Name")
   PRINT "Position: "; pRecordset.AbsolutePosition; " "; cvRes

   ' // Fetch the next row
   IF pRecordset.MoveNext <> S_OK THEN EXIT DO
LOOP
```

# <a name="ActiveCommand"></a>ActiveCommand

Indicates the **Command** object that created the associated **Recordset** object.

```
PROPERTY ActiveCommand () AS Afx_ADOCommand PTR
```

#### Return value

A **Command** object reference.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' ========================================================================================
' The ShowActiveCommand routine is given only a Recordset object, yet it must print the
' command text and parameter that created the Recordset. This can be done because the
' Recordset object's ActiveCommand property yields the associated Command object.
' The Command object's CommandText property yields the parameterized command that was
' substituted for the command's parameter placeholder ("?").
' ========================================================================================
SUB ShowActiveCommand (BYREF pConnection AS CAdoConnection, BYREF pRecordset AS CAdoRecordset)

   DIM pCommand AS CAdoCommand = pRecordset.ActiveCommand
   DIM cbsCommandText AS CBSTR = pCommand.CommandText
   DIM pParameters AS CAdoParameters = pCommand.Parameters
   DIM pParameter AS CAdoParameter = pParameters.Item("Name")
   DIM cvValue AS CVAR = pParameter.Value
   PRINT "Command text: "; cbsCommandText
   PRINT "Parameter: "; cvValue

   IF pRecordset.BOF THEN
      PRINT "Name = '"; cvValue; "'not found"
   ELSE
      DIM cvAuthor AS CVAR = pRecordset.Collect("Author")
      DIM cvID AS CVAR = pRecordset.Collect("Author")
      PRINT "Name= "; cvAuthor; ","; cvID
   END IF

END SUB
' ========================================================================================

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' Set the ADOCommand active connection
DIM pCommand AS CAdoCommand
pCommand.ActiveConnection = pConnection

' //Set the command text
pCommand.CommandText = "SELECT * FROM Authors WHERE Author = ?"
' // Create the parameter
DIM pParameter AS CADOParameter = pCommand.CreateParameter("Name", adChar, adParamInput, 255, "Bard, Dick")
' // Add the parameter to the collection
DIM pParameters AS CAdoParameters = pCommand.Parameters
pParameters.Append(pParameter)

' // Create the recordset by executing the command string
DIM pRecordset AS CAdoRecordset = pCommand.Execute
' // Display the results
ShowActiveCommand pConnection, pRecordset
```

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

Use the **ActiveConnection** property to determine the **Connection** object over which the specified **Recordset** will be opened.

For open **Recordset** objects or for **Recordset** objects whose **Source** property is set to a valid **Command** object, the **ActiveConnection** property is read-only. Otherwise, it is read/write.

You can set this property to a valid **Connection** object or to a valid connection string. In this case, the provider creates a new **Connection** object using this definition and opens the connection. Additionally, the provider may set this property to the new **Connection** object to give you a way to access the **Connection** object for extended error information or to execute other commands.

If you use the **ActiveConnection** argument of the **Open** method to open a **Recordset** object, the **ActiveConnection** property will inherit the value of the argument.

If you set the **Source** property of the **Recordset** object to a valid **Command** object variable, the **ActiveConnection** property of the **Recordset** inherits the setting of the **Command** object's **ActiveConnection** property.

#### Remote Data Service Usage

When used on a client-side **Recordset** object, this property can be set only to a connection string or to NULL.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Create a Recordset object
DIM pRecordset AS CAdoRecordset
' // Set the active connection
pRecordset.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"
' // Open the recordset
DIM cvSource AS CVAR = "SELECT TOP 20 * FROM Authors ORDER BY Author"
DIM hr AS HRESULT = pRecordset.Open(cvSource)
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
DIM ccConStr AS CVAR = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"
pConnection.Open cvConStr
' // Set the active connection
pRecordset.ActiveConnection = pConnection
' // Open the recordset
DIM cvSource AS CVAR = "SELECT TOP 20 * FROM Authors ORDER BY Author"
DIM hr AS HRESULT = pRecordset.Open(cvSource)
```

#### Disconnected recordset

A disconnected recordset is one of the powerful features of ADO wherein the connection is removed from a populated recordset. This recordset can be manipulated and again connected to the database for updating. Remote Data Services (RDS) uses this feature to send recordsets through either HTTP or Distributed Component Object Model (DCOM) protocols to a remote computer, however, you are not limited to using Remote Data Service (RDS) to generate a disconnected recordset.

We can manipulate ADO directly to disconnect a recordset without using either RDS Server or Client side components.

This technique is demonstrated below and is accomplished by setting the **ActiveConnection** property.

One of the primary requisites for a recordset to become a disconnected recordset is that it should use client side cursors. That is, the **CursorLocation** should be initialized to **adUseClient**.

Disconnecting a recordset can be done by setting the **ActiveConnection** property to NULL.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection PTR = NEW CAdoConnection
pConnection->Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Open the recordset
DIM pRecordset AS CAdoRecordset
' // Setting the cursor location to client side is important to get a disconnected recordset
pRecordset.CursorLocation = adUseClient
' // Open the recordset
DIM cvSource AS CVAR = "SELECT * FROM Authors"
pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

' // Disconnect the recordset by setting its active connection to null.
' // Casting to Afx_ADOConnection PTR is needed to get the correct overloaded method called;
' // otherwise, the CVAR version will be called and it will fail.
pRecordset.ActiveConnection = cast(Afx_ADOConnection PTR, NULL)

' // Close and release the connection
Delete pConnection

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

# <a name="AddNew"></a>AddNew

Creates a new record for an updatable **Recordset** object.

```
FUNCTION AddNew ( _
   BYREF cvFieldList AS CVAR = TYPE(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
   BYREF cvValues AS CVAR = TYPE(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND) _
) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cvFieldList* | Optional. A single name, or an array of names or ordinal positions of the fields in the new record. |
| *cvValues* | Optional. A single value, or an array of values for the fields in the new record. If *vFieldlist* is an array, *vValues* must also be an array with the same number of members; otherwise, an error occurs. The order of field names must match the order of field values in each array. |

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use the **AddNew** method to create and initialize a new record. Use the **Supports** method with **adAddNew** (a **CursorOptionEnum** value) to verify whether you can add records to the current **Recordset** object.

After you call the **AddNew** method, the new record becomes the current record and remains current after you call the **Update** method. Since the new record is appended to the **Recordset**, a call to **MoveNext** following the **Update** will move past the end of the **Recordset**, making **EOF** True. If the **Recordset** object does not support bookmarks, you may not be able to access the new record once you move to another record. Depending on your cursor type, you may need to call the **Requery** method to make the new record accessible.

If you call **AddNew** while editing the current record or while adding a new record, ADO calls the **Update** method to save any changes and then creates the new record.

The behavior of the **AddNew** method depends on the updating mode of the **Recordset** object and whether you pass the *vFieldlist* and *vValues* arguments.

In immediate update mode (in which the provider writes changes to the underlying data source once you call the **Update** method), calling the **AddNew** method without arguments sets the **EditMode** property to **adEditAdd** (an **EditModeEnum** value). The provider caches any field value changes locally. Calling the **Update** method posts the new record to the database and resets the **EditMode** property to **adEditNone** (an **EditModeEnum** value). If you pass the *vFieldlist* and *vValues* arguments, ADO immediately posts the new record to the database (no **Update** call is necessary); the **EditMode** property value does not change (**adEditNone**).

In batch update mode (in which the provider caches multiple changes and writes them to the underlying data source only when you call the **UpdateBatch** method), calling the **AddNew** method without arguments sets the **EditMode** property to **adEditAdd**. The provider caches any field value changes locally. Calling the **Update** method adds the new record to the current **Recordset**, but the provider does not post the changes to the underlying database, or reset the **EditMode** to **adEditNone**, until you call the **UpdateBatch** method. If you pass the *vFieldlist* and *vValues* arguments, ADO sends the new record to the provider for storage in a cache and sets the **EditMode** to **adEditAdd**; you need to call the **UpdateBatch** method to post the new record to the underlying database.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Create a Connection object
DIM pConnection AS CAdoConnection
' // Create a Recordset object
DIM pRecordset AS CAdoRecordset

' // Open the connection
DIM cvConStr AS CVAR = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"
pConnection.Open cvConStr

' // Open the recordset
DIM cvSource AS CVAR = "Publishers"
pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdTableDirect)

' // Add a new record
pRecordset.AddNew
   pRecordset.Collect("PubID") = CLNG(10000)
   pRecordset.Collect("Name") = "Wile E. Coyote"
   pRecordset.Collect("Company Name") = "Warner Brothers Studios"
   pRecordset.Collect("Address") = "4000 Warner Boulevard"
   pRecordset.Collect("City") = "Burbank, CA. 91522"
pRecordset.Update
```

# <a name="Attach"></a>Attach

Attaches a recordset to the class.

```
FUNCTION Attach (BYVAL pRecordset AS AFX_ADORecordset PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pRecordset* | A reference to an **Afx_ADORecordset** object. |
| *fAddRef* | TRUE = increase the reference count; FALSE = don't increase the reference count. |

#### Return value

S_OK or an HRESULT code.

# <a name="BOF"></a>BOF

Indicates that the current record position is before the first record in a **Recordset** object.

```
PROPERTY BOF () AS BOOLEAN
```

#### Return value

TRUE if the current record position is before the first record; FALSE, otherwise.

# <a name="EOF"></a>EOF

Indicates that the current record position is after the last record in a **Recordset** object.

```
PROPERTY EOF () AS BOOLEAN
```

#### Return value

TRUE if the current record position is after the last record; FALSE, otherwise.

# <a name="Bookmark"></a>Bookmark

Sets or returns a Variant expression that evaluates to a valid bookmark.

```
PROPERTY Bookmark () AS CVAR
PROPERTY Bookmark (BYREF cvBookmark AS CVAR)
```

#### Return value

The bookmark.

#### Remarks

Use the **Bookmark** property to save the position of the current record and return to that record at any time. Bookmarks are available only in **Recordset** objects that support bookmark functionality.

When you open a **Recordset** object, each of its records has a unique bookmark. To save the bookmark for the current record, assign the value of the **Bookmark** property to a variable. To quickly return to that record at any time after moving to a different record, set the **Recordset** object's **Bookmark** property to the value of that variable.

The user may not be able to view the value of the bookmark. Also, users should not expect bookmarks to be directly comparable—two bookmarks that refer to the same record may have different values.

If you use the **Clone** method to create a copy of a **Recordset** object, the **Bookmark** property settings for the original and the duplicate **Recordset** objects are identical and you can use them interchangeably. However, you cannot use bookmarks from different **Recordset** objects interchangeably, even if they were created from the same source or command.

#### Remote Data Service Usage

When used on a client-side Recordset object, the **Bookmark** property is always available.


# <a name="CacheSize"></a>CacheSize

Sets or returns a Long value that must be greater than 0. Default is 1.

```
PROPERTY CacheSize () AS LONG
PROPERTY CacheSize (BYVAL size AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *size* | A value that must be greater than 0. |

#### Return value

The cache size.

#### Remarks

Use the **CacheSize** property to control how many records to retrieve at one time into local memory from the provider. For example, if the **CacheSize** is 10, after first opening the **Recordset** object, the provider retrieves the first 10 records into local memory. As you move through the **Recordset** object, the provider returns the data from the local memory buffer. As soon as you move past the last record in the cache, the provider retrieves the next 10 records from the data source into the cache.

**Note**: **CacheSize** is based on the Maximum Open Rows provider-specific property (in the **Properties** collection of the **Recordset** object). You cannot set **CacheSize** to a value greater than Maximum Open Rows. To modify the number of rows which can be opened by the provider, set Maximum Open Rows.

The value of **CacheSize** can be adjusted during the life of the **Recordset** object, but changing this value only affects the number of records in the cache after subsequent retrievals from the data source. Changing the property value alone will not change the current contents of the cache.

If there are fewer records to retrieve than **CacheSize** specifies, the provider returns the remaining records and no error occurs.

A **CacheSize** setting of zero is not allowed and returns an error.

Records retrieved from the cache don't reflect concurrent changes that other users made to the source data. To force an update of all the cached data, use the **Resync** method.

If **CacheSize** is set to a value greater than one, the navigation methods (**Move**, **MoveFirst**, **MoveLast**, **MoveNext**, and **MovePrevious**) may result in navigation to a deleted record, if deletion occurs after the records were retrieved. After the initial fetch, subsequent deletions will not be reflected in your data cache until you attempt to access a data value from a deleted row. However, setting **CacheSize** to one eliminates this issue since deleted rows cannot be fetched.

# <a name="Cancel"></a>Cancel

Cancels execution of a pending, asynchronous method call.

```
FUNCTION Cancel () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use the **Cancel** method to terminate execution of an asynchronous method call (that is, a method invoked with the **adAsyncConnect**, **adAsyncExecute**, or **adAsyncFetch** option).

For a **Recordset** object, the last asynchronous call to the **Open** method is terminated.

# <a name="CancelBatch"></a>CancelBatch

Cancels a pending batch update.

```
FUNCTION CancelBatch (BYVAL AffectRecords AS AffectEnum = adAffectAll) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *AffectRecords* | Optional. An **AffectEnum** value that indicates how many records the **CancelBatch** method will affect. |

#### AffectEnum

Specifies which records are affected by an operation.

| Constant   | Description |
| ---------- | ----------- |
| **adAffectAll** | If there is not a Filter applied to the **Recordset**, affects all records. If the **Filter** property is set to a string criteria (such as "Author='Smith'"), then the operation affects visible records in the current chapter. If the **Filter** property is set to a member of the **FilterGroupEnum** or an array of Bookmarks, then the operation will affect all rows of the **Recordset**. |
| **adAffectAllChapters** | Affects all records in all sibling chapters of the **Recordset**, including those not visible via any **Filter** that is currently applied. |
| **adAffectCurrent** | Affects only the current record. |
| **adAffectGroup** | Affects only records that satisfy the current **Filter** property setting. You must set the **Filter** property to a **FilterGroupEnum** value or an array of Bookmarks to use this option. |

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use the **CancelBatch** method to cancel any pending updates in a **Recordset** in batch update mode. If the **Recordset** is in immediate update mode, calling **CancelBatch** without **adAffectCurrent** generates an error.

If you are editing the current record or are adding a new record when you call **CancelBatch**, ADO first calls the **CancelUpdate** method to cancel any cached changes. After that, all pending changes in the **Recordset** are canceled.

It's possible that the current record will be indeterminable after a **CancelBatch** call, especially if you were in the process of adding a new record. For this reason, it is prudent to set the current record position to a known location in the **Recordset** after the **CancelBatch** call. For example, call the **MoveFirst** method.

If the attempt to cancel the pending updates fails because of a conflict with the underlying data (for example, a record has been deleted by another user), the provider returns warnings to the **Errors** collection but does not halt program execution. A run-time error occurs only if there are conflicts on all the requested records. Use the **Filter** property (**adFilterAffectedRecords**) and the **Status** property to locate records with conflicts.

# <a name="CancelUpdate"></a>CancelUpdate

Cancels any changes made to the current or new row of a **Recordset** object before calling the **Update** method.

```
FUNCTION CancelUpdate () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use the **CancelUpdate** method to cancel any changes made to the current row or to discard a newly added row. You cannot cancel changes to the current row or a new row after you call the Update method, unless the changes are either part of a transaction that you can roll back with the **RollbackTrans** method, or part of a batch update. In the case of a batch update, you can cancel the **Update** with the **CancelUpdate** or **CancelBatch** method.

If you are adding a new row when you call the **CancelUpdate** method, the current row becomes the row that was current before the **AddNew** call.

If you are in edit mode and want to move off the current record (for example, with **Move**, **NextRecordset**, or **Close**), you can use **CancelUpdate** to cancel any pending changes. You may need to do this if the update cannot successfully be posted to the data source (for example, an attempted delete that fails due to referential integrity violations will leave the **Recordset** in edit mode after a call to **Delete_**).

# <a name="Clone"></a>Clone

Creates a duplicate **Recordset** object from an existing **Recordset object**. Optionally, specifies that the clone be read-only.

```
FUNCTION Clone (BYVAL nLockType AS LockTypeEnum = adLockUnspecified) AS Afx_AdoRecordset PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nLockType* | Optional. A **LockTypeEnum** value that specifies either the lock type of the original **Recordset**, or a read-only Recordset. Valid values are **adLockUnspecified** or **adLockReadOnly**. |

#### LockTypeEnum

Specifies the type of lock placed on records during editing.

| Constant   | Description |
| ---------- | ----------- |
| **adLockBatchOptimistic** | Indicates optimistic batch updates. Required for batch update mode. |
| **adLockOptimistic** | Indicates optimistic locking, record by record. The provider uses optimistic locking, locking records only when you call the **Update** method. |
| **adLockPessimistic** | Indicates pessimistic locking, record by record. The provider does what is necessary to ensure successful editing of the records, usually by locking records at the data source immediately after editing. |
| **adLockReadOnly** | Indicates read-only records. You cannot alter the data. |
| **adLockUnspecified** | Does not specify a type of lock. For clones, the clone is created with the same lock type as the original. |

#### Return value

An **Afx_ADORecordset** object reference.

#### Remarks

Use the **Clone** method to create multiple, duplicate **Recordset** objects, particularly if you want to maintain more than one current record in a given set of records. Using the **Clone** method is more efficient than creating and opening a new **Recordset** object with the same definition as the original.

The **Filter** property of the original **Recordset**, if any, will not be applied to the clone. Set the **Filter** property of the new **Recordset** in order to filter the results. The simplest way to copy any existing **Filter** value is to assign it directly, like this: 

The current record of a newly created clone is set to the first record.

Changes you make to one **Recordset** object are visible in all of its clones regardless of cursor type. However, after you execute **Requery** on the original **Recordset**, the clones will no longer be synchronized to the original.

Closing the original **Recordset** does not close its copies, nor does closing a copy close the original or any of the other copies.

You can only clone a **Recordset** object that supports bookmarks. Bookmark values are interchangeable; that is, a bookmark reference from one **Recordset** object refers to the same record in any of its clones.

Some **Recordset** events that are triggered will also fire in all **Recordset** clones. However, because the current record can differ between cloned **Recordsets**, the events may not be valid for the clone. For example, if you change a value of a field, a **WillChangeField** event will occur in the changed **Recordset** and in all clones. The **Fields** parameter of the **WillChangeField** event of a cloned **Recordset** (where the change was not made) will simply refer to the fields of the current record of the clone, which may be a different record than the current record of the original **Recordset** where the change occurred.

The following table provided a full listing of all **Recordset** events and indicates whether they are valid and triggered for any recordset clones generated using the **Clone** method.

| Event      | Triggered in clones? |
| ---------- | -------------------- |
| **EndOfRecordset** | No |
| **FetchComplete** | No |
| **FetchProgress** | No |
| **FieldChangeComplete** | Yes |
| **MoveComplete** | No |
| **RecordChangeComplete** | Yes |
| **RecordsetChangeComplete** | No |
| **WillChangeField** | Yes |
| **WillChangeRecord** | Yes |
| **WillChangeRecordset** | No |
| **WillMove** | No |

# <a name="Close"></a>Close

Closes a **Recordset** object and any dependent objects.

```
FUNCTION Close () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use the **Close** method to close a **Recordset** to free any associated system resources. Closing an object does not remove it from memory; you can change its property settings and open it again later. To completely eliminate an object from memory, release the connection calling the **Release** method.

# <a name="Collect"></a>Collect

Sets or returns a Variant value that indicates the value of the object

The ADO Recorset object exposes a hidden member: the **Collect** property. This property is functionally similar to the **Field**'s **Value** property, but it doesn't need a reference (explicit or implicit) to the **Field** object. You can pass either a numeric index or a field's name to this property.

```
PROPERTY Collect (BYREF cvIndex AS CVAR) AS CVAR
PROPERTY Collect (BYREF cvIndex AS CVAR, BYREF cvValue AS CVAR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cvIndex* | The zero-based ordinal number of the field or the name of the field. |
| *cvValue* | The value to assign to the field. |

#### Return value

The value of the field.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Create a Connection object
DIM pConnection AS CAdoConnection
' // Create a Recordset object
DIM pRecordset AS CAdoRecordset

' // Open the connection
DIM cvConStr AS CVAR = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"
pConnection.Open cvConStr

' // Open the recordset
DIM cvSource AS CVAR = "Publishers"
pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdTableDirect)

' // Add a new record
pRecordset.AddNew
   pRecordset.Collect("PubID") = CLNG(10000)
   pRecordset.Collect("Name") = "Wile E. Coyote"
   pRecordset.Collect("Company Name") = "Warner Brothers Studios"
   pRecordset.Collect("Address") = "4000 Warner Boulevard"
   pRecordset.Collect("City") = "Burbank, CA. 91522"
pRecordset.Update
```

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Set the cursor location
DIM pRecordset AS CAdoRecordset
pRecordset.CursorLocation = adUseClient

' // Open the recordset
DIM cvSource AS CVAR = "Publishers"
pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdTable)

' // Parse the recordset
DO
   ' // While not at the end of the recordset...
   IF pRecordset.EOF THEN EXIT DO
   ' // Get the contents of the "City" and "Name" columns
   DIM cvRes AS CVAR = pRecordset.Collect("Name")
   PRINT "Position: "; pRecordset.AbsolutePosition; " "; cvRes

   ' // Fetch the next row
   IF pRecordset.MoveNext <> S_OK THEN EXIT DO
LOOP
```

# <a name="CompareBookmarks"></a>CompareBookmarks

Compares two bookmarks and returns an indication of their relative values.

```
FUNCTION CompareBookmarks (BYREF cvBookmark1 AS CVAR, BYREF cvBookmark2 AS CVAR) AS CompareEnum
```

| Parameter  | Description |
| ---------- | ----------- |
| *cvBookmark1* | The bookmark of the first row. |
| *cvBookmark2* | The bookmark of the second row. |

#### Return value

A **CompareEnum** value.

#### CompareEnum

Specifies the relative position of two records represented by their bookmarks.

| Constant   | Description |
| ---------- | ----------- |
| **adCompareEqual** | Indicates that the bookmarks are equal. |
| **adCompareGreaterThan** | Indicates that the first bookmark is after the second. |
| **adCompareLessThan** | Indicates that the first bookmark is before the second. |
| **adCompareNotComparable** | Indicates that the bookmarks cannot be compared. |
| **adCompareNotEqual** | Indicates that the bookmarks are not equal and not ordered. |

#### Remarks

The bookmarks must apply to the same **Recordset** object, or a **Recordset** object and its clone. You cannot reliably compare bookmarks from different **Recordset** objects, even if they were created from the same source or command. Nor can you compare bookmarks for a **Recordset** object whose underlying provider does not support comparisons.

A bookmark uniquely identifies a row in a **Recordset** object. Use the current row's **Bookmark** property to obtain its bookmark.

Because the data type of a bookmark is provider specific, ADO exposes it as a Variant. For example, SQL Server bookmarks are of type DBTYPE_R8 (Double). ADO would expose this type as a Variant with a subtype of Double.

When comparing bookmarks, ADO does not attempt any type of coercion. The values are simply passed to the provider where the compare occurs. If bookmarks passed to the **CompareBookmarks** method are stored in variables of differing types, it can generate the type mismatch error, "Arguments are of the wrong type, are out of the acceptable range, or are in conflict with each other."

A bookmark that is not valid or incorrectly formed will cause an error.

# <a name="CursorLocation"></a>CursorLocation

Indicates the location of the cursor service.

```
PROPERTY CursorLocation () AS CursorLocationEnum
PROPERTY CursorLocation (BYVAL lCursorLoc AS CursorLocationEnum)
```

| Parameter  | Description |
| ---------- | ----------- |
| *lCursorLoc* | One of the **CursorLocationEnum** values |

#### Return value

A **CursorLocationEnum** value.

#### CursorLocationEnum

Specifies the location of the cursor service.

| Constant   | Description |
| ---------- | ----------- |
| **adUseClient** | Uses client-side cursors supplied by a local cursor library. Local cursor services often will allow many features that driver-supplied cursors may not, so using this setting may provide an advantage with respect to features that will be enabled. For backward compatibility, the synonym **adUseClientBatch** is also supported. |
| **adUseNone** | Does not use cursor services. (This constant is obsolete and appears solely for the sake of backward compatibility.) |
| **adUseServer** | Default. Uses data-provider or driver-supplied cursors. These cursors are sometimes very flexible and allow for additional sensitivity to changes others make to the data source. However, some features of the Microsoft Cursor Service for OLE DB (such as disassociated Recordset objects) cannot be simulated with server-side cursors and these features will be unavailable with this setting. |

# <a name="CursorType"></a>CursorType

Sets or returns a **CursorTypeEnum** value. The default value is **adOpenForwardOnly**.

```
PROPERTY CursorType () AS CursorTypeEnum
PROPERTY CursorType (BYVAL lCursorType AS CursorTypeEnum)
```

| Parameter  | Description |
| ---------- | ----------- |
| *lCursorType* | One of the **CursorTypeEnum** values. |

#### Return value

A **CursorTypeEnum** value.

#### CursorTypeEnum

Specifies the type of cursor used in a **Recordset** object.

| Constant   | Description |
| ---------- | ----------- |
| **adOpenDynamic** | Uses a dynamic cursor. Additions, changes, and deletions by other users are visible, and all types of movement through the **Recordset** are allowed, except for bookmarks, if the provider doesn't support them. |
| **adOpenForwardOnly** | Default. Uses a forward-only cursor. Identical to a static cursor, except that you can only scroll forward through records. This improves performance when you need to make only one pass through a **Recordset**. |
| **adOpenKeyset** | Uses a keyset cursor. Like a dynamic cursor, except that you can't see records that other users add, although records that other users delete are inaccessible from your Recordset. Data changes by other users are still visible. |
| **adOpenStatic** | Uses a static cursor, which is a static copy of a set of records that you can use to find data or generate reports. Additions, changes, or deletions by other users are not visible. |
| **adOpenUnspecified** | Does not specify the type of cursor. |

# <a name="DataMember"></a>DataMember

Indicates the name of the data member that will be retrieved from the object referenced by the **DataSource** property.

```
PROPERTY DataMember () AS CBSTR
PROPERTY DataMember (BYVAL cbsDataMember AS CBSTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsDataMember* | Name of the data member. Not case sensitive. |

#### Return value

The name of the data member.

#### Remarks

This property is used to create data-bound controls with the Data Environment. The Data Environment maintains collections of data (data sources) containing named objects (data members) that will be represented as a **Recordset** object.

The **DataMember** and **DataSource** properties must be used in conjunction.

The **DataMember** property determines which object specified by the **DataSource** property will be represented as a **Recordset** object. The **Recordset** object must be closed before this property is set. An error is generated if the **DataMember** property isn't set before the **DataSource** property, or if the **DataMember** name isn't recognized by the object specified in the **DataSource** property.

# <a name="DataSource"></a>DataSource

Indicates an object that contains data to be represented as a **Recordset** object.

```
PROPERTY DataSource () AS IUnknown PTR
PROPERTY DataSource (BYVAL punkDataSource AS IUnknown PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *punkDataSource* | An object reference to the data source. |

#### Return value

An object reference to the data source.

#### Remarks

This property is used to create data-bound controls with the Data Environment. The Data Environment maintains collections of data (data sources) containing named objects (data members) that will be represented as a **Recordset** object.

The **DataMember** and **DataSource** properties must be used in conjunction.

The object referenced must implement the **IDataSource** interface and must contain an **IRowset** interface.


# <a name="Delete_"></a>Delete_

Deletes the current record or a group of records.

```
FUNCTION Delete_ (BYVAL AffectRecords AS AffectEnum = adAffectCurrent) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *AffectRecords* | Optional. An **AffectEnum** value that determines how many records the **Delete_** method will affect. The default value is **adAffectCurrent**. Note: **adAffectAll** and **adAffectAllChapters** are not valid arguments to Delete_. |

#### AffectEnum

Specifies which records are affected by a **Delete_** operation.

| Constant   | Description |
| ---------- | ----------- |
| **adAffectCurrent** | Affects only the current record. |
| **adAffectGroup** | Affects only records that satisfy the current **Filter** property setting. You must set the **Filter** property to a **FilterGroupEnum** value or an array of Bookmarks to use this option. |

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Using the **Delete_** method marks the current record or a group of records in a **Recordset** object for deletion. If the **Recordset** object doesn't allow record deletion, an error occurs. If you are in immediate update mode, deletions occur in the database immediately. If a record cannot be successfully deleted (due to database integrity violations, for example), the record will remain in edit mode after the call to **Update**. This means that you must cancel the update with **CancelUpdate** before moving off the current record (for example, with **Close**, **Move**, or **NextRecordset**).

If you are in batch update mode, the records are marked for deletion from the cache and the actual deletion happens when you call the **UpdateBatch** method. (Use the **Filter** property to view the deleted records.)

Retrieving field values from the deleted record generates an error. After deleting the current record, the deleted record remains current until you move to a different record. Once you move away from the deleted record, it is no longer accessible.

If you nest deletions in a transaction, you can recover deleted records with the RollbackTrans method. If you are in batch update mode, you can cancel a pending deletion or group of pending deletions with the **CancelBatch** method.

If the attempt to delete records fails because of a conflict with the underlying data (for example, a record has already been deleted by another user), the provider returns warnings to the **Errors** collection but does not halt program execution. A run-time error occurs only if there are conflicts on all the requested records.

If the Unique Table dynamic property is set, and the **Recordset** is the result of executing a JOIN operation on multiple tables, then the **Delete_** method will only delete rows from the table named in the Unique Table property.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Open the recordset
DIM pRecordset AS CAdoRecordset
DIM cvSource AS CVAR = "SELECT * FROM Publishers WHERE PubID=10000"
pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

DIM cvRes AS CVAR = pRecordset.Collect("PubID")
IF cvRes.ValInt = 10000 THEN
   pRecordset.Delete_ adAffectCurrent
   PRINT "Record deleted"
ELSE
   PRINT "Record not found"
END IF
```

# <a name="EditMode"></a>EditMode

Indicates the editing status of the current record.

```
PROPERTY EditMode () AS LONG
```

#### EditModeEnum

Specifies the editing status of a record.

| Constant   | Description |
| ---------- | ----------- |
| **adEditNone** | Indicates that no editing operation is in progress. |
| **adEditInProgress** | Indicates that data in the current record has been modified but not saved. |
| **adEditAdd** | Indicates that the **AddNew** method has been called, and the current record in the copy buffer is a new record that has not been saved in the database. |
| **adEditDelete** | Indicates that the current record has been deleted. |

# <a name="Fields"></a>Fields

Gets a reference to the **Fields** collection of a **Record** object.

```
PROPERTY Fields () AS Afx_ADOFields PTR
```

#### Return value

An **Afx_ADOFields** object reference.

# <a name="Filter"></a>Filter

Indicates a filter for data in a **Recordset**.

```
PROPERTY Filter () AS CVAR
PROPERTY Filter (BYVAL cvCriteria AS CVAR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cvCriteria* | Can be one of the following values:<br><ul><li>Criteria string — a string made up of one or more individual clauses concatenated with AND or OR operators.</li><li>Array of bookmarks — an array of unique bookmark values that point to records in the **Recordset** object.</li></li><li>A **FilterGroupEnum** value.</li></ul> |

#### Return value

The filter criteria value.

#### Remarks

Use the **Filter** property to selectively screen out records in a **Recordset** object. The filtered **Recordset** becomes the current cursor. Other properties that return values based on the current cursor are affected, such as **AbsolutePosition**, **AbsolutePage**, **RecordCount**, and **PageCount**. This is because setting the **Filter** property to a specific value will move the current record to the first record that satisfies the new value.

The criteria string is made up of clauses in the form FieldName-Operator-Value (for example, "LastName = 'Smith'"). You can create compound clauses by concatenating individual clauses with AND (for example, "LastName = 'Smith' AND FirstName = 'John'") or OR (for example, "LastName = 'Smith' OR LastName = 'Jones'"). Use the following guidelines for criteria strings:

* FieldName must be a valid field name from the Recordset. If the field name contains spaces, you must enclose the name in square brackets.
* Operator must be one of the following: <, >, <=, >=, <>, =, or LIKE.
* Value is the value with which you will compare the field values (for example, 'Smith', #8/24/95#, 12.345, or $50.00). Use single quotes with strings and pound signs (#) with dates. For numbers, you can use decimal points, dollar signs, and scientific notation. If Operator is LIKE, Value can use wildcards. Only the asterisk (\*) and percent sign (%) wild cards are allowed, and they must be the last character in the string. Value cannot be null.

    * **Note**: To include single quotation marks (') in the filter Value, use two single quotation marks to represent one. For example, to filter on O'Malley, the criteria string should be "col1 = 'O''Malley'". To include single quotation marks at both the beginning and the end of the filter value, enclose the string with pound signs (#). For example, to filter on '1', the criteria string should be "col1 = #'1'#".

* There is no precedence between AND and OR. Clauses can be grouped within parentheses. However, you cannot group clauses joined by an OR and then join the group to another clause with an AND, like this:

    (LastName = 'Smith' OR LastName = 'Jones') AND FirstName = 'John'

    Instead, you would construct this filter as

    (LastName = 'Smith' AND FirstName = 'John') OR (LastName = 'Jones' AND FirstName = 'John')

* In a LIKE clause, you can use a wildcard at the beginning and end of the pattern (for example, LastName Like '*mit*'), or only at the end of the pattern (for example, LastName Like 'Smit*').

The filter constants make it easier to resolve individual record conflicts during batch update mode by allowing you to view, for example, only those records that were affected during the last **UpdateBatch** method call.

Setting the **Filter** property itself may fail because of a conflict with the underlying data (for example, a record has already been deleted by another user). In such a case, the provider returns warnings to the Errors collection but does not halt program execution. A run-time error occurs only if there are conflicts on all the requested records. Use the **Status** property to locate records with conflicts.

Setting the **Filter** property to a zero-length string ("") has the same effect as using the **adFilterNone** constant.

Whenever the **Filter** property is set, the current record position moves to the first record in the filtered subset of records in the Recordset. Similarly, when the Filter property is cleared, the current record position moves to the first record in the **Recordset**.

When a **Recordset** is filtered based on a field of some variant type (e.g., sql_variant), an error (DISP_E_TYPEMISMATCH or 80020005) will result if the subtypes of the field and filter values used in the criteria string do not match. For example, if a **Recordset** object (*lpRecordset*) contains a column (C) of the sql_variant type and a field of this column has been assigned a value of 1 of the I4 type, setting the criteria string of **Recordset** **Filter** "C='A'" on the field will produce the error at run time. However, **Recordset** **Filter** "C=2" applied on the same field will not produce any error although the field will be filtered out of the current record set.

See the **Bookmark** property for an explanation of bookmark values from which you can build an array to use with the **Filter** property.

Only Filters in the form of Criteria Strings (e.g. OrderDate > '12/31/1999') affect the contents of a persisted **Recordset**. Filters created with an Array of Bookmarks or using a value from the **FilterGroupEnum** will not affect the contents of the persisted Recordset. These rules apply to Recordsets created with either client-side or server-side cursors.

**Note**: When you apply the **adFilterPendingRecords** flag to a filtered and modified **Recordset** in the batch update mode, the resultant **Recordset** is empty if the filtering was based on the key field of a single-keyed table and the modification was made on the key field values. The resultant **Recordset** will be non-empty if one of the following is true:

* The filtering was based on non-key fields in a single-keyed table.
* The filtering was based on any fields in a multiple-keyed table.
* Modifications were made on non-key fields in a single-keyed table.
* Modifications were made on any fields in a multiple-keyed table.

The following table summarizes the effects of **adFilterPendingRecords** in different combinations of filtering and modifications. The left column shows the possible modifications; modifications can be made on any of the non-keyed fields, on the key field in a single-keyed table, or on any of the key fields in a multiple-keyed table. The top row shows the filtering criterion; filtering can be based on any of the non-keyed fields, the key field in a single-keyed table, or any of the key fields in a multiple-keyed table. The intersecting cells show the results: + means that applying adFilterPendingRecords results in a non-empty **Recordset**; - means an empty **Recordset**.

|     | Non Keys    | Single Key  | Multiple Keys |
| --- | ----------- | ----------- | ------------- |
| Non Keys | + | + | + |
| Single Key | + | - | N/A |
| Multiple Key | + | N/A | + |

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Set the cursor location
DIM pRecordset AS CAdoRecordset
pRecordset.CursorLocation = adUseClient

' // Open the recordset
DIM cvSource AS CVAR = "Publishers"
DIM hr AS HRESULT = pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdTableDirect)

' // Set the Filter property
pRecordset.Filter = "City = 'New York'"

' // Parse the recordset
DO
   ' // While not at the end of the recordset...
   IF pRecordset.EOF THEN EXIT DO
   ' // Get the contents of the "City" and "Name" columns
   DIM cvRes1 AS CVAR = pRecordset.Collect("City")
   DIM cvRes2 AS CVAR = pRecordset.Collect("Name")
   PRINT cvRes1 & " " & cvRes2
   ' // Fetch the next row
   IF pRecordset.MoveNext <> S_OK THEN EXIT DO
LOOP
```

# <a name="Find"></a>Find

Searches a **Recordset** for the row that satisfies the specified criteria. Optionally, the direction of the search, starting row, and offset from the starting row may be specified. If the criteria is met, the current row position is set on the found record; otherwise, the position is set to the end (or start) of the **Recordset**.

**Note**: The **Find** method is a single column operation only because the OLE DB specification defines **IRowsetFind** this way.


```
FUNCTION Find (BYREF cbsCriteria AS CBSTR, _
   BYREF cvStart AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
   BYVAL SkipRecords AS LONG = 0, BYVAL SearchDirection AS SearchDirectionEnum = adSearchForward) AS HRESULT
```
```
FUNCTION Find ( BYREF cbsCriteria AS CBSTR, BYVAL SkipRecords AS LONG = 0, _
   BYVAL SearchDirection AS SearchDirectionEnum = adSearchForward) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsCriteria* | A string value that contains a statement specifying the column name, comparison operator, and value to use in the search. |
| *cvStart* | Optional. A bookmark that functions as the starting position for the search. |
| *SkipRecords* | Optional. A Long value, whose default value is zero, that specifies the row offset from the current row or start bookmark to begin the search. By default, the search will start on the current row. |
| *SearchDirection* | Optional. A **SearchDirectionEnum** value that specifies whether the search should begin on the current row or the next available row in the direction of the search. An unsuccessful search stops at the end of the **Recordset** if the value is **adSearchForward**. An unsuccessful search stops at the start of the **Recordset** if the value is **adSearchBackward**. |

#### SearchDirectionEnum

Specifies the direction of a record search within a Recordset.

| Constant   | Description |
| ---------- | ----------- |
| **adSearchBackward** | Searches backward, stopping at the beginning of the **Recordset**. If a match is not found, the record pointer is positioned at **BOF**. |
| **adSearchForward** | Searches forward, stopping at the end of the **Recordset**. If a match is not found, the record pointer is positioned at **EOF**. |


#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Only a single-column name may be specified in criteria. This method does not support multi-column searches.

The comparison operator in **Criteria** may be ">" (greater than), "<" (less than), "=" (equal), ">=" (greater than or equal), "<=" (less than or equal), "<>" (not equal), or "like" (pattern matching).

The value in **Criteria** may be a string, floating-point number, or date. String values are delimited with single quotes or "#" (number sign) marks (for example, "state = 'WA'" or "state = #WA#"). Date values are delimited with "#" (number sign) marks (for example, "start_date > #7/22/97#"). These values can contain hours, minutes, and seconds to indicate time stamps, but should not contain milliseconds or errors will occur.

If the comparison operator is "like", the string value may contain an asterisk (\*) to find one or more occurrences of any character or substring. For example, "state like 'M*'" matches Maine and Massachusetts. You can also use leading and trailing asterisks to find a substring contained within the values. For example, "state like '*as*'" matches Alaska, Arkansas, and Massachusetts.

Asterisks can be used only at the end of a criteria string, or together at both the beginning and end of a criteria string, as shown above. You cannot use the asterisk as a leading wildcard ('\*str'), or embedded wildcard ('s\*r'). This will cause an error.

**Note**: An error will occur if a current row position is not set before calling **Find**. Any method that sets row position, such as **MoveFirst**, should be called before calling **Find**.

**Note**: If you call the **Find** method on a recordset, and the current position in the recordset is at the last record or end of file (EOF), you will not find anything. You need to call the **MoveFirsr** method to set the current position/cursor to the beginning of the recordset.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
DIM cvConStr AS CVAR = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"
pConnection.Open cvConStr

' // Open the recordset
DIM pRecordset AS CAdoRecordset
DIM cvSource AS CVAR = "SELECT * FROM Publishers ORDER By PubID"
DIM hr AS HRESULT = pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

pRecordset.Find "PubID = #70#"
DIM cvRes1 AS CVAR = pRecordset.Collect("PubID")
DIM cvRes2 AS CVAR = pRecordset.Collect("Name")
PRINT cvRes1 & " " & cvRes2
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

# <a name="GetRows"></a>GetRows

Retrieves multiple records of a **Recordset** object into an array.

```
FUNCTION GetRows (BYVAL Rows AS LONG = adGetRowsRest, BYREF cvStart AS CVAR = "", _
   BYREF cvFields AS CVAR = "") AS SAFEARRAY PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *Rows* | Optional. A **GetRowsOptionEnum** value that indicates the number of records to retrieve. The default is **adGetRowsRest**. |
| *cvStart* | Optional. A string value or Variant that evaluates to the bookmark for the record from which the **GetRows** operation should begin. You can also use a **BookmarkEnum** value. |
| *cvFields* | Optional. A Variant that represents a single field name or ordinal position, or an array of field names or ordinal position numbers. ADO returns only the data in these fields. |

#### BookmarkEnum

Specifies a bookmark indicating where the operation should begin.

| Parameter  | Description |
| ---------- | ----------- |
| **adBookmarkCurrent** | Starts at the current record. |
| **adBookmarkFirst** | Starts at the first record. |
| **adBookmarkLast** | Starts at the last record. |

#### Return value

SAFEARRAY. The array of records.

#### Remarks

Use the **GetRows** method to copy records from a **Recordset** into a two-dimensional array. The first subscript identifies the field and the second identifies the record number. The array variable is automatically dimensioned to the correct size when the **GetRows** method returns the data.

If you do not specify a value for the Rows argument, the **GetRows** method automatically retrieves all the records in the **Recordset** object. If you request more records than are available, **GetRows** returns only the number of available records.

If the **Recordset** object supports bookmarks, you can specify at which record the **GetRows** method should begin retrieving data by passing the value of that record's **Bookmark** property in the **Start** argument.

If you want to restrict the fields that the **GetRows** call returns, you can pass either a single field name/number or an array of field names/numbers in the **Fields** argument.

After you call **GetRows**, the next unread record becomes the current record, or the **EOF** property is set to True if there are no more records.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Create a Connection object
DIM pConnection AS CAdoConnection
' // Create a Recordset object
DIM pRecordset AS CAdoRecordset
' // Open the connection
DIM cvConStr AS CVAR = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"
pConnection.Open cvConStr
' // Open the recordset
DIM cvSource AS CVAR = "SELECT TOP 20 * FROM Publishers ORDER BY Name"
DIM hr AS HRESULT = pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

' // GetRows returns a pointer to a two-dimensional safe array
' // that we are going to attach to an instance of the CSafeArray class
DIM csa AS CSafeArray
csa.Attach(pRecordset.GetRows)
'--or--
' DIM csa AS CSafeArray = CSafeArray(pRecordset.GetRows, TRUE)

' // Print the contents of the safe array
FOR j AS LONG = csa.LBound(2) TO csa.UBound(2)
   FOR i AS LONG = csa.LBound(1) TO csa.UBound(1)
      PRINT csa.GetVar(i, j)
   NEXT
   PRINT
NEXT
```

# <a name="GetString"></a>GetString

Returns the **Recordset** as a string.

```
FUNCTION GetString (BYVAL StringFormat AS StringFormatEnum, BYVAL NumRows AS LONG, _
   BYREF ColumnDelimeter AS CBSTR = "", BYREF RowDelimeter AS CBSTR = "", _
   BYREF NullExpr AS CBSTR = "") AS CBSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *StringFormat* | Optional. A **StringFormatEnum** value that specifies how the **Recordset** should be converted to a string. The *RowDelimiter*, *ColumnDelimiter*, and *NullExpr* parameters are used only with a **StringFormat** of **adClipString**. |
| *NumRows* | Optional. The number of rows to be converted in the **Recordset**. If *NumRows* is not specified, or if it is greater than the total number of rows in the **Recordset**, then all the rows in the **Recordset** are converted. |
| *ColumnDelimiter* | Optional. A delimiter used between columns, if specified, otherwise the TAB character |
| *RowDelimiter* | Optional. A delimiter used between rows, if specified, otherwise the CARRIAGE RETURN character. |
| *NullExpr* | Optional. An expression used in place of a null value, if specified, otherwise the empty string. |

#### Return value

The **Recordset** as a string.

#### Remarks

Row data, but no schema data, is saved to the string. Therefore, a **Recordset** cannot be reopened using this string.

This method is equivalent to the **RDO** **GetClipString** method.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Create a Connection object
DIM pConnection AS CAdoConnection
' // Create a Recordset object
DIM pRecordset AS CAdoRecordset
' // Open the connection
DIM cvConStr AS CVAR = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"
pConnection.Open cvConStr
' // Open the recordset
DIM cvSource AS CVAR = "SELECT * FROM Publishers"
DIM hr AS HRESULT = pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)
PRINT pRecordset.GetString
```

# <a name="Index"></a>Index

Indicates the name of the index currently in effect for a **Recordset** object.

Sets or returns a string value, which is the name of the index.

```
PROPERTY Index () AS CBSTR
PROPERTY Index (BYREF cbsIndex AS CBSTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsIndex* | The name of the index. |

#### Return value

CBSTR. The name of the current index.

#### Remarks

The index named by the **Index** property must have previously been declared on the base table underlying the **Recordset** object. That is, the index must have been declared programmatically either as an **ADOX** **Index** object, or when the base table was created.

A run-time error will occur if the index cannot be set. The **Index** property cannot be set:

* Within a **WillChangeRecordset** or **RecordsetChangeComplete** event handler.
* If the **Recordset** is still executing an operation (which can be determined by the **State** property).
* If a filter has been set on the **Recordset** with the **Filter** property.

The **Index** property can always be set successfully if the **Recordset** is closed, but the **Recordset** will not open successfully, or the index will not be usable, if the underlying provider does not support indexes.

If the index can be set, the current row position may change. This will cause an update to the **AbsolutePosition** property, and the generation of **WillChangeRecordset**, **RecordsetChangeComplete**, **WillMove**, and **MoveComplete** events.

If the index can be set and the **LockType** property is **adLockPessimistic** or **adLockOptimistic**, then an implicit **UpdateBatch** operation is performed. This releases the current and affected groups. Any existing filter is released, and the current row position is changed to the first row of the reordered **Recordset**.

The **Index** property is used in conjunction with the **Seek** method. If the underlying provider does not support the **Index** property, and thus the **Seek** method, consider using the **Find** method instead. Determine whether the **Recordset** object supports indexes with the **Supports**(*adIndex*) method.

The built-in **Index** property is not related to the dynamic **Optimize** property, although they both deal with indexes.

**Note**: The SQL Server 6.5 or 7.0 doesn't support the **Seek** or **Index** methods of the **Recordset**. You can however, use the **Index** and **Seek** method with an Access 2000 database and the OLE DB 4.0 Provider for Jet.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Set the cursor location
DIM pRecordset AS CAdoRecordset
pRecordset.CursorLocation = adUseServer

' // Open the recordset
DIM cvSource AS CVAR = "Publishers"
DIM hr AS HRESULT = pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdTableDirect)

' // Set the index
pRecordset.Index = "PrimaryKey"

' // Seek the record 70
pRecordset.Seek 70, 1

' // Parse the recordset
DO
   ' // While not at the end of the recordset...
   IF pRecordset.EOF THEN EXIT DO
   ' // Get the contents
   DIM cvRes1 AS CVAR = pRecordset.Collect("PubID")
   DIM cvRes2 AS CVAR = pRecordset.Collect("Name")
   DIM cvRes3 AS CVAR = pRecordset.Collect("Company Name")
   PRINT cvRes1 & " " & cvRes2 & " " & cvRes3
   ' // Fetch the next row
   IF pRecordset.MoveNext <> S_OK THEN EXIT DO
LOOP
```

# <a name="LockType"></a>LockType

Sets or returns the lock type, a Long value that must be greater than 0. Default is 1.

```
PROPERTY LockType () AS LockTypeEnum
PROPERTY LockType (BYVAL lLockType AS LockTypeEnum)
```

| Parameter  | Description |
| ---------- | ----------- |
| *lLockType* | A **LockTypeEnum** value. The default value is **adLockReadOnly**. |

#### Return value

A **LockTypeEnum** value.

#### LockTypeENum

Specifies the type of lock placed on records during editing.

| Constant   | Description |
| ---------- | ----------- |
| **adLockBatchOptimistic** | Indicates optimistic batch updates. Required for batch update mode. |
| **adLockOptimistic** | Indicates optimistic locking, record by record. The provider uses optimistic locking, locking records only when you call the **Update** method. |
| **adLockPessimistic** | Indicates pessimistic locking, record by record. The provider does what is necessary to ensure successful editing of the records, usually by locking records at the data source immediately after editing. |
| **adLockReadOnly** | Indicates read-only records. You cannot alter the data. |
| **adLockUnspecified** | Does not specify a type of lock. For clones, the clone is created with the same lock type as the original. |

#### Remarks

Set the **LockType** property before opening a **Recordset** to specify what type of locking the provider should use when opening it. Read the property to return the type of locking in use on an open **Recordset** object.

Providers may not support all lock types. If a provider cannot support the requested **LockType** setting, it will substitute another type of locking. To determine the actual locking functionality available in a **Recordset** object, use the **Supports** method with **adUpdate** and **adUpdateBatch**.

The **adLockPessimistic** setting is not supported if the **CursorLocation** property is set to **adUseClient**. If an unsupported value is set, then no error will result; the closest supported **LockType** will be used instead.

The **LockType** property is read/write when the **Recordset** is closed and read-only when it is open.

**Remote Data Service Usage**: When used on a client-side **Recordset** object, the **LockType** property can only be set to **adLockBatchOptimistic**.

# <a name="MarshalOptions"></a>MarshalOptions

Indicates which records are to be marshaled back to the server.

```
PROPERTY MarshalOptions () AS MarshalOptionsEnum
PROPERTY MarshalOptions (BYVAL eMarshal AS MarshalOptionsEnum)
```

| Parameter  | Description |
| ---------- | ----------- |
| *eMarshal* | A **MarshalOptionEnum** value. The default value is **adMarshalAll**. |

#### Return value

A **MarshalOptionEnum** value.

#### MarshalOptionEnum

Specifies which records should be returned to the server.

| Constant   | Description |
| ---------- | ----------- |
| **adMarshalAll** | Default. Returns all rows to the server. |
| **adMarshalModifiedOnly** | Returns only modified rows to the server. |

#### Remarks

When using a client-side **Recordset**, records that have been modified on the client are written back to the middle tier or Web server through a technique called marshaling, the process of packaging and sending interface method parameters across thread or process boundaries. Setting the **MarshalOptions** property can improve performance when modified remote data is marshaled for updating back to the middle tier or Web server.

**Remote Data Service Usage**: This property is used only on a client-side **Recordset**.

# <a name="MaxRecords"></a>MaxRecords

Indicates the maximum number of records to return to a **Recordset** from a query.

Sets or returns a Long value that indicates the maximum number of records to return. Default is zero (no limit).

```
PROPERTY MaxRecords () AS LONG
PROPERTY MaxRecords (BYVAL lMaxRecords AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *lMaxRecords* | A value that indicates the maximum number of records to return. Default is zero (no limit). |

#### Return value

The maximum number of records.

#### Remarks

Use the **MaxRecords** property to limit the number of records that the provider returns from the data source. The default setting of this property is zero, which means the provider returns all requested records.

The **MaxRecords** property is read/write when the **Recordset** is closed and read-only when it is open.

#### Problem

Using the **MaxRecords** property of an ADO Recordset object does not affect the number of records returned in queries against Microsoft Access databases.

The **MaxRecords** property depends on functionality exposed by the underlying OLE DB provider or ODBC driver to limit the number of rows returned by the query. This functionality is optional for ODBC drivers and OLE DB providers. The Microsoft Access ODBC driver and Jet OLEDB provider do not expose this functionality. This behavior may occur for other OLE DB providers and ODBC drivers as well.

If you want to limit the number of records returned in a query against a Microsoft Access database, use the **TOP** syntax in the query string rather than the **Recordset**'s **MaxRecords** property.

For example:

```
SELECT TOP 2 EmployeeID, LastName, FirstName FROM Employees ORDER BY LastName, FirstName;
```

or

```
SELECT TOP 20 PERCENT EmployeeID, LastName, FirstName FROM Employees ORDER BY LastName, FirstName;
```

# <a name="Move"></a>Move

Moves the position of the current record in a **Recordset** object.

```
FUNCTION Move (BYVAL NumRecords AS LONG, BYVAL start AS BookmarkEnum = adBookmarkCurrent) AS HRESULT
FUNCTION Move (BYVAL NumRecords AS LONG, BYREF cvStart AS CVAR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *NumRecords* | A signed Long expression that specifies the number of records that the current record position moves. |
| *start* | Optional. A **BookmarkEnum** value. |
| *cvStart* | Optional. A Variant that evaluates to a bookmark. |

#### BookmarkEnum

Specifies a bookmark indicating where the operation should begin.

| Parameter  | Description |
| ---------- | ----------- |
| **adBookmarkCurrent** | Starts at the current record. |
| **adBookmarkFirst** | Starts at the first record. |
| **adBookmarkLast** | Starts at the last record. |

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

The **Move** method is supported on all **Recordset** objects.

If the *NumRecords* argument is greater than zero, the current record position moves forward (toward the end of the **Recordset**). If *NumRecords* is less than zero, the current record position moves backward (toward the beginning of the **Recordset**).

If the **Move** call would move the current record position to a point before the first record, ADO sets the current record to the position before the first record in the recordset (**BOF** is True). An attempt to move backward when the **BOF** property is already True generates an error.

If the **Move** call would move the current record position to a point after the last record, ADO sets the current record to the position after the last record in the recordset (**EOF** is True). An attempt to move forward when the **EOF** property is already True generates an error.

Calling the Move method from an empty **Recordset** object generates an error.

If you pass the *start* argument, the move is relative to the record with this bookmark, assuming the **Recordset** object supports bookmarks. If not specified, the move is relative to the current record.

If you are using the **CacheSize** property to locally cache records from the provider, passing a *NumRecords* argument that moves the current record position outside the current group of cached records forces ADO to retrieve a new group of records, starting from the destination record. The **CacheSize** property determines the size of the newly retrieved group, and the destination record is the first record retrieved.

If the **Recordset** object is forward only, a user can still pass a *NumRecords* argument less than zero, provided the destination is within the current set of cached records. If the **Move** call would move the current record position to a record before the first cached record, an error will occur. Thus, you can use a record cache that supports full scrolling over a provider that supports only forward scrolling. Because cached records are loaded into memory, you should avoid caching more records than is necessary. Even if a forward-only **Recordset** object supports backward moves in this way, calling the **MovePrevious** method on any forward-only **Recordset** object will still generate an error.

**Note**: Support for moving backwards in a forward-only **Recordset** is not predictable, depending upon your provider. If the current record has been postioned after the last record in the **Recordset**, Move backwards may not result in the correct current position.

# <a name="MoveFirst"></a>MoveFirst

Moves to the first record in a specified **Recordset** object and makes that record the current record.

```
FUNCTION MoveFirst () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use the **MoveFirst** method to move the current record position to the first record in the **Recordset**.

Use the **MoveLast** method to move the current record position to the last record in the **Recordset**. The **Recordset** object must support bookmarks or backward cursor movement; otherwise, the method call will generate an error.

A call to either **MoveFirst** or **MoveLast** when the **Recordset** is empty (both BOF and EOF are True) generates an error.

Use the **MoveNext** method to move the current record position one record forward (toward the bottom of the **Recordset**). If the last record is the current record and you call the **MoveNext** method, ADO sets the current record to the position after the last record in the Recordset (**EOF** is True). An attempt to move forward when the **EOF** property is already True generates an error.

In ADO 2.5 and later, when the **Recordset** has been filtered or sorted and the data of the current record is changed, calling the **MoveNext** method moves the cursor two records forward from the current record. This is because when the current record is changed, the next record becomes the new current record. Calling **MoveNext** after the change moves the cursor one record forward from the new current record. This is different from the behavior in ADO 2.1 and earlier. In these earlier versions, changing the data of a current record in the sorted or filtered **Recordset** does not change the position of the current record, and **MoveNext** moves the cursor to the next record immediately after the current record.

Use the **MovePrevious** method to move the current record position one record backward (toward the top of the **Recordset**). The **Recordset** object must support bookmarks or backward cursor movement; otherwise, the method call will generate an error. If the first record is the current record and you call the **MovePrevious** method, ADO sets the current record to the position before the first record in the Recordset (BOF is True). An attempt to move backward when the **BOF** property is already True generates an error. If the **Recordset** object does not support either bookmarks or backward cursor movement, the **MovePrevious** method will generate an error.

If the **Recordset** is forward only and you want to support both forward and backward scrolling, you can use the **CacheSize** property to create a record cache that will support backward cursor movement through the **Move** method. Because cached records are loaded into memory, you should avoid caching more records than is necessary. You can call the **MoveFirst** method in a forward-only **Recordset** object; doing so may cause the provider to re-execute the command that generated the **Recordset** object.

# <a name="MoveLast"></a>MoveLast

Moves to the last record in a specified **Recordset** object and makes that record the current record.

```
FUNCTION MoveLast () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use the **MoveFirst** method to move the current record position to the first record in the **Recordset**.

Use the **MoveLast** method to move the current record position to the last record in the **Recordset**. The **Recordset** object must support bookmarks or backward cursor movement; otherwise, the method call will generate an error.

A call to either **MoveFirst** or **MoveLast** when the **Recordset** is empty (both BOF and EOF are True) generates an error.

Use the **MoveNext** method to move the current record position one record forward (toward the bottom of the **Recordset**). If the last record is the current record and you call the **MoveNext** method, ADO sets the current record to the position after the last record in the Recordset (**EOF** is True). An attempt to move forward when the **EOF** property is already True generates an error.

In ADO 2.5 and later, when the **Recordset** has been filtered or sorted and the data of the current record is changed, calling the **MoveNext** method moves the cursor two records forward from the current record. This is because when the current record is changed, the next record becomes the new current record. Calling **MoveNext** after the change moves the cursor one record forward from the new current record. This is different from the behavior in ADO 2.1 and earlier. In these earlier versions, changing the data of a current record in the sorted or filtered **Recordset** does not change the position of the current record, and **MoveNext** moves the cursor to the next record immediately after the current record.

Use the **MovePrevious** method to move the current record position one record backward (toward the top of the **Recordset**). The **Recordset** object must support bookmarks or backward cursor movement; otherwise, the method call will generate an error. If the first record is the current record and you call the **MovePrevious** method, ADO sets the current record to the position before the first record in the Recordset (BOF is True). An attempt to move backward when the **BOF** property is already True generates an error. If the **Recordset** object does not support either bookmarks or backward cursor movement, the **MovePrevious** method will generate an error.

If the **Recordset** is forward only and you want to support both forward and backward scrolling, you can use the **CacheSize** property to create a record cache that will support backward cursor movement through the **Move** method. Because cached records are loaded into memory, you should avoid caching more records than is necessary. You can call the **MoveFirst** method in a forward-only **Recordset** object; doing so may cause the provider to re-execute the command that generated the **Recordset** object.

# <a name="MoveNext"></a>MoveNext

Moves to the next record in a specified **Recordset** object and makes that record the current record.

```
FUNCTION MoveNext () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use the **MoveFirst** method to move the current record position to the first record in the **Recordset**.

Use the **MoveLast** method to move the current record position to the last record in the **Recordset**. The **Recordset** object must support bookmarks or backward cursor movement; otherwise, the method call will generate an error.

A call to either **MoveFirst** or **MoveLast** when the **Recordset** is empty (both BOF and EOF are True) generates an error.

Use the **MoveNext** method to move the current record position one record forward (toward the bottom of the **Recordset**). If the last record is the current record and you call the **MoveNext** method, ADO sets the current record to the position after the last record in the Recordset (**EOF** is True). An attempt to move forward when the **EOF** property is already True generates an error.

In ADO 2.5 and later, when the **Recordset** has been filtered or sorted and the data of the current record is changed, calling the **MoveNext** method moves the cursor two records forward from the current record. This is because when the current record is changed, the next record becomes the new current record. Calling **MoveNext** after the change moves the cursor one record forward from the new current record. This is different from the behavior in ADO 2.1 and earlier. In these earlier versions, changing the data of a current record in the sorted or filtered **Recordset** does not change the position of the current record, and **MoveNext** moves the cursor to the next record immediately after the current record.

Use the **MovePrevious** method to move the current record position one record backward (toward the top of the **Recordset**). The **Recordset** object must support bookmarks or backward cursor movement; otherwise, the method call will generate an error. If the first record is the current record and you call the **MovePrevious** method, ADO sets the current record to the position before the first record in the Recordset (BOF is True). An attempt to move backward when the **BOF** property is already True generates an error. If the **Recordset** object does not support either bookmarks or backward cursor movement, the **MovePrevious** method will generate an error.

If the **Recordset** is forward only and you want to support both forward and backward scrolling, you can use the **CacheSize** property to create a record cache that will support backward cursor movement through the **Move** method. Because cached records are loaded into memory, you should avoid caching more records than is necessary. You can call the **MoveFirst** method in a forward-only **Recordset** object; doing so may cause the provider to re-execute the command that generated the **Recordset** object.

# <a name="MovePrevious"></a>MovePrevious

Moves to the previous record in a specified **Recordset** object and makes that record the current record.

```
FUNCTION MovePrevious () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use the **MoveFirst** method to move the current record position to the first record in the **Recordset**.

Use the **MoveLast** method to move the current record position to the last record in the **Recordset**. The **Recordset** object must support bookmarks or backward cursor movement; otherwise, the method call will generate an error.

A call to either **MoveFirst** or **MoveLast** when the **Recordset** is empty (both BOF and EOF are True) generates an error.

Use the **MoveNext** method to move the current record position one record forward (toward the bottom of the **Recordset**). If the last record is the current record and you call the **MoveNext** method, ADO sets the current record to the position after the last record in the Recordset (**EOF** is True). An attempt to move forward when the **EOF** property is already True generates an error.

In ADO 2.5 and later, when the **Recordset** has been filtered or sorted and the data of the current record is changed, calling the **MoveNext** method moves the cursor two records forward from the current record. This is because when the current record is changed, the next record becomes the new current record. Calling **MoveNext** after the change moves the cursor one record forward from the new current record. This is different from the behavior in ADO 2.1 and earlier. In these earlier versions, changing the data of a current record in the sorted or filtered **Recordset** does not change the position of the current record, and **MoveNext** moves the cursor to the next record immediately after the current record.

Use the **MovePrevious** method to move the current record position one record backward (toward the top of the **Recordset**). The **Recordset** object must support bookmarks or backward cursor movement; otherwise, the method call will generate an error. If the first record is the current record and you call the **MovePrevious** method, ADO sets the current record to the position before the first record in the Recordset (BOF is True). An attempt to move backward when the **BOF** property is already True generates an error. If the **Recordset** object does not support either bookmarks or backward cursor movement, the **MovePrevious** method will generate an error.

If the **Recordset** is forward only and you want to support both forward and backward scrolling, you can use the **CacheSize** property to create a record cache that will support backward cursor movement through the **Move** method. Because cached records are loaded into memory, you should avoid caching more records than is necessary. You can call the **MoveFirst** method in a forward-only **Recordset** object; doing so may cause the provider to re-execute the command that generated the **Recordset** object.

# <a name="NextRecordset"></a>NextRecordset

Moves to the previous record in a specified **Recordset** object and makes that record the current record.

```
FUNCTION NextRecordset (BYVAL RecordsAffected AS LONG PTR = NULL) AS Afx_AdoRecordset PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *RecordsAffected* | Optional. A pointer to a Long variable to which the provider returns the number of records that the current operation affected. Note: This parameter only returns the number of records affected by an operation; it does not return a count of records from a select statement used to generate the **Recordset**. |

#### Return value

An **Afx_ADORecordset** object reference.

#### Remarks

Use the **NextRecordset** method to return the results of the next command in a compound command statement or of a stored procedure that returns multiple results. If you open a **Recordset** object based on a compound command statement (for example, "SELECT * FROM table1;SELECT * FROM table2") using the **Execute** method on a **Command** or the **Open** method on a **Recordset**, ADO executes only the first command and returns the results to recordset. To access the results of subsequent commands in the statement, call the **NextRecordset** method.

As long as there are additional results and the **Recordset** containing the compound statements is not disconnected or marshaled across process boundaries, the NextRecordset method will continue to return **Recordset** objects. If a row-returning command executes successfully but returns no records, the returned **Recordset** object will be open but empty. Test for this case by verifying that the **BOF** and **EOF** properties are both True. If a non–row-returning command executes successfully, the returned Recordset object will be closed, which you can verify by testing the State property on the **Recordset**. When there are no more results, recordset will be set to NULL.

The **NextRecordset** method is not available on a disconnected **Recordset** object, where **ActiveConnection** has been set to NULL.

If an edit is in progress while in immediate update mode, calling the **NextRecordset** method generates an error; call the Update or **CancelUpdate** method first.

To pass parameters for more than one command in the compound statement by filling the Parameters collection, or by passing an array with the original **Open** or **Execute** call, the parameters must be in the same order in the collection or array as their respective commands in the command series. You must finish reading all the results before reading output parameter values.

Your OLE DB provider determines when each command command in a compound statement is executed. The Microsoft OLE DB Provider for SQL Server, for example, executes all commands in a batch upon receiving the compound statement. The resulting Recordsets are simply returned when you call **NextRecordset**.

However, other providers may execute the next command in a statement only after **NextRecordset** is called. For these providers, if you explicitly close the **Recordset** object before stepping through the entire command statement, ADO never executes the remaining commands.

# <a name="Open"></a>Open

Opens a connection to a data source.

```
FUNCTION Open (BYREF cvSource AS CVAR, BYREF cActiveConnection AS CAdoConnection, _
   BYVAL nCursorType AS CursorTypeEnum = adOpenUnspecified, _
   BYVAL nLockType AS LockTypeEnum = adLockUnspecified, _
   BYVAL nOptions AS LONG = adCmdUnspecified) AS HRESULT
```
```
FUNCTION Open (BYREF cvSource AS CVAR, _
   BYREF cActiveConnection AS CAdoConnection = TYPE(<VARIANT>VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
   BYVAL nCursorType AS CursorTypeEnum = adOpenUnspecified, _
   BYVAL nLockType AS LockTypeEnum = adLockUnspecified, _
   BYVAL nOptions AS LONG = adCmdUnspecified) AS HRESULT
```
```
FUNCTION Open (BYREF cvSource AS CVAR, _
   BYVAL pActiveConnection AS CAdoConnection PTR, _
   BYVAL nCursorType AS CursorTypeEnum = adOpenUnspecified, _
   BYVAL nLockType AS LockTypeEnum = adLockUnspecified, _
   BYVAL nOptions AS LONG = adCmdUnspecified) AS HRESULT
```
```
FUNCTION Open (BYVAL nCursorType AS CursorTypeEnum = adOpenUnspecified, _
   BYVAL nLockType AS LockTypeEnum = adLockUnspecified, _
   BYVAL nOptions AS LONG = adCmdUnspecified) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cvSource* | Optional. A Variant that evaluates to a valid **Command** object, an SQL statement, a table name, a stored procedure call, a URL, or the name of a file or **Stream** object containing a persistently stored **Recordset**. |
| *cActiveConnection* | Optional. A valid **Connection** object. |
| *cvActiveConnection* | Optional. Either a Variant that evaluates to a valid Connection object variable name, or a string that contains **ConnectionString** parameters. |
| *pActiveConnection* | Pointer to a valid connection object. |
| *nCursorType* | Optional. A **CursorTypeEnum** value that determines the type of cursor that the provider should use when opening the **Recordset**. The default value is **adOpenForwardOnly**. |
| *nLockType* | Optional. A **LockTypeEnum** value that determines what type of locking (concurrency) the provider should use when opening the **Recordset**. The default value is **adLockReadOnly**. |
| *nOptions* | Optional. A Long value that indicates how the provider should evaluate the **Source** argument if it represents something other than a **Command** object, or that the **Recordset** should be restored from a file where it was previously saved. Can be one or more **CommandTypeEnum** or **ExecuteOptionEnum** values, which can be combined with a bitwise **AND** operator.<br>**Note**: If you open a **Recordset** from a **Stream** containing a persisted **Recordset**, using an **ExecuteOptionEnum** value of **adAsyncFetchNonBlocking** will not have an effect; the fetch will be synchronous and blocking.<br>The **ExecuteOptionEnum** values of **adExecuteNoRecords** or **adExecuteStream** should not be used with Open. |

#### CursorTypeEnum

Specifies the type of cursor used in a **Recordset** object.

| Constant   | Description |
| ---------- | ----------- |
| **adOpenDynamic** | Uses a dynamic cursor. Additions, changes, and deletions by other users are visible, and all types of movement through the **Recordset** are allowed, except for bookmarks, if the provider doesn't support them. |
| **adOpenForwardOnly** | Default. Uses a forward-only cursor. Identical to a static cursor, except that you can only scroll forward through records. This improves performance when you need to make only one pass through a **Recordset**. |
| **adOpenKeyset** | Uses a keyset cursor. Like a dynamic cursor, except that you can't see records that other users add, although records that other users delete are inaccessible from your Recordset. Data changes by other users are still visible. |
| **adOpenStatic** | Uses a static cursor, which is a static copy of a set of records that you can use to find data or generate reports. Additions, changes, or deletions by other users are not visible. |
| **adOpenUnspecified** | Does not specify the type of cursor. |

#### CommandTypeEnum

Specifies how a command argument should be interpreted.

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

S_OK or an HRESULT code.

#### Remarks

The default cursor for an ADO **Recordset** is a forward-only, read-only cursor located on the server.

Using the **Open** method on a **Recordset** object opens a cursor that represents records from a base table, the results of a query, or a previously saved **Recordset**.

Use the optional **Source** argument to specify a data source using one of the following: a **Command** object variable, an SQL statement, a stored procedure, a table name, a URL, or a complete file path name. If **Source** is a file path name, it can be a full path ("c:\dir\file.rst"), a relative path ("..\file.rst"), or a URL ("http://files/file.rst").

It is not a good idea to use the **Source** argument of the **Open** method to perform an action query that doesn't return records because there is no easy way to determine whether the call succeeded. The **Recordset** returned by such a query will be closed. Call the **Execute** method of a **Command** object or the **Execute** method of a **Connection** object instead to perform a query that, such as a SQL INSERT statement, that doesn't return records.

The **ActiveConnection** argument corresponds to the **ActiveConnection** property and specifies in which connection to open the **Recordset** object. If you pass a connection definition for this argument, ADO opens a new connection using the specified parameters. After opening the **Recordset** with a client-side cursor (**CursorLocation = adUseClient**), you can change the value of this property to send updates to another provider. Or you can set this property to NULL to disconnect the **Recordset** from any provider. Changing **ActiveConnection** for a server-side cursor generates an error, however.

For the other arguments that correspond directly to properties of a **Recordset** object (**Source**, **CursorType**, and **LockType**), the relationship of the arguments to the properties is as follows:

* The property is read/write before the **Recordset** object is opened.

* The property settings are used unless you pass the corresponding arguments when executing the **Open** method. If you pass an argument, it overrides the corresponding property setting, and the property setting is updated with the argument value.

* After you open the **Recordset** object, these properties become read-only.

**Note**: The **ActiveConnection** property is read only for **Recordset** objects whose **Source** property is set to a valid **Command** object, even if the **Recordset** object isn't open.

If you pass a **Command** object in the **Source** argument and also pass an **ActiveConnection** argument, an error occurs. The **ActiveConnection** property of the **Command** object must already be set to a valid **Connection** object or connection string.

If you pass something other than a **Command** object in the **Source** argument, you can use the **Options** argument to optimize evaluation of the **Source** argument. If the **Options** argument is not defined, you may experience diminished performance because ADO must make calls to the provider to determine if the argument is an SQL statement, a stored procedure, a URL, or a table name. If you know what **Source** type you're using, setting the **Options** argument instructs ADO to jump directly to the relevant code. If the **Options** argument does not match the **Source** type, an error occurs.

If you pass a **Stream** object in the **Source** argument, you should not pass information into the other arguments. Doing so will generate an error. The **ActiveConnection** information is not retained when a **Recordset** is opened from a **Stream**.

The default for the **Options** argument is **adCmdFile** if no connection is associated with the **Recordset**. This will typically be the case for persistently stored **Recordset** objects.

If the data source returns no records, the provider sets both the **BOF** and **EOF** properties to True, and the current record position is undefined. You can still add new data to this empty **Recordset** object if the cursor type allows it.

When you have concluded your operations over an open **Recordset** object, use the **Close** method to free any associated system resources. Closing an object does not remove it from memory; you can change its property settings and use the **Open** method to open it again later. To completely eliminate an object from memory, cal the **Release** method of the interface.

Before the **ActiveConnection** property is set, call **Open** with no operands to create an instance of a **Recordset** created by appending fields to the **Recordset** **Fields** collection.

If you have set the **CursorLocation** property to **adUseClient**, you can retrieve rows asynchronously in one of two ways. The recommended method is to set **Options** to **adAsyncFetch**. Alternatively, you can use the "Asynchronous Rowset Processing" dynamic property in the **Properties** collection, but related retrieved events can be lost if you do not set the **Options** parameter to **adAsyncFetch**.

**Note**: Background fetching in the MS Remote provider is supported only through the **Open** method's **Options** parameter.

**Note**: URLs using the http scheme will automatically invoke the Microsoft OLE DB Provider for Internet Publishing.

Certain combinations of **CommandTypeEnum** and **ExecuteOptionEnum** values are not valid. For information about which options cannot be combined, see the topics for the **ExecuteOptionEnum**, and **CommandTypeEnum**.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Open the recordset
DIM pRecordset AS CAdoRecordset
DIM cvSource AS CVAR = "SELECT TOP 20 * FROM Authors ORDER BY Author"
DIM hr AS HRESULT = pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

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

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the recordset
DIM pRecordset AS CAdoRecordset
DIM cbsSource AS CBSTR = "SELECT TOP 20 * FROM Authors ORDER BY Author"
DIM cvConStr AS CVAR = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"
DIM hr AS HRESULT = pRecordset.Open(cbsSource, cvConStr, adOpenKeyset, adLockOptimistic, adCmdText)

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

# <a name="PageCount"></a>PageCount

Indicates how many pages of data the **Recordset** object contains.

```
PROPERTY PageCount () AS LONG
```

#### Remarks

Use the **PageCount** property to determine how many pages of data are in the **Recordset** object. Pages are groups of records whose size equals the **PageSize** property setting. Even if the last page is incomplete because there are fewer records than the **PageSize** value, it counts as an additional page in the **PageCount** value. If the **Recordset** object does not support this property, the value will be -1 to indicate that the **PageCount** is indeterminable.

See the **PageSize** and **AbsolutePage** properties for more on page functionality.

# <a name="PageSize"></a>PageSize

Indicates how many records constitute one page in the **Recordset**.

Sets or returns a Long value that indicates how many records are on a page. The default is 10.

```
PROPERTY PageSize () AS LONG
PROPERTY PageSize (BYVAL nPageSize AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nPageSize* | Number of records on a page. The default is 10. |

#### Return value

The number of records on a page.

#### Remarks

Use the **PageSize** property to determine how many records make up a logical page of data. Establishing a page size allows you to use the **AbsolutePage** property to move to the first record of a particular page. This is useful in Web-server scenarios when you want to allow the user to page through data, viewing a certain number of records at a time.

This property can be set at any time, and its value will be used for calculating the location of the first record of a particular page.

# <a name="Properties"></a>Properties

Returns a reference to the **Properties** collection.

```
PROPERTY Properties () AS Afx_ADOProperties PTR
```

#### Return value

An **Afx_ADOProperties** object reference.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Create an instance of the CAdoProperties class
' // with a reference to the Peoperties collection.
DIM pProperties AS CAdoProperties = pConnection.Properties
PRINT "Number of properties: "; pProperties.Count

' // Create an instance of the CAdoProperty class
' // with a reference to a Property object.
DIM pProperty AS CAdoProperty = pProperties.Item("DBMS Version")

' // Print the value of the property
PRINT "DBMS version : "; pProperty.Value.ToStr
```

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Open the recordset
DIM pRecordset AS CAdoRecordset
DIM cbsSource AS CBSTR = "SELECT * FROM Publishers ORDER BY PubID"
DIM hr AS HRESULT = pRecordset.Open(cbsSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

' // Parse the Properties collection
DIM pProperties AS CAdoProperties = pRecordset.Properties
DIM nCount AS LONG = pProperties.Count
FOR i AS LONG = 0 TO nCount - 1
   DIM pProperty AS CAdoProperty = pProperties.Item(i)
   PRINT "Property name: "; pProperty.Name; " - Attributes: "; WSTR(pProperty.Attributes)
NEXT
```

Alternate version using a compound syntax:

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Retrieve the "DBMS Version" property
PRINT CAdoProperty(CAdoProperties(pConnection.Properties).Item("DBMS Version")).Value
```

# <a name="RecordCount"></a>RecordCount

Indicates the number of records in a **Recordset** object.

```
PROPERTY RecordCount () AS LONG
```

#### Return value

The number of records in the **Recordset**.

#### Remarks

Use the **RecordCount** property to find out how many records are in a **Recordset** object. The property returns -1 when ADO cannot determine the number of records or if the provider or cursor type does not support **RecordCount**. Reading the **RecordCount** property on a closed **Recordset** causes an error.

If the **Recordset** object supports approximate positioning or bookmarks—that is, **Supports**(*adApproxPosition*) or **Supports**(*adBookmark*), respectively, return True—this value will be the exact number of records in the **Recordset**, regardless of whether it has been fully populated. If the **Recordset** object does not support approximate positioning, this property may be a significant drain on resources because all records will have to be retrieved and counted to return an accurate **RecordCount** value.

**Note**: In ADO versions 2.8 and earlier, the SQLOLEDB provider fetches all records when a server-side cursor is used despite the fact that it returns True for both **Supports**(*adApproxPosition*) and **Supports**(*adBookmark*).

The cursor type of the **Recordset** object affects whether the number of records can be determined. The **RecordCount** property will return -1 for a forward-only cursor; the actual count for a static or keyset cursor; and either -1 or the actual count for a dynamic cursor, depending on the data source.

#### Problem

**RecordCount May Return -1**

The number of records in a dynamic cursor may change. Forward only cursors do not return a **RecordCount**.

Use either **adOpenKeyset** or **adOpenStatic** as the **CursorType** for server side cursors or use a client side cursor. Client side cursors use only **adOpenStatic** for **CursorTypes** regardless of which **CursorType* you select.

This behavior is by design.

# <a name="Requery"></a>Requery

Updates the data in a **Recordset** object by re-executing the query on which the object is based.

```
FUNCTION Requery (BYVAL Options AS LONG = adOpenUnspecified) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *Options* | Optional. A bitmask that contains **ExecuteOptionEnum** and **CommndTypeEnum** values affecting this operation. **Note**: If *Options* is set to **adAsyncExecute**, this operation will execute asynchronously and a **RecordsetChangeComplete** event will be issued when it concludes. The **ExecuteOptionEnum** values of **adExecuteNoRecords** or **adExecuteStream** should not be used with **Requery**. |

#### CommandTypeEnum

Specifies how a command argument should be interpreted.

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
| *adAsyncFetch* | Indicates that the remaining rows after the initial quantity specified in the **CacheSize** property should be retrieved asynchronously. |
| *adAsyncFetchNonBlocking* | Indicates that the main thread never blocks while retrieving. If the requested row has not been retrieved, the current row automatically moves to the end of the file. If you open a **Recordset** from a **Stream** containing a persistently stored **Recordset**, **adAsyncFetchNonBlocking** will not have an effect; the operation will be synchronous and blocking. **adAsynchFetchNonBlocking** has no effect when the **CmdTableDirect** option is used to open the **Recordset**. |
| *adExecuteNoRecords* | Indicates that the command text is a command or stored procedure that does not return rows (for example, a command that only inserts data). If any rows are retrieved, they are discarded and not returned. **adExecuteNoRecords** can only be passed as an optional parameter to the Command or **Connection** **Execute** method. |
| *adExecuteStream* | Indicates that the results of a command execution should be returned as a stream. **adExecuteStream** can only be passed as an optional parameter to the **Command** **Execute** method. |
| *adExecuteRecord* | Indicates that the **CommandText** is a command or stored procedure that returns a single row which should be returned as a **Record** object. |
| *adOptionUnspecified* | Indicates that the command is unspecified. |

#### Return value

S_OK or an HRESULT code.

### Remarks

Use the **Requery** method to refresh the entire contents of a Recordset object from the data source by reissuing the original command and retrieving the data a second time. Calling this method is equivalent to calling the **Close** and **Open** methods in succession. If you are editing the current record or adding a new record, an error occurs.

While the **Recordset** object is open, the properties that define the nature of the cursor (**CursorType**, **LockType**, **MaxRecords**, and so forth) are read-only. Thus, the **Requery** method can only refresh the current cursor. To change any of the cursor properties and view the results, you must use the **Close** method so that the properties become read/write again. You can then change the property settings and call the **Open** method to reopen the cursor.

# <a name="Resync"></a>Resync

Refreshes the data in the current **Recordset** object from the underlying database.

```
FUCTION Resync (BYVAL AffectRecords AS AffectEnum = adAffectAll, _
   BYVAL ResyncValues AS ResyncEnum = adResyncAllValues) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *AffectRecords* | Optional. An **AffectEnum** value that determines how many records the **Resync** method will affect. The default value is **adAffectAll**. This value is not available with the **Resync** method of the **Fields** collection of a **Record** object. |
| *ResyncValues* | Optional. A **ResyncEnum** value that specifies whether underlying values are overwritten. The default value is **adResyncAllValues**. |

#### AffectEnum

Specifies which records are affected by an operation.

| Constant   | Description |
| ---------- | ----------- |
| **adAffectAll** | If there is not a Filter applied to the **Recordset**, affects all records. If the **Filter** property is set to a string criteria (such as "Author='Smith'"), then the operation affects visible records in the current chapter. If the **Filter** property is set to a member of the **FilterGroupEnum** or an array of Bookmarks, then the operation will affect all rows of the **Recordset**. |
| **adAffectAllChapters** | Affects all records in all sibling chapters of the **Recordset**, including those not visible via any **Filter** that is currently applied. |
| **adAffectCurrent** | Affects only the current record. |
| **adAffectGroup** | Affects only records that satisfy the current **Filter** property setting. You must set the **Filter** property to a **FilterGroupEnum** value or an array of Bookmarks to use this option. |

#### ResyncEnum

Specifies whether underlying values are overwritten by a call to **Resync**.

| Constant   | Description |
| ---------- | ----------- |
| **adResyncAllValues** | Default. Overwrites data, and pending updates are canceled. |
| **adResyncUnderlyingValues** | Does not overwrite data, and pending updates are not canceled. |

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use the **Resync** method to resynchronize records in the current **Recordset** with the underlying database. This is useful if you are using either a static or forward-only cursor, but you want to see any changes in the underlying database.

If you set the **CursorLocation** property to **adUseClient**, **Resync** is only available for non-read-only **Recordset** objects.

Unlike the **Requery** method, the **Resync** method does not re-execute the **Recordset** object's underlying command. New records in the underlying database will not be visible.

If the attempt to resynchronize fails because of a conflict with the underlying data (for example, a record has been deleted by another user), the provider returns warnings to the **Errors** collection and a run-time error occurs. Use the **Filter** property (**adFilterConflictingRecords**) and the **Status** property to locate records with conflicts.

If the Unique Table and Resync Command dynamic properties are set, and the **Recordset** is the result of executing a JOIN operation on multiple tables, then the **Resync** method will execute the command given in the **Resync** **Command** property only on the table named in the Unique Table property.

# <a name="Save"></a>Save

Saves the **Recordset** in a file or **Stream** object.

```
FUNCTION Save (BYREF Destination AS CVAR, _
   BYVAL PersistFormat AS PersistFormatEnum = adPersistADTG) AS HRESULT
FUNCTON Save (BYREF pDestination AS Afx_AdoStream PTR, _
   BYVAL PersistFormat AS PersistFormatEnum = adPersistADTG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *Destination* | Optional. The complete path name of the file where the **Recordset** is to be saved. |
| *pDestination* | Optional. A reference to a **Stream** object. |
| *PersistFormat* | Optional. A **PersistFormatEnum** value that specifies the format in which the **Recordset** is to be saved (XML or ADTG). The default value is **adPersistADTG**. |

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

The **Save** method can only be invoked on an open **Recordset**. Use the **Open** method to later restore the **Recordset** from *Destination*.

If the **Filter** property is in effect for the **Recordset**, then only the rows accessible under the filter are saved. If the **Recordset** is hierarchical, then the current child **Recordset** and its children are saved, including the parent **Recordset**. If the **Save** method of a child **Recordset** is called, the child and all its children are saved, but the parent is not.

The first time you save the **Recordset**, it is optional to specify *Destination*. If you omit *Destination*, a new file will be created with a name set to the value of the **Source** property of the **Recordset**.

Omit *Destination* when you subsequently call **Save** after the first save, or a run-time error will occur. If you subsequently call **Save** with a new *Destination*, the **Recordset** is saved to the new destination. However, the new destination and the original destination will both be open.

**Save** does not close the **Recordset** or *Destination*, so you can continue to work with the **Recordset** and save your most recent changes. *Destination* remains open until the Recordset is closed.

For reasons of security, the **Save** method permits only the use of low and custom security settings from a script executed by Microsoft Internet Explorer.

If the **Save** method is called while an asynchronous **Recordset** fetch, execute, or update operation is in progress, then **Save** waits until the asynchronous operation is complete.

Records are saved beginning with the first row of the **Recordset**. When the **Save** method is finished, the current row position is moved to the first row of the **Recordset**.

For best results, set the **CursorLocation** property to **adUseClient** with **Save**. If your provider does not support all of the functionality necessary to save **Recordset** objects, the Cursor Service will provide that functionality.

When a **Recordset** is persisted with the **CursorLocation** property set to **adUseServer**, the update capability for the **Recordset** is limited. Typically, only single-table updates, insertions, and deletions are allowed (dependant upon provider functionality). The **Resync** method is also unavailable in this configuration.

**Note**: Saving a **Recordset** with **Fields** of type **adVariant**, **adIDispatch**, or **adIUnknown** is not supported by ADO and can cause unpredictable results.

Only Filters in the form of Criteria Strings (e.g. OrderDate > '12/31/1999') affect the contents of a persisted **Recordset**. Filters created with an Array of Bookmarks or using a value from the **FilterGroupEnum** will not affect the contents of the persisted **Recordset**. These rules apply to Recordsets created with either client-side or server-side cursors.

Because the *Destination* parameter can accept any object that supports the OLE DB IStream interface, you can save a **Recordset** directly to the ASP Response object.

You can also save a **Recordset** in XML format to an instance of an MSXML DOM object.

**Note**: Two limitations apply when saving hierarchical Recordsets (data shapes) in XML format. You cannot save into XML if the hierarchical **Recordset** contains pending updates, and you cannot save a parameterized hierarchical **Recordset**.

A **Recordset** saved in XML format is saved using UTF-8 format. When such a file is loaded into an ADO **Stream**, the **Stream** object will not attempt to open a **Recordset** from the stream unless the **Charset** property of the stream is set to the appropriate value for UTF-8 format.

**Example** (save as XML)

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Open the recordset
DIM pRecordset AS CAdoRecordset
DIM cvSource AS CBSTR = "SELECT * FROM Publishers"
DIM hr AS HRESULT = pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

' // Save the recordset as XML
DIM wszFileName AS WSTRING * MAX_PATH = "Publishers.xml"
IF AfxFileExists(wszFileName) THEN KILL wszFileName
IF pRecordset.Save(wszFileName, adPersistXML) = S_OK THEN
   PRINT "Recordset saved"
ELSE
   PRINT "Save failed"
END IF
```

#### Persisting data

Portable computing (for example, using laptops) has generated the need for applications that can run in both a connected and disconnected state. ADO has added support for this by giving the developer the ability to save a client cursor **Recordset** to disk and reload it later.

There are several scenarios in which you could use this type of feature, including the following:

Traveling: When taking the application on the road, it is vital to supply the ability to make changes and add new records that can then be reconnected to the database later and committed.

Infrequently updated lookups: Often in an application, tables are used as lookups—for example, state tax tables. They are infrequently updated and are read-only. Instead of rereading this data from the server each time the application is started, the application can simply load the data from a locally persisted **Recordset**.

In ADO, to save and load Recordsets, use the **Recordset.Save** and **Recordset.Open**(*,,,,adCmdFile*) methods on the ADO **Recordset** object.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Create an scoped instance of the CAdoRecordset class
SCOPE
   DIM pRecordset AS CAdoRecordset
   ' // Set the cursor location
   pRecordset.CursorLocation = adUseClient

   ' // Open the recordset
   DIM cbsSource AS CBSTR = "SELECT * FROM Publishers"
   DIM hr AS HRESULT = pRecordset.Open(cbsSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

   ' // Save the recordset as XML
   DIM wszFileName AS WSTRING * MAX_PATH = "Publishers.dat"
   IF AfxFileExists(wszFileName) THEN KILL wszFileName
   IF pRecordset.Save(wszFileName, adPersistADTG) = S_OK THEN
      PRINT "Recordset saved"
   ELSE
      PRINT "Save failed"
   END IF
END SCOPE

IF AfxFileExists("Publishers.dat") THEN
   ' // Reopen the recordset
   DIM pRecordset AS CAdoRecordset
   DIM cbsSource AS CBSTR = "Publishers.dat"
   DIM hr AS HRESULT = pRecordset.Open(cbsSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdFile)
   IF hr = S_OK THEN
      ' // Parse the recordset
      DO
         ' // While not at the end of the recordset...
         IF pRecordset.EOF THEN EXIT DO
         ' // Get the contents
         DIM cvRes1 AS CVAR = pRecordset.Collect("PubID")
         DIM cvRes2 AS CVAR = pRecordset.Collect("Name")
         DIM cvRes3 AS CVAR = pRecordset.Collect("Company Name")
         PRINT cvRes1 & " " & cvRes2 & " " & cvRes3
         ' // Fetch the next row
         IF pRecordset.MoveNext <> S_OK THEN EXIT DO
      LOOP
   END IF
END IF
```

# <a name="Seek"></a>Seek

Searches the index of a **Recordset** to quickly locate the row that matches the specified values, and changes the current row position to that row.

```
FUNCTION Seek (BYREF KeyValues AS CVAR, BYVAL SeekOption AS SeekEnum = adSeekFirstEQ) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cvKeyValues* | An array of Variant values. An index consists of one or more columns and the array contains a value to compare against each corresponding column. |
| *SeekOption* | Optional. A **SeekEnum** value that specifies the type of comparison to be made between the columns of the index and the corresponding *KeyValues*. |

#### SeekEnum

Specifies the type of Seek to execute.

| Constant   | Description |
| ---------- | ----------- |
| **adSeekFirstEQ** | Seeks the first key equal to *KeyValues*. |
| **adSeekLastEQ** | Seeks the last key equal to *KeyValues*. |
| **adSeekAfterEQ** | Seeks either a key equal to *KeyValues* or just after where that match would have occurred. |
| **adSeekAfter** | Seeks a key just after where a match with *KeyValues* would have occurred. |
| **adSeekBeforeEQ** | Seeks either a key equal to *KeyValues* or just before where that match would have occurred. |
| **adSeekBefore** | Seeks a key just before where a match with *KeyValues* would have occurred. |

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use the **Seek** method in conjunction with the **Index** property if the underlying provider supports indexes on the **Recordset** object. Use the **Supports**(*adSeek*) method to determine whether the underlying provider supports Seek, and the **Supports**(*adIndex*) method to determine whether the provider supports indexes. (For example, the OLE DB Provider for Microsoft Jet supports **Seek** and **Index**.)

If **Seek** does not find the desired row, no error occurs, and the row is positioned at the end of the **Recordset**. Set the **Index** property to the desired index before executing this method.

This method is supported only with server-side cursors. Seek is not supported when the **Recordset** object's **CursorLocation** property value is **adUseClient**.

This method can only be used when the **Recordset** object has been opened with a **CommandTypeEnum** value of **adCmdTableDirect**.

#### Note

The SQL Server 6.5 or 7.0 doesn't support the **Seek** or **Index** methods of the **Recordset**. You can however, use the **Index** and **Seek** method with an Access 2000 database and the OLE DB 4.0 Provider for Jet. 

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Set the cursor location
DIM pRecordset AS CAdoRecordset
pRecordset.CursorLocation = adUseServer

' // Open the recordset
DIM cvSource AS CVAR = "Publishers"
DIM hr AS HRESULT = pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdTableDirect)

' // Set the index
pRecordset.Index = "PrimaryKey"

' // Seek the record 70
pRecordset.Seek 70, 1

' // Parse the recordset
DO
   ' // While not at the end of the recordset...
   IF pRecordset.EOF THEN EXIT DO
   ' // Get the contents
   DIM cvRes1 AS CVAR = pRecordset.Collect("PubID")
   DIM cvRes2 AS CVAR = pRecordset.Collect("Name")
   DIM cvRes3 AS CVAR = pRecordset.Collect("Company Name")
   PRINT cvRes1 & " " & cvRes2 & " " & cvRes3
   ' // Fetch the next row
   IF pRecordset.MoveNext <> S_OK THEN EXIT DO
LOOP
```

# <a name="Sort"></a>Sort

Indicates one or more field names on which the **Recordset** is sorted, and whether each field is sorted in ascending or descending order.

Sets or returns a string value that indicates the field names in the **Recordset** on which to sort.

```
PROPERTY Sort() AS CBSTR
PROPERTY Sort (BYVAL cbsCriteria AS CBSTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsCriteria* | A value that indicates the field names in the **Recordset** on which to sort. Each name is separated by a comma, and is optionally followed by a blank and the keyword, ASC, which sorts the field in ascending order, or DESC, which sorts the field in descending order. By default, if no keyword is specified, the field is sorted in ascending order. |

#### Return value

The sorting criteria.

#### Remarks

This property requires the **CursorLocation** property to be set to **adUseClient**. A temporary index will be created for each field specified in the **Sort** property if an index does not already exist.

The sort operation is efficient because data is not physically rearranged, but is simply accessed in the order specified by the index.

When the value of the **Sort** property is anything other than an empty string, the **Sort** property order takes precedence over the order specified in an ORDER BY clause included in the SQL statement used to open the **Recordset**.

The **Recordset** does not have to be opened before accessing the **Sort** property; it can be set at any time after the **Recordset** object is instantiated.

Setting the **Sort** property to an empty string will reset the rows to their original order and delete temporary indexes. Existing indexes will not be deleted.

Suppose a **Recordset** contains three fields named *firstName*, *middleInitial*, and *lastName*. Set the **Sort** property to the string, "lastName DESC, firstName ASC", which will order the Recordset by last name in descending order, then by first name in ascending order. The middle initial is ignored.

No field can be named "ASC" or "DESC" because those names conflict with the keywords ASC and DESC. Give a field with a conflicting name an alias by using the AS keyword in the query that returns the Recordset.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Set the cursor location
DIM pRecordset AS CAdoRecordset
pRecordset.CursorLocation = adUseClient

' // Open the recordset
DIM cvSource AS CBSTR = "Publishers"
DIM hr AS HRESULT = pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdTableDirect)

' // Set the Sort property
pRecordset.Sort = "City ASC, Name ASC"

' // Parse the recordset
DO
   ' // While not at the end of the recordset...
   IF pRecordset.EOF THEN EXIT DO
   ' // Get the contents of the "City" and "Name" columns
   DIM cvRes1 AS CVAR = pRecordset.Collect("City")
   DIM cvRes2 AS CVAR = pRecordset.Collect("Name")
   PRINT cvRes1 & " " & cvRes2
   ' // Fetch the next row
   IF pRecordset.MoveNext <> S_OK THEN EXIT DO
LOOP
```

# <a name="Source"></a>Source

Indicates the data source for a **Recordset** object.

Sets a string value or **Command** object reference; returns only a string value that indicates the source of the **Recordset**.

```
PROPERTY Source (BYREF cbsConn AS CBSTR)
PROPERTY Source (BYVAL pcmd AS Afx_ADOCommand PTR)
PROPERTY Source () AS CVAR
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsConn* | The source of the **Recordset**. |
| *pcmd* | A **Command** object reference. |

#### Return value

The data source of the **Recordset**.

#### Remarks

Use the **Source** property to specify a data source for a **Recordset** object using one of the following: a **Command** object variable, an SQL statement, a stored procedure, or a table name.

If you set the **Source** property to a **Command** object, the **ActiveConnection** property of the **Recordset** object will inherit the value of the **ActiveConnection** property for the specified **Command** object. However, reading the **Source** property does not return a **Command** object; instead, it returns the **CommandText** property of the **Command** object to which you set the Source property.

If the **Source** property is an SQL statement, a stored procedure, or a table name, you can optimize performance by passing the appropriate **Options** argument with the **Open** method call.

The **Source** property is read/write for closed **Recordset** objects and read-only for open **Recordset** objects.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Create a Recordset object
DIM pRecordset AS CAdoRecordset
' // Set the active connection
pRecordset.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"
' // Set the source
pRecordset.Source = "SELECT TOP 20 * FROM Authors ORDER BY Author"
' // Open the recordset
DIM hr AS HRESULT = pRecordset.Open
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
DIM cvConStr AS CVAR = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"
pConnection.Open cvConStr
' // Set the active connection
pRecordset.ActiveConnection = pConnection
' // Set the source
DIM cvSource AS CVAR = "SELECT TOP 20 * FROM Authors ORDER BY Author"
pRecordset.Source = cvSource
' // Open the recordset
DIM hr AS HRESULT = pRecordset.Open
```

# <a name="State"></a>State

Indicates for whether the state of the **Recordset** object is open or closed.

```
PROPERTY State () AS LONG
```

#### Return value

The current **Recordset** state.

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

# <a name="Status"></a>Status

Indicates the status of the current record with respect to batch updates or other bulk operations.

```
PROPERTY Status () AS LONG
```

#### Return value

The status of the current record.

#### EventStatusEnum

Specifies the current status of the execution of an event.

| Constant   | Description |
| ---------- | ----------- |
| **adStatusCancel** | Requests cancellation of the operation that caused the event to occur. |
| **adStatusCantDeny** | Indicates that the operation cannot request cancellation of the pending operation. |
| **adStatusErrorsOccurred** | Indicates that the operation that caused the event failed due to an error or errors. |
| **adStatusOK** | Indicates that the operation that caused the event was successful. |
| **adStatusUnwantedEvent** | Prevents subsequent notifications before the event method has finished executing. |

#### Remarks

Use the **Status** property to see what changes are pending for records modified during batch updating. You can also use the **Status** property to view the status of records that fail during bulk operations, such as when you call the **Resync**, **UpdateBatch**, or **CancelBatch** methods on a **Recordset** object, or set the **Filter** property on a **Recordset** object to an array of bookmarks. With this property, you can determine how a given record failed and resolve it accordingly.

# <a name="StayInSync"></a>StayInSync

Indicates, in a hierarchical **Recordset** object, whether the reference to the underlying child records (that is, the chapter) changes when the parent row position changes.

```
PROPERTY StayInSync (BYVAL pbStayInSync AS BOOLEAN)
PROPERTY StayInSync () AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *bStayInSync* | Sets a Boolean value. The default value is True. If True, the chapter will be updated if the parent **Recordset** object changes row position; if False, the chapter will continue to refer to data in the previous chapter even though the parent **Recordset** object has changed row position. |

#### Return value

TRUE or FALSE.

#### Remarks

This property applies to hierarchical recordsets, such as those supported by the Microsoft Data Shaping Service for OLE DB, and must be set on the parent Recordset before the child **Recordset** is retrieved. This property simplifies navigating hierarchical recordsets.

# <a name="Supports"></a>Supports

Determines whether a specified **Recordset** object supports a particular type of functionality.

```
FUNCTION Supports (BYVAL CursorOptions AS CursorOptionEnum) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *CursorOptions* | A **CursorOptionEnum** value. |

#### CursorOptionEnum

Specifies what functionality the **Supports** method should test for.

| Constant   | Description |
| ---------- | ----------- |
| **adAddNew** | Supports the **AddNew** method to add new records. |
| **adApproxPosition** | Supports the **AbsolutePosition** and **AbsolutePage** properties. |
| **adBookmark** | Supports the **Bookmark** property to gain access to specific records. |
| **adDelete** | Supports the **Delete_** method to delete records. |
| **adFind** | Supports the Find method to locate a row in a **Recordset**. |
| **adHoldRecords** | Retrieves more records or changes the next position without committing all pending changes. |
| **adIndex** | Supports the **Index** property to name an index. |
| **adMovePrevious** | Supports the **MoveFirst** and **MovePrevious** methods, and **Move** or **GetRows** methods to move the current record position backward without requiring bookmarks. |
| **adNotify** | Indicates that the underlying data provider supports notifications (which determines whether **Recordset** events are supported). |
| **adResync** | Supports the **Resync** method to update the cursor with the data that is visible in the underlying database. |
| **adSeek** | Supports the **Seek** method to locate a row in a **Recordset**. |
| **adUpdate** | Supports the **Update** method to modify existing data. |
| **adUpdateBatch** | Supports batch updating (**UpdateBatch** and **CancelBatch** methods) to transmit groups of changes to the provider. |

#### Return value

True if the object supports the specified type of functionality; False, otherwise.

# <a name="Update"></a>Update

Saves any changes you make to the current row of a **Recordset** object.

```
FUNCTION Update ( _
   BYREF cvFieldList AS CVAR = TYPE(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
   BYREF cvValues AS CVAR = TYPE(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND) _
) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cvFieldList* | Optional. A Variant that represents a single name, or a Variant array that represents names or ordinal positions of the field or fields you wish to modify. |
| *cvValues* | Optional. A Variant that represents a single value, or a Variant array that represents values for the field or fields in the new record. |

#### Return value

S_OK (0) or an HRESULT code.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Create a Connection object
DIM pConnection AS CAdoConnection
' // Create a Recordset object
DIM pRecordset AS CAdoRecordset

' // Open the connection
DIM cvConStr AS CVAR = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"
pConnection.Open cvConStr

' // Open the recordset
DIM cvSource AS CVAR = "Publishers"
DIM hr AS HRESULT = pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdTableDirect)

' // Add a new record
pRecordset.AddNew
   pRecordset.Collect("PubID") = CLNG(10000)
   pRecordset.Collect("Name") = "Wile E. Coyote"
   pRecordset.Collect("Company Name") = "Warner Brothers Studios"
   pRecordset.Collect("Address") = "4000 Warner Boulevard"
   pRecordset.Collect("City") = "Burbank, CA. 91522"
pRecordset.Update
```

# <a name="UpdateBatch"></a>UpdateBatch

Writes all pending batch updates to disk.

```
FUNCTION UpdateBatch (BYVAL AffectRecords AS AffectEnum = adAffectAll) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *AffectRecords* | Optional. An **AffectEnum** value that indicates how many records the **UpdateBatch** method will affect. |

#### AffectEnum

Specifies which records are affected by an operation.

| Constant   | Description |
| ---------- | ----------- |
| **adAffectAll** | If there is not a Filter applied to the **Recordset**, affects all records. If the **Filter** property is set to a string criteria (such as "Author='Smith'"), then the operation affects visible records in the current chapter. If the **Filter** property is set to a member of the **FilterGroupEnum** or an array of Bookmarks, then the operation will affect all rows of the **Recordset**. |
| **adAffectAllChapters** | Affects all records in all sibling chapters of the **Recordset**, including those not visible via any **Filter** that is currently applied. |
| **adAffectCurrent** | Affects only the current record. |
| **adAffectGroup** | Affects only records that satisfy the current **Filter** property setting. You must set the **Filter** property to a **FilterGroupEnum** value or an array of Bookmarks to use this option. |

#### Return value

S_OK (0) or an HRESULT code.
