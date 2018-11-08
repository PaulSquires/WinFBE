# CADOFields Class

Contains all the **Field** objects of a **Recordset** or **Record** object.

**Include file**: CADOFields.inc (include CADODB.inc)

A **Recordset** object has a **Fields** collection made up of **Field** objects. Each **Field** object corresponds to a column in the **Recordset**. You can populate the **Fields** collection before opening the **Recordset** by calling the **Refresh** method on the collection.

See the **Field** object topic for a more detailed explanation of how to use **Field** objects.

The **Fields** collection has an **Append** method, which provisionally creates and adds a Field object to the collection, and an Update method, which finalizes any additions or deletions.

A **Record** object has two special fields that can be indexed with **FieldEnum** constants. One constant accesses a field containing the default stream for the **Record**, and the other accesses a field containing the absolute URL string for the **Record**.

Certain providers (for example, the Microsoft OLE DB Provider for Internet Publishing) may populate the **Fields** collection with a subset of available fields for the **Record** or **Recordset**. Other fields will not be added to the collection until they are first referenced by name or indexed by your code.

If you attempt to reference a nonexistent field by name, a new **Field** object will be appended to the **Fields** collection with a **Status** of **adFieldPendingInsert**. When you call Update, ADO will create a new field in your data source if allowed by your provider.

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [Append](#Append) | Appends an object to a collection. |
| [Attach](#Attach1) | Attaches a reference to a **Fields** object to the class. |
| [CancelUpdate](#CancelUpdate) | Cancels any changes made to the **Fields** collection of a **Record** object before calling the **Update** method. |
| [Count](#Count) | Retrieves the number of objects in the **Fields** collection. |
| [Delete_](#Delete_) | Deletes an object from the **Fields** collection. |
| [Item](#Item) | Indicates a specific member of the **Fields** collection, by name or ordinal number. |
| [Refresh](#Refresh) | Refreshes the contents of the **Fields** collection. |
| [Resync](#Resync) | Resynchronizes the contents of the **Fields** collection. |
| [Update](#Update) | Saves any changes you make to the current **Fields** collection of a **Record** object. |

# CADOField Class

Represents a column of data with a common data type.

**Include file**: CADOField.inc (include CADODB.inc)

Each **Field** object corresponds to a column in the **Recordset**. You use the **Value** property of **Field** objects to set or return data for the current record. Depending on the functionality the provider exposes, some collections, methods, or properties of a **Field** object may not be available.

With the collections, methods, and properties of a **Field** object, you can do the following:

* Return the name of a field with the **Name** property.
* View or change the data in the field with the **Value** property. **Value** is the default property of the **Field** object.
* Return the basic characteristics of a field with the **Type_**, **Precision**, and **NumericScale** properties.
* Return the declared size of a field with the **DefinedSize** property.
* Return the actual size of the data in a given field with the **ActualSize** property.
* Determine what types of functionality are supported for a given field with the **Attributes** property and **Properties** collection.
* Manipulate the values of fields containing long binary or long character data with the **AppendChunk** and **GetChunk** methods.
* If the provider supports batch updates, resolve discrepancies in field values during batch updating with the **OriginalValue** and **UnderlyingValue** properties.

All of the metadata properties (**Name**, **Type_**, **DefinedSize**, **Precision**, and **NumericScale**) are available before opening the **Field** object's **Recordset**. Setting them at that time is useful for dynamically constructing forms.

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [ActualSize](#ActualSize) | Indicates the actual length of a field's value. |
| [AppendChunk](#AppendChunk) | Appends data to a large text or binary data **Field**. |
| [Attach](#Attach2) | Attaches a reference to a **Field** object to the class. |
| [Attributes](#Attributes) | Indicates one or more characteristics of a field. |
| [DataFormat](#DataFormat) | Links the current **Field** object to a data-bound control. |
| [DefinedSize](#DefinedSize) | Indicates the data capacity of a field. |
| [GetChunk](#GetChunk) | Returns all, or a portion, of the contents of a large text or binary data field. |
| [Name](#Name) | Returns the name of the field. |
| [NumericScale](#NumericScale) | Sets or returns a Byte value that indicates the number of decimal places to which numeric values will be resolved. |
| [OriginalValue](#OriginalValue) | Indicates the value of a field that existed in the record before any changes were made. |
| [Precision](#Precision) | Sets or returns a Byte value that indicates the maximum number of digits used to represent values. |
| [Properties](#Properties) | Returns a reference to the Properties collection. |
| [Status](#Status) | Indicates the status of a field. |
| [Type_](#Type_) | Sets or returns a **DataTypeEnum** value. |
| [UnderlyingValue](#UnderlyingValue) | Indicates a field's current value in the database. |
| [Value](#Value) | Sets or returns a Variant value that indicates the value of the object |

# <a name="Append"></a>Append (CADOFields)

Appends an object to a collection. A new **Field** object may be created before it is appended to the collection.

```
FUNCTION Append (BYREF cbsName AS CBSTR, BYVAL nType AS DataTypeEnum, _
   BYVAL DefinedSize AS ADO_LONGPTR = 0, BYVAL Attrib AS FieldAttributeEnum = 0, _
   BYREF cvFieldValue AS CVAR = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsName* | An string value that contains the name of the new **Field** object, and must not be the same name as any other object in fields. |
| *nType* | A **DataTypeEnum** value, whose default value is **adEmpty**, that specifies the data type of the new field. The following data types are not supported by ADO, and should not be used when appending new fields to a Recordset: **adIDispatch**, **adIUnknown**, **adVariant**. |
| *DefinedSize* | Optional. An ADO_LONGPTR value that represents the defined size, in characters or bytes, of the new field. The default value for this parameter is derived from Type. Fields with a **DefinedSize** greater than 255 bytes, and treated as variable length columns. (The default **DefinedSize** is unspecified.) |
| *Attrib* | Optional. A **FieldAttributeEnum** value, whose default value is **adFldDefault**, that specifies attributes for the new field. If this value is not specified, the field will contain attributes derived from **Type_**. |
| *FieldValue* | Optional. A Vaiant that represents the value for the new field. If not specified, then the field is appended with a null value. |

#### Return value

S_OK (0) or an HESULT code.

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

The **FieldValue** parameter is only valid when adding a **Field** object to a **Record** object, not to a **Recordset** object. With a **Record** object, you may append fields and provide values at the same time. With a **Recordset** object, you must create fields while the **Recordset** is closed, then open the **Recordset** and assign values to the fields.

**Notes**: For new **Field** objects that have been appended to the **Fields** collection of a **Record** object, the Value property must be set before any other **Field** properties can be specified. First, a specific value for the **Value** property must have been assigned and **Update** on the **Fields** collection called. Then, other properties such as Type or **Attributes** can be accessed.

**Field** objects of the following data types (**DataTypeEnum**) cannot be appended to the **Fields** collection and will cause an error to occur: **adArray**, **adChapter**, **adEmpty**, **adPropVariant**, and **adUserDefined**. Also, the following data types are not supported by ADO: **adIDispatch**, **adIUnknown**, and **adIVariant**. For these types, no error will occur when appended, but usage can produce unpredictable results including memory leaks.

# <a name="Attach1"></a>Attach (CADOFields)

Attaches a reference to an ADO **Fields** object to the class.

```
SUB Attach (BYVAL pFields AS Afx_ADOFields PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pFields* | A reference to an ADO **Fields** object. |
| *fAddRef* | TRUE = increase the reference count; FALSE = don't increase the reference count. |

# <a name="CancelUpdate"></a>CancelUpdate (CADOFields)

Cancels any changes made to the **Fields** collection of a **Record** object before calling the **Update** method.

```
FUNCTION CancelUpdate () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

The **CancelUpdate** method cancels any pending insertions or deletions of Field objects, and cancels pending updates of existing fields and restores them to their original values. The **Status** property of all fields in the Fields collection is set to **adFieldOK**.

# <a name="Count"></a>Count (CADOFields)

Retrieves the number of objects in the **Fields** collection.

```
PROPERTY Count () AS LONG
```

#### Return value

The number of objects in the **Fields** collection.

#### Remarks

Because numbering for members of a collection begins with zero, you should always code loops starting with the zero member and ending with the value of the **Count** property minus 1.

If the **Count** property is zero, there are no objects in the collection.

# <a name="CoDelete_unt"></a>Delete_ (CADOFields)

Deletes an object from the **Fields** collection.

```
FUNCTION Delete_ (BYREF cvIndex AS CVAR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cvIndex* | A Variant that designates the **Field** object to delete. This parameter can be the name of the **Field** object or the ordinal position of the **Field** object itself. |

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Calling the **Delete_** method on an open **Recordset** causes a run-time error.

# <a name="Item"></a>Item (CADOFields)

Indicates a specific member of the **Fields** collection, by name or ordinal number.

```
PROPERTY Item (BYREF cvIndex AS CVAR) AS Afx_ADOField PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *cvIndex* | A Variant expression that evaluates either to the name or to the ordinal number of an object in a collection. |

#### Return value

An ADO **Fields** object reference.

#### Remarks

If **Item** cannot find an object in the collection corresponding to the *Index* argument, an error occurs.

# <a name="Refresh"></a>Refresh (CADOFields)

Refreshes the contents of the **Fields** collection.

```
FUNCTION Refresh () AS HRESULT
```
#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Using the **Refresh** method on the **Fields** collection has no visible effect. To retrieve changes from the underlying database structure, you must use either the **Requery** method or, if the **Recordset** object does not support bookmarks, the **MoveFirst** method.

# <a name="Resync"></a>Resync (CADOFields)

Resynchronizes the contents of the **Fields** collection.

```
FUNCTION Resync (BYVAL ResyncValues AS ResyncEnum = adResyncAllValues) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *ResyncValues* | Optional. A **ResyncEnum** value that specifies whether underlying values are overwritten. The default value is **adResyncAllValues**. |

#### ResyncEnum

Specifies whether underlying values are overwritten by a call to **Resync**.

| Constant   | Description |
| ---------- | ----------- |
| **adResyncAllValues** | Default. Overwrites data, and pending updates are canceled. |
| **adResyncUnderlyingValues** | Does not overwrite data, and pending updates are not canceled. |

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use the **Resync** method to resynchronize the values of the **Fields** collection of a **Record** object with the underlying data source. The **Count** property is not affected by this method.

If **ResyncValues** is set to **adResyncAllValues** (the default value), then the **UnderlyingValue**, **Value**, and **OriginalValue** properties of **Field** objects in the collection are synchronized. If **ResyncValues** is set to **adResyncUnderlyingValues**, only the **UnderlyingValue** property is synchronized.

The value of the **Status** property for each **Field** object at the time of the call also affects the behavior of **Resync**. For **Field** objects with **Status** values of **adFieldPendingUnknown** or **adFieldPendingInsert**, **Resync** has no effect. For **Status** values of **adFieldPendingChange** or **adFieldPendingDelete**, **Resync** synchronizes data values for fields that still exist at the data source.

**Resync** will not modify **Status** values of **Field** objects unless an error occurs when **Resync** is called. For example, if the field no longer exists, the provider will return an appropriate Status value for the **Field** object, such as **adFieldDoesNotExist**. Returned **Status** values may be logically combined within the value of the **Status** property.

# <a name="Update"></a>Update (CADOFields)

Saves any changes you make to the current **Fields** collection of a **Record** object.

```
FUNCTION Update () AS HRESULT
```

#### Return value

S_OK or an HRESULT code.

#### Remarks

The **Update** method finalizes additions, deletions, and updates to fields in the **Fields** collection of a **Record** object.

For example, fields deleted with the **Delete_** method are marked for deletion immediately but remain in the collection. The **Update** method must be called to actually delete these fields from the provider's collection.

# <a name="ActualSize"></a>ActualSize (CADOField)

Indicates the actual length of a field's value. Some providers may allow this property to be set to reserve space for BLOB data, in which case the default value is 0.

```
PROPERTY ActualSize () AS ADO_LONGPTR
```

#### Return value

The actual length of the field's value.

#### Remarks

Use the **ActualSize** property to return the actual length of a field value. For all fields, the **ActualSize** property is read-only. If ADO cannot determine the length of the field value, the **ActualSize** property returns **adUnknown**.

The **ActualSize** and **DefinedSize** properties are different. A field with a declared type of **adVarChar** and a maximum length of 50 characters returns a **DefinedSize** property value of 50, but the **ActualSize** property value it returns is the length of the data stored in the field for the current record. **Fields** with a **DefinedSize** greater than 255 bytes are treated as variable length columns.

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

' // Get a reference to the Fields collection
DIM pFields AS CAdoFields = pRecordset.Fields

' // Parse the recordset
DO
   ' // While not at the end of the recordset...
   IF pRecordset.EOF THEN EXIT DO
   ' // Get the contents of the fields
   DIM pField AS CAdoField
   pField.Attach(pFields.Item("Name"))
   DIM cvRes AS CVAR = pField.Value
   PRINT "Name: "; cvRes; " - "; WSTR(pField.ActualSize); " - "; WSTR(pField.DefinedSize)
   ' // Fetch the next row
   IF pRecordset.MoveNext <> S_OK THEN EXIT DO
LOOP
```

# <a name="AppendChunk"></a>AppendChunk (CADOField)

Appends data to a large text or binary data **Field*.

```
FUNCTION AppendChunk (BYREF cvData AS CVAR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cvData* | A Variant that contains the data to append to the object. |

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Use the **AppendChunk** method on a Field object to fill it with long binary or character data. In situations where system memory is limited, you can use the **AppendChunk** method to manipulate long values in portions rather than in their entirety.

If the **adFldLong** bit in the **Attributes** property of a **Field** object is set to true, you can use the **AppendChunk** method for that field.

The first **AppendChunk** call on a **Field** object writes data to the field, overwriting any existing data. Subsequent **AppendChunk** calls add to existing data. If you are appending data to one field and then you set or read the value of another field in the current record, ADO assumes that you are finished appending data to the first field. If you call the **AppendChunk** method on the first field again, ADO interprets the call as a new **AppendChunk** operation and overwrites the existing data. Accessing fields in other **Recordset** objects that are not clones of the first **Recordset** object will not disrupt **AppendChunk** operations.

If there is no current record when you call **AppendChunk** on a **Field** object, an error occurs.

**Note**: The **AppendChunk** method does not operate on Field objects of a **Record** object. It does not perform any operation and will produce a run-time error.

# <a name="Attach2"></a>Attach (CADOField)

Attaches a reference to an ADO **Field** object to the class.

```
SUB Attach (BYVAL pField AS Afx_ADOField PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pField* | A reference to an ADO **Field** object. |
| *fAddRef* | TRUE = increase the reference count; FALSE = don't increase the reference count. |

# <a name="Attributes"></a>Attributes (CADOField)

Indicates one or more characteristics of a field.

```
PROPERTY Attributes () AS LONG
PROPERTY Attributes (BYVAL lAttr AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *lAttr* | Can be the sum of one or more **FieldAttributeEnum** values. It is normally read-only, However, for new **Field** objects that have been appended to the **Fields** collection of a **Record**, **Attributes** is read/write only after the **Value** property for the **Field** has been specified and the new **Field** has been successfully added by the data provider by calling the **Update** method of the **Fields** collection. |

#### Return value

One or more **FieldAttributeEnum** values.

#### FieldAttributeEnum

Specifies one or more attributes of a **Field** object.

| Parameter  | Description |
| ---------- | ----------- |
| **adFldCacheDeferred** | Indicates that the provider caches field values and that subsequent reads are done from the cache. |
| **adFldFixed** | Indicates that the field contains fixed-length data. |
| **adFldIsChapter** | Indicates that the field contains a chapter value, which specifies a specific child recordset related to this parent field. Typically chapter fields are used with data shaping or filters. |
| **adFldIsCollection** | Indicates that the field specifies that the resource represented by the record is a collection of other resources, such as a folder, rather than a simple resource, such as a text file. |
| **adFldIsDefaultStream** | Indicates that the field contains the default stream for the resource represented by the record. For example, the default stream can be the HTML content of a root folder on a Web site, which is automatically served when the root URL is specified. |
| **adFldIsNullable** | Indicates that the field accepts null values. |
| **adFldIsRowURL** | Indicates that the field contains the URL that names the resource from the data store represented by the record.  |
| **adFldLong** | Indicates that the field is a long binary field. Also indicates that you can use the **AppendChunk** and **GetChunk** methods. |
| **adFldMayBeNull** | Indicates that you can read null values from the field. |
| **adFldMayDefer** | Indicates that the field is deferred—that is, the field values are not retrieved from the data source with the whole record, but only when you explicitly access them. |
| **adFldNegativeScale** | Indicates that the field represents a numeric value from a column that supports negative scale values. The scale is specified by the **NumericScale** property. |
| **adFldRowID** | Indicates that the field contains a persistent row identifier that cannot be written to and has no meaningful value except to identify the row (such as a record number, unique identifier, and so forth). |
| **adFldRowVersion** | Indicates that the field contains some kind of time or date stamp used to track updates. |
| **adFldUnknownUpdatable** | Indicates that the provider cannot determine if you can write to the field. |
| **adFldUnspecified** | Indicates that the provider does not specify the field attributes. |
| **adFldUpdatable** | Indicates that you can write to the field. |

#### Remarks

When you set multiple attributes, you can sum the appropriate constants. If you set the property value to a sum including incompatible constants, an error occurs.

#### Remote Data Service Usage

This property is not available on a client-side **Connection** object.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Open the recordset
DIM pRecordset AS CAdoRecordset
DIM cvSource AS CVAR = "SELECT * FROM Publishers ORDER BY PubID"
DIM hr AS HRESULT = pRecordset.Open(cbsSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

' // Get a reference to the Fields collection
DIM pFields AS CAdoFields = pRecordset.Fields
DIM nCount AS LONG = pFields.Count
IF nCount THEN
   PRINT
   PRINT "Nullable fields:"
   PRINT "================"
   PRINT
   FOR i AS LONG = 0 TO nCount - 1
      DIM pField AS CAdoField
      pField.Attach(pFields.Item(i))
      ' // Get the attributes of the field
      DIM lAttr AS LONG = pField.Attributes
      ' // Display fields that are nullable
      IF (lAttr AND adFldIsNullable) = adFldIsNullable THEN
         PRINT "Field name: "; pField.Name
      END IF
   NEXT
END IF
```

# <a name="DataFormat"></a>DataFormat (CADOField)

Links the current **Field** object to a data-bound control.

```
PROPERTY DataFormat () AS IUnknown PTR
PROPERTY DataFormat (BYVAL piDF AS IUnknown PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *piDF* | A reference to a **StdDataFormat** object. |

#### Return value

A reference to a **StdDataFormat** object.

#### Remarks

The **DataFormat** property is both read- and write-enabled. It accepts and returns a **StdDataFormat** object that is used to attach a bound object.

# <a name="DefinedSize"></a>DefinedSize (CADOField)

Indicates the data capacity of a field.

```
PROPERTY DefinedSize () AS LONG
PROPERTY DefinedSize (BYVAL lSize AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *lSize* | The data capacity of the field's value. |

#### Return value

The actual length of the field's value.

#### Remarks

Use the **DefinedSize** property to determine the data capacity of a field.

The **DefinedSize** and **ActualSize** properties are different. For example, consider a field with a declared type of *adVarChar* and a **DefinedSize** property value of 50, containing a single character. The **ActualSize** property value it returns is the length in bytes of the single character.

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

' // Get a reference to the Fields collection
DIM pFields AS CAdoFields = pRecordset.Fields

' // Parse the recordset
DO
   ' // While not at the end of the recordset...
   IF pRecordset.EOF THEN EXIT DO
   ' // Get the contents of the fields
   DIM pField AS CAdoField
   pField.Attach(pFields.Item("Name"))
   DIM cvRes AS CVAR = pField.Value
   PRINT "Name: "; cvRes; " - "; WSTR(pField.ActualSize); " - "; WSTR(pField.DefinedSize)
   ' // Fetch the next row
   IF pRecordset.MoveNext <> S_OK THEN EXIT DO
LOOP
```

# <a name="GetChunk"></a>GetChunk (CADOField)

Returns all, or a portion, of the contents of a large text or binary data field.

```
FUNCTION GetChunk (BYVAL nLen AS LONG) AS CVAR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nLen* | The number of bytes to retrieve. |

#### Return value

The retrieved data.

#### Remarks

Use the **GetChunk** method on a **Field** object to retrieve part or all of its long binary or character data. In situations where system memory is limited, you can use the **GetChunk** method to manipulate long values in portions, rather than in their entirety.

The data that a **GetChunk** call returns is assigned to variable. If **Size** is greater than the remaining data, the **GetChunk** method returns only the remaining data without padding variable with empty spaces. If the field is empty, the **GetChunk** method returns a null value.

Each subsequent **GetChunk** call retrieves data starting from where the previous **GetChunk** call left off. However, if you are retrieving data from one field and then you set or read the value of another field in the current record, ADO assumes you have finished retrieving data from the first field. If you call the **GetChunk** method on the first field again, ADO interprets the call as a new **GetChunk** operation and starts reading from the beginning of the data. Accessing fields in other **Recordset** objects that are not clones of the first **Recordset** object will not disrupt **GetChunk** operations.

If the **adFldLong** bit in the **Attributes** property of a **Field** object is set to True, you can use the **GetChunk** method for that field.

If there is no current record when you use the **GetChunk** method on a **Field** object, error 3021 (no current record) occurs.

**Note**: The **GetChunk** method does not operate on **Field** objects of a Record object. It does not perform any operation and will produce a run-time error.

# <a name="Name"></a>Name (CADOField)

Returns the name of the field.

```
PROPERTY Name () AS CBSTR
```

# <a name="NumericScale"></a>NumericScale (CADOField)

Sets or returns a Byte value that indicates the number of decimal places to which numeric values will be resolved.

```
PROPERTY NumericScale () AS BYTE
PROPERTY NumericScale (BYVAL bScale AS BYTE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *bScale* | The number of decimal places. |

#### Return value

The number of decimal places. Returns -1 if it is not a numeric field.

#### Remarks

Use the **NumericScale** property to determine how many digits to the right of the decimal point will be used to represent values for a numeric field.

For a field, **NumericScale** is normally read-only. However, for new fields that have been appended to a **Record**, **NumericScale** is read/write only after the **Value** property for the field has been specified and the data provider has successfully added the new field by calling the Update method of the **Fields** collection.

# <a name="OriginalValue"></a>OriginalValue (CADOField)

Indicates the value of a field that existed in the record before any changes were made.

```
PROPERTY OriginalValue () AS CVAR
```

#### Return value

The original field value.

#### Remarks

Use the **OriginalValue** property to return the original field value for a field from the current record.

In immediate update mode (in which the provider writes changes to the underlying data source after you call the **Update** method), the **OriginalValue** property returns the field value that existed prior to any changes (that is, since the last **Update** method call). This is the same value that the **CancelUpdate** method uses to replace the Value property.

In batch update mode (in which the provider caches multiple changes and writes them to the underlying data source only when you call the **UpdateBatch** method), the **OriginalValue** property returns the field value that existed prior to any changes (that is, since the last **UpdateBatch** method call). This is the same value that the **CancelBatch** method uses to replace the **Value** property. When you use this property with the **UnderlyingValue** property, you can resolve conflicts that arise from batch updates.

# <a name="Name"></a>Name (CADOField)

Returns the name of the field.

```
PROPERTY Name () AS CBSTR
```

# <a name="Precision"></a>Precision (CADOField)

Sets or returns a Byte value that indicates the maximum number of digits used to represent values.

```
PROPERTY Precision () AS BYTE
PROPERTY Precision (BYVAL bPrecision AS BYTE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *bPrecision* | The maximum number of digits used to represent values. |

#### Return value

The degress of precision. Returns -1 if it is not a numeric field.

#### Remarks

Use the **Precision** property to determine the maximum number of digits used to represent values for a numeric parameter or field.

For a field, **Precision** is normally read-only. However, for new fields that have been appended to a **Record**, **Precision** is read/write only after the **Value** property for the field has been specified and the data provider has successfully added the new field by calling the **Update** method of the **Fields** collection.

# <a name="Properties"></a>Properties (CADOField)

Returns a reference to the **Properties** collection.

```
PROPERTY Properties () AS Afx_ADOProperties
```

#### Return value

An ADO **Properties** object reference.

# <a name="Status"></a>Status (CADOField)

Indicates the status of a field. Returns a **FieldStatusEnum** value.

```
PROPERTY Status () AS LONG
```

#### Return value

A **FieldStatusEnum** value. The default value is adFieldOK.

#### FieldStatusEnum

Specifies the status of a Field object. 

The adFieldPending* values indicate the operation that caused the status to be set, and may be combined with other status values.

| Constant   | Description |
| ---------- | ----------- |
| **adFieldAlreadyExists** | Indicates that the specified field already exists. |
| **adFieldBadStatus** | Indicates that an invalid status value was sent from ADO to the OLE DB provider. Possible causes include an OLE DB 1.0 or 1.1 provider, or an improper combination of **Value** and **Status**. |
| **adFieldCannotComplete** | Indicates that the server of the URL specified by Source could not complete the operation. |
| **adFieldCannotDeleteSource** | Indicates that during a move operation, a tree or subtree was moved to a new location, but the source could not be deleted. |
| **adFieldCantConvertValue** | Indicates that the field cannot be retrieved or stored without loss of data. |
| **adFieldCantCreate** | Indicates that the field could not be added because the provider exceeded a limitation (such as the number of fields allowed). |
| **adFieldDataOverflow** | Indicates that the data returned from the provider overflowed the data type of the field. |
| **adFieldDefault** | Indicates that the default value for the field was used when setting data. |
| **adFieldDoesNotExist** | Indicates that the field specified does not exist. |
| **adFieldIgnore** | Indicates that this field was skipped when setting data values in the source. The provider set no value. |
| **adFieldIntegrityViolation** | Indicates that the field cannot be modified because it is a calculated or derived entity. |
| **adFieldInvalidURL** | Indicates that the data source URL contains invalid characters. |
| **adFieldIsNull** | Indicates that the provider returned a VARIANT value of type VT_NULL and that the field is not empty. |
| **adFieldOK** | Default. Indicates that the field was successfully added or deleted. |
| **adFieldOutOfSpace** | Indicates that the provider is unable to obtain enough storage space to complete a move or copy operation. |
| **adFieldPendingChange** | Indicates either that the field has been deleted and then re-added, perhaps with a different data type, or that the value of the field which previously had a status of adFieldOK has changed. The final form of the field will modify the Fields collection after the **Update** method is called. |
| **adFieldPendingDelete** | Indicates that the **Delete_** operation caused the status to be set. The field has been marked for deletion from the **Fields** collection after the **Update** method is called. |
| **adFieldPendingInsert** | Indicates that the **Append** operation caused the status to be set. The **Field** has been marked to be added to the Fields collection after the **Update** method is called. |
| **adFieldPendingUnknown** | Indicates that the provider cannot determine what operation caused field status to be set. |
| **adFieldPendingUnknownDelete** | Indicates that the provider cannot determine what operation caused field status to be set, and that the field will be deleted from the **Fields** collection after the **Update** method is called. |
| **adFieldPermissionDenied** | Indicates that the field cannot be modified because it is defined as read-only. |
| **adFieldReadOnly** | Indicates that the field in the data source is defined as read-only. |
| **adFieldResourceExists** | Indicates that the provider was unable to perform the operation because an object already exists at the destination URL and it is not able to overwrite the object. |
| **adFieldResourceLocked** | Indicates that the provider was unable to perform the operation because the data source is locked by one or more other application or process. |
| **adFieldResourceOutOfScope** | Indicates that a source or destination URL is outside the scope of the current record. |
| **adFieldSchemaViolation** | Indicates that the value violated the data source schema constraint for the field. |
| **adFieldSignMismatch** | Indicates that data value returned by the provider was signed but the data type of the ADO field value was unsigned. |
| **adFieldTruncated** | Indicates that variable-length data was truncated when reading from the data source. |
| **adFieldUnavailable** | Indicates that the provider could not determine the value when reading from the data source. For example, the row was just created, the default value for the column was not available, and a new value had not yet been specified. |
| **adFieldVolumeNotFound** | Indicates that the provider is unable to locate the storage volume indicated by the URL. |

#### Remarks

Changes to the value of a **Field** object in the **Fields** collection of a **Record** object are cached until the object's **Update** method is called. At that point, if the change to the **Field**'s value caused an error, OLE DB raises the error DB_E_ERRORSOCCURRED (2147749409). The **Status** property of any of the **Field** objects in the **Fields** collection that caused the error will contain a value from the **FieldStatusEnum** describing the cause of the problem.

To enhance performance, additions and deletions to the **Fields** collections of the **Record** object are cached until the **Update** method is called, and then the changes are made in a batch optimistic update. If the **Update** method is not called, the server is not updated. If any updates fail then an OLE DB provider error (DB_E_ERRORSOCCURRED) is returned and the **Status** property indicates the combined values of the operation and error status code. For example, **adFieldPendingInsert OR adFieldPermissionDenied**. The **Status** property for each **Field** can be used to determine why the **Field** was not added, modified, or deleted.

Many types of problems encountered when adding, modifying, or deleting a **Field** are reported through the **Status** property. For example, if the user deletes a **Field**, it is marked for deletion from the **Fields** collection. If the subsequent **Update** returns an error because the user tried to delete a **Field** for which they do not have permission, the **Field** will have a **Status** of **adFieldPermissionDenied** or **adFieldPendingDelete**. Calling the **CancelUpdate** method restores original values and sets the **Status** to **adFieldOK**.

Similarly, the **Update** method may return an error because a new **Field** was added and given an inappropriate value. In that case the new **Field** will be in the **Fields** collection and have a status of **adFieldPendingInsert** and perhaps **adFieldCantCreate** (depending upon your provider). You can supply an appropriate value for the new **Field** and call **Update** again.

#### Recordset Field Status

Changes to the value of a **Field** object in the **Fields** collection of either a **Recordset** are cached until the object's **Update** method is called. At that point, if the change to the **Field**'s value caused an error, OLE DB raises the error DB_E_ERRORSOCCURRED (2147749409). The **Status** property of any of the **Field** objects in the **Fields** collection that caused the error will contain a value from the **FieldStatusEnum** describing the cause of the problem.

# <a name="Type_"></a>Type_ (CADOField)

Sets or returns a **DataTypeEnum** value.

```
PROPERTY Type_ () AS DataTypeEnum
PROPERTY Type_ (BYVAL nType AS DataTypeEnum)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nType* | A **DataTypeEnum** value. |

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

#### Remarks

For a new fields that have been appended to the **Fields** collection of a **Record**, **Type_** is read/write only after the Value property for the field has been specified and the data provider has successfully added the new field by calling the **Update** method of the **Fields** collection.

For all other objects, the **Type_** property is read-only.

# <a name="UnderlyingValue"></a>UnderlyingValue (CADOField)

Returns the current field's value in the database.

```
PROPERTY UnderlyingValue () AS CVAR
```

# <a name="Value"></a>Value (CADOField)

Sets or returns a Variant value that indicates the value of the object

```
PROPERTY Value () AS CVAR
PROPERTY Value (BYREF cvValue AS CVAR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cvValue* | The value of the field. |

#### Return value

The value of the field.

#### Remarks

The ADO Recorset object exposes a hidden member: the **Collect** property. This property is functionally similar to the **Field**'s **Value** property, but it doesn't need a reference (explicit or implicit) to the **Field** object. You can pass either a numeric index or a field's name to this property.

#### Example

```
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Open the recordset
DIM pRecordset AS CAdoRecordset
DIM cvSource AS CVAR = "SELECT * FROM Publishers ORDER BY PubID"
DIM hr AS HRESULT = pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

' // Get a reference to the Fields collection
DIM pFields AS CAdoFields = pRecordset.Fields

' // Parse the recordset
DO
   ' // While not at the end of the recordset...
   IF pRecordset.EOF THEN EXIT DO
   ' // Get the contents of the fields
   ' // Note: Instead of pField.Attach / pField.Value, we can use pRecordset.Collect
    DIM pField AS CAdoField
    pField.Attach(pFields.Item("PubID"))
    DIM cvRes1 AS CVAR = pField.Value
    pField.Attach(pFields.Item("Name"))
    DIM cvRes2 AS CVAR = pField.Value
    pField.Attach(pFields.Item("Company Name"))
    DIM cvRes3 AS CVAR = pField.Value
    PRINT cvRes1; " "; cvRes2; " "; cvRes3
   ' // Fetch the next row
   IF pRecordset.MoveNext <> S_OK THEN EXIT DO
LOOP
```
