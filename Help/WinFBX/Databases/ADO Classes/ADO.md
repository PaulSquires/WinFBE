# CADODB

Microsoft ActiveX Data Objects (ADO) enable your client applications to access and manipulate data from a variety of sources through an OLE DB provider. Its primary benefits are ease of use, high speed, low memory overhead, and a small disk footprint. ADO supports key features for building client/server and Web-based applications. 

`CADODB` are a collection of classes to allow to work with ADO using the FreeBasic compilers.

The `CADOBase` object, from which the other ADO classes inherit, initializes and uninitializes the COM library and implements two methods, **GetLastResult** and **SetResult** used by the derived classes to store the result codes of ADO calls.

**Folder**: Afx/CADODB<br>
**Files**: CADODB.inc, CADOCommand.inc, CADOConnection.inc, CADOErrors.inc, CADOFields.inc, CADOParameters.inc, CADOProperties.inc, CADORecord.inc, CADORecordset.inc, CADOStream.inc.

| Name       | Description |
| ---------- | ----------- |
| [CADOBase Class](#CADOBase) | Base class for all the other ADO classes. |
| [ADO Object Model](#ADOObjectModel) | ADO objects and their collections. |
| [ADO Identifiers](#ADOIdentifiers) | PROGIDs, CLSIDs and IIDs. |
| [ADO Errors](#ADOErrors) | ADO errors. |
| [ADO Properties](#ADOProperties) | Represents a dynamic characteristic of an ADO object that is defined by the provider. |

# <a name="CADOBase"></a>CADOBase Class

The **CAdoBase** class, from which the other ADO classes inherit, initializes and uninitializes the COM library and implements two methods, **GetLastResult** and **SetResult** used by the derived classes to store the result codes of ADO calls.

### <a name="GetLastResult"></a>GetLastResult

Returns the last result code.

```
FUNCTION GetLastResult () AS HRESULT
```

### <a name="SetResult"></a>SetResult

Sets the last result code.

```
FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
```

#### Return value

The result code returned by the last executed method.

# <a name="ADOObjectModel"></a>ADO Object Model

ADO objects and their collections:

```
Connection
  |_ Errors     - Error
  |_ Properties - Property

Command
  |_ Parameters - Parameter
  |_ Properties - Property

Recordset
  |_ Fields - Field
  |_ Properties - Property

Record
  |_ Fields - Field

Stream
```

# <a name="ADOIdentifiers"></a>ADO Identifiers

```
' ========================================================================================
' ProgIDs (Program identifiers)
' ========================================================================================

' CLSID = {00000507-0000-0010-8000-00AA006D2EA4}
CONST AFX_PROGID_Command60 = "ADODB.Command.6.0"
' CLSID = {00000514-0000-0010-8000-00AA006D2EA4}
CONST AFX_PROGID_Connection60 = "ADODB.Connection.6.0"
' CLSID = {0000050B-0000-0010-8000-00AA006D2EA4}
CONST AFX_PROGID_Parameter60 = "ADODB.Parameter.6.0"
' CLSID = {00000560-0000-0010-8000-00AA006D2EA4}
CONST AFX_PROGID_Record60 = "ADODB.Record.6.0"
' CLSID = {00000535-0000-0010-8000-00AA006D2EA4}
CONST AFX_PROGID_Recordset60 = "ADODB.Recordset.6.0"
' CLSID = {00000566-0000-0010-8000-00AA006D2EA4}
CONST AFX_PROGID_Stream60 = "ADODB.Stream.6.0"

' ========================================================================================
' Version independent ProgIDs
' ========================================================================================

' CLSID = {00000507-0000-0010-8000-00AA006D2EA4}
CONST AFX_PROGID_Command = "ADODB.Command"
' CLSID = {00000514-0000-0010-8000-00AA006D2EA4}
CONST AFX_PROGID_Connection = "ADODB.Connection"
' CLSID = {0000050B-0000-0010-8000-00AA006D2EA4}
CONST AFX_PROGID_Parameter = "ADODB.Parameter"
' CLSID = {00000560-0000-0010-8000-00AA006D2EA4}
CONST AFX_PROGID_Record = "ADODB.Record"
' CLSID = {00000535-0000-0010-8000-00AA006D2EA4}
CONST AFX_PROGID_Recordset = "ADODB.Recordset"
' CLSID = {00000566-0000-0010-8000-00AA006D2EA4}
CONST AFX_PROGID_Stream = "ADODB.Stream"

' ========================================================================================
' ClsIDs (Class identifiers)
' ========================================================================================

CONST AFX_CLSID_Command = "{00000507-0000-0010-8000-00AA006D2EA4}"
CONST AFX_CLSID_Connection = "{00000514-0000-0010-8000-00AA006D2EA4}"
CONST AFX_CLSID_Parameter = "{0000050B-0000-0010-8000-00AA006D2EA4}"
CONST AFX_CLSID_Record = "{00000560-0000-0010-8000-00AA006D2EA4}"
CONST AFX_CLSID_Recordset = "{00000535-0000-0010-8000-00AA006D2EA4}"
CONST AFX_CLSID_Stream = "{00000566-0000-0010-8000-00AA006D2EA4}"

' ========================================================================================
' IIDs (Interface identifiers)
' ========================================================================================

CONST AFX_IID_ADO = "{00000534-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_Collection = "{00000512-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_Command = "{986761E8-7269-4890-AA65-AD7C03697A6D}"
CONST AFX_IID_Connection = "{00001550-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_DynaCollection = "{00000513-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_Parameter = "{0000150C-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_Record = "{00001562-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_Recordset = "{00001556-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_Stream = "{00001565-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_ADOCommandConstruction = "{00000517-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_ADOConnectionConstruction = "{00000551-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_ADORecordConstruction = "{00000567-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_ADORecordsetConstruction = "{00000283-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_ADOStreamConstruction = "{00000568-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_ConnectionEvents = "{00001400-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_ConnectionEventsVt = "{00001402-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_Error = "{00000500-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_Errors = "{00000501-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_Field = "{00001569-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_Fields = "{00001564-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_Fields_Deprecated = "{00000564-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_Parameters = "{0000150D-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_Properties = "{00000504-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_Property = "{00000503-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_RecordsetEvents = "{00001266-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_RecordsetEventsVt = "{00001403-0000-0010-8000-00AA006D2EA4}"
```

# <a name="ADOErrors"></a>ADO Errors

Any operation involving ADO objects can generate one or more provider errors. As each error occurs, one or more **Error** objects can be placed in the **Errors collection** of the **Connection object**. When another ADO operation generates an error, the **Errors** collection is cleared, and the new set of **Error** objects can be placed in the **Errors** collection.

Each **Error** object represents a specific provider error, not an ADO error. ADO operations that don't generate an error have no effect on the **Errors** collection. Use the **Clear** method to manually clear the **Errors** collection.

The set of **Error** objects in the **Errors** collection describes all errors that occurred in response to a single statement. Enumerating the specific errors in the **Errors** collection enables your error-handling routines to more precisely determine the cause and origin of an error, and take appropriate steps to recover.

Some properties and methods return warnings that appear as **Error** objects in the **Errors** collection but do not halt a program's execution. Before you call the **Resync**, **UpdateBatch**, or **CancelBatch** methods on a **Recordset** object, the **Open** method on a **Connection** object, or set the **Filter** property on a **Recordset** object, call the **Clear** method on the **Errors** collection. That way you can read the **Count** property of the **Errors** collection to test for returned warnings.

The **Error** object contains details about data access errors that pertain to a single operation involving the provider. Any operation involving ADO objects can generate one or more provider errors. As each error occurs, one or more **Error** objects are placed in the **Errors** collection of the **Connection** object. When another ADO operation generates an error, the **Errors** collection is cleared, and the new set of **Error** objects is placed in the **Errors** collection.

**Note**: Each **Error** object represents a specific provider error, not an ADO error. ADO errors are exposed to the run-time exception-handling mechanism. For example, in Microsoft Visual Basic, the occurrence of an ADO-specific error will trigger an **On Error** event and appear in the **Error** object. For a complete list of ADO errors, see the **ErrorValueEnum** table below.

You can read an **Error** object's properties to obtain specific details about each error, including the following:

* The **Description** property, which contains the text of the error. This is the default property.
* The **Number** property, which contains the Long integer value of the error constant.
* The **Source** property, which identifies the object that raised the error. This is particularly useful when you have several Error objects in the Errors collection following a request to a data source.
* The **SQLState** and **NativeError** properties, which provide information from SQL data sources.

When a provider error occurs, it is placed in the **Errors** collection of the **Connection** object. ADO supports the return of multiple errors by a single ADO operation to allow for error information specific to the provider. To obtain this rich error information in an error handler, use the appropriate error-trapping features of the language or environment you are working with, then use nested loops to enumerate the properties of each **Error** object in the **Errors** collection.

To retrieve information about ADO errors, call the following function:

```
FUNCTION AfxAdoGetErrorInfo (BYVAL pConnection AS AFX_ADOConnection PTR, BYVAL nError AS HRESULT = 0) AS CBSTR
```

Where *pConnection* is a reference to the ADO **Connection** object and *nError* the ADO HRESULT code.

### ErrorValueEnum Enumeration

Specifies the type of ADO run-time error.

Three forms of the error number are listed:

* Positive decimal—the low two bytes of the full number in decimal format.
* Negative decimal—The decimal translation of the full error number.
* Hexadecimal—The hexadecimal representation of the full error number. The Windows facility code is in the fourth digit. The facility code for ADO error numbers is A. For example: &H800A0E7B.

**Note**: OLE DB errors may be passed to your ADO application. Typically, these can be identified by a Windows facility code of 4. For example, &H8004....

| Constant   | Value | Description |
| ---------- | ----- | ----------- |
| **adErrBoundToCommand** | 3707<br>-2146824581<br>&H800A0E7B | Cannot change the **ActiveConnection** property of a **Recordset** object which has a **Command** object as its source. |
| **adErrCannotComplete** | 3732<br>-2146824556<br>&H800A0E94 | Server cannot complete the operation. |
| **adErrCantChangeConnection** | 3748<br>-2146824540<br>&H800A0EA4 | Connection was denied. New connection you requested has different characteristics than the one already in use. |
| **adErrCantChangeProvider** | 3220<br>-2146825068<br>&H800A0C94 | Supplied provider is different from the one already in use. |
| **adErrCantConvertvalue** | 3724<br>-2146824564<br>&H800A0E8C | Data value cannot be converted for reasons other than sign mismatch or data overflow. For example, conversion would have truncated data. |
| **adErrCantCreate** | 3725<br>-2146824563<br>&H800A0E8D | Data value cannot be set or retrieved because the field data type was unknown, or the provider had insufficient resources to perform the operation. |
| **adErrCatalogNotSet** | 3747<br>-2146824541<br>&H800A0EA3 | Operation requires a valid **ParentCatalog**. |
| **adErrColumnNotOnThisRow** | 3726<br>-2146824562<br>&H800A0E8E | Record does not contain this field. |
| **adErrDataConversion** | 3421<br>-2146824867<br>&H800A0D5D | Application uses a value of the wrong type for the current operation. |
| **adErrDataOverflow** | 3721<br>-2146824567<br>&H800A0E89 | Data value is too large to be represented by the field data type. |
| **adErrDelResOutOfScope** | 3738<br>-2146824550<br>&H800A0E9A | URL of the object to be deleted is outside the scope of the current record. |
| **adErrDenyNotSupported** | 3750<br>-2146824538<br>&H800A0EA6 | Provider does not support sharing restrictions. |
| **adErrDenyTypeNotSupported** | 3751<br>-2146824537<br>&H800A0EA7 | Provider does not support the requested kind of sharing restriction. |
| **adErrFeatureNotAvailable** | 3251<br>-2146825037<br>&H800A0CB3 | Object or provider is not capable of performing requested operation. |
| **adErrFieldsUpdateFailed** | 3749<br>-2146824539<br>&H800A0EA5 | Fields update failed. For further information, examine the Status property of individual field objects. |
| **adErrIllegalOperation** | 3219<br>-2146825069<br>&H800A0C93 | Operation is not allowed in this context. |
| **adErrIntegrityViolation** | 3719<br>-2146824569<br>&H800A0E87 | Data value conflicts with the integrity constraints of the field. |
| **adErrInTransaction** | 3246<br>-2146825042<br>&H800A0CAE | Connection object cannot be explicitly closed while in a transaction. |
| **adErrInvalidArgument** | 3001<br>-2146825287<br>&H800A0BB9 | Arguments are of the wrong type, are out of acceptable range, or are in conflict with one another. |
| **adErrInvalidConnection** | 3709<br>-2146824579<br>&H800A0E7D | The connection cannot be used to perform this operation. It is either closed or invalid in this context. |
| **adErrInvalidParamInfo** | 3708<br>-2146824580<br>&H800A0E7C | Parameter object is improperly defined. Inconsistent or incomplete information was provided. |
| **adErrInvalidTransaction** | 3714<br>-2146824574<br>&H800A0E82 | Coordinating transaction is invalid or has not started. |
| **adErrInvalidURL** | 3729<br>-2146824559<br>&H800A0E91 | URL contains invalid characters. Make sure the URL is typed correctly. |
| **adErrItemNotFound** | 3265<br>-2146825023<br>&H800A0CC1 | Item cannot be found in the collection corresponding to the requested name or ordinal. |
| **adErrNoCurrentRecord** | 3021<br>-2146825267<br>&H800A0BCD | Either BOF or EOF is True, or the current record has been deleted. Requested operation requires a current record. |
| **adErrNotExecuting** | 3715<br>-2146824573<br>&H800A0E83 | Operation cannot be performed while not executing. |
| **adErrNotReentrant** | 3710<br>-2146824578<br>&H800A0E7E | Operation cannot be performed while processing event. |
| **adErrObjectClosed** | 3704<br>-2146824584<br>&H800A0E78 | Operation is not allowed when the object is closed. |
| **adErrObjectInCollection** | 3367<br>-2146824921<br>&H800A0D27 | Object is already in collection. Cannot append. |
| **adErrObjectNotSet** | 3420<br>-2146824868<br>&H800A0D5C | Object is no longer valid. |
| **adErrObjectOpen** | 3705<br>-2146824583<br>&H800A0E79 | Operation is not allowed when the object is open. |
| **adErrOpeningFile** | 3002<br>-2146825286<br>&H800A0BBA | File could not be opened. |
| **adErrOperationCancelled** | 3712<br>-2146824576<br>&H800A0E80 | Operation has been cancelled by the user. |
| **adErrOutOfSpace** | 3734<br>-2146824554<br>&H800A0E96 | Operation cannot be performed. Provider cannot obtain enough storage space. |
| **adErrPermissionDenied** | 3720<br>-2146824568<br>&H800A0E88 | Insufficent permission prevents writing to the field. |
| **adErrProviderFailed** | 3000<br>-2146825288<br>&H800A0BB8 | Provider failed to perform the requested operation. |
| **adErrProviderNotFound** | 3706<br>-2146824582<br>&H800A0E7A | Provider cannot be found. It may not be properly installed. |
| **adErrReadFile** | 3003<br>-2146825285<br>&H800A0BBB | File could not be read. |
| **adErrResourceExists** | 3731<br>-2146824557<br>&H800A0E93 | Copy operation cannot be performed. Object named by destination URL already exists. Specify **adCopyOverwrite** to replace the object. |
| **adErrResourceLocked** | 3730<br>-2146824558<br>&H800A0E92 | Object represented by the specified URL is locked by one or more other processes. Wait until the process has finished and attempt the operation again. |
| **adErrResourceOutOfScope** | 3735<br>-2146824553<br>&H800A0E97 | Source or destination URL is outside the scope of the current record. |
| **adErrSchemaViolation** | 3722<br>-2146824566<br>&H800A0E8A | Data value conflicts with the data type or constraints of the field. |
| **adErrSignMismatch** | 3723<br>-2146824565<br>&H800A0E8B | Conversion failed because the data value was signed and the field data type used by the provider was unsigned. |
| **adErrStillConnecting** | 3713<br>-2146824575<br>&H800A0E81 | Operation cannot be performed while connecting asynchronously. |
| **adErrStillExecuting** | 3711<br>-2146824577<br>&H800A0E7F | Operation cannot be performed while executing asynchronously. |
| **adErrTreePermissionDenied** | 3728<br>-2146824560<br>&H800A0E90 | Permissions are insufficient to access tree or subtree. |
| **adErrUnavailable** | 3736<br>-2146824552<br>&H800A0E98 | Operation failed to complete and the status is unavailable. The field may be unavailable or the operation was not attempted. |
| **adErrUnsafeOperation** | 3716<br>-2146824572<br>&H800A0E84 | Safety settings on this computer prohibit accessing a data source on another domain. |
| **adErrURLDoesNotExist** | 3727<br>-2146824561<br>&H800A0E8F | Either the source URL or the parent of the destination URL does not exist. |
| **adErrURLNamedRowDoesNotExist** | 3737<br>-2146824551<br>&H800A0E99 | Record named by this URL does not exist. |
| **adErrVolumeNotFound** | 3733<br>-2146824555<br>&H800A0E95 | Provider cannot locate the storage device indicated by the URL. Make sure the URL is typed correctly. |
| **adErrWriteFile** | 3004<br>-2146825284<br>&H800A0BBC | Write to file failed. |
| **adWrnSecurityDialog** | 3717<br>-2146824571<br>&H800A0E85 | For internal use only. Don't use. |
| **adWrnSecurityDialogHeader** | 3718<br>-2146824570<br>&H800A0E86 | For internal use only. Don't use. |

# <a name="ADOProperties"></a>ADO Properties

Represents a dynamic characteristic of an ADO object that is defined by the provider.

ADO objects have two types of properties: built-in and dynamic.

Built-in properties are those properties implemented in ADO and immediately available to any new object, using the *MyObject.Property* syntax. They do not appear as **Property** objects in an object's **Properties** collection, so although you can change their values, you cannot modify their characteristics.

Dynamic properties are defined by the underlying data provider, and appear in the **Properties** collection for the appropriate ADO object. For example, a property specific to the provider may indicate if a **Recordset** object supports transactions or updating. These additional properties will appear as **Property** objects in that **Recordset** object's **Properties** collection. Dynamic properties can be referenced only through the collection, using the *MyObject.Properties(0)* or *MyObject.Properties("Name")* syntax.

You cannot delete either kind of property.

A dynamic **Property** object has four built-in properties of its own:

* The **Name** property is a string that identifies the property.
* The **Type_** property is an integer that specifies the property data type.
* The **Value** property is a variant that contains the property setting. **Value** is the default property for a **Property** object.
* The **Attributes** property is a long value that indicates characteristics of the property specific to the provider.

The **Properties** collection contains all the **Property** objects for a specific instance of an object.

# CADOProperty Class Methods

**Include file**: CAdoProperties.inc (include CADODB.inc).

## Attributes (CADOProperty Class)

For a **Property** object, the **Attributes** property is read-only, and its value can be the sum of any one or more **PropertyAttributesEnum** values.

```
PROPERTY Attributes () AS LONG
PROPERTY Attributes (BYVAL lAttr AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *lAttr* | Can be the sum of any one or more **ParameterAttributesEnum** values. The default is **adParamSigned**. |

#### Return value

LONG. One or more **ParameterAttributesEnum** values.

#### PropertyAttributesEnum

| Constant   | Value       |
| ---------- | ----------- |
| **adPropNotSupported** | Indicates that the property is not supported by the provider. |
| **adPropRequired** | Indicates that the user must specify a value for this property before the data source is initialized. |
| **adPropOptional** | Indicates that the user does not need to specify a value for this property before the data source is initialized. |
| **adPropRead** | Indicates that the user can read the property. |
| **adPropWrite** | Indicates that the user can set the property. |

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

## Name (CADOProperty Class)

Returns the name of a **Property**.

```
PROPERTY Name () AS CBSTR
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
DIM cvSource AS CVAR = "SELECT * FROM Publishers ORDER BY PubID"
DIM hr AS HRESULT = pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

' // Parse the Properties collection
DIM pProperties AS CAdoProperties = pRecordset.Properties
DIM nCount AS LONG = pProperties.Count
FOR i AS LONG = 0 TO nCount - 1
   DIM pProperty AS CAdoProperty = pProperties.Item(i)
   PRINT "Property name: "; pProperty.Name; " - Attributes: "; WSTR(pProperty.Attributes)
NEXT
```

## Type_ (CADOProperty Class)

Returns a **DataTypeEnum** value that indicates the operational type or data type of a **Property** object.

```
PROPERTY Type_ () AS DataTypeEnum
```

#### DataTypeEnum

| Constant   | Value       |
| ---------- | ----------- |
| **AdArray* | A flag value, always combined with another data type constant, that indicates an array of that other data type. |
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
| **adIDispatch** | Indicates a pointer to an IDispatch interface on a COM object (DBTYPE_IDISPATCH). This data type is currently not supported by ADO. Usage may cause unpredictable results. |
| **adInteger** | Indicates a four-byte signed integer (DBTYPE_I4). |
| **adIUnknown** | Indicates a pointer to an IUnknown interface on a COM object (DBTYPE_IUNKNOWN). This data type is currently not supported by ADO. Usage may cause unpredictable results. |
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
| **adVariant** | Indicates an Automation Variant (DBTYPE_VARIANT). This data type is currently not supported by ADO. Usage may cause unpredictable results. |
| **adVarNumeric** | Indicates a numeric value. |
| **adVarWChar** | Indicates a null-terminated Unicode character string. |
| **adWChar** | Indicates a null-terminated Unicode character string (DBTYPE_WSTR). |

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

' // Parse the Properties collection
DIM pProperties AS CAdoProperties = pRecordset.Properties
DIM nCount AS LONG = pProperties.Count
FOR i AS LONG = 0 TO nCount - 1
   DIM pProperty AS CAdoProperty = pProperties.Item(i)
   PRINT "Property name: "; pProperty.Name; " - Type: "; WSTR(pProperty.Type_)
NEXT
```

## Value (CADOProperty Class)

Sets or returns a Variant value that indicates the value of the object. Default value depends on the **Type_** property.

```
PROPERTY Value () AS CVAR
PROPERTY Value (BYREF cvValue AS CVAR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cvValue* | CVAR. The value of the **Property** object. |

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

' // Parse the Properties collection
DIM pProperties AS CAdoProperties = pRecordset.Properties
DIM nCount AS LONG = pProperties.Count
FOR i AS LONG = 0 TO nCount - 1
   DIM pProperty AS CAdoProperty = pProperties.Item(i)
   PRINT "Property name: "; pProperty.Name; " - Value: "; WSTR(pProperty.Value)
NEXT
```

# CADOProperties Class Methods

**Include file**: CAdoProperties.inc (include CADODB.inc).

## Count (CADOProperties Class)

Retrieves the number of objects of the **Properties** collection.

```
PROPERTY Count () AS LONG
```

#### Remarks

Because numbering for members of a collection begins with zero, you should always code loops starting with the zero member and ending with the value of the Count property minus 1.

If the **Count** property is zero, there are no objects in the collection.

## Item (CADOProperties Class)

Indicates a specific member of the **Properties** collection, by name or ordinal number.

```
PROPERTY Item (BYREF cvIndex AS CVAR) AS Afx_ADOProperty PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *cvIndex* | A Variant expression that evaluates either to the name or to the ordinal number of an object in a collection. |

#### Return value

An **Afx_ADOProperty** object reference.

#### Remarks

If **Item** cannot find an object in the collection corresponding to the **Index** argument, an error occurs.

## Refresh (CADOProperties Class)

Refreshes the contents of the **Properties** collection.

```
FUNCTION Refresh () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

Using the **Refresh** method on a **Properties** collection of some objects populates the collection with the dynamic properties that the provider exposes. These properties provide information about functionality specific to the provider, beyond the built-in properties ADO supports.
