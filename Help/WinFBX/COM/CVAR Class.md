# CVAR Class

The `CVAR` class implements a `VARIANT` data type. The variant data type is a tagged union that can be used to represent any other data type. While lacking in efficiency, they are heavily used in COM Automation for its flexibility. The main purpose of the `CVAR` class is to make its use as easy as possible when you need to use them to work with COM Automation objects.

**Include file**: CVAR.INC.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#Constructors) | Initialize the class with the specified value. |
| [Operators](#Operators) | Procedures that perform a certain function with their operands. |
| [vType](#vType) | Returns the VARIANT type. |
| [bstr](#bstr) | Extracts the content of the underlying variant and returns it as a CBSTR. |
| [sptr](#sptr) | Returns the address of the underlying variant. Same as vptr but without clearing the variant. |
| [vptr](#vptr) | Clears the CVAR and returns the address of the underlying variant. |
| [wstr](#wstr) | Extracts the content of the underlying variant and returns it as a CWSTR. |
| [Attach](#Attach) | Attaches a variant to the class. |
| [Detach](#Detach) | Detaches the variant data from this class and transfers ownership to the passed variant. |
| [ChangeType](#ChangeType) | Converts the variant from one type to another. |
| [ChangeTypeEx](#ChangeTypeEx) | Converts the variant from one type to another. |
| [GetDim](#GetDim) | Gets the number of dimensions in the array. |
| [GetLBound](#GetLBound) | Gets the lower bound for the specified dimension of the safe array. |
| [GetUBound](#GetUBound) | Gets the upper bound for the specified dimension of the safe array. |
| [GetElementCount](#GetElementCount) | Gets the number of elements in the array. |
| [DecToCY](#DecToCY) | Converts a CVAR of type decimal to a CY structure. |
| [DecToDouble](#DecToDouble) | Converts a CVAR of type decimal to a double. |
| [Round](#Round) | Rounds a variant to the specified number of decimal places. |
| [FormatNumber](#FormatNumber) | Formats a CVAR containing numbers into a string form. |
| [GetBooleanElem](#GetBooleanElem) | Extracts a single boolean element from a safe array of booleans. |
| [GetDoubleElem](#GetDoubleElem) | Extracts a single DOUBLE element from a safe array of doubles. |
| [GetLongElem](#GetLongElem) | Extracts a single LONG element from a safe array of longs. |
| [GetLongIntElem](#GetLongIntElem) | Extracts a single LONGINT element from a safe array of long integers. |
| [GetShortElem](#GetShortElem) | Extracts a single SHORT element from a safe array of shorts. |
| [GetStringElem](#GetStringElem) | Extracts a single BSTR element from a safe array of unicode strings. |
| [GetULongElem](#GetULongElem) | Extracts a single ULONG element from a safe array of unsigned longs. |
| [GetULongIntElem](#GetULongIntElem) | Extracts a single ULONGINT element from a safe array of unsigned long integers. |
| [GetUShortElem](#GetUShortElem) | Extracts a single USHORT element from a safe array of unsigned shorts. |
| [GetVariantElem](#GetVariantElem) | Extracts a single Variant element from a safe array of variants. |
| [Put](#Put) | Assigns values to a CVAR. |
| [PutNull](#PutNull) | Assigns a null value. |
| [PutBool](#PutBool) | Assigns a boolean value. |
| [PutBoolean](#PutBoolean) | Assigns a boolean value. |
| [PutByte](#PutByte) | Assigns a byte value. |
| [PutUByte](#PutUByte) | Assigns an ubyte value. |
| [PutShort](#PutShort) | Assigns a short value. |
| [PutUShort](#PutUShort) | Assigns an ushort value. |
| [PutInt](#PutInt) | Assigns an int_ value. |
| [PutUInt](#PutUInt) | Assigns an uint value. |
| [PutLong](#PutLong) | Assigns a long value. |
| [PutULong](#PutULong) | Assigns an ulong value. |
| [PutLongInt](#PutLongInt) | Assigns a longint value. |
| [PutULongInt](#PutULongInt) | Assigns an ulongint value. |
| [PutSingle](#PutSingle) | Assigns a single value. |
| [PutFloat](#PutFloat) | Assigns a single value. |
| [PutDouble](#PutDouble) | Assigns a double value. |
| [PutBooleanArray](#PutBooleanArray) | Initializes CVAR from an array of Boolean values. |
| [PutShortArray](#PutShortArray) | Initializes CVAR from an array of signed 16-bit integer values. |
| [PutUShortArray](#PutUShortArray) | Initializes CVAR from an array of unsigned 16-bit integer values. |
| [PutLongArray](#PutLongArray) | Initializes CVAR from an array of signed 32-bit integer values. |
| [PutULongArray](#PutULongArray) | Initializes CVAR from an array of 32-bit unsigned integer values. |
| [PutLongIntArray](#PutLongIntArray) | Initializes CVAR from an array of signed 64-bit integer values. |
| [PutULongIntArray](#PutULongIntArray) | Initializes CVAR from an array of unsigned 64-bit integer values. |
| [PutDoubleArray](#PutDoubleArray) | Initializes CVAR from an array of unsigned 64-bit integer values. |
| [PutStringArray](#PutStringArray) | Initializes CVAR from an array of unsigned 64-bit integer values. |
| [PutBuffer](#PutBuffer) | Initializes CVAR with the contents of a buffer. |
| [PutDateString](#PutDateString) | Initializes CVAR VT_DATE from a string. |
| [PutDec](#PutDec) | Initializes CVAR with the contents of a DECIMAL structure. |
| [PutDecFromCY](#PutDecFromCY) | Converts a currency value to a variant of type VT_DECIMAL. |
| [PutDecFromDouble](#PutDecFromDouble) | Converts a double value to a variant of type VT_DECIMAL. |
| [PutDecFromStr](#PutDecFromStr) | Initializes CVAR as VT_DECIMAL from a string. |
| [PutFileTime](#PutFileTime) | Initializes CVAR with the contents of a FILETIME structure. |
| [PutFileTimeArray](#PutFileTimeArray) | Initializes CVAR with an array of FILETIME structures. |
| [PutGuid](#PutGuid) | Initializes CVAR from a GUID. |
| [PutPropVariant](#PutPropVariant) | Initializes CVAR from the contents of a PROPVARIANT structure. |
| [PutRecord](#PutRecord) | Initializes CVAR with a reference to an UDT. |
| [PutRef](#PutRef) | Assigns a value by reference (a pointer to a variable). |
| [PutResource](#PutResource) | Initializes the CVAR based on a string resource imbedded in an executable file. |
| [PutSafeArray](#PutSafeArray) | Initializes CVAR from a safe array. |
| [PutStrRet](#PutStrRet) | Initializes CVAR with string stored in a STRRET structure. |
| [PutSystemTime](#PutSystemTime) | Initializes CVAR with the contents of a SYSTEMTIME structure. |
| [PutUtf8](#PutUtf8) | Initializes CVAR with the contents of an UTF-8 string. |
| [PutVariantArrayElem](#PutVariantArrayElem) | Initializes CVAR with a value stored in another VARIANT structure. |
| [PutVbDate](#PutVbDate) | Initializes CVAR with the contents of a DATE value. |
| [ToBooleanArray](#ToBooleanArray) | Extracts an array of boolean values from CVAR. |
| [ToBooleanArrayAlloc](#ToBooleanArrayAlloc) | Extracts an array of boolean values from CVAR. |
| [ToBstr](#ToBstr) | Extracts the content of the underlying variant and returns it as a CBSTR. |
| [ToBuffer](#ToBuffer) | Extracts the contents of a CVAR of type VT_ARRRAY OR VT_UI1 to a buffer. |
| [ToBuffer (STRING)](#ToBufferString) | Extracts the contents of a CVAR of type VT_ARRRAY OR VT_UI1 to a string used as a buffer. |
| [ToDosDateTime](#ToDosDateTime) | Extracts a date and time value in Microsoft MS-DOS format from a CVAR of type VT_DATE. |
| [ToDoubleArray](#ToDoubleArray) | Extracts an array of DOUBLE values from CVAR. |
| [ToDoubleArrayAlloc](#ToDoubleArrayAlloc) | Extracts an array of DOUBLE values from CVAR. |
| [ToFileTime](#ToFileTime) | Returns the contents of a CVAR of type VT_DATE as a FILETIME structure. |
| [ToGuid](#ToGuid) | Returns the contents of a CVAR containing a GUID string as a GUID structure. |
| [ToGuidBStr](#ToGuidBStr) | Returns the contents of a CVAR containing a GUID string as an unicode GUID string. |
| [ToGuidStr](#ToGuidStr) | Returns the contents of a CVAR containing a GUID string as an unicode GUID string. |
| [ToGuidWStr](#ToGuidWStr) | Returns the contents of a CVAR containing a GUID string as an unicode GUID string. |
| [ToLongArray](#ToLongArray) | Extracts an array of LONG values from CVAR. |
| [ToLongArrayAlloc](#ToLongArrayAlloc) | Extracts an array of LONG values from CVAR. |
| [ToLongIntArray](#ToLongIntArray) | Extracts an array of LONGINT values from CVAR. |
| [ToLongIntArrayAlloc](#ToLongIntArrayAlloc) | Extracts an array of LONGINT values from CVAR. |
| [ToShortArray](#ToShortArray) | Extracts an array of Int16 values from CVAR. |
| [ToShortArrayAlloc](#ToShortArrayAlloc) | Extracts an array of SHORT values from CVAR. |
| [ToStr](#ToStr) | Extracts the content of the underlying variant and returns it as a CWSTR. |
| [ToStringArray](#ToStringArray) | Extracts data from a vector structure into a PWSTR array. |
| [ToStringArrayAlloc](#ToStringArrayAlloc) | Extracts an array of PWSTR values from CVAR. |
| [ToStrRet](#ToStrRet) | Returns the contents of a CVAR of type VT_BSTR to a STRRET stucture. |
| [ToSystemTime](#ToSystemTime) | Returns the contents of CVAR of type VT_DATE as a FILETIME structure. |
| [ToULongArray](#ToULongArray) | Extracts an array of ULONG values from CVAR. |
| [ToULongArrayAlloc](#ToULongArrayAlloc) | Extracts an array of ULONG values from CVAR. |
| [ToULongIntArray](#ToULongIntArray) | Extracts an array of ULONGINT values from CVAR. |
| [ToULongIntArrayAlloc](#ToULongIntArrayAlloc) | Extracts an array of ULONGINT values from CVAR. |
| [ToUShortArray](#ToUShortArray) | Extracts an array of USHORT values from CVAR. |
| [ToUShortArrayAlloc](#ToUShortArrayAlloc) | Extracts an array of USHORT values from CVAR. |
| [ToUtf8](#ToUtf8) | Returns the contents of a CVAR containing a BSTR as an UTF-8 encoded string. |
| [ToVbDate](#ToVbDate) | Returns the contents of a CVAR of type VT_DATE as a DATE value. |
| [ToWStr](#ToWStr) | Extracts the content of the underlying variant and returns it as a CWSTR. |

#### Numeric Conversions

| Name       | Description |
| ---------- | ----------- |
| [ValDouble](#ValDouble) | Converts the variant to a floating point number (DOUBLE). |
| [ValInt](#ValInt) | Converts the variant to a signed 32-bit integer (LONG). |
| [ValLong](#ValLong) | Converts the variant to a signed 32-bit integer (LONG). |
| [ValLongInt](#ValLongInt) | Converts the variant to a signed 64-bit integer (LONGINT). |
| [Value](#Value) | Converts the variant to a floating point number (DOUBLE). |
| [ValUInt](#ValUInt) | Converts the variant to a 32.bit unsigned integer (ULONG). |
| [ValULong](#ValULong) | Converts the variant to a 32-bit unsigned integer (ULONG). |
| [ValULongInt](#ValULongInt) | Converts the variant to a 64-bit unsigned integer (ULONGINT). |

#### Helper Procedures

| Name       | Description |
| ---------- | ----------- |
| [AfxCVarToStr](#AfxCVarToStr) | Extracts the contents of a CVAR to a CWSTR. |
| [AfxCVariantToBuffer](#AfxCVariantToBuffer) | Extracts the contents of a variant that contains an array of bytes. |
| [AfxCVarOptPrm](#AfxCVarOptPrm) | Returns a CVAR suitable to be used with optional parameters. |

# <a name="Constructors"></a>Constructors

Creates a an instance of the CVAR class.

```
CONSTRUCTOR
CONSTRUCTOR (BYREF cv AS CVAR)
CONSTRUCTOR (BYVAL v AS VARIANT)
CONSTRUCTOR (BYREF wsz AS WSTRING)
CONSTRUCTOR (BYREF cws AS CWSTR)
CONSTRUCTOR (BYREF cbs AS CBSTR)
CONSTRUCTOR (BYVAL pvar AS VARIANT PTR)
CONSTRUCTOR (BYVAL cy AS CURRENCY)
CONSTRUCTOR (BYVAL dec AS DECIMAL)
CONSTRUCTOR (BYREF pDisp AS IDispatch PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
CONSTRUCTOR (BYREF pUnk AS IUnknown PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
CONSTRUCTOR (BYVAL _value AS LONGINT, BYVAL _vType AS WORD = VT_I4)
CONSTRUCTOR (BYVAL _value AS DOUBLE, BYVAL _vType AS WORD = VT_R8)
CONSTRUCTOR (BYVAL _value AS LONGINT, BYREF strType AS STRING)
CONSTRUCTOR (BYVAL _value AS DOUBLE, BYREF strType AS STRING)
CONSTRUCTOR (BYVAL _pvar AS ANY PTR, BYVAL _vType AS WORD)
CONSTRUCTOR (BYVAL _pvar AS ANY PTR, BYREF strType AS STRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cv* | A CVAR. |
| *v* | A VARIANT. |
| *pvar* | Pointer to a VARIANT. |
| *cy* | A currency structure. |
| *dec* | A decimal structure. |
| *pwsz* | Pointer to an unicode string. You can also pass a Free Basic ansi string or a string literal. |
| *cbs* | A CBSTR. |
| *cws* | A CWSTR. |
| *pDisp* | Pointer to a DISPATCH interface. |
| *pUnk* | Pointer to a UNKNOWN interface. |
| *_value* | A numeric value or variable. |
| *_pvar* | Pointer to a variable. This will create a VT_BYREF variant of the specified type. |
| *_vtype* | The variant type, e.g. VT_I4, VT_UI4. |
| *strType* | The variant type as a string: "BOOL", "BYTE", "UBYTE", "SHORT", "USHORT, "INT", UINT", "LONG", "ULONG", "LONGINT", "SINGLE, "DOUBLE", "NULL". |
| *fAddRef* | TRUE or FALSE. If TRUE, increases the reference count of the passed interface. |

# <a name="Operators"></a>Operators

Procedures that perform a certain function with their operands. They do the same actions that the native FreeBasic operators but with variants. For detailed descriptions see the FreeBasic documentation.

Global Operators

```
OPERATOR & (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS CVAR
OPERATOR * (BYREF cv AS CVAR) AS VARIANT PTR
```

Cast Operators

```
OPERATOR Cast () AS VARIANT
OPERATOR Cast () AS VARIANT PTR
```

Assignment operators

```
OPERATOR Let (BYREF cv AS CVAR)
OPERATOR Let (BYVAL v AS VARIANT)
OPERATOR Let (BYVAL pvar AS VARIANT PTR)
OPERATOR Let (BYVAL cy AS CURRENCY)
OPERATOR Let (BYVAL dec AS DECIMAL)
OPERATOR Let (BYREF cbs AS CBSTR)
OPERATOR Let (BYREF cws AS CWSTR)
OPERATOR Let (BYREF pDisp AS IDispatch PTR)
OPERATOR Let (BYREF pUnk AS IUnknown PTR)
OPERATOR Let (BYVAL _value AS LONGINT)
OPERATOR Let (BYVAL _value AS DOUBLE)
OPERATOR += (BYREF cv AS CVAR)
OPERATOR -= (BYREF cv AS CVAR)
OPERATOR *= (BYREF cv AS CVAR)
OPERATOR /= (BYREF cv AS CVAR)
OPERATOR \= (BYREF cv AS CVAR)
OPERATOR Mod= (BYREF cv AS CVAR)
OPERATOR Imp= (BYREF cv AS CVAR)
OPERATOR Eqv= (BYREF cv AS CVAR)
OPERATOR ^= (BYREF cv AS CVAR)
```

Arithmetic operators

```
OPERATOR + (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS CVAR
OPERATOR - (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS CVAR
OPERATOR * (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS CVAR
OPERATOR / (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS CVAR
OPERATOR \ (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS CVAR
OPERATOR ^ (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS DOUBLE
OPERATOR Mod (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS INTEGER
OPERATOR - (BYREF cv AS CVAR) AS CVAR
```

Relational operators

```
OPERATOR = (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS BOOLEAN
OPERATOR <> (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS BOOLEAN
OPERATOR < (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS BOOLEAN
OPERATOR > (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS BOOLEAN
OPERATOR <= (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS BOOLEAN
OPERATOR >= (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS BOOLEAN
```

Bitwise operators

```
OPERATOR And (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS INTEGER
OPERATOR Eqv (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS INTEGER
OPERATOR Imp (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS INTEGER
OPERATOR Not (BYREF cv AS CVAR) AS INTEGER
OPERATOR Or (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS INTEGER
OPERATOR Xor (BYREF cv1 AS CVAR, BYREF cv2 AS CVAR) AS INTEGER
```

# <a name="vType"></a>vType

Returns the VARIANT type.

```
FUNCTION vType () AS VARTYPE
```

The following table shows the available data types and where these values can be used.

| Type                | Description                              | VARIANT | Typedesc | Property set | Safe array |
| ------------------- | ---------------------------------------- | ------- | -------- | ------------ | ---------- |
| VT_EMPTY            | Not specified.                           |    X    |          |      X       |            |
| VT_NULL             | Null.                                    |    X    |          |      X       |            |
| VT_I1               | BYTE. A character.                       |    X    |    X     |      X       |     X      |
| VT_UI1              | UBYTE. An unsigned character.            |    X    |    X     |      X       |     X      |
| VT_I2               | SHORT. A 2-byte integer.                 |    X    |    X     |      X       |     X      |
| VT_UI2              | USHORT. An unsigned short.               |    X    |    X     |      X       |     X      |
| VT_I4               | LONG. A 4-byte integer.                  |    X    |    X     |      X       |     X      |
| VT_UI4              | ULONG. An unsigned long.                 |    X    |    X     |      X       |     X      |
| VT_I8               | LONGINT. A 64-bit integer.               |    X    |    X     |      X       |     X      |
| VT_UI8              | ULONGINT. A 64-bit unsigned integer.     |    X    |    X     |      X       |     X      |
| VT_INT              | LONG. An integer.                        |    X    |    X     |      X       |     X      |
| VT_UINT             | ULOG. An unsigned integer.               |    X    |    X     |              |     X      |
| VT_R4               | SINGLE. A 4-byte real.                   |    X    |    X     |      X       |     X      |
| VT_R8               | DOUBLE. A 8-byte real.                   |    X    |    X     |      X       |     X      |
| VT_CY               | CY. Currency.                            |    X    |    X     |      X       |     X      |
| VT_DATE             | DOUBLE. A date.                          |    X    |    X     |      X       |     X      |
| VT_BSTR             | BSTR. A string.                          |    X    |    X     |      X       |     X      |
| VT_DISPATCH         | IDispatch PTR. An IDispatch pointer.     |    X    |    X     |              |     X      |
| VT_ERROR            | SCODE. An SCODE value.                   |    X    |    X     |      X       |     X      |
| VT_BOOL             | BOOLEAN. A Boolean value<br>(True = -1, False = 0)   |    X    |    X     |      X       |     X      |
| VT_VARIANT          | VARIANT PTR. A variant pointer.          |    X    |    X     |      X       |     X      |
| VT_UNKNOWN          | IUnknown PTR. An IUnknown pointer.       |    X    |    X     |              |     X      |
| VT_DECIMAL          | DECIMAL PTR. A 16-byte fixed-pointer value.           |    X    |    X     |              |     X      |
| VT_VOID             | NULL. A C-style void.                    |         |          |      X       |            |
| VT_HRESULT          | HRESULT. An HRESULT value.               |         |          |      X       |            |
| VT_PTR              | A pointer type.                          |         |          |      X       |            |
| VT_SAFEARRAY        | SAFEARRAY PTR. A safe array.<br>Use VT_ARRAY in VARIANT.   |         |    X     |              |            |
| VT_CARRAY           | A C-style array.                         |         |    X     |              |            |
| VT_USERDEFINED      | A user-defined type.                     |         |          |      X       |            |
| VT_LPSTR            | ZSTRING. A null-terminated string.       |         |    X     |      X       |            |
| VT_LPWSTR           | WSTRING. A wide null-terminated string.  |         |    X     |      X       |            |
| VT_RECORD           | A user-defined type.                     |    X    |    X     |              |     X      |
| VT_INT_PTR          | A signed machine register size width.    |         |          |      X       |            |
| VT_UINT_PTR         | An unsigned machine register size width. |         |          |      X       |            |
| VT_FILETIME         | FILETIME. A FILETIME value.              |         |          |      X       |            |
| VT_BLOB             | Length-prefixed bytes.                   |         |          |      X       |            |
| VT_STREAM           | IStream PTR. The name of the stream follows.          |         |          |      X       |            |
| VT_STORAGE          | The name of the storage follows.         |         |          |      X       |            |
| VT_STREAMED_OBJECT  | The stream contains an object.           |         |          |      X       |            |
| VT_STORED_OBJECT    | The storage contains an object.          |         |          |      X       |            |
| VT_BLOB_OBJECT      | The blob contains an object.             |         |          |      X       |            |
| VT_CF               | A clipboard format.                      |         |          |      X       |            |
| VT_CLSID            | CLSID. A class ID.                       |         |          |      X       |            |
| VT_VERSIONED_STREAM | A stream with a GUID version.            |         |          |      X       |            |
| VT_BSTR_BLOB        | Reserved for system use.                 |         |          |              |            |
| VT_VECTOR           | A simple counted array.                  |         |          |      X       |            |
| VT_ARRAY            | SAFEARRAY PTR. A SAFEARRAY pointer.      |    X    |          |              |            |
| VT_BYREF            | A void pointer for local use.            |    X    |          |              |            |
| VT_RESERVED         | Reserved.                                |         |          |              |            |
| VT_ILLEGAL          |                                          |         |          |              |            |
| VT_ILLEGALMASKED    |                                          |         |          |              |            |
| VT_TYPEMASK         |                                          |         |          |              |            |


# <a name="bstr"></a>bstr

Extracts the content of the underlying variant and returns it as a CBSTR.

```
FUNCTION bstr () AS CBSTR
```

# <a name="sptr"></a>sptr

Returns the address of the underlying variant. Same as **vptr** but without clearing the variant.

```
FUNCTION sptr () AS VARIANT PTR
```

# <a name="vptr"></a>vptr

Clears the CVAR and returns the address of the underlying variant. Can be used to pass the variant to an OUT BYVAL VARIANT PTR parameter. If we pass a CVAR to a function with an OUT variant parameter without first clearing the contents of the CVAR, we may have a memory leak.

```
FUNCTION vptr () AS VARIANT PTR
```

# <a name="wstr"></a>wstr

Extracts the content of the underlying variant and returns it as a CWSTR.

```
FUNCTION wstr () AS CWSTR
```

# <a name="Attach"></a>Attach

Attaches a variant to the class.

```
FUNCTION Attach (BYVAL pvar AS VARIANT PTR) AS HRESULT
FUNCTION Attach (BYREF v AS VARIANT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pvar* | Pointer to the variant to attach. |
| *v* | The variant to attach. |

#### Remark

Marks the source variant as VT_EMPTY instead of clearing it with VariantClear because we aren't making a duplicate of the contents, but transfering ownership.

#### Return value

Returns S_OK (0) or an HRESULT error code.

# <a name="Detach"></a>Detach

Detaches the variant data from this class and transfers ownership to the passed variant.

```
FUNCTION Detach (BYVAL pvar AS VARIANT PTR) AS HRESULT
FUNCTION Detach (BYREF v AS VARIANT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pvar* | Pointer to the variant where the contents of the variant data will be moved. |
| *v* | Variant where the contents of the variant data will be moved. |

#### Remark

This method transfers ownership of the underlying variant and marks it as empty.

#### Return value

Returns S_OK (0) or an HRESULT error code.

# <a name="ChangeType"></a>ChangeType

Converts the variant from one type to another.

```
FUNCTION ChangeType (BYVAL vtNew AS VARTYPE, BYVAL wFlags AS USHORT = 0) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *vtNew* | The new variant type. |
| *wFlags* | *VARIANT_NOVALUEPROP* : Prevents the function from attempting to coerce an object to a fundamental type by getting the Value property. Applications should set this flag only if necessary, because it makes their behavior inconsistent with other applications.<br>*VARIANT_ALPHABOOL* : Converts a VT_BOOL value to a string containing either "True" or "False".<br>*VARIANT_NOUSEROVERRIDE* : For conversions to or from VT_BSTR, passes LOCALE_NOUSEROVERRIDE to the core coercion routines.<br>*VARIANT_LOCALBOOL* : For conversions from VT_BOOL to VT_BSTR and back, uses the language specified by the locale in use on the local computer. |

#### Return value

Returns S_OK (0) or an HRESULT error code.

# <a name="ChangeTypeEx"></a>ChangeTypeEx

Converts the variant from one type to another.

```
FUNCTION ChangeTypeEx (BYVAL vtNew AS VARTYPE, BYVAL lcid AS LCID = 0, BYVAL wFlags AS USHORT = 0) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *vtNew* | The new variant type. |
| *lcid* | The locale identifier. The LCID is useful when the type of the source or destination VARIANTARG is VT_BSTR, VT_DISPATCH, or VT_DATE. |
| *wFlags* | *VARIANT_NOVALUEPROP* : Prevents the function from attempting to coerce an object to a fundamental type by getting the Value property. Applications should set this flag only if necessary, because it makes their behavior inconsistent with other applications.<br>*VARIANT_ALPHABOOL* : Converts a VT_BOOL value to a string containing either "True" or "False".<br>*VARIANT_NOUSEROVERRIDE* : For conversions to or from VT_BSTR, passes LOCALE_NOUSEROVERRIDE to the core coercion routines.<br>*VARIANT_LOCALBOOL* : For conversions from VT_BOOL to VT_BSTR and back, uses the language specified by the locale in use on the local computer. |

#### Return value

Returns S_OK (0) or an HRESULT error code.

# <a name="GetDim"></a>GetDim

Gets the number of dimensions in the array.

```
FUNCTION GetDim () AS ULONG
```

#### Return value

Returns the number of dimensions for variants of type VT_ARRAY; returns 0 otherwise.

# <a name="GetLBound"></a>GetLBound

Gets the lower bound for the specified dimension of the safe array.

```
FUNCTION GetLBound (BYVAL nDim AS UINT = 1) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nDim* | The dimension of the array. |

#### Return value

Returns the lower bound for the specified dimension of the safe array for variants of type VT_ARRAY; returns 0 otherwise.

# <a name="GetUBound"></a>GetUBound

Gets the upper bound for the specified dimension of the safe array.

```
FUNCTION GetUBound (BYVAL nDim AS UINT = 1) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nDim* | The dimension of the array. |

#### Return value

Returns the upper bound for the specified dimension of the safe array for variants of type VT_ARRAY; returns 0 otherwise.

# <a name="GetElementCount"></a>GetElementCount

Gets the number of elements in the array.

```
FUNCTION GetElementCount () AS ULONG
```

#### Return value

Returns the number of elements for variants of type VT_ARRAY; returns 1 otherwise.


# <a name="DecToCY"></a>DecToCY

Converts a CVAR of type decimal to a CY structure.

```
FUNCTION DecToCY () AS CY
```

#### Return value

Returns the contents of a VT_DECIMAL variant as a CY structure.

# <a name="DecToDouble"></a>DecToDouble

Converts a CVAR of type decimal to a double.

```
FUNCTION DecToDouble () AS DOUBLE
```

#### Return value

Returns the contents of a VT_DECIMAL variant as a DOUBLE.

# <a name="Round"></a>Round

Rounds a variant to the specified number of decimal places.

```
FUNCTION Round (BYREF cv AS CVAR, BYVAL cDecimals AS LONG) AS CVAR
```

| Parameter  | Description |
| ---------- | ----------- |
| *cv* | The CVAR to round. |
| *cDecimals* | The number of decimal places. |

#### Return value

A CVAR containing the rounded result.

# <a name="FormatNumber"></a>FormatNumber

Formats a CVAR containing numbers into a string form.

```
FUNCTION FormatNumber (BYVAL iNumDig AS LONG = -1, BYVAL ilncLead AS LONG = -2, _
   BYVAL iUseParens AS LONG = -2, BYVAL iGroup AS LONG = -2, BYVAL dwFlags AS DWORD = 0) AS CBSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *iNumDig* | The number of digits to pad to after the decimal point. Specify -1 to use the system default value. |
| *ilncLead* | Specifies whether to include the leading digit on numbers.<br>-2 : Use the system default.<br>-1 : Include the leading digit.<br> 0 : Do not include the leading digit. |
| *iUseParens* | Specifies whether negative numbers should use parentheses.<br>-2 : Use the system default.<br>-1 : Use parentheses.<br>0 : Do not use parentheses. |
| *iGroup* | Specifies whether thousands should be grouped. For example 10,000 versus 10000.<br>-2 : Use the system default.<br>-1 : Group thousands.<br> 0 : Do not group thousands. |
| *dwFlags* | VAR_CALENDAR_HIJRI is the only flag that can be set. |

#### Return value

A CWSTR containing the formatted value.

#### Remarks

This function uses the user's default locale while calling VarTokenizeFormatString and VarFormatFromTokens.

# <a name="GetBooleanElem"></a>GetBooleanElem

Extracts a single boolean element from a safe array of booleans.

```
FUNCTION GetBooleanElem (BYVAL iElem AS ULONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *iElem* | The index of the element of the array. |

#### Return value

The retrieved value.

# <a name="GetDoubleElem"></a>GetDoubleElem

Extracts a single DOUBLE element from a safe array of doubles.

```
FUNCTION GetDoubleElem (BYVAL iElem AS ULONG) AS DOUBLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *iElem* | The index of the element of the array. |

#### Return value

The retrieved value.

# <a name="GetLongElem"></a>GetLongElem

Extracts a single LONG element from a safe array of longs.

```
FUNCTION GetLongElem (BYVAL iElem AS ULONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *iElem* | The index of the element of the array. |

#### Return value

The retrieved value.

# <a name="GetLongIntElem"></a>GetLongIntElem

Extracts a single LONGINT element from a safe array of long integers.

```
FUNCTION GetLongIntElem (BYVAL iElem AS ULONG) AS LONGINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *iElem* | The index of the element of the array. |

#### Return value

The retrieved value.

# <a name="GetShortElem"></a>GetShortElem

Extracts a single SHORT element from a safe array of short integers.

```
FUNCTION GetShortElem (BYVAL iElem AS ULONG) AS SHORT
```

| Parameter  | Description |
| ---------- | ----------- |
| *iElem* | The index of the element of the array. |

#### Return value

The retrieved value.

# <a name="GetStringElem"></a>GetStringElem

Extracts a single BSTR element from a safe array of unicode strings.

```
FUNCTION GetStringElem (BYVAL iElem AS ULONG) AS BSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *iElem* | The index of the element of the array. |

#### Return value

The retrieved value.

# <a name="GetULongElem"></a>GetULongElem

Extracts a single ULONG element from a safe array of unsigned longs.

```
FUNCTION GetULongElem (BYVAL iElem AS ULONG) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *iElem* | The index of the element of the array. |

#### Return value

The retrieved value.

# <a name="GetULongIntElem"></a>GetULongIntElem

Extracts a single ULONGINT element from a safe array of unsigned long integers.

```
FUNCTION GetULongIntElem (BYVAL iElem AS ULONG) AS ULONGINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *iElem* | The index of the element of the array. |

#### Return value

The retrieved value.

# <a name="GetUShortElem"></a>GetUShortElem

Extracts a single USHORT element from a safe array of unsigned shorts.

```
FUNCTION GetUShortElem (BYVAL iElem AS ULONG) AS USHORT
```

| Parameter  | Description |
| ---------- | ----------- |
| *iElem* | The index of the element of the array. |

#### Return value

The retrieved value.

# <a name="GetVariantElem"></a>GetVariantElem

Extracts a single Variant element from a safe array of variants.

```
FUNCTION GetVariantElem (BYVAL iElem AS ULONG) AS CVAR
```

| Parameter  | Description |
| ---------- | ----------- |
| *iElem* | The index of the element of the array. |

#### Return value

The retrieved value.

# <a name="Put"></a>Put

Assigns values to a CVAR.

```
SUB Put (BYREF wsz AS WSTRING)
SUB Put (BYREF cws AS CWSTR)
SUB Put (BYREF cbs AS CBSTR)
FUNCTION Put (BYREF cv AS CVAR) AS HRESULT
FUNCTION Put (BYVAL v AS VARIANT) AS HRESULT
FUNCTION Put (BYVAL pvar AS VARIANT PTR) AS HRESULT
FUNCTION Put (BYREF pDisp AS IDispatch PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
FUNCTION Put (BYREF pUnk AS IUnknown PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
SUB Put (BYVAL _value AS LONGINT, BYVAL _vType AS WORD = VT_I4)
SUB Put (BYVAL _value AS DOUBLE, BYVAL _vType AS WORD = VT_R8)
SUB Put (BYVAL _value AS LONGINT, BYREF strType AS STRING)
SUB Put (BYVAL _value AS DOUBLE, BYREF strType AS STRING)
SUB Put (BYVAL _pv AS ANY PTR, BYVAL _vType AS WORD)
SUB Put (BYVAL _pv AS ANY PTR, BYREF strType AS STRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wsz* | An unicode string. You can also pass a Free Basic ansi string or a string literal. |
| *cws* | A CWSTR variable. |
| *cbs* | A CBSTR variable. |
| *cv* | A CVAR variable. |
| *v* | A VARIANT variable. |
| *v* | A VARIANT variable. |
| *pvar* | Pointer to a VARIANT variable. |
| *pDisp* | Pointer to a DISPATCH interface. |
| *pUnk* | Pointer to a UNKNOWN interface. |
| *_pv* | Pointer to a variable. This will create a VT_BYREF variant of the specified type. |
| *_vtype* | The variant type, e.g. VT_I4, VT_UI4. |
| *strType* | The variant type as a string: "BOOL", "BYTE", "UBYTE", "SHORT", "USHORT, "INT", UINT", "LONG", "ULONG", "LONGINT", "SINGLE, "DOUBLE", "NULL". |
| *fAddRef* | TRUE or FALSE. If TRUE, increases the reference count of the interface. |

# <a name="PutNull"></a>PutNull

Assigns a null value to the CVAR.

```
SUB PutNull
```

# <a name="PutBool"></a>PutBool

Assigns a boolean value to the CVAR.

```
SUB PutBool (BYVAL _value AS BOOL)
```

# <a name="PutBoolean"></a>PutBoolean

Assigns a boolean value to the CVAR.

```
SUB PutBoolean (BYVAL _value AS BOOLEAN)
```

# <a name="PutByte"></a>PutByte

Assigns a byte value to the CVAR.

```
SUB PutByte (BYVAL _value AS BYTE)
```

# <a name="PutUByte"></a>PutUByte

Assigns an unsigned ubyte value to the CVAR.

```
SUB PutUByte (BYVAL _value AS UBYTE)
```

# <a name="PutShort"></a>PutShort

Assigns a short integer value to the CVAR.

```
SUB PutShort (BYVAL _value AS SHORT)
```

# <a name="PutUShort"></a>PutUShort

Assigns an unsigned short integer value to the CVAR.

```
SUB PutUShort (BYVAL _value AS USHORT)
```

# <a name="PutInt"></a>PutInt

Assigns an INT_ (long) value to the CVAR.

```
SUB PutInt (BYVAL _value AS INT_)
```

### Remark

Don't confuse an INT_ (LONG) with the Free Basic INTEGER data type.

# <a name="PutUInt"></a>PutUInt

Assigns an UINT (unsigned long) value to the CVAR.

```
SUB PutUInt (BYVAL _value AS UINT)
```

### Remark

Don't confuse an UINT (ULONG) with the Free Basic UINTEGER data type.

# <a name="PutLong"></a>PutLong

Assigns a LONG value to the CVAR.

```
SUB PutLong (BYVAL _value AS LONG)
```

# <a name="PutULong"></a>PutULong

Assigns a ULONG value to the CVAR.

```
SUB PutULong (BYVAL _value AS ULONG)
```

# <a name="PutLongInt"></a>PutLongInt

Assigns a LONGINT value to the CVAR.

```
SUB PutULong (BYVAL _value AS LONGINT)
```

# <a name="PutULongInt"></a>PutULongInt

Assigns a ULONGINT value to the CVAR.

```
SUB PutULongInt (BYVAL _value AS ULONGINT)
```

# <a name="PutSingle"></a>PutSingle

Assigns a SINGLE value to the CVAR.

```
SUB PutSingle (BYVAL _value AS SINGLE)
```

# <a name="PutFloat"></a>PutFloat

Assigns a SINGLE value to the CVAR.

```
SUB PutFloat (BYVAL _value AS SINGLE)
```

# <a name="PutDouble"></a>PutDouble

Assigns a DOUBLE value to the CVAR.

```
SUB PutDouble (BYVAL _value AS DOUBLE)
```

# <a name="PutBooleanArray"></a>PutBooleanArray

Initializes CVAR from an array of Boolean values.

```
FUNCTION PutBooleanArray (BYVAL prgf AS WINBOOL PTR, BYVAL cElems AS ULONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgf* | Pointer to source array of Boolean values. |
| *cElems* | The number of elements in the array. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_ARRAY OR VT_BOOL variant.

# <a name="PutShortArray"></a>PutShortArray

Initializes CVAR from an array of signed 16-bit integer values.

```
FUNCTION PutShortArray (BYVAL prgf AS SHORT PTR, BYVAL cElems AS ULONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgf* | Pointer to source array of SHORT values. |
| *cElems* | The number of elements in the array. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_ARRAY OR VT_I2 variant.

# <a name="PutUShortArray"></a>PutUshortArray

Initializes CVAR from an array of 16-bit unsigned integer values.

```
FUNCTION PutUshortArray (BYVAL prgf AS USHORT PTR, BYVAL cElems AS ULONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgf* | Pointer to source array of USHORT values. |
| *cElems* | The number of elements in the array. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_ARRAY OR VT_UI2 variant.

# <a name="PutLongArray"></a>PutLongArray

Initializes CVAR from an array of signed 32-bit integer values.

```
FUNCTION PutLongArray (BYVAL prgf AS LONG PTR, BYVAL cElems AS ULONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgf* | Pointer to source array of LONG values. |
| *cElems* | The number of elements in the array. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_ARRAY OR VT_I4 variant.

# <a name="PutULongArray"></a>PutULongArray

Initializes CVAR from an array of 32-bit unsigned integer values.

```
FUNCTION PutULongArray (BYVAL prgf AS ULONG PTR, BYVAL cElems AS ULONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgf* | Pointer to source array of ULONG values. |
| *cElems* | The number of elements in the array. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_ARRAY OR VT_UI4 variant.

# <a name="PutLongIntArray"></a>PutLongIntArray

Initializes CVAR from an array of signed 64-bit integer values.

```
FUNCTION PutLongIntArray (BYVAL prgf AS LONGINT PTR, BYVAL cElems AS ULONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgf* | Pointer to source array of LONGINT values. |
| *cElems* | The number of elements in the array. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_ARRAY OR VT_I8 variant.

# <a name="PutULongIntArray"></a>PutULongIntArray

Initializes CVAR from an array of unsigned 64-bit integer values.

```
FUNCTION PutULongIntArray (BYVAL prgf AS ULONGINT PTR, BYVAL cElems AS ULONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgf* | Pointer to source array of ULONGINT values. |
| *cElems* | The number of elements in the array. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_ARRAY OR VT_UI8 variant.

# <a name="PutDoubleArray"></a>PutDoubleArray

Initializes CVAR from an array of unsigned 64-bit integer values.

```
FUNCTION PutDoubleArray (BYVAL prgf AS DOUBLE PTR, BYVAL cElems AS ULONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgf* | Pointer to source array of DOUBLE values. |
| *cElems* | The number of elements in the array. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_ARRAY OR VT_R8 variant.

# <a name="PutStringArray"></a>PutStringArray

Initializes CVAR from an array of unicode strings.

```
FUNCTION PutStringArray (BYVAL prgsz AS PCWSTR, BYVAL cElems AS ULONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgsz* | Pointer to source array of unicode strings. |
| *cElems* | The number of elements in the array. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_ARRAY OR VT_BSTR variant.

# <a name="PutBuffer"></a>PutBuffer

Initializes CVAR with the contents of a buffer.

```
FUNCTION PutBuffer (BYVAL pv AS VOID PTR, BYVAL cb AS UINT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pv* | Pointer to the source buffer. |
| *cb* | The length of the buffer, in bytes. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_ARRAY OR VT_UI1 variant.

# <a name="PutDateString"></a>PutDateString

Initializes CVAR VT_DATE from a string.

```
FUNCTION PutDateString (BYVAL pwszDate AS WSTRING PTR, BYVAL lcid AS LCID = 0, _
   BYVAL dwFlags AS ULONG = 0) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszDate* | The date value to convert, e.g. "2018-08-20 19:42". |
| *lcid* | The locale identifier. |
| *dwFlags* | One or more of the following flags.<br>*LOCALE_NOUSEROVERRIDE* : Uses the system default locale settings, rather than custom locale settings.<br>*VAR_CALENDAR_HIJRI* : If set then the Hijri calendar is used. Otherwise the calendar set in the control panel is used.<br>*VAR_TIMEVALUEONLY* : Omits the date portion of a VT_DATE and returns only the time. Applies to conversions to or from dates.<br>*VAR_DATEVALUEONLY* : Omits the time portion of a VT_DATE and returns only the date. Applies to conversions to or from dates. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_DATE variant.

# <a name="PutDec"></a>PutDec

Initializes CVAR with the contents of a DECIMAL structure.

```
FUNCTION PutDec (BYCAL dec AS DECIMAL) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *dec* | A DECIMAL structure. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_DECIMAL OR VT_BYREF variant.

# <a name="PutDecFromCY"></a>PutDecFromCY

Converts a currency value to a variant of type VT_DECIMAL.

```
FUNCTION PutDecFromCY (BYVAL cyIn AS CY) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cyIn* | The currency value to convert. |

#### Return value

This function can return one of these values.

| Value      | Meaning     |
| ---------- | ----------- |
| *S_OK* | Success. |
| *DISP_E_TYPEMISMATCH* | The argument could not be coerced to the specified type. |
| *E_INVALIDARG* | One of the arguments is not valid. |
| *E_OUTOFMEMORY* | Insufficient memory to complete the operation. |

#### Remarks

Creates a VT_DECIMAL variant.

# <a name="PutDecFromDouble"></a>PutDecFromDouble

Converts a double value to a variant of type VT_DECIMAL.

```
FUNCTION PutDecFromDouble (BYVAL dbIn AS DOUBLE) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *dbIn* | The DOUBLE value to convert. |

#### Return value

This function can return one of these values.

| Value      | Meaning     |
| ---------- | ----------- |
| *S_OK* | Success. |
| *DISP_E_TYPEMISMATCH* | The argument could not be coerced to the specified type. |
| *E_INVALIDARG* | One of the arguments is not valid. |
| *E_OUTOFMEMORY* | Insufficient memory to complete the operation. |

#### Remarks

Creates a VT_DECIMAL variant.


# <a name="PutDecFromStr"></a>PutDecFromStr

Initializes CVAR as VT_DECIMAL from a string.

```
FUNCTION PutDecFromStr (BYVAL pwszIn AS WSTRING PTR, BYVAL lcid AS LCID = 0, _
   BYVAL dwFlags AS ULONG = 0) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszIn* | The string value to convert. |
| *lcid* | The locale identifier. |
| *dwFlags* | One or more of the following flags.<br>*LOCALE_NOUSEROVERRIDE* : Uses the system default locale settings, rather than custom locale settings.<br>*VAR_TIMEVALUEONLY* : Omits the date portion of a VT_DATE and returns only the time. Applies to conversions to or from dates.<br>*VAR_DATEVALUEONLY* : Omits the time portion of a VT_DATE and returns only the date. Applies to conversions to or from dates. |

#### Return value

This function can return one of these values.

| Value      | Meaning     |
| ---------- | ----------- |
| *S_OK* | Success. |
| *DISP_E_TYPEMISMATCH* | The argument could not be coerced to the specified type. |
| *E_INVALIDARG* | One of the arguments is not valid. |
| *E_OUTOFMEMORY* | Insufficient memory to complete the operation. |

#### Remarks

Creates a VT_DECIMAL variant.

# <a name="PutFileTime"></a>PutFileTime

Initializes CVAR with the contents of a FILETIME structure.

```
FUNCTION PutFileTime (BYVAL pft AS FILETIME PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pft* | Pointer to a FILETIME structure. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_DATE variant.

# <a name="PutFileTimeArray"></a>PutFileTimeArray

Initializes CVAR with an array of FILETIME structures.

```
FUNCTION PutFileTimeArray (BYVAL prgft AS FILETIME PTR, BYVAL cElems AS ULONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgft* | Pointer to an array of FILETIME structures. |
| *cElems* | The number of elements in the array pointed to by *prgft*. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_ARRAY OR VT_DATE variant.

# <a name="PutGuid"></a>PutGuid

Initializes CVAR from a GUID.

```
FUNCTION PutGuid (BYVAL rguid AS GUID PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *rguid* | Reference to the source GUID. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_ARRAY OR VT_UI1 variant.

# <a name="PutPropVariant"></a>PutPropVariant

Initializes CVAR from the contents of a PROPVARIANT structure.

```
FUNCTION PutPropVariant (BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pPropVar* | Pointer to a source PROPVARIANT structure. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Normally, the data stored in the PROPVARIANT is copied to the VARIANT without a datatype change. However, in the following cases, there is no direct VARIANT support for the datatype, and they are converted as shown.

VT_BLOB, VT_STREAM<br>
Converted to VT_UNKNOWN. The punkVal member will contain a pointer to an IStream that contains the source data.

VT_LPSTR, VT_LPWSTR, VT_CLSID<br>
Converted to VT_BSTR,

VT_FILETIME<br>
Converted to VT_DATE.

VT_VECTOR OR x<br>
Converted to VT_ARRAY OR x

The following types cannot be converted with this function.

VT_STORAGE<br>
VT_BLOB_OBJECT<br>
VT_STREAMED_OBJECT<br>
VT_STORED_OBJECT<br>
VT_CF<br>
VT_VECTOR OR VT_CF

# <a name="PutRecord"></a>PutRecord

Initializes CVAR with a reference to an UDT.

```
FUNCTION PutRecord (BYVAL pIRecordInfo AS IRecordInfo PTR, BYVAL pRec AS VOID PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pIRecordInfo* | Pointer to the IRecordInfo interface. |
| *pRec* | Pointer to the UDT. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_RECORD variant.

# <a name="PutRef"></a>PutRef

Assigns a value by reference (a pointer to a variable).

```
FUNCTION PutRef (BYVAL _pvar AS ANY PTR, BYVAL _vType AS WORD) AS HRESULT
FUNCTION PutRef (BYVAL _pvar AS ANY PTR, BYREF strType AS STRING) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *_pvar* | Pointer to a variable. |
| *_vType* | Type of the variant: VT_BOOL, VT_I1, VT_UI1, VT_I2, VT_UI2, VT_INT, VT_UINT, VT_I4, VT_UI4, VT_I8, VT_UI8, VT_R4, VT_R8, VT_BSTR, VT_UNKNOWN, VT_DISPATCH, VT_DECIMAL, VT_CY, VT_DATE, VT_VARIANT, VT_SAFEARRAY, VT_ERROR. |
| *strType* | Type of the variant: "BOOL", "BYTE", "UBYTE", "SHORT", "USHORT", "INT", "UINT", "LONG", "ULONG", "LONGINT", "ULONGINT", "SINGLE", "DOUBLE", "BSTR", "UNKNOWN", "DISPATCH", "DECIMAL", "CY", "DATE", "VARIANT", "SAFEARRAY", "ERROR". |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_BYREF variant of the sepecified type.

# <a name="PutResource"></a>PutResource

Initializes the CVAR based on a string resource imbedded in an executable file.

```
FUNCTION PutResource (BYVAL hinst AS HINSTANCE, BYVAL id AS UINT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *hinst* | The instance handle. |
| *id* | Integer identifier of the string to be loaded. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_BSTR variant. If the resource does not exist, this function initializes the VARIANT as VT_EMPTY and returns a failure code.

# <a name="PutSafeArray"></a>PutSafeArray

Initializes CVAR from a safe array.

```
FUNCTION PutSafeArray (BYVAL parray AS SAFEARRAY PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *parray* | Pointer to safe array. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_ARRAY variant.

# <a name="PutStrRet"></a>PutStrRet

Initializes CVAR with the string stored in a STRRET structure.

```
FUNCTION PutStrRet (BYVAL pstrret AS STRRET PTR, BYVAL pidl AS PCUITEMID_CHILD) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pstrret* | Pointer to a STRRET structure. |
| *pidl* | PIDL of the item whose details are being retrieved. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_BSTR variant. This function frees the resources used for the STRRET contents.


# <a name="PutSystemTime"></a>PutSystemTime

Initializes CVAR with the contents of a SYSTEMTIME structure.

```
FUNCTION PutSystemTime (BYVAL pst AS SYSTEMTIME PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pst* | Pointer to a SYSTEMTIME structure. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_DATE variant.


# <a name="PutUtf8"></a>PutUtf8

Initializes CVAR with the contents of an UTF-8 string.

```
FUNCTION PutUtf8 (BYREF strUtf8 AS STRING) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *strUtf8* | The UTF-8 encoded string. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_BSTR variant.

# <a name="PutVariantArrayElem"></a>PutVariantArrayElem

Initializes CVAR with a value stored in another VARIANT structure.

```
FUNCTION PutVariantArrayElem (BYVAL pvarIn AS VARIANT PTR, BYVAL iElem AS ULONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pvarIn* | Reference to the source VARIANT structure. |
| *iElem* | Index of one of the source VARIANT structure elements. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

This helper function works for VARIANT structures of the following types: 

    VT_BSTR
    VT_BOOL
    VT_I2
    VT_I4
    VT_I8
    VT_U12
    VT_U14
    VT_U18
    VT_DATE
    VT_ARRAY | (any one of VT_BSTR, VT_BOOL, VT_I2, VT_I4, VT_I8, VT_U12, VT_U14, VT_U18, VT_DATE)

Additional types may be supported in the future.

This function extracts a single value from the source VARIANT structure and uses that value to initialize the output VARIANT structure. The calling application must use VariantClear to free the VARIANT referred to by pvar when it is no longer needed.

If the source VARIANT is an array, *iElem* must be less than the number of elements in the array.

If the source VARIANT has a single value, *iElem* must be 0.

If the source VARIANT is empty, this function always returns an error code.

You can use **GetElementCount** to obtain the number of elements in the array or array.

# <a name="PutVbDate"></a>PutVbDate

Initializes CVAR with the contents of a DATE value.

```
FUNCTION PutVbDate (BYREF vbDate AS DATE_) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *vbDate* | The DATE value. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

Creates a VT_DATE variant.

# <a name="ToBooleanArray"></a>ToBooleanArray

Extracts an array of boolean values from CVAR.

```
FUNCTION ToBooleanArray (BYVAL prgf AS WINBOOL PTR, BYVAL crgn AS ULONG) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgf* | Pointer to a buffer that contains *crgn* boolean values. When this function returns, the buffer has been initialized with elements extracted from the source VARIANT structure. |
| *crgn* | The number of elements in the buffer pointed to by *prgf*. |

#### Return value

The count of WINBOOL elements extracted from the CVAR.

#### Remarks

This helper function is used when the calling application expects a VARIANT to hold an array that consists of a fixed number of boolean values.

If the source VARIANT is of type VT_ARRAY OR VT_BOOL, this function extracts up to *crgn* WINBOOL values and places them into the buffer pointed to by *prgf*. If the VARIANT contains more elements than will fit into the *prgf* buffer, this function returns 0.

# <a name="ToBooleanArrayAlloc"></a>ToBooleanArrayAlloc

Extracts an array of boolean values from CVAR.

```
FUNCTION ToBooleanArrayAlloc (BYVAL pprgf AS WINBOOL PTR PTR) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgf* | Pointer to a WINBOOL PTR variable that will recive a pointer to an array of WINBOOL values extracted from the source CVAR. |

#### Return value

The count of WINBOOL elements extracted from the CVAR.

#### Remarks

This helper function is used when the calling application expects a CVAR to hold an array of WINBOOL values.

If CVAR is of type VT_ARRAY OR VT_BOOL, this function extracts an array of WINBOOL values into a newly allocated array. The calling application is responsible for using **CoTaskMemFree** to release the array pointed to by *pprgf* when it is no longer needed.

# <a name="ToBstr"></a>ToBstr

Extracts the content of the underlying variant and returns it as a CBSTR.

```
FUNCTION ToBstr () AS CBSTR
```

#### Return value

The contents of the variant as a CBSTR.

#### Example

```
DIM cv AS CVAR = "Test string"
DIM cbs AS CBSTR = cv.ToBstr
```

# <a name="ToBuffer"></a>ToBuffer

Extracts the contents of a CVAR of type VT_ARRRAY OR VT_UI1 to a buffer.

```
FUNCTION ToBuffer (BYVAL pv AS VOID PTR, BYVAL cb AS UINT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pv* | Pointer to a buffer of length *cb* bytes. When this function returns, contains the first *cb* bytes of the extracted buffer value. |
| *cb* | The size of the *pv* buffer, in bytes. The buffer should be the same size as the data to be extracted, or smaller. |

#### Return value

Returns one of the following values:

| Value      | Meaning |
| ---------- | ----------- |
| *S_OK* | Data successfully extracted. |
| *E_INVALIDARG* | The VARIANT was not of type VT_ARRRAY OR VT_UI1. |
| *E_FAIL* |The VARIANT buffer value had fewer than cb bytes. |

#### Remarks

This function is used when the calling application expects a VARIANT to hold a buffer value. The calling application should check that the value has the expected length before it calls this function.

If CVAR has type VT_ARRAY OR VT_UI1, this function extracts the first *cb* bytes from the structure and places them in the buffer pointed to by *pv*.

If the stored value has fewer than *cb* bytes, then function fails and the buffer is not modified.

If the value has more than *cb* bytes, then function succeeds and truncates the value.

To retrieve the size of the array call **GetElementCount**.

# <a name="ToBufferString"></a>ToBuffer (STRING)

Extracts the contents of a CVAR of type VT_ARRRAY OR VT_UI1 to a string used as a buffer.

```
FUNCTION ToBuffer () AS STRING
```

#### Return value

A string with the contents of the array.

# <a name="ToDosDateTime"></a>ToDosDateTime

Extracts a date and time value in Microsoft MS-DOS format from a CVAR of type VT_DATE.

```
FUNCTION ToDosDateTime (BYVAL pwDate AS USHORT PTR, BYVAL pwTime AS USHORT PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwDate* | When this function returns, contains the extracted USHORT that represents a MS-DOS date. |
| *pwTime* | When this function returns, contains the extracted contains the extracted WORD that represents a MS-DOS time. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

This helper function is used when the calling application expects a CVAR to hold a datetime value.

If CVAR is of type VT_DATE, this function extracts the datetime value.

If CVAR is not of type VT_DATE, the function attempts to convert the value in the VARIANT structure into the right format. If a conversion is not possible, it returns a failure code.

# <a name="ToDoubleArray"></a>ToDoubleArray

Extracts an array of DOUBLE values from CVAR.

```
FUNCTION ToDoubleArray (BYVAL prgn AS DOUBLE PTR, BYVAL crgn AS ULONG) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgn* | Pointer to a buffer that contains crgn DOUBLE variables. When this function returns, the buffer has been initialized with DOUBLE elements extracted from CVAR. |
| *crgn* | The number of elements in the buffer pointed to by *prgn*. |

#### Return value

This helper function is used when the calling application expects a VARIANT to hold an array that consists of a fixed number of DOUBLE values.

If the source VARIANT is of type VT_ARRAY OR VT_R8, this function extracts up to *crgn* DOUBLE values and places them into the buffer pointed to by *prgn*. If the VARIANT contains more elements than will fit into the *prgn* buffer, this function returns 0.

# <a name="ToDoubleArrayAlloc"></a>ToDoubleArrayAlloc

Extracts an array of DOUBLE values from CVAR.

```
FUNCTION ToDoubleArrayAlloc (BYVAL pprgn AS ULONGINT PTR PTR) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pprgf* | Pointer to a DOUBLE PTR variable that will recive a pointer to an array of DOUBLE values extracted from the source CVAR. |

#### Return value

The count of DOUBLE elements extracted from the CVAR.

#### Remarks

This helper function is used when the calling application expects a CVAR to hold an array of DOUBLE values.

If CVAR is of type VT_ARRAY OR VT_R8, this function extracts an array of DOUBLE values into a newly allocated array. The calling application is responsible for using **CoTaskMemFree** to release the array pointed to by *pprgn* when it is no longer neede

# <a name="ToFileTime"></a>ToFileTime

Returns the contents of a CVAR of type VT_DATE as a FILETIME structure.

```
FUNCTION ToFileTime (BYVAL stfOut AS AFX_PSTIME_FLAGS) AS FILETIME
```

| Parameter  | Description |
| ---------- | ----------- |
| *stfOut* | Specifies one of the following time flags:<br>*PSTF_UTC (0)* : Indicates coordinated universal time.<br>*PSTF_LOCAL (1)* : Indicates local time. |

#### Return value

A FILETIME structure.

# <a name="ToGuid"></a>ToGuid

Returns the contents of a CVAR containing a GUID string as a GUID structure.

```
FUNCTION ToGuid () AS GUID
```

#### Return value

Returns the contents of a CVAR containing a GUID string as an unicode GUID string.

# <a name="ToGuidBStr"></a>ToGuidBStr

Returns the contents of a CVAR containing a GUID string as an unicode GUID string.

```
FUNCTION ToGuidBStr () AS CBSTR
```

#### Return value

A GUID string.

# <a name="ToGuidStr"></a>ToGuidStr

Returns the contents of a CVAR containing a GUID string as an unicode GUID string.

```
FUNCTION ToGuidStr () AS CWSTR
```

#### Return value

A GUID string.

# <a name="ToGuidWStr"></a>ToGuidWStr

Returns the contents of a CVAR containing a GUID string as an unicode GUID string.

```
FUNCTION ToGuidWStr () AS CWSTR
```

#### Return value

A GUID string.

# <a name="ToLongArray"></a>ToLongArray

Extracts an array of LONG values from CVAR.

```
FUNCTION ToLongArray (BYVAL prgn AS LONG PTR, BYVAL crgn AS ULONG) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgn* | Pointer to a buffer that contains *crgn* LONG variables. When this function returns, the buffer has been initialized with Int32 elements extracted from CVAR. |
| *crgn* | The number of elements in the buffer pointed to by *prgn*. |

#### Return value

This helper function is used when the calling application expects a VARIANT to hold an array that consists of a fixed number of Int32 values.

If the source VARIANT is of type VT_ARRAY OR VT_I4, this function extracts up to *crgn* Int32 values and places them into the buffer pointed to by *prgn*. If the VARIANT contains more elements than will fit into the *prgn* buffer, this function returns 0.

# <a name="ToLongArrayAlloc"></a>ToLongArrayAlloc

Extracts an array of LONG values from CVAR.

```
FUNCTION ToLongArrayAlloc (BYVAL pprgn AS LONG PTR PTR) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pprgn* | Pointer to a LONG PTR variable that will recive a pointer to an array of LONG values extracted from the source CVAR. |

#### Return value

The count of LONG elements extracted from the CVAR.

#### Remarks

This helper function is used when the calling application expects a CVAR to hold an array of LONG values.

If CVAR is of type VT_ARRAY OR VT_I4, this function extracts an array of LONG values into a newly allocated array. The calling application is responsible for using **CoTaskMemFree** to release the array pointed to by *pprgn* when it is no longer needed.

# <a name="ToLongIntArray"></a>ToLongIntArray

Extracts an array of LONGINT values from CVAR.

```
FUNCTION ToLongIntArray (BYVAL prgn AS LONGINT PTR, BYVAL crgn AS ULONG) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgn* | Pointer to a buffer that contains *crgn* LONGINT variables. When this function returns, the buffer has been initialized with Int64 elements extracted from CVAR. |
| *crgn* | The number of elements in the buffer pointed to by *prgn*. |

#### Return value

This helper function is used when the calling application expects a VARIANT to hold an array that consists of a fixed number of Int64 values.

If the source VARIANT is of type VT_ARRAY OR VT_UI4, this function extracts up to *crgn* Int64 values and places them into the buffer pointed to by *prgn*. If the VARIANT contains more elements than will fit into the *prgn* buffer, this function returns 0.

# <a name="ToLongIntArrayAlloc"></a>ToLongIntArrayAlloc

Extracts an array of LONGINT values from CVAR.

```
FUNCTION ToLongIntArrayAlloc (BYVAL pprgn AS LONGINT PTR PTR) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pprgn* | Pointer to a LONGINT PTR variable that will recive a pointer to an array of LONGINT values extracted from the source CVAR.
 |

#### Return value

The count of LONGINT elements extracted from the CVAR.

#### Remarks

This helper function is used when the calling application expects a CVAR to hold an array of LONGINT values.

If CVAR is of type VT_ARRAY OR VT_I8, this function extracts an array of LONGINT values into a newly allocated array. The calling application is responsible for using **CoTaskMemFree** to release the array pointed to by *pprgn* when it is no longer needed.

# <a name="ToShortArray"></a>ToShortArray

Extracts an array of Int16 values from CVAR.

```
FUNCTION ToShortArray (BYVAL prgn AS SHORT PTR, BYVAL crgn AS ULONG) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgn* | Pointer to a buffer that contains *crgn* Int16 variables. When this function returns, the buffer has been initialized with Int16 elements extracted from CVAR. |
| *crgn* | The number of elements in the buffer pointed to by *prgn*. |

#### Return value

The count of Int16 elements extracted from the CVAR.

#### Remarks

This helper function is used when the calling application expects a VARIANT to hold an array that consists of a fixed number of Int16 values.

If the source VARIANT is of type VT_ARRAY OR VT_I2, this function extracts up to *crgn* Int16 values and places them into the buffer pointed to by *prgn*. If the VARIANT contains more elements than will fit into the *prgn* buffer, this function returns 0.

# <a name="ToShortArrayAlloc"></a>ToShortArrayAlloc

Extracts an array of SHORT values from CVAR.

```
FUNCTION ToShortArrayAlloc (BYVAL pprgn AS SHORT PTR PTR) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pprgn* | Pointer to a SHORT PTR variable that will recive a pointer to an array of SHORT values extracted from the source CVAR. |

#### Return value

The count of SHORT elements extracted from the CVAR.

#### Remarks

This helper function is used when the calling application expects a CVAR to hold an array of SHORT values.

If CVAR is of type VT_ARRAY OR VT_I2, this function extracts an array of SHORT values into a newly allocated array. The calling application is responsible for using **CoTaskMemFree** to release the array pointed to by *pprgn* when it is no longer needed.

# <a name="ToStr"></a>ToStr

Extracts the content of the underlying variant and returns it as a CWSTR.

```
FUNCTION ToStr () AS CWSTR
```

#### Return value

The contents of the variant as a CWSTR.

#### Example

```
DIM cv AS CVAR = "Test string"
DIM cws AS CWSTR = cv.ToStr
```

# <a name="ToStringArray"></a>ToStringArray

Extracts data from a vector structure into a PWSTR array.

```
FUNCTION ToStringArray (BYVAL prgsz AS PWSTR, BYVAL crgsz AS ULONG) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgsz* | Pointer to a buffer that contains *crgn* PWSTR values. When this function returns, the buffer has been initialized with elements extracted from the source VARIANT structure. |
| *crgsz* | The number of elements in the buffer pointed to by *prgsz*. |

#### Return value

The count of PWSTR elements extracted from the CVAR.

#### Remarks

This helper function is used when the calling application expects a CVAR to hold an array of PWSTR values. If the VARIANT contains more elements than will fit into the *prgsz* buffer, this function returns 0.

# <a name="ToStringArrayAlloc"></a>ToStringArrayAlloc

Extracts an array of PWSTR values from CVAR.

```
FUNCTION ToStringArrayAlloc (BYVAL pprgsz AS PWSTR PTR) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pprgsz* | Pointer to a PWSTR PTR variable that will recive a pointer to an array of PWSTR values extracted from the source CVAR. |

#### Return value

The count of PWSTR elements extracted from the CVAR.

#### Remarks

This helper function is used when the calling application expects a CVAR to hold an array of PWSTR values.

This function extracts an array of PWSTR values into a newly allocated array. The calling application is responsible for using **CoTaskMemFree** to free the memory used by each of the strings and to release the array pointed to by pprgn when it is no longer needed.

# <a name="ToStrRet"></a>ToStrRet

Returns the contents of a CVAR of type VT_BSTR to a STRRET stucture.

```
FUNCTION ToStrRet () AS STRRET
```

#### Return value

A STRRET structure.

# <a name="ToSystemTime"></a>ToSystemTime

Returns the contents of CVAR of type VT_DATE as a FILETIME structure.

```
FUNCTION ToSystemTime () AS SYSTEMTIME
```

#### Return value

A SYSTEMTIME structure.

# <a name="ToULongArray"></a>ToULongArray

Extracts an array of ULONG values from CVAR.

```
FUNCTION ToULongArray (BYVAL prgn AS ULONG PTR, BYVAL crgn AS ULONG) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgn* | Pointer to a buffer that contains *crgn* ULONG variables. When this function returns, the buffer has been initialized with UInt32 elements extracted from CVAR. |
| *crgn* | The number of elements in the buffer pointed to by *prgn*. |

#### Return value

The count of UInt32 elements extracted from the CVAR.

#### Remarks

This helper function is used when the calling application expects a VARIANT to hold an array that consists of a fixed number of UInt32 values.

If the source VARIANT is of type VT_ARRAY OR VT_UI4, this function extracts up to *crgn* UInt32 values and places them into the buffer pointed to by *prgn*. If the VARIANT contains more elements than will fit into the *prgn* buffer, this function returns 0.

# <a name="ToULongArrayAlloc"></a>ToULongArrayAlloc

Extracts an array of ULONG values from CVAR.

```
FUNCTION ToULongArrayAlloc (BYVAL pprgn AS ULONG PTR PTR) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pprgn* | Pointer to a ULONG PTR variable that will recive a pointer to an array of ULONG values extracted from the source CVAR. |

#### Return value

The count of ULONG elements extracted from the CVAR.

#### Remarks

This helper function is used when the calling application expects a CVAR to hold an array of ULONG values.

If CVAR is of type VT_ARRAY OR VT_UI4, this function extracts an array of ULONG values into a newly allocated array. The calling application is responsible for using **CoTaskMemFree** to release the array pointed to by *pprgn* when it is no longer needed.

# <a name="ToULongIntArray"></a>ToULongIntArray

Extracts an array of ULONGINT values from CVAR.

```
FUNCTION ToULongIntArray (BYVAL prgn AS ULONGINT PTR, BYVAL crgn AS ULONG) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgn* | Pointer to a buffer that contains *crgn* ULONGINT variables. When this function returns, the buffer has been initialized with UInt64 elements extracted from CVAR. |
| *crgn* | The number of elements in the buffer pointed to by *prgn*. |

#### Return value

The count of UInt64 elements extracted from the CVAR.

#### Remarks

This helper function is used when the calling application expects a VARIANT to hold an array that consists of a fixed number of UInt64 values.

If the source VARIANT is of type VT_ARRAY OR VT_UI8, this function extracts up to *crgn* UInt64 values and places them into the buffer pointed to by *prgn*. If the VARIANT contains more elements than will fit into the *prgn* buffer, this function returns 0.

# <a name="ToULongIntArrayAlloc"></a>ToULongIntArrayAlloc

Extracts an array of ULONGINT values from CVAR.

```
FUNCTION ToULongIntArrayAlloc (BYVAL pprgn AS ULONGINT PTR PTR) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pprgn* | Pointer to a ULONGINT PTR variable that will recive a pointer to an array of ULONGINT values extracted from the source CVAR. |

#### Return value

The count of ULONGINT elements extracted from the CVAR.

#### Remarks

This helper function is used when the calling application expects a CVAR to hold an array of ULONGINT values.

If CVAR is of type VT_ARRAY OR VT_UI8, this function extracts an array of ULONGINT values into a newly allocated array. The calling application is responsible for using **CoTaskMemFree** to release the array pointed to by *pprgn* when it is no longer needed.

# <a name="ToUShortArray"></a>ToUShortArray

Extracts an array of UInt16 values from CVAR.

```
FUNCTION ToUShortArray (BYVAL prgn AS USHORT PTR, BYVAL crgn AS ULONG) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgn* | Pointer to a buffer that contains *crgn* UInt16 variables. When this function returns, the buffer has been initialized with UInt16 elements extracted from CVAR. |
| *crgn* | The number of elements in the buffer pointed to by *prgn*. |

#### Return value

The count of UInt16 elements extracted from the CVAR.

#### Remarks

This helper function is used when the calling application expects a VARIANT to hold an array that consists of a fixed number of UInt16 values.

If the source VARIANT is of type VT_ARRAY OR VT_I2, this function extracts up to *crgn* UInt16 values and places them into the buffer pointed to by *prgn*. If the VARIANT contains more elements than will fit into the *prgn* buffer, this function returns 0.

# <a name="ToUShortArrayAlloc"></a>ToUShortArrayAlloc

Extracts an array of USHORT values from CVAR.

```
FUNCTION ToUShortArrayAlloc (BYVAL pprgn AS USHORT PTR PTR) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pprgn* | Pointer to a USHORT PTR variable that will recive a pointer to an array of USHORT values extracted from the source CVAR. |

#### Return value

The count of USHORT elements extracted from the CVAR.

#### Remarks

This helper function is used when the calling application expects a CVAR to hold an array of USHORT values.

If CVAR is of type VT_ARRAY OR VT_UI2, this function extracts an array of SHORT values into a newly allocated array. The calling application is responsible for using **CoTaskMemFree** to release the array pointed to by *pprgn* when it is no longer needed.

# <a name="ToUtf8"></a>ToUtf8

Returns the contents of a CVAR containing a BSTR as an UTF-8 encoded string.

```
FUNCTION ToUtf8 () AS STRING
```

#### Return value

The UTF-8 string.

# <a name="ToVbDate"></a>ToVbDate

Returns the contents of a CVAR of type VT_DATE as a DATE value.

```
FUNCTION ToVbDate () AS DATE_
```

#### Return value

A DATE_ value (double).

# <a name="ToWStr"></a>ToWStr

Extracts the content of the underlying variant and returns it as a CWSTR.

```
FUNCTION ToWStr () AS CWSTR
```

#### Return value

The contents of the variant as a CWSTR.

#### Example

```
DIM cv AS CVAR = "Test string"
DIM cws AS CWSTR = cv.ToWStr
```

# <a name="ValDouble"></a>ValDouble

Converts the variant to a floating point number (DOUBLE)

```
FUNCTION ValDouble () AS DOUBLE
```

# <a name="ValInt"></a>ValInt

Converts the variant to a signed 32-bit integer (LONG).

```
FUNCTION ValInt () AS LONG
```

# <a name="ValLong"></a>ValLong

Converts the variant to a signed 32-bit integer (LONG).

```
FUNCTION ValLong () AS LONG
```

# <a name="ValLongInt"></a>ValLongInt

Converts the variant to a signed 64-bit integer (LONGINT).

```
FUNCTION ValLongInt () AS LONGINT
```

# <a name="Value"></a>Value

Converts the variant to a floating point number (DOUBLE).

```
FUNCTION Value () AS DOUBLE
```

# <a name="ValUInt"></a>ValUInt

Converts the variant to a 32.bit unsigned integer (ULONG).

```
FUNCTION ValUInt () AS ULONG
```

# <a name="ValULong"></a>ValULong

Converts the variant to a 32.bit unsigned integer (ULONG).

```
FUNCTION ValULong () AS ULONG
```

# <a name="ValULongInt"></a>ValULongInt

Converts the variant to a 64-bit unsigned integer (ULONGINT).

```
FUNCTION ValULongInt () AS ULONGINT
```

# <a name="AfxCVarToStr"></a>AfxCVarToStr

Extracts the contents of a CVAR to a CWSTR.

```
FUNCTION AfxCVarToStr OVERLOAD (BYREF cv AS CVAR) AS CWSTR
```

# <a name="AfxCVarToStr"></a>AfxCVarToStr

Extracts the contents of a CVAR to a CWSTR.

```
FUNCTION AfxCVarToStr OVERLOAD (BYVAL pcv AS CVAR PTR) AS CWSTR
```

# <a name="AfxCVariantToBuffer"></a>AfxCVariantToBuffer

Extracts the contents of a variant that contains an array of bytes.

```
FUNCTION AfxCVariantToBuffer (BYREF cvIn AS CVAR, BYVAL pv AS LPVOID, BYVAL cb AS ULONG) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pv* | Pointer to a buffer of length *cb* bytes. When this function returns, contains the first *cb* bytes of the extracted buffer value. |
| *cb* | The size of the *pv* buffer, in bytes. The buffer should be the same size as the data to be extracted, or smaller. |

# <a name="AfxCVarOptPrm"></a>AfxCVarOptPrm

Returns a CVAR suitable to be used with optional parameters.

```
FUNCTION AfxCVarOptPrm () AS CVAR
```

#### Remarks

If you want to call a method that has optional variant parameters, you still have to pass the parameters, but in a way That the methods knows that they were omitted. Specifically, you have to pass them as type VT_ERROR, code DISP_E_PARAM_NOT_FOUND.

```
DIM v AS VARIANT = TYPE(VT_ERROR, 0, 0, 0, DISP_E_PARAMNOTFOUND)
```
