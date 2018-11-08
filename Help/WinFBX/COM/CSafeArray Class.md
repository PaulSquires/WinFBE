# CSafeArray Class

**CSafeArray** is a class that provides wrapper methods for the SAFEARRAY structure, making it easy to create and manage single- and multidimensional arrays of almost any of the VARIANT-supported types. The lower bound of a CSafeArray can start at any user-defined value. Additional overloaded methods are provided for one and two-dimensional safe arrays, that are the ones most often used in COM programming.

When the class is destroyed, the array descriptor and all of the data in the array is destroyed. If objects are stored in the array, Release is called on each object in the array. Safe arrays of variant will have VariantClear called on each member and safe arrays of BSTR will have SysFreeString called on each element. IRecordInfo.RecordClear will be called to release object references and other values of a record without deallocating the record.

**Include file**: CSafeArray.INC.

# Structures and VARTYPE

| Name       | Description |
| ---------- | ----------- |
| [SAFEARRAY](#SAFEARRAY) | Represents a safe array. |
| [SAFEARRAYBOUND](#SAFEARRAYBOUND) | Represents the bounds of one dimension of the array. |
| [VARTYPE](#VARTYPE) | Safe array Variant type. |

# Constructors

| Name       | Description |
| ---------- | ----------- |
| [Constructor (SAFEARRAYBOUND)](#Constructor1) | Creates a CSafeArray. |
| [Constructor (CSafeArray)](#Constructor2) | Creates a CSafeArray from another CSafeArray. |
| [Constructor (SAFEARRAY PTR)](#Constructor3) | Creates a CSafeArray from a safe array. |
| [Constructor (VARIANT PTR)](#Constructor4) | Creates a CSafeArray from a Variant of type VT_ARRAY. |

# Operators

| Name       | Description |
| ---------- | ----------- |
| [Operator \*](#Operator1) | Returns a pointer to the safe array descriptor. |
| [Operator LET](#Operator2) | Assigns a CSafeArray, a safe array or a Variant. |

# Methods

| Name       | Description |
| ---------- | ----------- |
| [AccessData](#AccessData) | Increments the lock count of an array, and retrieves a pointer to the array data. |
| [Append](#Append) | Appends a value to the end of the one-dimensional safe array. |
| [Attach](#Attach) | Attaches a safe array to a CSafeArray object. |
| [Clear](#DestroyData) | Like DestroyData, destroys all the data in a safe array. It is the same that Erase and Reset. |
| [Copy](#Copy) | Creates a copy of the safe array. |
| [CopyData](#CopyData) | Copies the source array to the target array after releasing any resources in the target array. |
| [CopyFrom](#CopyFrom) | Copies the contents of a safe array. |
| [CopyFromVariant](#CopyFromVariant) | Copies the contents of a VARIANT of type VT_ARRAY to the object. |
| [CopyToVariant](#CopyToVariant) | Copies the safe array to the passed variant. |
| [Count](#Count) | Returns the number of elements in the specified dimension of the array. |
| [Create](#Create) | Creates a safe array from the given VARTYPE, number of dimensions and bounds. |
| [CreateEx](#CreateEx) | Creates a safe array from the given VARTYPE, number of dimensions and bounds. |
| [CreateVector](#CreateVector) | Creates a one-dimensional safe array from the given VARTYPE, lower bound and number elements. |
| [CreateVectorEx](#CreateVectorEx) | Creates a one-dimensional safe array from the given VARTYPE, lower bound and number elements. |
| [Destroy](#Destroy) | Destroys an existing array descriptor and all of the data in the array. |
| [DestroyData](#DestroyData) | Destroys all the data in a safe array. |
| [Detach](#Detach) | Detaches the sage array descriptor from the CSafeArray. |
| [ElemSize](#ElemSize) | Returns the size of an element. |
| [Erase](#DestroyData) | Like DestroyData, destroys all the data in a safe array. It is the same that Clear and Reset. |
| [Features](#Features) | Returns the flags used by the safe array. This is the same that the Flags method. |
| [Find](#Find) | Scans the array to search for the specified string. |
| [Flags](#Features) | Returns the flags used by the safe array. This is the same that the **Features** method. |
| [Get](#Get) | Retrieves a single element of the array. |
| [GetIID](#GetIID) | Returns the GUID of the interface contained within a given safe array. |
| [GetPtr](#Operator1) | Returns the address of the safe array. |
| [GetRecordInfo](#GetRecordInfo) | Retrieves the IRecordInfo interface of a UDT contained in a given safe array. |
| [GetType](#GetType) | Returns the VARTYPE stored in the given safe array. |
| [Insert](#Insert) | Inserts a value at the specified position of the safe array. |
| [IsResizable](#IsResizable) | Tests if the safe array can be resized. |
| [LBound](#LBound) | Returns the lower bound for any dimension of a safe array. |
| [LocksCount](#LocksCount) | Returns the number of times the array has been locked without the corresponding unlock. |
| [MoveFromVariant](#MoveFromVariant) | Transfers ownership of the safe array contained in the variant parameter to this object. |
| [MoveToVariant](#MoveToVariant) | Transfers ownership of the safe array to a variant and detaches it from the class. |
| [NumDims](#NumDims) | Returns the number of dimensions in the array. |
| [PtrOfIndex](#PtrOfIndex) | Returns a pointer to an array element. |
| [Put](#Put) | Stores the data element at a given location in the array. |
| [Redim](#Redim) | Changes the right-most (least significant) bound of a safe array. |
| [Remove](#Remove) | Deletes the specified array element. |
| [Reset](#DestroyData) | Like DestroyData, destroys all the data in a safe array. It is the same that Clear and Erase. |
| [SetIID](#SetIID) | Sets the GUID of the interface contained within a given safe array. |
| [SetRecordInfo](#SetRecordInfo) | Sets the IRecordInfo interface of the UDT contained in a given safe array. |
| [Sort](#Sort) | Sorts a one-dimensional VT_BSTR CSafeArray calling the C qsort function. |
| [UBound](#UBound) | Returns the upper bound for any dimension of a safe array. |
| [UnaccessData](#UnaccessData) | Decrements the lock count of an array, and invalidates the pointer retrieved by AccessData. |

# Helper Procedures

| Name       | Description |
| ---------- | ----------- |
| [AfxStrJoin](#AfxStrJoin) | Returns a string consisting of all of the strings in an array, each separated by a delimiter. |
| [AfxStrSplit](#AfxStrSplit) | Splits a string into tokens. |
| [AfxXmlBase64Decode](#AfxXmlBase64Decode) | Converts the contents of a Base64 mime encoded string to an ascii string. |
| [AfxXmlBase64Encode](#AfxXmlBase64Encode) | Converts the contents of a string to Base64 mime encoding. |

# <a name="SAFEARRAY"></a>SAFEARRAY Structure

Represents a safe array.

```
TYPE tagSAFEARRAY
   cDims as USHORT
   fFeatures as USHORT
   cbElements as ULONG
   cLocks as ULONG
   pvData as PVOID
   rgsabound(0 to 0) as SAFEARRAYBOUND
EBD TYPE
```

| Member     | Description |
| ---------- | ----------- |
| **cDims** | Count of dimensions of the array. |
| **fFeatures** | Flags. |
| **cbElements** | Size of an element of the array. |
| **cLocks** | Number of times the array has been locked without corresponding unlock. |
| **pvData** | Pointer to the data. |
| **rgsabound** | One bound for each dimension. |

#### fFeatures Flags

| Flag       | Description |
| ---------- | ----------- |
| FADF_AUTO | An array that is allocated on the stack. |
| FADF_STATIC | An array that is statically allocated. |
| FADF_EMBEDDED | An array that is embedded in a structure. |
| FADF_FIXEDSIZE | An array that may not be resized or reallocated. |
| FADF_RECORD | An array that contains records. When set, there will be a pointer to the IRecordinfo interface at negative offset 4 in the array descriptor. |
| FADF_HAVEIID | An array that has an IID identifying interface. When set, there will be a GUID at negative offset 16 in the safe array descriptor. Flag is set only when FADF_DISPATCH or FADF_UNKNOWN is also set. |
| FADF_HAVEVARTYPE | An array that has a VT type. When set, there will be a VT tag at negative offset 4 in the array descriptor that specifies the element type. |
| FADF_BSTR | An array of BSTRs. |
| FADF_UNKNOWN | An array of IUnknown pointers. |
| FADF_DISPATCH | An array of IDispatch pointers. |
| FADF_VARIANT | An array of VARIANTs. |
| FADF_RESERVED | Bits reserved for future use. |

#### Remarks

The array rgsabound is stored with the left-most dimension in rgsabound[0] and the right-most dimension in rgsabound(cDims - 1).

The **fFeatures** flags describe attributes of an array that can affect how the array is released. The fFeatures field describes what type of data is stored in the SAFEARRAY and how the array is allocated. This allows freeing the array without referencing its containing variant. The bits are accessed using the following constants:

#### Thread Safety

All public static members of the SAFEARRAY data type are thread safe. Instance members are not guaranteed to be thread safe.

For example, consider an application that uses the SafeArrayLock and SafeArrayUnlock functions. If these functions are called concurrently from different threads on the same SAFEARRAY data type instance, an inconsistent lock count may be created. This will eventually cause the SafeArrayUnlock function to return E_UNEXPECTED. You can prevent this by providing your own synchronization code.

# <a name="SAFEARRAYBOUND"></a>SAFEARRAYBOUND Structure

Represents the bounds of one dimension of the array. The lower bound of the dimension is represented by **lLbound**, and **cElements** represents the number of elements in the dimension. The structure is defined as follows:

```
TYPE SAFEARRAYBOUND
   cElements as ULONG
   lLbound as LONG
END TYPE
```

| Member     | Description |
| ---------- | ----------- |
| **cElements** | Number of elements in the dimension. |
| **lLbound** | The lower bound of the dimension. |

# <a name="VARTYPE"></a>Safe array Variant type

The base type of the array (the VARTYPE of each element of the array). The VARTYPE is restricted to a subset of the variant types. Neither the VT_ARRAY nor the VT_BYREF flag can be set. VT_EMPTY and VT_NULL are not valid base types for the array. All other types are legal.

| VarType    | Meaning     | Data type | strType |
| ---------- | ----------- | --------- | ------- |
| VT_BSTR | Unicode string | BSTR | "BSTR" |
| VT_I1 | Signed byte | BYTE | "BYTE" |
| VT_UI1 | Unsigned byte | UBYTE | "UBYTE" |
| VT_I2 | Signed short | SHORT | "SHORT" |
| VT_UI2 | Unsigned short | USHORT | "USHORT" |
| VT_I4 | Signed long | LONG | "LONG" |
| VT_INT | Signed long | LONG | "ULONG" |
| VT_UI4 | Unsigned long | ULONG | "LONG" |
| VT_UINT | Unsigned long | ULONG | "ULONG" |
| VT_I8 | Signed quad | LONGINT | "LONGINT" |
| VT_UI8 | Unnsigned quad | ULONGINT | "ULONGINT" |
| VT_R4 | Single | SINGLE | "SINGLE" |
| VT_R8 | Double | DOUBLE | "DOUBLE" |
| VT_CUR | Currency | CY | "CURRENCY" |
| VT_BOOL | Boolean (cast to a signed short) | SHORT | "BOOL" |
| VT_DATE | Date | DATE_ | "DATE" |
| VT_DECIMAL | Decimal structure | DECIMAL | "DECIMAL" |
| VT_VARIANT | Variant | VARIANT | "VARIANT" |
| VT_UNKNOWN | IUnknown pointer | IUnknown PTR | "UNKNOWN" |
| VT_DISPATCH | IDispatch pointer | IDispatch PTR | "DISPATCH" |

# <a name="Constructor1"></a>Constructor (SAFEARRAYBOUND)

Creates a CSafeArray.

Multidimensional array:

```
CONSTRUCTOR CSafeArray (BYVAL vt AS VARTYPE, BYVAL cDims AS UINT, BYVAL prgsabounds AS SAFEARRAYBOUND PTR)
CONSTRUCTOR CSafeArray (BYREF strType AS STRING, BYVAL cDims AS UINT, BYVAL prgsabounds AS SAFEARRAYBOUND PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *vt* | The base type of the array (the VARTYPE of each element of the array). The VARTYPE is restricted to a subset of the variant types. Neither the VT_ARRAY nor the VT_BYREF flag can be set. VT_EMPTY and VT_NULL are not valid base types for the array. All other types are legal. |
| *strType* | The base type of the array as a string literal. |
| *cDims* | Number of dimensions in the array. The number cannot be changed after the array is created. |
| *rgsabound* | Pointer to a vector of bounds (one for each dimension) to allocate for the array. |

One-dimensional array:

```
CONSTRUCTOR CSafeArray (BYVAL vt AS VARTYPE, BYVAL cElements AS ULONG = 0, BYVAL lLBound AS LONG = 0)
CONSTRUCTOR (BYREF strType AS STRING, BYVAL cElements AS ULONG = 0, BYVAL lLBound AS LONG = 0)
```

| Parameter  | Description |
| ---------- | ----------- |
| *vt* | The base type of the array (the VARTYPE of each element of the array). The VARTYPE is restricted to a subset of the variant types. Neither the VT_ARRAY nor the VT_BYREF flag can be set. VT_EMPTY and VT_NULL are not valid base types for the array. All other types are legal. |
| *strType* | The base type of the array as a string literal. |
| *cElements* | Optional. Number of elements in the array. |
| *lLBound* | Optional. The lower bound of the array. |

Two-dimensional array:

```
CONSTRUCTOR CSafeArray (BYVAL vt AS VARTYPE, BYVAL cElements1 AS ULONG, BYVAL lLBound1 AS LONG, _
   BYVAL cElements2 AS ULONG, BYVAL lLBound2 AS LONG)
CONSTRUCTOR (BYREF strType AS STRING, BYVAL cElements1 AS ULONG, BYVAL lLBound1 AS LONG, _
   BYVAL cElements2 AS ULONG, BYVAL lLBound2 AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *vt* | The base type of the array (the VARTYPE of each element of the array). The VARTYPE is restricted to a subset of the variant types. Neither the VT_ARRAY nor the VT_BYREF flag can be set. VT_EMPTY and VT_NULL are not valid base types for the array. All other types are legal. |
| *strType* | The base type of the array as a string literal. |
| *cElements1* | Number of elements in the first dimension of the array. |
| *lLBound1* | The lower bound of the first dimension of the array. |
| *cElements2* | Number of elements in the second dimension of the array. |
| *lLBound2* | The lower bound of the second dimension of the array. |

#### Usage examples

```
' // Two-dimensional array of BSTR
' // 2D: elements = 5, lower bound = 1
' // 2D: elements = 3, lower bound = 1
DIM rgsabounds(0 TO 1) AS SAFEARRAYBOUND = {(5, 1), (3, 1)}
DIM csa AS CSafeArray = CSafeArray(VT_BSTR, 2, @rgsabounds(0))
-or-
' // Two-dimensional array of BSTR
DIM csa AS CSafeArray = CSafeArray(VT_BSTR, 5, 1, 3, 1)

' // One-dimensional array of VT_VARIANT with 0 elements and a lower-bound of 0
DIM csa AS CSafeArray = CSafeArray(VT_VARIANT, 0, 0)
-or-
DIM csa AS CSafeArray = CSafeArray(VT_VARIANT)

' // One-dimensional array of VT_BSTR with 5 elements and a lower-bound of 1
DIM csa AS CSafeArray = CSafeArray(VT_BSTR, 5, 1)
```

# <a name="Constructor2"></a>Constructor (SAFEARRAYBOUND)

Creates a CSafeArray from another CSafeArray.

```
CONSTRUCTOR CSafeArray (BYREF csa AS CSafeArray)
```

| Parameter  | Description |
| ---------- | ----------- |
| *csa* | A CSafeArray object. |

# <a name="Constructor3"></a>CConstructor (SAFEARRAY PTR)

Creates a CSafeArray from a safe array.
```
CONSTRUCTOR CSafeArray (BYVAL psa AS SafeArray PTR)
CONSTRUCTOR CSafeArray (BYVAL psa AS SafeArray PTR, BYVAL fAttach AS BOOLEAN)
```

| Parameter  | Description |
| ---------- | ----------- |
| *psa* | Pointer to a safe array. |
| *fAttach* | If TRUE, the safe array is attached, else a copy is made. |

# <a name="Constructor4"></a>CConstructor (VARIANT PTR)

Creates a CSafeArray from a Variant of type VT_ARRAY.

```
CONSTRUCTOR CSafeArray (BYVAL pvar AS VARIANT PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pvar* | Pointer to the VARIANT. |

# <a name="Operator1"></a>Operator * (GetPtr)

Returns a pointer to the underlying safe array descriptor.

```
OPERATOR * () AS ANY PTR
```

#### Remark

You can also call the **GetPtr** method.

# <a name="Operator2"></a>Operator LET ( = )

Assigns a CSafeArray to another CSafeArray.<br>
Assigns a safe array to a CSafeArray.<br>
Assigns a variant of type VT_ARRAY to a CSafeArray.

```
OPERATOR Let (BYREF csa AS CSafeArray)
OPERATOR Let (BYVAL psa AS SAFEARRAY PTR)
OPERATOR Let (BYVAL pvar AS VARIANT PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *csa* | An instance of the CSafeArray class. |
| *psa* | A safe array pointer. |
| *pvar* | Pointer to a VARIANT. |

# <a name="AfxStrJoin"></a>AfxStrJoin

Returns a string consisting of all of the strings in an array, each separated by a delimiter. If the delimiter is a null (zero-length) string then no separators are inserted between the string sections. If the delimiter expression is the 3-byte value of "," which may be expressed in your source code as the string literal """,""" or as Chr(34,44,34) then a leading and trailing double-quote is added to each string section. This ensures that the returned string contains standard comma-delimited quoted fields that can be easily parsed.

```
FUNCTION AfxStrJoin (BYREF cwsa AS CSafeArray, BYREF wszDelimiter AS WSTRING = " ") AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *cwsa* | The one-dimensional VT_BSTR CSafeArray to join. |
| *wszDelimiter* | The delimiter character. |

#### Return value

A CWSTR containing the joined string.

#### Usage example

```
DIM csa AS CSafeArray = CSafeArray("STRING", 3, 1)
csa.PutStr(1, "One")
csa.PutStr(2, "Two")
csa.PutStr(3, "Three")
DIM cws AS CWSTR = AfxStrJoin(csa, ",")
PRINT cws   ' ouput: One,Two,Three
```

# <a name="AfxStrSplit"></a>AfxStrSplit

Splits a string into tokens, which are sequences of contiguous characters separated by any of the characters that are part of delimiters.

```
FUNCTION AfxStrSplit (BYREF wszStr AS CONST WSTRING, BYREF wszDelimiters AS WSTRING = " ") AS CSafeArray
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszStr* | The string to split. |
| *wszDelimiters* | The delimiter characters. |

#### Return value

A CSafeArray containing a token in each element.

#### Usage example

```
DIM cws AS CWSTR = "- This, a sample string."
DIM cwsa AS CSafeArray = AfxStrSplit(cws, " ,.-")
FOR i AS LONG = cwsa.LBound TO cwsa.UBound
  PRINT cwsa.GetStr(i)
NEXT
```
# <a name="AfxXmlBase64Decode"></a>AfxXmlBase64Decode

Converts the contents of a Base64 mime encoded string to an ascii string.

```
FUNCTION AfxXmlBase64Decode (BYREF strData AS STRING) AS STRING
```
| Parameter  | Description |
| ---------- | ----------- |
| *strData* | The string to decode. |

#### Return value

The decoded string on success, or a null string on failure.

Remaks

Base64 is a group of similar encoding schemes that represent binary data in an ASCII string format by translating it into a radix-64 representation. The Base64 term originates from a specific MIME content transfer encoding.

Base64 encoding schemes are commonly used when there is a need to encode binary data that needs be stored and transferred over media that are designed to deal with textual data. This is to ensure that the data remains intact without modification during transport. Base64 is used commonly in a number of applications including email via MIME, and storing complex data in XML.

# <a name="AfxXmlBase64Encode"></a>AfxXmlBase64Encode

Converts the contents of a string to Base64 mime encoding.

```
FUNCTION AfxXmlBase64Encode (BYREF strData AS STRING) AS STRING
```
| Parameter  | Description |
| ---------- | ----------- |
| *strData* | The string to encode. |

#### Return value

The encoded string on succeess, or a null string on failure.

Remaks

Base64 is a group of similar encoding schemes that represent binary data in an ASCII string format by translating it into a radix-64 representation. The Base64 term originates from a specific MIME content transfer encoding.

Base64 encoding schemes are commonly used when there is a need to encode binary data that needs be stored and transferred over media that are designed to deal with textual data. This is to ensure that the data remains intact without modification during transport. Base64 is used commonly in a number of applications including email via MIME, and storing complex data in XML.

# <a name="AccessData"></a>AccessData

Retrieves a pointer to the array data and increments the lock count of an array.

```
FUNCTION AccessData () AS ANY PTR
```

#### Return value

IF it succeeds, it returns a pointer to the array data. If it fails, it returns a null pointer.

# <a name="Append"></a>Append

Appends a value to the end of the one-dimensional safe array.

```
FUNCTION Append (BYVAL pData AS ANY PTR) AS HRESULT
FUNCTION Append (BYREF cbsData AS CBSTR) AS HRESULT
FUNCTION Append (BYREF cvData AS CVAR) AS HRESULT
FUNCTION Append (BYVAL vData AS VARIANT) AS HRESULT
```

Appends a string to the end of the one-dimensional safe array.

```
FUNCTION AppendStr (BYVAL pwszData AS WSTRING PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pData* | Pointer to a variable of the appropriate data type. |
| *cbsData* | The CBSTR to insert, if the safe array is of type VT_BSTR. |
| *cvData* | The CVAR to insert, if the safe array is of type VT_VARIANT. |
| *vData* | The VARIANT to insert, if the safe array is of type VT_VARIANT. |

#### Usage examples

```
' // Create a one-dimensional array of doubles
'DIM csa AS CSafeArray = CSafeArray(VT_R8, 2, 1)
DIM csa AS CSafeArray = CSafeArray("DOUBLE", 2, 1)

DIM dblVal AS DOUBLE = 12345.12
csa.Put(1, @dblVal)
dblVal = 74447.34
csa.Put(2, @dblVal)
dblVal = 63535.567
csa.Append(@dblVal)

csa.Get(1, @dblVal)
print dblVal
csa.Get(2, @dblVal)
print dblVal
csa.Get(3, @dblVal)
print dblVal
```

```
' // Create a one-dimensional array of variants
'DIM csa AS CSafeArray = CSafeArray(VT_VARIANT, 2, 1)
DIM csa AS CSafeArray = CSafeArray("VARIANT", 2, 1)

csa.Put(1, CVAR("Test string"))
csa.Put(2, CVAR(12345))
csa.Append(CVAR("Test string appended"))

DIM cvOut AS CVAR
csa.Get(1, cvOut)
print cvOut
csa.Get(2, cvOut)
print cvOut
csa.Get(3, cvOut)
print cvOut
```

#### Return value

S_OK (0) on success or an HREUSLT code on failure.

| HREUSLT  | Description |
| ---------- | ----------- |
| DISP_E_BADINDEX | The specified index is not valid. |
| E_INVALIDARG | One of the arguments is not valid. |
| E_OUTOFMEMORY | Memory could not be allocated for the element. |
| E_FAIL | Failure. The array descriptor is null. |

# <a name="Attach"></a>Attach

Attaches a safe array to a CSafeArray object.

```
FUNCTION Attach (BYVAL psaSrc AS SAFEARRAY PTR) AS HRESULT
```

#### Return value

S_OK (0) on success or an HREUSLT code.

# <a name="DestroyData"></a>DestroyData / Clear / Erase / Reset

Destroys all the data in a safe array.

```
FUNCTION DestroyData () AS HRESULT
```

#### Return value

S_OK (0) on success or an HREUSLT code  on failure.

| HREUSLT  | Description |
| ---------- | ----------- |
| DISP_E_ARRAYISLOCKED | The array is locked. |

# <a name="Copy"></a>Copy

Creates a copy of the safe array.

```
FUNCTION Copy () AS SAFEARRAY PTR
```

#### Return value

Pointer of the new array descriptor. You must free this pointer calling the API function **SafeArrayDestroy**.

# <a name="CopyData"></a>CopyData

Copies the source array to the target array after releasing any resources in the target array. This is similar to **Copy**, except that the target array has to be set up by the caller. The target is not allocated or reallocated.

```
FUNCTION CopyData (BYVAL psaTarget AS SAFEARRAY PTR) AS HRESULT
```

On exit, the array referred to by *psaTarget* contains a copy of the data if the call succeeds.

#### Return value

S_OK (0) on success or an HREUSLT code.

| HREUSLT  | Description |
| ---------- | ----------- |
| E_INVALIDARG | The dimensions or the number of dimensions don't match. |
| E_OUTOFMEMORY | Insufficient memory to create the copy. |

# <a name="CopyFrom"></a>CopyFrom

Copies the contents of the passed safe array.

```
FUNCTION CopyFrom (BYVAL psaSrc AS SAFEARRAY PTR) AS HRESULT
```
#### Return value

S_OK (0) on success or an HREUSLT code on failure.

#### Remarks

**CopyFrom** calls the string or variant manipulation functions if the array to copy contains either of these data types. If the array being copied contains object references, the reference counts for the objects are incremented.


# <a name="CopyFromVariant"></a>CopyFromVariant

Copies the contents of a VARIANT of type VT_ARRAY to the object. The VARIANT remains unaltered.

```
FUNCTION CopyFromVariant (BYVAL pvar AS VARIANT PTR) AS HRESULT
```
*pvar* must be variant of type VT_ARRAY, i.e. containing a safe array. If it is of another type, an invalid data error is returned.

#### Return value

S_OK (0) on success or an HREUSLT code on failure.

#### Remarks

**CopyFromVariant** calls the string or variant manipulation functions if the array to copy contains either of these data types. If the array being copied contains object references, the reference counts for the objects are incremented.

# <a name="CopyToVariant"></a>CopyToVariant

Copies the safe array to the passed variant.

```
FUNCTION CopyToVariant (BYVAL pvar AS VARIANT PTR) AS HRESULT
```
*pvar* is a pointer to the variant where the safe array will be copied.

#### Return value

S_OK (0) on success or an HREUSLT code on failure.

# <a name="Count"></a>Count

Returns the number of elements in the specified dimension of the array.

```
FUNCTION Count (BYVAL nDim AS UINT = 1) AS UINT
```

The optional nDim parameter is the array dimension for which to get the number of elements. You don't need to pass this parameter if the safe array in one-dimensional.

#### Return value

S_OK (0) on success or an HREUSLT code on failure.

# <a name="Create"></a>Create

Creates a safe array from the given VARTYPE, number of dimensions and bounds.

Multidimensional array:

```
FUNCTION Create (BYVAL vt AS VARTYPE, BYVAL cDims AS UINT, BYVAL prgsabound AS SAFEARRAYBOUND PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *vt* | The base type of the array (the VARTYPE of each element of the array). The VARTYPE is restricted to a subset of the variant types. Neither the VT_ARRAY nor the VT_BYREF flag can be set. VT_EMPTY and VT_NULL are not valid base types for the array. All other types are legal. |
| *cDims* | Number of dimensions in the array. The number cannot be changed after the array is created. |
| *rgsabound* | Pointer to a vector of bounds (one for each dimension) to allocate for the array. |

One-dimensional array:

```
FUNCTION Create (BYVAL vt AS VARTYPE, BYVAL cElements AS ULONG, BYVAL lLBound AS LONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *vt* | The base type of the array (the VARTYPE of each element of the array). The VARTYPE is restricted to a subset of the variant types. Neither the VT_ARRAY nor the VT_BYREF flag can be set. VT_EMPTY and VT_NULL are not valid base types for the array. All other types are legal. |
| *cElements* | Optional. Number of elements in the array. |
| *lLBound* | Optional. The lower bound of the array. |

Two-dimensional array:

```
FUNCTION Create (BYVAL vt AS VARTYPE, BYVAL cElements1 AS ULONG, BYVAL lLBound1 AS LONG, _
   BYVAL cElements2 AS ULONG, BYVAL lLBound2 AS LONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *vt* | The base type of the array (the VARTYPE of each element of the array). The VARTYPE is restricted to a subset of the variant types. Neither the VT_ARRAY nor the VT_BYREF flag can be set. VT_EMPTY and VT_NULL are not valid base types for the array. All other types are legal. |
| *cElements1* | Number of elements in the first dimension of the array. |
| *lLBound1* | The lower bound of the first dimension of the array. |
| *cElements2* | Number of elements in the second dimension of the array. |
| *lLBound2* | The lower bound of the second dimension of the array. |

#### Return value

S_OK (0) on success or an HREUSLT code on failure.

# <a name="CreateEx"></a>CreateEx

Creates a safe array from the given VARTYPE, number of dimensions and bounds.

Multidimensional array:

```
FUNCTION CreateEx (BYVAL vt AS VARTYPE, BYVAL cDims AS UINT, _
   BYVAL prgsabound AS SAFEARRAYBOUND PTR, BYVAL pvExtra AS PVOID) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *vt* | The base type of the array (the VARTYPE of each element of the array). The VARTYPE is restricted to a subset of the variant types. Neither the VT_ARRAY nor the VT_BYREF flag can be set. VT_EMPTY and VT_NULL are not valid base types for the array. All other types are legal. The FADF_RECORD flag can be set for a variant type VT_RECORD, The FADF_HAVEIID flag can be set for VT_DISPATCH or VT_UNKNOWN, and FADF_HAVEVARTYPE can be set for all other VARTYPEs. For more information about the FADF_RECORD, FADF_HAVEIID or FADF_HAVEVARTYPE flags see SAFEARRAY Data Type. |
| *cDims* | Number of dimensions in the array. The number cannot be changed after the array is created. |
| *rgsabound* | Pointer to a vector of bounds (one for each dimension) to allocate for the array. |
| *pvExtra* | Points to the type information of the user-defined type, if you are creating a safe array of user-defined types. If the *vt* parameter is VT_RECORD, then *pvExtra* will be a pointer to an **IRecordInfo** interface describing the record. If the *vt* parameter is VT_DISPATCH or VT_UNKNOWN, then *pvExtra* will contain a pointer to a GUID representing the type of interface being passed to the array. |

One-dimensional array:

```
FUNCTION CreateEx (BYVAL vt AS VARTYPE, BYVAL cElements AS ULONG, _
   BYVAL lLBound AS LONG, BYVAL pvExtra AS PVOID) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *vt* | The base type of the array (the VARTYPE of each element of the array). The VARTYPE is restricted to a subset of the variant types. Neither the VT_ARRAY nor the VT_BYREF flag can be set. VT_EMPTY and VT_NULL are not valid base types for the array. All other types are legal. The FADF_RECORD flag can be set for a variant type VT_RECORD, The FADF_HAVEIID flag can be set for VT_DISPATCH or VT_UNKNOWN, and FADF_HAVEVARTYPE can be set for all other VARTYPEs. For more information about the FADF_RECORD, FADF_HAVEIID or FADF_HAVEVARTYPE flags see SAFEARRAY Data Type. |
| *cElements* | Optional. Number of elements in the array. |
| *lLBound* | Optional. The lower bound of the array. |
| *pvExtra* | Points to the type information of the user-defined type, if you are creating a safe array of user-defined types. If the *vt* parameter is VT_RECORD, then *pvExtra* will be a pointer to an **IRecordInfo** interface describing the record. If the *vt* parameter is VT_DISPATCH or VT_UNKNOWN, then *pvExtra* will contain a pointer to a GUID representing the type of interface being passed to the array. |

Two-dimensional array:

```
FUNCTION CreateEx (BYVAL vt AS VARTYPE, BYVAL cElements1 AS ULONG, BYVAL lLBound1 AS LONG, _
   BYVAL cElements2 AS ULONG, BYVAL lLBound2 AS LONG, BYVAL pvExtra AS PVOID) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *vt* | The base type of the array (the VARTYPE of each element of the array). The VARTYPE is restricted to a subset of the variant types. Neither the VT_ARRAY nor the VT_BYREF flag can be set. VT_EMPTY and VT_NULL are not valid base types for the array. All other types are legal. The FADF_RECORD flag can be set for a variant type VT_RECORD, The FADF_HAVEIID flag can be set for VT_DISPATCH or VT_UNKNOWN, and FADF_HAVEVARTYPE can be set for all other VARTYPEs. For more information about the FADF_RECORD, FADF_HAVEIID or FADF_HAVEVARTYPE flags see SAFEARRAY Data Type. |
| *cElements1* | Number of elements in the first dimension of the array. |
| *lLBound1* | The lower bound of the first dimension of the array. |
| *cElements2* | Number of elements in the second dimension of the array. |
| *lLBound2* | The lower bound of the second dimension of the array. |
| *pvExtra* | Points to the type information of the user-defined type, if you are creating a safe array of user-defined types. If the *vt* parameter is VT_RECORD, then *pvExtra* will be a pointer to an **IRecordInfo** interface describing the record. If the *vt* parameter is VT_DISPATCH or VT_UNKNOWN, then *pvExtra* will contain a pointer to a GUID representing the type of interface being passed to the array. |

#### Return value

S_OK (0) on success or an HRESULT code on failure.

# <a name="CreateVector"></a>CreateVector

Creates a fixed size safe array from the given VARTYPE, lower bound and number of elements.

```
FUNCTION CreateVector (BYVAL vt AS VARTYPE, BYVAL cElements AS ULONG, BYVAL lLBound AS LONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *vt* | The base type of the array (the VARTYPE of each element of the array). The VARTYPE is restricted to a subset of the variant types. Neither the VT_ARRAY nor the VT_BYREF flag can be set. VT_EMPTY and VT_NULL are not valid base types for the array. All other types are legal. |
| *cElements* | The number of elements in the array. |
| *lLBound* | The lower bound for the array. Can be negative. |

#### Return value

S_OK (0) on success or an HRESULT code on failure.

#### Remarks

**CreateVector** allocates a single block of memory containing a SAFEARRAY structure for a single-dimension array (24 bytes), immediately followed by the array data. All of the existing safe array functions work correctly for safe arrays that are allocated with CreateVector.

A safe array created with CreateVector is allocated as a single block of memory. Both the SafeArray descriptor and the array data block are allocated contiguously in one allocation, which speeds up array allocation. However, a user can allocate the descriptor and data area separately using the **SafeArrayAllocDescriptor** and **SafeArrayAllocData*** calls.

# <a name="CreateVectorEx"></a>CreateVectorEx

Creates a fixed size safe array from the given VARTYPE, lower bound and number of elements.

```
FUNCTION CreateVectorEx (BYVAL vt AS VARTYPE, BYVAL cElements AS ULONG, BYVAL lLBound AS LONG, _
   BYVAL pvExtra AS ANY PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *vt* | The base type of the array (the VARTYPE of each element of the array). The VARTYPE is restricted to a subset of the variant types. Neither the VT_ARRAY nor the VT_BYREF flag can be set. VT_EMPTY and VT_NULL are not valid base types for the array. All other types are legal. The FADF_RECORD flag can be set for a variant type VT_RECORD, The FADF_HAVEIID flag can be set for VT_DISPATCH or VT_UNKNOWN, and FADF_HAVEVARTYPE can be set for all other VARTYPEs. For more information about the FADF_RECORD, FADF_HAVEIID or FADF_HAVEVARTYPE flags see SAFEARRAY Data Type. |
| *cElements* | The number of elements in the array. |
| *lLBound* | The lower bound for the array. Can be negative. |
| *pvExtra* | Points to the type information of the user-defined type, if you are creating a safe array of user-defined types. If the *vt* parameter is VT_RECORD, then pvExtra will be a pointer to an **IRecordInfo** interface describing the record. If the *vt* parameter is VT_DISPATCH or VT_UNKNOWN, then *pvExtra* will contain a pointer to a GUID representing the type of interface being passed to the array. |

#### Return value

S_OK (0) on success or an HRESULT code on failure.

# <a name="Destroy"></a>Destroy

Destroys an existing array descriptor and all of the data in the array. If objects are stored in the array, Release is called on each object in the array.

```
FUNCTION Destroy () AS HRESULT
```

#### Return value

S_OK (0) on success or an HRESULT code on failure.

| HRESULT  | Description |
| ---------- | ----------- |
| E_INVALIDARG | The dimensions or the number of dimensions don't match. |
| DISP_E_ARRAYISLOCKED | The array is currently locked. |

# <a name="Detach"></a>Detach

Detaches the safe array descriptor from the CSafeArray.

```
FUNCTION Detach () AS SAFEARRAY PTR
```

#### Return value

Returns a pointer to a SAFEARRAY descriptor.

#### Remarks

The caller takes ownership of it and must destroy it when no longer needed.

# <a name="ElemSize"></a>ElemSize

Returns the size of an element.

```
FUNCTION ElemSize () AS UINT
```

#### Return value

Returns the size (in bytes) of an element in a safe array. Does not include size of pointed-to data.

# <a name="Features"></a>Features / Flags

Returns the flags used by the safe array. This is the same that the **Flags** method.

```
FUNCTION Features () AS USHORT
```

# <a name="Find"></a>Find

Scans the array to search for the specified string.

```
FUNCTION Find (BYREF wszFind AS WSTRING, BYVAL bNoCase AS BOOLEAN = FALSE) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFind* | The string to find. |
| *bNoCase* | Optional. TRUE = Ignore case. |

#### Return value

The index of the retrieved array element, or 0 on failue.

# <a name="Get"></a>Get / GetStr / GetVar

Retrieves a single element of the array.

Multidimensional array:

```
FUNCTION Get (BYVAL prgIndices AS LONG PTR, BYVAL pData AS ANY PTR) AS HRESULT
FUNCTION Get (BYVAL prgIndices AS LONG PTR, BYREF cbsData AS CBSTR) AS HRESULT
FUNCTION Get (BYVAL prgIndices AS LONG PTR, BYREF cvData AS CVAR) AS HRESULT
FUNCTION GetStr (BYVAL prgIndices AS LONG PTR) AS CBSTR
FUNCTION GetVar (BYVAL prgIndices AS LONG PTR) AS CVAR
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgIndices* | Pointer to a vector of indexes for each dimension of the array. The right-most (least significant) dimension is rgIndices(0). The left-most dimension is stored at pgIndices(@psa.cDims – 1). |
| *pData* | Pointer to the location to place the element of the array. |
| *cbsData* | A CBSTR passed by reference that will receive the result. |
| *cvData* | A CVAR passed by reference that will receive the result. |

One-dimensional array:

```
FUNCTION Get (BYVAL idx AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
FUNCTION Get (BYVAL idx AS LONG, BYREF cbsData AS CBSTR) AS HRESULT
FUNCTION Get (BYVAL idx AS LONG, BYREF cvData AS CVAR) AS HRESULT
FUNCTION GetStr (BYVAL idx AS LONG) AS CBSTR
FUNCTION GetVar (BYVAL idx AS LONG) AS CVAR
```

| Parameter  | Description |
| ---------- | ----------- |
| *idx* | Index of the element. |
| *pData* | Pointer to the location to place the element of the array. |
| *cbsData* | A CBSTR passed by reference that will receive the result. |
| *cvData* | A CVAR passed by reference that will receive the result. |

Two-dimensional array:

```
FUNCTION Get (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
FUNCTION Get (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF cbsData AS CBSTR) AS HRESULT
FUNCTION Get (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF cvData AS CVAR) AS HRESULT
FUNCTION GetStr (BYVAL cElem AS LONG, BYVAL cDim AS LONG) AS CVAR
FUNCTION GetVar (BYVAL cElem AS LONG, BYVAL cDim AS LONG) AS CVAR
```

| Parameter  | Description |
| ---------- | ----------- |
| *cElem* | Index of the element. |
| *cDim* | Dimension of the array. |
| *pData* | Pointer to the location where to place the element of the array. |
| *cbsData* | A CBSTR passed by reference that will receive the result. |
| *cvData* | A CVAR passed by reference that will receive the result. |

First element, then dimension, e.g. 2, 1 (element 2, first dimension), 1, 2 (element 1, 2nd dimension).

#### Usage example

```
' // One-dimensional array of VT_BSTR
DIM csa AS CSafeArray = CSafeArray(VT_BSTR, 2, 1)
csa.PutStr(1, "Test string 1")
print csa.GetStr(1)
csa.PutStr(2, "Test string 2")
print csa.GetStr(2)
```

```
' // One-dimensional array of VT_VARIANT
DIM csa AS CSafeArray = CSafeArray(VT_VARIANT, 5, 1)
DIM cv1 AS CVAR = "Test variant 1"
csa.Put(1, cv1)
DIM cvOut AS CVAR
csa.Get(1, cvOut)
print cvOut
DIM cv2 AS CVAR = "Test variant 2"
csa.Put(1, cv2)
csa.Get(1, cvOut)
print cvOut
```

```
' // Two-dimensional array of BSTR
' // 2D: elements = 5, lower bound = 1
' // 2D: elements = 3, lower bound = 1
DIM rgsabounds(0 TO 1) AS SAFEARRAYBOUND = {(5, 1), (3, 1)}
DIM csa AS CSafeArray = CSafeArray(VT_BSTR, 2, @rgsabounds(0))

' // array index: first element, first dimension
DIM rgidx(0 TO 1) AS LONG = {1, 1}
DIM cbs1 AS CBSTR = "Test string 1"
' // Put the value
csa.Put(@rgidx(0), cbs1)
' // Get the value
DIM cbsOut AS CBSTR
csa.Get(@rgidx(0), cbsOut)
print cbsOut
' // array index: second element, first dimension
rgidx(0) = 2 : rgidx(1) = 1
' // Put the value
DIM cbs2 AS CBSTR = "Test string 2"
csa.Put(@rgidx(0), cbs2)
' // Get the value
csa.Get(@rgidx(0), cbsOut)
print cbsOut

' // array index: first element, second dimension
rgidx(0) = 1 : rgidx(1) = 2
' // Put the value
DIM cbs3 AS CBSTR = "Test string 3"
csa.Put(@rgidx(0), cbs3)
' // Get the value
csa.Get(@rgidx(0), cbsOut)
print cbsOut

' // array index: second element, second dimension
rgidx(0) = 2 : rgidx(1) = 2
' // Put the value
DIM cbs4 AS CBSTR = "Test string 4"
csa.Put(@rgidx(0), cbs4)
' // Get the value
csa.Get(@rgidx(0), cbsOut)
print cbsOut
```

#### Return value (Get)

S_OK (0) on success or an HRESULT code on failure.

| HRESULT  | Description |
| ---------- | ----------- |
| DISP_E_BADINDEX | The specified index is invalid. |
| E_INVALIDARG | One of the arguments is invalid. |
| E_OUTOFMEMORY | Memory could not be allocated for the element. |

#### Return value (GetStr / GetVar)

The string or variant element.

#### Remarks

This method calls **SafearrayLock** and **SafearrayUnlock** automatically, before and after retrieving the element. The caller must provide a storage area of the correct size to receive the data. If the data element is a string, object, or variant, the function copies the element in the correct way.

# <a name="GetIID"></a>GetIID

Returns the GUID of the interface contained within a given safe array.

```
FUNCTION GetIID () AS GUID
```

#### Return value

The GUID of the interface, on success, or a null guid on failure.

# <a name="GetRecordInfo"></a>GetRecordInfo

Retrieves the IRecordInfo interface of a UDT contained in a given safe array.

```
FUNCTION GetRecordInfo () AS IRecordInfo PTR
```

# <a name="GetType"></a>GetType

Returns the VARTYPE stored in the given safe array.

```
FUNCTION GetType () AS VARTYPE
```

# <a name="Insert"></a>Insert

Inserts a value at the specified position of the safe array.

```
FUNCTION Insert (BYVAL nPos AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
FUNCTION Insert (BYVAL nPos AS LONG, BYREF cbsData AS CBSTR) AS HRESULT
FUNCTION Insert (BYVAL nPos AS LONG, BYREF cvData AS CVAR) AS HRESULT
FUNCTION Insert (BYVAL nPos AS LONG, BYVAL vData AS VARIANT) AS HRESULT
```
Inserts a value at the beginning of the safe array.

```
FUNCTION Insert (BYVAL pData AS ANY PTR) AS HRESULT
FUNCTION Insert (BYREF cbsData AS CBSTR) AS HRESULT
FUNCTION Insert (BYREF cvData AS CVAR) AS HRESULT
FUNCTION Insert (BYVAL vData AS VARIANT) AS HRESULT
```

Inserts a string.

```
FUNCTION InsertStr (BYVAL nPos AS LONG, BYVAL pwszData AS WSTRING PTR) AS HRESULT
FUNCTION InsertStr (BYVAL pwszData AS WSTRING PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *nPos* | Optional. Index of the array in which the data will be inserted. If *nPos* is not specified, the item is inserted at the beginning of the array. If the array is empty, it is redimensioned to one element. |
| *pData* | Pointer to a variable of the appropriate data type. |
| *cbsData* | The CBSTR to insert, if the safe array is of type VT_BSTR. |
| *cvData* | The CVAR to insert, if the safe array is of type VT_VARIANT. |
| *vData* | The VARIANT to insert, if the safe array is of type VT_VARIANT. |

#### Return value

S_OK (0) on success or an HRESULT code on failure.

| HRESULT  | Description |
| ---------- | ----------- |
| DISP_E_BADINDEX | The specified index is invalid. |
| E_INVALIDARG | One of the arguments is invalid. |
| E_OUTOFMEMORY | Memory could not be allocated for the element. |
| E_FAIL | Failure. The array descriptor is null. |

#### Usage example

```
' // Create a one-dimensional array of variants
'DIM csa AS CSafeArray = CSafeArray(VT_VARIANT, 2, 1)
DIM csa AS CSafeArray = CSafeArray("VARIANT", 2, 1)

csa.Put(1, CVAR("Test string 1"))
csa.Put(2, CVAR("Test string 2"))
csa.Insert(2, CVAR(12345.67, "DOUBLE"))

DIM cvOut AS CVAR
csa.Get(1, cvOut)
print cvOut
csa.Get(2, cvOut)
print cvOut
csa.Get(3, cvOut)
print cvOut
```

# <a name="IsResizable"></a>IsResizable

Tests if the safe array can be resized.

```
FUNCTION IsResizable () AS BOOLEAN
```

#### Return value

TRUE if the array can be resized; FALSE, otherwise.

# <a name="LBound"></a>LBound

Returns the lower bound for any dimension of a safe array.

```
FUNCTION LBound (BYVAL nDim AS UINT = 1) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nDim* | Optional. The array dimension for which to get the lower bound. You don't need to pass this parameter if it is a one-dimensional array. |

# <a name="LBound"></a>LBound

Returns the lower bound for any dimension of a safe array.

```
FUNCTION LBound (BYVAL nDim AS UINT = 1) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nDim* | Optional. The array dimension for which to get the lower bound. You don't need to pass this parameter if it is a one-dimensional array. |

# <a name="UBound"></a>UBound

Returns the upper bound for any dimension of a safe array.

```
FUNCTION UBound (BYVAL nDim AS UINT = 1) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nDim* | Optional. The array dimension for which to get the upper bound. You don't need to pass this parameter if it is a one-dimensional array. |

# <a name="LocksCount"></a>LocksCount

Returns the number of times the array has been locked without the corresponding unlock.

```
FUNCTION LocksCount () AS UINT
```

# <a name="MoveFromVariant"></a>MoveFromVariant

Transfers ownership of the safe array contained in the variant parameter to this object. The variant is then changed to VT_EMPTY.

```
FUNCTION MoveFromVariant (BYVAL pvar AS VARIANT PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pvar* | Pointer to the variant containing the safe array. |

#### Return value

S_OK (0) on success or an HRESULT code on failure.

# <a name="MoveToVariant"></a>MoveToVariant

Transfers ownership of the safe array to a variant and detaches it from the class.

```
FUNCTION MoveToVariant (BYVAL pvar AS VARIANT PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pvar* | Pointer to the variant where the safe array will be moved. |

#### Usage example

```
' // One-dimensional array of VT_BSTR
DIM csa AS CSafeArray = CSafeArray(VT_BSTR, 2, 1)
csa.Put(1, CBSTR("Test string 1"))
DIM cbsOut AS CBSTR
csa.Get(1, cbsOut.vptr)
print cbsOut
csa.Put(2, CBSTR("Test string 2"))
csa.Get(2, cbsOut)
print cbsOut

DIM v AS VARIANT
csa.MoveToVariant(@v)
print AfxVarToStr(@v)
VariantClear @v
```

#### Return value

S_OK (0) on success or an HRESULT code on failure.

# <a name="NumDims"></a>NumDims

Returns the number of dimensions in the array.

```
FUNCTION NumDims () AS UINT
```

# <a name="PtrOfIndex"></a>PtrOfIndex

Returns a pointer to an array element.

Multidimensional array:

```
FUNCTION PtrOfIndex (BYVAL prgIndices AS LONG PTR) AS ANY PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgIndices* | An array of index values that identify an element of the array. All indexes for the element must be specified. |

One-dimensional array:

```
FUNCTION PtrOfIndex (BYVAL idx AS LONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *idx* | Index of the element. |

Two-dimensional array:

```
FUNCTION PtrOfIndex (BYVAL cElem AS LONG, BYVAL cDim AS LONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cElem* | Index of the element. |
| *cDim* | Dimension number of the array. |

# <a name="Put"></a>Put

Stores the data element at a given location in the array.

Multidimensional array:

```
FUNCTION Put (BYVAL prgIndices AS LONG PTR, BYVAL pData AS ANY PTR) AS HRESULT
FUNCTION PutStr (BYVAL prgIndices AS LONG PTR, BYREF cbsData AS CBSTR) AS HRESULT
FUNCTION PutVar (BYVAL prgIndices AS LONG PTR, BYREF cvData AS CVAR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *prgIndices* | Pointer to a vector of indexes for each dimension of the array. The right-most (least significant) dimension is rgIndices(0). The left-most dimension is stored at prgIndices(@psa.cDims – 1). |
| *pData* | Pointer to the data to assign to the array. The variant types VT_DISPATCH, VT_UNKNOWN, and VT_BSTR are pointers, and do not require another level of indirection. |
| *cbsData* | A CBSTR. |
| *cbsData* | A CVAR. |

One-dimensional array:

```
FUNCTION Put (BYVAL idx AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
FUNCTION PutStr (BYVAL idx AS LONG, BYREF cbsData AS CBSTR) AS HRESULT
FUNCTION PutVar (BYVAL idx AS LONG, BYREF cvData AS CVAR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *idx* | Index of the element of the array. |
| *pData* | Pointer to the data to assign to the array. The variant types VT_DISPATCH, VT_UNKNOWN, and VT_BSTR are pointers, and do not require another level of indirection. |
| *cbsData* | A CBSTR. |
| *cbsData* | A CVAR. |

| Parameter  | Description |
| ---------- | ----------- |
| *idx* | Index of the element. |

Two-dimensional array:

```
FUNCTION Put (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYVAL pData AS ANY PTR) AS HRESULT
FUNCTION PutStr (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF cbsData AS CBSTR) AS HRESULT
FUNCTION PutVar (BYVAL cElem AS LONG, BYVAL cDim AS LONG, BYREF cvData AS CVAR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cElem* | Index of the element of the array. |
| *cDim* | Dimension number of the array. |
| *pData* | Pointer to the data to assign to the array. The variant types VT_DISPATCH, VT_UNKNOWN, and VT_BSTR are pointers, and do not require another level of indirection. |
| *cbsData* | A CBSTR. |
| *cbsData* | A CVAR. |

#### Return value

S_OK (0) on success or an HRESULT code on failure.

| HRESULT  | Description |
| ---------- | ----------- |
| DISP_E_BADINDEX | The specified index is invalid. |
| E_INVALIDARG | One of the arguments is invalid. |
| E_OUTOFMEMORY | Memory could not be allocated for the element. |
| E_FAIL | Failure. The array descriptor is null. |

#### Remarks

This function automatically calls **SAfeArrayLock** and **SafeArrayUnlock**  before and after assigning the element. If the data element is a string, object, or variant, the function copies it correctly when the safe array is destroyed. If the existing element is a string, object, or variant, it is cleared correctly. If the data element is a VT_DISPATCH or VT_UNKNOWN, **AddRef** is called to increment the object's reference count. 

Multiple locks can be on an array. Elements can be put into an array while the array is locked by other operations.

# <a name="Redim"></a>Redim

Changes the right-most (least significant) bound of a safe array.

Multidimensional array:

```
FUNCTION Redim (BYVAL pnewsabounds AS SAFEARRAYBOUND PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pnewsabounds* | Pointer to a new safe array bound structure that contains the new array boundary. You can change only the least significant dimension of an array. |

One-dimensional array:

```
FUNCTION Redim (BYVAL cElements AS ULONG, BYVAL lLBound AS LONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cElements* | Number of elements in the array. |
| *lLBound* | The lower bound of the array. |

Two-dimensional array:

```
FUNCTION Redim (BYVAL cElements1 AS ULONG, BYVAL lLBound1 AS LONG, _
   BYVAL cElements2 AS ULONG, BYVAL lLBound2 AS LONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cElements1* | Number of elements in the first dimension of the array |
| *lLBound1* | The lower bound of the first dimension of the array. |
| *cElements2* | Number of elements in the second dimension of the array |
| *lLBound2* | The lower bound of the second dimension of the array. |

#### Usage example

```
' // One-dimensional array of VT_VARIANT
' // Two elements, lower-bound 1
'DIM csa AS CSafeArray = CSafeArray(VT_VARIANT, 2, 1)
DIM csa AS CSafeArray = CSafeArray("VARIANT", 2, 1)
DIM cv1 AS CVAR = "Test variant 1"
csa.Put(1, cv1)
DIM cvOut AS CVAR
csa.Get(1, cvOut)
print cvOut

DIM cv2 AS CVAR = "Test variant 2"
csa.Put(1, cv2)
csa.Get(1, cvOut)
print cvOut

' // Redim (preserve) the safe array
csa.Redim(1, 3)
DIM cv3 AS CVAR = "Test variant 3"
csa.Put(3, cv3)
csa.Get(3, cvOut)
print cvOut
```

#### Return value

S_OK (0) on success or an HRESULT code on failure.

| HRESULT    | Description |
| ---------- | ----------- |
| DISP_E_ARRAYISLOCKED | The array is currently locked. |
| E_INVALIDARG | Invalid safe array descriptor. |
| E_FAIL | Failure. The array descriptor is null. |

# <a name="Remove"></a>Remove

Removes the specified array element.

```
FUNCTION Remove (BYVAL nPos AS LONG) AS HRESULT
FUNCTION RemoveStr (BYVAL nPos AS LONG) AS HRESULT
FUNCTION RemoveVar (BYVAL nPos AS LONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *nPos* | Index of the array element which will be deleted. |

#### Return value

S_OK (0) on success or an HRESULT code on failure.

| HRESULT    | Description |
| ---------- | ----------- |
| DISP_E_BADINDEX | The specified index is not valid. |
| E_INVALIDARG | One of the arguments is not valid. |
| E_OUTOFMEMORY | Memory could not be allocated for the element. |
| E_FAIL | Failure. The array descriptor is null. |

# <a name="SetIID"></a>SetIID

Sets the IID (a GUID) of the interface contained within a given safe array.

```
FUNCTION SetIID (BYVAL pguid AS GUID PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pguid* | Pointer to the IID. |

#### Return value

S_OK (0) on success or an HRESULT code on failure.

| HRESULT    | Description |
| ---------- | ----------- |
| E_INVALIDARG | If the array descriptor does not have the FADF_HAVEIID flag set. |
| E_FAIL | Failure. The array descriptor is null. |

# <a name="SetRecordInfo"></a>SetRecordInfo

Sets the IRecordInfo interface of the UDT contained in a given safe array.

```
FUNCTION SetRecordInfo (BYVAL prinfo AS IRecordInfo PTR) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *prinfo* | Pointer to an IRecordInfo interface. |

#### Return value

S_OK (0) on success or an HRESULT code on failure.

| HRESULT    | Description |
| ---------- | ----------- |
| E_INVALIDARG | If the array descriptor does not have the FADF_RECORD flag set. |
| E_FAIL | Failure. The array descriptor is null. |

# <a name="Sort"></a>Sort

Sorts a one-dimensional VT_BSTR CSafeArray calling the C **qsort** function.

```
FUNCTION Sort (BYVAL bAscend AS BOOLEAN = TRUE) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *bAscend* | TRUE for sorting in ascending order; FALSE for sorting in descending order. |

#### Return value

S_OK (0) on success or an HRESULT code on failure.

| HRESULT    | Description |
| ---------- | ----------- |
| E_FAIL | The array descriptor is null or the safe array is not of of the type VT_BSTR. |
| E_UNEXPECTED | The array could not be locked. |

#### Usage example

```
' // Create a one-dimensional array of BSTR
DIM csa AS CSafeArray = CSafeArray(VT_BSTR, 3, 1)

DIM cbsVal AS CBSTR = "bcde"
csa.Put(1, cbsVal)
cbsVal = "abc"
csa.Put(2, cbsVal)
cbsVal = "abcfg"
csa.Put(3, cbsVal)
' // Sort the safe array
csa.Sort

csa.Get(1, cbsVal)
print cbsVal
csa.Get(2, cbsVal)
print cbsVal
csa.Get(3, cbsVal)
print cbsVal

--- or ---

print csa.GetStr(1)
print csa.GetStr(2)
print csa.GetStr(3)
```

# <a name="UnaccessData"></a>UnaccessData

Decrements the lock count of an array, and invalidates the pointer retrieved by **AccessData**.

```
FUNCTION UnaccessData () AS HRESULT
```

#### Return value

S_OK (0) on success or an HRESULT code on failure.
