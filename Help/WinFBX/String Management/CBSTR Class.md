# CBSTR Class

The **CBSTR** class implements a dynamic unicode null terminated string. Free Basic has a dynamic string data type (STRING) and a fixed length unicode data type (WSTRING), but it lacks a dynamic unicode string. **CBSTR** behaves as if it was a native data type, working directly with the intrinsic Free Basic string functions and operators.

**Include file**: CWSTR.INC.

Quirks:

* MID as a function: Something like MID(cbs, 2) doesn't work with languages such Russian and Chinese. Using MID(\*\*cbs, 2), MID(cbs.wstr, 2) or cbs.MidChars(2) works.
* MID as a statement: Something like MID(cbs, 2, 1) = "x" compiles but does not change the contents of the dynamic unicode string. MID(\*\*cbs, 2, 1) = "x" works.
* TRIM, LTRIM, RTRIM, LSET and RSET don't work with languages like Russian or Chinese unless we prepend \*\* to the variable name.
* SELECT CASE: Something like SELECT CASE LEFT(cbs, 2) does not compile; we have to use SELECT CASE LEFT(\*\*cbs, 2).

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#Constructors) | Initialize the class with the specified value. |
| [Operator \*](#Operator*) | One * returns the address of the CBSTR buffer.<br> Two ** returns the address of the start of the string data. |
| [bptr](#bptr) | Returns the underlying BSTR pointer. |
| [sptr](#sptr) | Returns the address of the string data. Same as \*\*. |
| [vptr](#vptr) | Frees the underlying BSTR and returns the BSTR pointer. |
| [wstr](#wstr) | Returns the string data. Same as \*\*. |
| [Operator &](#Operator&) | Concatenates strings. |
| [Operator +=](#Operator+=) | Appends a string to the CBSTR. |
| [Operator &=](#Operator&=) | Appends a string to the CBSTR. |
| [Operator []](#Operator[]) | Gets or sets the corresponding unicode integer representation of the character at the specified position. |
| [Operator Let](#OperatorLet) | Assigns a string to the CBSTR. It implements the = operator. |
| [Operator Cast](#OperatorCast) | Returns a pointer to the CBSTR buffer or the string data.<br>Casting is automatic. You don't have to call this operator. |
| [wstr](#wstr) | Returns the string data. Same as \*\*. |
| [wchar](#wchar) | Returns the string data as a new unicode string allocated with CoTaskMemAlloc. |
| [Utf8](#Utf8) | Converts from UTF8 to Unicode and from Unicode to UTF8. |
| [Attach](#Attach) | Attaches a BSTR to the CBSTR class. |
| [Detach](#Detach) | Detaches the underlying BSTR from the CBSTR class and returns it as the result of the function. |
| [Append](#Append) | Appends a string to the CBSTR. |
| [Clear](#Clear) | Frees the underlying BSTR. |
| [Empty](#Empty) | Frees the underlying BSTR. |
| [Left](#Left) | Returns the leftmost substring of the string. Same as LEFT. |
| [Right](#Right) | Returns the rightmost substring of the string. Same as RIGHT. |
| [Char](#Char) | Gets or sets the corresponding unicode integer representation of the character at the specified position. |
| [LeftChars](#LeftChars) | Returns the leftmost substring of the string. Same as Left. |
| [MidChars](#MidChars) | Returns a substring of the string. Same as Mid. |
| [RightChars](#RightChars) | Returns the rightmost substring of the string. Same as Right. |
| [Val](#Val) | Converts the string to a floating point number (DOUBLE). |
| [ValDouble](#ValDouble) | Converts the string to a floating point number (DOUBLE). |
| [ValInt](#ValInt) | Converts the string to a signed 32-bit integer (LONG). |
| [ValLong](#ValLong) | Converts the string to a signed 32-bit integer (LONG). |
| [ValLongInt](#ValLongInt) | Converts the string to a signed 64-bit integer (LONGINT). |
| [Value](#Value) | Converts the string to a floating point number (DOUBLE). |
| [ValUInt](#ValUInt) | Converts the string to a 32.bit unsigned integer (ULONG). |
| [ValULong](#ValULong) | Converts the string to a 32-bit unsigned integer (ULONG). |
| [ValULongInt](#ValULongInt) | Converts the string to a 64-bit unsigned integer (ULONGINT). |

# <a name="Constructors"></a>Constructors

Initialize the class with the specified value.

```
CONSTRUCTOR CBStr
CONSTRUCTOR CBStr (BYREF wszStr AS CONST WSTRING = "")
CONSTRUCTOR CBStr (BYREF cws AS CWSTR)
CONSTRUCTOR CBStr (BYREF cbs AS CBSTR)
CONSTRUCTOR CBStr (BYREF ansiStr AS STRING, BYVAL nCodePage AS UINT = 0)
CONSTRUCTOR (BYVAL n AS LONGINT)
CONSTRUCTOR (BYVAL n AS DOUBLE)
CONSTRUCTOR (BYREF bstrHandle AS AFX_BSTR = NULL, BYVAL fAttach AS LONG = TRUE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszStr* | A WSTRING. |
| *cws* | A CWSTR. |
| *cbs* | A CBSTR. |
| *ansiStr* | An ansi string or string literal. |
| *nCodePage* | The code page to be used for ansi to unicode conversions. |
| *n* | A number. |
| *bstrHandle* | A handle to a BSTR.. |
| *fAttach* | TRUE = Attaches the BSTR handle to the class. It will be released when the class is destroyed; therefore, don't free the passed handle. FALSE = The contents of the passed BSTR are copied.. |

For a list of code pages see: [Code Page Identifiers](https://msdn.microsoft.com/en-us/library/windows/desktop/dd317756(v=vs.85).aspx)

# Operators

## <a name="Operator*"></a>Operator *

Deferences the CBSTR.<br>One * returns the address of the underlying BSTR pointer.<br>Two ** returns the address of the start of the string data.

```
OPERATOR * (BYREF cbs AS CBSTR) AS AFX_BSTR
```

## <a name="bptr"></a>bptr

Returns the underlying BSTR pointer.

```
FUNCTION bptr () BYREF AS AFX_BSTR
```

## <a name="sptr"></a>sptr

Returns the address of the CBSTR string data (same as **)

```
FUNCTION sptr () AS WSTRING PTR
```

## <a name="vptr"></a>vptr

Frees the underlying BSTR and returns the BSTR pointer.

```
FUNCTION vptr () AS AFX_BSTR PTR
```

Must be used to pass the underlying BSTR to an OUT BYVAL BSTR PTR parameter. If we pass a CBSTR to a function with an OUT BSTR parameter without first freeing it we will have a memory leak.

## <a name="wstr"></a>wstr

Returns the string data. Same as **.

```
FUNCTION wstr () BYREF AS CONST WSTRING
```

## <a name="Operator&"></a>Operator &

Concatenates strings.

```
OPERATOR & (BYREF cbs1 AS CBSTR, BYREF cbs2 AS CBSTR) AS CBSTR
```

## <a name="Operator+="></a>Operator +=

Appends a string to the CBSTR.

```
OPERATOR += (BYREF wszStr AS CONST WSTRING)
OPERATOR += (BYREF cws AS CWStr)
OPERATOR += (BYREF cbs AS CBStr)
OPERATOR += (BYREF ansiStr AS STRING)
```

## <a name="Operator&="></a>Operator &=

Appends a string to the CBSTR.

```
OPERATOR &= (BYREF wszStr AS CONST WSTRING)
OPERATOR &= (BYREF cws AS CWStr)
OPERATOR &= (BYREF cbs AS CBStr)
OPERATOR &= (BYREF ansiStr AS STRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszStr* | A WSTRING. |
| *cws* | A CWSTR. |
| *cbs* | A CBSTR. |
| *ansiStr* | An ansi string or string literal. |

## <a name="Operator[]"></a>Operator []

Gets or sets the corresponding unicode integer representation of the character at the specified position. The index parameter is zero based ((0 for the first character, 1 for the second, etc.). This operator must not be used in case of empty string because reference is undefined (inducing runtime error). Otherwise, the user must ensure that the index does not exceed the range "\[0, Len(cbs) - 1]". Outside this range, results are undefined.

```
OPERATOR [] (BYVAL nIndex AS LONG) BYREF AS USHORT
```
#### Example
```
DIM cbs as CBSTR = "1234567890"
print cbs[1]
cbs[1] = ASC("X")
print cbs
```

## <a name="OperatorLet"></a>Operator Let

Assigns a string to the CBSTR.

```
OPERATOR LET (BYREF wszStr AS CONST WSTRING)
OPERATOR LET (BYREF cws AS CWStr)
OPERATOR LET (BYREF cbs AS CBStr)
OPERATOR LEY (BYREF ansiStr AS STRING)
OPERATOR LET (BYVAL n AS LONGINT)
OPERATOR LET (BYVAL n AS DOUBLE)
OPERATOR LET (BYREF bstrHandle AS AFX_BSTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszStr* | A WSTRING. |
| *cws* | A CWSTR. |
| *cbs* | A CBSTR. |
| *ansiStr* | An ansi string or string literal. |
| *n* | A number. |
| *bstrHandle* | A handle to a BSTR. Don't free it, because it will be attached. |

# Casting and Conversions

## <a name="OperatorCast"></a>Operator Cast

Return a pointer to the underlying BSTR or the string data. These operators aren't called directly.

```
OPERATOR CAST () BYREF AS CONST WSTRING
OPERATOR CAST () AS ANY PTR
```

## <a name="wchar"></a>wchar

Returns the string data as a new unicode string allocated with CoTaskMemAlloc.
Useful when we need to pass a pointer to a null terminated wide string to a function or method that will release it. If we pass a WSTRING it will GPF. If the length of the input string is 0, CoTaskMemAlloc allocates a zero-length item and returns a valid pointer to that item. If there is insufficient memory available, CoTaskMemAlloc returns NULL.

```
FUNCTION wchar () AS WSTRING PTR
```

## <a name="Utf8"></a>Utf8

Converts from UTF8 to Unicode and from Unicode to UTF8.

```
PROPERTY Utf8() AS STRING
PROPERTY Utf8 (BYREF utf8String AS STRING)
```

# Methods

## <a name="Attach"></a>Attach

Attaches a BSTR to the CBSTR class.

```
SUB Attach (BYVAL pbstrSrc AS AFX_BSTR)
```

## <a name="Detach"></a>Detach

Detaches the underlying BSTR from the CBSTR class and returns it as the result of the function. The returned pointer must be freed by SysFreeString.

```
FUNCTION Detach () AS AFX_BSTR
```

This method frees the *m_bstr* member of the CBSTR class and returns it as the result of the function. Because it no longer belongs to the class, it must be freed by SysFreeString.

## <a name="Append"></a>Append

Appends a string to the CBSTR. 

```
SUB Append (BYREF wszStr AS CONST WSTRING)
```

*Remark*: The string can be a literal or a FB STRING, a WSTRING, a CWSTR or a CBSTR variable.

## <a name="Clear"></a>Clear

Frees the underlying BSTR.

```
SUB Clear
```

## <a name="Empty"></a>Empty

Frees the underlying BSTR.

```
SUB Empty
```

## <a name="Left"></a>Left

Returns the leftmost substring of the string.

```
FUNCTION Left OVERLOAD (BYREF cws AS CBSTR, BYVAL nChars AS INTEGER) AS CBSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbs* | The source CBSTR. |
| *nChars* | The number of characters to return from the source string. |

## <a name="Right"></a>Right

Returns the rightmost substring of the string.

```
FUNCTION Right OVERLOAD (BYREF cbs AS CBSTR, BYVAL nChars AS INTEGER) AS CBSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbs* | The source CBSTR. |
| *nChars* | The substring length, in characters. |

## <a name="Char"></a>Char

Gets or sets the corresponding unicode integer representation of the character at the position specified by the *nIndex* parameter.

```
PROPERTY Char (BYVAL nIndex AS UINT) AS USHORT
PROPERTY Char (BYVAL nIndex AS UINT, BYVAL nValue AS USHORT)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nIndex* | The one based index of the character in the string (1 for the first character, 2 for the second, etc.). If nIndex is beyond the current length of the string, a 0 is returned. |
| *nValue* | The unicode integer representation of the character. |

## <a name="LeftChars"></a>LeftChars

Returns the leftmost substring of the string.

```
FUNCTION LeftChars (BYVAL nChars AS LONG) AS CBSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nChars* | The number of characters to return from the source string. |

## <a name="MidChars"></a>MidChars

Returns a substring of the string.

```
MidChars (BYVAL nStart AS LONG, BYVAL nChars AS LONG = 0) AS CBSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nStart* | The start position in CBSTR of the substring. The first character starts at position 1. |
| *nChars* | The substring length, in characters. If nChars < 0 or nChars >= length of the CBSTR then all of the remaining characters are returned. |

If CBSTR is empty then the null string ("") is returned. If *nStart* <= 0 then the null string ("") is returned.

## <a name="RightChars"></a>RightChars

Returns the rightmost substring of the string.

```
FUNCTION RightChars (BYVAL nChars AS LONG) AS CBSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nChars* | The substring length, in characters. |

## <a name="Val"></a>Val

Converts the string to a floating point number (DOUBLE).

```
FUNCTION Val OVERLOAD (BYREF cbs AS CBSTR) AS DOUBLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbs* | The source CBSTR. |

## <a name="ValDouble"></a>ValDouble

Converts the string to a floating point number (DOUBLE).

```
FUNCTION ValDouble () AS DOUBLE
```

## <a name="ValInt"></a>ValInt

Converts the string to a signed 32-bit integer (LONG).

```
FUNCTION ValInt () AS LONG
```

## <a name="ValLong"></a>ValLong

Converts the string to a signed 32-bit integer (LONG).

```
FUNCTION ValLong () AS LONG
```

## <a name="ValLongInt"></a>ValLongInt

Converts the string to a signed 64-bit integer (LONGINT).

```
FUNCTION ValLongInt () AS LONGINT
```

## <a name="Value"></a>Value

Converts the string to a floating point number (DOUBLE).

```
FUNCTION Value () AS DOUBLE
```

## <a name="ValUInt"></a>ValUInt

Converts the string to a 32.bit unsigned integer (ULONG).

```
FUNCTION ValUInt () AS ULONG
```

## <a name="ValULong"></a>ValULong

Converts the string to a 32-bit unsigned integer (ULONG).

```
FUNCTION ValULong () AS ULONG
```

## <a name="ValULongInt"></a>ValULongInt

Converts the string to a 64-bit unsigned integer (ULONGINT).

```
FUNCTION ValULongInt () AS ULONGINT
```
