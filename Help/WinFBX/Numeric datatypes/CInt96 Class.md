# CInt96 Class

`CInt96` is a wrapper class for the DECIMAL data type. Holds signed 128-bit (16-byte) values representing 96-bit (12-byte) integer numbers. It supports up to 29 significant digits and can represent values in excess of 7.9228 x 10^28. The largest possible value is +/-79,228,162,514,264,337,593,543,950,335.

### Constructors

Creates a new instance of the `CInt96` class and assigns the values passed to it.

```
CONSTRUCTOR CInt96
CONSTRUCTOR CInt96 (BYREF cSrc AS CInt96)
CONSTRUCTOR CInt96 (BYREF decSrc AS DECIMAL)
CONSTRUCTOR CInt96 (BYVAL nInteger AS LONGINT)
CONSTRUCTOR CInt96 (BYVAL nUInteger AS ULONGINT)
CONSTRUCTOR CInt96 (BYREF wszSrc AS WSTRING)
```

#### Usage example

```
DIM int96 AS CInt96 = 1234567890
```

#### Remarks

Because the bigger numeric variable natively supported by Free Basic is a long integer, if we want to set bigger values we need to use strings, e.g.

```
DIM int96 AS CInt96 = "79228162514264337593543950335"
```

### Operators

| Name       | Description |
| ---------- | ----------- |
| [Operator LET](#Operator1) | Assigns a value to a `CInt96`variable. |
| [CAST operators](#Operator2) | Converts a `CInt96` into another data type. |
| [Operator \*](#Operator3) | Returns the address of the underlying `DECIMAL` structure. |
| [Comparison operators](#Operator4) | Compares `CInt96` numbers. |
| [Math operators](#Operator5) | Add, subtract, multiply or divide currency numbers. |

### Methods

| Name       | Description |
| ---------- | ----------- |
| [Abs_](#Abs_) | Returns the absolute value of a decimal data type. |
| [IsSigned](#IsSigned) | Returns true if this number is signed or false otherwise. |
| [IsUnsigned](#IsUnsigned) | Returns true if this number is unsigned or false otherwise. |
| [Sign](#Sign) | Returns 0 if it is not signed of &h80 (128) if it is signed. |
| [ToVar](#ToVar) | Returns the currency as a VT_CY variant. |

# Error Handling

The class uses the API function `SetLastError` to set error information. After calling a method of the class, you can check its success or failure calling the API function `GetLastError`. To get a localized description of the error, you can call a function such `AfxGetWinErrMsg`, passing the error code returned by `GetLastError`:

```
PRIVATE FUNCTION AfxGetWinErrMsg (BYVAL dwError AS DWORD) AS CWSTR
   DIM cbLen AS DWORD, pBuffer AS WSTRING PTR, cwsMsg AS CWSTR
   cbLen = FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER OR _
           FORMAT_MESSAGE_FROM_SYSTEM OR FORMAT_MESSAGE_IGNORE_INSERTS, _
           NULL, dwError, BYVAL MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), _
           cast(LPWSTR, @pBuffer), 0, NULL)
   IF cbLen THEN
      cwsMsg = *pBuffer
      LocalFree pBuffer
   END IF
   RETURN cwsMsg
END FUNCTION
```

# <a name="Operator1"></a>Operator LET (=)

Assigns a value to a `CInt96` variable.

```
OPERATOR LET (BYREF cSrc AS CInt96)
OPERATOR LET (BYVAL nInteger AS LONGINT)
OPERATOR LET (BYVAL cySrc AS CURRENCY)
OPERATOR LET (BYVAL nUInteger AS ULONGINT)
OPERATOR LET (BYREF wszSrc AS WSTRING)
```

#### Examples

```
DIM int96 AS CInt96
int96 = 1234567890
-or-
DIM int96 AS CInt96 = 1234567890
-or-
DIM int96 AS CInt96 = "1234567890"
```

# <a name="Operator2"></a>CAST Operators

Converts a `Cint16` into another data type.

```
OPERATOR CAST () AS BYTE
OPERATOR CAST () AS UBYTE
OPERATOR CAST () AS SHORT
OPERATOR CAST () AS USHORT
OPERATOR CAST () AS LONG
OPERATOR CAST () AS ULONG
OPERATOR CAST () AS LONGINT
OPERATOR CAST () AS ULONGINT
OPERATOR CAST () AS SINGLE
OPERATOR CAST () AS DOUBLE
OPERATOR CAST () AS CURRENCY
OPERATOR CAST () AS VARIANT
OPERATOR CAST () AS STRING
```

#### Remarks

These operators aren't called directly, They perform the conversion when the target is not a `CInt96` variable.

# <a name="Operator3"></a>Operator *

Returns the address of the underlying `DECIMAL` structure.

```
OPERATOR * (BYREF int96 AS CInt96) AS DECIMAL PTR
```

# <a name="Operator4"></a>Comparison operators

```
OPERATOR = (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS BOOLEAN
OPERATOR <> (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS BOOLEAN
OPERATOR < (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS BOOLEAN
OPERATOR > (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS BOOLEAN
OPERATOR <= (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS BOOLEAN
OPERATOR >= (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS BOOLEAN
```

# <a name="Operator5"></a>Math operators

```
OPERATOR + (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS CInt96
OPERATOR + (BYREF int96 AS CInt96, BYREF wszValue AS WSTRING) AS CInt96
OPERATOR + (BYREF int96 AS CInt96, BYVAL nValue AS LONGINT) AS CInt96
OPERATOR + (BYREF int96 AS CInt96, BYVAL nValue AS ULONGINT) AS CInt96
OPERATOR + (BYVAL nValue AS LONGINT, BYREF int96 AS CInt96) AS CInt96
OPERATOR + (BYVAL nValue AS ULONGINT, BYREF int96 AS CInt96) AS CInt96
OPERATOR += (BYREF int96 AS CInt96)
OPERATOR += (BYVAL nValue AS LONGINT)
OPERATOR += (BYVAL nValue AS ULONGINT)
OPERATOR += (BYVAL wszValue AS WSTRING)
OPERATOR - (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS CInt96
OPERATOR - (BYREF int96 AS CInt96, BYVAL nValue AS LONGINT) AS CInt96
OPERATOR - (BYREF int96 AS CInt96, BYVAL nValue AS ULONGINT) AS CInt96
OPERATOR - (BYVAL nValue AS LONGINT, BYREF int96 AS CInt96) AS CInt96
OPERATOR - (BYVAL nValue AS ULONGINT, BYREF int96 AS CInt96) AS CInt96
OPERATOR - (BYREF int96 AS CInt96, BYVAL wszValue AS WSTRING) AS CInt96
OPERATOR -= (BYREF int96 AS CInt96)
OPERATOR -= (BYREF nValue AS LONGINT)
OPERATOR -= (BYREF nValue AS ULONGINT)
OPERATOR -= (BYREF wszValue AS WSTRING)
OPERATOR * (BYREF int961 AS CInt96, BYREF int962 AS CInt96) AS CInt96
OPERATOR * (BYREF int96 AS CInt96, BYVAL nOperand AS LONGINT) AS CInt96
OPERATOR * (BYREF int96 AS CInt96, BYVAL nOperand AS ULONGINT) AS CInt96
OPERATOR * (BYVAL nOperand AS LONGINT, BYREF int96 AS CInt96) AS CInt96
OPERATOR * (BYVAL nOperand AS ULONGINT, BYREF int96 AS CInt96) AS CInt96
OPERATOR * (BYREF int96 AS CInt96, BYREF wszValue AS WSTRING) AS CInt96
OPERATOR *= (BYREF int96 AS CInt96)
OPERATOR *= (BYVAL nOperand AS LONGINT)
OPERATOR *= (BYVAL nOperand AS ULONGINT)
OPERATOR *= (BYVAL wszValue AS WSTRING)
OPERATOR / (BYREF int96 AS CInt96, BYVAL cOperand AS CInt96) AS CInt96
OPERATOR / (BYREF int96 AS CInt96, BYVAL nOperand AS LONGINT) AS CInt96
OPERATOR / (BYREF int96 AS CInt96, BYVAL nOperand AS ULONGINT) AS CInt96
OPERATOR / (BYVAL nValue AS LONGINT, BYREF int96 AS CInt96) AS CInt96
OPERATOR / (BYVAL nValue AS ULONGINT, BYREF int96 AS CInt96) AS CInt96
OPERATOR / (BYREF int96 AS CInt96, BYREF wszValue AS WSTRING) AS CInt96
OPERATOR /= (BYREF cOperand AS CInt96)
OPERATOR /= (BYVAL nOperand AS LONGINT)
OPERATOR /= (BYVAL nOperand AS ULONGINT)
OPERATOR /= (BYREF wszValue AS WSTRING)
```

#### Examples

```
DIM int96 AS CInt96 = 1234567890
int96 = int96 + 111
int96 = int96 - 111
int96 = int96 * 2
int96 = int96 / 2
int96 += 123
int96 -= 123
int96 *= 2
int96 /= 2
```

#### Remarks

The version OPERATOR - (BYREF int96 AS CInt96) AS CInt96 changes the sign of a decimal number.

```
DIM int96 AS CInt96 = 123
int96 = -int96
```

# <a name="Abs_"></a>Abs_

Returns the absolute value of a decimal data type.

```
FUNCTION Abs_ () AS CInt96
```

# <a name="IsSigned"></a>IsSigned

Returns true if this number is signed or false otherwise.

```
FUNCTION IsSigned () AS BOOLEAN
```

# <a name="IsUnsigned"></a>IsUnsigned

Returns true if this number is unsigned or false otherwise.

```
FUNCTION IsUnsigned () AS BOOLEAN
```

# <a name="Sign"></a>Sign

Returns 0 if it is not signed of &h80 (128) if it is signed.

```
FUNCTION Sign () AS UBYTE
```

# <a name="ToVar"></a>ToVar

Returns the DECIMAL as a VT_DECIMAL variant.

```
FUNCTION ToVar () AS VARIANT
```

#### Remarks

Can be used to assign a currency directly to a CVAR variable.

```
DIM it96 AS CInt95 = 1234567890
DIM cv AS CVAR = int96.ToVar
```
