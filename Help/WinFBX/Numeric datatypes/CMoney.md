# CMoney Class

**CMoney** is especially useful for financial calculations, as it avoids the round-off errors associated with floating point data types. The **WinFBX** framework also provides **CCUR**, that implements a currency data type with 4 digits of precision after the decimal point. **CMoney** allows te use of 0 to 9 digits of precision (the default value is 2 digits) for a greater flexibility (some currencies have 0 decimal places, e.g. the Japanese yen, and some three decimal places, e.g. the Kuwaiti dinar). The number of decimal places defaults to 2, but you can choose between 0-9 by defining them before includind CMoney.inc, e.g. #define AFX_CMONEY_DECIMAL_PLACES 4. If you choose more than 9 decimal places, the output will be truncated. The internal storage of **CCUR** and **CMoney** differs substantially: **CCUR** uses a CURRENCY structure and **CMoney** a DECIMAL structure.

When using 2 decimal digits, the greatest possible value is +/-792,281,625,142,643,375,935,439,503.35. However, the formatting functions, **FormatCurrency** and **FormatNumber** can only return a maximum value of +/-922,337,203,685,477.58, which are more than enough for financial calculations, because they use a currency data type. For greater values, you can format the returned string by yourself.

The locale identifier defaults to 1033 (US), for consistency with FreeBasic numeric data types, but you can change it by defining it before including CMoney.inc, e.g. #define AFX_CMONEY_LCID 3082 (Spain). The locale identifier (LCID) instructs the methods of this class about how input and output strings should be treated. For example, if you choose the Spanish LCID, the formated string input and output will use a comma, instead of a point, as the dcimal separator, and a point, instead of a comma, as the thousands separator.

### Constructors

Creates a new instance of the **CMoney** class and assigns the value passed to it.

```
CONSTRUCTOR CMoney
CONSTRUCTOR (BYREF cSrc AS CMoney)
CONSTRUCTOR (BYVAL nInteger AS LONGINT)
CONSTRUCTOR (BYVAL nUInteger AS ULONGINT)
CONSTRUCTOR (BYREF decSrc AS DECIMAL)
CONSTRUCTOR (BYVAL cy AS CURRENCY)
CONSTRUCTOR (BYVAL nValue AS DOUBLE)
CONSTRUCTOR (BYREF wszSrc AS WSTRING)
```

#### Remarks

We can also use strings to set the value, e.g.:

```
DIM money AS CMoney = "1234567890.12"
```
--or--
```
DIM money AS CMoney = "1,234,567,890.12"
```

When using a locale identifier other than the default one, we may need to use different separators for the thousands and the decimal point. For example, if using the locale identifier for Spain we need to use "," as the decimal separator and "." as the thousands separator.

```
DIM money AS CMoney = "1.234.567.890,12"
```

### Operators

| Name       | Description |
| ---------- | ----------- |
| [Operator LET](#Operator1) | Assigns a value to a **CMoney** variable. |
| [CAST operators](#Operator2) | Converts a **CMoney** into another data type. |
| [Operator \*](#Operator3) | Returns the address of the underlying **DECIMAL** structure. |
| [Comparison operators](#Operator4) | Compares **CMoney** numbers. |
| [Math operators](#Operator5) | Add, subtract, multiply or divide **CMoney** numbers. |

### Methods

| Name       | Description |
| ---------- | ----------- |
| [Abs_](#CyAbs) | Returns the absolute value of a money data type. |
| [FormatCurrency](#FormatCurrency) | Formats a currency into a string form. |
| [FormatNumber](#FormatNumber) | Formats a currency into a string form. |
| [IsSigned](#IsSigned) | Returns true if this number is signed or false otherwise. |
| [IsUnsigned](#IsUnsigned) | Returns true if this number is unsigned or false otherwise. |
| [Scale](#Scale) | Returns the scale of the money number. |
| [Sign](#Sign) | Returns 0 if it is not signed of &h80 (128) if it is signed. |
| [ToCyVar](#ToCyVar) | Returns the money number as a VT_CURRENCY variant. |
| [ToVar](#ToVar) | Returns the money number as a VT_DECIMAL variant. |

# <a name="Operator1"></a>Operator LET (=)

Assigns a value to a **CMoney** variable.

#### Examples

```
DIM money AS CMoney
money = 123
money = 12345.12
```

# <a name="Operator2"></a>CAST Operators

Converts a **CMoney** into another data type.

```
OPERATOR CAST () AS DOUBLE
OPERATOR CAST () AS CURRENCY
OPERATOR CAST () AS DECIMAL
OPERATOR CAST () AS VARIANT
OPERATOR CAST () AS STRING
```

#### Remarks

These operators aren't called directly, They perform the conversion when the target is not a **CMoney** variable.

# <a name="Operator3"></a>Operator *

Returns the address of the underlying **DECIMAL** structure.

```
OPERATOR * (BYREF money AS CMoney) AS DECIMAL PTR
```

# <a name="Operator4"></a>Comparison operators

Compares **CMoney** numbers.

```
OPERATOR = (BYREF money1 AS CMoney, BYREF money2 AS CMoney) AS BOOLEAN
OPERATOR <> (BYREF money1 AS CMoney, BYREF money2 AS CMoney) AS BOOLEAN
OPERATOR < (BYREF money1 AS CMoney, BYREF money2 AS CMoney) AS BOOLEAN
OPERATOR > (BYREF money1 AS CMoney, BYREF money2 AS CMoney) AS BOOLEAN
OPERATOR <= (BYREF money1 AS CMoney, BYREF money2 AS CMoney) AS BOOLEAN
OPERATOR >= (BYREF money1 AS CMoney, BYREF money2 AS CMoney) AS BOOLEAN
```

# <a name="Operator5"></a>Math operators

Add, subtract, multiply or divide **CMoney** numbers.

```
OPERATOR + (BYREF money1 AS CMoney, BYREF money2 AS CMoney) AS CMoney
OPERATOR += (BYREF money AS CMoney)
OPERATOR - (BYREF money1 AS CMoney, BYREF money2 AS CMoney) AS CMoney
OPERATOR - (BYREF money AS CMoney) AS CMoney
OPERATOR -= (BYREF money AS CMoney)
OPERATOR * (BYREF money1 AS CMoney, BYREF money2 AS CMoney) AS CMoney
OPERATOR *= (BYREF money AS CMoney)
OPERATOR / (BYREF money1 AS CMoney, BYREF money2 AS CMoney) AS CMoney
OPERATOR /= (BYREF money AS CMoney)
```

#### Examples

```
DIM money AS CMoney = 12345.12
money = money + 111.11
money = money - 111.11
money = money * 2
money = money / 2
money += 123
money -= 123
money *= 2.3
money /= 2.3
```

#### Remarks

The version OPERATOR - (BYREF money AS CMoney) AS CMoney changes the sign of a decimal number.

```
DIM money AS CMoney = 123
money = -money
```

# <a name="CyAbs"></a>Abs_

Returns the absolute value of a money data type.

```
FUNCTION Abs_ () AS CMoney
```
# <a name="FormatCurrency"></a>FormatCurrency

Formats a currency into a string form.

```
FUNCTION FormatCurrency (BYVAL iNumDig AS LONG = -1, BYVAL ilncLead AS LONG = -2, _
   BYVAL iUseParens AS LONG = -2, BYVAL iGroup AS LONG = -2, BYVAL dwFlags AS DWORD = 0) AS STRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *iNumDig* | The number of digits to pad to after the decimal point. Specify -1 to use the system default value. |
| *ilncLead* | Specifies whether to include the leading digit on numbers.<br>-2 : Use the system default.<br>-1 : Include the leading digit.<br> 0 : Do not include the leading digit. |
| *iUseParens* | Specifies whether negative numbers should use parentheses.<br>-2 : Use the system default.<br>-1 : Use parentheses.<br> 0 : Do not use parentheses. |
| *iGroup* | Specifies whether thousands should be grouped. For example 10,000 versus 10000.<br>-2 : Use the system default.<br>-1 : Group thousands.<br> 0 : Do not group thousands. |
| *dwFlags* | VAR_CALENDAR_HIJRI is the only flag that can be set.  |

#### Return value

A string containing the formatted value.

#### Remarks

Same as **FormatNumber** but adding the currency symbol.

#### Example

```
DIM c AS CCUR = 12345.1234
PRINT c.FormatCurrency   --> 12.345,12 € (Spain)
```

# <a name="FormatNumber"></a>FormatNumber

Formats a currency into a string form.

```
FUNCTION FormatNumer (BYVAL iNumDig AS LONG = -1, BYVAL ilncLead AS LONG = -2, _
   BYVAL iUseParens AS LONG = -2, BYVAL iGroup AS LONG = -2, BYVAL dwFlags AS DWORD = 0) AS STRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *iNumDig* | The number of digits to pad to after the decimal point. Specify -1 to use the system default value. |
| *ilncLead* | Specifies whether to include the leading digit on numbers.<br>-2 : Use the system default.<br>-1 : Include the leading digit.<br> 0 : Do not include the leading digit. |
| *iUseParens* | Specifies whether negative numbers should use parentheses.<br>-2 : Use the system default.<br>-1 : Use parentheses.<br> 0 : Do not use parentheses. |
| *iGroup* | Specifies whether thousands should be grouped. For example 10,000 versus 10000.<br>-2 : Use the system default.<br>-1 : Group thousands.<br> 0 : Do not group thousands. |
| *dwFlags* | VAR_CALENDAR_HIJRI is the only flag that can be set.  |

#### Return value

A string containing the formatted value.

#### Remarks

Same as **FormatCurrency** but without adding the currency symbol.

#### Example

```
DIM c AS CCUR = 12345.1234
PRINT c.FormatNumber   --> 12.345,12 (Spain)
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

# <a name="Scale"></a>Scale

Returns the scale of the money number.

```
FUNCTION Scale () AS UBYTE
```

# <a name="Sign"></a>Sign

Returns 0 if it is not signed of &h80 (128) if it is signed.

```
FUNCTION Sign () AS UBYTE
```

# <a name="ToCyVar"></a>ToCyVar

Returns the money number as a VT_CURRENCY variant.

```
FUNCTION ToCyVar () AS VARIANT
```

# <a name="ToVar"></a>ToVar

Returns the money number as a VT_DECIMAL variant.

```
FUNCTION ToVar () AS VARIANT
```
