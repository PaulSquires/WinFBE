# CCUR Class

`CCUR` is a wrapper class for the `CURRENCY` data type. `CURRENCY` is implemented as an 8-byte two's-complement integer value scaled by 10,000. This gives a fixed-point number with 15 digits to the left of the decimal point and 4 digits to the right. The `CURRENCY` data type is extremely useful for calculations involving money, or for any fixed-point calculations where accuracy is important.

The `CCUR` wrapper implements arithmetic, assignment, and comparison operations for this fixed-point type, and provides access to the numbers on either side of the decimal point in the form of two components: an integer component which stores the value to the left of the decimal point, and a fractional component which stores the value to the right of the decimal point. The fractional component is stored internally as an integer value between -9999 (CY_MIN_FRACTION) and +9999 (CY_MAX_FRACTION). The function `GetFraction` returns a value scaled by a factor of 10000 (CY_SCALE).

When specifying the integer and fractional components of a `CCUR` object, remember that the fractional component is a number in the range 0 to 9999. This is important when dealing with a currency such as the US dollar that expresses amounts using only two significant digits after the decimal point. Even though the last two digits are not displayed, they must be taken into account.

### Constructors

Create a new instance of the `CCUR` class and assigns the values passed to it.

```
CONSTRUCTOR CCur
CONSTRUCTOR CCur (BYREF cSrc AS CCUR)
CONSTRUCTOR CCur (BYVAL cySrc AS CURRENCY)
CONSTRUCTOR CCur (BYVAL nInteger AS LONGLONG)
CONSTRUCTOR CCur (BYVAL nInteger AS LONGLONG, BYVAL nFraction AS SHORT)
CONSTRUCTOR CCur (BYVAL bSrc AS BYTE)
CONSTRUCTOR CCur (BYVAL ubSrc AS UBYTE)
CONSTRUCTOR CCur (BYVAL sSrc AS SHORT)
CONSTRUCTOR CCur (BYVAL usSrc AS USHORT)
CONSTRUCTOR CCur (BYVAL lSrc AS LONG)
CONSTRUCTOR CCur (BYVAL ulSrc AS ULONG)
CONSTRUCTOR CCur (BYVAL fSrc AS SINGLE)
CONSTRUCTOR CCur (BYVAL dSrc AS DOUBLE)
CONSTRUCTOR CCur (BYVAL dSrc AS DECIMAL)
CONSTRUCTOR CCur (BYREF szSrc AS STRING)
CONSTRUCTOR CCur (BYVAL varSrc AS VARIANT)
CONSTRUCTOR CCur (BYVAL pDispSrc AS IDispatch PTR)
```

#### Examples

```
DIM c AS CCUR = 123
DIM c AS CCUR = CCUR(123, 45)
DIM c AS CCUR = 12345.1234
DIM c AS CCUR = "77777.999"
```
#### Remark

The constructor that accepts a DOUBLE value is particulary useful, because it allows to set the integer and fractionary parts at the same time with just a number, e.g.

```
DIM c AS CCUR = 12345.1234
```

#### Testing code:

```
'#CONSOLE ON
#INCLUDE ONCE "Afx/CCur.inc"
using Afx

DIM c AS CCUR = 12345.1234
print c
c = c + 111.11
print c
c = c - 111.11
print c
c = c * 2
print c
c = c / 2
print c
c += 123
print c
c -= 123
print c
c *= 2.3
print c
c /= 2.3
print c
c = c ^ 2
print c
c = SQR(c)
print c
DIM c2 AS CCUR = c
print c2
DIM c3 AS CCUR = c * 2
print c3
DIM c4 AS CCUR = c3 / 2
print c4
DIM c5 AS CCUR = "1234.789"
print c5
DIM c6 AS CCUR
c6 = "77777.999"
print c6
DIM c7 AS CCUR
c7 = c6
print c7
DIM cl AS CCUR = 3
cl = LOG(cl)
print cl
DIM v AS VARIANT = cl
dim cv AS CCUR = v
print cv
print "--------------"
DIM cx AS CCUR
FOR i AS LONG = 1 TO 1000000
   cx += 0.0001
NEXT
PRINT "0.0001 added 1,000,000 times = "; cx

PRINT
PRINT "Press any key..."
SLEEP
```

### Operators

| Name       | Description |
| ---------- | ----------- |
| [Operator LET](#Operator1) | Assigns a value to a **CCUR** variable. |
| [CAST operators](#Operator2) | Converts a **CCUR** into another data type. |
| [Operator \*](#Operator3) | Returns the address of the underlying **CURRENCY** structure. |
| [Comparison operators](#Operator4) | Compares currency numbers. |
| [Math operators](#Operator5) | Add, subtract, multiply or divide currency numbers. |

### Methods

| Name       | Description |
| ---------- | ----------- |
| [FormatCurrency](#FormatCurrency) | Formats a currency into a string form. |
| [FormatNumber](#FormatNumber) | Formats a currency into a string form. |
| [GetFraction](#GetFraction) | Returns the fractional component of a currency value. |
| [GetInteger](#GetInteger) | Returns the integer component of a currency value. |
| [Round](#Round) | Rounds the currency to a specified number of decimal places. |
| [SetValues](#SetValues) | Sets the integer and fractional components. |
| [ToVar](#ToVar) | Returns the currency as a VT_CY variant. |

# <a name="Operator1"></a>Operator LET (=)

Assigns a value to a `CCUR` variable.

```
OPERATOR LET (BYREF cSrc AS CCUR)
OPERATOR LET (BYVAL cySrc AS CURRENCY)
OPERATOR LET (BYVAL nInteger AS LONGLONG)
OPERATOR LET (BYVAL bSrc AS BYTE)
OPERATOR LET (BYVAL ubSrc AS UBYTE)
OPERATOR LET (BYVAL sSrc AS SHORT)
OPERATOR LET (BYVAL usSrc AS USHORT)
OPERATOR LET (BYVAL lSrc AS LONG)
OPERATOR LET (BYVAL ulSrc AS ULONG)
OPERATOR LET (BYVAL fSrc AS SINGLE)
OPERATOR LET (BYVAL dSrc AS DOUBLE)
OPERATOR LET (BYVAL dSrc AS DECIMAL)
OPERATOR LET (BYREF szSrc AS STRING)
OPERATOR LET (BYVAL varSrc AS VARIANT)
OPERATOR LET (BYVAL pDispSrc AS IDispatch PTR)
```

#### Examples

```
DIM c AS CCUR
c = 123
c = 12345.1234
c = "77777.999"
```

The operator that accepts a DOUBLE value is particulary useful, because it allows to set the integer and fractionary parts at the same time with just a number, e.g.

```
DIM c AS CCUR = 12345.1234
```

# <a name="Operator2"></a>CAST Operators

Converts a `CCUR` into another data type.

```
OPERATOR CAST () AS DOUBLE
OPERATOR CAST () AS STRING
```

These operators aren't called directly, They perform the conversion when the target is not a CCUR variable. For example, when assigning a `CCUR` variable to an standard numeric variable, the CAST() AS DOUBLE operator is implicity called.

```
OPERATOR CAST () AS CURRENCY
OPERATOR CAST () AS VARIANT
```

The casts to CURRENCY and VARIANT allow to assign a currency value to a `CVAR`:

```
DIM c AS CCUR = 12345.12
DIM cv AS CVAR
cv = cast(CY, c)
```
--or--
```
cv = cast(VARIANT, c)
```

# <a name="Operator3"></a>Operator *

Returns the address of the underlying `CURRENCY` structure.

```
OPERATOR * (BYREF cur AS CCUR) AS CURRENCY PTR
```

# <a name="Operator4"></a>Comparison operators

```
OPERATOR = (BYREF cur1 AS CCUR, BYREF cur2 AS CCUR) AS BOOLEAN
OPERATOR <> (BYREF cur1 AS CCUR, BYREF cur2 AS CCUR) AS BOOLEAN
OPERATOR < (BYREF cur1 AS CCUR, BYREF cur2 AS CCUR) AS BOOLEAN
OPERATOR > (BYREF cur1 AS CCUR, BYREF cur2 AS CCUR) AS BOOLEAN
OPERATOR <= (BYREF cur1 AS CCUR, BYREF cur2 AS CCUR) AS BOOLEAN
OPERATOR >= (BYREF cur1 AS CCUR, BYREF cur2 AS CCUR) AS BOOLEAN
```

#### Examples

```
DIM c AS CCUR = 1.23
IF c = 1.23 THEN PRINT "equal" ELSE PRINT "different"
DIM c2 AS CCUR = 1.23
IF c = c2 THEN PRINT "equal" ELSE PRINT "different"
DIM c3 AS SINGLE = 1.23
IF c = c3 THEN PRINT "equal" ELSE PRINT "different"
```

# <a name="Operator5"></a>Math operators

```
OPERATOR + (BYREF cur1 AS CCUR, BYREF cur2 AS CCUR) AS CCUR
OPERATOR + (BYREF cur AS CCUR, BYVAL nValue AS LONG) AS CCUR
OPERATOR + (BYVAL nValue AS LONG, BYREF cur AS CCUR) AS CCUR
OPERATOR + (BYREF cur AS CCUR, BYVAL nValue AS DOUBLE) AS CCUR
OPERATOR + (BYVAL nValue AS DOUBLE, BYREF cur AS CCUR) AS CCUR
OPERATOR += (BYREF cur AS CCUR)
OPERATOR += (BYVAL nValue AS LONG)
OPERATOR += (BYVAL nValue AS DOUBLE)
OPERATOR - (BYREF cur1 AS CCUR, BYREF cur2 AS CCUR) AS CCUR
OPERATOR - (BYREF cur AS CCUR, BYVAL nValue AS LONG) AS CCUR
OPERATOR - (BYVAL nValue AS LONG, BYREF cur AS CCUR) AS CCUR
OPERATOR - (BYREF cur AS CCUR, BYVAL nValue AS DOUBLE) AS CCUR
OPERATOR - (BYVAL nValue AS DOUBLE, BYREF cur AS CCUR) AS CCUR
OPERATOR -= (BYREF cur AS CCUR)
OPERATOR -= (BYREF nValue AS LONG)
OPERATOR -= (BYREF nValue AS DOUBLE)
OPERATOR * (BYREF cur1 AS CCUR, BYREF cur2 AS CCUR) AS CCUR
OPERATOR * (BYREF cur AS CCUR, BYVAL nOperand AS LONG) AS CCUR
OPERATOR * (BYVAL nOperand AS LONG, BYREF cur AS CCUR) AS CCUR
OPERATOR * (BYREF cur AS CCUR, BYVAL nOperand AS DOUBLE) AS CCUR
OPERATOR * (BYVAL nOperand AS DOUBLE, BYREF cur AS CCUR) AS CCUR
OPERATOR *= (BYREF cur AS CCUR)
OPERATOR *= (BYVAL nOperand AS LONG)
OPERATOR *= (BYVAL nOperand AS DOUBLE)
OPERATOR / (BYREF cur AS CCUR, BYVAL cOperand AS CCUR) AS CCUR
OPERATOR / (BYREF cur AS CCUR, BYVAL nOperand AS LONG) AS CCUR
OPERATOR / (BYREF cur AS CCUR, BYVAL nOperand AS DOUBLE) AS CCUR
OPERATOR / (BYVAL nValue AS LONG, BYREF cur AS CCUR) AS CCUR
OPERATOR / (BYVAL nValue AS DOUBLE, BYREF cur AS CCUR) AS CCUR
OPERATOR /= (BYREF cOperand AS CCUR)
OPERATOR /= (BYVAL nOperand AS LONG)
OPERATOR /= (BYVAL nOperand AS DOUBLE)
```

#### Examples

```
DIM c AS CCUR = 12345.1234
c = c + 111.11
c = c - 111.11
c = c * 2
c = c / 2
c += 123
c -= 123
c *= 2.3
c /= 2.3
```

#### Remarks

The version OPERATOR - (BYREF cur AS CCUR) AS CCUR changes the sign of a currency number.

```
DIM c AS CCUR = 123
c = -c
```

Or you can use the `NOT` operator:

```
DIM c AS CCUR = 123
c = NOT c
```

Other FreeBasic operators such `AND`, `MOD`, `OR`, `SHL` and `SHR` can also be used with `CCUR` variables, e.g. c = c AND 1, c = c MOD 5, etc.

# <a name="FormatCurrency"></a>FormatCurrency

Formats a currency into a string form.

```
FUNCTION FormatCurrency (BYVAL iNumDig AS LONG = -1, BYVAL ilncLead AS LONG = -2, _
   BYVAL iUseParens AS LONG = -2, BYVAL iGroup AS LONG = -2, BYVAL dwFlags AS DWORD = 0) AS STRING
```
```
FUNCTION FormatCurrencyW (BYVAL iNumDig AS LONG = -1, BYVAL ilncLead AS LONG = -2, _
   BYVAL iUseParens AS LONG = -2, BYVAL iGroup AS LONG = -2, BYVAL dwFlags AS DWORD = 0) AS CBSTR
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
FUNCTION FormatCurrency (BYVAL iNumDig AS LONG = -1, BYVAL ilncLead AS LONG = -2, _
   BYVAL iUseParens AS LONG = -2, BYVAL iGroup AS LONG = -2, BYVAL dwFlags AS DWORD = 0) AS STRING
```
```
FUNCTION FormatCurrencyW (BYVAL iNumDig AS LONG = -1, BYVAL ilncLead AS LONG = -2, _
   BYVAL iUseParens AS LONG = -2, BYVAL iGroup AS LONG = -2, BYVAL dwFlags AS DWORD = 0) AS CBSTR
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

# <a name="GetFraction"></a>GetFraction

Returns the fractional component of a currency value.

```
FUNCTION GetFraction () AS SHORT
```

# <a name="GetInteger"></a>GetInteger

Returns the integer component of a currency value.

```
FUNCTION GetInteger () AS LONGLONG
```

# <a name="Round"></a>Round

Rounds the currency to a specified number of decimal places.

```
FUNCTION Round (BYVAL nDecimals AS LONG) AS CCUR
```

# <a name="SetValues"></a>SetValues

Sets the integer and fractional components.

```
FUNCTION SetValues (BYVAL nInteger AS LONGLONG, BYVAL nFraction AS SHORT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *nInteger* | The value to be assigned to the integer component of the m_currency data member. The sign of the integer component must match the sign of the existing fractional component. *nInteger* must be in the range CY_MIN_INTEGER to CY_MAX_INTEGER inclusive. |
| *nFraction* | The value to be assigned to the fractional component of the m_currency data member.The sign of the fractional component must the same as the integer component, and the value must be in range -9999 (CY_MIN_FRACTION) to +9999 (CY_MAX_FRACTION). |

#### Remarks

Based on 4 digits. To set .2, pass 2000, to set .0002, pass a 2.

# <a name="ToVar"></a>ToVar

Returns the currency as a VT_CY variant.

```
FUNCTION ToVar () AS VARIANT
```

#### Remarks

Can be used to assign a currency directly to a `CVAR` variable.

```
DIM c AS CCUR = 12345.1234
DIM cv AS CVAR = c.ToVar
```
