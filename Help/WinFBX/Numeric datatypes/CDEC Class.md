# CDEC Class

**CDEC** is a wrapper class for the DECIMAL data type. Holds signed 128-bit (16-byte) values representing 96-bit (12-byte) integer numbers scaled by a variable power of 10. The scaling factor specifies the number of digits to the right of the decimal point; it ranges from 0 through 28. With a scale of 0 (no decimal places), the largest possible value is +/-79,228,162,514,264,337,593,543,950,335 (+/-7.9228162514264337593543950335E+28). With 28 decimal places, the largest value is +/-7.9228162514264337593543950335, and the smallest nonzero value is +/-0.0000000000000000000000000001 (+/-1E-28).

In this implementation the scale is dynamic, dictated by the number of decimal places.

```
DIM dec AS CDEC = 12345.12
print dec   ' = 12345,12
print "scale: ", dec.m_dec.scale  ' --> scale = 2
dec += dec + 11.1111
print dec   ' = 24701,3511
print "new scale: ", dec.m_dec.scale  ' --> scale = 4
```

### Constructors

Creates a new instance of the **CDEC** class and assigns the values passed to it.

```
CONSTRUCTOR CDec
CONSTRUCTOR CDec (BYVAL lcid AS LCID = 0)
CONSTRUCTOR CDec (BYREF cSrc AS CDec, BYVAL lcid AS LCID = 0)
CONSTRUCTOR CDec (BYREF decSrc AS DECIMAL, BYVAL lcid AS LCID = 0)
CONSTRUCTOR CDec (BYVAL cySrc AS CURRENCY, BYVAL lcid AS LCID = 0)
CONSTRUCTOR CDec (BYREF cur AS CCUR, BYVAL lcid AS LCID = 0)
CONSTRUCTOR CDec (BYVAL bSrc AS BYTE, BYVAL lcid AS LCID = 0)
CONSTRUCTOR CDec (BYVAL ubSrc AS UBYTE, BYVAL lcid AS LCID = 0)
CONSTRUCTOR CDec (BYVAL sSrc AS SHORT, BYVAL lcid AS LCID = 0)
CONSTRUCTOR CDec (BYVAL usSrc AS USHORT, BYVAL lcid AS LCID = 0)
CONSTRUCTOR CDec (BYVAL lSrc AS LONG, BYVAL lcid AS LCID = 0)
CONSTRUCTOR CDec (BYVAL ulSrc AS ULONG, BYVAL lcid AS LCID = 0)
CONSTRUCTOR CDec (BYVAL nInteger AS LONGINT, BYVAL lcid AS LCID = 0)
CONSTRUCTOR CDec (BYVAL nUInteger AS ULONGINT, BYVAL lcid AS LCID = 0)
CONSTRUCTOR CDec (BYVAL fSrc AS SINGLE, BYVAL lcid AS LCID = 0)
CONSTRUCTOR CDec (BYVAL dSrc AS DOUBLE, BYVAL lcid AS LCID = 0)
CONSTRUCTOR CDec (BYREF wszSrc AS WSTRING, BYVAL lcid AS LCID = 0, BYVAL dwFlags AS ULONG = 0)
CONSTRUCTOR CDec (BYVAL pDispSrc AS IDispatch PTR, BYVAL lcid AS LCID = 0)
CONSTRUCTOR CDec (BYVAL varSrc AS VARIANT, BYVAL lcid AS LCID = 0)
```

#### Remarks

As the bigger numeric variable supported by FreeBasic is a long integer, if we want to set bigger values we need to use strings, e.g.

```
DIM dec AS CDEC = "-79228162514264337593543950.335"
```
--or--
```
DIM dec AS CDEC = "-79,228,162,514,264,337,593,543,950.335"
```

By default, the locale user identifier is used. Therefore, in my Spanish computer I need to use "," as the decimal separator and "." as the thousands separator.

```
DIM dec AS CDEC = "-79.228.162.514.264.337.593.543.950,335"
```

But it can be overriden by passing an LCID value (1033 for US).

```
DIM dec AS CDEC = CDEC("-79,228,162,514,264,337,593,543,950.335", 1033)
```

### Operators

| Name       | Description |
| ---------- | ----------- |
| [Operator LET](#Operator1) | Assigns a value to a **CDEC** variable. |
| [CAST operators](#Operator2) | Converts a **CDEC** into another data type. |
| [Operator \*](#Operator3) | Returns the address of the underlying **CURRENCY** structure. |
| [Comparison operators](#Operator4) | Compares decimal numbers. |
| [Math operators](#Operator5) | Add, subtract, multiply or divide decimal numbers. |

### Methods

| Name       | Description |
| ---------- | ----------- |
| [DecAbs](#DecAbs) | Returns the absolute value of a decimal data type. |
| [DecFix](#DecFix) | Returns integer portion of a decimal data type. |
| [DecInt](#DecInt) | Returns integer portion of a decimal data type. |
| [DecRound](#DecRound) | Rounds the DECIMAL to a specified number of decimal places. |
| [IsSigned](#IsSigned) | Returns true if this number is signed or false otherwise. |
| [IsUnsigned](#IsUnsigned) | Returns true if this number is unsigned or false otherwise. |
| [Scale](#Scale) | Returns the scale of the DECIMAL number. |
| [Sign](#Sign) | Returns 0 if it is not signed of &h80 (128) if it is signed. |
| [ToVar](#ToVar) | Returns the currency as a VT_CY variant. |

# <a name="Operator1"></a>Operator LET (=)

Assigns a value to a **CDEC** variable.

```
OPERATOR LET (BYREF cSrc AS CDec)
```

#### Examples

```
DIM c AS CDEC
c = 123
c = 12345.1234
```

# <a name="Operator2"></a>CAST Operators

Converts a CDEC into another data type.

```
OPERATOR CAST () AS CURRENCY
OPERATOR CAST () AS VARIANT
OPERATOR CAST () AS STRING
```

#### Remarks

These operators aren't called directly, They perform the conversion when the target is not a CDEC variable.

# <a name="Operator3"></a>Operator *

Returns the address of the underlying **DECIMAL** structure.

```
OPERATOR * (BYREF dec AS CDEC) AS DECIMAL PTR
```

# <a name="Operator4"></a>Comparison operators

```
OPERATOR = (BYREF dec1 AS CDEC, BYREF dec2 AS CDEC) AS BOOLEAN
OPERATOR <> (BYREF dec1 AS CDEC, BYREF dec2 AS CDEC) AS BOOLEAN
OPERATOR < (BYREF dec1 AS CDEC, BYREF dec2 AS CDEC) AS BOOLEAN
OPERATOR > (BYREF dec1 AS CDEC, BYREF dec2 AS CDEC) AS BOOLEAN
OPERATOR <= (BYREF dec1 AS CDEC, BYREF dec2 AS CDEC) AS BOOLEAN
OPERATOR >= (BYREF dec1 AS CDEC, BYREF dec2 AS CDEC) AS BOOLEAN
```

# <a name="Operator5"></a>Math operators

```
OPERATOR + (BYREF dec1 AS CDEC, BYREF dec2 AS CDEC) AS CDEC
OPERATOR += (BYREF dec AS CDEC)
OPERATOR - (BYREF dec1 AS CDEC, BYREF dec2 AS CDEC) AS CDEC
OPERATOR -= (BYREF dec AS CDEC)
OPERATOR * (BYREF dec1 AS CDEC, BYREF dec2 AS CDEC) AS CDEC
OPERATOR *= (BYREF dec AS CDEC)
OPERATOR / (BYREF dec AS CDEC, BYVAL cOperand AS CDEC) AS CDEC
OPERATOR /= (BYREF cOperand AS CDEC)
```

#### Examples

```
DIM dec AS CDEC = 12345.1234
dec = dec + 111.11
dec = dec - 111.11
dec = dec * 2
dec = dec / 2
dec += 123
dec -= 123
dec *= 2.3
dec /= 2.3
```

#### Remarks

The version OPERATOR - (BYREF dec AS CDEC) AS CDEC changes the sign of a decimal number.

```
DIM dec AS CDEC = 123
dec = -dec
```

Other FreeBasic operators such AND, MOD, OR, SHL and SHR can also be used with CDEC variables, e.g. c = c AND 1, c = c MOD 5, etc.

# <a name="DecAbs"></a>DecAbs

Returns the absolute value of a decimal data type.

```
FUNCTION DecAbs () AS CDEC
```

# <a name="DecFix"></a>DecFix

Returns integer portion of a decimal data type. If the value is negative, then the first negative integer greater than or equal to the value is returned.

```
FUNCTION DecFix () AS CDEC
```

# <a name="DecInt"></a>DecInt

Returns integer portion of a decimal data type. If the value is negative, then the first negative integer less than or equal to the value is returned.

```
FUNCTION DecInt () AS CDEC
```

# <a name="DecRound"></a>DecRound

Rounds the DECIMAL to a specified number of decimal places.

```
FUNCTION DecRound () AS CDEC
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

Returns the scale of the DECIMAL number.

```
FUNCTION Scale () AS UBYTE
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
DIM dec AS CDEC = 12345.1234
DIM cv AS CVAR = dec.ToVar
```
