# CComplex Class

`CComplex` is a class to work with complex numbers with Free Basic. Complex numbers are represented using the type `_complex`. The real and imaginary part are stored in the members `x` and `y`. There is also a flat api version that you can use alone or in combination with this class (see: [Complex Numbers Procedures](https://github.com/JoseRoca/WinFBX/blob/master/docs/Numeric%20datatypes/Complex%20Numbers%20Procedures.md)

```
TYPE _complex
   x AS DOUBLE   ' real part
   y AS DOUBLE   ' imaginary part
END TYPE
```

**Include file**: CComplex.inc.

### Constructors

Create a new instance of the **CComplex** class and assigns the values passed to it.

```
CONSTRUCTOR CComplex
CONSTRUCTOR CComplex (BYVAL x AS DOUBLE = 0, BYVAL y AS DOUBLE = 0)
CONSTRUCTOR CComplex (BYREF cpx AS CComplex)
CONSTRUCTOR CComplex (BYREF cpx AS _complex)
```

| Parameter  | Description |
| ---------- | ----------- |
| *x, y* | Uses the cartesian components (x,y) to set the real and imaginary parts of the complex number. |
| *cpx* | An instance of the **CComplex** class or a **\_complex** structure. |

#### Examples

Uses the cartesian components (x,y) to set the real and imaginary parts of the complex number.

```
DIM cpx AS CComplex = TYPE(3, 4)
DIM cpx2 AS CComplex = cpx
```

An instance of the **CComplex** class.

```
DIM cpx AS CComplex = TYPE(3, 4)
DIM cpx2 AS CComplex = cpx
```

A **\_complex** structure.

```
DIM cpx AS CComplex = TYPE<_complex>(3, 4)
```

### Operators

| Name       | Description |
| ---------- | ----------- |
| [Operator LET](#Operator1) | Assigns a value to a **CComplex** variable. |
| [CAST operators](#Operator2) | Converts a **CComplex** into another data type. |
| [Comparison operators](#Operator3) | Compares complex numbers. |
| [Math operators](#Operator4) | Add, subtract, multiply or divide complex numbers. |

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [CAbs](#CAbs) | Returns the magnitude of this complex number. |
| [CAbs2](#CAbs2) | Returns the squared magnitude of this complex number, otherwise known as the complex norm. |
| [CAbsSqr](#CAbsSqr) | Returns the absolute square (squared norm) of a complex number. |
| [CACos](#CArcCos) | Returns the complex arccosine of this complex number. |
| [CACosH](#CArcCosH) | Returns the complex hyperbolic arccosine of this complex number. |
| [CACosHReal](#CArcCosHReal) | Returns the complex arccosine of this complex number. |
| [CACosReal](#CArcCosReal) | Returns the complex arccosine of a real number. |
| [CACot](#CArcCot) | Returns the complex arccotangent of this complex number. |
| [CACotH](#CArcCotH) | Returns the complex hyperbolic arccotangent of this complex number. |
| [CACsc](#CArcCsc) | Returns the complex arccosecant of this complex number. |
| [CACscH](#CArcCscH) | Returns the complex hyperbolic arccosecant of this complex number. |
| [CACscReal](#CArcCscReal) | Returns the complex arccosecant of a real number. |
| [CAdd](#CAdd) | Adds a complex number. |
| [CAddImag](#CAddImag) | Adds an imaginary number. |
| [CAddReal](#CAddReal) | Adds a real number. |
| [CArcCos](#CArcCos) | Returns the complex arccosine of this complex number. |
| [CArcCosH](#CArcCosH) | Returns the complex hyperbolic arccosine of this complex number. |
| [CArcCosHReal](#CArcCosHReal) | Returns the complex arccosine of this complex number. |
| [CArcCosReal](#CArcCosReal) | Returns the complex arccosine of a real number. |
| [CArcCot](#CArcCot) | Returns the complex arccotangent of this complex number. |
| [CArcCotH](#CArcCotH) | Returns the complex hyperbolic arccotangent of this complex number. |
| [CArcCsc](#CArcCsc) | Returns the complex arccosecant of this complex number. |
| [CArcCscH](#CArcCscH) | Returns the complex hyperbolic arccosecant of this complex number. |
| [CArcCscReal](#CArcCscReal) | Returns the complex arccosecant of a real number. |
| [CArcSec](#CArcSec) | Returns the complex arcsecant of this complex number. |
| [CArcSecH](#CArcSecH) | Returns the complex hyperbolic arcsecant of this complex number. |
| [CArcSecReal](#CArcSecReal) | Returns the complex arcsecant of a real number. |
| [CArcSin](#CArcSin) | Returns the complex arcsine of this complex number. |
| [CArcSinH](#CArcSinH) | Returns the complex hyperbolic arcsine of this complex number. |
| [CArcSinReal](#CArcSinReal) | Returns the complex arcsine of a real number. |
| [CArcTan](#CArcTan) | Returns the complex arctangent of this complex number. |
| [CArcTanH](#CArcTanH) | Returns the complex hyperbolic arctangent of this complex number. |
| [CArcTanHReal](#CArcTanHReal) | Returns the complex hyperbolic arctangent of a real number. |
| [CArg](#CArg) | Returns the argument of this complex number. |
| [CArgument](#CArg) | Returns the argument of this complex number. |
| [CASec](#CArcSec) | Returns the complex arcsecant of this complex number. |
| [CASecH](#CArcSecH) | Returns the complex hyperbolic arcsecant of this complex number. |
| [CASecReal](#CArcSecReal) | Returns the complex arcsecant of a real number. |
| [CASin](#CArcSin) | Returns the complex arcsine of this complex number. |
| [CASinH](#CArcSinH) | Returns the complex hyperbolic arcsine of this complex number. |
| [CASinReal](#CArcSinReal) | Returns the complex arcsine of a real number. |
| [CATan](#CArcTan) | Returns the complex arctangent of this complex number. |
| [CATanH](#CArcTanH) | Returns the complex hyperbolic arctangent of this complex number. |
| [CATanHReal](#CArcTanHReal) | Returns the complex hyperbolic arctangent of a real number. |
| [CConj](#CConjugate) | Returns the complex conjugate of this complex number. |
| [CConjugate](#CConjugate) | Returns the complex conjugate of this complex number. |
| [CCos](#CCos) | Returns the complex cosine of this complex number. |
| [CCosH](#CCosH) | Returns the complex hyperbolic cosine of this complex number. |
| [CCot](#CCot) | Returns the complex cotangent of this complex number. |
| [CCotH](#CCotH) | Returns the complex hyperbolic cotangent of this complex number. |
| [CCsc](#CCsc) | Returns the complex cosecant of this complex number. |
| [CCscH](#CCscH) | Returns the complex hyperbolic cosecant of this complex number. |
| [CDiv](#CDiv) | Divides by a complex number. |
| [CDivImag](#CDivImag) | Divides by an imaginary number. |
| [CDivReal](#CDivReal) | Divides by a real number. |
| [CExp](#CExp) | Returns the complex exponential of this complex number. |
| [CImag](#CImag) | Gets/sets the imaginary part of a complex number. |
| [CInverse](#CReciprocal) | Returns the inverse, or reciprocal, of a complex number. |
| [CLog](#CLog) | Returns the complex natural logarithm (base e) of this complex number. |
| [CLog10](#CLog10) | Returns the complex base-10 logarithm of this complex number. |
| [CLogAbs](#CLogAbs) | Returns the natural logarithm of the magnitude of a complex number. |
| [CMagnitude](#CAbs) | Returns the magnitude of this complex number. |
| [CMod](#CModulus) | Returns the modulus of a complex number. |
| [CModulus](#CModulus) | Returns the modulus of a complex number. |
| [CMul](#CMul) | Multiplies by a complex number. |
| [CMulImag](#CMulImag) | Multiplies by an imaginary number. |
| [CMulReal](#CMulReal) | Multiplies by a real number. |
| [CNeg](#CNegate) | Negates the complex number. |
| [CNegate](#CNegate) | Negates the complex number. |
| [CNegative](#CNegate) | Negates the complex number. |
| [CNorm](#CAbs2) | Returns the squared magnitude of this complex number, otherwise known as the complex norm. |
| [CNthRoot](#CNthRoot) | Returns the kth nth root of a complex number where k = 0, 1, 2, 3,...,n - 1. |
| [CPhase](#CArg) | Returns the argument of this complex number. |
| [CPolar](#CPolar) | Sets the complex number from the polar representation. |
| [CPow](#CPow) | Returns this complex number raised to a complex power or to a real number. |
| [CReal](#CReal) | Gets/sets the real part of a complex number. |
| [CRect](#CSet) | Uses the cartesian components (x,y) to set the real and imaginary parts of the complex number. |
| [CReciprocal](#CReciprocal) | Returns the inverse, or reciprocal, of a complex number. |
| [CSec](#CSec) | Returns the complex secant of this complex number. |
| [CSecH](#CSecH) | Returns the complex hyperbolic secant of this complex number. |
| [CSet](#CSet) | Uses the cartesian components (x,y) to set the real and imaginary parts of the complex number. |
| [CSgn](#CSgn) | Returns the sign of this complex number. |
| [CSin](#CSin) | Returns the complex sine of this complex number. |
| [CSinH](#CSinH) | Returns the complex hyperbolic sine of this complex number. |
| [CSqr](#CSqr) | Returns the square root of a complex number. |
| [CSqrt](#CSqr) | Returns the square root of a complex number. |
| [CSub](#CSub) | Subtracts a complex number. |
| [CSubImag](#CSubImag) | Subtracts an imaginary number. |
| [CSubReal](#CSubReal) | Subtracts a real number. |
| [CSwap](#CSwap) | Exchanges the contents of two complex numbers. |
| [CTan](#CTan) | Returns the complex tangent of this complex number. |
| [CTanH](#CTanH) | Returns the complex hyperbolic tangent of this complex number. |

### Helper Methods

| Name       | Description |
| ---------- | ----------- |
| [ArcCosH](#ArcCosH) | Calculates the inverse hyperbolic cosine. |
| [ArcTanH](#ArcTanH) | Returns the inverse hyperbolic tangent of a number. |
| [IsInf](#IsInfinity) | Determines whether the argument is an infinity. |
| [IsInfinity](#IsInfinity) | Determines whether the argument is an infinity. |

# <a name="Operator1"></a>Operator LET (=)

Assigns a value to a CComplex variable.

```
OPERATOR LET (BYREF z AS CComplex)
OPERATOR LET (BYREF z AS _complex)
```

| Parameter  | Description |
| ---------- | ----------- |
| *z* | An instance of the **CComplex** class or a **\_complex** structure. |

An instance of the **CComplex** class.

```
DIM cpx AS CComplex = TYPE(3, 4)
DIM cpx2 AS CComplex = cpx
```
```
DIM cpx AS CComplex
cpx = CComplex(3, 4)
```

A **\_complex** structure.

```
DIM cpx AS CComplex = TYPE<_complex>(3, 4)
```

# <a name="Operator2"></a>Operator CAST

Returns the underlying **\_complex** number.

```
OPERATOR CAST () AS _complex
OPERATOR CAST () AS STRING
```

#### Remarks

This overloaded operator is not called directly.

The second overloaded operator returns the complex number as a formatted string that can be used with the PRINT statement.

# <a name="Operator3"></a>Comparison operators

Compares complex numbers for equality or inequality.

```
OPERATOR = (BYREF z1 AS CComplex, BYREF z2 AS CComplex) AS BOOLEAN
OPERATOR <> (BYREF z1 AS CComplex, BYREF z2 AS CComplex) AS BOOLEAN
```

#### Exammple

```
DIM cpx1 AS CComplex = CComplex(1, 2)
DIM cpx2 AS CComplex = CComplex(3, 4)
IF cpx1 = cpx2 THEN PRINT "equal" ELSE PRINT "different"
```

# <a name="Operator4"></a>Math operators

```
OPERATOR + (BYREF z1 AS CComplex, BYREF z2 AS CComplex) AS CComplex
OPERATOR + (BYVAL a AS DOUBLE, BYREF z AS CComplex) AS CComplex
OPERATOR + (BYREF z AS CComplex, BYVAL a AS DOUBLE) AS CComplex
OPERATOR - (BYREF z1 AS CComplex, BYREF z2 AS CComplex) AS CComplex
OPERATOR - (BYVAL a AS DOUBLE, BYREF z AS CComplex) AS CComplex
OPERATOR - (BYREF z AS CComplex, BYVAL a AS DOUBLE) AS CComplex
OPERATOR * (BYREF z1 AS CComplex, BYREF z2 AS CComplex) AS CComplex
OPERATOR * (BYVAL a AS DOUBLE, BYREF z AS CComplex) AS CComplex
OPERATOR * (BYREF z AS CComplex, BYVAL a AS DOUBLE) AS CComplex
OPERATOR / (BYREF leftside AS CComplex, BYREF rightside AS CComplex) AS CComplex
OPERATOR / (BYVAL a AS DOUBLE, BYREF z AS CComplex) AS CComplex
OPERATOR / (BYREF z AS CComplex, BYVAL a AS DOUBLE) AS CComplex
OPERATOR += (BYREF z AS CComplex)
OPERATOR -= (BYREF z AS CComplex)
OPERATOR *= (BYREF z AS CComplex)
OPERATOR /= (BYREF z AS CComplex)
OPERATOR - (BYREF z AS CComplex, BYVAL a AS DOUBLE) AS CComplex
OPERATOR ^ (BYREF value AS CComplex, BYREF power AS CComplex) AS CComplex
OPERATOR ^ (BYREF value AS CComplex, BYVAL power AS DOUBLE) AS CComplex
```

#### Examples

```
DIM cpx1 AS CComplex = CComplex(3, 4)
DIM cpx2 AS CComplex(5, 6)
cpx2 = cpx1 + cpx2
```
```
DIM cpx1 AS CComplex = CComplex(3, 4)
DIM cpx2 AS CComplex = CComplex(5, 6)
cpx2 = cpx1 + 11
```
```
DIM cpx1 AS CComplex = CComplex(3, 4)
DIM cpx2 AS CComplex = CComplex(5, 6)
cpx2 = 11 + cpx1
```
```
DIM cpx1 AS CComplex = CComplex(3, 4)
DIM cpx2 AS CComplex = CComplex(5, 6)
cpx2 = cpx1 - cpx2
```
```
DIM cpx1 AS CComplex = CComplex(3, 4)
DIM cpx2 AS CComplex
cpx2 = cpx1 - 11
```
```
DIM cpx1 AS CComplex = CComplex(3, 4)
DIM cpx2 AS CComplex
cpx2 = 11 - cpx1
```
```
DIM cpx1 AS CComplex = CComplex(3, 4)
DIM cpx2 AS CComplex = CComplex(5, 6)
cpx2 = cpx1 * cpx2
```
```
DIM cpx1 AS CComplex = CComplex(3, 4)
DIM cpx2 AS CComplex
cpx2 = cpx1 * 11
```
```
DIM cpx1 AS CComplex = CComplex(3, 4)
DIM cpx2 AS CComplex
cpx2 = 11 * cpx1
```
```
DIM cpx1 AS CComplex = CComplex(5, 6)
DIM cpx2 AS CComplex = CComplex(2, 3)
cpx2 = cpx1 / cpx2
```
```
DIM cpx1 AS CComplex = CComplex(5, 6)
DIM cpx2 AS CComplex
cpx2 = cpx1 / 11
```
```
DIM cpx1 AS CComplex = CComplex(5, 6)
DIM cpx2 AS CComplex
cpx2 = 11 / cpx1
```
```
DIM cpx1 AS CComplex = CComplex(3, 4)
DIM cpx2 AS CComplex = CComplex(5, 6)
cpx1 += cpx2
```
```
DIM cpx1 AS CComplex = CComplex(3, 4)
DIM cpx2 AS CComplex = CComplex(5, 6)
cpx2 -= cpx1
```
```
DIM cpx1 AS CComplex = CComplex(3, 4)
DIM cpx2 AS CComplex = CComplex(5, 6)
cpx1 *= cpx2
```
```
DIM cpx1 AS CComplex = CComplex(3, 4)
DIM cpx2 AS CComplex = CComplex(5, 6)
cpx2 /= cpx1
```
```
DIM cpx1 AS CComplex = CComplex(3, 4)
DIM cpx2 AS CComplex = -cpx1
```

# <a name="CAbs"></a>CAbs / CMagnitude

Returns the magnitude of this complex number.

```
FUNCTION CAbs () AS DOUBLE
FUNCTION CMagnitude () AS DOUBLE
```

#### Example

```
DIM cpx AS CComplex = CComplex(2, 3)
PRINT cpx.CAbs
Output: 3.60555127546399
```

# <a name="CAbs2"></a>CAbs2 / CNorm

Returns the squared magnitude of this complex number, otherwise known as the complex norm.

```
FUNCTION CAbs2 () AS DOUBLE
FUNCTION CNorm () AS DOUBLE
```

#### Example

```
DIM cpx AS CComplex = CComplex(2, 3)
PRINT cpx.CAbs2
Output: 13
```
# <a name="CAbsSqr"></a>CAbsSqr

Returns the absolute square (squared norm) of a complex number.

```
FUNCTION CAbsSqr () AS DOUBLE
```

#### Example

```
DIM cpx AS CComplex = CComplex(1.2345, -2.3456)
print cpx.CAbsSqr
Output: 7.025829610000001
```

# <a name="CAdd"></a>CAdd

Adds a complex number.

```
FUNCTION CAdd (BYREF z AS CComplex) AS CComplex
FUNCTION CAdd (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE) AS CComplex
```

| Parameter  | Description |
| ---------- | ----------- |
| *z* | The complex number to add. |
| *x, y* | Double values representing the real and imaginary parts. |

#### Examples

```
DIM cpx AS CComplex = CComplex(5, 6)
DIM cpx2 AS CComplex = CComplex(2, 3)
cpx = cpx.CAdd(cpx2)
' --or-- cpx = cpx.CAdd(CComplex(2, 3))
```
```
DIM cpx AS CComplex = CComplex(5, 6) : cpx = cpx.CAdd(2, 3)
```

# <a name="CAddImag"></a>CAddImag

Adds an imaginary number.

```
FUNCTION CAddImag (BYVAL x AS DOUBLE) AS CComplex
```

| Parameter  | Description |
| ---------- | ----------- |
| *y* | A double value. |

#### Example

```
DIM cpx AS CComplex = CComplex(5, 6)
cpx = cpx.CAddImag(10)
```

# <a name="CAddReal"></a>CAddReal

Adds a real number.

```
FUNCTION CAddReal (BYVAL x AS DOUBLE) AS CComplex
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | A double value. |

#### Example

```
DIM cpx AS CComplex = CComplex(5, 6)
cpx = cpx.CAddReal(10)
```

# <a name="CArcCos"></a>CArcCos / CACos

Returns the complex arccosine of this complex number.

```
FUNCTION CArcCos () AS CComplex
FUNCTION CACos () AS CComplex
```

#### Example

```
DIM cpx AS CComplex = CComplex(1, 1)
print cpx.CArcCos(z)
Output: 0.9045568943023814 -1.061275061905036 * i
```

# <a name="CArcCosH"></a>CArcCosH / CACosH

Returns the complex hyperbolic arccosine of this complex number. The branch cut is on the real axis, less than 1.

```
FUNCTION CArcCosH () AS CComplex
FUNCTION CACosH () AS CComplex
```

#### Example

```
DIM cpx AS CComplex = CComplex(1, 1)
print cpx.CArcCosH
Output: 1.061275061905036 +0.9045568943023813 * i
```

# <a name="CArcCosHReal"></a>CArcCosHReal / CACosHReal

Returns the complex hyperbolic arccosine of a real number.

```
FUNCTION CArcCosHReal (BYVAL value AS DOUBLE) AS CComplex
FUNCTION CACosHReal (BYVAL value AS DOUBLE) AS CComplex
```

| Parameter  | Description |
| ---------- | ----------- |
| *value* | A double value representing the real part of a complex number. |

# <a name="CArcCosReal"></a>CArcCosReal / CACosReal

Returns the complex arccosine of a real number.<br>
For a between -1 and 1, the function returns a real value in the range \[0, pi].<br>
For a less than -1 the result has a real part of pi and a negative imaginary part.<br>
For a greater than 1 the result is purely imaginary and positive.

```
FUNCTION CArcCosReal (BYVAL value AS DOUBLE) AS CComplex
FUNCTION CACosReal (BYVAL value AS DOUBLE) AS CComplex
```

| Parameter  | Description |
| ---------- | ----------- |
| *value* | A double value representing the real part of a complex number. |

#### Example

```
DIM cpx AS CComplex
print cpx.CArcCosReal(1) ' = 0 0 * i
print cpx.CArcCosReal(-1) ' = 3.141592653589793 0 * i
print cpx.CArcCosReal(2) ' = 0 +1.316957896924817 * i
```

# <a name="CArcCot"></a>CArcCot / CACot

Returns the complex arccotangent of this complex number.

```
FUNCTION CArcCot () AS CComplex
FUNCTION CACot () AS CComplex
```

#### Example

```
DIM cpx AS CComplex = CComplex(1, 1)
print cpx.CArcCot
Output: 0.5535743588970452 -0.4023594781085251 * i
```

# <a name="CArcCotH"></a>CArcCotH / CACotH

Returns the complex hyperbolic arccotangent of this complex number.

```
FUNCTION CArcCotH () AS CComplex
FUNCTION CACotH () AS CComplex
```

#### Example

```
DIM cpx AS CComplex = CComplex(1, 1)
PRINT cpx.CArcCotH
Output: 0.4023594781085251 -0.5535743588970452 * i
```

# <a name="CArcCsc"></a>CArcCsc / CACsc

Returns the complex arccosecant of this complex number.

```
FUNCTION CArcCsc () AS CComplex
FUNCTION CACsc () AS CComplex
```

#### Example

```
DIM cpx AS CComplex = CComplex(1, 1)
print cpx.CArcCsc
Output: 0.4522784471511907 -0.5306375309525178 * i
```

# <a name="CArcCscH"></a>CArcCscH / CACscH

Returns the complex hyperbolic arccosecant of this complex number.

```
FUNCTION CArcCscH () AS CComplex
FUNCTION CACscH () AS CComplex
```

#### Example

```
DIM cpx AS CComplex = CComplex(1, 1)
PRINT cpx.CArcCscH
Output: 0.5306375309525179 -0.4522784471511906 * i
```

# <a name="CArcCscReal"></a>CArcCscReal / CACscReal

Returns the complex arccosecant of a real number. 

```
FUNCTION CArcCscReal (BYVAL value AS DOUBLE) AS CComplex
FUNCTION CACscReal (BYVAL value AS DOUBLE) AS CComplex
```

#### Example

```
DIM cpx AS CComplex
print cpx.CArcCscReal(1)
Output: 1.570796326794897 0 * i
```

# <a name="CArcSec"></a>CArcSec / CASec

Returns the complex arcsecant of this complex number.

```
FUNCTION CArcSec () AS CComplex
FUNCTION CASec () AS CComplex
```

#### Example

```
DIM cpx AS CComplex = CComplex(1, 1)
print cpx.CArcSec
Output: 1.118517879643706 +0.5306375309525176 * i
```

# <a name="CArcSecH"></a>CArcSecH / CASecH

Returns the complex hyperbolic arcsecant of this complex number.

```
FUNCTION CArcSecH () AS CComplex
FUNCTION CASecH () AS CComplex
```

#### Example

```
DIM cpx AS CComplex = CComplex(1, 1)
PRINT cpx.CArcSecH
Output: 0.5306375309525178 -1.118517879643706 * i
```

# <a name="CArcSecReal"></a>CArcSecReal / CASecReal

Returns the complex arcsecant of a real number.

```
FUNCTION CArcSecReal (BYVAL value AS DOUBLE) AS CComplex
FUNCTION CASecReal (BYVAL value AS DOUBLE) AS CComplex
```

| Parameter  | Description |
| ---------- | ----------- |
| *value* | A double value representing the real part of a complex number. |

#### Example

```
DIM cpx AS CComplex
print cpx.CArcSecReal(1.1)
Output: 0.4296996661514246 0 * i
```

# <a name="CArcSin"></a>CArcSin / CASin

Returns the complex arcsine of this complex number. The branch cuts are on the real axis, less than -1 and greater than 1.

```
FUNCTION CArcSin () AS CComplex
FUNCTION CASin () AS CComplex
```

#### Example

```
DIM cpx AS CComplex
PRINT cpx.CArcSin(1, 1)
Output: 0.6662394324925152 +1.061275061905036 * i
```

# <a name="CArcSinH"></a>CArcSinH / CASinH

Returns the complex hyperbolic arcsine of this complex number. The branch cuts are on the imaginary axis, below -i and above i.

```
FUNCTION CArcSinH () AS CComplex
FUNCTION CASinH () AS CComplex
```

#### Example

```
DIM cpx AS CComplex = CComplex(1, 1)
PRINT cpx.CArcSinH
Output: 1.061275061905036 +0.6662394324925153 * i
```

# <a name="CArcSinReal"></a>CArcSinReal / CASinReal

Returns the complex arcsine of a real number.<br>
For a between -1 and 1, the function returns a real value in the range \[-pi/2, pi/2].<br>
For a less than -1 the result has a real part of -pi/2 and a positive imaginary part.<br>
For a greater than 1 the result has a real part of pi/2 and a negative imaginary part.

```
FUNCTION CArcSinReal (BYVAL value AS DOUBLE) AS CComplex
FUNCTION CASinReal (BYVAL value AS DOUBLE) AS CComplex
```

| Parameter  | Description |
| ---------- | ----------- |
| *value* | A double value representing the real part of a complex number. |

#### Example

```
DIM cpx AS CComplex
PRINT cpx.CArcSinReal(1)
Output: 1.570796326794897 +0 * i
```

# <a name="CArcTan"></a>CArcTan / CATan

Returns the complex arctangent of this complex number. The branch cuts are on the imaginary axis, below -i and above i.

```
FUNCTION CArcTan () AS CComplex
FUNCTION CATan () AS CComplex
```

#### Example

```
DIM cpx AS CComplex
PRINT cpx.CArcTan(1, 1)
Output: 1.017221967897851 +0.4023594781085251 * i
```

# <a name="CArcTanH"></a>CArcTanH / CATanH

Returns the complex hyperbolic arctangent of this complex number. The branch cuts are on the real axis, less than -1 and greater than 1.

```
FUNCTION CArcTanH () AS CComplex
FUNCTION CATanH () AS CComplex
```

#### Example

```
DIM cpx AS CComplex = CComplex(1, 1)
PRINT cpx.CArcTanH
Output: 0.4023594781085251 +1.017221967897851 * i
```

# <a name="CArcTanHReal"></a>CArcTanHReal / CATanHReal

Returns the complex hyperbolic arctangent of a real number.

```
FUNCTION CArcTanHReal (BYVAL value AS DOUBLE) AS CComplex
FUNCTION CATanHReal (BYVAL value AS DOUBLE) AS CComplex
```

| Parameter  | Description |
| ---------- | ----------- |
| *value* | A double value representing the real part of a complex number. |

# <a name="CArg"></a>CArg / CArgument / CPhase

Returns the argument of this complex number.

```
FUNCTION CArg () AS DOUBLE
FUNCTION CArgument () AS DOUBLE
FUNCTION CPhase () AS DOUBLE
```

#### Example

```
DIM cpx AS CComplex = CComplex(1, 0)
PRINT cpx.CArg
Output: 0.0
```
```
DIM cpx AS CComplex = CComplex(0, 1)
PRINT cpx.CArg
Output: 1.570796326794897
```
```
DIM cpx AS CComplex = CComplex(0, -1)
PRINT cpx.CArg
Output: -1.570796326794897
```
```
DIM cpx AS CComplex = CComplex(-1, 0)
PRINT cpx.CArg
Output: 3.141592653589793
```

# <a name="CConjugate"></a>CConjugate / CConj

Returns the complex conjugate of this complex number.

```
FUNCTION CConjugate () AS CComplex
FUNCTION CConj () AS CComplex
```

#### Example

```
DIM cpx AS CComplex = CComplex(5, 6)
cpx = cpx.CConjugate
```

# <a name="CCos"></a>CCos

Returns the complex cosine of this complex number.

```
FUNCTION CCos () AS CComplex
```

#### Example

```
DIM cpx AS CComplex = CComplex(1, 1)
PRINT cpx.CCos
Output: 0.8337300251311491 -0.9888977057628651 * i
```

# <a name="CCosH"></a>CCosH

Returns the complex hyperbolic cosine of this complex number.

```
FUNCTION CCosH () AS CComplex
```

#### Example

```
DIM cpx AS CComplex = CComplex(1, 1)
PRINT cpx.CCosH
Output: 0.8337300251311491 +0.9888977057628651 * i
```

# <a name="CCot"></a>CCot

Returns the complex cotangent of this complex number.

```
FUNCTION CCot () AS CComplex
```

#### Example

```
DIM cpx AS CComplex = CComplex(1, 1)
PRINT cpx.CCot
Output: 0.2176215618544027 -0.8680141428959249 * i
```

# <a name="CCotH"></a>CCotH

Returns the complex hyperbolic cotangent of this complex number.

```
FUNCTION CCotH () AS CComplex
```

#### Example

```
DIM cpx AS CComplex = CComplex(1, 1)
PRINT cpx.CCotH
Output: 0.8680141428959249 -0.2176215618544029 * i
```

# <a name="CCsc"></a>CCsc

Returns the complex cosecant of this complex number.

```
FUNCTION CCsc () AS CComplex
```

#### Example

```
DIM cpx AS CComplex = CComplex(1, 1)
PRINT cpx.CCsc
Output: 0.6215180171704285 -0.3039310016284265 * i
```

# <a name="CCscH"></a>CCscH

Returns the complex hyperbolic cosecant of this complex number.

```
FUNCTION CCscH () AS CComplex
```

#### Example

```
DIM cpx AS CComplex = CComplex(1, 1)
PRINT cpx.CCscH
Output: 0.3039310016284265 -0.6215180171704285 * i
```

# <a name="CDiv"></a>CDiv

Divides by a complex number.<br>
Divides by a real and an imaginary number.

```
FUNCTION CDiv (BYREF z AS CComplex) AS CComplex
FUNCTION CDiv (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE) AS CComplex
```

| Parameter  | Description |
| ---------- | ----------- |
| *z* | A complex number. |
| *x, y* | Double values representing the real and imaginary parts. |

#### Example

```
DIM cpx AS CComplex = CComplex(5, 6)
DIM cpx2 AS CComplex = CComplex(2, 3)
cpx = cpx.CDiv(cpx2)
' --or-- cpx = cpx.CDiv(CComplex(2, 3))
```
```
DIM cpx AS CComplex = CComplex(5, 6)
cpx = cpx.CDiv(2, 3)
```

# <a name="CDivImag"></a>CDivImag

Divides by an imaginary number.

```
FUNCTION CDivImag (BYVAL y AS DOUBLE) AS CComplex
```


| Parameter  | Description |
| ---------- | ----------- |
| *y* | A double value. |

#### Example

```
DIM cpx AS CComplex = CComplex(5, 6)
cpx = cpx.CDivImag(10)
```

# <a name="CDivReal"></a>CDivReal

Divides by a real number.

```
FUNCTION CDivReal (BYVAL x AS DOUBLE) AS CComplex
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | A double value. |

#### Example

```
DIM cpx AS CComplex = CComplex(5, 6)
cpx = cpx.CDivReal(10)
```

# <a name="CExp"></a>CExp

Returns the complex exponential of this complex number.

```
FUNCTION CExp () AS CComplex
```

#### Example

```
DIM cpx AS CComplex = CComplex(1, 1)
PRINT cpx.CExp
Output: 1.468693939915885 +2.287355287178842 * i
```

# <a name="CImag"></a>CImag

Gets/sets the imaginary part of a complex number.

```
PROPERTY CImag () AS DOUBLE
PROPERTY CImag (BYVAL x AS DOUBLE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | A double value. |

#### Example

```
DIM cpx AS CComplex : cpx.CImag = 4
DIM d AS DOUBLE = cpx.CImag
```

# <a name="CLog"></a>CLog

Returns the complex natural logarithm (base e) of this complex number. The branch cut is the negative real axis.

```
FUNCTION CLog () AS CComplex
FUNcTION CLog (BYVAL baseValue AS DOUBLE) AS CComplex
```

| Parameter  | Description |
| ---------- | ----------- |
| *baseValue* | A double value. |

#### Examples

```
DIM cpx AS CComplex = CComplex(1, 1)
PRINT cpx.CLog
Output: 0.3465735902799727 +0.7853981633974483 * i
```
```
DIM cpx AS CComplex = CComplex(0, 0)
PRINT cpx.CLog
Output: -1.#INF
```
```
DIM cpx AS CComplex = CComplex(1, 1)
PRINT cpx.CLog(10)
Output: 0.1505149978319906 +0.3410940884604603 * i
```

# <a name="CLog10"></a>CLog10

Returns the complex base-10 logarithm of this complex number.

```
FUNCTION CLog10 () AS CComplex
```

#### Example

```
DIM cpx AS CComplex = CComplex(1, 1)
PRINT cpx.CLog10
Output: 0.1505149978319906 +0.3410940884604603 * i
```

# <a name="CLogAbs"></a>CLogAbs

Returns the natural logarithm of the magnitude of the complex number z, log|z|. It allows an accurate evaluation of \log|z| when |z| is close to one. The direct evaluation of log(CAbs(z)) would lead to a loss of precision in this case.

```
FUNCTION CLogAbs () AS DOUBLE
```

#### Example

```
DIM cpx AS CComplex = CComplex(1.1, 0.1)
PRINT cpx.CLogAbs
Output: 0.09942542937258279
```
# <a name="CModulus"></a>CModulus / CMod

Returns the modulus of a complex number.

```
FUNCTION CModulus () AS DOUBLE
FUNCTION CMod () AS DOUBLE
```

#### Example

```
DIM cpx AS CComplex = CComplex(2.3, -4.5)
print cpx.CModulus
Output: 5.053711507397311
```

# <a name="CMul"></a>CMul

Multiplies by a complex number.<br>
Multiplies by a real and an imaginary number.

```
FUNCTION CMul (BYREF z AS CComplex) AS CComplex
FUNCTION CMul (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE) AS CComplex
```

| Parameter  | Description |
| ---------- | ----------- |
| *z* | A complex number. |
| *x, y* | Double values representing the real and imaginary parts. |

#### Examples

```
DIM cpx AS CComplex = CComplex(5, 6)
DIM cpx2 AS CComplex = CComplex(2, 3)
cpx = cpx.CMul(cpx2)
' --or-- cpx = cpx.CMul(CComplex(2, 3))
```
```
DIM cpx AS CComplex = CComplex(5, 6)
cpx = cpx.CMul(2, 3)
```

# <a name="CMulImag"></a>CMulImag

Multiplies by an imaginary number.

```
FUNCTION CMulImag (BYVAL y AS DOUBLE) AS CComplex
```

| Parameter  | Description |
| ---------- | ----------- |
| *y* | A double value. |

#### Example

```
DIM cpx AS CComplex = CComplex(5, 6)
cx = cpx.CMulImag(3)
```

# <a name="CMulReal"></a>CMulReal

Multiplies by a real number.

```
FUNCTION CMulReal (BYVAL x AS DOUBLE) AS CComplex
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | A double value. |

#### Example

```
DIM cpx AS CComplex = CComplex(5, 6)
cpx = cpx.CMulReal(2)
```

# <a name="CNegate"></a>CNegate / CNeg / CNegative

Negates the complex number.

```
FUNCTION CNeg (BYREF z AS _complex) AS _complex
FUNCTION CNegate (BYREF z AS _complex) AS _complex
FUNCTION CNegative (BYREF z AS _complex) AS _complex
```
# <a name="CNthRoot"></a>CNthRoot

Returns the kth nth root of a complex number where k = 0, 1, 2, 3,...,n - 1.<br>
De Moivre's formula states that for any complex number x and integer n it holds that<br>
  cos(x)+ i*sin(x))^n = cos(n*x) + i*sin(n*x)<br>
where i is the imaginary unit (i2 = -1).<br>
Since z = r*e^(i*t) = r * (cos(t) + i sin(t))<br>
  where<br>
  z = (a, ib)<br>
  r = modulus of z<br>
  t = argument of z<br>
  i = sqrt(-1.0)<br>
we can calculate the nth root of z by the formula:<br>
  z^(1/n) = r^(1/n) * (cos(x/n) + i sin(x/n))<br>
by using log division.

```
FUNCTION CNthRoot (BYVAL n AS LONG, BYVAL k AS LONG = 0) AS _complex
```

#### Example

```
DIM cpx AS CComplex = CComplex(2.3, -4.5)
DIM n AS LONG = 5
FOR i AS LONG = 0 TO n - 1
   print CStr(cpx.CNthRoot(n, i))
NEXT
Output:
 1.349457704883236  -0.3012830564679053 * i
 0.7035425781022545 +1.190308959128094 * i
-0.9146444790833151 +1.036934450322577 * i
-1.268823953798186  -0.5494482247230521 * i
 0.1304681498960107 -1.376512128259714 * i
```

# <a name="CPolar"></a>CPolar

Sets the complex number from the polar representation.

```
FUNCTION CPolar (BYVAL r AS DOUBLE, BYVAL theta AS DOUBLE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *r* | The modulus of complex number. |
| *theta* | The angle with the positive direction of x-axis. |

# <a name="CPow"></a>CPow

Returns this complex number raised to a complex power.<br>
Returns this complex number raised to a real number.

```
FUNCTION CPow (BYREF power AS CComplex) AS CComplex
FUNCTION CPow (BYVAL power AS DOUBLE) AS CComplex
```

| Parameter  | Description |
| ---------- | ----------- |
| *power* | A complex or a double number. |

#### Examples

```
DIM cpx AS CComplex = CComplex(1, 1)
DIM b AS CComplex = CComplex(2, 2)
PRINT cpx.CPow(b)
Output: -0.2656539988492412 +0.3198181138561362 * i
```
```
DIM cpx AS CComplex = CComplex(1, 1)
PRINT cpx.CPow(2)
Output: 1.224606353822378e-016 +2 * i
```

# <a name="CReal"></a>CReal

Gets/sets the real part of a complex number.

```
PROPERTY CReal () AS DOUBLE
PROPERTY CReal (BYVAL x AS DOUBLE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | A double value. |

#### Example

```
DIM cpx AS CComplex : cpx.CReal = 3
DIM d AS DOUBLE = cpx.CReal
```

# <a name="CReciprocal"></a>CReciprocal / CInverse

Returns the inverse, or reciprocal, of a complex number.

```
FUNCTION CReciprocal () AS CComplex
FUNCTION CInverse () AS CComplex
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | A double value. |

#### Example

```
DIM cpx AS CComplex = CComplex(1, 1)
Print cpx.CReciprocal
Output: 0.5 -0.5 * i
```

# <a name="CSec"></a>CSec

Returns the complex secant of this complex number.

```
FUNCTION CSec () AS CComplex
```

#### Example

```
DIM cpx AS CComplex = CComplex(1, 1)
PRINT cpx.CSec
Output: 0.4983370305551869 +0.591083841721045 * i
```

# <a name="CSecH"></a>CSecH

Returns the complex hyperbolic secant of this complex number.

```
FUNCTION CSecH () AS CComplex
```

#### Example

```
DIM cpx AS CComplex = CComplex(1, 1)
PRINT cpx.CSecH
Output: 0.4983370305551869 -0.591083841721045 * i
```

# <a name="CSet"></a>CSet / CRect

Uses the cartesian components (x,y) to set the real and imaginary parts of the complex number.

```
PROPERTY CSet (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE)
PROPERTY CRect (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *x, y* | Double values. |

#### Example

```
DIM cpx AS CComplex : cpx.CSet = 3, 4
DIM cpx AS CComplex : cpx.CRect = 3, 4
```

# <a name="CSgn"></a>CSgn

Returns the sign of this complex number.<br>
If number is greater than zero, then CSgn returns 1.<br>
If number is equal to zero, then CSgn returns 0.<br>
If number is less than zero, then CSgn returns -1.

```
PROPERTY CSgn () AS LONG
```

# <a name="CSin"></a>CSin

Returns the complex sine of this complex number.

```
PROPERTY CSin () AS CSin
```

#### Example

```
DIM cpx AS CComplex = CComplex(1, 1)
PRINT cpx.CSin
Output: 1.298457581415977 +0.6349639147847361 * i
```

# <a name="CSinH"></a>CSinH

Returns the complex hyperbolic sine of this complex number.

```
FUNCTION CSinH () AS CComplex
```

#### Example

```
DIM cpx AS CComplex = CComplex(1, 1)
PRINT cpx.CSinH
Output: 0.6349639147847361 +1.298457581415977 * i
```

# <a name="CSqr"></a>CSqr / CSqrt

Returns the square root of the complex number z. The branch cut is the negative real axis. The result always lies in the right half of the complex plane.

```
FUNCTION CSqr () AS CComplex
FUNCTION CSqrt () AS CComplex
```

Returns the complex square root of the real number x, where x may be negative.

```
FUNCTION CSqr (BYVAL value AS DOUBLE) AS CComplex
FUNCTION CSqrt (BYVAL value AS DOUBLE) AS CComplex
```

| Parameter  | Description |
| ---------- | ----------- |
| *value* | A double value representing a real number. |

#### Examples

```
DIM cpx AS CComplex = CComplex(2, 3)
PRINT cpx.CSqrt
Output: 1.67414922803554 +0.895977476129838 * i
```

Compute the square root of -1:

```
DIM cpx AS CComplex = CComplex(-1)
PRINT cpx.CSqr
Output: 0 +1.0 * i
```

# <a name="CSub"></a>CSub

Subtracts a complex number.

```
FUNCTION CSub (BYREF z AS CComplex) AS CComplex
FUNCTION CSub (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE) AS CComplex
```

| Parameter  | Description |
| ---------- | ----------- |
| *z* | The complex number to subtract. |
| *x, y* | Double values representing the real and imaginary parts. |

#### Examples

```
DIM cpx AS CComplex = CComplex(5, 6)
DIM cpx2 AS CComplex = CComplex(2, 3)
cpx = cpx.CSub(cpx2)
' --or-- cpx = cpx.CSub(CComplex(2, 3))
```
```
DIM cpx AS CComplex = CComplex(5, 6)
cpx = cpx.CSub(2, 3)
```

# <a name="CSubImag"></a>CSubImag

Subtracts an imaginary number.

```
FUNCTION CSubImag (BYVAL y AS DOUBLE) AS CComplex
```

| Parameter  | Description |
| ---------- | ----------- |
| *y* | A double value. |

#### Example

```
DIM cpx AS CComplex = CComplex(5, 6)
cx = cpx.CSubImag(3)
```

# <a name="CSubReal"></a>CSubReal

Subtracts a real number.

```
FUNCTION CSubReal (BYVAL x AS DOUBLE) AS CComplex
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | A double value. |

#### Example

```
DIM cpx AS CComplex = CComplex(5, 6)
cpx = cpx.CSubReal(2)
```

# <a name="CSwap"></a>CSwap

Exchanges the contents of two complex numbers.

```
SUB CSwap (BYREF z AS CComplex)
```

| Parameter  | Description |
| ---------- | ----------- |
| *z* | The complex number to swap. |

# <a name="CTan"></a>CTan

Returns the complex tangent of this complex number.

```
FUNCTION CTan () AS CComplex
```

#### Example

```
DIM cpx AS CComplex = CComplex(1, 1)
PRINT cpx.CTan
Output: 0.2717525853195117 +1.083923327338695 * i
```

# <a name="CTanH"></a>CTanH

Returns the complex hyperbolic tangent of this complex number.

```
FUNCTION CTanH () AS CComplex
```

#### Example

```
DIM cpx AS CComplex = CComplex(1, 1)
PRINT cpx.CTanH
Output: 1.083923327338695 +0.2717525853195119 * i
```

# <a name="ArcCosH"></a>ArcCosH

Calculates the inverse hyperbolic cosine.

```
FUNCTION ArcCosH (BYVAL x AS DOUBLE) AS DOUBLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | A double value. |

#### Example

```
DIM AS double pi = 3.1415926535
DIM AS double x, y
DIM cpx AS CComplex
x = cosh(pi / 4)
y = cpx.ArcCosH(x)
print "cosh = ", pi/4, x
print "ArcCosH = ", x, y

Output:
cosh =  0.785398163375      1.324609089232506
acosh = 1.324609089232506   0.7853981633749999
```

# <a name="ArcTanH"></a>ArcTanH

Returns the inverse hyperbolic tangent of a number.

```
FUNCTION ArcTanH (BYVAL x AS DOUBLE) AS DOUBLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | A double value. |

#### Example

```
DIM cpx AS CComplex
print cpx.ArcTanh(0.76159416)
Output: 1.00000000962972

print cpx.ArcTanH(-0.1)
Output: -0.1003353477310756
```

# <a name="IsInfinity"></a>IsInfinity / IsInf

Determines whether the argument is an infinity.

```
FUNCTION IsInfinity (BYVAL x AS DOUBLE) AS LONG
FUNCTION IsInf (BYVAL x AS DOUBLE) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | A double value. |

#### Remarks

Returns +1 if x is positive infinity, -1 if x is negative infinity and 0 otherwise.
