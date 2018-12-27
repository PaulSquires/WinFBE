# Complex Numbers Procedures

Procedures to work with complex numbers. Complex numbers are represented using the type `_complex`. The real and imaginary part are stored in the members `x` and `y`.

```
TYPE _complex
   x AS DOUBLE   ' real part
   y AS DOUBLE   ' imaginary part
END TYPE
```

**Include file**: AfxComplex.inc.

### Operators

| Name       | Description |
| ---------- | ----------- |
| [Comparison operators](#Operator3) | Compares complex numbers. |
| [Math operators](#Operator4) | Add, subtract, multiply or divide complex numbers. |

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [CAbs](#CAbs) | Returns the magnitude of a complex number. |
| [CAbs2](#CAbs2) | Returns the squared magnitude of a complex number, otherwise known as the complex norm. |
| [CAbsSqr](#CAbsSqr) | Returns the absolute square (squared norm) of a complex number. |
| [CACos](#CArcCos) | Returns the complex arccosine of a complex number. |
| [CACosH](#CArcCosH) | Returns the complex hyperbolic arccosine of a complex number. The branch cut is on the real axis, less than 1. |
| [CACosHReal](#CArcCosHReal) | Returns the complex arccosine of a complex number. |
| [CACosReal](#CArcCosReal) | Returns the complex arccosine of a real number. |
| [CACot](#CArcCot) | Returns the complex arccotangent of a complex number. |
| [CACotH](#CArcCotH) | Returns the complex hyperbolic arccotangent of a complex number. |
| [CACsc](#CArcCsc) | Returns the complex arccosecant of a complex number. |
| [CACscH](#CArcCscH) | Returns the complex hyperbolic arccosecant of a complex number. |
| [CACscReal](#CArcCscReal) | Returns the complex arccosecant of a real number. |
| [CAddImag](#CAddImag) | Adds an imaginary number. |
| [CAddReal](#CAddReal) | Adds a real number. |
| [CArcCos](#CArcCos) | Returns the complex arccosine of a complex number. |
| [CArcCosH](#CArcCosH) | Returns the complex hyperbolic arccosine of a complex number. |
| [CArcCosHReal](#CArcCosHReal) | Returns the complex arccosine of a complex number. |
| [CArcCosReal](#CArcCosReal) | Returns the complex arccosine of a real number. |
| [CArcCot](#CArcCot) | Returns the complex arccotangent of a complex number. |
| [CArcCotH](#CArcCotH) | Returns the complex hyperbolic arccotangent of a complex number. |
| [CArcCsc](#CArcCsc) | Returns the complex arccosecant of a complex number. |
| [CArcCscH](#CArcCscH) | Returns the complex hyperbolic arccosecant of a complex number. |
| [CArcCscReal](#CArcCscReal) | Returns the complex arccosecant of a real number. |
| [CArcSec](#CArcSec) | Returns the complex arcsecant of a complex number. |
| [CArcSecH](#CArcSecH) | Returns the complex hyperbolic arcsecant of a complex number. |
| [CArcSecReal](#CArcSecReal) | Returns the complex arcsecant of a real number. |
| [CArcSin](#CArcSin) | Returns the complex arcsine of a complex number. |
| [CArcSinH](#CArcSinH) | Returns the complex hyperbolic arcsine of a complex number. |
| [CArcSinReal](#CArcSinReal) | Returns the complex arcsine of a real number. |
| [CArcTan](#CArcTan) | Returns the complex arctangent of a complex number. |
| [CArcTanH](#CArcTanH) | Returns the complex hyperbolic arctangent of a complex number. |
| [CArcTanHReal](#CArcTanHReal) | Returns the complex hyperbolic arctangent of a real number. |
| [CArg](#CArg) | Returns the argument of a complex number. |
| [CArgument](#CArg) | Returns the argument of a complex number. |
| [CASec](#CArcSec) | Returns the complex arcsecant of a complex number. |
| [CASecH](#CArcSecH) | Returns the complex hyperbolic arcsecant of a complex number. |
| [CASecReal](#CArcSecReal) | Returns the complex arcsecant of a real number. |
| [CASin](#CArcSin) | Returns the complex arcsine of a complex number. |
| [CASinH](#CArcSinH) | Returns the complex hyperbolic arcsine of a complex number. |
| [CASinReal](#CArcSinReal) | Returns the complex arcsine of a real number. |
| [CATan](#CArcTan) | Returns the complex arctangent of a complex number. |
| [CATanH](#CArcTanH) | Returns the complex hyperbolic arctangent of a complex number. |
| [CATanHReal](#CArcTanHReal) | Returns the complex hyperbolic arctangent of a real number. |
| [CConj](#CConjugate) | Returns the complex conjugate of a complex number. |
| [CConjugate](#CConjugate) | Returns the complex conjugate of a complex number. |
| [CCos](#CCos) | Returns the complex cosine of a complex number. |
| [CCosH](#CCosH) | Returns the complex hyperbolic cosine of a complex number. |
| [CCot](#CCot) | Returns the complex cotangent of a complex number. |
| [CCotH](#CCotH) | Returns the complex hyperbolic cotangent of a complex number. |
| [CCsc](#CCsc) | Returns the complex cosecant of thais complex number. |
| [CCscH](#CCscH) | Returns the complex hyperbolic cosecant of a complex number. |
| [CDivImag](#CDivImag) | Divides by an imaginary number. |
| [CDivReal](#CDivReal) | Divides by a real number. |
| [CExp](#CExp) | Returns the complex exponential of s complex number. |
| [CGetImag](#CGetImag) | Returns the imaginary part of a complex number. |
| [CGetReal](#CGetReal) | Returns the real part of a complex number. |
| [CInverse](#CReciprocal) | Returns the inverse, or reciprocal, of a complex number. |
| [CLog](#CLog) | Returns the complex natural logarithm (base e) of s complex number. |
| [CLog10](#CLog10) | Returns the complex base-10 logarithm of a complex number. |
| [CLogAbs](#CLogAbs) | Returns the natural logarithm of the magnitude of a complex number. |
| [CMagnitude](#CAbs) | Returns the magnitude of a complex number. |
| [CMod](#CModulus) | Returns the modulus of a complex number. |
| [CModulus](#CModulus) | Returns the modulus of a complex number. |
| [CMulImag](#CMulImag) | Multiplies by an imaginary number. |
| [CMulReal](#CMulReal) | Multiplies by a real number. |
| [CNeg](#CNegate) | Negates a complex number. |
| [CNegate](#CNegate) | Negates a complex number. |
| [CNegative](#CNegate) | Negates the complex number. |
| [CNorm](#CAbs2) | Returns the squared magnitude of a complex number, otherwise known as the complex norm. |
| [CNthRoot](#CNthRoot) | Returns the kth nth root of a complex number where k = 0, 1, 2, 3,...,n - 1. |
| [CPhase](#CArg) | Returns the argument of a complex number. |
| [CPolar](#CPolar) | Sets the complex number from the polar representation. |
| [CPow](#CPow) | Returns this complex number raised to a complex power or to a real number. |
| [CRect](#CSet) | Uses the cartesian components (x,y) to set the real and imaginary parts of the complex number. |
| [CReciprocal](#CReciprocal) | Returns the inverse, or reciprocal, of a complex number. |
| [CSec](#CSec) | Returns the complex secant of a complex number. |
| [CSecH](#CSecH) | Returns the complex hyperbolic secant of a complex number. |
| [CSet](#CSet) | Uses the cartesian components (x,y) to set the real and imaginary parts of a complex number. |
| [CSetImag](#CSetImag) | Sets the imaginary part of a complex number. |
| [CSetReal](#CSetReal) | Sets the real part of a complex number. |
| [CSgn](#CSgn) | Returns the sign of a complex number. |
| [CSin](#CSin) | Returns the complex sine of a complex number. |
| [CSinH](#CSinH) | Returns the complex hyperbolic sine of a complex number. |
| [CSqr](#CSqr) | Returns the square root of a complex. |
| [CSqrt](#CSqr) | Returns the square root of a complex number. |
| [CSqrReal](#CSqrReal) | Returns the complex square root of a real number. |
| [CSubImag](#CSubImag) | Subtracts an imaginary number. |
| [CSubReal](#CSubReal) | Subtracts a real number. |
| [CSwap](#CSwap) | Exchanges the contents of two complex numbers. |
| [CTan](#CTan) | Returns the complex tangent of a complex number. |
| [CTanH](#CTanH) | Returns the complex hyperbolic tangent of a complex number. |

### Helper Procedures

| Name       | Description |
| ---------- | ----------- |
| [acosh](#acosh) | Calculates the inverse hyperbolic cosine. |
| [asinh](#asinh) | Calculates the inverse hyperbolic sine. |
| [atanh](#atanh) | Calculates the inverse hyperbolic tangent of a number. |
| [hypot3](#hypot3) | Computes the value of sqr(x^2 + y^2 + z^2). |
| [IsInfinity](#IsInfinity) | Determines whether the argument is an infinity. |

# <a name="Operator3"></a>Comparison operators

Compares complex numbers for equality or inequality.

```
OPERATOR = (BYREF z1 AS _complex, BYREF z2 AS _complex) AS BOOLEAN
OPERATOR <> (BYREF z1 AS _complex, BYREF z2 AS _complex) AS BOOLEAN
```

#### Example

```
DIM c1 AS _complex = (1, 2)
DIM c2 AS _complex = (3, 4)
IF c1 = c2 THEN PRINT "equal" ELSE PRINT "different"
```

# <a name="Operator4"></a>Math operators

```
OPERATOR + (BYREF z1 AS _complex, BYREF z2 AS _complex) AS _complex
OPERATOR + (BYVAL a AS DOUBLE, BYREF z AS _complex) AS _complex
OPERATOR + (BYREF z AS _complex, BYVAL a AS DOUBLE) AS _complex
OPERATOR - (BYREF z1 AS _complex, BYREF z2 AS _complex) AS _complex
OPERATOR - (BYVAL a AS DOUBLE, BYREF z AS _complex) AS _complex
OPERATOR - (BYREF z AS _complex, BYVAL a AS DOUBLE) AS _complex
OPERATOR * (BYREF z1 AS _complex, BYREF z2 AS _complex) AS _complex
OPERATOR * (BYVAL a AS DOUBLE, BYREF z AS _complex) AS _complex
OPERATOR * (BYREF z AS _complex, BYVAL a AS DOUBLE) AS _complex
OPERATOR / (BYREF leftside AS _complex, BYREF rightside AS _complex) AS _complex
OPERATOR / (BYVAL a AS DOUBLE, BYREF z AS _complex) AS _complex
OPERATOR / (BYREF z AS _complex, BYVAL a AS DOUBLE) AS _complex
OPERATOR - (BYREF z AS _complex, BYVAL a AS DOUBLE) AS _complex
OPERATOR ^ (BYREF value AS _complex, BYREF power AS _complex) AS _complex
OPERATOR ^ (BYREF value AS _complex, BYVAL power AS DOUBLE) AS _complex
```

#### Examples

```
DIM c1 AS _complex = (3, 4)
DIM c2 AS _complex = (5, 6)
c2 = c1 + c2
print CStr(c2)
```
```
DIM c1 AS _complex = (3, 4)
DIM c2 AS _complex = (5, 6)
c2 = c1 + 11
print CStr(c2)
```
```
DIM c1 AS _complex = (3, 4)
DIM c2 AS _complex = (5, 6)
c2 = 11 + c1
print CStr(c2)
```
```
DIM c1 AS _complex = (3, 4)
DIM c2 AS _complex = (5, 6)
c2 = c1 - c2
print CStr(c2)
```
```
DIM c1 AS _complex = (3, 4)
DIM c2 AS _complex = c1 - 11
print CStr(c2)
```
```
DIM c1 AS _complex = (3, 4)
DIM c2 AS _complex = 11 - c1
print CStr(c2)
```
```
DIM c1 AS _complex = (3, 4)
DIM c2 AS _complex = (5, 6)
c2 = c1 * c2
print CStr(c2)
```
```
DIM c1 AS _complex = (3, 4)
DIM c2 AS _complex = c1 * 11
print CStr(c2)
```
```
DIM c1 AS _complex = (3, 4)
DIM c2 AS _complex = 11 * c1
print CStr(c2)
```
```
DIM c1 AS _complex = (3, 4)
DIM c2 AS _complex = (5, 6)
c2 = c1 / c2
print CStr(c2)
```
```
DIM c1 AS _complex = (3, 4)
DIM c2 AS _complex = c1 / 11
print CStr(c2)
```
```
DIM cpx1 AS _complex = (5, 6)
DIM cpx2 AS _complex
cpx2 = 11 / cpx1
```
```
DIM c1 AS _complex = (3, 4)
DIM c2 AS _complex = 11 * c1
print CStr(c2)
```
```
DIM c1 AS _complex = (3, 4)
DIM c2 AS _complex = (5, 6)
c2 -= c1
print CStr(c2)
```
```
DIM c1 AS _complex = (3, 4)
DIM c2 AS _complex = (5, 6)
c1 *= c2
print CStr(c2)
```
```
DIM c1 AS _complex = (3, 4)
DIM c2 AS _complex = (5, 6)
c1 /= c2
print CStr(c2)
```
```
DIM c1 AS _complex = (3, 4)
DIM c2 AS _complex = -c1
print CStr(c2)
```

# <a name="CAbs"></a>CAbs / CMagnitude

Returns the magnitude of a complex number.

```
FUNCTION CAbs (BYREF z AS _complex) AS DOUBLE
FUNCTION CMagnitude (BYREF z AS _complex) AS DOUBLE
```

#### Example

```
DIM c AS _complex = (2, 3)
PRINT CAbs(c)
Output: 3.60555127546399
```

# <a name="CAbs2"></a>CAbs2 / CNorm

Returns the squared magnitude of a complex number, otherwise known as the complex norm.

```
FUNCTION CAbs2 () AS DOUBLE
FUNCTION CNorm () AS DOUBLE
```

#### Example

```
DIM z AS _complex = (2, 3)
DIM d AS DOUBLE = CAbs2(z)
PRINT d
Output: 13
--or--
DIM z AS _complex = (2, 3)
DIM d AS DOUBLE = CNorm(z)
PRINT d
Output: 13
```

# <a name="CAbsSqr"></a>CAbsSqr

Returns the absolute square (squared norm) of a complex number.

```
FUNCTION CAbsSqr (BYREF z AS _complex) AS DOUBLE
```

#### Example

```
DIM z AS _complex = (1.2345, -2.3456)
print CAbsSqr(z)
Output: 7.025829610000001
```

# <a name="CAddImag"></a>CAddImag

Adds an imaginary number.

```
FUNCTION CAddImag (BYREF z AS _complex, BYVAL y AS DOUBLE) AS _complex
```

| Parameter  | Description |
| ---------- | ----------- |
| *y* | A double value. |

#### Example

```
DIM c AS _complex = (5, 6)
c = CAddImag(c, 10)
print CStr(c)
```

# <a name="CAddReal"></a>CAddReal

Adds a real number.

```
FUNCTION CAddReal (BYREF z AS _complex, BYVAL x AS DOUBLE) AS _complex
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | A double value. |

#### Examples

```
DIM c AS _complex = (5, 6)
c = CAddReal(c, 10)
print CStr(c)
```

# <a name="CArcCos"></a>CArcCos / CACos

Returns the complex arccosine of a complex number.

```
FUNCTION CArcCos (BYREF value AS _complex) AS _complex
FUNCTION CACos (BYREF value AS _complex) AS _complex
```

#### Example

```
DIM c AS _complex = (1, 1)
print CStr(CArcCos(c))
Output: 0.9045568943023814 -1.061275061905036 * i
```

# <a name="CArcCosH"></a>CArcCosH / CACosH

Returns the complex hyperbolic arccosine of a complex number. The branch cut is on the real axis, less than 1.

```
FUNCTION CArcCosH (BYREF value AS _complex) AS _complex
FUNCTION CACosH (BYREF value AS _complex) AS _complex
```

#### Example

```
DIM c AS _complex = (1, 1)
print CStr(CArcCosH(c))
Output: 1.061275061905036 +0.9045568943023813 * i
```

# <a name="CArcCosHReal"></a>CArcCosHReal / CACosHReal

Returns the complex hyperbolic arccosine of a real number.

```
FUNCTION CArcCosHReal (BYVAL value AS DOUBLE) AS _complex
FUNCTION CACosHReal (BYVAL value AS DOUBLE) AS _complex
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
FUNCTION CArcCosReal (BYVAL value AS DOUBLE) AS _complex
FUNCTION CACosReal (BYVAL value AS DOUBLE) AS _complex
```

| Parameter  | Description |
| ---------- | ----------- |
| *value* | A double value representing the real part of a complex number. |

#### Example

```
print CStr(CArcCosReal(1)) ' = 0 0 * i
print CStr(CArcCosReal(-1)) ' = 3.141592653589793 0 * i
print CStr(CArcCosReal(2)) ' = 0 +1.316957896924817 * i
```

# <a name="CArcCot"></a>CArcCot / CACot

Returns the complex arccotangent of a complex number.

```
FUNCTION CArcCot (BYREF value AS _complex) AS _complex
FUNCTION CACot (BYREF value AS _complex) AS _complex
```

#### Example

```
DIM z AS _complex = (1, 1)
z = CArcCot(z)
PRINT CStr(z)
Output: 0.5535743588970452 -0.4023594781085251 * i
```

# <a name="CArcCotH"></a>CArcCotH / CACotH

Returns the complex hyperbolic arccotangent of a complex number.

```
FUNCTION CArcCotH (BYREF value AS _complex) AS _complex
FUNCTION CACotH (BYREF value AS _complex) AS _complex
```

#### Example

```
DIM c AS _complex = (1, 1)
print CStr(CArcCotH(c))
Output: 0.4023594781085251 -0.5535743588970452 * i
```

# <a name="CArcCsc"></a>CArcCsc / CACsc

Returns the complex arccosecant of a complex number.

```
FUNCTION CArcCsc (BYREF value AS _complex) AS _complex
FUNCTION CACsc (BYREF value AS _complex) AS _complex
```

#### Example

```
DIM c AS _complex = (1, 1)
print CStr(CArcCsc(c))
Output: 0.4522784471511907 -0.5306375309525178 * i
```

# <a name="CArcCscH"></a>CArcCscH / CACscH

Returns the complex hyperbolic arccosecant of a complex number.

```
FUNCTION CArcCscH (BYREF value AS _complex) AS _complex
FUNCTION CACscH (BYREF value AS _complex) AS _complex
```

#### Example

```
DIM c AS _complex = (1, 1)
print CStr(CArcCscH(c))
Output: 0.5306375309525179 -0.4522784471511906 * i
```

# <a name="CArcCscReal"></a>CArcCscReal / CACscReal

Returns the complex arccosecant of a real number. 

```
FUNCTION CArcCscReal (BYVAL value AS DOUBLE) AS _complex
FUNCTION CACscReal (BYVAL value AS DOUBLE) AS _complex
```

#### Example

```
print CStr(CArcCscReal(1))
Output: 1.570796326794897 0 * i
```

# <a name="CArcSec"></a>CArcSec / CASec

Returns the complex arcsecant of a complex number.

```
FUNCTION CArcSec (BYREF value AS _complex) AS _complex
FUNCTION CASec (BYREF value AS _complex) AS _complex
```

#### Example

```
DIM z AS _complex = (1, 1)
z = CArcSec(z)
PRINT CStr(z)
Output: 1.118517879643706 +0.5306375309525176 * i
```

# <a name="CArcSecH"></a>CArcSecH / CASecH

Returns the complex hyperbolic arcsecant of a complex number.

```
FUNCTION CArcSecH (BYREF value AS _complex) AS _complex
FUNCTION CASecH (BYREF value AS _complex) AS _complex
```

#### Example

```
DIM z AS _complex = (1, 1)
z = CArcSecH(z)
PRINT CStr(z)
Output: 0.5306375309525178 -1.118517879643706 * i
```

# <a name="CArcSecReal"></a>CArcSecReal / CASecReal

Returns the complex arcsecant of a real number.

```
FUNCTION CArcSecReal (BYVAL value AS DOUBLE) AS _complex
FUNCTION CASecReal (BYVAL value AS DOUBLE) AS _complex
```

| Parameter  | Description |
| ---------- | ----------- |
| *value* | A double value representing the real part of a complex number. |

#### Example

```
print CStr(CArcSecReal(1.1))
Output: 0.4296996661514246 0 * i
```

# <a name="CArcSin"></a>CArcSin / CASin

Returns the complex arcsine of a complex number. The branch cuts are on the real axis, less than -1 and greater than 1.

```
FUNCTION CArcSin (BYREF value AS _complex) AS _complex
FUNCTION CASin (BYREF value AS _complex) AS _complex
```

#### Example

```
DIM z AS _complex = (1, 1)
z = CArcSin(z)
PRINT CStr(z)
Output: 0.6662394324925152 +1.061275061905036 * i
```

# <a name="CArcSinH"></a>CArcSinH / CASinH

Returns the complex hyperbolic arcsine of a complex number. The branch cuts are on the imaginary axis, below -i and above i.

```
FUNCTION CArcSinH (BYREF value AS _complex) AS _complex
FUNCTION CASinH (BYREF value AS _complex) AS _complex
```

#### Example

```
DIM z AS _complex = (1, 1)
z = CArcSinH(z)
PRINT CStr(z)
Output: 1.061275061905036 +0.6662394324925153 * i
```

# <a name="CArcSinReal"></a>CArcSinReal / CASinReal

Returns the complex arcsine of a real number.<br>
For a between -1 and 1, the function returns a real value in the range \[-pi/2, pi/2].<br>
For a less than -1 the result has a real part of -pi/2 and a positive imaginary part.<br>
For a greater than 1 the result has a real part of pi/2 and a negative imaginary part.

```
FUNCTION CArcSinReal (BYVAL value AS DOUBLE) AS _complex
FUNCTION CASinReal (BYVAL value AS DOUBLE) AS _complex
```

| Parameter  | Description |
| ---------- | ----------- |
| *value* | A double value representing the real part of a complex number. |

#### Example

```
DIM z AS _complex = CArcSinReal(1)
PRINT CStr(z)
Output: 1.570796326794897 +0 * i
```

# <a name="CArcTan"></a>CArcTan / CATan

Returns the complex arctangent of a complex number. The branch cuts are on the imaginary axis, below -i and above i.

```
FUNCTION CArcTan (BYREF value AS _complex) AS _complex
FUNCTION CATan (BYREF value AS _complex) AS _complex
```

#### Example

```
DIM z AS _complex = (1, 1)
z = CArcTan(z)
PRINT CStr(z)
Output: 1.017221967897851 +0.4023594781085251 * i
```

# <a name="CArcTanH"></a>CArcTanH / CATanH

Returns the complex hyperbolic arctangent of a complex number. The branch cuts are on the real axis, less than -1 and greater than 1.

```
FUNCTION CArcTanH (BYREF value AS _complex) AS _complex
FUNCTION CATanH (BYREF value AS _complex) AS _complex
```

#### Example

```
DIM z AS _complex = (1, 1)
z = CArcTanH(z)
PRINT CStr(z)
Output: 0.4023594781085251 +1.017221967897851 * i
```

# <a name="CArcTanHReal"></a>CArcTanHReal / CATanHReal

Returns the complex hyperbolic arctangent of a real number.

```
FUNCTION CArcTanhReal (BYVAL value AS DOUBLE) AS _complex
FUNCTION CATanhReal (BYVAL value AS DOUBLE) AS _complex
```

| Parameter  | Description |
| ---------- | ----------- |
| *value* | A double value representing the real part of a complex number. |

# <a name="CArg"></a>CArg / CArgument / CPhase

Returns the argument of this complex number.

```
FUNCTION CArg (BYREF z AS _complex) AS DOUBLE
FUNCTION CArgument (BYREF z AS _complex) AS DOUBLE
FUNCTION CPhase (BYREF z AS _complex) AS DOUBLE
```

#### Example

```
DIM z AS _complex = (1, 0)
DIM d AS DOUBLE = CArg(z)
PRINT d
Output: 0
```
```
DIM z AS _complex = (0, 1)
DIM d AS DOUBLE = CArg(z)
PRINT d
Output: 1.570796326794897
```
```
DIM z AS _complex = (0, -1)
DIM d AS DOUBLE = CArg(z)
PRINT d
Output: -1.570796326794897
```
```
DIM z AS _complex = (-1, 0)
DIM d AS DOUBLE = CArg(z)
PRINT d
Output: 3.141592653589793
```

# <a name="CConjugate"></a>CConjugate / CConj

Returns the complex conjugate of a complex number.

```
FUNCTION CConjugate (BYREF z AS _complex) AS _complex
FUNCTION CConj (BYREF z AS _complex) AS _complex
```

#### Example

```
DIM z AS _complex = (2, 3)
z = CConj(z)
PRINT CStr(z)
Output: 2 - 3 * i
```

# <a name="CCos"></a>CCos

Returns the complex cosine of a complex number.

```
FUNCTION CCos (BYREF value AS _complex) AS _complex
```

#### Example

```
DIM z AS _complex = (1, 1)
z = CCos(z)
PRINT CStr(z)
Output: 0.8337300251311491 -0.9888977057628651 * i
```

# <a name="CCosH"></a>CCosH

Returns the complex hyperbolic cosine of a complex number.

```
FUNCTION CCosH (BYREF value AS _complex) AS _complex
```

#### Example

```
DIM z AS _complex = (1, 1)
z = CCosH(z)
PRINT CStr(z)
Output: 0.8337300251311491 +0.9888977057628651 * i
```

# <a name="CCot"></a>CCot

Returns the complex cotangent of a complex number.

```
FUNCTION CCot (BYREF value AS _complex) AS _complex
```

#### Example

```
DIM z AS _complex = (1, 1)
z = CCot(z)
PRINT CStr(z)
Output: 0.2176215618544027 -0.8680141428959249 * i
```

# <a name="CCotH"></a>CCotH

Returns the complex hyperbolic cotangent of a complex number.

```
FUNCTION CCotH (BYREF value AS _complex) AS _complex
```

#### Example

```
DIM z AS _complex = (1, 1)
z = CCotH(z)
PRINT CStr(z)
Output: 0.8680141428959249 -0.2176215618544029 * i
```

# <a name="CCsc"></a>CCsc

Returns the complex cosecant of a complex number.

```
FUNCTION CCsc (BYREF value AS _complex) AS _complex
```

#### Example

```
DIM z AS _complex = (1, 1)
z = CCsc(z)
PRINT CStr(z)
Output: 0.6215180171704285 -0.3039310016284265 * i
```

# <a name="CCscH"></a>CCscH

Returns the complex hyperbolic cosecant of a complex number.

```
FUNCTION CCscH (BYREF value AS _complex) AS _complex
```

#### Example

```
DIM z AS _complex = (1, 1)
z = CCscH(z)
PRINT CStr(z)
Output: 0.3039310016284265 -0.6215180171704285 * i
```

# <a name="CDivImag"></a>CDivImag

Divides by an imaginary number.

```
FUNCTION CDivImag (BYREF z AS _complex, BYVAL y AS DOUBLE) AS _complex
```

| Parameter  | Description |
| ---------- | ----------- |
| *y* | A double value. |

# <a name="CDivReal"></a>CDivReal

Divides by a real number.

```
FUNCTION CDivReal (BYREF z AS _complex, BYVAL x AS DOUBLE) AS _complex
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | A double value. |

# <a name="CExp"></a>CExp

Returns the complex exponential of a complex number.

```
FUNCTION CExp (BYREF value AS _complex) AS _complex
```

#### Example

```
DIM z AS _complex = (1, 1)
z = CExp(z)
PRINT CStr(z)
Output: 1.468693939915885 +2.287355287178842 * i
```

# <a name="CGetImag"></a>CGetImag

Returns the imaginary part of a complex number.

```
FUNCTION CGetImag (BYREF z AS _complex) AS DOUBLE
```

# <a name="CGetReal"></a>CGetReal

Returns the real part of a complex number.

```
FUNCTION CGetReal (BYREF z AS _complex) AS DOUBLE
```

# <a name="CLog"></a>CLog

Returns the complex natural logarithm (base e) of a complex number. The branch cut is the negative real axis.

```
FUNCTION CLog OVERLOAD (BYREF value AS _complex) AS _complex
FUNCTION CLog OVERLOAD (BYREF value AS _complex, BYVAL baseValue AS DOUBLE) AS _complex
```

| Parameter  | Description |
| ---------- | ----------- |
| *baseValue* | A double value. |

#### Examples

```
DIM z AS _complex = (1, 1)
z = CLog(z)
PRINT CStr(z)
Output: 0.3465735902799727 +0.7853981633974483 * i
```
```
DIM z AS _complex = (0, 0)
z = CLog(z)
PRINT CStr(z)
Output: -1.#INF
```
```
DIM z AS _complex = (1, 1)
z = CLog(z, 10)
PRINT CStr(z)
Output: 0.1505149978319906 +0.3410940884604603 * i
```

# <a name="CLog10"></a>CLog10

Returns the complex base-10 logarithm of a complex number.

```
FUNCTION CLog10 (BYREF value AS _complex) AS _complex
```

#### Example

```
DIM z AS _complex = (1, 1)
z = CLog10(z)
PRINT CStr(z)
Output: 0.1505149978319906 +0.3410940884604603 * i
```

# <a name="CLogAbs"></a>CLogAbs

Returns the natural logarithm of the magnitude of the complex number z, log|z|. It allows an accurate evaluation of \log|z| when |z| is close to one. The direct evaluation of log(CAbs(z)) would lead to a loss of precision in this case.

```
FUNCTION CLogAbs (BYREF value AS _complex) AS DOUBLE
```

#### Example

```
DIM z AS _complex = (1.1, 0.1)
DIM d AS DOUBLE = CLogAbs(z)
print d
Output: 0.09942542937258279
```

# <a name="CModulus"></a>CModulus / CMod

Returns the modulus of a complex number.

```
FUNCTION CModulus (BYREF z AS _complex) AS DOUBLE
FUNCTION CMod (BYREF z AS _complex) AS DOUBLE
```

#### Example

```
DIM z AS _complex = (2.3, -4.5)
print CModulus(z)
Output: 5.053711507397311
```
```
DIM z AS _complex = CPolar(0.2938, -0.5434)
print CModulus(z)
Output: 0.2938
```

# <a name="CMulImag"></a>CMulImag

Multiplies by an imaginary number.

```
FUNCTION CMulImag (BYREF z AS _complex, BYVAL y AS DOUBLE) AS _complex
```

| Parameter  | Description |
| ---------- | ----------- |
| *y* | A double value. |

#### Example

```
DIM z AS _complex = (5, 6)
z = CMulImag(z, 3)
PRINT CStr(z)
```

# <a name="CMulReal"></a>CMulReal

Multiplies by a real number.

```
FUNCTION CMulReal (BYREF z AS _complex, BYVAL x AS DOUBLE) AS _complex
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | A double value. |

#### Example

```
DIM z AS _complex = (5, 6)
z = CMulReal(z, 3)
PRINT CStr(z)
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
FUNCTION CNthRoot (BYREF z AS _complex, BYVAL n AS LONG, BYVAL k AS LONG = 0) AS _complex
```

#### Example

```
DIM z AS _complex = (2.3, -4.5)
DIM n AS LONG = 5
FOR i AS LONG = 0 TO n - 1
   print CStr(CNthRoot(z, n, i))
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
FUNCTION CPolar (BYVAL r AS DOUBLE, BYVAL theta AS DOUBLE) AS _complex
```

| Parameter  | Description |
| ---------- | ----------- |
| *r* | The modulus of complex number. |
| *theta* | The angle with the positive direction of x-axis. |

#### Example

```
DIM z AS _complex = CPolar(3, 4)
```

# <a name="CPow"></a>CPow

Returns this complex number raised to a complex power.<br>
Returns this complex number raised to a real number.

```
FUNCTION CPow OVERLOAD (BYREF value AS _complex, BYREF power AS _complex) AS _complex
FUNCTION CPow OVERLOAD (BYREF value AS _complex, BYVAL power AS DOUBLE) AS _complex
```

| Parameter  | Description |
| ---------- | ----------- |
| *power* | A complex or a double number. |

#### Examples

```
DIM z AS _complex = (1, 1)
DIM b AS _complex = (2, 2)
print CStr(CPow(z, b))
Output: -0.2656539988492412 +0.3198181138561362 * i
```
```
DIM z AS _complex = (1, 1)
PRINT CStr(CPow(z, 2))
Output: 1.224606353822378e-016 +2 * i
```

# <a name="CReciprocal"></a>CReciprocal / CInverse

Returns the inverse, or reciprocal, of a complex number.

```
FUNCTION CReciprocal (BYREF value AS _complex) AS _complex
FUNCTION CInverse (BYREF value AS _complex) AS _complex
```

#### Example

```
DIM z AS _complex = (1, 1)
z = CReciprocal(z)
PRINT CStr(z)
Output: 0.5 -0.5 * i
```

# <a name="CSec"></a>CSec

Returns the complex secant of a complex number.

```
FUNCTION CSec (BYREF value AS _complex) AS _complex
```

#### Example

```
DIM z AS _complex = (1, 1)
z = CSec(z)
PRINT CStr(z)
Output: 0.4983370305551869 +0.591083841721045 * i
```

# <a name="CSecH"></a>CSecH

Returns the complex hyperbolic secant of a complex number.

```
FUNCTION CSecH (BYREF value AS _complex) AS _complex
```

#### Example

```
DIM z AS _complex = (1, 1)
z = CSecH(z)
PRINT CStr(z)
Output: 0.4983370305551869 -0.591083841721045 * i
```

# <a name="CSet"></a>CSet / CRect

Uses the cartesian components (x,y) to set the real and imaginary parts of the complex number.

```
FUNCTION CSet (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE) AS _complex
FUNCTION CRet (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE) AS _complex
```

| Parameter  | Description |
| ---------- | ----------- |
| *x, y* | Double values. |

#### Example

```
DIM z AS _complex = CSet(3, 4)
```

# <a name="CSetImag"></a>CSetImag

Sets the imaginary part of a complex number.

```
SUB CSetImag (BYREF z AS _complex, BYVAL y AS DOUBLE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *z* | The complex number. |
| *y* | A double value. |

# <a name="CSetReal"></a>CSetReal

Sets the real part of a complex number.

```
SUB CSetReal (BYREF z AS _complex, BYVAL y AS DOUBLE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *z* | The complex number. |
| *y* | A double value. |

# <a name="CSgn"></a>CSgn

Returns the sign of this complex number.<br>
If number is greater than zero, then CSgn returns 1.<br>
If number is equal to zero, then CSgn returns 0.<br>
If number is less than zero, then CSgn returns -1.

```
FUNCTION CSgn (BYREF z AS _complex) AS LONG
```

# <a name="CSin"></a>CSin

Returns the complex sine of a complex number.

```
FUNCTION CSin (BYREF value AS _complex) AS _complex
```

#### Example

```
DIM z AS _complex = (1, 1)
z = CSin(z)
PRINT CStr(z)
Output: 1.298457581415977 +0.6349639147847361 * i
```

# <a name="CSinH"></a>CSinH

Returns the complex hyperbolic sine of a complex number.

```
FUNCTION CSinH (BYREF value AS _complex) AS _complex
```

#### Example

```
DIM z AS _complex = (1, 1)
z = CSinH(z)
PRINT CStr(z)
Output: 0.6349639147847361 +1.298457581415977 * i
```

# <a name="CSqr"></a>CSqr / CSqrt

Returns the square root of the complex number z. The branch cut is the negative real axis. The result always lies in the right half of the complex plane.

```
FUNCTION CSqr (BYREF value AS _complex) AS _complex
FUNCTION CSqrt (BYREF value AS _complex) AS _complex
```

| Parameter  | Description |
| ---------- | ----------- |
| *value* | A complex number. |

#### Examples

```
DIM z AS _complex = (2, 3)
z = CSqr(z)
PRINT CStr(z)
Output: 1.67414922803554 +0.895977476129838 * i
```

Compute the square root of -1:

```
DIM z AS _complex = (-1)
z = CSqr(z)
PRINT CStr(z)
Output: 0 +1.0 * i
```

# <a name="CSqrReal"></a>CSqrReal

Returns the complex square root of the real number x, where x may be negative.

```
FUNCTION CSqrReal (BYVAL value AS DOUBLE) AS _complex
```

| Parameter  | Description |
| ---------- | ----------- |
| *value* | A double value representing a real number. |


# <a name="CSubImag"></a>CSubImag

Subtracts an imaginary number.

```
FUNCTION CSubImag (BYREF z AS _complex, BYVAL y AS DOUBLE) AS _complex
```

| Parameter  | Description |
| ---------- | ----------- |
| *y* | A double value. |

#### Example

```
DIM z AS _complex = (5, 6)
z = CSubImag(z, 3)
PRINT CStr(z)
Output:  5 +3 * i
```

# <a name="CSubReal"></a>CSubReal

Subtracts a real number.

```
FUNCTION CSubReal (BYREF z AS _complex, BYVAL x AS DOUBLE) AS _complex
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | A double value. |

#### Example

```
DIM z AS _complex = (5, 6)
z = CSubReal(z, 2)
PRINT CStr(z)
Output: 3 +6 * i
```

# <a name="CSwap"></a>CSwap

Exchanges the contents of two complex numbers.

```
SUB CSwap (BYREF z1 AS _complex, BYREF z2 AS _complex)
```

| Parameter  | Description |
| ---------- | ----------- |
| *z1 / z2* | The complex numbers to swap. |

# <a name="CTan"></a>CTan

Returns the complex tangent of a complex number.

```
FUNCTION CTan (BYREF value AS _complex) AS _complex
```

#### Example

```
DIM z AS _complex = (1, 1)
z = CTan(z)
PRINT CStr(z)
Output: 0.2717525853195117 +1.083923327338695 * i
```

# <a name="CTanH"></a>CTanH

Returns the complex hyperbolic tangent of a complex number.

```
FUNCTION CTanH (BYREF value AS _complex) AS _complex
```

#### Example

```
DIM z AS _complex = (1, 1)
z = CTanH(z)
PRINT CStr(z)
Output: 1.083923327338695 +0.2717525853195119 * i
```

# <a name="acosh"></a>acosh

Calculates the inverse hyperbolic cosine.

```
FUNCTION acosh (BYVAL x AS DOUBLE) AS DOUBLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | A double value. |

# <a name="asinh"></a>asinh

Calculates the inverse hyperbolic sine.

```
FUNCTION asinh (BYVAL x AS DOUBLE) AS DOUBLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | A double value. |


# <a name="atanh"></a>atanh

Returns the inverse hyperbolic tangent of a number.

```
FUNCTION atanh (BYVAL x AS DOUBLE) AS DOUBLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | A double value. |


# <a name="hypot3"></a>hypot3

Computes the value of sqr(x^2 + y^2 + z^2).

```
FUNCTION hypot3 (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE, BYVAL z AS DOUBLE) AS DOUBLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *x, y, z* | Double values. |

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
