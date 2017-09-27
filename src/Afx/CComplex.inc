' ########################################################################################
' Microsoft Windows
' File: CComplex.inc
' Contents: Complex numbers.
' Compiler: FreeBasic 32 & 64-bit
' Written in 2017 by José Roca. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

' ########################################################################################
' Note: Some methods are based in the .NET Complex.cs class
' https://github.com/Microsoft/referencesource/blob/master/System.Numerics/System/Numerics/Complex.cs
' and, therefore, subject to the MIT license.
' Copyright (c) Microsoft Corporation.  All rights reserved.
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "crt/math.bi"
#include once "crt/limits.bi"
#INCLUDE ONCE "Afx/AfxComplex.inc"
using Afx

NAMESPACE Afx

#ifndef DBL_EPSILON
#define DBL_EPSILON      2.2204460492503131e-016   ' /* smallest such that 1.0+DBL_EPSILON != 1.0 */
#endif
#ifndef SQRT_DBL_EPSILON
#define SQRT_DBL_EPSILON 1.4901161193847656e-08
#endif

#ifndef gsl_complex
type gsl_complex
	x as double
	y as double
end type
#endif

' ========================================================================================
' Macro for debug
' To allow debugging, define _CCOMPLEX_DEBUG_ 1 in your application before including this file.
' ========================================================================================
#ifndef _CCOMPLEX_DEBUG_
   #define _CCOMPLEX_DEBUG_ 0
#ENDIF
#ifndef _CCOMPLEX_DP_
   #define _CCOMPLEX_DP_ 1
   #MACRO CCOMPLEX_DP(st)
      #IF (_CCOMPLEX_DEBUG_ = 1)
         OutputDebugStringW(st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

' ########################################################################################
' CComplex class
' ########################################################################################
TYPE CComplex

   m_cpx AS _complex

   ' // Constructors and destructor
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL x AS DOUBLE = 0, BYVAL y AS DOUBLE = 0)
   DECLARE CONSTRUCTOR (BYREF cpx AS CComplex)
   DECLARE CONSTRUCTOR (BYREF cpx AS _complex)
   DECLARE CONSTRUCTOR (BYREF cpx AS gsl_complex)
   DECLARE DESTRUCTOR
   ' // Assignment and cast
   DECLARE OPERATOR LET (BYREF z AS CComplex)
   DECLARE OPERATOR LET (BYREF z AS _complex)
   DECLARE OPERATOR LET (BYREF z AS gsl_complex)
   DECLARE OPERATOR CAST () AS _complex
   DECLARE OPERATOR CAST () AS gsl_complex
   DECLARE OPERATOR CAST () AS STRING
   ' // Complex number methods
   DECLARE FUNCTION IsInf (BYVAL x AS DOUBLE) AS LONG
   DECLARE FUNCTION IsInfinity (BYVAL x AS DOUBLE) AS LONG
   DECLARE FUNCTION ArcCosH (BYVAL x AS DOUBLE) AS DOUBLE
   DECLARE FUNCTION ArcTanH (BYVAL x AS DOUBLE) AS DOUBLE
   DECLARE PROPERTY CSet (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE)
   DECLARE PROPERTY CRect (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE)
   DECLARE PROPERTY CReal () AS DOUBLE
   DECLARE PROPERTY CReal (BYVAL x AS DOUBLE)
   DECLARE PROPERTY CImag () AS DOUBLE
   DECLARE PROPERTY CImag (BYVAL y AS DOUBLE)
   DECLARE PROPERTY CPolar (BYVAL r AS DOUBLE, BYVAL theta AS DOUBLE)
   DECLARE SUB CSwap (BYREF z AS CComplex)
   DECLARE SUB CNegate
   DECLARE SUB CNeg
   DECLARE SUB CNegative
   ' // Operators
   DECLARE OPERATOR += (BYREF z AS CComplex)
   DECLARE OPERATOR -= (BYREF z AS CComplex)
   DEClaRE OPERATOR *= (BYREF z AS CComplex)
   DEClaRE OPERATOR /= (BYREF z AS CComplex)
   ' // Complex arithmetic methods
   DECLARE FUNCTION CAdd (BYREF z AS CComplex) AS CComplex
   DECLARE FUNCTION CAdd (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE) AS CComplex
   DECLARE FUNCTION CAddReal (BYVAL x AS DOUBLE) AS CComplex
   DECLARE FUNCTION CAddImag (BYVAL y AS DOUBLE) AS CComplex
   DECLARE FUNCTION CSub (BYREF z AS CComplex) AS CComplex
   DECLARE FUNCTION CSub (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE) AS CComplex
   DECLARE FUNCTION CSubReal (BYVAL x AS DOUBLE) AS CComplex
   DECLARE FUNCTION CSubImag (BYVAL y AS DOUBLE) AS CComplex
   DECLARE FUNCTION CMul (BYREF z AS CComplex) AS CComplex
   DECLARE FUNCTION CMul (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE) AS CComplex
   DECLARE FUNCTION CMulReal (BYVAL x AS DOUBLE) AS CComplex
   DECLARE FUNCTION CMulImag (BYVAL y AS DOUBLE) AS CComplex
   DECLARE FUNCTION CDiv (BYREF z AS CComplex) AS CComplex
   DECLARE FUNCTION CDiv (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE) AS CComplex
   DECLARE FUNCTION CDivReal (BYVAL x AS DOUBLE) AS CComplex
   DECLARE FUNCTION CDivImag (BYVAL y AS DOUBLE) AS CComplex
   DECLARE FUNCTION CConj () AS CComplex
   DECLARE FUNCTION CConjugate () AS CComplex
   '// Properties of complex numbers
   DECLARE FUNCTION CArg () AS DOUBLE
   DECLARE FUNCTION CArgument () AS DOUBLE
   DECLARE FUNCTION CPhase () AS DOUBLE
   DECLARE FUNCTION CAbs () AS DOUBLE
   DECLARE FUNCTION CMagnitude () AS DOUBLE
   DECLARE FUNCTION CAbs2 () AS DOUBLE
   DECLARE FUNCTION CLogAbs () AS DOUBLE
   DECLARE FUNCTION CNorm () AS DOUBLE
   DECLARE FUNCTION CInverse () AS CComplex
   DECLARE FUNCTION CReciprocal () AS CComplex
   DECLARE FUNCTION CFromPolarCoordinates (BYVAL magnitude AS DOUBLE, BYVAL phase AS DOUBLE) AS CComplex
   DECLARE FUNCTION CSqr () AS CComplex
   DECLARE FUNCTION CSqrt () AS CComplex
   DECLARE FUNCTION CSqr (BYVAL x AS DOUBLE) AS CComplex
   DECLARE FUNCTION CSqrt (BYVAL x AS DOUBLE) AS CComplex
   DECLARE FUNCTION CExp () AS CComplex
   DECLARE FUNCTION CLog () AS CComplex
   DECLARE FUNCTION CLog (BYVAL baseValue AS DOUBLE) AS CComplex
   DECLARE FUNCTION CLog10 () AS CComplex
   DECLARE FUNCTION CPow (BYREF power AS CComplex) AS CComplex
   DECLARE FUNCTION CPow (BYVAL power AS DOUBLE) AS CComplex
   DECLARE FUNCTION CSin () AS CComplex
   DECLARE FUNCTION CCos () AS CComplex
   DECLARE FUNCTION CSec () AS CComplex
   DECLARE FUNCTION CCsc () AS CComplex
   DECLARE FUNCTION CTan () AS CComplex
   DECLARE FUNCTION CCot () AS CComplex
   DECLARE FUNCTION CArcSinReal (BYVAL value AS DOUBLE) AS CComplex
   DECLARE FUNCTION CASinReal (BYVAL value AS DOUBLE) AS CComplex
   DECLARE FUNCTION CArcSin () AS CComplex
   DECLARE FUNCTION CASin () AS CComplex
   DECLARE FUNCTION CArcCosReal (BYVAL value AS DOUBLE) AS CComplex
   DECLARE FUNCTION CACosReal (BYVAL value AS DOUBLE) AS CComplex
   DECLARE FUNCTION CArcCos () AS CComplex
   DECLARE FUNCTION CACos () AS CComplex
   DECLARE FUNCTION CArcSecReal (BYVAL value AS DOUBLE) AS CComplex
   DECLARE FUNCTION CASecReal (BYVAL value AS DOUBLE) AS CComplex
   DECLARE FUNCTION CArcSec () AS CComplex
   DECLARE FUNCTION CASec () AS CComplex
   DECLARE FUNCTION CArcCscReal (BYVAL value AS DOUBLE) AS CComplex
   DECLARE FUNCTION CACscReal (BYVAL value AS DOUBLE) AS CComplex
   DECLARE FUNCTION CArcCsc () AS CComplex
   DECLARE FUNCTION CACsc () AS CComplex
   DECLARE FUNCTION CArcTan () AS CComplex
   DECLARE FUNCTION CATan () AS CComplex
   DECLARE FUNCTION CArcCot () AS CComplex
   DECLARE FUNCTION CACot () AS CComplex
   DECLARE FUNCTION CSinH () AS CComplex
   DECLARE FUNCTION CCosH () AS CComplex
   DECLARE FUNCTION CSecH () AS CComplex
   DECLARE FUNCTION CCscH () AS CComplex
   DECLARE FUNCTION CTanH () AS CComplex
   DECLARE FUNCTION CCotH () AS CComplex
   DECLARE FUNCTION CArcSinH () AS CComplex
   DECLARE FUNCTION CASinH () AS CComplex
   DECLARE FUNCTION CArcCosHReal (BYVAL value AS DOUBLE) AS CComplex
   DECLARE FUNCTION CACosHReal (BYVAL value AS DOUBLE) AS CComplex
   DECLARE FUNCTION CACosHReal () AS CComplex
   DECLARE FUNCTION CArcCosH () AS CComplex
   DECLARE FUNCTION CACosH () AS CComplex
   DECLARE FUNCTION CArcSecH () AS CComplex
   DECLARE FUNCTION CASecH () AS CComplex
   DECLARE FUNCTION CArcCscH () AS CComplex
   DECLARE FUNCTION CACscH () AS CComplex
   DECLARE FUNCTION CArcTanHReal (BYVAL value AS DOUBLE) AS CComplex
   DECLARE FUNCTION CATanHReal (BYVAL value AS DOUBLE) AS CComplex
   DECLARE FUNCTION CArcTanH () AS CComplex
   DECLARE FUNCTION CATanH () AS CComplex
   DECLARE FUNCTION CArcCotH () AS CComplex
   DECLARE FUNCTION CACotH () AS CComplex
   DECLARE FUNCTION CSgn () AS LONG

END TYPE

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CComplex
   CCOMPLEX_DP("CCOMPLEX CONSTRUCTOR")
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' * Uses the cartesian components (x,y) to set the real and imaginary parts of the complex number.
' Examples:
' DIM cpx AS CComplex = CComplex(3, 4)
' DIM cpx AS CComplex = TYPE(3, 4)
' ========================================================================================
PRIVATE CONSTRUCTOR CComplex (BYVAL x AS DOUBLE = 0, BYVAL y AS DOUBLE = 0)
   CCOMPLEX_DP("CCOMPLEX CONSTRUCTOR x, y")
   m_cpx = TYPE<_complex>(x, y)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Example: DIM cpx AS CComplex = TYPE<_complex>(3, 4)
' ========================================================================================
PRIVATE CONSTRUCTOR CComplex (BYREF cpx AS _complex)
   CCOMPLEX_DP("CCOMPLEX CONSTRUCTOR _complex")
   m_cpx = cpx
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Example: DIM cpx AS CComplex = TYPE<gsl_complex>(3, 4)
' ========================================================================================
PRIVATE CONSTRUCTOR CComplex (BYREF cpx AS gsl_complex)
   CCOMPLEX_DP("CCOMPLEX CONSTRUCTOR gsl_complex")
   m_cpx = TYPE<_complex>(cpx.x, cpx.y)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Example: DIM cpx AS CComplex = TYPE(3, 4) : DIM cpx2 AS CComplex = cpx
' ========================================================================================
PRIVATE CONSTRUCTOR CComplex (BYREF cpx AS CComplex)
   CCOMPLEX_DP("CCOMPLEX CONSTRUCTOR CComplex")
   m_cpx = cpx.m_cpx
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE DESTRUCTOR CComplex
   CCOMPLEX_DP("CCOMPLEX DESTRUCTOR")
   ' // Nothing to do
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' * Assigns a complex number.
' Example: DIM cpx AS CComplex = TYPE(3, 4) : DIM cpx2 AS CComplex : cpx2 = cpx
' Example: DIM cpx AS CComplex : cpx = CComplex(3, 4)
' ========================================================================================
PRIVATE OPERATOR CComplex.LET (BYREF z AS CComplex)
   CCOMPLEX_DP("CCOMPLEX LET CComplex")
   m_cpx = z.m_cpx
END OPERATOR
' ========================================================================================
' ========================================================================================
' Example: DIM cpx AS CComplex : cpx = TYPE<_complex>(3, 4)
' ========================================================================================
PRIVATE OPERATOR CComplex.LET (BYREF z AS _complex)
   CCOMPLEX_DP("CCOMPLEX LET _complex")
   m_cpx = z
END OPERATOR
' ========================================================================================
' ========================================================================================
' Example: DIM cpx AS CComplex : cpx = TYPE<gsl_complex>(3, 4)
' ========================================================================================
PRIVATE OPERATOR CComplex.LET (BYREF z AS gsl_complex)
   CCOMPLEX_DP("CCOMPLEX LET gsl_complex")
   m_cpx = TYPE<_complex>(z.x, z.y)
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Returns the underlying _complex number.
' ========================================================================================
PRIVATE OPERATOR CComplex.CAST () AS _complex
   CCOMPLEX_DP("CCOMPLEX CAST _complex")
   OPERATOR = m_cpx
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CComplex.CAST () AS gsl_complex
   CCOMPLEX_DP("CCOMPLEX CAST gsl_complex")
   OPERATOR = TYPE<gsl_complex>(m_cpx.x, m_cpx.y)
END OPERATOR
' ========================================================================================
' ========================================================================================
' * Returns the underlying _complex number as a formated string.
' ========================================================================================
PRIVATE OPERATOR CComplex.CAST () AS STRING
   CCOMPLEX_DP("CCOMPLEX CAST String")
   OPERATOR = STR(m_cpx.x) & IIF(SGN(m_cpx.y) = -1 OR SGN(m_cpx.y) = 0, " ", " +") & STR(m_cpx.y) & " * i"
END OPERATOR
' ========================================================================================

' ========================================================================================
' Determines whether the argument is an infinity.
' Returns +1 if x is positive infinity, -1 if x is negative infinity and 0 otherwise.
' ========================================================================================
PRIVATE FUNCTION CComplex.IsInf (BYVAL x AS DOUBLE) AS LONG
   CCOMPLEX_DP("CCOMPLEX IsInf")
   IF _finite(x) = TRUE AND _isnan(x) = FALSE THEN
      FUNCTION = IIF(x > 0, 1, -1)
   END IF
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CComplex.IsInfinity (BYVAL x AS DOUBLE) AS LONG
   CCOMPLEX_DP("CCOMPLEX IsInfinity")
   IF _finite(x) = TRUE AND _isnan(x) = FALSE THEN
      FUNCTION = IIF(x > 0, 1, -1)
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Calculates the inverse hyperbolic cosine.
' Example:
'   DIM AS double pi = 3.1415926535
'   DIM AS double x, y
'   DIM c AS CComplex
'   x = cosh(pi / 4)
'   y = c.AfxArcCosH(x)
'   print "cosh = ", pi/4, x
'   print "ArcCosH = ", x, y
' Output:
'   cosh =         0.785398163375              1.324609089232506
'   acosh =        1.324609089232506           0.7853981633749999
' ========================================================================================
PRIVATE FUNCTION CComplex.ArcCosH (BYVAL x AS DOUBLE) AS DOUBLE
   CCOMPLEX_DP("CCOMPLEX ArcCosH")
   DIM t AS DOUBLE
   IF x > 1.0 / SQRT_DBL_EPSILON THEN
      FUNCTION = LOG(x) + M_LN2
   ELSEIF x > 2 THEN
      FUNCTION = log(2 * x - 1 / (SQR(x * x - 1) + x))
   ELSEIF x > 1 THEN
      t = x - 1
      FUNCTION = log1p(t + SQR(2 * t + t * t))
   ELSEIF x = 1 THEN
      FUNCTION = 0
   ELSE
      FUNCTION = 0 / 0  ' NAN
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the inverse hyperbolic tangent of a number.
' Examples:
' print atanh(0.76159416)
'    1.00000000962972
' print atanh(-0.1)
'   -0.1003353477310756
' ========================================================================================
PRIVATE FUNCTION CComplex.ArcTanH (BYVAL x AS DOUBLE) AS DOUBLE
   CCOMPLEX_DP("CCOMPLEX ArcTanH")
   DIM AS DOUBLE a, s
   a = fabs(x)
   s = IIF(x < 0, -1, 1)
   IF a > 1 THEN
      FUNCTION = 0 / 0   ' NAN
   ELSEIF a = 1 THEN
      FUNCTION = IIF(x < 0, HUGE_VALF, -HUGE_VALF)
   ELSEIF a >= 0.5 THEN
      FUNCTION = s * 0.5 * log1p(2 * a / (1 - a))
   ELSEIF a > DBL_EPSILON THEN
      FUNCTION = s * 0.5 * log1p(2 * a + 2 * a * a / (1 - a))
   ELSE
      FUNCTION = x
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Uses the cartesian components (x,y) to set the real and imaginary parts of the complex number.
' Example: DIM cpx AS CComplex : cpx.CSe = 3, 4
' Example: DIM cpx AS CComplex : cpx.CRect = 3, 4
' ========================================================================================
PRIVATE PROPERTY CComplex.CSet (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE)
   CCOMPLEX_DP("CCOMPLEX CSet")
   m_cpx = TYPE<_complex>(x, y)
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CComplex.CRect (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE)
   CCOMPLEX_DP("CCOMPLEX CRect")
   m_cpx = TYPE<_complex>(x, y)
END PROPERTY
' ========================================================================================

' ========================================================================================
' * Gets/sets the real part of a complex numbe.
' ========================================================================================
PRIVATE PROPERTY CComplex.CReal () AS DOUBLE
   CCOMPLEX_DP("CCOMPLEX PROPGET CReal")
   RETURN m_cpx.x
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CComplex.CReal (BYVAL x AS DOUBLE)
   CCOMPLEX_DP("CCOMPLEX PROPSET CReal")
   m_cpx.x = x
END PROPERTY
' ========================================================================================

' ========================================================================================
' * Gets/sets the imaginary part of a complex numbe.
' ========================================================================================
PRIVATE PROPERTY CComplex.CImag () AS DOUBLE
   CCOMPLEX_DP("CCOMPLEX PROPGET CImag")
   RETURN m_cpx.y
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CComplex.CImag (BYVAL y AS DOUBLE)
   CCOMPLEX_DP("CCOMPLEX PROPSET CImag")
   m_cpx.y = y
END PROPERTY
' ========================================================================================

' =====================================================================================
' Sets the complex number from the polar representation.
' =====================================================================================
PRIVATE PROPERTY CComplex.CPolar (BYVAL r AS DOUBLE, BYVAL theta AS DOUBLE)
   CCOMPLEX_DP("CCOMPLEX CPolar")
   m_cpx = TYPE<_complex> (r * cos(theta), r * sin(theta))
END PROPERTY
' =====================================================================================

' ========================================================================================
' Exchanges the contents of two complex numbers.
' ========================================================================================
PRIVATE SUB CComplex.CSwap (BYREF z AS CComplex)
   CCOMPLEX_DP("CCOMPLEX CSwap")
   DIM cpxTemp AS _complex = m_cpx
   m_cpx = z.m_cpx
   z.m_cpx = cpxTemp
END SUB
' ========================================================================================

'/* Comparison operators */

' ========================================================================================
' * Returns true if the two complex numbers are equal, or false otherwise.
' Example: DIM cpx1 AS CComplex = CComplex(1, 2) : DIM cpx2 AS CComplex = CComplex(3,4)
' IF cpx1 = cpx2 THEN PRINT "equal" ELSE PRINT "different"
' ========================================================================================
PRIVATE OPERATOR = (BYREF z1 AS CComplex, BYREF z2 AS CComplex) AS BOOLEAN
   CCOMPLEX_DP("CCOMPLEX OPERATOR = (CComplex, CComplex)")
   RETURN (z1.m_cpx.x = z2.m_cpx.x AND z1.m_cpx.y = z2.m_cpx.y)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns true if the two complex numbers are different, or false otherwise.
' Example: DIM cpx1 AS CComplex = CComplex(1, 2) : DIM cpx2 AS CComplex = CComplex(3,4)
' IF cpx1 <> cpx2 THEN PRINT "different" ELSE PRINT "equal"
' ========================================================================================
PRIVATE OPERATOR <> (BYREF z1 AS CComplex, BYREF z2 AS CComplex) AS BOOLEAN
   CCOMPLEX_DP("CCOMPLEX OPERATOR <> (CComplex, CComplex)")
   RETURN (z1.m_cpx.x <> z2.m_cpx.x OR z1.m_cpx.y <> z2.m_cpx.y)
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Returns the negative of this complex number.
' Example: DIM cpx1 AS CComplex = CComplex(3, 4) : DIM cpx2 AS CComplex = -cpx1
' ========================================================================================
PRIVATE OPERATOR - (BYREF z AS CComplex) AS CComplex
   CCOMPLEX_DP("CCOMPLEX OPERATOR - (CComplex)")
   RETURN TYPE<CComplex>(-z.m_cpx.x, -z.m_cpx.y)
END OPERATOR
' ========================================================================================
' ========================================================================================
' Negates the complex number.
' ========================================================================================
PRIVATE SUB CComplex.CNegate
   CCOMPLEX_DP("CCOMPLEX CNegate")
   m_cpx = TYPE<_complex>(-m_cpx.x, -m_cpx.y)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CComplex.CNeg
   CCOMPLEX_DP("CCOMPLEX CNeg")
   m_cpx = TYPE<_complex>(-m_cpx.x, -m_cpx.y)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CComplex.CNegative
   CCOMPLEX_DP("CCOMPLEX CNegative")
   m_cpx = TYPE<_complex>(-m_cpx.x, -m_cpx.y)
END SUB
' ========================================================================================

' ========================================================================================
' * Returns the sum of complex numbers.
' Example: DIM cpx1 AS CComplex(3, 4) : DIM cpx2 AS CComplex = CComplex(5, 6) : cpx2 = cpx1 + cpx2
' Example: DIM cpx1 AS CComplex(3, 4) : DIM cpx2 = CComplex(5, 6) : cpx2 = cpx1 + 11
' Example: DIM cpx1 AS CComplex(3, 4) : DIM cpx2 = CComplex(5, 6) : cpx2 = 11 + cpx1
' ========================================================================================
PRIVATE OPERATOR + (BYREF z1 AS CComplex, BYREF z2 AS CComplex) AS CComplex
   CCOMPLEX_DP("CCOMPLEX OPERATOR + (CComplex, CComplex)")
   RETURN TYPE<CComplex>(z1.m_cpx.x + z2.m_cpx.x, z1.m_cpx.y + z2.m_cpx.y)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR + (BYVAL a AS DOUBLE, BYREF z AS CComplex) AS CComplex
   CCOMPLEX_DP("CCOMPLEX OPERATOR + (Double, CComplex)")
   RETURN TYPE<CComplex>(z.m_cpx.x + a, z.m_cpx.y)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR + (BYREF z AS CComplex, BYVAL a AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX OPERATOR + (CComplex, Double)")
   RETURN TYPE<CComplex>(z.m_cpx.x + a, z.m_cpx.y)
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Returns the difference of complex numbers.
' Example: DIM cpx1 AS CComplex(3, 4) : DIM cpx2 AS CComplex = CComplex(5, 6) : cpx2 = cpx1 - cpx2
' Example: DIM cpx1 AS CComplex(3, 4) : DIM cpx2 AS CComplex : cpx2 = cpx1 - 11
' Example: DIM cpx1 AS CComplex(3, 4) : DIM cpx2 AS CComplex: cpx2 = 11 - cpx1
' ========================================================================================
PRIVATE OPERATOR - (BYREF z1 AS CComplex, BYREF z2 AS CComplex) AS CComplex
   CCOMPLEX_DP("CCOMPLEX OPERATOR - (CComplex, CComplex)")
   RETURN TYPE<CComplex>(z1.m_cpx.x - z2.m_cpx.x, z1.m_cpx.y - z2.m_cpx.y)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR - (BYVAL a AS DOUBLE, BYREF z AS CComplex) AS CComplex
   CCOMPLEX_DP("CCOMPLEX OPERATOR - (Double, CComplex)")
   RETURN TYPE<CComplex>(a - z.m_cpx.x, -z.m_cpx.y)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR - (BYREF z AS CComplex, BYVAL a AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX OPERATOR - (CComplex, Double)")
   RETURN TYPE<_complex> (z.m_cpx.x - a, z.m_cpx.y)
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Returns the product of complex numbers.
' Example: DIM cpx1 AS CComplex(3, 4) : DIM cpx2 AS CComplex = CComplex(5, 6) : cpx2 = cpx1 * cpx2
' Example: DIM cpx1 AS CComplex(3, 4) : DIM cpx2 AS CComplex = CComplex : cpx2 = cpx1 * 11
' Example: DIM cpx1 AS CComplex(3, 4) : DIM cpx2 AS CComplex = CComplex : cpx2 = 11 * cpx1
' ========================================================================================
PRIVATE OPERATOR * (BYREF z1 AS CComplex, BYREF z2 AS CComplex) AS CComplex
   CCOMPLEX_DP("CCOMPLEX OPERATOR * (CComplex, CComplex)")
   RETURN TYPE<_complex>(z1.m_cpx.x * z2.m_cpx.x - z1.m_cpx.y * z2.m_cpx.y, z1.m_cpx.x * z2.m_cpx.y + z1.m_cpx.y * z2.m_cpx.x)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR * (BYVAL a AS DOUBLE, BYREF z AS CComplex) AS CComplex
   CCOMPLEX_DP("CCOMPLEX OPERATOR * (Double, CComplex)")
   RETURN TYPE<_complex> (a * z.m_cpx.x, a * z.m_cpx.y)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR * (BYREF z AS CComplex, BYVAL a AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX OPERATOR * (CComplex, Double)")
   RETURN TYPE<_complex> (a * z.m_cpx.x, a * z.m_cpx.y)
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Returns the quotient of complex numbers.
' Example: DIM cpx1 AS CComplex(5, 6) : DIM cpx2 AS CComplex(2, 3) : cpx2 = cpx1 / cpx2
' Example: DIM cpx1 AS CComplex(5, 6) : DIM cpx2 AS CComplex : cpx2 = cpx1 / 11
' Example: DIM cpx1 AS CComplex(5, 6) : DIM cpx2 AS CComplex : cpx2 = 11 / cpx1
' ========================================================================================
' ========================================================================================
' * Returns the quotient of complex numbers.
' - .NET 4.7 code:
' public static Complex operator /(Complex left, Complex right) {
'    // Division : Smith's formula.
'    double a = left.m_real;
'    double b = left.m_imaginary;
'    double c = right.m_real;
'    double d = right.m_imaginary;
'    if (Math.Abs(d) < Math.Abs(c)) {
'       double doc = d / c;
'       return new Complex((a + b * doc) / (c + d * doc), (b - a * doc) / (c + d * doc));
'    } else {
'       double cod = c / d;
'       return new Complex((b + a * cod) / (d + c * cod), (-a + b * cod) / (d + c * cod));
'    }
' }
' ========================================================================================
PRIVATE OPERATOR / (BYREF leftside AS CComplex, BYREF rightside AS CComplex) AS CComplex
   CCOMPLEX_DP("CCOMPLEX OPERATOR / (CComplex, CComplex)")
   ' // Division : Smith's formula.
   DIM a AS double = leftside.m_cpx.x
   DIM b AS DOUBLE = leftside.m_cpx.y
   DIM c AS DOUBLE = rightside.m_cpx.x
   DIM d AS DOUBLE = rightside.m_cpx.y
   IF ABS(d) < ABS(c) THEN
      DIM doc AS DOUBLE = d / c
      RETURN TYPE<_complex>((a + b * doc) / (c + d * doc), (b - a * doc) / (c + d * doc))
   ELSE
      DIM cod AS DOUBLE = c / d
      RETURN TYPE<_complex>((b + a * cod) / (d + c * cod), (-a + b * cod) / (d + c * cod))
   END IF
END OPERATOR
' ========================================================================================
'PRIVATE OPERATOR / (BYREF z1 AS CComplex, BYREF z2 AS CComplex) AS CComplex
'   DIM d AS DOUBLE = 1 / (z2.m_cpx.x * z2.m_cpx.x + z2.m_cpx.y * z2.m_cpx.y)
'   RETURN TYPE <_complex>((z1.m_cpx.x * z2.m_cpx.x + z1.m_cpx.y * z2.m_cpx.y) * d, _
'                          (z1.m_cpx.y * z2.m_cpx.x - z1.m_cpx.x * z2.m_cpx.y) * d)
'END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR / (BYVAL a AS DOUBLE, BYREF z AS CComplex) AS CComplex
   CCOMPLEX_DP("CCOMPLEX OPERATOR / (Double, CComplex)")
   DIM d AS DOUBLE = a / (z.m_cpx.x * z.m_cpx.x + z.m_cpx.y * z.m_cpx.y)
   RETURN TYPE<_complex>(z.m_cpx.x * d, -z.m_cpx.y * d)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR / (BYREF z AS CComplex, BYVAL a AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX OPERATOR / (CComplex, Double)")
   RETURN TYPE<_complex>(z.m_cpx.x / a, z.m_cpx.y / a)
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Adds a number.
' Example: DIM cpx AS CComplex = CComplex(3, 4)
' DIM cpx2 AS CComplex = CComplex(5, 6) : cpx += cpx2
' ========================================================================================
PRIVATE OPERATOR CComplex.+= (BYREF z AS CComplex)
   CCOMPLEX_DP("CCOMPLEX OPERATOR += (CComplex)")
   m_cpx = CComplex(m_cpx) + z
END OPERATOR
' ========================================================================================
' ========================================================================================
' * Subtracts a number.
' Example: DIM cpx AS CComplex = CComplex(3, 4)
' DIM cpx2 AS CComplex = CComplex(5, 6) : cpx2 -= cpx
' ========================================================================================
PRIVATE OPERATOR CComplex.-= (BYREF z AS CComplex)
   CCOMPLEX_DP("CCOMPLEX OPERATOR -= (CComplex)")
   m_cpx = CComplex(m_cpx) - z
END OPERATOR
' ========================================================================================
' ========================================================================================
' * Multiplies by a number.
' Example: DIM cpx AS CComplex = CComplex(3, 4)
' DIM cpx2 AS CComplex = CComplex(5, 6) : cpx *= cpx2
' ========================================================================================
PRIVATE OPERATOR CComplex.*= (BYREF z AS CComplex)
   CCOMPLEX_DP("CCOMPLEX OPERATOR *= (CComplex)")
   m_cpx = CComplex(m_cpx) * z
END OPERATOR
' ========================================================================================
' ========================================================================================
' * Divides by a number.
' Example: DIM cpx AS CComplex = CComplex(3, 4)
' DIM cpx2 AS CComplex = CComplex(5, 6) : cpx2 /= cpx
' ========================================================================================
PRIVATE OPERATOR CComplex./= (BYREF z AS CComplex)
   CCOMPLEX_DP("CCOMPLEX OPERATOR /= (CComplex)")
   m_cpx = CComplex(m_cpx) / z
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Adds a complex number.
' Example: DIM cpx AS CComplex = CComplex(5, 6)
' DIM cpx2 AS CComplex = CComplex(2, 3) : cpx = cpx.CAdd(cpx2)
' --or-- cpx = cpx.CAdd(CComplex(2, 3))
' ========================================================================================
PRIVATE FUNCTION CComplex.CAdd (BYREF z AS CComplex) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CAdd (CComplex)")
   RETURN TYPE<_complex>(m_cpx.x + z.m_cpx.x, m_cpx.y + z.m_cpx.y)
END FUNCTION
' ========================================================================================
' * Adds a real and an imaginary number.
' DIM cpx AS CComplex = CComplex(5, 6) : cpx = cpx.CAdd(2, 3)
' ========================================================================================
PRIVATE FUNCTION CComplex.CAdd (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CAdd (Double, Double)")
   RETURN TYPE<_complex>(m_cpx.x + x, m_cpx.y + y)
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Adds a real number.
' DIM cpx AS CComplex = CComplex(5, 6) : cpx = cpx.CAddReal(10)
' ========================================================================================
PRIVATE FUNCTION CComplex.CAddReal (BYVAL x AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CAddReal (Double)")
   RETURN TYPE<_complex>(m_cpx.x + x, m_cpx.y)
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Adds an imaginary number.
' DIM cpx AS CComplex = CComplex(5, 6) : cpx = cpx.CAddImag(10)
' ========================================================================================
PRIVATE FUNCTION CComplex.CAddImag (BYVAL y AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CAddImag (Double)")
   RETURN TYPE<_complex>(m_cpx.x, m_cpx.y + y)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Subtracts a complex number.
' Example: DIM cpx AS CComplex = CComplex(5, 6)
' DIM cpx2 AS CComplex = CComplex(2, 3) : cpx = cpx.CSub(cpx2)
' --or-- cpx = cpx.CSub(CComplex(2, 3))
' ========================================================================================
PRIVATE FUNCTION CComplex.CSub (BYREF z AS CComplex) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CSub (CComplex)")
   RETURN TYPE<_complex>(m_cpx.x - z.m_cpx.x, m_cpx.y - z.m_cpx.y)
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Subtracts a real number.
' Example: DIM cpx AS CComplex = CComplex(5, 6) : cpx = cpx.CSub(2, 3)
' ========================================================================================
PRIVATE FUNCTION CComplex.CSub (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CSub (Double, Double)")
   RETURN TYPE<_complex>(m_cpx.x - x, m_cpx.y - y)
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Subtracts a real number.
' Example: DIM cpx AS CComplex = CComplex(5, 6) : cpx = cpx.CSubReal(2)
' ========================================================================================
PRIVATE FUNCTION CComplex.CSubReal (BYVAL x AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CSubReal (Double)")
   RETURN TYPE<_complex>(m_cpx.x - x, m_cpx.y)
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Subtracts an imaginary number.
' Example: DIM cpx AS CComplex = CComplex(5, 6) : cx = cpx.CSubImag(3)
' ========================================================================================
PRIVATE FUNCTION CComplex.CSubImag (BYVAL y AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CSubImag (Double)")
   RETURN TYPE<_complex>(m_cpx.x, m_cpx.y - y)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Multiplies by a complex number.
' Example: DIM cpx AS CComplex = CComplex(5, 6)
' DIM cpx2 AS CComplex = CComplex(2, 3) : cpx = cpx.CMul(cpx2)
' --or-- cpx = cpx.CMul(CComplex(2, 3))
' ========================================================================================
PRIVATE FUNCTION CComplex.CMul (BYREF z AS CComplex) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CMul (CComplex)")
   RETURN TYPE<_complex>(m_cpx.x * z.m_cpx.x - m_cpx.y * z.m_cpx.y, m_cpx.x * z.m_cpx.y + m_cpx.y * z.m_cpx.x)
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Multiplies by a real and an imaginary number.
' Example: DIM cpx AS CComplex = CComplex(5, 6) : cpx = cpx.CMul(2, 3)
' ========================================================================================
PRIVATE FUNCTION CComplex.CMul (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CMul (Double, Double)")
   RETURN TYPE<_complex>(m_cpx.x * x - m_cpx.y * y, m_cpx.x * y + m_cpx.y * x)
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Multiplies by a real number.
' ========================================================================================
PRIVATE FUNCTION CComplex.CMulReal (BYVAL x AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CMulReal (Double)")
   RETURN TYPE<_complex>(m_cpx.x * x, m_cpx.y * x)
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Multiplies by an imaginary number.
' ========================================================================================
PRIVATE FUNCTION CComplex.CMulImag (BYVAL y AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CMulImag (Double)")
   RETURN TYPE<_complex>(-y * m_cpx.y, y * m_cpx.x)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Divides by a complex number.
' ========================================================================================
PRIVATE FUNCTION CComplex.CDiv (BYREF z AS CComplex) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CDiv (CComplex)")
   RETURN CComplex(m_cpx) / z
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Divides by a real and an imaginary number.
' ========================================================================================
PRIVATE FUNCTION CComplex.CDiv (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CDiv (Double, Double)")
   RETURN CComplex(m_cpx) / CComplex(x, y)
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Divides by a real number.
' ========================================================================================
PRIVATE FUNCTION CComplex.CDivReal (BYVAL x AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CDivReal (Double)")
   RETURN TYPE<_complex>(m_cpx.x / x, m_cpx.y / x)
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Divides by an imaginar number.
' ========================================================================================
PRIVATE FUNCTION CComplex.CDivImag (BYVAL y AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CDivImag (Double)")
   RETURN TYPE<_complex>(m_cpx.y / y, - m_cpx.x / y)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex conjugate of this complex number.
' Example:
'   DIM cpx AS CComplex = CComplex(2, 3)
'   cpx = cpx.CConj
'   PRINT cpx
' Output: 2 - 3 * i
' ========================================================================================
' ========================================================================================
' - .NET 4.7 code:
' public static Complex Conjugate(Complex value) {
'   // Conjugate of a Complex number: the conjugate of x+i*y is x-i*y
'   return (new Complex(value.m_real, (-value.m_imaginary)));
' }
' ========================================================================================
' Conjugates this complex number.
' ========================================================================================
PRIVATE FUNCTION CComplex.CConjugate () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CConjugate")
   RETURN TYPE<_complex>(m_cpx.x, -m_cpx.y)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CComplex.CConj () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CConj")
   RETURN TYPE<_complex>(m_cpx.x, -m_cpx.y)
END FUNCTION
' ========================================================================================

'/* Properties of complex numbers */

' =====================================================================================
' Returns the argument of this complex number.
' Examples:
'   DIM cpx AS CComplex = CComplex(1, 0)
'   PRINT cpx.CArg
' Output: 0.0
'   DIM cpx AS CComplex = CComplex(0, 1)
'   PRINT cpx.CArg
' Output: 1.570796326794897
'   DIM cpx AS CComplex = CComplex(0, -1)
'   PRINT cpx.CArg
' Output: -1.570796326794897
'   DIM cpx AS CComplex = CComplex(-1, 0)
'   PRINT cpx.CArg
' Output: 3.141592653589793
' =====================================================================================
PRIVATE FUNCTION CComplex.CArg () AS DOUBLE
   CCOMPLEX_DP("CCOMPLEX CArg")
   RETURN atan2(m_cpx.y, m_cpx.x)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CComplex.CArgument () AS DOUBLE
   CCOMPLEX_DP("CCOMPLEX CArgument")
   RETURN atan2(m_cpx.y, m_cpx.x)
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Gets the phase of this complex number (same as argument, above).
' - NET 4.7 code:
' public Double Phase {
'   get { return Math.Atan2(m_imaginary, m_real); }
' }
' =====================================================================================
PRIVATE FUNCTION CComplex.CPhase () AS DOUBLE
   CCOMPLEX_DP("CCOMPLEX CPhase")
   RETURN atan2(m_cpx.y, m_cpx.x)
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Returns the magnitude of this complex number.
' Example:
'   DIM cpx AS CComplex = CComplex(2, 3)
'   PRINT cpx.CAbs
' Output: 3.60555127546399
' =====================================================================================
' =====================================================================================
' - NET 4.7 code:
' public static Double Abs(Complex value) {
'    if(Double.IsInfinity(value.m_real) || Double.IsInfinity(value.m_imaginary)) {
'       return double.PositiveInfinity;
'    }
'    // |value| == sqrt(a^2 + b^2)
'    // sqrt(a^2 + b^2) == a/a * sqrt(a^2 + b^2) = a * sqrt(a^2/a^2 + b^2/a^2)
'    // Using the above we can factor out the square of the larger component to dodge overflow.
'    double c = Math.Abs(value.m_real);
'    double d = Math.Abs(value.m_imaginary);
'    if (c > d) {
'       double r = d / c;
'       return c * Math.Sqrt(1.0 + r * r);
'    } else if (d == 0.0) {
'       return c;  // c is either 0.0 or NaN
'    } else {
'       double r = c / d;
'       return d * Math.Sqrt(1.0 + r * r);
'    }
' }
' =====================================================================================
PRIVATE FUNCTION CComplex.CAbs () AS DOUBLE
   CCOMPLEX_DP("CCOMPLEX CAbs")
   IF this.IsInfinity(m_cpx.x) OR this.IsInfinity(m_cpx.y) THEN RETURN HUGE_VALF
   ' // |value| == sqrt(a^2 + b^2)
   ' // sqrt(a^2 + b^2) == a/a * sqrt(a^2 + b^2) = a * sqrt(a^2/a^2 + b^2/a^2)
   ' // Using the above we can factor out the square of the larger component to dodge overflow.
   DIM c AS DOUBLE = ABS(m_cpx.x)
   DIM d AS DOUBLE = ABS(m_cpx.y)
   IF c > d THEN
      DIM r AS DOUBLE = d / c
      RETURN c * SQR(1.0 + r * r)
   ELSEIF d = 0 THEN
      RETURN c   ' // c is either 0.0 or NaN
   ELSE
      DIM r AS DOUBLE = c / d
      RETURN d * SQR(1.0 + r * r)
   END IF
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CComplex.CMagnitude () AS DOUBLE
   CCOMPLEX_DP("CCOMPLEX CMagnitude")
   IF this.IsInfinity(m_cpx.x) OR this.IsInfinity(m_cpx.y) THEN RETURN HUGE_VALF
   ' // |value| == sqrt(a^2 + b^2)
   ' // sqrt(a^2 + b^2) == a/a * sqrt(a^2 + b^2) = a * sqrt(a^2/a^2 + b^2/a^2)
   ' // Using the above we can factor out the square of the larger component to dodge overflow.
   DIM c AS DOUBLE = ABS(m_cpx.x)
   DIM d AS DOUBLE = ABS(m_cpx.y)
   IF c > d THEN
      DIM r AS DOUBLE = d / c
      RETURN c * SQR(1.0 + r * r)
   ELSEIF d = 0 THEN
      RETURN c   ' // c is either 0.0 or NaN
   ELSE
      DIM r AS DOUBLE = c / d
      RETURN d * SQR(1.0 + r * r)
   END IF
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Returns the squared magnitude of this complex number, otherwise known as the complex norm.
' Examples:
'   DIM cpx AS CComplex = CComplex(2, 3)
'   PRINT cpx.CAbs2
' Output: 13
' --or--
'   DIM cpx AS CComplex = CComplex(2, 3)
'   PRINT cpx.CNorm
' Output: 13
' =====================================================================================
PRIVATE FUNCTION CComplex.CAbs2 () AS DOUBLE
   CCOMPLEX_DP("CCOMPLEX CAbs2")
   FUNCTION = (m_cpx.x * m_cpx.x) + (m_cpx.y * m_cpx.y)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CComplex.CNorm () AS DOUBLE
   CCOMPLEX_DP("CCOMPLEX CNorm")
   FUNCTION = (m_cpx.x * m_cpx.x) + (m_cpx.y * m_cpx.y)
END FUNCTION
' =====================================================================================

' ========================================================================================
' * Returns the natural logarithm of the magnitude of the complex number z, log|z|.
' It allows an accurate evaluation of \log|z| when |z| is close to one. The direct
' evaluation of log(CAbs(z)) would lead to a loss of precision in this case.
' Example:
'   DIM cpx AS CComplex = CComplex(1.1, 0.1)
'   PRINT cpx.CLogAbs
' Output: 0.09942542937258279
' ========================================================================================
PRIVATE FUNCTION CComplex.CLogAbs () AS DOUBLE
   CCOMPLEX_DP("CCOMPLEX CLogAbs")
   RETURN LOG(this.CAbs)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the inverse, or reciprocal, of this complex number.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   Print cpx.CReciprocal
' Output: 0.5 -0.5 * i
' ========================================================================================
' ========================================================================================
' .NET 4.7: System.Numerics/System/Numerics/Complex.cs
' public static Complex Reciprocal(Complex value) {
'    // Reciprocal of a Complex number : the reciprocal of x+i*y is 1/(x+i*y)
'    if ((value.m_real == 0) && (value.m_imaginary == 0)) {
'       return Complex.Zero;
'    }
'       return Complex.One / value;
' }
' ========================================================================================
PRIVATE FUNCTION CComplex.CReciprocal () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CReciprocal")
   IF m_cpx.x = 0 AND m_cpx.y = 0 THEN RETURN CComplex(0, 0)
   RETURN CComplex(1, 0) / CComplex(m_cpx)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CComplex.CInverse () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CInverse")
   IF m_cpx.x = 0 AND m_cpx.y = 0 THEN RETURN CComplex(0, 0)
   RETURN CComplex(1, 0) / CComplex(m_cpx)
END FUNCTION
' ========================================================================================

'/* Elementary Complex Functions */

' ========================================================================================
' * Returns the square root of the complex number z. The branch cut is the negative
' real axis. The result always lies in the right half of the complex plane.
' Example:
'   DIM cpx AS CComplex = CComplex(2, 3)
'   PRINT cpx.CSqrt
' Output: 1.67414922803554 +0.895977476129838 * i
' Compute the square root of -1:
'   DIM cpx AS CComplex = CComplex(-1)
'   PRINT cpx.CSqr
' Output: 0 +1.0 * i
' ========================================================================================
' ========================================================================================
' // Make a complex number from polar coordinates
' - NET 4.7 code:
'public static Complex FromPolarCoordinates(Double magnitude, Double phase) /* Factory method to take polar inputs and create a Complex object */
'   { return new Complex((magnitude * Math.Cos(phase)), (magnitude * Math.Sin(phase))); }
' ========================================================================================
PRIVATE FUNCTION CComplex.CFromPolarCoordinates (BYVAL magnitude AS DOUBLE, BYVAL phase AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CFromPolarCoordinates")
   RETURN TYPE<_complex>(magnitude * COS(phase), magnitude * SIN(phase))
END FUNCTION
' ========================================================================================
' ========================================================================================
' - NET 4.7 code:
'public static Complex Sqrt(Complex value) /* Square root ot the complex number */
'{ return Complex.FromPolarCoordinates(Math.Sqrt(value.Magnitude), value.Phase / 2.0); }
' ========================================================================================
PRIVATE FUNCTION CComplex.CSqr () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CSQr")
   DIM z AS CComplex = this.CFromPolarCoordinates(SQR(this.CMagnitude), this.CPhase / 2.0)
   IF m_cpx.x < 0 AND m_cpx.y = 0 THEN z.m_cpx.x = 0   ' // For negative real numbers
   RETURN z
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CComplex.CSqrt () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CSQrt")
   DIM z AS CComplex = this.CFromPolarCoordinates(SQR(this.CMagnitude), this.CPhase / 2.0)
   IF m_cpx.x < 0 AND m_cpx.y = 0 THEN z.m_cpx.x = 0   ' // For negative real numbers
   RETURN z
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex square root of the real number x, where x may be negative.
' ========================================================================================
PRIVATE FUNCTION CComplex.CSqr (BYVAL value AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CSQr (Double)")
   IF value >= 0 THEN RETURN TYPE<_complex>(SQR(value), 0)
   RETURN TYPE<_complex>(0, SQR(-value))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CComplex.CSqrt (BYVAL value AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CSQrt (Double)")
   IF value >= 0 THEN RETURN TYPE<_complex>(SQR(value), 0)
   RETURN TYPE<_complex>(0, SQR(-value))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex exponential of this complex number.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   PRINT cpx.CExp
' Output: 1.468693939915885 +2.287355287178842 * i
' ========================================================================================
' ========================================================================================
' - NET 4.7 code:
'public static Complex Exp(Complex value) /* The complex number raised to e */
'{
'   Double temp_factor = Math.Exp(value.m_real);
'   Double result_re = temp_factor * Math.Cos(value.m_imaginary);
'   Double result_im = temp_factor * Math.Sin(value.m_imaginary);
'   return (new Complex(result_re, result_im));
'}
' ========================================================================================
PRIVATE FUNCTION CComplex.CExp () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CExp")
   DIM factor AS DOUBLE = EXP(m_cpx.x)
   RETURN TYPE<_complex>(factor * COS(m_cpx.y), factor * SIN(m_cpx.y))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex natural logarithm (base e) of this complex number.
' The branch cut is the negative real axis.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   PRINT cpx.CLog
' Output: 0.3465735902799727 +0.7853981633974483 * i
'   DIM cpx AS CComplex = CComplex(0, 0)
'   PRINT cpx.CLog
' Output: -1.#INF
' ========================================================================================
' ========================================================================================
' .NET 4.7:
' public static Complex Log(Complex value) /* Log of the complex number value to the base of 'e' */
' { return (new Complex((Math.Log(Abs(value))), (Math.Atan2(value.m_imaginary, value.m_real)))); }
' ========================================================================================
PRIVATE FUNCTION CComplex.CLog () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CLog")
   RETURN TYPE<_complex>(LOG(this.CAbs), atan2(m_cpx.y, m_cpx.x))
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Returns the complex base-value logarithm of this complex number.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   PRINT cpx.CLog(10)
' Output: 0.1505149978319906 +0.3410940884604603 * i
' ========================================================================================
' ========================================================================================
' - NET 4.7:
' public static Complex Log(Complex value, Double baseValue) /* Log of the complex number to a the base of a double */
' { return (Log(value) / Log(baseValue)); }
' ========================================================================================
PRIVATE FUNCTION CComplex.CLog (BYVAL baseValue AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CLog (Double)")
   DIM z AS CComplex = this.CLog
   z.m_cpx.x /= LOG(baseValue)
   z.m_cpx.y /= LOG(baseValue)
   RETURN z
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex base-10 logarithm of this complex number.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   PRINT cpx.CLog10
' Output: 0.1505149978319906 +0.3410940884604603 * i
' ========================================================================================
' ========================================================================================
' - NET 4.7 code:
' private static Complex Scale(Complex value, Double factor) {
'    Double result_re = factor * value.m_real;
'    Double result_im = factor * value.m_imaginary;
'    return (new Complex(result_re, result_im));
' }
' public static Complex Log10(Complex value) /* Log to the base of 10 of the complex number */
' {
'    Complex temp_log = Log(value);
'    return (Scale(temp_log, (Double)LOG_10_INV));
' }
' ========================================================================================
PRIVATE FUNCTION CComplex.CLog10 () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CLog10")
   DIM z AS _complex = this.CLog
   z.x = z.x * M_LOG10E
   z.y = z.y * M_LOG10E
   RETURN z
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex number a raised to the complex power b.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   DIM b AS CComplex = CComplex(2, 2)
'   PRINT cpx.CPow(b)
' Output: -0.2656539988492412 +0.3198181138561362 * i
' ========================================================================================
' ========================================================================================
' - NET 4.7 code:
' public static Complex Pow(Complex value, Complex power) /* A complex number raised to another complex number */
' {
'    if (power == Complex.Zero) { return Complex.One; }
'    if (value == Complex.Zero) { return Complex.Zero; }
'    double a = value.m_real;
'    double b = value.m_imaginary;
'    double c = power.m_real;
'    double d = power.m_imaginary;
'    double rho = Complex.Abs(value);
'    double theta = Math.Atan2(b, a);
'    double newRho = c * theta + d * Math.Log(rho);
'    double t = Math.Pow(rho, c) * Math.Pow(Math.E, -d * theta);
'    return new Complex(t * Math.Cos(newRho), t * Math.Sin(newRho));
' }
' ========================================================================================
PRIVATE FUNCTION CComplex.CPow (BYREF power AS CComplex) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CPow (CComplex)")
   IF m_cpx.x = 0 AND m_cpx.y = 0 THEN RETURN TYPE<_complex>(0, 0)
   IF power.m_cpx.x = 0 AND power.m_cpx.y = 0 THEN RETURN TYPE<_complex>(1, 0)
   DIM rho AS DOUBLE = this.CAbs
   DIM theta AS DOUBLE = atan2(m_cpx.y, m_cpx.x)
   DIM newRho AS DOUBLE = power.m_cpx.x * theta + power.m_cpx.y * LOG(rho)
   DIM t AS DOUBLE = POW(rho, power.m_cpx.x) * POW(M_E, -power.m_cpx.y * theta)
   RETURN TYPE<_complex>(t * COS(newRho), t * SIN(newRho))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR ^ (BYREF value AS CComplex, BYREF power AS CComplex) AS CComplex
   CCOMPLEX_DP("CCOMPLEX ^ (CComplex, CComplex)")
   OPERATOR = value.CPow(power)
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Returns this complex number raised to a real number.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   PRINT cpx.CPow(2)
' Output: 1.224606353822378e-016 +2 * i
' ========================================================================================
' ========================================================================================
' - NET 4.7 code:
'public static Complex Pow(Complex value, Double power) // A complex number raised to a real number
'{ return Pow(value, new Complex(power, 0)); }
' ========================================================================================
PRIVATE FUNCTION CComplex.CPow (BYVAL power AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CPow (Double)")
   RETURN CComplex(m_cpx) ^ CComplex(power, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR ^ (BYREF value AS CComplex, BYVAL power AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX ^ (CComplex, Double)")
   OPERATOR = value.CPow(power)
END OPERATOR
' ========================================================================================

'/* Complex Trigonometric Functions */

' ========================================================================================
' * Returns the complex sine of this complex number.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   PRINT cpx.CSin
' Output: 1.298457581415977 +0.6349639147847361 * i
' ========================================================================================
PRIVATE FUNCTION CComplex.CSin () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CSin")
   IF m_cpx.y = 0 THEN RETURN TYPE<_complex>(SIN(m_cpx.x), 0)
   RETURN TYPE<_complex>(SIN(m_cpx.x) * cosh(m_cpx.y), COS(m_cpx.x) * sinh(m_cpx.y))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex cosine of this complex number.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   PRINT cpx.CCos
' Output: 0.8337300251311491 -0.9888977057628651 * i
' ========================================================================================
PRIVATE FUNCTION CComplex.CCos () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CCos")
   IF m_cpx.y = 0 THEN RETURN TYPE<_complex>(COS(m_cpx.x), 0)
   RETURN TYPE<_complex>(COS(m_cpx.x) * cosh(m_cpx.y), SIN(m_cpx.x) * sinh(-m_cpx.y))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex secant of this complex number.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   PRINT cpx.CSec
' Output: 0.4983370305551869 +0.591083841721045 * i
' ========================================================================================
PRIVATE FUNCTION CComplex.CSec () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CSec")
   DIM z AS CComplex = this.CCos
   RETURN z.CReciprocal
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex cosecant of this complex number.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   PRINT cpx.CCsc
' Output: 0.6215180171704285 -0.3039310016284265 * i
' ========================================================================================
PRIVATE FUNCTION CComplex.CCsc () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CCsc")
   DIM z AS CComplex = this.CSin
   RETURN z.CReciprocal
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex tangent of this complex number.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   PRINT cpx.CTan
' Output: 0.2717525853195117 +1.083923327338695 * i
' ========================================================================================
' - NET 4.7 code:
' public static Complex Tan(Complex value) { return (Sin(value) / Cos(value)); }
' ========================================================================================
PRIVATE FUNCTION CComplex.CTan () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CTan")
   DIM z AS CComplex = this.CSin
   RETURN z / this.CCos
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the complex cotangent of this complex number.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   PRINT cpx.CCot
' Output: 0.2176215618544027 -0.8680141428959249 * i
' ========================================================================================
PRIVATE FUNCTION CComplex.CCot () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CCot")
   DIM z AS CComplex = this.CTan
   RETURN z.CReciprocal
END FUNCTION
' ========================================================================================

'/* Inverse Complex Trigonometric Functions */

' ========================================================================================
' Returns the complex arcsine of a real number.
' For a between -1 and 1, the function returns a real value in the range [-pi/2, pi/2].
' For a less than -1 the result has a real part of -pi/2 and a positive imaginary part.
' For a greater than 1 the result has a real part of pi/2 and a negative imaginary part.
' Example:
'   DIM cpx AS CComplex
'   PRINT cpx.CArcSinReal(1)
' Output: 1.570796326794897 +0 * i
' ========================================================================================
PRIVATE FUNCTION CComplex.CArcSinReal (BYVAL value AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CArcSinReal")
   IF ABS(value) <= 1 THEN RETURN TYPE<_complex>(ASIN(value))
   IF value < 0 THEN RETURN TYPE<_complex>(-M_PI_2, this.ArcCosH(-value))
   RETURN TYPE<_complex>(M_PI_2, this.ArcCosH(value))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CComplex.CASinReal (BYVAL value AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CASinReal")
   IF ABS(value) <= 1 THEN RETURN TYPE<_complex>(ASIN(value))
   IF value < 0 THEN RETURN TYPE<_complex>(-M_PI_2, this.ArcCosH(-value))
   RETURN TYPE<_complex>(M_PI_2, this.ArcCosH(value))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex arcsine of this complex number.
' The branch cuts are on the real axis, less than -1 and greater than 1.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   PRINT cpx.CArcSin
' Output: 0.6662394324925152 +1.061275061905036 * i
' ========================================================================================
PRIVATE FUNCTION CComplex.CArcSin () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CArcSin")
   RETURN Afx.CArcSin(m_cpx)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CComplex.CASin () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CAsin")
   RETURN Afx.CArcSin(m_cpx)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the complex arccosine of a real number.
' For a between -1 and 1, the function returns a real value in the range [0, pi].
' For a less than -1 the result has a real part of pi and a negative imaginary part.
' For a greater than 1 the result is purely imaginary and positive.
' Examples:
' DIM cpx AS CComplex
' print cpx.CArcCosReal(1) ' = 0 0 * i
' print cpx.CArcCosReal(-1) ' = 3.141592653589793 0 * i
' print cpx.CArcCosReal(2) ' = 0 +1.316957896924817 * i
' ========================================================================================
PRIVATE FUNCTION CComplex.CArcCosReal (BYVAL value AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CArcCosReal")
   IF ABS(value) <= 1.0 THEN RETURN TYPE<_complex>(acos(value))
   IF value < 0 THEN RETURN TYPE<_complex>(M_PI, -this.ArcCosH(-value))
   RETURN TYPE<_complex>(0, this.ArcCosH(value))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CComplex.CACosReal (BYVAL value AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CACosReal")
   IF ABS(value) <= 1.0 THEN RETURN TYPE<_complex>(acos(value))
   IF value < 0 THEN RETURN TYPE<_complex>(M_PI, -this.ArcCosH(-value))
   RETURN TYPE<_complex>(0, this.ArcCosH(value))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex arccosine of this complex number.
' The branch cuts are on the real axis, less than -1 and greater than 1.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   print cpx.CArcCos(z)
' Output: 0.9045568943023814 -1.061275061905036 * i
' ========================================================================================
' ========================================================================================
' - NET 4.7 code:
' public static Complex Acos(Complex value) /* Arccos */
' { return (-ImaginaryOne) * Log(value + ImaginaryOne*Sqrt(One - (value * value))); }
' ========================================================================================
PRIVATE FUNCTION CComplex.CArcCos () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CArcCos")
   RETURN Afx.CArcCos(m_cpx)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CComplex.CACos () AS CComplex
      CCOMPLEX_DP("CCOMPLEX CACos")
RETURN Afx.CArcCos(m_cpx)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the complex arcsecant of the a number.
' Example
'   DIM cpx AS CComplex
'   print cpx.CArcSecReal(1.1)
' Output: 0.4296996661514246 0 * i
' ========================================================================================
PRIVATE FUNCTION CComplex.CArcSecReal (BYVAL value AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CArcSecReal")
   IF value <= -1 OR value >= 1 THEN RETURN TYPE<_complex>(acos(1 / value), 0)
   IF value >= 0 THEN RETURN TYPE<_complex>(0, this.ArcCosH(1 / value))
   RETURN TYPE<_complex>(M_PI, -this.ArcCosH(-1 / value))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CComplex.CASecReal (BYVAL value AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CASecReal")
   IF value <= -1 OR value >= 1 THEN RETURN TYPE<_complex>(acos(1 / value), 0)
   IF value >= 0 THEN RETURN TYPE<_complex>(0, this.ArcCosH(1 / value))
   RETURN TYPE<_complex>(M_PI, -this.ArcCosH(-1 / value))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex arcsecant of this complex number.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   print cpx.CArcSec
' Output: 1.118517879643706 +0.5306375309525176 * i
' ========================================================================================
PRIVATE FUNCTION CComplex.CArcSec () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CArcSec")
   DIM z AS CComplex = this.CReciprocal
   RETURN z.CArcCos
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CComplex.CASec () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CSec")
   DIM z AS CComplex = this.CReciprocal
   RETURN z.CArcCos
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the complex arccosecant of a real number.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   print cpx.CArcCscReal(1)
' Output: 1.570796326794897 0 * i
' ========================================================================================
PRIVATE FUNCTION CComplex.CArcCscReal (BYVAL value AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CArcCscReal")
   IF value <= -1 OR value >= 1 THEN RETURN TYPE<_complex>(asin(1 / value))
   IF value >= 0 THEN RETURN TYPE<_complex>(M_PI_2, -ArcCosH(1 / value))
   RETURN TYPE<_complex>(-M_PI_2, ArcCosH(-1 / value))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CComplex.CACscReal (BYVAL value AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CACscReal")
   IF value <= -1 OR value >= 1 THEN RETURN TYPE<_complex>(asin(1 / value))
   IF value >= 0 THEN RETURN TYPE<_complex>(M_PI_2, -ArcCosH(1 / value))
   RETURN TYPE<_complex>(-M_PI_2, ArcCosH(-1 / value))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex arccosecant of this complex number.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   print cpx.CArcCsc
' Output: 0.4522784471511907 -0.5306375309525178 * i
' ========================================================================================
PRIVATE FUNCTION CComplex.CArcCsc () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CArcCsc")
   DIM z AS CComplex = this.CReciprocal
   RETURN z.CArcSin
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CComplex.CACsc () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CACsc")
   DIM z AS CComplex = this.CReciprocal
   RETURN z.CArcSin
END FUNCTION
' ========================================================================================

' ========================================================================================
' * This function returns the complex arctangent of this complex number.
' The branch cuts are on the imaginary axis, below -i and above i.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   print cpx.CArcTan
' Output: 1.017221967897851 +0.4023594781085251 * i
' ========================================================================================
' ========================================================================================
' - NET 4.7 code:
' public static Complex Atan(Complex value) /* Arctan */
' {
'    Complex Two = new Complex(2.0, 0.0);
'    return (ImaginaryOne / Two) * (Log(One - ImaginaryOne * value) - Log(One + ImaginaryOne * value));
' }
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CComplex.CArcTan () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CArcTan")
   RETURN Afx.CArcTan(m_cpx)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CComplex.CATan () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CATan")
   RETURN Afx.CArcTan(m_cpx)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex arccotangent of this complex number.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   print cpx.CArcCot
' Output: 0.5535743588970452 -0.4023594781085251 * i
' ========================================================================================
PRIVATE FUNCTION CComplex.CArcCot () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CArcCot")
   IF m_cpx.x = 0 AND m_cpx.y = 0 THEN RETURN TYPE<_complex>(M_PI_2)
   DIM z AS CComplex = this.CReciprocal
   RETURN z.CArcTan
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CComplex.CACot () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CACot")
   IF m_cpx.x = 0 AND m_cpx.y = 0 THEN RETURN TYPE<_complex>(M_PI_2)
   DIM z AS CComplex = this.CReciprocal
   RETURN z.CArcTan
END FUNCTION
' ========================================================================================

'/* Complex Hyperbolic Functions */

' ========================================================================================
' * Returns the complex hyperbolic sine of this complex number.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   print cpx.CSinH
' Output: 0.6349639147847361 +1.298457581415977 * i
' ========================================================================================
PRIVATE FUNCTION CComplex.CSinH () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CSinH")
   RETURN TYPE<_complex>(sinh(m_cpx.x) * cos(m_cpx.y), cosh(m_cpx.x) * sin(m_cpx.y))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex hyperbolic cosine of this complex number.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   print cpx.CCosH
' Output: 0.8337300251311491 +0.9888977057628651 * i
' ========================================================================================
PRIVATE FUNCTION CComplex.CCosH () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CCosH")
   RETURN TYPE<_complex>(cosh(m_cpx.x) * cos(m_cpx.y), sinh(m_cpx.x) * sin(m_cpx.y))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex hyperbolic secant of this complex number.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   print cpx.CSecH
' Output: 0.4983370305551869 -0.591083841721045 * i
' ========================================================================================
PRIVATE FUNCTION CComplex.CSecH () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CSecH")
   DIM z AS CComplex = this.CCosH
   RETURN z.CReciprocal
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex hyperbolic cosecant of this complex number.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   print cpx.CCscH
' Output: 0.3039310016284265 -0.6215180171704285 * i
' ========================================================================================
PRIVATE FUNCTION CComplex.CCscH () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CCScH")
   DIM z AS CComplex = this.CSinH
   RETURN z.CReciprocal
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex hyperbolic tangent of this complex number.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   print cpx.CTanH
' Output: 1.083923327338695 +0.2717525853195119 * i
' ========================================================================================
' ========================================================================================
' - NET 4.7 code:
' public static Complex Tanh(Complex value) /* Hyperbolic tan */
' { return (Sinh(value) / Cosh(value)); }
' ========================================================================================
PRIVATE FUNCTION CComplex.CTanH () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CTanH")
   RETURN this.CSinH / this.CCosH
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex hyperbolic cotangent of this complex number.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   print cpx.CCotH
' Output: 0.8680141428959249 -0.2176215618544029 * i
' ========================================================================================
PRIVATE FUNCTION CComplex.CCotH () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CCotH")
   DIM z AS CComplex = this.CTanH
   RETURN z.CReciprocal
END FUNCTION
' ========================================================================================

'/* Inverse Complex Hyperbolic Functions */

' ========================================================================================
' * Returns the complex hyperbolic arcsine of this complex number.
' The branch cuts are on the imaginary axis, below -i and above i.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   print cpx.CArcSinH
' Output: 1.061275061905036 +0.6662394324925153 * i
' ========================================================================================
PRIVATE FUNCTION CComplex.CArcSinH () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CArcSinH")
   DIM z AS CComplex = this.CMulImag(1.0)
   z = z.CArcSin
   RETURN z.CMulImag(-1.0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CComplex.CASinH () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CASinH")
   DIM z AS CComplex = this.CMulImag(1.0)
   z = z.CArcSin
   RETURN z.CMulImag(-1.0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the complex hyperbolic arccosine of a real number.
' ========================================================================================
PRIVATE FUNCTION CComplex.CArcCosHReal (BYVAL value AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CArcCosReal")
   IF value >= 1 THEN RETURN TYPE<_complex>(this.ArcCosH(value), 0)
   IF value >= -1 THEN RETURN TYPE<_complex>(0, acos(value))
   RETURN TYPE<_complex>(this.ArcCosH(-value), M_PI)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CComplex.CACosHReal (BYVAL value AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CACosReal")
   IF value >= 1 THEN RETURN TYPE<_complex>(this.ArcCosH(value), 0)
   IF value >= -1 THEN RETURN TYPE<_complex>(0, acos(value))
   RETURN TYPE<_complex>(this.ArcCosH(-value), M_PI)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex hyperbolic arccosine of this complex number.
' The branch cut is on the real axis, less than 1.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   print cpx.CArcCosH
' Output: 1.061275061905036 +0.9045568943023813 * i
' ========================================================================================
PRIVATE FUNCTION CComplex.CArcCosH () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CArcCosH")
   DIM z AS CComplex = this.CArcCos
   RETURN z.CMulImag(IIF(z.m_cpx.y > 0, -1.0, 1.0))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CComplex.CACosH () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CACosH")
   DIM z AS CComplex = this.CArcCos
   RETURN z.CMulImag(IIF(z.m_cpx.y > 0, -1.0, 1.0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex hyperbolic arcsecant of this complex number.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   print cpx.CArcSecH
' Output: 0.5306375309525178 -1.118517879643706 * i
' ========================================================================================
PRIVATE FUNCTION CComplex.CArcSecH () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CArcSecH")
   DIM z AS CComplex = this.CReciprocal
   RETURN z.CArcCosH
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CComplex.CASecH () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CASecH")
   DIM z AS CComplex = this.CReciprocal
   RETURN z.CArcCosH
END FUNCTION
' ========================================================================================

' ========================================================================================
' * This function returns the complex hyperbolic arccosecant of this complex number.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   print cpx.CArcCscH
' Output: 0.5306375309525179 -0.4522784471511906 * i
' ========================================================================================
PRIVATE FUNCTION CComplex.CArcCscH () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CArcCscH")
   DIM z AS CComplex = this.CReciprocal
   RETURN z.CArcSinH
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CComplex.CACscH () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CACscH")
   DIM z AS CComplex = this.CReciprocal
   RETURN z.CArcSinH
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the complex hyperbolic arctangent of a real number.
' ========================================================================================
PRIVATE FUNCTION CComplex.CArcTanHReal (BYVAL value AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CArcTanHeRal")
   IF value > -1 AND value < 1 THEN RETURN TYPE<_complex>(this.ArcTanH(value), 0)
   RETURN TYPE<_complex>(ArcTanH(1 / value), IIF(value < 0, M_PI_2, -M_PI_2))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CComplex.CATanHReal (BYVAL value AS DOUBLE) AS CComplex
   CCOMPLEX_DP("CCOMPLEX CATanHReal")
   IF value > -1 AND value < 1 THEN RETURN TYPE<_complex>(this.ArcTanH(value), 0)
   RETURN TYPE<_complex>(ArcTanH(1 / value), IIF(value < 0, M_PI_2, -M_PI_2))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex hyperbolic arctangent of this complex number.
' The branch cuts are on the real axis, less than -1 and greater than 1.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   print cpx.CArcTanH
' Output: 0.4023594781085251 +1.017221967897851 * i
' ========================================================================================
PRIVATE FUNCTION CComplex.CArcTanH () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CArcTanH")
   IF m_cpx.y = 0 THEN RETURN this.CArcTanHReal(m_cpx.x)
   DIM z AS CComplex = this.CMulImag(1)
   z = z.CArcTan
   RETURN z.CMulImag(-1)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CComplex.CATanH () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CATanH")
   IF m_cpx.y = 0 THEN RETURN this.CArcTanHReal(m_cpx.x)
   DIM z AS CComplex = this.CMulImag(1)
   z = z.CArcTan
   RETURN z.CMulImag(-1)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex hyperbolic arccotangent of this complex number.
' Example:
'   DIM cpx AS CComplex = CComplex(1, 1)
'   print cpx.CArcCotH
' Output: 0.4023594781085251 -0.5535743588970452 * i
' ========================================================================================
PRIVATE FUNCTION CComplex.CArcCotH () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CArcCotH")
   DIM z AS CComplex = this.CReciprocal
   RETURN z.CArcTanH
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CComplex.CACotH () AS CComplex
   CCOMPLEX_DP("CCOMPLEX CACotH")
   DIM z AS CComplex = this.CReciprocal
   RETURN z.CArcTanH
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the sign of this complex number.
' If number is greater than zero, then CSgn returns 1.
' If number is equal to zero, then CSgn returns 0.
' If number is less than zero, then CSgn returns -1.
' ========================================================================================
PRIVATE FUNCTION CComplex.CSgn () AS LONG
   CCOMPLEX_DP("CCOMPLEX CSgn")
   IF m_cpx.x > 0 THEN RETURN 1
   IF m_cpx.x < 0 THEN RETURN -1
   IF m_cpx.y > 0 THEN RETURN 1
   IF m_cpx.y < 0 THEN RETURN -1
   RETURN 0
END FUNCTION
' ========================================================================================

END NAMESPACE
