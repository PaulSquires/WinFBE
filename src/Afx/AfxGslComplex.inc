' ########################################################################################
' Microsoft Windows
' File: AfxGslComplex.inc
' Contents: Complex numbers.
' Compiler: FreeBasic 32 & 64-bit
' Note: Because many of the functions are based in ansi C functions found in the GNU
' Scientific Library ( http://www.gnu.org/software/gsl/  ), distributed under the terms of
' the GNU General Public License (GPL), this module is also subject to that license and,
' therefore, is only suitable to be used in open source projects.
' GNU General Public License: http://www.gnu.org/copyleft/gpl.html
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "crt/math.bi"
#include once "crt/limits.bi"

NAMESPACE Gsl

' ########################################################################################
' Additional math functions.
' ########################################################################################

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
' A function returns this value when the result of a mathematical operation yields a value
' that is too large in magnitude to be representable with its return type. This is one of
' the possible range errors, and is signaled by setting errno to ERANGE.
' Actually, functions may either return a positive or a negative HUGE_VAL (HUGE_VAL or -HUGE_VAL)
' to indicate the sign of the result.
' Returns: 1.#INF
' Note: Use HUGE_VALF, that is defined as #define HUGE_VALF (1.0F/0.0F) and returns the same value.
' ========================================================================================
PRIVATE FUNCTION HUGE_VAL () AS DOUBLE
   DIM hLib AS HMODULE, phuge AS DOUBLE PTR
   hLib = LoadLibraryW("msvcrt.dll")
   IF hLib THEN
      phuge = cast(DOUBLE PTR, GetProcAddress(hLib, "_HUGE"))
      IF phuge THEN FUNCTION = *phuge
      FreeLibrary hLib
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Computes the value of exp(x)-1 in a way that is accurate for small x.
' The true value of exp(1e-10) - 1 is 1.00000000005e-010 to about 32 significant digits.
' This example shows the superiority of expm1 in this case.
' expm1(1e-10)
'   1.00000000005e-010
' exp(1e-10) - 1
'   1.000000082740371e-010
' ========================================================================================
PRIVATE FUNCTION expm1 (BYVAL x AS DOUBLE) AS DOUBLE
   DIM AS DOUBLE i, sum, term
   IF fabs(x) < M_LN2 THEN
      ' /* Compute the Taylor series S = x + (1/2!) x^2 + (1/3!) x^3 + ... */
      i = 1
      sum = x
      term = x / 1
      DO
         i += 1
         term *= x / i
         sum += term
         IF fabs(term) <= fabs(sum) * DBL_EPSILON THEN EXIT DO
      LOOP
      FUNCTION = sum
   ELSE
      FUNCTION = exp(x) - 1
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Calculates the inverse hyperbolic cosine.
' Example:
'   DIM AS double pi = 3.1415926535
'   DIM AS double x, y
'   x = cosh(pi / 4)
'   y = acosh(x)
'   print "cosh = ", pi/4, x
'   print "acosh = ", x, y
' Output:
'   cosh =         0.785398163375              1.324609089232506
'   acosh =        1.324609089232506           0.7853981633749999
' ========================================================================================
PRIVATE FUNCTION acosh (BYVAL x AS DOUBLE) AS DOUBLE
   DIM t AS DOUBLE
   IF x > 1.0 / SQRT_DBL_EPSILON THEN
      FUNCTION = log(x) + M_LN2
   ELSEIF x > 2 THEN
      FUNCTION = log(2 * x - 1 / (sqr(x * x - 1) + x))
   ELSEIF x > 1 THEN
      t = x - 1
      FUNCTION = log1p(t + sqr(2 * t + t * t))
   ELSEIF x = 1 THEN
      FUNCTION = 0
   ELSE
      FUNCTION = 0 / 0  ' NAN
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Calculates the inverse hyperbolic sine.
' Example:
'   DIM AS double pi = 3.1415926535
'   DIM AS double x, y
'   x = sinh(pi / 4)
'   y = asinh(x)
'   print "sinh = ", pi/4, x
'   print "asinh = ", x, y
' Output:
'   sinh =         0.785398163375              0.8686709614562744
'   asinh =        0.8686709614562744          0.7853981633750001
' ========================================================================================
PRIVATE FUNCTION asinh (BYVAL x AS DOUBLE) AS DOUBLE
   DIM AS DOUBLE a, s, a2
   a = fabs(x)
   s = IIF(x < 0, -1, 1)
   IF a > 1 / SQRT_DBL_EPSILON THEN
      FUNCTION = s * (log(a) + M_LN2)
   ELSEIF a > 2 THEN
      FUNCTION = s * log(2 * a + 1 / (a + sqr(a * a + 1)))
   ELSEIF a > SQRT_DBL_EPSILON THEN
      a2 = a * a
      FUNCTION = s * log1p(a + a2 / (1 + sqr(1 + a2)))
   ELSE
      FUNCTION = x
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
PRIVATE FUNCTION atanh (BYVAL x AS DOUBLE) AS DOUBLE
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
' Computes the value of sqr(x^2 + y^2 + z^2) in a way that avoids overflow.
' ========================================================================================
FUNCTION hypot3 (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE, BYVAL z AS DOUBLE) AS DOUBLE
   DIM AS DOUBLE xabs, yabs, zabs, w, r
   xabs = fabs(x)
   yabs = fabs(y)
   zabs = fabs(z)
   w = MAX(xabs, MAX(yabs, zabs))
   IF w = 0 THEN
      FUNCTION = 0
   ELSE
      r = w * sqr((xabs / w) * (xabs / w) + _
                  (yabs / w) * (yabs / w) + _
                  (zabs / w) * (zabs / w))
      FUNCTION = r
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether the argument is an infinity.
' Returns +1 if x is positive infinity, -1 if x is negative infinity and 0 otherwise.
' ========================================================================================
PRIVATE FUNCTION isinf (BYVAL x AS DOUBLE) AS LONG
   IF _finite(x) = TRUE AND _isnan(x) = FALSE THEN
      FUNCTION = IIF(x > 0, 1, -1)
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether x and y are approximately equal to a relative accuracy EPSILON.
' ========================================================================================
PRIVATE FUNCTION fcmp (BYVAL x1 AS DOUBLE, BYVAL x2 AS DOUBLE, BYVAL epsilon AS DOUBLE) AS LONG
   DIM exponent AS LONG
   DIM AS DOUBLE delta, diff, dmax
   ' /* Find exponent of largest absolute value */
   dmax = IIF(fabs(x1) > fabs(x2), x1, x2)
   frexp(dmax, @exponent)
   ' /* Form a neighborhood of size  2 * delta */
   delta = ldexp(epsilon, exponent)
   diff = x1 - x2
   IF diff > delta THEN        ' /* x1 > x2 */
      FUNCTION = 1
   ELSEIF diff < -delta THEN   ' /* x1 < x2 */
      FUNCTION = -1
   ELSE                        ' /* -delta <= diff <= delta */
      FUNCTION = 0             ' /* x1 ~=~ x2 */
   END IF
END FUNCTION
' ========================================================================================

' ########################################################################################
' Complex numbers wrapper functions
' ########################################################################################

' ========================================================================================
' * Sets the real and imaginary parts of the complex number.
' Example: DIM cx AS gsl_complex = CSet(3, 4)
' But you can al so use: DIM cx AS gsl_complex = (3, 4)
' ========================================================================================
PRIVATE FUNCTION CSet (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE) AS gsl_complex
   RETURN TYPE<gsl_complex>(x, y)
END FUNCTION
' ========================================================================================
' =====================================================================================
' * Sets the real and imaginary parts of the complex number.
' Example: DIM cx AS gsl_complex = CRect(3, 4)
' =====================================================================================
PRIVATE FUNCTION CRect (BYVAL x AS DOUBLE, BYVAL y AS DOUBLE) AS gsl_complex
   RETURN TYPE<gsl_complex>(x, y)
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets the real and imaginary parts of the complex number in polar format.
' Example: DIM cx AS gsl_complex = CPolar(3, 4)
' =====================================================================================
PRIVATE FUNCTION CPolar (BYVAL r AS DOUBLE, BYVAL theta AS DOUBLE) AS gsl_complex
   RETURN TYPE<gsl_complex> (r * cos(theta), r * sin(theta))
END FUNCTION
' =====================================================================================

' ========================================================================================
' * Returns the real part of a complex number.
' ========================================================================================
PRIVATE FUNCTION CGetReal (BYREF z AS gsl_complex) AS DOUBLE
   RETURN z.x
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the imaginary part of a complex number.
' ========================================================================================
PRIVATE FUNCTION CGetImag (BYREF z AS gsl_complex) AS DOUBLE
   RETURN z.y
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the real part of a complex number.
' ========================================================================================
PRIVATE SUB CSetReal (BYREF z AS gsl_complex, BYVAL x AS DOUBLE)
   z.x = x
END SUB
' ========================================================================================

' ========================================================================================
' * Sets the imaginary part of a complex number.
' ========================================================================================
PRIVATE SUB CSetImag (BYREF z AS gsl_complex, BYVAL y AS DOUBLE)
   z.y = y
END SUB
' ========================================================================================

'/* Complex numbers */

' ========================================================================================
' * Returns true if the two complex numbers are equal, or false otherwise.
' ========================================================================================
PRIVATE OPERATOR = (BYREF z1 AS gsl_complex, BYREF z2 AS gsl_complex) AS BOOLEAN
   RETURN (z1.x = z2.x AND z1.y = z2.y)
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Returns true if the two complex numbers are different, or false otherwise.
' ========================================================================================
PRIVATE OPERATOR <> (BYREF z1 AS gsl_complex, BYREF z2 AS gsl_complex) AS BOOLEAN
   RETURN (z1.x <> z2.x OR z1.y <> z2.y)
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Exchanges the contents of two complex numbers.
' ========================================================================================
PRIVATE SUB CSwap (BYREF z1 AS gsl_complex, BYREF z2 AS gsl_complex)
   DIM z AS gsl_complex = z1 : z1 = z2 : z2 = z
END SUB
' ========================================================================================

'/* Complex arithmetic operators */

' ========================================================================================
' * Returns the negative of a complex number.
' ========================================================================================
PRIVATE FUNCTION CNeg (BYREF z AS gsl_complex) AS gsl_complex
   RETURN TYPE<gsl_complex>(-z.x, -z.y)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CNegate (BYREF z AS gsl_complex) AS gsl_complex
   RETURN TYPE<gsl_complex>(-z.x, -z.y)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CNegative (BYREF z AS gsl_complex) AS gsl_complex
   RETURN TYPE<gsl_complex>(-z.x, -z.y)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the sum of complex numbers.
' ========================================================================================
PRIVATE OPERATOR + (BYREF z1 AS gsl_complex, BYREF z2 AS gsl_complex) AS gsl_complex
   RETURN TYPE<gsl_complex>(z1.x + z2.x, z1.y + z2.y)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR + (BYVAL a AS DOUBLE, BYREF z AS gsl_complex) AS gsl_complex
   RETURN TYPE<gsl_complex>(z.x + a, z.y)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR + (BYREF z AS gsl_complex, BYVAL a AS DOUBLE) AS gsl_complex
   RETURN TYPE<gsl_complex>(z.x + a, z.y)
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Returns the difference of complex numbers.
' ========================================================================================
PRIVATE OPERATOR - (BYREF z1 AS gsl_complex, BYREF z2 AS gsl_complex) AS gsl_complex
   RETURN TYPE<gsl_complex>(z1.x - z2.x, z1.y - z2.y)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR - (BYVAL a AS DOUBLE, BYREF z AS gsl_complex) AS gsl_complex
   RETURN TYPE<gsl_complex>(a - z.x, -z.y)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR - (BYREF z AS gsl_complex, BYVAL a AS DOUBLE) AS gsl_complex
  RETURN TYPE<gsl_complex> (z.x - a, z.y)
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Returns the negative of a complex number.
' ========================================================================================
PRIVATE OPERATOR - (BYREF z AS gsl_complex) AS gsl_complex
   RETURN TYPE<gsl_complex>(-z.x, -z.y)
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Returns the product of complex numbers.
' ========================================================================================
PRIVATE OPERATOR * (BYREF z1 AS gsl_complex, BYREF z2 AS gsl_complex) AS gsl_complex
   RETURN TYPE<gsl_complex>(z1.x * z2.x - z1.y * z2.y, z1.x * z2.y + z1.y * z2.x)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR * (BYVAL a AS DOUBLE, BYREF z AS gsl_complex) AS gsl_complex
   RETURN TYPE<gsl_complex> (a * z.x, a * z.y)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR * (BYREF z AS gsl_complex, BYVAL a AS DOUBLE) AS gsl_complex
   RETURN TYPE<gsl_complex> (a * z.x, a * z.y)
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Returns the quotient of complex numbers.
' ========================================================================================
PRIVATE OPERATOR / (BYREF z1 AS gsl_complex, BYREF z2 AS gsl_complex) AS gsl_complex
   DIM d AS DOUBLE = 1 / (z2.x * z2.x + z2.y * z2.y)
   RETURN TYPE <gsl_complex>((z1.x * z2.x + z1.y * z2.y) * d, _
                          (z1.y * z2.x - z1.x * z2.y) * d)
END OPERATOR
' ========================================================================================
' ========================================================================================
' * Returns the quotient of the complex number a and the real number.
' ========================================================================================
PRIVATE OPERATOR / (BYVAL a AS DOUBLE, BYREF z AS gsl_complex) AS gsl_complex
   DIM d AS DOUBLE = a / (z.x * z.x + z.y * z.y)
   RETURN TYPE<gsl_complex>(z.x * d, -z.y * d)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR / (BYREF z AS gsl_complex, BYVAL a AS DOUBLE) AS gsl_complex
   RETURN TYPE<gsl_complex>(z.x / a, z.y / a)
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Returns the sum of a complex number and a real number.
' ========================================================================================
PRIVATE FUNCTION CAddReal (BYREF z AS gsl_complex, BYVAL x AS DOUBLE) AS gsl_complex
   RETURN TYPE<gsl_complex>(z.x + x, z.y)
END FUNCTION
' ========================================================================================
' ========================================================================================
' * This function returns the sum of a complex number and an imaginary number.
' ========================================================================================
PRIVATE FUNCTION CAddImag (BYREF z AS gsl_complex, BYVAL y AS DOUBLE) AS gsl_complex
   RETURN TYPE<gsl_complex>(z.x, z.y + y)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the difference of a complex number and a real number.
' ========================================================================================
PRIVATE FUNCTION CSubReal (BYREF z AS gsl_complex, BYVAL x AS DOUBLE) AS gsl_complex
   RETURN TYPE<gsl_complex>(z.x - x, z.y)
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Returns the difference of a complex number and an imaginary number.
' ========================================================================================
PRIVATE FUNCTION CSubImag (BYREF z AS gsl_complex, BYVAL y AS DOUBLE) AS gsl_complex
   RETURN TYPE<gsl_complex>(z.x, z.y - y)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the product of a complex number and a real number.
' ========================================================================================
PRIVATE FUNCTION CMulReal (BYREF z AS gsl_complex, BYVAL x AS DOUBLE) AS gsl_complex
   RETURN TYPE<gsl_complex>(z.x * x, z.y * x)
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Returns the product of a complex number and an imaginary number.
' ========================================================================================
PRIVATE FUNCTION CMulImag (BYREF z AS gsl_complex, BYVAL y AS DOUBLE) AS gsl_complex
   RETURN TYPE<gsl_complex>(-y * z.y, y * z.x)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the quotient of a complex number and a real number.
' ========================================================================================
PRIVATE FUNCTION CDivReal (BYREF z AS gsl_complex, BYVAL x AS DOUBLE) AS gsl_complex
   RETURN TYPE<gsl_complex>(z.x / x, z.y / x)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the quotient of a complex number and an imaginary number.
' ========================================================================================
PRIVATE FUNCTION CDivImag (BYREF z AS gsl_complex, BYVAL y AS DOUBLE) AS gsl_complex
   RETURN TYPE<gsl_complex>(z.y / y, -z.x / y)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex conjugate of a complex number.
' Example:
'   DIM cpx AS gsl_complex = (2, 3)
'   cpx = CConj(cpx)
'   PRINT CStr(cpx)
' Output: 2 - 3 * i
' ========================================================================================
PRIVATE FUNCTION CConj (BYREF z AS gsl_complex) AS gsl_complex
   RETURN TYPE<gsl_complex>(z.x, -z.y)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CConjugate (BYREF z AS gsl_complex) AS gsl_complex
   RETURN TYPE<gsl_complex>(z.x, -z.y)
END FUNCTION
' ========================================================================================

'/* Properties of complex numbers */

' =====================================================================================
' * Returns the argument of a complex number.
' Examples:
'   DIM cpx AS gsl_complex = (1, 0)
'   DIM d AS DOUBLE = CArg(cpx)
'   PRINT d
' Output: 0.0
'   DIM cpx AS gsl_complex = (0, 1)
'   DIM d AS DOUBLE = CArg(cpx)
'   PRINT d
' Output: 1.570796326794897
'   DIM cpx AS gsl_complex = (0, -1)
'   DIM d AS DOUBLE = CArg(cpx)
' PRINT d
' Output: -1.570796326794897
'   DIM cpx AS gsl_complex = (-1, 0)
'   DIM d AS DOUBLE = CArg(cpx)
' PRINT d
' Output: 3.141592653589793
' =====================================================================================
PRIVATE FUNCTION CArg (BYREF z AS gsl_complex) AS DOUBLE
   FUNCTION = atan2(z.y, z.x)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CArgument (BYREF z AS gsl_complex) AS DOUBLE
   FUNCTION = atan2(z.y, z.x)
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Returns the magnitude of a complex number.
' Example:
'   DIM cpx AS gsl_complex = (2, 3)
'   DIM d AS DOUBLE = CAbs(cpx)
'   PRINT d
' Output: 3.605551275463989
' =====================================================================================
PRIVATE FUNCTION CAbs (BYREF z AS gsl_complex) AS DOUBLE
   FUNCTION = _hypot(z.x, z.y)
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Returns the squared magnitude of a complex number, otherwise known as the complex norm.
' Examples:
'   DIM cpx AS gsl_complex = (2, 3)
'   DIM d AS DOUBLE = CAbs2(cpx)
'   PRINT d
' Output: 13
' --or--
'   DIM cpx AS gsl_complex = (2, 3)
'   DIM d AS DOUBLE = CNorm(cpx)
'   PRINT d
' Output: 13
' =====================================================================================
PRIVATE FUNCTION CAbs2 (BYREF z AS gsl_complex) AS DOUBLE
   FUNCTION = (z.x * z.x) + (z.y * z.y)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CNorm (BYREF z AS gsl_complex) AS DOUBLE
   FUNCTION = (z.x * z.x) + (z.y * z.y)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CPhase (BYREF z AS gsl_complex) AS DOUBLE
   FUNCTION = (z.x * z.x) + (z.y * z.y)
END FUNCTION
' =====================================================================================





' ========================================================================================
' * Returns the inverse, or reciprocal, of a complex number.
' ========================================================================================
PRIVATE FUNCTION CInverse (BYREF a AS gsl_complex) AS gsl_complex
   DIM d AS DOUBLE, z AS gsl_complex
   d = 1.0 / CAbs(a)
   z.x = (a.x * d) * d
   z.y = -(a.y * d) * d
   RETURN z
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CReciprocal (BYREF a AS gsl_complex) AS gsl_complex
   DIM d AS DOUBLE, z AS gsl_complex
   d = 1.0 / CAbs(a)
   z.x = (a.x * d) * d
   z.y = -(a.y * d) * d
   RETURN z
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the natural logarithm of the magnitude of the complex number z, log|z|.
' It allows an accurate evaluation of \log|z| when |z| is close to one.  The direct
' evaluation of log(CAbs(z)) would lead to a loss of precision in this case.
' Example:
'   DIM cpx AS gsl_complex = (1.1, 0.1)
'   DIM d AS DOUBLE = CLogAbs(cpx)
'   print d
' Output: 0.09942542937258267
' ========================================================================================
PRIVATE FUNCTION CLogAbs (BYREF z AS gsl_complex) AS DOUBLE
   DIM AS DOUBLE xabs, yabs, dmax, u
   xabs = fabs(z.x)
   yabs = fabs(z.y)
   IF xabs >= yabs THEN
      dmax = xabs
      u = yabs / xabs
   ELSE
      dmax = yabs
      u = xabs / yabs
   END IF
   ' /* Handle underflow when u is close to 0 */
   FUNCTION = log(dmax) + 0.5 * log1p(u * u)
END FUNCTION
' ========================================================================================

'/* Elementary Complex Functions */

' ========================================================================================
' * Returns the square root of the complex number z. The branch cut is the negative
' real axis. The result always lies in the right half of the complex plane.
' Example:
'   DIM cpx AS gsl_complex = (2, 3)
'   cpx = CSqr(cpx)
'   PRINT CStr(cpx)
' Output: 1.67414922803554 +0.8959774761298383 * i
' Compute the square root of -1:
'   DIM cpx AS gsl_complex = (-1)
'   cpx = CSqr(cpx)
'   PRINT CStr(cpx)
' Output: 0 +1.0 * i
' ========================================================================================
FUNCTION CSqr (BYREF a AS gsl_complex) AS gsl_complex
   DIM z AS gsl_complex
   DIM AS DOUBLE x, y, w, t, ai, vi
   IF a.x = 0 AND a.y = 0 THEN RETURN z
   x = fabs(a.x)
   y = fabs(a.y)
   IF x >= y THEN
      t = y / x
      w = sqr(x) * sqr(0.5 * (1.0 + sqr(1.0 + t * t)))
   ELSE
      t = x / y
      w = sqr(y) * sqr(0.5 * (t + sqr(1.0 + t * t)))
   END IF
   IF a.x >= 0 THEN
      ai = a.y
      z.x = w
      z.y = ai / (2.0 * w)
   ELSE
      ai = a.y
      vi = IIF(ai >= 0, w, -w)
      z.x = ai / (2.0 * vi)
      z.y = vi
   END IF
   RETURN z
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex square root of the real number x, where x may be negative.
' ========================================================================================
PRIVATE FUNCTION CSqrReal (BYVAL x AS DOUBLE) AS gsl_complex
   DIM z AS gsl_complex
   IF x >= 0 THEN z.x = sqr(x) ELSE z.y = sqr(-x)
   RETURN z
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex exponential of a complex number.
' Example:
'   DIM cpx AS gsl_complex = (1, 1)
'   cpx = CExp(cpx)
'   PRINT CStr(cpx)
' Output: 1.468693939915885 +2.287355287178842 * i
' ========================================================================================
PRIVATE FUNCTION CExp (BYREF a AS gsl_complex) AS gsl_complex
   DIM z AS gsl_complex
   DIM AS DOUBLE rho, theta
   rho = exp(a.x)
   theta = a.y
   z.x = rho * cos(theta)
   z.y = rho * sin(theta)
   RETURN z
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex natural logarithm (base e) of a complex number.
' The branch cut is the negative real axis.
' Example:
'   DIM cpx AS gsl_complex = (1, 1)
'   cpx = CLog(cpx)
'   PRINT CStr(cpx)
' Output: 0.3465735902799726 +0.7853981633974483 * i
'   DIM cpx AS gsl_complex = (0, 0)
'   cpx = CLog(cpx)
'   PRINT CStr(cpx)
' Output: -1.#IND (indetermined)
' ========================================================================================
PRIVATE FUNCTION CLog (BYREF a AS gsl_complex) AS gsl_complex
   DIM z AS gsl_complex
   z.x = CLogAbs(a)
   z.y = CArg(a)
   RETURN z
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex base-10 logarithm of a complex number.
' Example:
'   DIM cpx AS gsl_complex = (1, 1)
'   cpx = CLog10(cpx)
'   PRINT CStr(cpx)
' Output: 0.1505149978319906 +0.3410940884604603 * i
' ========================================================================================
PRIVATE FUNCTION CLog10 (BYREF a AS gsl_complex) AS gsl_complex
   DIM z AS gsl_complex
   z = CLog(a)
   RETURN CMulReal(z, 1 / log(10))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex base-b logarithm of a complex number.
' This quantity is computed as the ratio \log(a)/\log(b).
' Example:
'   DIM cpx AS gsl_complex = (1, 1)
'   DIM d AS DOUBLE = CAbs2(cpx)
'   DIM cpxlog AS gsl_complex = (10, 0)
'   cpx = CLogb(cpx, cpxlog)
'   PRINT CStr(cpx)
' Output: 0.1505149978319906 +0.3410940884604602 * i
' ========================================================================================
PRIVATE FUNCTION CLogb (BYREF a AS gsl_complex, BYREF b AS gsl_complex) AS gsl_complex
   DIM AS gsl_complex z1, z2
   z1 = CLog(a)
   z2 = CLog(b)
   RETURN z1 / z2
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CLog_b (BYREF a AS gsl_complex, BYREF b AS gsl_complex) AS gsl_complex
   DIM AS gsl_complex z1, z2
   z1 = CLog(a)
   z2 = CLog(b)
   RETURN z1 / z2
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex number a raised to the complex power b.
' ========================================================================================
PRIVATE FUNCTION CPow (BYREF a AS gsl_complex, BYREF b AS gsl_complex) AS gsl_complex
   DIM z AS gsl_complex
   DIM AS DOUBLE logr, theta, rho, beta
   IF a.x = 0 AND a.y = 0 THEN
      IF b.x = 0 AND b.y = 0 THEN
         z.x = 1
      END IF
   ELSEIF b.x = 1 AND b.y = 0 THEN
      z = a
   ELSEIF b.x = -1 AND b.y = 0 THEN
      z = CInverse(a)
   ELSE
      logr = CLogAbs(a)
      theta = CArg(a)
      rho = exp(logr * b.x - b.y * theta)
      beta = theta * b.x + b.y * logr
      z.x = rho * cos(beta)
      z.y = rho * sin(beta)
   END IF
   RETURN z
END FUNCTION
' ========================================================================================
' ========================================================================================
OPERATOR ^ (BYREF a AS gsl_complex, BYREF b AS gsl_complex) AS gsl_complex
   OPERATOR = CPow(a, b)
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Returns a complex number a raised to the real power b.
' ========================================================================================
PRIVATE FUNCTION CPowReal (BYREF a AS gsl_complex, BYVAL b AS DOUBLE) AS gsl_complex
   DIM z AS gsl_complex
   DIM AS DOUBLE logr, theta, br, rho, beta
   IF a.x = 0 AND a.y = 0 THEN RETURN z
   logr = CLogAbs(a)
   theta = CArg(a)
   rho = exp(logr * b)
   beta = theta * b
   z.x = rho * cos(beta)
   z.y = rho * sin(beta)
   RETURN z
END FUNCTION
' ========================================================================================
' ========================================================================================
OPERATOR ^ (BYREF a AS gsl_complex, BYREF b AS DOUBLE) AS gsl_complex
   OPERATOR = CPowReal(a, b)
END OPERATOR
' ========================================================================================
' ========================================================================================
OPERATOR ^ (BYREF a AS DOUBLE, BYREF b AS gsl_complex) AS gsl_complex
   OPERATOR = CExp(b * log(a))
END OPERATOR
' ========================================================================================


'/* Complex Trigonometric Functions */

' ========================================================================================
' * Returns the complex sine of a complex number.
' Example:
'   DIM cpx AS gsl_complex = (1, 1)
'   cpx = CSin(cpx)
'   PRINT CStr(cpx)
' Output: 1.298457581415977 +0.6349639147847361 * i
' ========================================================================================
PRIVATE FUNCTION CSin (BYREF a AS gsl_complex) AS gsl_complex
   DIM z AS gsl_complex
   IF a.y = 0 THEN
      z.x = sin(a.x)
   ELSE
      z.x = sin(a.x) * cosh(a.y)
      z.y = cos(a.x) * sinh(a.y)
   END IF
   RETURN z
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex cosine of a complex number.
' Example:
'   DIM cpx AS gsl_complex = (1, 1)
'   cpx = CCos(cpx)
'   PRINT CStr(cpx)
' Output: 0.8337300251311491 -0.9888977057628651 * i
' ========================================================================================
PRIVATE FUNCTION CCos (BYREF a AS gsl_complex) AS gsl_complex
   DIM z AS gsl_complex
   IF a.y = 0 THEN
      z.x = cos(a.x)
   ELSE
      z.x = cos(a.x) * cosh(a.y)
      z.y = sin(a.x) * sinh(-a.y)
   END IF
   RETURN z
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex secant of a complex number.
' Example:
'   DIM cpx AS gsl_complex = (1, 1)
'   cpx = CSec(cpx)
'   PRINT CStr(cpx)
' Output: 0.4983370305551869 +0.591083841721045 * i
' ========================================================================================
PRIVATE FUNCTION CSec (BYREF a AS gsl_complex) AS gsl_complex
   DIM z AS gsl_complex = CCos(a)
   RETURN CInverse(z)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex cosecant of a complex number.
' Example:
'   DIM cpx AS gsl_complex = (1, 1)
'   cpx = CCsc(cpx)
'   PRINT CStr(cpx)
' Output: 0.6215180171704285 -0.3039310016284265 * i
' ========================================================================================
PRIVATE FUNCTION CCsc (BYREF a AS gsl_complex) AS gsl_complex
   DIM z AS gsl_complex
   z = CSin(a)
   RETURN CInverse(z)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex tangent of a complex number.
' Example:
'   DIM cpx AS gsl_complex = (1, 1)
'   cpx = CTan(cpx)
'   PRINT CStr(cpx)
' Output: 0.2717525853195117 +1.083923327338695 * i
' ========================================================================================
PRIVATE FUNCTION CTan (BYREF a AS gsl_complex) AS gsl_complex
   DIM z AS gsl_complex
   DIM AS DOUBLE u, C, D, S, T
   IF fabs(a.y) < 1 THEN
      D = pow(cos(a.x), 2.0) + pow(sinh(a.y), 2.0)
      z.x = 0.5 * sin(2 * a.x) / D
      z.y = 0.5 * sinh(2 * a.y) / D
   ELSE
      u = exp(-a.y)
      C = 2 * u / (1 - pow(u, 2.0))
      D = 1 + pow(cos(a.x), 2.0) * pow(C, 2.0)
      S = pow(C, 2.0)
      T = 1.0 / tanh(a.y)
      z.x = 0.5 * sin(2 * a.x) * S / D
      z.y = T / D
   END IF
   RETURN z
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the complex cotangent of a complex number.
' ========================================================================================
PRIVATE FUNCTION CCot (BYREF a AS gsl_complex) AS gsl_complex
   DIM z AS gsl_complex
   z = CTan(a)
   RETURN CInverse(z)
END FUNCTION
' ========================================================================================

'/* Inverse Complex Trigonometric Functions */

' ========================================================================================
' Returns the complex arcsine of a real number.
' For a between -1 and 1, the function returns a real value in the range [-pi/2, pi/2].
' For a less than -1 the result has a real part of -pi/2 and a positive imaginary part.
' For a greater than 1 the result has a real part of pi/2 and a negative imaginary part.
' Example:
'   DIM cpx AS gsl_complex
'   cpx = CArcSinReal(1)
'   PRINT CStr(cpx)
' Output: 1.570796326794897 +0 * i
' ========================================================================================
PRIVATE FUNCTION CArcSinReal (BYVAL a AS DOUBLE) AS gsl_complex
   DIM z AS gsl_complex
   IF fabs(a) <= 1 THEN
      z.x = asin(a)
   ELSE
      IF a < 0 THEN
         z.x = -M_PI_2
         z.y = acosh(-a)
      ELSE
         z.x = M_PI_2
         z.y = -acosh(a)
      END IF
   END IF
   RETURN z
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex arcsine of a complex number.
' The branch cuts are on the real axis, less than -1 and greater than 1.
' Example:
'   DIM cpx AS gsl_complex = (1, 1)
'   cpx = CArcSin(cpx)
'   PRINT CStr(cpx)
' Output: 0.6662394324925152 +1.061275061905036 * i
' ========================================================================================
PRIVATE FUNCTION CArcSin (BYREF a AS gsl_complex) AS gsl_complex

   DIM z AS gsl_complex
   DIM AS DOUBLE x, y, r_, s, A_, B, y2, real, imag
   DIM AS DOUBLE A_crossover, B_crossover, D, Apx, Am1

   IF a.y = 0 THEN
      z = CArcSinReal(a.x)
   ELSE
      x = fabs(a.x)
      y = fabs(a.y)
      r_ = hypot(x + 1, y)
      s = hypot(x - 1, y)
      A_ = 0.5 * (r_ + s)
      B = x / A_
      y2 = y * y
      A_crossover = 1.5
      B_crossover = 0.6417
      IF B <= B_crossover THEN
         real = asin(B)
      ELSE
         IF x <= 1 THEN
            D = 0.5 * (A_ + x) * (y2 / (r_ + x + 1) + (s + (1 - x)))
            real = atan_(x / sqr(D))
         ELSE
            Apx = A_ + x
            D = 0.5 * (Apx / (r_ + x + 1) + Apx / (s + (x - 1)))
            real = atan_(x / (y * sqr(D)))
         END IF
      END IF
      IF A_ <= A_crossover THEN
         IF x < 1 THEN
            Am1 = 0.5 * (y2 / (r_ + (x + 1)) + y2 / (s + (1 - x)))
         ELSE
            Am1 = 0.5 * (y2 / (r_ + (x + 1)) + (s + (x - 1)))
         END IF
         imag = log1p(Am1 + sqr(Am1 * (A_ + 1)))
      ELSE
         imag = log(A_ + sqr(A_ * A_ - 1))
      END IF
      z.x = IIF(a.x >= 0, real, -real)
      z.y = IIF(a.y >= 0, imag, -imag)
   END IF

   RETURN z

END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the complex arccosine of a real number.
' For a between -1 and 1, the function returns a real value in the range [0, pi].
' For a less than -1 the result has a real part of pi and a negative imaginary part.
' For a greater than 1 the result is purely imaginary and positive.
' ========================================================================================
PRIVATE FUNCTION CArcCosReal (BYVAL a AS DOUBLE) AS gsl_complex
   DIM z AS gsl_complex
   IF fabs(a) <= 1.0 THEN
      z.x = acos(a)
   ELSE
      IF a < 0 THEN
         z.x = M_PI
         z.y = -acosh(-a)
      ELSE
         z.y = acosh(a)
      END IF
   END IF
   RETURN z
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex arccosine of a complex number.
' The branch cuts are on the real axis, less than -1 and greater than 1.
' Example:
'   DIM cpx AS gsl_complex = (1, 1)
'   cpx = CArcCos(cpx)
'   print CStr(cpx)
' Output: 0.9045568943023814 -1.061275061905036 * i
' ========================================================================================
PRIVATE FUNCTION CArcCos (BYREF a AS gsl_complex) AS gsl_complex

   DIM z AS gsl_complex
   DIM AS DOUBLE x, y, r_, s, A_, B, y2, real, imag
   DIM AS DOUBLE A_crossover, B_crossover, D, Apx, Am1
   IF a.y = 0 THEN
      z = CArcCosReal(a.x)
   ELSE
      x = fabs(a.x)
      y = fabs(a.y)
      r_ = hypot(x + 1, y)
      s = hypot(x - 1, y)
      A_ = 0.5 * (r_ + s)
      B = x / A_
      y2 = y * y
      A_crossover = 1.5
      B_crossover = 0.6417
      IF B <= B_crossover THEN
         real = acos(B)
      ELSE
         IF x <= 1 THEN
            D = 0.5 * (A_ + x) * (y2 / (r_ + x + 1) + (s + (1 - x)))
            real = atan_(sqr(D) / x)
         ELSE
            Apx = A_ + x
            D = 0.5 * (Apx / (r_ + x + 1) + Apx / (s + (x - 1)))
            real = atan_((y * sqr(D)) / x)
         END IF
      END IF
      IF A_ <= A_crossover THEN
         IF x < 1 THEN
            Am1 = 0.5 * (y2 / (r_ + (x + 1)) + y2 / (s + (1 - x)))
         ELSE
            Am1 = 0.5 * (y2 / (r_ + (x + 1)) + (s + (x - 1)))
         END IF
         imag = log1p(Am1 + sqr(Am1 * (A_ + 1)))
      ELSE
         imag = log(A_ + sqr(A_ * A_ - 1))
      END IF
      z.x = IIF(a.x >= 0, real, M_PI - real)
      z.y = IIF(a.y >= 0, -imag, imag)
   END IF
   RETURN z

END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex arcsecant of a complex number.
' Example:
'   DIM cpx AS gsl_complex = (1, 1)
'   cpx = CArcSec(cpx)
'   PRINT CStr(cpx)
' Output: 1.118517879643706 +0.5306375309525176 * i
' ========================================================================================
PRIVATE FUNCTION CArcSec (BYREF a AS gsl_complex) AS gsl_complex
   DIM z AS gsl_complex = CInverse(a)
   RETURN CArcCos(z)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the complex arcsecant of the a number.
' ========================================================================================
PRIVATE FUNCTION CArcSecReal (BYVAL a AS DOUBLE) AS gsl_complex
   DIM z AS gsl_complex
   IF a <= -1 OR a >= 1 THEN
      z.x = acos(1 / a)
   ELSE
      IF a >= 0 THEN
         z.y = acosh(1 / a)
      ELSE
         z.x = M_PI
         z.y = -acosh(-1 / a)
      END IF
   END IF
   RETURN z
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex arccosecant of a complex number.
' Example:
'   DIM cpx AS gsl_complex = (1, 1)
'   cpx = CArcCsc(cpx)
'   PRINT CStr(cpx)
' Output: 0.4522784471511905 -0.5306375309525176 * i
' ========================================================================================
PRIVATE FUNCTION CArcCsc (BYREF a AS gsl_complex) AS gsl_complex
   DIM z AS gsl_complex = CInverse(a)
   RETURN CArcSin(z)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the complex arccosecant of a real number.
' ========================================================================================
PRIVATE FUNCTION CArcCscReal (BYVAL a AS DOUBLE) AS gsl_complex
   DIM z AS gsl_complex
   IF a <= -1 OR a >= 1 THEN
      z.x = asin(1 / a)
   ELSE
      IF a >= 0 THEN
         z.x = M_PI_2
         z.y = -acosh(1 / a)
      ELSE
         z.x = -M_PI_2
         z.y = acosh(-1 / a)
      END IF
   END IF
   RETURN z
END FUNCTION
' ========================================================================================

' ========================================================================================
' * This function returns the complex arctangent of a complex number.
' The branch cuts are on the imaginary axis, below -i and above i.
' Example:
'   DIM cpx AS gsl_complex = (1, 1)
'   cpx = CArcTan(cpx)
'   PRINT CStr(cpx)
' Output: 1.017221967897851 +0.4023594781085251 * i
' ========================================================================================
PRIVATE FUNCTION CArcTan (BYREF a AS gsl_complex) AS gsl_complex

   DIM z AS gsl_complex
   DIM AS DOUBLE r_, imag, u, A_, B

   IF a.y = 0 THEN
      z.x = atan_(a.x)
   ELSE
      ' /* FIXME: This is a naive implementation which does not fully
      '    take into account cancellation errors, overflow, underflow
      '    etc.  It would benefit from the Hull et al treatment. */
      r_ = hypot(a.x, a.y)
      u = 2 * a.y / (1 + r_ * r_)

      ' /* FIXME: the following cross-over should be optimized but 0.1
      '    seems to work ok */
      IF fabs(u) < 0.1 THEN
         imag = 0.25 * (log1p(u) - log1p(-u))
      ELSE
         A_ = hypot(a.x, a.y + 1)
         B = hypot(a.x, a.y - 1)
         imag = 0.5 * log(A_ / B)
      END IF

      IF a.x = 0 THEN
         IF a.y > 1 THEN
            z.x = M_PI_2
            z.y = imag
         ELSEIF a.y < -1 THEN
            z.x = -M_PI_2
            z.y = imag
         ELSE
            z.y = imag
         END IF
      ELSE
         z.x = 0.5 * atan2(2 * a.x, ((1 + r_) * (1 - r_)))
         z.y = imag
      END IF
   END IF

   RETURN z

END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex arccotangent of a complex number.
' Example:
'   DIM cpx AS gsl_complex = (1, 1)
'   cpx = CArcCot(cpx)
'   PRINT CStr(cpx)
' Output: 0.5535743588970451 -0.4023594781085251 * i
' ========================================================================================
PRIVATE FUNCTION CArcCot (BYREF a AS gsl_complex) AS gsl_complex
   DIM z AS gsl_complex
   IF a.x = 0 AND a.y = 0 THEN
      z.x = M_PI_2
   ELSE
      z = CInverse(a)
      z = CArcTan(z)
   END IF
   RETURN z
END FUNCTION
' ========================================================================================

'/* Complex Hyperbolic Functions */

' ========================================================================================
' * Returns the complex hyperbolic sine of a complex number.
' Example:
'   DIM cpx AS gsl_complex = (1, 1)
'   cpx = CSinH(cpx)
'   PRINT CStr(cpx)
' Output: 0.6349639147847361 +1.298457581415977 * i
' ========================================================================================
PRIVATE FUNCTION CSinH (BYREF a AS gsl_complex) AS gsl_complex
   DIM z AS gsl_complex
   z.x = sinh(a.x) * cos(a.y)
   z.y = cosh(a.x) * sin(a.y)
   RETURN z
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex hyperbolic cosine of a complex number.
' Example:
'   DIM cpx AS gsl_complex = (1, 1)
'   cpx = CCosH(cpx)
'   PRINT CStr(cpx)
' Output: 0.8337300251311491 +0.9888977057628651 * i
' ========================================================================================
PRIVATE FUNCTION CCosH (BYREF a AS gsl_complex) AS gsl_complex
   DIM z AS gsl_complex
   z.x = cosh(a.x) * cos(a.y)
   z.y = sinh(a.x) * sin(a.y)
   RETURN z
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex hyperbolic secant of a complex number.
' Example:
'   DIM cpx AS gsl_complex = (1, 1)
'   cpx = CSecH(cpx)
'   PRINT CStr(cpx)
' Output: 0.4983370305551869 -0.591083841721045 * i
' ========================================================================================
PRIVATE FUNCTION CSecH (BYREF a AS gsl_complex) AS gsl_complex
   DIM z AS gsl_complex = CCosH(a)
   RETURN CInverse(z)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex hyperbolic cosecant of a complex number.
' Example:
'   DIM cpx AS gsl_complex = (1, 1)
'   cpx = CCscH(cpx)
'   PRINT CStr(cpx)
' Output: 0.3039310016284265 -0.6215180171704285 * i
' ========================================================================================
PRIVATE FUNCTION CCscH (BYREF a AS gsl_complex) AS gsl_complex
   DIM z AS gsl_complex = CSinH(a)
   RETURN CInverse(z)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex hyperbolic tangent of a complex number.
' Example:
'   DIM cpx AS gsl_complex = (1, 1)
'   cpx = CTanH(cpx)
'   PRINT CStr(cpx)
' Output: 1.083923327338695 +0.2717525853195117 * i
' ========================================================================================
PRIVATE FUNCTION CTanH (BYREF a AS gsl_complex) AS gsl_complex
   DIM z AS gsl_complex
   DIM AS DOUBLE D, F
   IF fabs(a.x) < 1 THEN
      D = pow(cos(a.y), 2.0) + pow(sinh(a.x), 2.0)
      z.x = sinh(a.x) * cosh(a.x) / D
      z.y = 0.5 * sin(2 * a.y) / D
   ELSE
      D = pow(cos(a.y), 2.0) + pow(sinh(a.x), 2.0)
      F = 1 + pow(cos(a.y) / sinh(a.x), 2.0)
      z.x = 1.0 / (tanh(a.x) * F)
      z.y = 0.5 * sin(2 * a.y) / D
   END IF
   RETURN z
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex hyperbolic cotangent of a complex number.
' Example:
'   DIM cpx AS gsl_complex = (1, 1)
'   cpx = CCotH(cpx)
'   PRINT CStr(cpx)
' Output: 0.868014142895925 -0.2176215618544027 * i
' ========================================================================================
PRIVATE FUNCTION CCotH (BYREF a AS gsl_complex) AS gsl_complex
   DIM z AS gsl_complex = CTanH(a)
   RETURN CInverse(z)
END FUNCTION
' ========================================================================================

'/* Inverse Complex Hyperbolic Functions */

' ========================================================================================
' * Returns the complex hyperbolic arcsine of a complex number.
' The branch cuts are on the imaginary axis, below -i and above i.
' Example:
'   DIM cpx AS gsl_complex = (1, 1)
'   cpx = CArcSinH(cpx)
'   PRINT CStr(cpx)
' Output: 1.061275061905036 +0.6662394324925152 * i
' ========================================================================================
PRIVATE FUNCTION CArcSinH (BYREF a AS gsl_complex) AS gsl_complex
   DIM z AS gsl_complex = CMulImag(a, 1.0)
   z = CArcSin(z)
   RETURN CMulImag(z, -1.0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex hyperbolic arccosine of a complex number.
' The branch cut is on the real axis, less than 1.
' Example:
'   DIM cpx AS gsl_complex = (1, 1)
'   cpx = CArcCosH(cpx)
'   PRINT CStr(cpx)
' Output: 1.061275061905036 +0.9045568943023814 * i
' ========================================================================================
PRIVATE FUNCTION CArcCosH (BYREF a AS gsl_complex) AS gsl_complex
   DIM z AS gsl_complex = CArcCos(a)
   RETURN CMulImag(z, IIF(z.y > 0, -1.0, 1.0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the complex hyperbolic arccosine of a real number.
' ========================================================================================
PRIVATE FUNCTION CArcCosHReal (BYVAL a AS DOUBLE) AS gsl_complex
   DIM z AS gsl_complex
   IF a >= 1 THEN
      z.x = acosh(a)
   ELSE
      IF a >= -1 THEN
         z.y = acos(a)
      ELSE
         z.x = acosh(-a)
         z.y = M_PI
      END IF
   END IF
   RETURN z
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex hyperbolic arcsecant of a complex number.
' Example:
'   DIM cpx AS gsl_complex = (1, 1)
'   cpx = CArcSecH(cpx)
'   PRINT CStr(cpx)
' Output: 0.5306375309525176 -1.118517879643706 * i
' ========================================================================================
PRIVATE FUNCTION CArcSecH (BYREF a AS gsl_complex) AS gsl_complex
   DIM z AS gsl_complex = CInverse(a)
   RETURN CArcCosH(z)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * This function returns the complex hyperbolic arccosecant of a complex number.
' Example:
'   DIM cpx AS gsl_complex = (1, 1)
'   cpx = CArcCscH(cpx)
'   PRINT CStr(cpx)
' Output: 0.5306375309525176 -0.4522784471511905 * i
' ========================================================================================
PRIVATE FUNCTION CArcCscH (BYREF a AS gsl_complex) AS gsl_complex
   DIM z AS gsl_complex = CInverse(a)
   RETURN CArcSinH(z)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the complex hyperbolic arctangent of a real number.
' ========================================================================================
PRIVATE FUNCTION CArcTanhReal (BYVAL a AS DOUBLE) AS gsl_complex
   DIM z AS gsl_complex
   IF a > -1 AND a < 1 THEN
      z.x = atanh(a)
   ELSE
      z.x = atanh(1 / a)
      z.y = IIF(a < 0, M_PI_2, -M_PI_2)
   END IF
   RETURN z
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex hyperbolic arctangent of a complex number.
' The branch cuts are on the real axis, less than -1 and greater than 1.
' Example:
'   DIM cpx AS gsl_complex = (1, 1)
'   cpx = CArcTanH(cpx)
'   PRINT CStr(cpx)
' Output: 0.4023594781085251 +1.017221967897851 * i
' ========================================================================================
PRIVATE FUNCTION CArcTanH (BYREF a AS gsl_complex) AS gsl_complex
   DIM z AS gsl_complex
   IF a.y = 0 THEN
      z = CArcTanhReal(a.x)
   ELSE
      z = CMulImag(a, 1)
      z = CArcTan(z)
      z = CMulImag(z, -1)
   END IF
   RETURN z
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the complex hyperbolic arccotangent of a complex number.
' Example:
'   DIM cpx AS gsl_complex = (1, 1)
'   cpx = CArcCotH(cpx)
'   PRINT CStr(cpx)
' Output: 0.4023594781085251 -0.5535743588970451 * i
' ========================================================================================
PRIVATE FUNCTION CArcCotH (BYREF a AS gsl_complex) AS gsl_complex
   DIM z AS gsl_complex = CInverse(a)
   RETURN CArcTanH(z)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the sign of a complex number.
' If number is greater than zero, then CSgn returns 1.
' If number is equal to zero, then CSgn returns 0.
' If number is less than zero, then CSgn returns -1.
' ========================================================================================
PRIVATE FUNCTION CSgn (BYREF z AS gsl_complex) AS LONG
   IF z.x > 0 THEN RETURN 1
   IF z.x < 0 THEN RETURN -1
   IF z.y > 0 THEN RETURN 1
   IF z.y < 0 THEN RETURN -1
   RETURN 0
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a complex number as a formated string.
' ========================================================================================
PRIVATE FUNCTION CStr (BYREF z AS gsl_complex) AS STRING
   FUNCTION = STR(z.x) & IIF(SGN(z.y) = -1 OR SGN(z.y) = -0, " ", " +") & STR(z.y) & " * i"
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a complex number in polar form: Z = |Z| * exp(Arg(Z) * i)
' ========================================================================================
PRIVATE FUNCTION CStrPolar (BYREF z AS gsl_complex) AS STRING
   FUNCTION = STR(CAbs(z)) & " * exp(" & STR(CArg(z)) & " * i)"
END FUNCTION
' ========================================================================================

END NAMESPACE
