# CGpMatrix Class

Encapsulates a 3-by-3 affine matrix that represents a geometric transform. A **Matrix** object represents a 3×3 matrix that, in turn, represents an affine transformation. A **Matrix** object stores only six of the 9 numbers in a 3×3 matrix because all 3×3 matrices that represent affine transformations have the same third column (0, 0, 1).

**Inherits from**: CGpBase.<br>
**Imclude file**: CGpMatrix.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#Constructors) | Creates and initializes a **Matrix** object that represents the identity matrix. |
| [Clone](#Clone) | Copies the contents of the existing **Matrix** object into a new **Matrix** object. |
| [Equals](#Equals) | Determines whether the elements of this matrix are equal to the elements of another matrix. |
| [GetElements](#GetElements) | Gets the elements of this matrix. |
| [Invert](#Invert) | If this matrix is invertible, the Invert method replaces the elements of this matrix with the elements of its inverse. |
| [IsIdentity](#IsIdentity) | Determines whether this matrix is the identity matrix. |
| [IsInvertible](#IsInvertible) | Determines whether this matrix is invertible. |
| [Multiply](#Multiply) | Updates this matrix with the product of itself and another matrix. |
| [OffsetX](#OffsetX) | Gets the horizontal translation value of this matrix, which is the element in row 3, column 1. |
| [OffsetY](#OffsetY) | Gets the vertical translation value of this matrix, which is the element in row 3, column 1. |
| [Reset](#Reset) | Updates this matrix with the elements of the identity matrix. |
| [Rotate](#Rotate) | Updates this matrix with the product of itself and a rotation matrix. |
| [RotateAt](#RotateAt) | Updates this matrix with the product of itself and a matrix that represents rotation about a specified point. |
| [Scale](#Scale) | Scales this matrix with the product of itself and a scaling matrix. |
| [SetElements](#SetElements) | Sets the elements of this matrix. |
| [Shear](#Shear) | Updates this matrix with the product of itself and a shearing matrix. |
| [TransformPoints](#TransformPoints) | Multiplies each point in an array by this matrix. Each point is treated as a row matrix. The multiplication is performed with the row matrix on the left and this matrix on the right. |
| [TransformVectors](#TransformVectors) | Multiplies each vector in an array by this matrix. |
| [Translate](#Translate) | Updates this matrix with the product of itself and a translation matrix. |

# <a name="Constructors"></a>Constructors

Encapsulates a 3-by-3 affine matrix that represents a geometric transform. A **Matrix** object represents a 3 ×3 matrix that, in turn, represents an affine transformation. A **Matrix** object stores only six of the 9 numbers in a 3 ×3 matrix because all 3 ×3 matrices that represent affine transformations have the same third column (0, 0, 1).

```
CONSTRUCTOR CGpMatrix
CONSTRUCTOR CGpMatrix (BYVAL pMatrix AS CGpMatrix PTR)
CONSTRUCTOR CGpMatrix (BYVAL m11 AS SINGLE, BYVAL m12 AS SINGLE,  BYVAL m21 AS SINGLE, _
   BYVAL m22 AS SINGLE, BYVAL dx AS SINGLE, BYVAL dy AS SINGLE)
CONSTRUCTOR CGpMatrix (BYVAL m11 AS INT_, BYVAL m12 AS INT_, BYVAL m21 AS INT_, _
   BYVAL m22 AS INT_, BYVAL dx AS INT_, BYVAL dy AS INT_)
CONSTRUCTOR CGpMatrix (BYVAL rc AS GpRectF PTR, BYVAL dstplg AS GpPointF PTR)
CONSTRUCTOR CGpMatrix (BYVAL rc AS GpRect PTR, BYVAL dstplg AS GpPoint PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMatrix* | A pointer to another **Matrix** object to be cloned. |
| *m11* | The element in the first row, first column. |
| *m12* | The element in the first row, second column. |
| *m21* | The element in in the second row, first column.  |
| *m22* | The element in the second row, second column.  |
| *dx* | The element in the third row, first column. |
| *dy* | The element in the third row, second column. |
| *rc* | Reference to a **GpRectF** or **GpRect** structure. The **X** data member of the rectangle specifies the matrix element in row 1, column 1. The **Y** data member of the rectangle specifies the matrix element in row 1, column 2. The **Width** data member of the rectangle specifies the matrix element in row 2, column 1. The **Height** data member of the rectangle specifies the matrix element in row 2, column 2. |
| *dstplg* | Pointer to a **GpPointF** or **GpPoint** structure. The **X** data member of the point specifies the matrix element in row 3, column 1. The **Y** data member of the point specifies the matrix element in row 3, column 2. |

# <a name="Clone"></a>Clone

Copies the contents of the existing **Matrix** object into a new **Matrix** object.

```
FUNCTION Clone (BYVAL pCloneMatrix AS CGpMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pCloneMatrix* | Pointer to the Matrix object where to copy the contents of the existing object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Matrix object that represents a horizontal scaling. The
' code creates a second matrix by calling the Matrix::Clone method of the first matrix.
' Then the code calls the Matrix::Translate method of the second matrix. At that point,
' the second matrix represents a composite transformation: first scale, then translate.
' The code passes the address of the first matrix to the SetTransform method of a Graphics
' object and then uses that Graphics object to draw a scaled rectangle. Then the code passes
' the address of the second matrix to the SetTransform method of the same Graphics object.
' Finally, the code calls the DrawRectangle method of the Graphics object a second time to
' draw a rectangle that is scaled and translated.
' ========================================================================================
SUB Example_Clone (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96

   ' // Horizontal scaling
   DIM matrix AS CGpMatrix = CGpMatrix(3.0, 0.0, 0.0, 1.0, 0.0, 0.0)
   ' // Clone the matrix
   DIM clonedMatrix AS CGpMatrix
   matrix.Clone(@clonedMatrix)
   ' // You can also use:
   ' DIM clonedMatrix AS CGpMatrix = @matrix)

   ' // Translate the cloned matrix
   clonedMatrix.Translate(40.0 * rxRatio, 25.0 * ryRatio, MatrixOrderAppend)

   ' // Create a pem
   DIM myPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255))
   ' // Scale
   graphics.SetTransform(@matrix)
   graphics.DrawRectangle(@myPen, 0, 0, 100 * rxRatio, 100 * ryRatio)
   ' // Scale and translate
   graphics.SetTransform(@clonedMatrix)
   graphics.DrawRectangle(@myPen, 0, 0, 100 * rxRatio, 100 * ryRatio)

END SUB
' ========================================================================================
```

# <a name="Equals"></a>Equals

Determines whether the elements of this matrix are equal to the elements of another matrix.

```
FUNCTION Equals (BYVAL pMatrix AS CGpMatrix PTR) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMatrix* | Pointer to a **Matrix** object that is compared with this **Matrix** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
#INCLUDE ONCE "Afx/CGdiPlus/CGdiPlus.inc"
USING Afx

' // Must be constructed with NEW to we able to delete them before the call to AfxGdipShutdown
DIM pMat1 AS CGpMatrix PTR = NEW CGpMatrix(1.0, 2.0, 1.0, 1.0, 2.0, 2.0)
DIM pMat2 AS CGpMatrix PTR = NEW CGpMatrix(1.0, 2.0, 1.0, 1.0, 2.0, 2.0)

IF pMat1->Equals(pMat2) THEN
   PRINT "They are the same"
ELSE
   PRINT "They are the different"
END IF

' // Must be deleted before the call to AfxGdipShutdown
Delete pMat1
Delete pMat2
```

# <a name="GetElements"></a>GetElements

Gets the elements of this matrix. The elements are placed in an array in the order m11, m12, m21, m22, m31, m32, where mij denotes the element in row i, column j.

```
FUNCTION GetElements (BYVAL m AS SINGLE PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *m* | Out. Pointer to an array that receives the matrix elements. The size of the array should be 6 × 4 (the size of a simple precision number). |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="Invert"></a>Invert

If this matrix is invertible, the **Invert** method replaces the elements of this matrix with the elements of its inverse.

```
FUNCTION Invert () AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *m* | Out. Pointer to an array that receives the matrix elements. The size of the array should be 6 × 4 (the size of a simple precision number). |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

If this matrix is not invertible, the method fails and returns **InvalidParameter**.

#### Example

```
' ========================================================================================
' The following example passes the address of a Matrix object to the SetTransform method
' of a Graphics object and then draws a rectangle. The rectangle is translated 30 units
' right and 20 units down by the world transformation of the Graphics object. The code
' calls the Matrix.Invert method of the Matrix object and sets the world transformation
' of the Graphics object to the inverted matrix. The code draws a second rectangle that is
' translated 30 units left and 20 units up.
' ========================================================================================
SUB Example_Invert (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96

   DIM matrix AS CGpMatrix = CGpMatrix(1.0, 0.0, 0.0, 1.0, 30.0 * rxRatio, 20.0 * ryRatio)
   DIM myPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), rxRatio)

   graphics.SetTransform(@matrix)
   graphics.DrawRectangle(@myPen, 0, 0, 200 * rxRatio, 100 * ryRatio)
   matrix.Invert
   graphics.SetTransform(@matrix)
   graphics.DrawRectangle(@myPen, 0, 0, 200 * rxRatio, 100 * ryRatio)

END SUB
' ========================================================================================
```

# <a name="IsIdentity"></a>IsIdentity

Determines whether this matrix is the identity matrix.

```
FUNCTION IsIdentity () AS BOOLEAN
```

#### Return value

If this matrix is the identity matrix, this method returns TRUE; otherwise, it returns FALSE.

#### Example

```
#INCLUDE ONCE "Afx/CGdiPlus/CGdiPlus.inc"
USING Afx

' // Must be constructed with NEW to we able to delete it before the call to AfxGdipShutdown
DIM pMatrix AS CGpMatrix PTR = NEW CGpMatrix(1.0, 0.0, 0.0, 1.0, 0.0, 2.0)

IF pMatrix->IsIdentity THEN
   PRINT "The matrix is the identity matrix"
ELSE
   PRINT "The matrix is not the identity matrix"
END IF

' // Must be deleted before the call to AfxGdipShutdown
Delete pMatrix
```

# <a name="IsInvertible"></a>IsInvertible

Determines whether this matrix is invertible.

```
FUNCTION IsInvertible () AS BOOLEAN
```

#### Return value

If this matrix is the invertible, this method returns TRUE; otherwise, it returns FALSE.

#### Example

```
#INCLUDE ONCE "Afx/CGdiPlus/CGdiPlus.inc"
USING Afx

' // Must be constructed with NEW to we able to delete it before the call to AfxGdipShutdown
DIM pMatrix AS CGpMatrix PTR = NEW CGpMatrix(3.0, 0.0, 0.0, 2.0, 20.0, 10.0)

IF pMatrix->IsInvertible THEN
   PRINT "The matrix is invertible"
ELSE
   PRINT "The matrix is not invertible"
END IF

' // Must be deleted before the call to AfxGdipShutdown
Delete pMatrix
```

# <a name="Multiply"></a>Multiply

Updates this matrix with the product of itself and another matrix.

```
FUNCTION Multiply (BYVAL pMatrix AS CGpMatrix PTR, _
   BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMatrix* | Pointer to a matrix that will be multiplied by this matrix. |
| *order* | Optional. Element of the **MatrixOrder** enumeration that specifies the order of the multiplication. **MatrixOrderPrepend** specifies that the passed matrix is on the left, and **MatrixOrderAppend** specifies that the passed matrix is on the right. The default value is **MatrixOrderPrepend**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates two Matrix objects: matrix1 and matrix2. The call to the
' Matrix.Multiply method updates matrix1 with the product ( matrix1)( matrix2). At that
' point, matrix1 represents a composite transformation: first scale, then translate. The
' code uses matrix1 to set the world transformation of a Graphics object and then draws a
' rectangle that is transformed according to that world transformation.
' ========================================================================================
SUB Example_Multiply (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96

   ' // Create a pen
   DIM myPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255))

   ' // Horizontal scale
   DIM matrix1 AS CGpMatrix = CGpMatrix(3.0, 0.0, 0.0, 1.0, 0.0, 0.0)
   ' // Translation
   DIM matrix2 AS CGpMatrix = CGpMatrix(1.0, 0.0, 0.0, 1.0, 20.0 * rxRatio, 40.0 * ryRatio)

   matrix1.Multiply(@matrix2, MatrixOrderAppend)
   graphics.SetTransform(@matrix1)
   graphics.DrawRectangle(@myPen, 0, 0, 100 * rxRatio, 100 * ryRatio)

END SUB
' ========================================================================================
```

# <a name="OffsetX"></a>OffsetX

Gets the horizontal translation value of this matrix, which is the element in row 3, column 1.

```
FUNCTION OffsetX () AS SINGLE
```

#### Return value

This method returns the horizontal translation value of this matrix, which is the element in row 3, column 1.

#### Example

```
' ========================================================================================
' The following example creates a Matrix object with a horizontal translation value of 50
' and a vertical translation value of 30. The code calls the Matrix::OffsetX and Matrix.OffsetY
' methods of the Matrix object to obtain those translation values. The code draws a line
' from (0, 0) to the point whose coordinates are the retrieved translation values. The code
' also uses the matrix to set the world transformation of a Graphics object and then draws
' a rectangle that is transformed according to that world transformation.
' ========================================================================================
SUB Example_Offset (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96

   ' // Create a pen
   DIM myPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), rxRatio)

   DIM matrix AS CGpMatrix = CGpMatrix(1.0, 0.0, 0.0, 1.0, 50.0 * rxRatio, 30.0 * ryRatio)
   DIM AS SINGLE xTranslation, yTranslation
   xTranslation = matrix.OffsetX
   yTranslation = matrix.OffsetY

   graphics.DrawLine(@myPen, 0.0, 0.0, xTranslation, yTranslation)
   graphics.SetTransform(@matrix)
   graphics.DrawRectangle(@myPen, 0, 0, 20 * rxRatio, 20 * ryRatio)

END SUB
' ========================================================================================
```

# <a name="OffsetY"></a>OffsetY

Gets the vertical translation value of this matrix, which is the element in row 3, column 1.

```
FUNCTION OffsetY () AS SINGLE
```

#### Return value

This method returns the vertical translation value of this matrix, which is the element in row 3, column 2.

#### Example

```
' ========================================================================================
' The following example creates a Matrix object with a horizontal translation value of 50
' and a vertical translation value of 30. The code calls the Matrix::OffsetX and Matrix.OffsetY
' methods of the Matrix object to obtain those translation values. The code draws a line
' from (0, 0) to the point whose coordinates are the retrieved translation values. The code
' also uses the matrix to set the world transformation of a Graphics object and then draws
' a rectangle that is transformed according to that world transformation.
' ========================================================================================
SUB Example_Offset (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96

   ' // Create a pen
   DIM myPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), rxRatio)

   DIM matrix AS CGpMatrix = CGpMatrix(1.0, 0.0, 0.0, 1.0, 50.0 * rxRatio, 30.0 * ryRatio)
   DIM AS SINGLE xTranslation, yTranslation
   xTranslation = matrix.OffsetX
   yTranslation = matrix.OffsetY

   graphics.DrawLine(@myPen, 0.0, 0.0, xTranslation, yTranslation)
   graphics.SetTransform(@matrix)
   graphics.DrawRectangle(@myPen, 0, 0, 20 * rxRatio, 20 * ryRatio)

END SUB
' ========================================================================================
```

# <a name="Reset"></a>Reset

Updates this matrix with the elements of the identity matrix.

```
FUNCTION Reset () AS GpStatus
```

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Matrix object that represents a horizontal scaling by a
' factor of 5 and a vertical scaling by a factor of 3. The code calls the Matrix.Reset
' method to replace the elements of that matrix with the elements of the identity matrix.
' Then the code calls the Matrix::Translate method to update the matrix with the product
' of itself (the identity) and a translation matrix. The result is that the matrix
' represents only the translation, not the scaling. The code uses the matrix to set the
' world transformation of a Graphics object and then draws a rectangle that is transformed
' according to that world transformation.
' ========================================================================================
SUB Example_Reset (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96

   ' // Create a pen
   DIM myPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), rxRatio)

   DIM matrix AS CGpMatrix = CGpMatrix(5.0, 0.0, 0.0, 3.0, 0.0, 0.0)
   matrix.Reset
   matrix.Translate(50.0 * rxRatio, 40.0 * ryRatio)

   graphics.SetTransform(@matrix)
   graphics.DrawRectangle(@myPen, 0, 0, 100 * rxRatio, 100 * ryRatio)

END SUB
' ========================================================================================
```

# <a name="Rotate"></a>Rotate

Updates this matrix with the product of itself and a rotation matrix.

```
FUNCTION Rotate (BYVAL angle AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *angle* | Simple precision number that specifies the angle of rotation in degrees. Positive values specify clockwise rotation. |
| *order* | Optional. Element of the **MatrixOrder** enumeration that specifies the order of the multiplication. **MatrixOrderPrepend** specifies that the rotation matrix is on the left, and **MatrixOrderAppend** specifies that the rotation matrix is on the right. The default value is **MatrixOrderPrepend**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Matrix object and calls the Matrix::Translate method to
' set the elements of that matrix to a translation. Then the code calls the Matrix.Rotate
' method to update the matrix with the product of itself and a rotation matrix. At that
' point, the matrix represents a composite transformation: first translate, then rotate.
' The code uses the matrix to set the world transformation of a Graphics object and then
' draws an ellipse that is transformed according to the composite transformation.
' ========================================================================================
SUB Example_Rotate (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96

   ' // Create a pen
   DIM myPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), rxRatio)

   DIM matrix AS CGpMatrix
   ' // First a translation
   matrix.Translate(40.0 * rxRatio, 0.0)
   ' // Then a rotation
   matrix.Rotate(30.0, MatrixOrderAppend)

   graphics.SetTransform(@matrix)
   graphics.DrawEllipse(@myPen, 0, 0, 100 * rxRatio, 50 * ryRatio)

END SUB
' ========================================================================================
```

# <a name="RotateAt"></a>RotateAt

Updates this matrix with the product of itself and a matrix that represents rotation about a specified point.

```
FUNCTION RotateAt (BYVAL angle AS SINGLE, BYVAL center AS GpPointF PTR, _
   BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *angle* | Simple precision number that specifies the angle of rotation in degrees. Positive values specify clockwise rotation. |
| *center* | Reference to a **GpPointF** object that specifies the center of the rotation. This is the point about which the rotation takes place. |
| *order* | Optional. Element of the **MatrixOrder** enumeration that specifies the order of the multiplication. **MatrixOrderPrepend** specifies that the rotation matrix is on the left, and **MatrixOrderAppend** specifies that the rotation matrix is on the right. The default value is **MatrixOrderPrepend**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Matrix object and calls the Matrix.Translate method to
' set the elements of that matrix to a translation (150 right, 100 down). Then the code
' calls the Matrix.RotateAt method to update the matrix with the product of itself and a
' matrix that represents a 30-degree rotation about the point (150, 100). The matrix then
' represents a composite transformation: first translate (150 right, 100 down), then rotate
' about the point (150, 100). The code uses the matrix to set the world transformation of
' a Graphics object and then draws an ellipse that is transformed according to the
' composite transformation.
' ========================================================================================
SUB Example_RotateAt (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96

   ' // Create a pen
   DIM myPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), rxRatio)

   DIM matrix AS CGpMatrix
   matrix.Translate(150.0 * rxRatio, 100.0 * ryRatio)
   matrix.RotateAt(30.0, @TYPE<GpPointF>(150.0 * rxRatio, 100.0 * ryRatio), MatrixOrderAppend)

   graphics.SetTransform(@matrix)

   ' // Draw a tramsformed ellipse. The composite transformation
   ' // is translate right 150, translate down 100, and then
   ' // rotate 30 degrees about the point (150, 100).
   graphics.DrawEllipse(@myPen, -40 * rxRatio, -20 * ryRatio, 80 * rxRatio, 40 * ryRatio)

   ' // Draw rotated axes with the origin at the center of the ellipse.
   graphics.DrawLine(@myPen, -50 * rxRatio, 0, 50 * ryRatio, 0)
   graphics.DrawLine(@myPen, 0, -50 * rxRatio, 0, 50 * ryRatio)

END SUB
' ========================================================================================
```

# <a name="Scale"></a>Scale

Scales this matrix with the product of itself and a scaling matrix.

```
FUNCTION Scale (BYVAL scaleX AS SINGLE, BYVAL scaleY AS SINGLE, _
   BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *scaleX* | Simple precision that specifies the horizontal scale factor. |
| *scaleY* | Simple precision that specifies the vertical scale factor. |
| *order* | Optional. Element of the MatrixOrder enumeration that specifies the order of the multiplication. M**atrixOrderPrepend** specifies that the rotation matrix is on the left, and **MatrixOrderAppend** specifies that the rotation matrix is on the right. The default value is **MatrixOrderPrepend**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Matrix object and calls the Matrix.Rotate method to set
' the elements of that matrix to a rotation. Then the code calls the Matrix.Scale method
' to update the matrix with the product of itself and a scaling matrix. At that point, the
' matrix represents a composite transformation: first rotate, then scale. The code uses
' the matrix to set the world transformation of a Graphics object and then draws an ellipse
' that is transformed according to the composite transformation.
' ========================================================================================
SUB Example_Scale (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96

   ' // Create a pen
   DIM myPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), rxRatio)

   DIM matrix AS CGpMatrix
   matrix.Rotate(30.0)
   matrix.Scale(3.0, 2.0, MatrixOrderAppend)

   graphics.SetTransform(@matrix)

   ' // Draw a transformed ellipse. The composite transformation
   ' // is rotate 30 degrees and then scale by a factor of 3
   ' // in the horizontal direction and 2 in the vertical direction.
   graphics.DrawEllipse(@myPen, 0, 0, 80 * rxRatio, 40 * ryRatio)

END SUB
' ========================================================================================
```

# <a name="SetElements"></a>SetElements

Sets the elements of this matrix.

```
FUNCTION SetElements (BYVAL m11 AS SINGLE, BYVAL m12 AS SINGLE, BYVAL m21 AS SINGLE, _
   BYVAL m22 AS SINGLE, BYVAL dx AS SINGLE, BYVAL dy AS SINGLE) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *m11* | Simple precision number that specifies the element in the first row, first column. |
| *m12* | Simple precision number that specifies the element in the first row, second column. |
| *m21* | Simple precision number that specifies the element in the second row, first column. |
| *m22* | Simple precision number that specifies the element in the second row, second column. |
| *dx* | Simple precision number that specifies the element in the third row, first column. |
| *dy* | Simple precision number that specifies the element in the third row, second column. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example calls the Matrix.SetElements method of a Matrix object. The code
' uses the matrix to set the world transformation of a Graphics object and then draws a
' rectangle that is transformed by that world transformation.
' ========================================================================================
SUB Example_MatrixSetElements (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96

   ' // Create a pen
   DIM myPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), rxRatio)

   DIM matrix AS CGpMatrix
   matrix.SetElements(1.0, 0.0, 0.0, 1.0, 30.0 * rxRatio, 50.0 * ryRatio)

   graphics.SetTransform(@matrix)
   graphics.DrawRectangle(@myPen, 0, 0, 80 * rxRatio, 40 * ryRatio)

END SUB
' ========================================================================================
```

# <a name="Shear"></a>Shear

Updates this matrix with the product of itself and a shearing matrix.

```
FUNCTION Shear (BYVAL shearX AS SINGLE, BYVAL shearY AS SINGLE, _
   BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *shearX* | Simple precision number that specifies the horizontal shear factor. |
| *shearY* | Simple precision number that specifies the vertical shear factor. |
| *order* | Optional. Element of the **MatrixOrder** enumeration that specifies the order of the multiplication. **MatrixOrderPrepend** specifies that the rotation matrix is on the left, and **MatrixOrderAppend** specifies that the rotation matrix is on the right. The default value is **MatrixOrderPrepend**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Matrix object and calls the Matrix::Scale method to set
' the elements of that matrix to a scaling transformation. Then the code calls the Matrix.Shear
' method to update the matrix with the product of itself and a shearing matrix. At that point,
' the matrix represents a composite transformation: first scale, then shear. The code uses
' the matrix to set the world transformation of a Graphics object and then draws a rectangle
' that is transformed according to the composite transformation.
' In the call to Matrix::Shear, the shearX parameter is 3 and the shearY parameter is 0.
' That particular shearing transformation slides the bottom edge of the rectangle to the
' right. The distance that the bottom edge slides is shearX multiplied by the height of
' the rectangle after it is stretched by the scaling transformation.
' ========================================================================================
SUB Example_Shear (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96

   ' // Create a pen
   DIM myPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255))

   DIM matrix AS CGpMatrix
   ' // First a scaling
   matrix.Scale(2.0, 2.0)
   ' // Then a shear
   matrix.Shear(3.0, 0.0, MatrixOrderAppend)

   graphics.SetTransform(@matrix)
   graphics.DrawRectangle(@myPen, 0, 0, 100 * rxRatio, 50 * ryRatio)

END SUB
' ========================================================================================
```

# <a name="TransformPoints"></a>TransformPoints

Multiplies each point in an array by this matrix. Each point is treated as a row matrix. The multiplication is performed with the row matrix on the left and this matrix on the right.

```
FUNCTION TransformPoints (BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_ = 1) AS GpStatus
FUNCTION TransformPoints (BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_ = 1) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pts* | Pointer to an array of **GpPointF** or **GpPoint** objects that, on input, contains the points to be transformed and, on output, receives the transformed points. Each point in the array is transformed (multiplied by this matrix) and updated with the result of the transformation. |
| *nCount* | Optional. Integer that specifies the number of points to be transformed. The default value is 1. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates an array of five points and then transforms those points
' by calling the Matrix.TransformPoints method of a Matrix object. The code passes the
' array of points to the DrawCurve Methods method before the transformation and again
' after the transformation.
' ========================================================================================
SUB Example_TransformPoints (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96

   ' // Create a pen
   DIM myPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), rxRatio)

   DIM rgPoints(0 TO 4) AS GpPointF = {GDIP_POINTF(50 * rxRatio, 100 * ryRatio), GDIP_POINTF(100 * rxRatio, 50 * ryRatio), _
       GDIP_POINTF(160 * rxRatio, 125 * rxRatio), GDIP_POINTF(200 * rxRatio, 100 * ryRatio), GDIP_POINTF(250 * rxRatio, 150 * ryRatio)}

'#ifdef __FB_64BIT__
'   DIM rgPoints(0 TO 4) AS GpPointF = {(50 * rxRatio, 100 * ryRatio), (100 * rxRatio, 50 * ryRatio), _
'       (160 * rxRatio, 125 * rxRatio), (200 * rxRatio, 100 * ryRatio), (250 * rxRatio, 150 * ryRatio)}
'#else
'      ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM rgPoints(0 TO 4) AS GpPointF
'   rgPoints(0).x = 50 * rxRatio : rgPoints(0).y = 100 * ryRatio : rgPoints(1).x = 100 * rxRatio
'   rgPoints(1). y = 50 * ryRatio : rgPoints(2).x = 160 * rxRatio : rgPoints(2).y = 125 * ryRatio
'   rgPoints(3).x = 200 * rxRatio : rgPoints(3).y = 100 * ryRatio rgPoints(4).x = 250 * rxRatio
'   rgPoints(4).y = 150 * ryRatio
'#endif

   DIM matrix AS CGpMatrix = CGpMatrix(1.0, 0.0, 0.0, 2.0, 0.0, 0.0)

   graphics.DrawCurve(@myPen, @rgPoints(0), 5)
   matrix.TransformPoints(@rgPoints(0), 5)
   graphics.DrawCurve(@myPen, @rgPoints(0), 5)

END SUB
' ========================================================================================
```

# <a name="TransformVectors"></a>TransformVectors

Multiplies each vector in an array by this matrix. The translation elements of this matrix (third row) are ignored. Each vector is treated as a row matrix. The multiplication is performed with the row matrix on the left and this matrix on the right.

```
FUNCTION TransformVectors (BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_ = 1) AS GpStatus
FUNCTION TransformVectors (BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_ = 1) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pts* | Pointer to an array of **GPointF** or **GpPoint** objects that, on input, contains the vectors to be transformed and, on output, receives the transformed vectors. Each vector in the array is transformed (multiplied by this matrix) and updated with the result of the transformation. |
| *nCount* | Optional. Integer that specifies the number of points to be transformed. The default value is 1. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a vector and a point. The tip of the vector and the point
' are at the same location: (100, 50). The code creates a Matrix object and initializes its
' elements so that it represents a clockwise rotation followed by a translation 100 units
' to the right. The code calls the TransformPoints Methods method of the matrix to transform
' the point and calls the Matrix.TransformVectors method of the matrix to transform the
' vector. The entire transformation (rotation followed by translation) is performed on the
' point, but only the rotation part of the transformation is performed on the vector. The
' elements of the matrix that represent translation are ignored by the Matrix.TransformVectors method.
' ========================================================================================
SUB Example_TransformVectors (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a pen
   DIM myPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), rxRatio)
   myPen.SetEndCap(LineCapArrowAnchor)
   DIM myBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 255)

   ' // A point and a vector, same representation but different behavior
   DIM pt AS GpPointF = GDIP_POINTF(100.0, 50.0)
   DIM vector AS GpPointF = GDIP_POINTF(100.0, 50.0)

   ' // Draw the original point and vector in blue
   graphics.FillEllipse(@myBrush, pt.X - 5.0, pt.Y - 5.0, 10.0, 10.0)
   graphics.DrawLine(@MyPen, 0, 0, vector.x, vector.y)

   ' // Transform
   DIM matrix AS CGpMatrix = CGpMatrix(0.8, 0.6, -0.6, 0.8, 100.0, 0.0)
   matrix.TransformPoints(@pt)
   matrix.TransformVectors(@vector)

   ' // Draw the transformed point and vector in red.
   myPen.SetColor(GDIP_ARGB(255, 255, 0, 0))
   myBrush.SetColor(GDIP_ARGB(255, 255, 0, 0))
   graphics.FillEllipse(@myBrush, pt.X - 5.0, pt.Y - 5.0, 10.0, 10.0)
   graphics.DrawLine(@MyPen, 0, 0, vector.x, vector.y)

END SUB
' ========================================================================================
```

# <a name="Translate"></a>Translate

Updates this matrix with the product of itself and a translation matrix.

```
FUNCTION Translate (BYVAL nOffsetX AS SINGLE, BYVAL nOffsetY AS SINGLE, _
   BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nOffsetX* | Simple precision number that specifies the horizontal component of the translation. |
| *nOffsetY* | Simple precision number that specifies the vertical component of the translation. |
| *order* | Optional. Element of the MatrixOrder enumeration that specifies the order of the multiplication. **MatrixOrderPrepend** specifies that the rotation matrix is on the left, and **MatrixOrderAppend** specifies that the rotation matrix is on the right. The default value is **MatrixOrderPrepend**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Matrix object and calls the Matrix.Rotate method to set
' the elements of that matrix to a rotation. Then the code calls the Matrix.Translate method
' to update the matrix with the product of itself and a translation matrix. At that point,
' the matrix represents a composite transformation: first rotate, then translate. The code
' uses the matrix to set the world transformation of a Graphics object and then draws an
' ellipse that is transformed according to the composite transformation.
' ========================================================================================
SUB Example_Translate (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96

   ' // Create a pen
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), rxRatio)

   DIM matrix AS CGpMatrix
   ' // First a rotation
   matrix.Rotate(30.0, MatrixOrderAppend)
   ' // Then a translation
   matrix.Translate(150.0 * rxRatio, 100.0 * ryRatio, MatrixOrderAppend)

   graphics.SetTransform(@matrix)

   ' // Draw a tramsformed ellipse. The composite transformation
   ' // is rotate 30 degrees, then translate 150 right and 100 down.
   graphics.DrawEllipse(@pen, -40 * rxRatio, -20 * ryRatio, 80 * rxRatio, 40 * ryRatio)

   ' // Draw rotated axes with the origin at the center of the ellipse.
   graphics.DrawLine(@pen, -50 * rxRatio, 0, 50 * ryRatio, 0)
   graphics.DrawLine(@pen, 0, -50 * rxRatio, 0, 50 * ryRatio)

END SUB
' ========================================================================================
```
