' ########################################################################################
' Microsoft Windows
' File: CGpMatrix.inc
' Contents: GDI+ matrix class.
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

' ========================================================================================
' * Creates and initializes a Matrix object that represents the identity matrix.
' ========================================================================================
PRIVATE CONSTRUCTOR CGpMatrix
   m_Status = GdipCreateMatrix(@m_pMatrix)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' * Creates and initializes a Matrix object based on six numbers that define an affine transformation.
' ========================================================================================
PRIVATE CONSTRUCTOR CGpMatrix (BYVAL m11 AS SINGLE, BYVAL m12 AS SINGLE, BYVAL m21 AS SINGLE, BYVAL m22 AS SINGLE, BYVAL dx AS SINGLE, BYVAL dy AS SINGLE)
   m_Status = GdipCreateMatrix2(m11, m12, m21, m22, dx, dy, @m_pMatrix)
END CONSTRUCTOR
' ========================================================================================
' =====================================================================================
' Creates a Matrix object based on a rectangle and a point.
' =====================================================================================
PRIVATE CONSTRUCTOR CGpMatrix (BYVAL rcf AS GpRectF PTR, BYVAL dstplg AS GpPointF PTR)
   m_Status = GdipCreateMatrix3(rcf, dstplg, @m_pMatrix)
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Creates a Matrix object based on a rectangle and a point.
' =====================================================================================
PRIVATE CONSTRUCTOR CGpMatrix (BYVAL rc AS GpRect PTR, BYVAL dstplg AS GpPoint PTR)
   m_Status = GdipCreateMatrix3I(rc, dstplg, @m_pMatrix)
END CONSTRUCTOR
' =====================================================================================

' ========================================================================================
' Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR CGpMatrix
   IF m_pMatrix THEN GdipDeleteMatrix(m_pMatrix)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last status code.
' ========================================================================================
PRIVATE FUNCTION CGpMatrix.GetLastStatus () AS GpStatus
   RETURN m_Status
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last status code.
' ========================================================================================
PRIVATE FUNCTION CGpMatrix.SetStatus (BYVAL status AS GpStatus) AS GpStatus
   m_Status = status
   RETURN m_Status
END FUNCTION
' ========================================================================================

' =====================================================================================
' Gets the elements of this matrix. The elements are placed in an array in the order
' m11, m12, m21, m22, m31, m32, where mij denotes the element in row i, column j.
' The size of the array should be 6 × sizeof(SINGLE).
' =====================================================================================
PRIVATE FUNCTION CGpMatrix.GetElements (BYVAL m AS SINGLE PTR) AS GpStatus
   RETURN SetStatus(GdipGetMatrixElements(m_pMatrix, m))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Gets the elements of this matrix. The elements are placed in an array in the order
' m11, m12, m21, m22, m31, m32, where mij denotes the element in row i, column j.
' =====================================================================================
PRIVATE FUNCTION CGpMatrix.SetElements (BYVAL m11 AS SINGLE, BYVAL m12 AS SINGLE, BYVAL m21 AS SINGLE, BYVAL m22 AS SINGLE, BYVAL dx AS SINGLE, BYVAL dy AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipSetMatrixElements(m_pMatrix, m11, m12, m21, m22, dx, dy))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Gets the horizontal translation value of this matrix, which is the element in row 3, column 1.
' =====================================================================================
PRIVATE FUNCTION CGpMatrix.OffsetX () AS SINGLE
   DIM elements(0 TO 5) AS SINGLE
   IF this.GetElements(@elements(0)) = OK THEN RETURN elements(4)
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Gets the vertical translation value of this matrix, which is the element in row 3, column 2.
' =====================================================================================
PRIVATE FUNCTION CGpMatrix.OffsetY () AS SINGLE
   DIM elements(0 TO 5) AS SINGLE
   IF this.GetElements(@elements(0)) = OK THEN RETURN elements(5)
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Updates this matrix with the elements of the identity matrix.
' =====================================================================================
PRIVATE FUNCTION CGpMatrix.Reset () AS GpStatus
   RETURN SetStatus(GdipSetMatrixElements(m_pMatrix, 1, 0, 0, 1, 0, 0))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Updates this matrix with the product of itself and another matrix.
' =====================================================================================
PRIVATE FUNCTION CGpMatrix.Multiply (BYVAL pMatrix AS CGpMatrix PTR, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   RETURN SetStatus(GdipMultiplyMatrix(m_pMatrix, pMatrix->m_pMatrix, order))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Updates this matrix with the product of itself and a translation matrix.
' =====================================================================================
PRIVATE FUNCTION CGpMatrix.Translate (BYVAL nOffsetX AS SINGLE, BYVAL nOffsetY AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   RETURN SetStatus(GdipTranslateMatrix(m_pMatrix, nOffsetX, nOffsetY, order))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Updates this matrix with the product of itself and a scaling matrix.
' =====================================================================================
PRIVATE FUNCTION CGpMatrix.Scale (BYVAL scaleX AS SINGLE, BYVAL scaleY AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   RETURN SetStatus(GdipScaleMatrix(m_pMatrix, scaleX, scaleY, order))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Updates this matrix with the product of itself and a rotation matrix.
' =====================================================================================
PRIVATE FUNCTION CGpMatrix.Rotate (BYVAL angle AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   RETURN SetStatus(GdipRotateMatrix(m_pMatrix, angle, order))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Updates this matrix with the product of itself and a matrix that represents rotation
' about a specified point.
' =====================================================================================
PRIVATE FUNCTION CGpMatrix.RotateAt (BYVAL angle AS SINGLE, BYVAL center AS GpPointF PTR, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   IF order = MatrixOrderPrepend THEN
      m_Status = GdipTranslateMatrix(m_pMatrix, center->X, center->Y, order)
      m_Status = GdipRotateMatrix(m_pMatrix, angle, order)
      RETURN SetStatus(GdipTranslateMatrix(m_pMatrix, -center->X, -center->Y, order))
   ELSE
      m_Status = GdipTranslateMatrix(m_pMatrix, -center->X, -center->Y, order)
      m_Status = GdipRotateMatrix(m_pMatrix, angle, order)
      RETURN SetStatus(GdipTranslateMatrix(m_pMatrix, center->X, center->Y, order))
   END IF
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Updates this matrix with the product of itself and a shearing matrix.
' =====================================================================================
PRIVATE FUNCTION CGpMatrix.Shear (BYVAL shearX AS SINGLE, BYVAL shearY AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   RETURN SetStatus(GdipShearMatrix(m_pMatrix, shearX, shearY, order))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * If the matrix is invertible, this function replaces its elements  with the elements of its inverse.
' =====================================================================================
PRIVATE FUNCTION CGpMatrix.Invert () AS GpStatus
   RETURN SetStatus(GdipInvertMatrix(m_pMatrix))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Multiplies each point in an array by this matrix. Each point is treated as a row matrix.
' The multiplication is performed with the row matrix on the left and this matrix on the right.
' =====================================================================================
PRIVATE FUNCTION CGpMatrix.TransformPoints (BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_ = 1) AS GpStatus
   RETURN SetStatus(GdipTransformMatrixPoints(m_pMatrix, pts, nCount))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpMatrix.TransformPoints (BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_ = 1) AS GpStatus
   RETURN SetStatus(GdipTransformMatrixPointsI(m_pMatrix, pts, nCount))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Multiplies each vector in an array by this matrix. The translation elements of this
' matrix (third row) are ignored. Each vector is treated as a row matrix. The multiplication
' is performed with the row matrix on the left and this matrix on the right.
' =====================================================================================
PRIVATE FUNCTION CGpMatrix.TransformVectors (BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_ = 1) AS GpStatus
   RETURN SetStatus(GdipVectorTransformMatrixPoints(m_pMatrix, pts, nCount))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpMatrix.TransformVectors (BYVAL pts AS GpPoint pTR, BYVAL nCount AS INT_ = 1) AS GpStatus
   RETURN SetStatus(GdipVectorTransformMatrixPointsI(m_pMatrix, pts, nCount))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Determines whether this matrix is invertible.
' =====================================================================================
PRIVATE FUNCTION CGpMatrix.IsInvertible () AS BOOLEAN
   DIM result AS BOOL
   m_Status = GdipIsMatrixInvertible(m_pMatrix, @result)
   RETURN result
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Determines whether this matrix is the identity matrix.
' =====================================================================================
PRIVATE FUNCTION CGpMatrix.IsIdentity () AS BOOLEAN
   DIM result AS BOOL
   m_Status = GdipIsMatrixIdentity(m_pMatrix, @result)
   RETURN result
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Determines whether the elements of this matrix are equal to the elements of another matrix.
' =====================================================================================
PRIVATE FUNCTION CGpMatrix.Equals (BYVAL pMatrix AS CGpMatrix PTR) AS BOOLEAN
   IF pMatrix = NULL THEN RETURN FALSE
   DIM result AS BOOL
   m_Status = GdipIsMatrixEqual(m_pMatrix, pMatrix->m_pMatrix, @result)
   RETURN result
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Creates a new Matrix object based on this Matrix object.
' =====================================================================================
PRIVATE FUNCTION CGpMatrix.Clone (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
   IF pMatrix = NULL THEN RETURN SetStatus(InvalidParameter)
   IF pMatrix->m_pMatrix THEN GdipDeleteMatrix(pMatrix->m_pMatrix)
   RETURN SetStatus(GdipCloneMatrix(m_pMatrix, @pMatrix->m_pMatrix))
END FUNCTION
' =====================================================================================
