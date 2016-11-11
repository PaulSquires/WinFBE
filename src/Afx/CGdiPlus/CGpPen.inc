' ########################################################################################
' Microsoft Windows
' File: CGpPen.inc
' Contents: GDI+ pen class.
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

' ########################################################################################
' CGpPen Class
' ########################################################################################

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CGpPen
END CONSTRUCTOR
' ========================================================================================

' =====================================================================================
' * Creates a Pen object that uses a specified color and width.
' =====================================================================================
PRIVATE CONSTRUCTOR CGpPen (BYVAL colour AS ARGB, BYVAL nWidth AS SINGLE = 1.0)
   m_Status = GdipCreatePen1(colour, nWidth, UnitWorld, @m_pPen)
END CONSTRUCTOR
' =====================================================================================

' =====================================================================================
' Creates a Pen object that uses the attributes of a brush and a real number to set the
' width of this Pen object.
' =====================================================================================
PRIVATE CONSTRUCTOR CGpPen (BYVAL pBrush AS CGpBrush PTR, BYVAL nWidth AS SINGLE = 1.0)
   m_Status = GdipCreatePen2(pBrush->m_pBrush, nWidth, UnitWorld, @m_pPen)
END CONSTRUCTOR
' =====================================================================================

' ========================================================================================
' Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR CGpPen
   IF m_pPen THEN GdipDeletePen(m_pPen)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last status code.
' ========================================================================================
PRIVATE FUNCTION CGpPen.GetLastStatus () AS GpStatus
   RETURN m_Status
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last status code.
' ========================================================================================
PRIVATE FUNCTION CGpPen.SetStatus (BYVAL status AS GpStatus) AS GpStatus
   m_Status = status
   RETURN m_Status
END FUNCTION
' ========================================================================================

' =====================================================================================
' * Creates a new Pen object based on this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.Clone (BYVAL pPen AS CGpPen PTR) AS GpStatus
   IF pPen = NULL THEN RETURN SetStatus(InvalidParameter)
   IF pPen->m_pPen THEN GdipDeletePen(pPen->m_pPen)
   RETURN SetStatus(GdipClonePen(m_pPen, @pPen->m_pPen))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets the width for this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.SetWidth (BYVAL nWidth AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipSetPenWidth(m_pPen, nWidth))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Gets the width currently set for this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.GetWidth () AS SINGLE
   DIM nWidth AS SINGLE
   m_Status = GdipGetPenWidth(m_pPen, @nWidth)
   RETURN nWidth
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets the cap styles for the start, end, and dashes in a line drawn with this pen.
' =====================================================================================
PRIVATE FUNCTION CGpPen.SetLineCap (BYVAL startCap AS LineCap, BYVAL endCap AS LineCap, BYVAL nDashCap AS DashCap) AS GpStatus
   RETURN SetStatus(GdipSetPenLineCap197819(m_pPen, startCap, endCap, nDashCap))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets the start cap for this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.SetStartCap (BYVAL startCap AS LineCap) AS GpStatus
   RETURN SetStatus(GdipSetPenStartCap(m_pPen, startCap))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets the end cap for this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.SetEndCap (BYVAL endCap AS LineCap) AS GpStatus
   RETURN SetStatus(GdipSetPenEndCap(m_pPen, endCap))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets the dash cap style for this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.SetDashCap (BYVAL nDashCap AS DashCap) AS GpStatus
   RETURN SetStatus(GdipSetPenDashCap197819(m_pPen, nDashCap))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Gets the start cap currently set for this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.GetStartCap () AS LineCap
   DIM startCap AS LineCap
   m_Status = GdipGetPenStartCap(m_pPen, @startCap)
   RETURN startCap
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Gets the end cap currently set for this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.GetEndCap () AS LineCap
   DIM endCap AS LineCap
   m_Status = GdipGetPenEndCap(m_pPen, @endCap)
   RETURN endCap
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Gets the dash cap style currently set for this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.GetDashCap () AS DashCap
   DIM nDashCap AS DashCap
   m_Status = GdipGetPenDashCap197819(m_pPen, @nDashCap)
   RETURN nDashCap
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets the the line join for this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.SetLineJoin (BYVAL nLineJoin AS LineJoin) AS GpStatus
   RETURN SetStatus(GdipSetPenLineJoin(m_pPen, nLineJoin))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Gets the line join for this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.GetLineJoin () AS LineJoin
   DIM nLineJoin AS LineJoin
   m_Status = GdipGetPenLineJoin(m_pPen, @nLineJoin)
   RETURN nLineJoin
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets the custom start cap for this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.SetCustomStartCap (BYVAL pCustomLineCap AS CGpCustomLineCap PTR) AS GpStatus
   RETURN SetStatus(GdipSetPenCustomStartCap(m_pPen, pCustomLineCap->m_pCustomLineCap))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Gets the custom start cap currently set for this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.GetCustomStartCap (BYVAL pCustomLineCap AS CGpCustomLineCap PTR) AS GpStatus
   RETURN SetStatus(GdipGetPenCustomStartCap(m_pPen, @pCustomLineCap->m_pCustomLineCap))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets the custom end cap for this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.SetCustomEndCap (BYVAL pCustomLineCap AS CGpCustomLineCap PTR) AS GpStatus
   RETURN SetStatus(GdipSetPenCustomEndCap(m_pPen, IIF(pCustomLineCap, pCustomLineCap->m_pCustomLineCap, NULL)))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Gets the custom end cap currently set for this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.GetCustomEndCap (BYVAL pCustomLineCap AS CGpCustomLineCap PTR) AS GpStatus
   RETURN SetStatus(GdipGetPenCustomEndCap(m_pPen, @pCustomLineCap->m_pCustomLineCap))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the miter limit of this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.SetMiterLimit (BYVAL miterLimit AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipSetPenMiterLimit(m_pPen, miterLimit))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the miter length currently set for this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.GetMiterLimit () AS SINGLE
   DIM miterLimit AS SINGLE
   m_Status = GdipGetPenMiterLimit(m_pPen, @miterLimit)
   RETURN miterLimit
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets the alignment for this Pen object relative to the line.
' =====================================================================================
PRIVATE FUNCTION CGpPen.SetAlignment (BYVAL nPenAlignment AS PenAlignment) AS GpStatus
   RETURN SetStatus(GdipSetPenMode(m_pPen, nPenAlignment))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the alignment for this Pen object relative to the line.
' =====================================================================================
PRIVATE FUNCTION CGpPen.GetAlignment () AS PenAlignment
   DIM nPenAlignment AS PenAlignment
   m_Status = GdipGetPenMode(m_pPen, @nPenAlignment)
   FUNCTION = nPenAlignment
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets the world transformation of this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.SetTransform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
   RETURN SetStatus(GdipSetPenTransform(m_pPen, pMatrix->m_pMatrix))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the world transformation of this Pen object.
' GdipGetPenTransform is wrongly defined in the FB headers. The second parameter must be
' BYVAL matrix AS GpMatrix PTR PTR, not BYVAL matrix AS GpMatrix PTR.
' =====================================================================================
PRIVATE FUNCTION CGpPen.GetTransform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
   RETURN SetStatus(GdipGetPenTransform(m_pPen, cast(GpMatrix PTR, @pMatrix->m_pMatrix)))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets the world transformation matrix of this Pen object to the identity matrix.
' =====================================================================================
PRIVATE FUNCTION CGpPen.ResetTransform () AS GpStatus
   RETURN SetStatus(GdipResetPenTransform(m_pPen))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Updates the world transformation matrix of this Pen object with the product of itself
' and another matrix. Default value for order is MatrixOrderPrepend (0).
' =====================================================================================
PRIVATE FUNCTION CGpPen.MultiplyTransform (BYVAL pMatrix AS CGpMatrix PTR, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   RETURN SetStatus(GdipMultiplyPenTransform(m_pPen, pMatrix->m_pMatrix, order))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Translates the local geometric transformation of this Pen object by the specified dimensions.
' This FUNCTION prepends the translation to the transformation.
' Default value for order is MatrixOrderPrepend (0).
' =====================================================================================
PRIVATE FUNCTION CGpPen.TranslateTransform (BYVAL dx AS SINGLE, BYVAL dy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   RETURN SetStatus(GdipTranslatePenTransform(m_pPen, dx, dy, order))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets the Pen object's world transformation matrix equal to the product of itself and
' a scaling matrix. Default value for order is MatrixOrderPrepend (0).
' =====================================================================================
PRIVATE FUNCTION CGpPen.ScaleTransform (BYVAL sx AS SINGLE, BYVAL sy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   RETURN SetStatus(GdipScalePenTransform(m_pPen, sx, sy, order))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Updates the world transformation matrix of this Pen object with the product of itself
' and a rotation matrix. Default value for order is MatrixOrderPrepend (0).
' =====================================================================================
PRIVATE FUNCTION CGpPen.RotateTransform (BYVAL angle AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   RETURN SetStatus(GdipRotatePenTransform(m_pPen, angle, order))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the type currently set for this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.GetPenType () AS PenType
   DIM nPenType AS PenType
   m_Status = GdipGetPenFillType(m_pPen, @nPenType)
   FUNCTION = nPenType
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets the color for this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.SetColor (BYVAL colour AS ARGB) AS GpStatus
   RETURN SetStatus(GdipSetPenColor(m_pPen, colour))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets the color for this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.SetBrush (BYVAL pBrush AS CGpBrush PTR) AS GpStatus
   RETURN SetStatus(GdipSetPenBrushFill(m_pPen, pBrush->m_pBrush))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Gets the color currently set for this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.GetColor (BYVAL colour AS ARGB PTR) AS GpStatus
   IF this.GetPenType <> PenTypeSolidColor THEN RETURN SetStatus(WrongState)
   RETURN SetStatus(GdipGetPenColor(m_pPen, colour))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the Brush object that is currently set for this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.GetBrush (BYVAL pBrush AS CGpBrush PTR) AS GpStatus
   IF pBrush = NULL THEN RETURN SetStatus(InvalidParameter)
   RETURN SetStatus(GdipGetPenBrushFill(m_pPen, @pBrush->m_pBrush))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the dash style currently set for this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.GetDashStyle () AS DashStyle
   DIM nDashStyle AS DashStyle
   m_Status = GdipGetPenDashStyle(m_pPen, @nDashStyle)
   RETURN nDashStyle
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the dash style for this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.SetDashStyle (BYVAL nDashStyle AS DashStyle) AS GpStatus
   RETURN SetStatus(GdipSetPenDashStyle(m_pPen, nDashStyle))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the distance from the start of the line to the start of the first space in a
' dashed line.
' =====================================================================================
PRIVATE FUNCTION CGpPen.GetDashOffset () AS SINGLE
   DIM dashOffset AS SINGLE
   m_Status = GdipGetPenDashOffset(m_pPen, @dashOffset)
   RETURN dashOffset
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets an array of custom dashes and spaces for this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.SetDashOffset (BYVAL dashOffset AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipSetPenDashOffset(m_pPen, dashOffset))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets an array of custom dashes and spaces for this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.SetDashPattern (BYVAL dashArray AS SINGLE PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipSetPenDashArray(m_pPen, dashArray, nCount))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the number of elements in a dash pattern array.
' =====================================================================================
PRIVATE FUNCTION CGpPen.GetDashPatternCount () AS INT_
   DIM nCount AS INT_
   m_Status = GdipGetPenDashCount(m_pPen, @nCount)
   RETURN nCount
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets an array of custom dashes and spaces currently set for this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.GetDashPattern (BYVAL dashArray AS SINGLE PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipGetPenDashArray(m_pPen, dashArray, nCount))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets the compound array for this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.SetCompoundArray (BYVAL compoundArray AS SINGLE PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipSetPenCompoundArray(m_pPen, compoundArray, nCount))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the number of elements in a dash pattern array.
' =====================================================================================
PRIVATE FUNCTION CGpPen.GetCompoundArrayCount () AS GpStatus
   DIM nCount AS INT_
   m_Status = GdipGetPenCompoundCount(m_pPen, @nCount)
   RETURN nCount
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets an array of custom dashes and spaces currently set for this Pen object.
' =====================================================================================
PRIVATE FUNCTION CGpPen.GetCompoundArray (BYVAL compoundArray AS SINGLE PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipGetPenCompoundArray(m_pPen, compoundArray, nCount))
END FUNCTION
' =====================================================================================
