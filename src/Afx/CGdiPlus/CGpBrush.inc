' ########################################################################################
' Microsoft Windows
' File: CGpBrush.inc
' Contents: GDI+ Brush class
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

' ########################################################################################
' CGpBrush Class
' ########################################################################################

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CGpBrush
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR CGpBrush
   IF m_pBrush THEN GdipDeleteBrush(m_pBrush)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last status code.
' ========================================================================================
PRIVATE FUNCTION CGpBrush.GetLastStatus () AS GpStatus
   RETURN m_Status
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last status code.
' ========================================================================================
PRIVATE FUNCTION CGpBrush.SetStatus (BYVAL status AS GpStatus) AS GpStatus
   m_Status = status
   RETURN m_Status
END FUNCTION
' ========================================================================================

' =====================================================================================
' * Creates a new Brush object based on this Brush object.
' =====================================================================================
PRIVATE FUNCTION CGpBrush.Clone (BYVAL pBrush AS CGpBrush PTR) AS GpStatus
   IF pBrush = NULL THEN RETURN SetStatus(InvalidParameter)
   IF pBrush->m_pBrush THEN GdipDeleteBrush(pBrush->m_pBrush)
   RETURN SetStatus(GdipCloneBrush(m_pBrush, @pBrush->m_pBrush))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Gets the type of this brush.
' =====================================================================================
PRIVATE FUNCTION CGpBrush.GetType () AS GpBrushType
   DIM nType AS GpBrushType = -1
   m_Status = GdipGetBrushType(m_pBrush, @nType)
   RETURN nType
END FUNCTION
' =====================================================================================

' #####################################################################################
' CGpSolidBrush interface.
' Extends CGpBrush.
' #####################################################################################

' ========================================================================================
' * Creates a solid brush
' ========================================================================================
PRIVATE CONSTRUCTOR CGpSolidBrush (BYVAL colour AS ARGB = &hFF000000)
   DIM brush AS GpSolidFill PTR
   m_Status = GdipCreateSolidFill(colour, @brush)
   m_pBrush = cast(GpBrush PTR, brush)
END CONSTRUCTOR
' ========================================================================================

' =====================================================================================
' * Gets the color of this solid brush.
' =====================================================================================
PRIVATE FUNCTION CGpSolidBrush.GetColor (BYVAL colour AS ARGB PTR) AS GpStatus
   RETURN SetStatus(GdipGetSolidFillColor(cast(GpSolidFill PTR, m_pBrush), colour))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpSolidBrush.GetColor () AS ARGB
   DIM colour AS ARGB
   m_Status = GdipGetSolidFillColor(cast(GpSolidFill PTR, m_pBrush), @colour)
   RETURN colour
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets the color of this solid brush.
' =====================================================================================
PRIVATE FUNCTION CGpSolidBrush.SetColor (BYVAL colour AS ARGB) AS GpStatus
   RETURN SetStatus(GdipSetSolidFillColor(cast(GpSolidFill PTR, m_pBrush), colour))
END FUNCTION
' =====================================================================================

' #####################################################################################
' CGpTextureBrush interface.
' Extends CGpBrush.
' #####################################################################################

' =====================================================================================
' * Creates a TextureBrush object based on an image and a wrap mode. The size of the brush
' defaults to the size of the image, so the entire image is used by the brush.
' =====================================================================================
PRIVATE CONSTRUCTOR CGpTextureBrush (BYVAL pImage AS CGpImage PTR, BYVAL nWrapMode AS WrapMode = WrapModeTile)
   DIM texture AS GpTexture PTR
   m_Status = GdipCreateTexture(IIF(pImage, pImage->m_pImage, NULL), nWrapmode, @texture)
   m_pBrush = cast(GpBrush PTR, texture)
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Creates a TextureBrush object based on an image, a defining set of coordinates, and
' a set of image properties.
' =====================================================================================
PRIVATE CONSTRUCTOR CGpTextureBrush (BYVAL pImage AS CGpImage PTR, BYVAL dstRect AS GpRectF PTR, BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL)
   DIM texture AS GpTexture PTR
   m_Status = GdipCreateTextureIA(IIF(pImage, pImage->m_pImage, NULL), IIF(pImageAttributes, pImageAttributes->m_pImgAttr, NULL), _
              dstRect->x, dstRect->y, dstRect->Width, dstRect->Height, @texture)
   m_pBrush = cast(GpBrush PTR, texture)
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
PRIVATE CONSTRUCTOR CGpTextureBrush (BYVAL pImage AS CGpImage PTR, BYVAL dstRect AS GpRect PTR, BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL)
   DIM texture AS GpTexture PTR
   m_Status = GdipCreateTextureIAI(IIF(pImage, pImage->m_pImage, NULL), IIF(pImageAttributes, pImageAttributes->m_pImgAttr, NULL), _
              dstRect->x, dstRect->y, dstRect->Width, dstRect->Height, @texture)
   m_pBrush = cast(GpBrush PTR, texture)
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Creates a TextureBrush object based on an image, a wrap mode, and a defining set of coordinates.
' =====================================================================================
PRIVATE CONSTRUCTOR CGpTextureBrush (BYVAL pImage AS CGpImage PTR, BYVAL nWrapMode AS WrapMode, BYVAL dstRect AS GpRect PTR)
   DIM texture AS GpTexture PTR
   m_Status = GdipCreateTexture2I(IIF(pImage, pImage->m_pImage, NULL), nWrapMode, dstRect->x, dstRect->y, dstRect->Width, dstRect->Height, @texture)
   m_pBrush = cast(GpBrush PTR, texture)
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
PRIVATE CONSTRUCTOR CGpTextureBrush (BYVAL pImage AS CGpImage PTR, BYVAL nWrapMode AS WrapMode, BYVAL dstRect AS GpRectF PTR)
   DIM texture AS GpTexture PTR
   m_Status = GdipCreateTexture2(IIF(pImage, pImage->m_pImage, NULL), nWrapMode, dstRect->x, dstRect->y, dstRect->Width, dstRect->Height, @texture)
   m_pBrush = cast(GpBrush PTR, texture)
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
PRIVATE CONSTRUCTOR CGpTextureBrush (BYVAL pImage AS CGpImage PTR, BYVAL nWrapMode AS WrapMode, BYVAL dstX AS SINGLE, BYVAL dstY AS SINGLE, BYVAL dstWidth AS SINGLE, BYVAL dstHeight AS SINGLE)
   DIM texture AS GpTexture PTR
   m_Status = GdipCreateTexture2(IIF(pImage, pImage->m_pImage, NULL), nWrapMode, dstX, dstY, dstWidth, dstHeight, @texture)
   m_pBrush = cast(GpBrush PTR, texture)
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
PRIVATE CONSTRUCTOR CGpTextureBrush (BYVAL pImage AS CGpImage PTR, BYVAL nWrapMode AS WrapMode, BYVAL dstX AS INT_, BYVAL dstY AS INT_, BYVAL dstWidth AS INT_, BYVAL dstHeight AS INT_)
   DIM texture AS GpTexture PTR
   m_Status = GdipCreateTexture2I(IIF(pImage, pImage->m_pImage, NULL), nWrapMode, dstX, dstY, dstWidth, dstHeight, @texture)
   m_pBrush = cast(GpBrush PTR, texture)
END CONSTRUCTOR
' =====================================================================================

' =====================================================================================
' * Sets the transformation matrix of this texture brush.
' =====================================================================================
PRIVATE FUNCTION CGpTextureBrush.SetTransform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
   RETURN SetStatus(GdipSetTextureTransform(cast(GpTexture PTR, m_pBrush), IIF(pMatrix, pMatrix->m_pMatrix, NULL)))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the transformation matrix of this texture brush.
' =====================================================================================
PRIVATE FUNCTION CGpTextureBrush.GetTransform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
   RETURN SetStatus(GdipGetTextureTransform(cast(GpTexture PTR, m_pBrush), IIF(pMatrix, pMatrix->m_pMatrix, NULL)))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets the transformation matrix of this texture brush.
' =====================================================================================
PRIVATE FUNCTION CGpTextureBrush.ResetTransform () AS GpStatus
   RETURN SetStatus(GdipResetTextureTransform(cast(GpTexture PTR, m_pBrush)))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Updates this brush's transformation matrix with the product of itself and another matrix.
' =====================================================================================
PRIVATE FUNCTION CGpTextureBrush.MultiplyTransform (BYVAL pMatrix AS CGpMatrix PTR, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   RETURN SetStatus(GdipMultiplyTextureTransform(cast(GpTexture PTR, m_pBrush), IIF(pMatrix, pMatrix->m_pMatrix, NULL), order))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Updates this brush's current transformation matrix with the product of itself and a
' translation matrix.
' =====================================================================================
PRIVATE FUNCTION CGpTextureBrush.TranslateTransform (BYVAL dx AS SINGLE, BYVAL dy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   RETURN SetStatus(GdipTranslateTextureTransform(cast(GpTexture PTR, m_pBrush), dx, dy, order))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Updates this texture brush's current transformation matrix with the product of itself
' and a scaling matrix.
' =====================================================================================
PRIVATE FUNCTION CGpTextureBrush.ScaleTransform (BYVAL sx AS SINGLE, BYVAL sy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   RETURN SetStatus(GdipScaleTextureTransform(cast(GpTexture PTR, m_pBrush), sx, sy, order))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Resets the transformation matrix of this texture brush to the identity matrix. This
' means that no transformation takes place.
' =====================================================================================
PRIVATE FUNCTION CGpTextureBrush.RotateTransform (BYVAL angle AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   RETURN SetStatus(GdipRotateTextureTransform(cast(GpTexture PTR, m_pBrush), angle, order))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets the wrap mode of this texture brush.
' =====================================================================================
PRIVATE FUNCTION CGpTextureBrush.SetWrapMode (BYVAL nWrapMode AS WrapMode) AS GpStatus
   RETURN SetStatus(GdipSetTextureWrapMode(cast(GpTexture PTR, m_pBrush), nWrapMode))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Gets the wrap mode currently set for this texture brush.
' =====================================================================================
PRIVATE FUNCTION CGpTextureBrush.GetWrapMode () AS WrapMode
   DIM nWrapMode AS WrapMode
   m_Status = GdipGetTextureWrapMode(cast(GpTexture PTR, m_pBrush), @nWrapMode)
   RETURN nWrapMode
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Gets a pointer to the Image object that is defined by this texture brush.
' =====================================================================================
PRIVATE FUNCTION CGpTextureBrush.GetImage (BYVAL pImage AS CGpImage PTR) AS GpStatus
   IF pImage = NULL THEN RETURN SetStatus(InvalidParameter)
   IF pImage->m_pImage THEN GdipDisposeImage(pImage->m_pImage)
   RETURN SetStatus(GdipGetTextureImage(cast(GpTexture PTR, m_pBrush), @pImage->m_pImage))
END FUNCTION
' =====================================================================================


' #####################################################################################
' CGpLinearGradientBrush interface.
' Extends CGpBrush.
' #####################################################################################

' =====================================================================================
' Creates a LinearGradientBrush object from a set of boundary points and boundary colors.
' =====================================================================================
PRIVATE CONSTRUCTOR CGpLinearGradientBrush (BYVAL point1 AS GpPointF PTR, BYVAL point2 AS GpPointF PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB)
   DIM brush AS GpLineGradient PTR
   m_Status = GdipCreateLineBrush(point1, point2, color1, color2, WrapModeTile, @brush)
   m_pBrush = cast(GpBrush PTR, brush)
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
PRIVATE CONSTRUCTOR CGpLinearGradientBrush (BYVAL point1 AS GpPoint PTR, BYVAL point2 AS GpPoint PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB)
   DIM brush AS GpLineGradient PTR
   m_Status = GdipCreateLineBrushI(point1, point2, color1, color2, WrapModeTile, @brush)
   m_pBrush = cast(GpBrush PTR, brush)
END CONSTRUCTOR
' =====================================================================================

' =====================================================================================
' Creates a LinearGradientBrush object from a set of boundary points and boundary colors.
' =====================================================================================
PRIVATE CONSTRUCTOR CGpLinearGradientBrush (BYVAL rc AS GpRectF PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, BYVAL nMode AS GpWrapMode)
   DIM brush AS GpLineGradient PTR
   m_Status = GdipCreateLineBrushFromRect(rc, color1, color2, nMode, WrapModeTile, @brush)
   m_pBrush = cast(GpBrush PTR, brush)
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
PRIVATE CONSTRUCTOR CGpLinearGradientBrush (BYVAL rc AS GpRect PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, BYVAL nMode AS GpWrapMode)
   DIM brush AS GpLineGradient PTR
   m_Status = GdipCreateLineBrushFromRectI(rc, color1, color2, nMode, WrapModeTile, @brush)
   m_pBrush = cast(GpBrush PTR, brush)
END CONSTRUCTOR
' =====================================================================================

' =====================================================================================
' Creates a LinearGradientBrush object from a set of boundary points and boundary colors.
' =====================================================================================
PRIVATE CONSTRUCTOR CGpLinearGradientBrush (BYVAL rc AS GpRectF PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, BYVAL angle AS SINGLE, BYVAL isAngleScalable AS BOOL = FALSE)
   DIM brush AS GpLineGradient PTR
   m_Status = GdipCreateLineBrushFromRectWithAngle(rc, color1, color2, angle, isAngleScalable, WrapModeTile, @brush)
   m_pBrush = cast(GpBrush PTR, brush)
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
PRIVATE CONSTRUCTOR CGpLinearGradientBrush (BYVAL rc AS GpRect PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, BYVAL angle AS SINGLE, BYVAL isAngleScalable AS BOOL = FALSE)
   DIM brush AS GpLineGradient PTR
   m_Status = GdipCreateLineBrushFromRectWithAngleI(rc, color1, color2, angle, isAngleScalable, WrapModeTile, @brush)
   m_pBrush = cast(GpBrush PTR, brush)
END CONSTRUCTOR
' =====================================================================================

' =====================================================================================
' Sets the the starting color and ending color of this linear gradient brush.
' =====================================================================================
PRIVATE FUNCTION CGpLinearGradientBrush.SetLinearColors (BYVAL color1 AS ARGB, BYVAL color2 AS ARGB) AS GpStatus
   RETURN SetStatus(GdipSetLineColors(cast(GpLineGradient PTR, m_pBrush), color1, color2))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the the starting color and ending color of this linear gradient brush.
' =====================================================================================
PRIVATE FUNCTION CGpLinearGradientBrush.GetLinearColors (BYVAL colors AS ARGB PTR) AS GpStatus
   RETURN SetStatus(GdipGetLineColors(cast(GpLineGradient PTR, m_pBrush), colors))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the rectangle that defines the boundaries of the gradient.
' =====================================================================================
PRIVATE FUNCTION CGpLinearGradientBrush.GetRectangle (BYVAL rc AS GpRectF PTR) AS GpStatus
   RETURN SetStatus(GdipGetLineRect(cast(GpLineGradient PTR, m_pBrush), rc))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpLinearGradientBrush.GetRectangle (BYVAL rc AS GpRect PTR) AS GpStatus
   RETURN SetStatus(GdipGetLineRectI(cast(GpLineGradient PTR, m_pBrush), rc))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Specifies whether gamma correction is enabled for this linear gradient brush.
' =====================================================================================
PRIVATE FUNCTION CGpLinearGradientBrush.SetGammaCorrection (BYVAL useGammaCorrection AS BOOL) AS GpStatus
   RETURN SetStatus(GdipSetLineGammaCorrection(cast(GpLineGradient PTR, m_pBrush), useGammaCorrection))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Determines whether gamma correction is enabled for this LinearGradientBrush object.
' =====================================================================================
PRIVATE FUNCTION CGpLinearGradientBrush.GetGammaCorrection () AS BOOLEAN
   DIM useGammaCorrection AS BOOL
   m_Status = GdipGetLineGammaCorrection(cast(GpLineGradient PTR, m_pBrush), @useGammaCorrection)
   RETURN useGammaCorrection
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the number of blend factors currently set for this LinearGradientBrush object.
' =====================================================================================
PRIVATE FUNCTION CGpLinearGradientBrush.GetBlendCount () AS INT_
   DIM nCount AS INT_
   m_Status = GdipGetLineBlendCount(cast(GpLineGradient PTR, m_pBrush), @nCount)
   RETURN nCount
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the blend factors and the blend positions of this linear gradient brush to
' create a custom blend.
' =====================================================================================
PRIVATE FUNCTION CGpLinearGradientBrush.SetBlend (BYVAL blendFactors AS SINGLE PTR, BYVAL blendPositions AS SINGLE PTR, BYVAL nCount AS LONG) AS GpStatus
   RETURN SetStatus(GdipSetLineBlend(cast(GpLineGradient PTR, m_pBrush), blendFactors, blendPositions, nCount))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the blend factors and their corresponding blend positions from a LinearGradientBrush object.
' =====================================================================================
PRIVATE FUNCTION CGpLinearGradientBrush.GetBlend (BYVAL blendFactors AS SINGLE PTR, BYVAL blendPositions AS SINGLE PTR, BYVAL nCount AS LONG) AS GpStatus
   RETURN SetStatus(GdipGetLineBlend(cast(GpLineGradient PTR, m_pBrush), blendFactors, blendPositions, nCount))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the number of preset colors currently specified for this path gradient brush.
' =====================================================================================
PRIVATE FUNCTION CGpLinearGradientBrush.GetInterpolationColorCount () AS INT_
   DIM nCount AS INT_
   m_Status = GdipGetLinePresetBlendCount(cast(GpLineGradient PTR, m_pBrush), @nCount)
   RETURN nCount
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the colors to be interpolated for this linear gradient brush and their
' corresponding blend positions.
' =====================================================================================
PRIVATE FUNCTION CGpLinearGradientBrush.SetInterpolationColors (BYVAL presetColors AS ARGB PTR, BYVAL blendPositions AS SINGLE PTR, BYVAL nCount AS LONG) AS GpStatus
   RETURN SetSTatus(GdipSetLinePresetBlend(cast(GpLineGradient PTR, m_pBrush), presetColors, blendPositions, nCount))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the colors currently set to be interpolated for this linear gradient brush and
' their corresponding blend positions.
' =====================================================================================
PRIVATE FUNCTION CGpLinearGradientBrush.GetInterpolationColors (BYVAL presetColors AS ARGB PTR, BYVAL blendPositions AS SINGLE PTR, BYVAL nCount AS LONG) AS GpStatus
   RETURN SetStatus(GdipGetLinePresetBlend(cast(GpLineGradient PTR, m_pBrush), presetColors, blendPositions, nCount))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the blend shape of this linear gradient brush to create a custom blend based on
' a bell-shaped curve. Default scale: 1.0!
' =====================================================================================
PRIVATE FUNCTION CGpLinearGradientBrush.SetBlendBellShape (BYVAL focus AS SINGLE, BYVAL scale AS SINGLE = 1.0) AS GpStatus
   RETURN SetStatus(GdipSetLineSigmaBlend(cast(GpLineGradient PTR, m_pBrush), focus, scale))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the blend shape of this linear gradient brush to create a custom blend based on
' a triangular shape. Default scale: 1.0!
' =====================================================================================
PRIVATE FUNCTION CGpLinearGradientBrush.SetBlendTriangularShape (BYVAL focus AS SINGLE, BYVAL scale AS SINGLE = 1.0) AS GpStatus
   RETURN SetStatus(GdipSetLineLinearBlend(cast(GpLineGradient PTR, m_pBrush), focus, scale))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the transformation matrix of this linear gradient brush.
' =====================================================================================
PRIVATE FUNCTION CGpLinearGradientBrush.SetTransform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
   RETURN SetStatus(GdipSetLineTransform(cast(GpLineGradient PTR, m_pBrush), pMatrix->m_pMatrix))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the transformation matrix of this linear gradient brush.
' =====================================================================================
PRIVATE FUNCTION CGpLinearGradientBrush.GetTransform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
   RETURN SetStatus(GdipGetLineTransform(cast(GpLineGradient PTR, m_pBrush), pMatrix->m_pMatrix))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Resets the transformation matrix of this linear gradient brush to the identity matrix.
' This means that no transformation takes place.
' =====================================================================================
PRIVATE FUNCTION CGpLinearGradientBrush.ResetTransform () AS GpStatus
   RETURN SetStatus(GdipResetLineTransform(cast(GpLineGradient PTR, m_pBrush)))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Updates this brush's transformation matrix with the product of itself and another matrix.
' =====================================================================================
PRIVATE FUNCTION CGpLinearGradientBrush.MultiplyTransform (BYVAL pMatrix AS CGpMatrix PTR, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   RETURN SetStatus(GdipMultiplyLineTransform(cast(GpLineGradient PTR, m_pBrush), pMatrix->m_pMatrix, order))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Updates this brush's current transformation matrix with the product of itself and a
' translation matrix.
' =====================================================================================
PRIVATE FUNCTION CGpLinearGradientBrush.TranslateTransform (BYVAL dx AS SINGLE, BYVAL dy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   RETURN SetStatus(GdipTranslateLineTransform(cast(GpLineGradient PTR, m_pBrush), dx, dy, order))
END FUNCTION
'' =====================================================================================

' =====================================================================================
' Updates this brush's current transformation matrix with the product of itself and a
' scaling matrix.
' =====================================================================================
PRIVATE FUNCTION CGpLinearGradientBrush.ScaleTransform (BYVAL sx AS SINGLE, BYVAL sy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   RETURN SetStatus(GdipScaleLineTransform(cast(GpLineGradient PTR, m_pBrush), sx, sy, order))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Updates this brush's current transformation matrix with the product of itself and a
' rotation matrix.
' =====================================================================================
PRIVATE FUNCTION CGpLinearGradientBrush.RotateTransform (BYVAL angle AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   RETURN SetStatus(GdipRotateLineTransform(cast(GpLineGradient PTR, m_pBrush), angle, order))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the wrap mode of this texture brush.
' =====================================================================================
PRIVATE FUNCTION CGpLinearGradientBrush.SetWrapMode (BYVAL nWrapMode AS WrapMode) AS GpStatus
   RETURN SetStatus(GdipSetLineWrapMode(cast(GpLineGradient PTR, m_pBrush), nWrapMode))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the wrap mode currently set for this texture brush.
' =====================================================================================
PRIVATE FUNCTION CGpLinearGradientBrush.GetWrapMode () AS WrapMode
   DIM nWrapMode AS WrapMode
   m_Status = GdipGetLineWrapMode(cast(GpLineGradient PTR, m_pBrush), @nWrapMode)
   RETURN nWrapMode
END FUNCTION
' =====================================================================================

' #####################################################################################
' CGpPathGradientBrush interface.
' Extends CGpBrush.
' #####################################################################################

' =====================================================================================
' Creates a PathGradient object.
' =====================================================================================
PRIVATE CONSTRUCTOR CGpPathGradientBrush (BYVAL pts AS GpPointF PTR, BYVAL nCount AS LONG, BYVAL nWrapMode AS WrapMode = WrapModeClamp)
   DIM brush AS GpPathGradient PTR
   m_Status = GdipCreatePathGradient(pts, nCount, nWrapMode, @brush)
   m_pBrush = cast(GpBrush PTR, brush)
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
PRIVATE CONSTRUCTOR CGpPathGradientBrush (BYVAL pts AS GpPoint PTR, BYVAL nCount AS LONG, BYVAL nWrapMode AS WrapMode = WrapModeClamp)
   DIM brush AS GpPathGradient PTR
   m_Status = GdipCreatePathGradientI(pts, nCount, nWrapMode, @brush)
   m_pBrush = cast(GpBrush PTR, brush)
END CONSTRUCTOR
' =====================================================================================

' =====================================================================================
' Creates a PathGradient object.
' =====================================================================================
PRIVATE CONSTRUCTOR CGpPathGradientBrush (BYVAL pGraphPath AS CGpGraphicsPath PTR)
   DIM brush AS GpPathGradient PTR
   m_Status = GdipCreatePathGradientFromPath(pGraphPath->m_pPath, @brush)
   m_pBrush = cast(GpBrush PTR, brush)
END CONSTRUCTOR
' =====================================================================================

' =====================================================================================
' Gets the color of the center point of this path gradient brush.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.GetCenterColor (BYVAL colour AS ARGB PTR) AS GpStatus
   RETURN SetStatus(GdipGetPathGradientCenterColor(cast(GpPathGradient PTR, m_pBrush), colour))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the color of the center point of this path gradient brush.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.SetCenterColor (BYVAL colour AS ARGB) AS GpStatus
   RETURN SetStatus(GdipSetPathGradientCenterColor(cast(GpPathGradient PTR, m_pBrush), colour))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the number of points in the array of points that defines this brush's boundary path.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.GetPointCount () AS INT_
   DIM nCount AS INT_
   m_Status = GdipGetPathGradientPointCount(cast(GpPathGradient PTR, m_pBrush), @nCount)
   RETURN nCount
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the number of colors that have been specified for the boundary path of this path
' gradient brush.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.GetSurroundColorCount () AS INT_
   DIM nCount AS INT_
   m_Status = GdipGetPathGradientSurroundColorCount(cast(GpPathGradient PTR, m_pBrush), @nCount)
   RETURN nCount
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the surround colors currently specified for this path gradient brush.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.GetSurroundColors (BYVAL colors AS ARGB PTR, BYVAL nCount AS INT_ PTR) AS GpStatus
   IF colors = NULL OR *nCount = 0 THEN RETURN SetStatus(InvalidParameter)
   DIM nCount1 AS INT_
   m_Status = GdipGetPathGradientSurroundColorCount(cast(GpPathGradient PTR, m_pBrush), @nCount1)
   IF m_Status <> OK THEN RETURN m_Status
   IF *nCount < nCount1 OR nCount1 <= 0 THEN RETURN SetStatus(InsufficientBuffer)
   m_Status = GdipGetPathGradientSurroundColorsWithCount(cast(GpPathGradient PTR, m_pBrush), colors, @nCount1)
   IF m_Status = Ok THEN *nCount = nCount1
   RETURN m_Status
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the surround colors currently specified for this path gradient brush.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.SetSurroundColors (BYVAL colors AS ARGB PTR, BYVAL nCount AS INT_ PTR) AS GpStatus
   IF colors = NULL OR *nCount = 0 THEN RETURN SetStatus(InvalidParameter)
   DIM nCount1 AS INT_
   nCount1 = this.GetPointCount
   IF *nCount > nCount1 OR nCount1 <= 0 THEN RETURN SetStatus(InvalidParameter)
   m_Status = GdipSetPathGradientSurroundColorsWithCount(cast(GpPathGradient PTR, m_pBrush), colors, @nCount1)
   IF m_Status = Ok THEN *nCount = nCount1
   RETURN m_Status
END FUNCTION
' =====================================================================================

' =====================================================================================
' Not implemented in Windows GDI+ version 1.0.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.GetGraphicsPath (BYVAL pGraphPath AS CGpGraphicsPath PTR) AS GpStatus
   IF pGraphPath = NULL THEN RETURN SetStatus(InvalidParameter)
   RETURN SetStatus(GdipGetPathGradientPath(cast(GpPathGradient PTR, m_pBrush), pGraphPath->m_pPath))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Not implemented in Windows GDI+ version 1.0.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.SetGraphicsPath (BYVAL pGraphPath AS CGpGraphicsPath PTR) AS GpStatus
   IF pGraphPath = NULL THEN RETURN SetStatus(InvalidParameter)
   RETURN SetStatus(GdipGetPathGradientPath(cast(GpPathGradient PTR, m_pBrush), pGraphPath->m_pPath))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the center point of this path gradient brush.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.GetCenterPoint (BYVAL pt AS GpPointF PTR) AS GpStatus
   RETURN SetStatus(GdipGetPathGradientCenterPoint(cast(GpPathGradient PTR, m_pBrush), pt))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.GetCenterPoint (BYVAL pt AS GpPoint PTR) AS GpStatus
   RETURN SetStatus(GdipGetPathGradientCenterPointI(cast(GpPathGradient PTR, m_pBrush), pt))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the center point of this path gradient brush.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.SetCenterPoint (BYVAL pt AS GpPointF PTR) AS GpStatus
   RETURN SetStatus(GdipSetPathGradientCenterPoint(cast(GpPathGradient PTR, m_pBrush), pt))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.SetCenterPoint (BYVAL pt AS GpPoint PTR) AS GpStatus
   RETURN SetStatus(GdipSetPathGradientCenterPointI(cast(GpPathGradient PTR, m_pBrush), pt))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the smallest rectangle that encloses the boundary path of this path gradient brush.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.GetRectangle (BYVAL rc AS GpRectF PTR) AS GpStatus
   RETURN SetStatus(GdipGetPathGradientRect(cast(GpPathGradient PTR, m_pBrush), rc))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.GetRectangle (BYVAL rc AS GpRect PTR) AS GpStatus
   RETURN SetStatus(GdipGetPathGradientRectI(cast(GpPathGradient PTR, m_pBrush), rc))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Specifies whether gamma correction is enabled for this path gradient brush.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.SetGammaCorrection (BYVAL useGammaCorrection AS BOOL) AS GpStatus
   RETURN SetStatus(GdipSetPathGradientGammaCorrection(cast(GpPathGradient PTR, m_pBrush), useGammaCorrection))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Specifies whether gamma correction is enabled for this path gradient brush.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.GetGammaCorrection () AS BOOLEAN
   DIM useGammaCorrection AS BOOL
   m_Status = GdipGetPathGradientGammaCorrection(cast(GpPathGradient PTR, m_pBrush), @useGammaCorrection)
   RETURN useGammaCorrection
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the number of blend factors currently set for this path gradient brush.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.GetBlendCount () AS INT_
   DIM nCount AS INT_
   m_Status = GdipGetPathGradientBlendCount(cast(GpPathGradient PTR, m_pBrush), @nCount)
   RETURN nCount
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the blend factors and the corresponding blend positions currently set for this
' path gradient brush.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.GetBlend (BYVAL blendFactors AS SINGLE PTR, BYVAL blendPositions AS SINGLE PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipGetPathGradientBlend(cast(GpPathGradient PTR, m_pBrush), blendFactors, blendPositions, nCount))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the blend factors and the corresponding blend positions currently set for this
' path gradient brush.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.SetBlend (BYVAL blendFactors AS SINGLE PTR, BYVAL blendPositions AS SINGLE PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipSetPathGradientBlend(cast(GpPathGradient PTR, m_pBrush), blendFactors, blendPositions, nCount))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the number of preset colors currently specified for this path gradient brush.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.GetInterpolationColorCount () AS INT_
   DIM nCount AS INT_
   m_Status = GdipGetPathGradientPresetBlendCount(cast(GpPathGradient PTR, m_pBrush), @nCount)
   RETURN nCount
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the preset colors and the blend positions of this path gradient brush.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.SetInterpolationColors (BYVAL presetColors AS ARGB PTR, BYVAL blendPositions AS SINGLE PTR, BYVAL nCount AS INT_) AS GpStatus
   IF VARPTR(presetColors) = NULL OR nCount <= 0 THEN RETURN SetStatus(InvalidParameter)
   RETURN SetStatus(GdipSetPathGradientPresetBlend(cast(GpPathGradient PTR, m_pBrush), presetColors, blendPositions, nCount))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the number of preset colors currently specified for this path gradient brush.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.GetInterpolationColors (BYVAL presetColors AS ARGB PTR, BYVAL blendPositions AS SINGLE PTR, BYVAL nCount AS INT_) AS GpStatus
   IF VARPTR(presetColors) = NULL OR nCount <= 0 THEN RETURN SetStatus(InvalidParameter)
   RETURN SetStatus(GdipGetPathGradientPresetBlend(cast(GpPathGradient PTR, m_pBrush), presetColors, blendPositions, nCount))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the blend shape of this path gradient brush. Default scale: 1.0!
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.SetBlendBellShape (BYVAL focus AS SINGLE, BYVAL scale AS SINGLE = 1.0) AS GpStatus
   RETURN SetStatus(GdipSetPathGradientSigmaBlend(cast(GpPathGradient PTR, m_pBrush), focus, scale))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the blend shape of this path gradient brush. Default scale: 1.0!
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.SetBlendTriangularShape (BYVAL focus AS SINGLE, BYVAL scale AS SINGLE = 1.0) AS GpStatus
   RETURN SetStatus(GdipSetPathGradientLinearBlend(cast(GpPathGradient PTR, m_pBrush), focus, scale))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the transformation matrix of this texture brush.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.GetTransform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
   RETURN SetStatus(GdipGetPathGradientTransform(cast(GpPathGradient PTR, m_pBrush), pMatrix->m_pMatrix))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the transformation matrix of this path gradient brush.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.SetTransform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
   RETURN SetStatus(GdipSetPathGradientTransform(cast(GpPathGradient PTR, m_pBrush), pMatrix->m_pMatrix))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the transformation matrix of this texture brush.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.ResetTransform () AS GpStatus
   RETURN SetStatus(GdipResetPathGradientTransform(cast(GpPathGradient PTR, m_pBrush)))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Updates this brush's transformation matrix with the product of itself and another matrix.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.MultiplyTransform (BYVAL pMatrix AS CGpMatrix PTR, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   RETURN SetStatus(GdipMultiplyPathGradientTransform(cast(GpPathGradient PTR, m_pBrush), pMatrix->m_pMatrix, order))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Updates this brush's current transformation matrix with the product of itself and a
' translation matrix.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.TranslateTransform (BYVAL dx AS SINGLE, BYVAL dy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   RETURN SetStatus(GdipTranslatePathGradientTransform(cast(GpPathGradient PTR, m_pBrush), dx, dy, order))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Updates this texture brush's current transformation matrix with the product of itself
' and a scaling matrix.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.ScaleTransform (BYVAL sx AS SINGLE, BYVAL sy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   RETURN SetStatus(GdipScalePathGradientTransform(cast(GpPathGradient PTR, m_pBrush), sx, sy, order))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Resets the transformation matrix of this texture brush to the identity matrix. This
' means that no transformation takes place.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.RotateTransform (BYVAL angle AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   RETURN SetStatus(GdipRotatePathGradientTransform(cast(GpPathGradient PTR, m_pBrush), angle, order))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the focus scales of this path gradient brush.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.GetFocusScales (BYVAL xScale AS SINGLE PTR, BYVAL yScale AS SINGLE PTR) AS GpStatus
   RETURN GdipGetPathGradientFocusScales(cast(GpPathGradient PTR, m_pBrush), xScale, yScale)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the focus scales of this path gradient brush.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.SetFocusScales (BYVAL xScale AS SINGLE, BYVAL yScale AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipSetPathGradientFocusScales(cast(GpPathGradient PTR, m_pBrush), xScale, yScale))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the wrap mode currently set for this texture brush.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.GetWrapMode () AS WrapMode
   DIM nWrapMode AS WrapMode
   m_Status = GdipGetPathGradientWrapMode(cast(GpPathGradient PTR, m_pBrush), @nWrapMode)
   FUNCTION = nWrapMode
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the wrap mode of this texture brush.
' =====================================================================================
PRIVATE FUNCTION CGpPathGradientBrush.SetWrapMode (BYVAL nWrapMode AS WrapMode) AS GpStatus
   RETURN SetStatus(GdipSetPathGradientWrapMode(cast(GpPathGradient PTR, m_pBrush), nWrapMode))
END FUNCTION
' =====================================================================================

' #####################################################################################
' CGpHatchBrush interface.
' Inherits from CGpBrush.
' #####################################################################################

' =====================================================================================
' Creates a HatchBrush object based on a hatch style, a foreground color, and a background color.
' =====================================================================================
PRIVATE CONSTRUCTOR CGpHatchBrush (BYVAL nHatchStyle AS HatchStyle, BYVAL foreColor AS ARGB, BYVAL backColor AS ARGB = &HFF000000)
   DIM brush AS GpHatch PTR
   m_Status = GdipCreateHatchBrush(nHatchStyle, foreColor, backColor, @brush)
   m_pBrush = cast(GpBrush PTR, brush)
END CONSTRUCTOR
' =====================================================================================

' =====================================================================================
' Gets the hatch style of this hatch brush.
' =====================================================================================
PRIVATE FUNCTION CGpHatchBrush.GetHatchStyle () AS HatchStyle
   DIM nHatchStyle AS HatchStyle
   m_Status = GdipGetHatchStyle(cast(GpHatch PTR, m_pBrush), @nHatchStyle)
   FUNCTION = nHatchStyle
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the foreground color of this hatch brush.
' =====================================================================================
PRIVATE FUNCTION CGpHatchBrush.GetForegroundColor (BYVAL colour AS ARGB PTR) AS GpStatus
   RETURN SetStatus(GdipGetHatchForegroundColor(cast(GpHatch PTR, m_pBrush), colour))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the background color of this hatch brush.
' =====================================================================================
PRIVATE FUNCTION CGpHatchBrush.GetBackgroundColor (BYVAL colour AS ARGB PTR) AS GpStatus
   RETURN SetStatus(GdipGetHatchBackgroundColor(cast(GpHatch PTR, m_pBrush), colour))
END FUNCTION
' =====================================================================================
