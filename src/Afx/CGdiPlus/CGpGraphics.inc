' ########################################################################################
' Microsoft Windows
' File: CGpGraphics.inc
' Contents: GDI+ graphics class.
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

' ########################################################################################
' CGpGraphics Class
' ########################################################################################

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CGpGraphics
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' * Creates a Graphics object that is associated with a specified device context.
' When you use these constructors to create a Graphics object, make sure that the Graphics
' object is deleted or goes out of scope before the device context is released.
' ========================================================================================
PRIVATE CONSTRUCTOR CGpGraphics (BYVAL hdc AS HDC)
   m_Status = GdipCreateFromHDC(hdc, @m_pGraphics)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CGpGraphics (BYVAL hdc AS HDC, BYVAL hDevice AS HANDLE)
   m_Status = GdipCreateFromHDC2(hdc, hDevice, @m_pGraphics)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a Graphics object that is associated with a specified window.
' ========================================================================================
PRIVATE CONSTRUCTOR CGpGraphics (BYVAL hwnd AS HWND, BYVAL icm AS BOOLEAN = FALSE)
   IF icm THEN
      m_Status = GdipCreateFromHWNDICM(hwnd, @m_pGraphics)
   ELSE
      m_Status = GdipCreateFromHWND(hwnd, @m_pGraphics)
   END IF
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a Graphics object that is associated with an Image object.
' ========================================================================================
PRIVATE CONSTRUCTOR CGpGraphics (BYVAL pImage AS CGpImage PTR)
   m_Status = GdipGetImageGraphicsContext(pImage->m_pImage, @m_pGraphics)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR CGpGraphics
   IF m_pGraphics THEN GdipDeleteGraphics(m_pGraphics)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a Graphics object that is associated with a specified device context.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.FromHDC (BYVAL hdc AS HDC) AS GpStatus
   RETURN SetStatus(GdipCreateFromHDC(hdc, @m_pGraphics))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.FromHDC (BYVAL hdc AS HDC, BYVAL hDevice AS HANDLE) AS GpStatus
   RETURN SetStatus(GdipCreateFromHDC2(hdc, hDevice, @m_pGraphics))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.FromHWND (BYVAL hwnd AS HWND, BYVAL icm AS BOOLEAN = FALSE) AS GpStatus
   IF icm THEN
      RETURN SetStatus(GdipCreateFromHWNDICM(hwnd, @m_pGraphics))
   ELSE
      RETURN SetStatus(GdipCreateFromHWND(hwnd, @m_pGraphics))
   END IF
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.FromImage (BYVAL pImage AS CGpImage PTR) AS GpStatus
   RETURN SetStatus(GdipGetImageGraphicsContext(pImage->m_pImage, @m_pGraphics))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Flushes all pending graphics operations.
' ========================================================================================
PRIVATE SUB CGpGraphics.Flush (BYVAL intention AS FlushIntention = FlushIntentionFlush)
   GdipFlush(m_pGraphics, intention)
END SUB
' ========================================================================================

' ========================================================================================
' Gets a handle to the device context associated with this Graphics object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.GetHDC () AS HDC
   DIM hdc AS HDC
   SetStatus(GdipGetDC(m_pGraphics, @hdc))
   RETURN hdc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Releases a device context handle obtained by a previous call to the GetHDC FUNCTION of
' this Graphics object.
' ========================================================================================
PRIVATE SUB CGpGraphics.ReleaseHDC (BYVAL hdc AS HDC)
   GdipReleaseDC(m_pGraphics, hdc)
END SUB
' ========================================================================================

' ========================================================================================
' Sets the rendering origin of this Graphics object. The rendering origin is used to
' set the dither origin for 8-bits-per-pixel and 16-bits-per-pixel dithering and is
' also used to set the origin for hatch brushes.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.SetRenderingOrigin (BYVAL x AS INT_, BYVAL y AS INT_) AS GpStatus
   RETURN SetStatus(GdipSetRenderingOrigin(m_pGraphics, x, y))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the rendering origin currently set for this Graphics object. The rendering origin
' is used to set the dither origin for 8-bits per pixel and 16-bits per pixel dithering
' and is also used to set the origin for hatch brushes.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.GetRenderingOrigin (BYVAL x AS LONG PTR, BYVAL y AS LONG PTR) AS GpStatus
   RETURN SetStatus(GdipGetRenderingOrigin(m_pGraphics, x, y))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the compositing mode of this Graphics object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.SetCompositingMode (BYVAL nCompositingMode AS CompositingMode) AS GpStatus
   RETURN SetStatus(GdipSetCompositingMode(m_pGraphics, nCompositingMode))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the compositing mode currently set for this Graphics object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.GetCompositingMode () AS CompositingMode
   DIM nCompositingMode AS CompositingMode
   m_Status = GdipGetCompositingMode(m_pGraphics, @nCompositingMode)
   RETURN nCompositingMode
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the compositing quality of this Graphics object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.SetCompositingQuality (BYVAL nQuality AS CompositingQuality) AS GpStatus
   RETURN SetStatus(GdipSetCompositingQuality(m_pGraphics, nQuality))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the compositing mode currently set for this Graphics object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.GetCompositingQuality () AS CompositingQuality
   DIM nQuality AS CompositingQuality
   m_Status = GdipGetCompositingQuality(m_pGraphics, @nQuality)
   RETURN nQuality
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text rendering mode of this Graphics object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.SetTextRenderingHint (BYVAL newMode AS TextRenderingHint) AS GpStatus
   RETURN SetStatus(GdipSetTextRenderingHint(m_pGraphics, newMode))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the compositing mode currently set for this Graphics object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.GetTextRenderingHint () AS TextRenderingHint
   DIM hint AS TextRenderingHint
   m_Status = GdipGetTextRenderingHint(m_pGraphics, @hint)
   RETURN hint
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text rendering mode of this Graphics object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.SetTextContrast (BYVAL contrast AS UINT) AS GpStatus
   RETURN SetStatus(GdipSetTextContrast(m_pGraphics, contrast))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the contrast value currently set for this Graphics object. The contrast value is
' used for antialiasing text.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.GetTextContrast () AS UINT
   DIM contrast AS UINT
   m_Status = GdipGetTextContrast(m_pGraphics, @contrast)
   RETURN contrast
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the compositing mode currently set for this Graphics object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.GetInterpolationMode () AS InterpolationMode
   DIM nMode AS InterpolationMode
   m_Status = GdipGetInterpolationMode(m_pGraphics, @nMode)
   RETURN nMode
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the interpolation mode of this Graphics object. The interpolation mode determines
' the algorithm that is used when images are scaled or rotated.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.SetInterpolationMode (BYVAL interpolationMode AS LONG) AS GpStatus
   RETURN SetStatus(GdipSetInterpolationMode(m_pGraphics, interpolationMode))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether smoothing (antialiasing) is applied to the Graphics object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.GetSmoothingMode () AS SmoothingMode
   DIM nMode AS SmoothingMode
   m_Status = GdipGetSmoothingMode(m_pGraphics, @nMode)
   FUNCTION = nMode
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the rendering quality of the Graphics object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.SetSmoothingMode (BYVAL smoothingMode AS LONG) AS GpStatus
   RETURN SetStatus(GdipSetSmoothingMode(m_pGraphics, smoothingMode))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the pixel offset mode currently set for this Graphics object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.GetPixelOffsetMode () AS PixelOffsetMode
   DIM nMode AS PixelOffsetMode
   m_Status = GdipGetPixelOffsetMode(m_pGraphics, @nMode)
   RETURN nMode
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the pixel offset mode of this Graphics object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.SetPixelOffsetMode (BYVAL nMode AS PixelOffsetMode) AS GpStatus
   RETURN SetStatus(GdipSetPixelOffsetMode(m_pGraphics, nMode))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the transformation matrix of this Graphics object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.SetTransform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
   RETURN SetStatus(GdipSetWorldTransform(m_pGraphics, pMatrix->m_pMatrix))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the transformation matrix of this Graphics object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.ResetTransform () AS GpStatus
   RETURN SetStatus(GdipResetWorldTransform(m_pGraphics))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Updates this Graphic object's transformation matrix with the product of itself and another matrix.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.MultiplyTransform (BYVAL pMatrix AS CGpMatrix PTR, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   RETURN SetStatus(GdipMultiplyWorldTransform(m_pGraphics, pMatrix->m_pMatrix, order))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Updates this Graphic object's current transformation matrix with the product of itself
' and a translation matrix.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.TranslateTransform (BYVAL dx AS SINGLE, BYVAL dy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   RETURN SetStatus(GdipTranslateWorldTransform(m_pGraphics, dx, dy, order))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Updates this Graphic object's current transformation matrix with the product of itself
' and a scaling matrix.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.ScaleTransform (BYVAL sx AS SINGLE, BYVAL sy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   RETURN SetStatus(GdipScaleWorldTransform(m_pGraphics, sx, sy, order))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Resets the transformation matrix of this Graphic o0bject to the identity matrix. This
' means that no transformation takes place.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.RotateTransform (BYVAL angle AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   RETURN SetStatus(GdipRotateWorldTransform(m_pGraphics, angle, order))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the transformation matrix of this Graphic object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.GetTransform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
   RETURN SetStatus(GdipGetWorldTransform(m_pGraphics, pMatrix->m_pMatrix))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the unit of measure for this Graphics object. The page unit belongs to the page
' transformation, which converts page coordinates to device coordinates.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.SetPageUnit (BYVAL unit AS GpUnit) AS GpStatus
   RETURN SetStatus(GdipSetPageUnit(m_pGraphics, unit))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets scaling factor for the page transformation of this Graphics object. The page
' transformation converts page coordinates to device coordinates.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.SetPageScale (BYVAL nScale AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipSetPageScale(m_pGraphics, nScale))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the unit of measure currently set for this Graphics object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.GetPageUnit () AS GpUnit
   DIM unit_ AS GpUnit
   m_Status = GdipGetPageUnit(m_pGraphics, @unit_)
   RETURN unit_
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets scaling factor currently set for the page transformation of this Graphics object.
' The page transformation converts page coordinates to device coordinates.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.GetPageScale () AS SINGLE
   DIM nScale AS SINGLE
   m_Status = GdipGetPageScale(m_pGraphics, @nScale)
   RETURN nScale
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Gets the horizontal resolution, in dots per inch, of the display device associated
' with this Graphics object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.GetDpiX () AS SINGLE
   DIM dpi AS SINGLE
   m_Status = GdipGetDpiX(m_pGraphics, @dpi)
   RETURN dpi
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Gets the vertical resolution, in dots per inch, of the display device associated
' with this Graphics object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.GetDpiY () AS SINGLE
   DIM dpi AS SINGLE
   m_Status = GdipGetDpiY(m_pGraphics, @dpi)
   RETURN dpi
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts an array of points from one coordinate space to another. The conversion is
' based on the current world and page transformations of this Graphics object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.TransformPoints (BYVAL destSpace AS CoordinateSpace, BYVAL srcSpace AS CoordinateSpace, BYVAL pts AS GpPointF PTR, BYVAL nCount AS LONG) AS GpStatus
   RETURN SetStatus(GdipTransformPoints(m_pGraphics, destSpace, srcSpace, pts, nCount))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts an array of points from one coordinate space to another. The conversion is
' based on the current world and page transformations of this Graphics object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.TransformPoints (BYVAL destSpace AS CoordinateSpace, BYVAL srcSpace AS CoordinateSpace, BYVAL pts AS GpPoint PTR, BYVAL nCount AS LONG) AS GpStatus
   RETURN SetStatus(GdipTransformPointsI(m_pGraphics, destSpace, srcSpace, pts, nCount))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the nearest color to the color that is passed in. This FUNCTION works on 8-bits per
' pixel or lower display devices for which there is an 8-bit color palette.
' Note: Alpha is ignored.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.GetNearestColor (BYVAL colour AS ARGB PTR) AS GpStatus
   RETURN SetStatus(GdipGetNearestColor(m_pGraphics, colour))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Draws a line that connects two points.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawLine (BYVAL pPen AS CGpPen PTR, BYVAL x1 AS SINGLE, BYVAL y1 AS SINGLE, BYVAL x2 AS SINGLE, BYVAL y2 AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipDrawLine(m_pGraphics, pPen->m_pPen, x1, y1, x2, y2))
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Draws a line that connects two points.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawLine (BYVAL pPen AS CGpPen PTR, BYVAL x1 AS INT_, BYVAL y1 AS INT_, BYVAL x2 AS INT_, BYVAL y2 AS INT_) AS GpStatus
   RETURN SetStatus(GdipDrawLineI(m_pGraphics, pPen->m_pPen, x1, y1, x2, y2))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawLine (BYVAL pPen AS CGpPen PTR, BYVAL pt1 AS GpPointF PTR, BYVAL pt2 AS GpPointF PTR) AS GpStatus
   RETURN SetStatus(GdipDrawLine(m_pGraphics, pPen->m_pPen, pt1->X, pt1->Y, pt2->X, pt2->Y))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawLine (BYVAL pPen AS CGpPen PTR, BYVAL pt1 AS GpPoint PTR, BYVAL pt2 AS GpPoint PTR) AS GpStatus
   RETURN SetStatus(GdipDrawLineI(m_pGraphics, pPen->m_pPen, pt1->X, pt1->Y, pt2->X, pt2->Y))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Draws a sequence of connected lines.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawLines (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, BYVAL nCount AS LONG) AS GpStatus
   RETURN SetStatus(GdipDrawLines(m_pGraphics, pPen->m_pPen, pts, nCount))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawLines (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPoint PTR, BYVAL nCount AS LONG) AS GpStatus
   RETURN SetStatus(GdipDrawLinesI(m_pGraphics, pPen->m_pPen, pts, nCount))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Draws an arc. The arc is part of an ellipse.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawArc (BYVAL pPen AS CGpPen PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipDrawArc(m_pGraphics, pPen->m_pPen, x, y, nWidth, nHeight, startAngle, sweepAngle))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawArc (BYVAL pPen AS CGpPen PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipDrawArcI(m_pGraphics, pPen->m_pPen, x, y, nWidth, nHeight, startAngle, sweepAngle))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawArc (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRectF PTR, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipDrawArc(m_pGraphics, pPen->m_pPen, rc->x, rc->y, rc->Width, rc->Height, startAngle, sweepAngle))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawArc (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRect PTR, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipDrawArcI(m_pGraphics, pPen->m_pPen, rc->x, rc->y, rc->Width, rc->Height, startAngle, sweepAngle))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Draws a Bézier spline.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawBezier (BYVAL pPen AS CGpPen PTR, BYVAL x1 AS SINGLE, BYVAL y1 AS SINGLE, _
   BYVAL x2 AS SINGLE, BYVAL y2 AS SINGLE, BYVAL x3 AS SINGLE, BYVAL y3 AS SINGLE, _
   BYVAL x4 AS SINGLE, BYVAL y4 AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipDrawBezier(m_pGraphics, pPen->m_pPen, x1, y1, x2, y2, x3, y3, x4, y4))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawBezier (BYVAL pPen AS CGpPen PTR, BYVAL x1 AS INT_, BYVAL y1 AS INT_, _
   BYVAL x2 AS INT_, BYVAL y2 AS INT_, BYVAL x3 AS INT_, BYVAL y3 AS INT_, _
   BYVAL x4 AS INT_, BYVAL y4 AS INT_) AS GpStatus
   RETURN SetStatus(GdipDrawBezierI(m_pGraphics, pPen->m_pPen, x1, y1, x2, y2, x3, y3, x4, y4))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawBezier (BYVAL pPen AS CGpPen PTR, BYVAL pt1 AS GpPointF PTR, BYVAL pt2 AS GpPointF PTR, BYVAL pt3 AS GpPointF PTR, BYVAL pt4 AS GpPointF PTR) AS GpStatus
   RETURN SetStatus(GdipDrawBezier(m_pGraphics, pPen->m_pPen, pt1->x, pt1->y, pt2->x, pt2->y, pt3->x, pt3->y, pt4->x, pt4->y))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawBezier (BYVAL pPen AS CGpPen PTR, BYVAL pt1 AS GpPoint PTR, BYVAL pt2 AS GpPoint PTR, BYVAL pt3 AS GpPoint PTR, BYVAL pt4 AS GpPoint PTR) AS GpStatus
   RETURN SetStatus(GdipDrawBezierI(m_pGraphics, pPen->m_pPen, pt1->x, pt1->y, pt2->x, pt2->y, pt3->x, pt3->y, pt4->x, pt4->y))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Draws a Bézier spline.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawBeziers (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipDrawBeziers(m_pGraphics, pPen->m_pPen, pts, nCount))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawBeziers (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipDrawBeziersI(m_pGraphics, pPen->m_pPen, pts, nCount))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Draws a rectangle.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawRectangle (BYVAL pPen AS CGpPen PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipDrawRectangle(m_pGraphics, pPen->m_pPen, x, y, nWidth, nHeight))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawRectangle (BYVAL pPen AS CGpPen PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS GpStatus
   RETURN SetStatus(GdipDrawRectangleI(m_pGraphics, pPen->m_pPen, x, y, nWidth, nHeight))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawRectangle (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRectF PTR) AS GpStatus
   RETURN SetStatus(GdipDrawRectangle(m_pGraphics, pPen->m_pPen, rc->x, rc->y, rc->Width, rc->Height))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawRectangle (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRect PTR) AS GpStatus
   RETURN SetStatus(GdipDrawRectangleI(m_pGraphics, pPen->m_pPen, rc->x, rc->y, rc->Width, rc->Height))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Draws a sequence of rectangles.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawRectangles (BYVAL pPen AS CGpPen PTR, BYVAL rects AS GpRectF PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipDrawRectangles(m_pGraphics, pPen->m_pPen, rects, nCount))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawRectangles (BYVAL pPen AS CGpPen PTR, BYVAL rects AS GpRect PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipDrawRectanglesI(m_pGraphics, pPen->m_pPen, rects, nCount))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Draws an ellipse.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawEllipse (BYVAL pPen AS CGpPen PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipDrawEllipse(m_pGraphics, pPen->m_pPen, x, y, nWidth, nHeight))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawEllipse (BYVAL pPen AS CGpPen PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS GpStatus
   RETURN SetStatus(GdipDrawEllipseI(m_pGraphics, pPen->m_pPen, x, y, nWidth, nHeight))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawEllipse (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRectF PTR) AS GpStatus
   RETURN SetStatus(GdipDrawEllipse(m_pGraphics, pPen->m_pPen, rc->x, rc->y, rc->Width, rc->Height))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawEllipse (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRect PTR) AS GpStatus
   RETURN SetStatus(GdipDrawEllipseI(m_pGraphics, pPen->m_pPen, rc->x, rc->y, rc->Width, rc->Height))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Draws a pie.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawPie (BYVAL pPen AS CGpPen PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipDrawPie(m_pGraphics, pPen->m_pPen, x, y, nWidth, nHeight, startAngle, sweepAngle))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawPie (BYVAL pPen AS CGpPen PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipDrawPieI(m_pGraphics, pPen->m_pPen, x, y, nWidth, nHeight, startAngle, sweepAngle))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawPie (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRectF PTR, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipDrawPie(m_pGraphics, pPen->m_pPen, rc->x, rc->y, rc->Width, rc->Height, startAngle, sweepAngle))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawPie (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRect PTR, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipDrawPieI(m_pGraphics, pPen->m_pPen, rc->x, rc->y, rc->Width, rc->Height, startAngle, sweepAngle))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Draws a polygon.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawPolygon (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, BYVAL nCount AS LONG) AS GpStatus
   RETURN SetStatus(GdipDrawPolygon(m_pGraphics, pPen->m_pPen, pts, nCount))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawPolygon (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPoint PTR, BYVAL nCount AS LONG) AS GpStatus
   RETURN SetStatus(GdipDrawPolygonI(m_pGraphics, pPen->m_pPen, pts, nCount))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Draws a path.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawPath (BYVAL pPen AS CGpPen PTR, BYVAL pPath AS CGpGraphicsPath PTR) AS GpStatus
   RETURN SetStatus(GdipDrawPath(m_pGraphics, pPen->m_pPen, pPath->m_pPath))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Draws a cardinal spline.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipDrawCurve(m_pGraphics, pPen->m_pPen, pts, nCount))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipDrawCurveI(m_pGraphics, pPen->m_pPen, pts, nCount))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_, BYVAL tension AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipDrawCurve2(m_pGraphics, pPen->m_pPen, pts, nCount, tension))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_, BYVAL tension AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipDrawCurve2I(m_pGraphics, pPen->m_pPen, pts, nCount, tension))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_, BYVAL offset AS INT_, BYVAL numberOfSegments AS INT_, BYVAL tension AS SINGLE = 0.5) AS GpStatus
   RETURN SetStatus(GdipDrawCurve3(m_pGraphics, pPen->m_pPen, pts, nCount, offset, numberOfSegments, tension))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_, BYVAL offset AS INT_, BYVAL numberOfSegments AS INT_, BYVAL tension AS SINGLE = 0.5) AS GpStatus
   RETURN SetStatus(GdipDrawCurve3I(m_pGraphics, pPen->m_pPen, pts, nCount, offset, numberOfSegments, tension))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Draws a a closed cardinal spline.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawClosedCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipDrawClosedCurve(m_pGraphics, pPen->m_pPen, pts, nCount))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawClosedCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipDrawClosedCurveI(m_pGraphics, pPen->m_pPen, pts, nCount))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawClosedCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_, BYVAL tension AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipDrawClosedCurve2(m_pGraphics, pPen->m_pPen, pts, nCount, tension))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawClosedCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_, BYVAL tension AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipDrawClosedCurve2I(m_pGraphics, pPen->m_pPen, pts, nCount, tension))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Clears a Graphics object to a specified color.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.Clear (BYVAL colour AS ARGB) AS GpStatus
   RETURN SetStatus(GdipGraphicsClear(m_pGraphics, colour))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Uses a brush to fill the interior of a rectangle.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.FillRectangle (BYVAL pBrush AS CGpBrush PTR, BYVAL rc AS GpRectF PTR) AS GpStatus
   RETURN SetStatus(GdipFillRectangle(m_pGraphics, pBrush->m_pBrush, rc->x, rc->y, rc->Width, rc->Height))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.FillRectangle (BYVAL pBrush AS CGpBrush PTR, BYVAL rc AS GpRect PTR) AS GpStatus
   RETURN SetStatus(GdipFillRectangleI(m_pGraphics, pBrush->m_pBrush, rc->x, rc->y, rc->Width, rc->Height))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.FillRectangle (BYVAL pBrush AS CGpBrush PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipFillRectangle(m_pGraphics, pBrush->m_pBrush, x, y, nWidth, nHeight))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.FillRectangle (BYVAL pBrush AS CGpBrush PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS GpStatus
   RETURN SetStatus(GdipFillRectangleI(m_pGraphics, pBrush->m_pBrush, x, y, nWidth, nHeight))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Uses a brush to fill the interior of a sequence of rectangles.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.FillRectangles (BYVAL pBrush AS CGpBrush PTR, BYVAL rects AS GpRectF PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipFillRectangles(m_pGraphics, pBrush->m_pBrush, rects, nCount))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.FillRectangles (BYVAL pBrush AS CGpBrush PTR, BYVAL rects AS GpRect PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipFillRectanglesI(m_pGraphics, pBrush->m_pBrush, rects, nCount))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Uses a brush to fill the interior of a sequence of rectangles.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.FillPolygon (BYVAL pBrush AS CGpBrush PTR, BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipFillPolygon(m_pGraphics, pBrush->m_pBrush, pts, nCount, FillModeAlternate))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.FillPolygon (BYVAL pBrush AS CGpBrush PTR, BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipFillPolygonI(m_pGraphics, pBrush->m_pBrush, pts, nCount, FillModeAlternate))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.FillPolygon (BYVAL pBrush AS CGpBrush PTR, BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_, BYVAL nFillMode AS FillMode) AS GpStatus
   RETURN SetStatus(GdipFillPolygon(m_pGraphics, pBrush->m_pBrush, pts, nCount, nFillMode))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.FillPolygon (BYVAL pBrush AS CGpBrush PTR, BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_, BYVAL nFillMode AS FillMode) AS GpStatus
   RETURN SetStatus(GdipFillPolygonI(m_pGraphics, pBrush->m_pBrush, pts, nCount, nFillMode))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Uses a brush to fill the interior of a polygon.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.FillEllipse (BYVAL pBrush AS CGpBrush PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS GpStatus
   RETURN Setstatus(GdipFillEllipse(m_pGraphics, pBrush->m_pBrush, x, y, nWidth, nHeight))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.FillEllipse (BYVAL pBrush AS CGpBrush PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS GpStatus
   RETURN Setstatus(GdipFillEllipseI(m_pGraphics, pBrush->m_pBrush, x, y, nWidth, nHeight))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.FillEllipse (BYVAL pBrush AS CGpBrush PTR, BYVAL rc AS GpRectF PTR) AS GpStatus
   RETURN SetStatus(GdipFillEllipse(m_pGraphics, pBrush->m_pBrush, rc->x, rc->y, rc->Width, rc->Height))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.FillEllipse (BYVAL pBrush AS CGpBrush PTR, BYVAL rc AS GpRect PTR) AS GpStatus
   RETURN SetStatus(GdipFillEllipseI(m_pGraphics, pBrush->m_pBrush, rc->x, rc->y, rc->Width, rc->Height))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Uses a brush to fill the interior of a pie.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.FillPie (BYVAL pBrush AS CGpBrush PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipFillPie(m_pGraphics, pBrush->m_pBrush, x, y, nWidth, nHeight, startAngle, sweepAngle))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.FillPie (BYVAL pBrush AS CGpBrush PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipFillPieI(m_pGraphics, pBrush->m_pBrush, x, y, nWidth, nHeight, startAngle, sweepAngle))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.FillPie (BYVAL pBrush AS CGpBrush PTR, BYVAL rc AS GpRectF PTR, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipFillPie(m_pGraphics, pBrush->m_pBrush, rc->x, rc->y, rc->Width, rc->Height, startAngle, sweepAngle))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.FillPie (BYVAL pBrush AS CGpBrush PTR, BYVAL rc AS GpRect PTR, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipFillPieI(m_pGraphics, pBrush->m_pBrush, rc->x, rc->y, rc->Width, rc->Height, startAngle, sweepAngle))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Uses a brush to fill the interior of a path. If a figure in the path is not closed,
' this function treats the nonclosed figure as if it were closed by a straight line that
' connects the figure's starting and ending points.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.FillPath (BYVAL pBrush AS CGpBrush PTR, BYVAL pPath AS CGpGraphicsPath PTR) AS GpStatus
   RETURN SetStatus(GdipFillPath(m_pGraphics, pBrush->m_pBrush, pPath->m_pPath))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a closed cardinal spline from an array of points and uses a brush to fill the
' interior of the spline.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.FillClosedCurve (BYVAL pBrush AS CGpBrush PTR, BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipFillClosedCurve(m_pGraphics, pBrush->m_pBrush, pts, nCount))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.FillClosedCurve (BYVAL pBrush AS CGpBrush PTR, BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_) AS GpStatus
   RETURN SetStatus(GdipFillClosedCurveI(m_pGraphics, pBrush->m_pBrush, pts, nCount))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.FillClosedCurve (BYVAL pBrush AS CGpBrush PTR, BYVAL pts AS GpPointF PTR, BYVAL nCount AS LONG, BYVAL nFillMode AS FillMode, BYVAL tension AS SINGLE = 0.5) AS GpStatus
   RETURN SetStatus(GdipFillClosedCurve2(m_pGraphics, pBrush->m_pBrush, pts, nCount, tension, nFillMode))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.FillClosedCurve (BYVAL pBrush AS CGpBrush PTR, BYVAL pts AS GpPoint PTR, BYVAL nCount AS LONG, BYVAL nFillMode AS FillMode, BYVAL tension AS SINGLE = 0.5) AS GpStatus
   RETURN SetStatus(GdipFillClosedCurve2I(m_pGraphics, pBrush->m_pBrush, pts, nCount, tension, nFillMode))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Uses a brush to fill a specified region.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.FillRegion (BYVAL pBrush AS CGpBrush PTR, BYVAL pRegion AS CGpRegion PTR) AS GpStatus
   RETURN SetStatus(GdipFillRegion(m_pGraphics, pBrush->m_pBrush, pRegion->m_pRegion))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Draws a string based on a font and an origin for the string.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL rcf AS GpRectF PTR, BYVAL pBrush AS CGpBrush PTR) AS GpStatus
   RETURN SetStatus(GdipDrawString(m_pGraphics, pwszString, length, pFont->m_pFont, rcf, NULL, pBrush->m_pBrush))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, _
   BYVAL nHeight AS SINGLE, BYVAL pBrush AS CGpBrush PTR) AS GpStatus
   DIM layoutRect AS GpRectF = TYPE<GpRectF>(x, y, nWidth, nHeight)
   RETURN SetStatus(GdipDrawString(m_pGraphics, pwszString, length, pFont->m_pFont, @layoutRect, NULL, pBrush->m_pBrush))
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL layoutRect AS GpRectF PTR, _
   BYVAL pStringFormat AS CGpStringFormat PTR, BYVAL pBrush AS CGpBrush PTR) AS GpStatus
   RETURN SetStatus(GdipDrawString(m_pGraphics, pwszString, length, pFont->m_pFont, _
      layoutRect, pStringFormat->m_pStringFormat, pBrush->m_pBrush))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE, _
   BYVAL pStringFormat AS CGpStringFormat PTR, BYVAL pBrush AS CGpBrush PTR) AS GpStatus
   DIM layoutRect AS GpRectF = TYPE<GpRectF>(x, y, nWidth, nHeight)
   RETURN SetStatus(GdipDrawString(m_pGraphics, pwszString, length, pFont->m_pFont, _
      @layoutRect, pStringFormat->m_pStringFormat, pBrush->m_pBrush))
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL origin as GpPointF PTR, BYVAL pBrush AS CGpBrush PTR) AS GpStatus
   DIM layoutRect AS GpRectF = TYPE<GpRectF>(origin->x, origin->y, 0.0, 0.0)
   RETURN SetStatus(GdipDrawString(m_pGraphics, pwszString, length, pFont->m_pFont, @layoutRect, NULL, pBrush->m_pBrush))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL pBrush AS CGpBrush PTR) AS GpStatus
   DIM layoutRect AS GpRectF = TYPE<GpRectF>(x, y, 0.0, 0.0)
   RETURN SetStatus(GdipDrawString(m_pGraphics, pwszString, length, pFont->m_pFont, @layoutRect, NULL, pBrush->m_pBrush))
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL origin AS GpPointF PTR, _
   BYVAL pStringFormat AS CGpStringFormat PTR, BYVAL pBrush AS CGpBrush PTR) AS GpStatus
   DIM layoutRect AS GpRectF = TYPE<GpRectF>(origin->x, origin->y, 0.0, 0.0)
   RETURN SetStatus(GdipDrawString(m_pGraphics, pwszString, length, pFont->m_pFont, _
      @layoutRect, pStringFormat->m_pStringFormat, pBrush->m_pBrush))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, _
   BYVAL pStringFormat AS CGpStringFormat PTR, BYVAL pBrush AS CGpBrush PTR) AS GpStatus
   DIM layoutRect AS GpRectF = TYPE<GpRectF>(x, y, 0.0, 0.0)
   RETURN SetStatus(GdipDrawString(m_pGraphics, pwszString, length, pFont->m_pFont, _
      @layoutRect, pStringFormat->m_pStringFormat, pBrush->m_pBrush))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Measures the extent of the string in the specified font, format, and layout rectangle.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.MeasureString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFont AS CGpFont PTR, _
   BYVAL layoutRect AS GpRectF PTR, BYVAL pStringFormat AS CGpStringFormat PTR, BYVAL boundingBox AS GpRectF PTR, _
   BYREF codepointsFitted AS INT_ = 0, BYREF linesFilled AS INT_ = 0) AS GpStatus
   RETURN SetStatus(GdipMeasureString(m_pGraphics, pwszString, length, _
      pFont->m_pFont, layoutRect, pStringFormat->m_pStringFormat, _
      boundingBox, @codepointsFitted, @linesFilled))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.MeasureString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFont AS CGpFont PTR, _
   BYVAL layoutRectSize AS GpSizeF PTR, BYVAL pStringFormat AS CGpStringFormat PTR, BYVAL size AS GpSizeF PTR, _
   BYREF codepointsFitted AS INT_ = 0, BYREF linesFilled AS INT_ = 0) AS GpStatus
   DIM layoutRect AS GpRectF = TYPE<GpRectF>(0, 0, layoutRectSize->Width, layoutRectSize->Height)
   DIM boundingBox AS GpRectF
   m_Status = GdipMeasureString(m_pGraphics, pwszString, length, _
      pFont->m_pFont, @layoutRect, pStringFormat->m_pStringFormat, _
      @boundingBox, @codepointsFitted, @linesFilled)
   IF m_Status = Ok THEN
      size->Width = boundingBox.Width
      size->Height = boundingBox.Height
   END IF
   RETURN m_Status
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.MeasureString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFont AS CGpFont PTR, _
   BYVAL origin AS GpPointF PTR, BYVAL pStringFormat AS CGpStringFormat PTR, BYVAL boundingBox AS GpRectF PTR) AS GpStatus
   DIM layoutRect AS GpRectF = TYPE<GpRectF>(origin->x, origin->y, 0.0, 0.0)
   RETURN SetStatus(GdipMeasureString(m_pGraphics, pwszString, length, _
      pFont->m_pFont, @layoutRect, pStringFormat->m_pStringFormat, boundingBox, NULL, NULL))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.MeasureString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFont AS CGpFont PTR, _
   BYVAL layoutRect AS GpRectF PTR, BYVAL pStringFormat AS CGpStringFormat PTR, BYVAL boundingBox AS GpRectF PTR) AS GpStatus
   RETURN SetStatus(GdipMeasureString(m_pGraphics, pwszString, length, _
      pFont->m_pFont, layoutRect, pStringFormat->m_pStringFormat, boundingBox, NULL, NULL))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.MeasureString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFont AS CGpFont PTR, _
   BYVAL layoutRect AS GpRectF PTR, BYVAL boundingBox AS GpRectF PTR) AS GpStatus
   RETURN SetStatus(GdipMeasureString(m_pGraphics, pwszString, length, _
      pFont->m_pFont, layoutRect, NULL, boundingBox, NULL, NULL))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.MeasureString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFont AS CGpFont PTR, _
   BYVAL origin AS GpPointF PTR, BYVAL boundingBox AS GpRectF PTR) AS GpStatus
   DIM layoutRect AS GpRectF = TYPE<GpRectF>(origin->x, origin->y, 0.0, 0.0)
   RETURN SetStatus(GdipMeasureString(m_pGraphics, pwszString, length, _
      pFont->m_pFont, @layoutRect, NULL, boundingBox, NULL, NULL))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a set of regions each of which bounds a range of character positions within a string.
' GdipMeasureCharacterRanges not declared in the FB headers for 64 bit.
' ========================================================================================
#ifndef __FB_64BIT__
PRIVATE FUNCTION CGpGraphics.MeasureCharacterRanges (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL layoutRect AS GpRectF PTR, BYVAL pStringFormat AS CGpStringFormat PTR, _
   BYVAL regionCount AS INT_, BYVAL regions AS CGpRegion PTR) AS GpStatus
   IF pwszString = NULL OR pFont = NULL OR layoutRect = NULL OR pStringFormat = NULL OR regions = NULL THEN RETURN SetStatus(InvalidParameter)
   DIM pRegions(regionCount - 1) AS GpRegion PTR
   FOR i AS LONG = 0 TO regionCount - 1
      pRegions(i) = regions[i].m_pRegion
   NEXT
   RETURN SetStatus(GdipMeasureCharacterRanges(m_pGraphics, pwszString, length, pFont->m_pFont, layoutRect, _
          pStringFormat->m_pStringFormat, regionCount, @pRegions(0)))
END FUNCTION
#endif
' ========================================================================================

' ========================================================================================
' Draws characters at the specified positions. The function gives the client complete
' control over the appearance of text. The f assumes that the client has already
' set up the format and layout to be applied.
' Note: Positions declared as ANY PTR because in the FB headers is declared as PointF PTR
' in one file and as PointF_ PTR in another. It should be GpPointF.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawDriverString (BYVAL pText AS UINT16 PTR, BYVAL length AS INT_, BYVAL pFont AS CGpFont PTR, _
   BYVAL pBrush AS CGpBrush PTR, BYVAL positions AS ANY PTR, BYVAL flags AS INT_, BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
   RETURN SetStatus(GdipDrawDriverString(m_pGraphics, pText, length, pFont->m_pFont, pBrush->m_pBrush, positions, flags, pMatrix->m_pMatrix))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Measures the bounding box for the specified characters and their corresponding positions.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.MeasureDriverString (BYVAL pText AS UINT16 PTR, BYVAL length AS LONG, BYVAL pFont AS CGpFont PTR, _
   BYVAL positions AS ANY PTR, BYVAL flags AS LONG, BYVAL pMatrix AS CGpMatrix PTR, BYVAL boundingBox AS GpRectF PTR) AS GpStatus
   RETURN SetStatus(GdipMeasureDriverString(m_pGraphics, pText, length, pFont->m_pFont, _
              positions, flags, pMatrix->m_pMatrix, boundingBox))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Draws a cached bitmap on this graphics destination offset by x, y. Note this will fail
' with WrongState if the CachedBitmap native format differs from this Graphics.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawCachedBitmap (BYVAL pCachedBitmap AS CGpCachedBitmap PTR, BYVAL x AS LONG, BYVAL y AS LONG) AS GpStatus
   RETURN SetStatus(GdipDrawCachedBitmap(m_pGraphics, IIF(pCachedBitmap, pCachedBitmap->m_pCachedBitmap, NULL), x, y))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Draws an image.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL pt AS GpPointF PTR) AS GpStatus
   RETURN SetStatus(GdipDrawImage(m_pGraphics, pImage->m_pImage, pt->x, pt->y))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipDrawImage(m_pGraphics, pImage->m_pImage, x, y))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipDrawImageRect(m_pGraphics, pImage->m_pImage, x, y, nWidth, nHeight))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL rcf AS GpRectF PTR) AS GpStatus
   RETURN this.DrawImage(pImage, rcf->x, rcf->y, rcf->Width, rcf->Height)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL pt AS GpPoint PTR) AS GpStatus
   RETURN SetStatus(GdipDrawImageI(m_pGraphics, pImage->m_pImage, pt->x, pt->y))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL x AS INT_, BYVAL y AS INT_) AS GpStatus
   RETURN SetStatus(GdipDrawImageI(m_pGraphics, pImage->m_pImage, x, y))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS GpStatus
   RETURN SetStatus(GdipDrawImageRectI(m_pGraphics, pImage->m_pImage, x, y, nWidth, nHeight))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL rc AS GpRect PTR) AS GpStatus
   RETURN SetStatus(GdipDrawImageRectI(m_pGraphics, pImage->m_pImage, rc->x, rc->y, rc->Width, rc->Height))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destPoints AS GpPointF PTR, BYVAL nCount AS INT_) AS GpStatus
   IF nCount <> 3 AND nCount <> 4 THEN RETURN SetStatus(InvalidParameter)
   RETURN SetStatus(GdipDrawImagePoints(m_pGraphics, pImage->m_pImage, destPoints, nCount))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destPoints AS GpPoint PTR, BYVAL nCount AS INT_) AS GpStatus
   IF nCount <> 3 AND nCount <> 4 THEN RETURN SetStatus(InvalidParameter)
   RETURN SetStatus(GdipDrawImagePointsI(m_pGraphics, pImage->m_pImage, destPoints, nCount))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, _
   BYVAL srcx AS SINGLE, BYVAL srcy AS SINGLE, BYVAL srcwidth AS SINGLE, BYVAL srcheight AS SINGLE, BYVAL srcUnit AS gpUnit) AS GpStatus
   RETURN SetStatus(GdipDrawImagePointRect(m_pGraphics, pImage->m_pImage, x, y, srcx, srcy, srcwidth, srcheight, srcUnit))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL x AS INT_, BYVAL y AS INT_, _
   BYVAL srcx AS INT_, BYVAL srcy AS INT_, BYVAL srcwidth AS INT_, BYVAL srcheight AS INT_, BYVAL srcUnit AS gpUnit) AS GpStatus
   RETURN SetStatus(GdipDrawImagePointRectI(m_pGraphics, pImage->m_pImage, x, y, srcx, srcy, srcwidth, srcheight, srcUnit))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destRect AS GpRectF PTR, _
   BYVAL srcx AS SINGLE, BYVAL srcy AS SINGLE, BYVAL srcwidth AS SINGLE, BYVAL srcheight AS SINGLE, BYVAL srcUnit AS GpUnit, _
   BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL, BYVAL pcallback AS DrawImageAbort = NULL, BYVAL pcallbackData AS ANY PTR = NULL) AS GpStatus
   RETURN SetStatus(GdipDrawImageRectRect(m_pGraphics, pImage->m_pImage, destRect->x, destRect->y, _
   destRect->Width, destRect->Height, srcx, srcy, srcwidth, srcheight, srcUnit, _
   IIF(pImageAttributes, pImageAttributes->m_pImgAttr, NULL), pCallback, pcallbackData))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destRect AS GpRect PTR, _
   BYVAL srcx AS INT_, BYVAL srcy AS INT_, BYVAL srcwidth AS INT_, BYVAL srcheight AS INT_, BYVAL srcUnit AS GpUnit, _
   BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL, BYVAL pcallback AS DrawImageAbort = NULL, BYVAL pcallbackData AS ANY PTR = NULL) AS GpStatus
   RETURN SetStatus(GdipDrawImageRectRectI(m_pGraphics, pImage->m_pImage, destRect->x, destRect->y, _
   destRect->Width, destRect->Height, srcx, srcy, srcwidth, srcheight, srcUnit, _
   IIF(pImageAttributes, pImageAttributes->m_pImgAttr, NULL), pCallback, pcallbackData))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destx AS SINGLE, _
   BYVAL desty AS SINGLE, BYVAL destWidth AS SINGLE, BYVAL destHeight AS SINGLE, _
   BYVAL srcx AS SINGLE, BYVAL srcy AS SINGLE, BYVAL srcwidth AS SINGLE, BYVAL srcheight AS SINGLE, BYVAL srcUnit AS GpUnit, _
   BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL, BYVAL pcallback AS DrawImageAbort = NULL, BYVAL pcallbackData AS ANY PTR = NULL) AS GpStatus
   RETURN SetStatus(GdipDrawImageRectRect(m_pGraphics, pImage->m_pImage, destx, desty, _
   destWidth, destHeight, srcx, srcy, srcwidth, srcheight, srcUnit, _
   IIF(pImageAttributes, pImageAttributes->m_pImgAttr, NULL), pCallback, pcallbackData))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destx AS INT_, _
   BYVAL desty AS INT_, BYVAL destWidth AS INT_, BYVAL destHeight AS INT_, _
   BYVAL srcx AS INT_, BYVAL srcy AS INT_, BYVAL srcwidth AS INT_, BYVAL srcheight AS INT_, BYVAL srcUnit AS GpUnit, _
   BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL, BYVAL pcallback AS DrawImageAbort = NULL, BYVAL pcallbackData AS ANY PTR = NULL) AS GpStatus
   RETURN SetStatus(GdipDrawImageRectRectI(m_pGraphics, pImage->m_pImage, destx, desty, _
   destWidth, destHeight, srcx, srcy, srcwidth, srcheight, srcUnit, _
   IIF(pImageAttributes, pImageAttributes->m_pImgAttr, NULL), pCallback, pcallbackData))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destPoints AS GpPointF PTR, BYVAL nCount AS INT_, _
   BYVAL srcx AS SINGLE, BYVAL srcy AS SINGLE, BYVAL srcwidth AS SINGLE, BYVAL srcheight AS SINGLE, BYVAL srcUnit AS GpUnit, _
   BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL, BYVAL pcallback AS DrawImageAbort = NULL, BYVAL pcallbackData AS ANY PTR = NULL) AS GpStatus
   RETURN SetStatus(GdipDrawImagePointsRect(m_pGraphics, pImage->m_pImage, destPoints, nCount, _
   srcx, srcy, srcwidth, srcheight, srcUnit, IIF(pImageAttributes, pImageAttributes->m_pImgAttr, NULL), pCallback, pcallbackData))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destPoints AS GpPoint PTR, BYVAL nCount AS INT_, _
   BYVAL srcx AS INT_, BYVAL srcy AS INT_, BYVAL srcwidth AS INT_, BYVAL srcheight AS INT_, BYVAL srcUnit AS GpUnit, _
   BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL, BYVAL pcallback AS DrawImageAbort = NULL, BYVAL pcallbackData AS ANY PTR = NULL) AS GpStatus
   RETURN SetStatus(GdipDrawImagePointsRectI(m_pGraphics, pImage->m_pImage, destPoints, nCount, _
   srcx, srcy, srcwidth, srcheight, srcUnit, IIF(pImageAttributes, pImageAttributes->m_pImgAttr, NULL), pCallback, pcallbackData))
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destRect AS GpRectF PTR, _
   BYVAL sourceRect AS GpRectF PTR, BYVAL srcUnit AS GpUnit, BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL) AS GpStatus
   RETURN SetStatus(GdipDrawImageRectRect(m_pGraphics, pImage->m_pImage, _
          destRect->x, destRect->y, destRect->Width, destRect->Height, sourceRect->x, sourceRect->y, _
          sourceRect->Width, sourceRect->Height, srcUnit, IIF(pImageAttributes, pImageAttributes->m_pImgAttr, NULL), NULL, NULL))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destRect AS GpRect PTR, _
   BYVAL sourceRect AS GpRect PTR, BYVAL srcUnit AS GpUnit, BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL) AS GpStatus
   RETURN SetStatus(GdipDrawImageRectRectI(m_pGraphics, pImage->m_pImage, _
          destRect->x, destRect->y, destRect->Width, destRect->Height, sourceRect->x, sourceRect->y, _
          sourceRect->Width, sourceRect->Height, srcUnit, IIF(pImageAttributes, pImageAttributes->m_pImgAttr, NULL), NULL, NULL))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destx AS SINGLE, BYVAL desty AS SINGLE, BYVAL destWidth AS SINGLE, BYVAL destHeight AS SINGLE, _
   BYVAL srcx AS SINGLE, BYVAL srcy AS SINGLE, BYVAL srcWidth AS SINGLE, BYVAL srcHeight AS SINGLE, BYVAL srcUnit AS GpUnit, BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL) AS GpStatus
   RETURN SetStatus(GdipDrawImageRectRect(m_pGraphics, pImage->m_pImage, _
          destx, desty, destWidth, destHeight, srcx, srcy, srcWidth, srcHeight, _
          srcUnit, IIF(pImageAttributes, pImageAttributes->m_pImgAttr, NULL), NULL, NULL))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destx AS INT_, BYVAL desty AS INT_, BYVAL destWidth AS INT_, BYVAL destHeight AS INT_, _
   BYVAL srcx AS INT_, BYVAL srcy AS INT_, BYVAL srcWidth AS INT_, BYVAL srcHeight AS INT_, BYVAL srcUnit AS GpUnit, BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL) AS GpStatus
   RETURN SetStatus(GdipDrawImageRectRectI(m_pGraphics, pImage->m_pImage, _
          destx, desty, destWidth, destHeight, srcx, srcy, srcWidth, srcHeight, _
          srcUnit, IIF(pImageAttributes, pImageAttributes->m_pImgAttr, NULL), NULL, NULL))
END FUNCTION
' ========================================================================================

'#IF (GDIPVER >= &H0110)

' ========================================================================================
'FUNCTION CGpGraphics.DrawImageFX (BYVAL pImage AS CGpImage PTR, BYVAL sourceRect AS GpRectF PTR, BYVAL pMatrix AS CGpMatrix PTR, _
'BYVAL pEffect AS CGpEffect PTR, BYVAL pImageAttributes AS CGpImageAttributes PTR, BYVAL srcUnit AS GpUnit) AS GpStatus
'   RETURN SetStatus(GdipDrawImageFX(m_pGraphics, pImage->m_pImage, sourceRect, pMatrix->m_pMatrix, _
'          IIF(pEffect, pEffect->m_pEFfect, NULL), IIF(pImageAttributes, pImageAttributes->m_pImgAtt, NULL), srcUnit))
'END FUNCTION
' ========================================================================================

'#ENDIF   ' //(GDIPVER >= 0x0110)

' ========================================================================================
' Calls an application-defined callback function for each record in a specified metafile.
' You can use this function to display a metafile by calling PlayRecord in the callback function.
' ========================================================================================
'FUNCTION CGpGraphics.EnumerateMetafileDestPoint (BYVAL pMetafile AS CGpMetafile PTR, BYVAL destPoint AS GpPointF PTR, _
'BYVAL pCallback AS EnumerateMetafileProc = NULL, BYVAL pCallbackData AS ANY PTR = NULL, BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL) AS GpStatus
'   RETURN SetStatus(GdipEnumerateMetafileDestPoint(m_pGraphics, IIF(pMetafile, pMetafile->m_pMetafile, NULL), _
'          destPoint, pCallback, pCallbackData, IIF(pImageAttributes, pImageAttributes->m_pImgAttr, NULL)))
'END FUNCTION
' ========================================================================================
' ========================================================================================
'FUNCTION CGpGraphics.EnumerateMetafileDestRect (BYVAL pMetafile AS CGpMetafile PTR, BYVAL destRect AS GpRectF PTR, BYVAL pCallback AS EnumerateMetafileProc, _
'BYVAL pCallbackData AS ANY PTR = NULL, BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL) AS GpStatus
'   RETURN SetStatus(GdipEnumerateMetafileDestRect(m_pGraphics, IIF(pMetafile, pMetafile->m_pMetafile, NULL), _
'          destRect, pCallback, pCallbackData, IIF(pImageAttributes, pImageAttributes->m_pImgAttr, NULL)))
'END FUNCTION
' ========================================================================================
' ========================================================================================
'FUNCTION CGpGraphics.EnumerateMetafileDestPoints (BYVAL pMetafile AS CGpMetafile PTR, BYVAL destPoints AS GpPointF PTR, BYVAL nCount AS LONG, BYVAL pCallback AS EnumerateMetafileProc, _
'BYVAL pCallbackData AS ANY PTR = NULL, BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL) AS GpStatus
'   RETURN SetStatus(GdipEnumerateMetafileDestPoints(m_pGraphics, IIF(ISOBJECT(pMetafile), pMetafile.Ptr, %NULL), _
'          destPoints, nCount, pCallback, pCallbackData, IIF(ISOBJECT(pImageAttributes), pImageAttributes->m_pImgAttr, NULL)))
'END FUNCTION
' ========================================================================================
' ========================================================================================
'FUNCTION CGpGraphics.EnumerateMetafileSrcRectDestPoint (BYVAL pMetafile AS CGpMetafile, BYVAL destPoint AS GpPointF PTR, _
'BYVAL srcRect AS GpRectF PTR, BYVAL srcUnit AS GpUnit, BYVAL pCallback AS EnumerateMetafileProc, _
'BYVAL pCallbackData AS ANY PTR = NULL, BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL) AS GpStatus
'   RETURN SetStatus(GdipEnumerateMetafileSrcRectDestPoint(m_pGraphics, IIF(pMetafile, pMetafile->m_pMetafile, NULL), _
'         destPoint, srcRect, srcUnit, pCallback, pCallbackData, IIF(pImageAttributes, pImageAttributes->m_pImgAttr, NULL)))
'END FUNCTION
' ========================================================================================
' ========================================================================================
'FUNCTION CGpGraphics.EnumerateMetafileSrcRectDestRect (BYVAL pMetafile AS CGpMetafile, BYVAL destRect AS GpRectF PTR, _
'BYVAL srcRect AS GpRectF PTR, BYVAL srcUnit AS GpUnit, BYVAL pCallback AS EnumerateMetafileProc, _
'BYVAL pCallbackData AS ANY PTR = NULL, BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL) AS GpStatus
'   RETURN SetStatus(GdipEnumerateMetafileSrcRectDestRect(m_pGraphics, IIF(pMetafile, pMetafile->m_pMetafile, NULL), _
'          destRect, srcRect, srcUnit, pCallback, pCallbackData, IIF(pImageAttributes, pImageAttributes->m_pImgAtt, NULL)))
'END FUNCTION
' ========================================================================================
' ========================================================================================
'FUNCTION CGpGraphics.EnumerateMetafileSrcRectDestPoints (BYVAL pMetafile AS CGpMetafile, BYVAL destPoints AS GpPointF PTR, _
'BYVAL nCount AS INT_, BYVAL srcRect AS GpRectF PTR, BYVAL srcUnit AS GpUnit, BYVAL pCallback AS EnumerateMetafileProc, _
'BYVAL pCallbackData AS ANY PTR = NULL, BYVAL pImageAttributes AS CGpImageAttributes = NULL) AS GpStatus
'   RETURN SetStatus(GdipEnumerateMetafileSrcRectDestPoints(m_pGraphics, IIF(pMetafile, pMetafile->m_pMetaFile, NULL), _
'          destPoints, nCount, srcRect, srcUnit, pCallback, pCallbackData, IIF(pImageAttributes, pImageAttributes->m_pImgAtt, NULL)))
'END FUNCTION
' ========================================================================================

' ========================================================================================
' Updates the clipping region of this Graphics object to a region that is the combination
' of itself and the clipping region of another Graphics object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.SetClip (BYVAL pGraphics AS CGpGraphics PTR, BYVAL nCombineMode AS CombineMode = CombineModeReplace) AS GpStatus
   RETURN SetStatus(GdipSetClipGraphics(m_pGraphics, IIF(pGraphics, pGraphics->m_pGraphics, NULL), nCombineMode))
END FUNCTION
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.SetClip (BYVAL rc AS GpRectF PTR, BYVAL nCombineMode AS CombineMode = CombineModeReplace) AS GpStatus
   RETURN SetStatus(GdipSetClipRect(m_pGraphics, rc->x, rc->y, rc->Width, rc->Height, nCombineMode))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.SetClip (BYVAL rc AS GpRect PTR, BYVAL nCombineMode AS CombineMode = CombineModeReplace) AS GpStatus
   RETURN SetStatus(GdipSetClipRect(m_pGraphics, rc->x, rc->y, rc->Width, rc->Height, nCombineMode))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.SetClip (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE, BYVAL nCombineMode AS CombineMode = CombineModeReplace) AS GpStatus
   RETURN SetStatus(GdipSetClipRect(m_pGraphics, x, y, nWidth, nHeight, nCombineMode))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.SetClip (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL nCombineMode AS CombineMode = CombineModeReplace) AS GpStatus
   RETURN SetStatus(GdipSetClipRect(m_pGraphics, x, y, nWidth, nHeight, nCombineMode))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Updates the clipping region of this Graphics object to a region that is the combination
' of itself and the region specified by a graphics path. If a figure in the path is not
' closed, this FUNCTION treats the nonclosed figure as if it were closed by a straight
' line that connects the figure's starting and ending points.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.SetClip (BYVAL pPath AS CGpGraphicsPath PTR, BYVAL nCombineMode AS CombineMode = CombineModeReplace) AS GpStatus
   RETURN SetStatus(GdipSetClipPath(m_pGraphics, pPath->m_pPath, nCombineMode))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Updates the clipping region of this Graphics object to a region that is the
' combination of itself and the region specified by a Region object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.SetClip (BYVAL pRegion AS CGpRegion PTR, BYVAL nCombineMode AS CombineMode = CombineModeReplace) AS GpStatus
   RETURN SetStatus(GdipSetClipRegion(m_pGraphics, pRegion->m_pRegion, nCombineMode))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Updates the clipping region of this Graphics object to a region that is the combination
' of itself and a Windows Graphics Device Interface (GDI) region.
' This is different than the other SetClip methods because it assumes that the HRGN is
' already in device units, so it doesn't transform the coordinates in the HRGN.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.SetClip (BYVAL hRgn AS HRGN, BYVAL nCombineMode AS CombineMode = CombineModeReplace) AS GpStatus
   RETURN SetStatus(GdipSetClipHrgn(m_pGraphics, hRgn, nCombineMode))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Updates the clipping region of this Graphics object to the portion of the specified
' rectangle that intersects with the current clipping region of this Graphics object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.IntersectClip (BYVAL rcf AS GpRectF PTR) AS GpStatus
   RETURN SetStatus(GdipSetClipRect(m_pGraphics, rcf->x, rcf->y, rcf->Width, rcf->Height, CombineModeIntersect))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.IntersectClip (BYVAL rc AS GpRect PTR) AS GpStatus
   RETURN SetStatus(GdipSetClipRectI(m_pGraphics, rc->x, rc->y, rc->Width, rc->Height, CombineModeIntersect))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.IntersectClip (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipSetClipRect(m_pGraphics, x, y, nWidth, nHeight, CombineModeIntersect))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.IntersectClip (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS GpStatus
   RETURN SetStatus(GdipSetClipRectI(m_pGraphics, x, y, nWidth, nHeight, CombineModeIntersect))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Updates the clipping region of this Graphics object to the portion of the specified
' region that intersects with the current clipping region of this Graphics object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.IntersectClip (BYVAL pRegion AS CGpRegion PTR) AS GpStatus
   RETURN SetStatus(GdipSetClipRegion(m_pGraphics, pRegion->m_pRegion, CombineModeIntersect))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Updates the clipping region to the portion of itself that does not intersect the specified rectangle.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.ExcludeClip (BYVAL rcf AS GpRectF PTR) AS GpStatus
   RETURN SetStatus(GdipSetClipRect(m_pGraphics, rcf->x, rcf->y, rcf->Width, rcf->Height, CombineModeExclude))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.ExcludeClip (BYVAL rc AS GpRect PTR) AS GpStatus
   RETURN SetStatus(GdipSetClipRectI(m_pGraphics, rc->x, rc->y, rc->Width, rc->Height, CombineModeExclude))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.ExcludeClip (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipSetClipRect(m_pGraphics, x, y, nWidth, nHeight, CombineModeExclude))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.ExcludeClip (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS GpStatus
   RETURN SetStatus(GdipSetClipRectI(m_pGraphics, x, y, nWidth, nHeight, CombineModeExclude))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Updates the clipping region with the portion of itself that does not overlap the specified region.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.ExcludeClip (BYVAL pRegion AS CGpRegion PTR) AS GpStatus
   RETURN SetStatus(GdipSetClipRegion(m_pGraphics, pRegion->m_pRegion, CombineModeExclude))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the clipping region of this Graphics object to an infinite region.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.ResetClip () AS GpStatus
   RETURN SetStatus(GdipResetClip(m_pGraphics))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates the clipping region of this Graphics object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.TranslateClip (BYVAL dx AS SINGLE, BYVAL dy AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipTranslateClip(m_pGraphics, dx, dy))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the clipping region of this Graphics object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.GetClip (BYVAL pRegion AS CGpRegion PTR) AS GpStatus
   IF pRegion = NULL THEN RETURN SetStatus(InvalidParameter)
   RETURN SetStatus(GdipGetClip(m_pGraphics, pRegion->m_pRegion))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a rectangle that encloses the clipping region of this Graphics object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.GetClipBounds (BYVAL rc AS GpRectF PTR) AS GpStatus
   RETURN SetStatus(GdipGetClipBounds(m_pGraphics, rc))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets a rectangle that encloses the clipping region of this Graphics object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.GetClipBounds (BYVAL rc AS GpRect PTR) AS GpStatus
   RETURN SetStatus(GdipGetClipBoundsI(m_pGraphics, rc))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether the clipping region of this Graphics object is empty.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.IsClipEmpty () AS BOOLEAN
   DIM booln AS BOOL
   m_Status = GdipIsClipEmpty(m_pGraphics, @booln)
   RETURN booln
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a rectangle that encloses the visible clipping region of this Graphics object.
' The visible clipping region is the intersection of the clipping region of this
' Graphics object and the clipping region of the window.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.GetVisibleClipBounds (BYVAL rc AS GpRectF PTR) AS GpStatus
   RETURN SetStatus(GdipGetVisibleClipBounds(m_pGraphics, rc))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.GetVisibleClipBounds (BYVAL rc AS GpRect PTR) AS GpStatus
   RETURN SetStatus(GdipGetVisibleClipBoundsI(m_pGraphics, rc))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether the visible clipping region of this Graphics object is empty. The
' visible clipping region is the intersection of the clipping region of this Graphics
' object and the clipping region of the window.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.IsVisibleClipEmpty () AS BOOLEAN
   DIM booln AS BOOL
   m_Status = GdipIsVisibleClipEmpty(m_pGraphics, @booln)
   RETURN booln
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether the specified point is inside the visible clipping region of this
' Graphics object. The visible clipping region is the intersection of the clipping
' region of this Graphics object and the clipping region of the window.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.IsVisible (BYVAL pt AS GpPointF PTR) AS BOOLEAN
   DIM booln AS BOOL
   m_Status = GdipIsVisiblePoint(m_pGraphics, pt->x, pt->y, @booln)
   RETURN booln
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.IsVisible (BYVAL pt AS GpPoint PTR) AS BOOLEAN
   DIM booln AS BOOL
   m_Status = GdipIsVisiblePoint(m_pGraphics, pt->x, pt->y, @booln)
   RETURN booln
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.IsVisible (BYVAL x AS SINGLE, BYVAL y AS SINGLE) AS BOOLEAN
   DIM booln AS BOOL
   m_Status = GdipIsVisiblePoint(m_pGraphics, x, y, @booln)
   RETURN booln
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.IsVisible (BYVAL x AS INT_, BYVAL y AS INT_) AS BOOLEAN
   DIM booln AS BOOL
   m_Status = GdipIsVisiblePoint(m_pGraphics, x, y, @booln)
   RETURN booln
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.IsVisible (BYVAL rcf AS GpRectF PTR) AS BOOLEAN
   DIM booln AS BOOL
   m_Status = GdipIsVisibleRect(m_pGraphics, rcf->x, rcf->y, rcf->Width, rcf->Height, @booln)
   RETURN booln
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.IsVisible (BYVAL rc AS GpRect PTR) AS BOOLEAN
   DIM booln AS BOOL
   m_Status = GdipIsVisibleRectI(m_pGraphics, rc->x, rc->y, rc->Width, rc->Height, @booln)
   RETURN booln
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.IsVisible (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS BOOLEAN
   DIM booln AS BOOL
   m_Status = GdipIsVisibleRect(m_pGraphics, x, y, nWidth, nHeight, @booln)
   RETURN booln
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.IsVisible (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS BOOLEAN
   DIM booln AS BOOL
   m_Status = GdipIsVisibleRectI(m_pGraphics, x, y, nWidth, nHeight, @booln)
   RETURN booln
END FUNCTION
' ========================================================================================

' ========================================================================================
' Saves the current state (transformations, clipping region, and quality settings) of
' this Graphics object. You can restore the state later by calling the Restore function.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.Save () AS GraphicsState
   DIM gstate AS GraphicsState
   m_Status = GdipSaveGraphics(m_pGraphics, @gstate)
   RETURN gstate
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the state of this Graphics object to the state stored by a previous call to the
' Save function. of this Graphics object.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.Restore (BYVAL gstate AS GraphicsState) AS GpStatus
   RETURN SetStatus(GdipRestoreGraphics(m_pGraphics, gstate))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Begins a new graphics container.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.BeginContainer () AS GraphicsContainer
   DIM state AS GraphicsContainer
   m_Status = GdipBeginContainer2(m_pGraphics, @state)
   RETURN state
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.BeginContainer (BYVAL destrect AS GpRectF PTR, BYVAL srcrect AS GpRectF PTR, BYVAL nUnit AS GpUnit) AS GraphicsContainer
   DIM state AS GraphicsContainer
   m_Status = GdipBeginContainer(m_pGraphics, destrect, srcrect, nUnit, @state)
   RETURN state
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.BeginContainer (BYVAL destrect AS GpRect PTR, BYVAL srcrect AS GpRect PTR, BYVAL nUnit AS GpUnit) AS GraphicsContainer
   DIM state AS GraphicsContainer
   m_Status = GdipBeginContainerI(m_pGraphics, destrect, srcrect, nUnit, @state)
   RETURN state
END FUNCTION
' ========================================================================================

' ========================================================================================
' Closes a graphics container that was previously opened by the BeginContainer FUNCTION.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.EndContainer (BYVAL state AS GraphicsContainer) AS GpStatus
   RETURN SetStatus(GdipEndContainer(m_pGraphics, state))
END FUNCTION
' ========================================================================================

' // Only valid when recording metafiles.

' ========================================================================================
' Adds a text comment to an existing metafile.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.AddMetafileComment (BYVAL pdata AS BYTE PTR, BYVAL sizeData AS UINT) AS GpStatus
   RETURN SetStatus(GdipComment(m_pGraphics, sizeData, pdata))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a Windows halftone palette.
' ========================================================================================
PRIVATE FUNCTION CGpGraphics.GetHalftonePalette () AS HPALETTE
   RETURN GdipCreateHalftonePalette
END FUNCTION
' ========================================================================================

'#IF (GDIPVER >= &H0110)
'' ========================================================================================
'' Not used in Windows GDI+ versions 1.0 and 1.1.
'' ========================================================================================
'FUNCTION SetAbort (BYVAL pIAbort AS DWORD) AS GpStatus
'   m_Status = GdipGraphicsSetAbort(m_pGraphics, pIAbort)
'   FUNCTION = m_Status
'END FUNCTION
'' ========================================================================================
'#ENDIF   ' //(GDIPVER >= 0x0110)
