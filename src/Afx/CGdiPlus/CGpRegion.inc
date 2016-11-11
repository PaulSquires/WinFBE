' ########################################################################################
' Microsoft Windows
' File: CGpRegion.inc
' Contents: GDI+ region class.
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

' ########################################################################################
' CGpRegion class
' ########################################################################################

' ========================================================================================
' Constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CGpRegion
   m_Status = GdipCreateRegion(@m_pRegion)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CGpRegion (BYVAL rcf AS GpRectF PTR)
   m_Status = GdipCreateRegionRect(rcf, @m_pRegion)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CGpRegion (BYVAL rc AS GpRect PTR)
   m_Status = GdipCreateRegionRectI(rc, @m_pRegion)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CGpRegion (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE)
   DIM rcf AS GpRectF : rcf.x = x : rcf.y = y : rcf.Width = nwidth : rcf.Height = nHeight
   m_Status = GdipCreateRegionRect(@rcf, @m_pRegion)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CGpRegion (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_)
   DIM rc AS GpRect : rc.x = x : rc.y = y : rc.Width = nwidth : rc.Height = nHeight
   m_Status = GdipCreateRegionRectI(@rc, @m_pRegion)
END CONSTRUCTOR
' ========================================================================================
' =====================================================================================
PRIVATE CONSTRUCTOR CGpRegion (BYVAL pPath AS CGpGraphicsPath PTR)
   m_Status = GdipCreateRegionPath(IIF(pPath, pPath->m_pPath, NULL), @m_pRegion)
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
PRIVATE CONSTRUCTOR CGpRegion (BYVAL regionData AS BYTE PTR, BYVAL nSize AS LONG)
   m_Status = GdipCreateRegionRgnData(regionData, nSize, @m_pRegion)
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
PRIVATE CONSTRUCTOR CGpRegion (BYVAL hRgn AS HRGN)
   m_Status = GdipCreateRegionHrgn(hRgn, @m_pRegion)
END CONSTRUCTOR
' =====================================================================================

' ========================================================================================
' Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR CGpRegion
   IF m_pRegion THEN GdipDeleteRegion(m_pRegion)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last status code.
' ========================================================================================
PRIVATE FUNCTION CGpRegion.GetLastStatus () AS GpStatus
   RETURN m_Status
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last status code.
' ========================================================================================
PRIVATE FUNCTION CGpRegion.SetStatus (BYVAL status AS GpStatus) AS GpStatus
   m_Status = status
   RETURN m_Status
END FUNCTION
' ========================================================================================

' =====================================================================================
' * Creates a new Region object based on this Region object.
' =====================================================================================
PRIVATE FUNCTION CGpRegion.Clone (BYVAL pRegion AS CGpRegion PTR) AS GpStatus
   IF pRegion = NULL THEN RETURN SetStatus(InvalidParameter)
   IF pRegion->m_pRegion THEN GdipDeleteRegion(pRegion->m_pRegion)
   RETURN SetStatus(GdipCloneRegion(m_pRegion, @pRegion->m_pRegion))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Updates this region to an infinite region.
' =====================================================================================
PRIVATE FUNCTION CGpRegion.MakeInfinite () AS GpStatus
   RETURN SetStatus(GdipSetInfinite(m_pRegion))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Updates this region to an empty region. In other words, the region occupies no space
' on the display device.
' =====================================================================================
PRIVATE FUNCTION CGpRegion.MakeEmpty () AS GpStatus
   RETURN SetStatus(GdipSetEmpty(m_pRegion))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Updates this region to the portion of itself that intersects the specified rectangle's interior.
' =====================================================================================
PRIVATE FUNCTION CGpRegion.Intersect (BYVAL rcf AS GpRectF PTR) AS GpStatus
   RETURN SetStatus(GdipCombineRegionRect(m_pRegion, rcf, CombineModeIntersect))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpRegion.Intersect (BYVAL rc AS GpRect PTR) AS GpStatus
   RETURN SetStatus(GdipCombineRegionRectI(m_pRegion, rc, CombineModeIntersect))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Updates this region to the portion of itself that intersects the specified path's interior.
' =====================================================================================
PRIVATE FUNCTION CGpRegion.Intersect (BYVAL pPath AS CGpGraphicsPath PTR) AS GpStatus
   RETURN SetStatus(GdipCombineRegionPath(m_pRegion, pPath->m_pPath, CombineModeIntersect))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Updates this region to the portion of itself that intersects another region.
' =====================================================================================
PRIVATE FUNCTION CGpRegion.Intersect (BYVAL pRegion AS CGpRegion PTR) AS GpStatus
   RETURN SetStatus(GdipCombineRegionRegion(m_pRegion, IIF(pRegion, pRegion->m_pRegion, NULL), CombineModeIntersect))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Updates this region to all portions (intersecting and nonintersecting) of itself and
' all portions of the specified rectangle's interior.
' =====================================================================================
PRIVATE FUNCTION CGpRegion.Union_ (BYVAL rc AS GpRectF PTR) AS GpStatus
   RETURN SetStatus(GdipCombineRegionRect(m_pRegion, rc, CombineModeUnion))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpRegion.Union_ (BYVAL rc AS GpRect PTR) AS GpStatus
   RETURN SetStatus(GdipCombineRegionRectI(m_pRegion, rc, CombineModeUnion))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Updates this region to all portions (intersecting and nonintersecting) of itself and
' all portions of the specified path's interior.
' =====================================================================================
PRIVATE FUNCTION CGpRegion.Union_ (BYVAL pPath AS CGpGraphicsPath PTR) AS GpStatus
   RETURN SetStatus(GdipCombineRegionPath(m_pRegion, IIF(pPath, pPath->m_pPath, NULL), CombineModeUnion))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Updates this region to all portions (intersecting and nonintersecting) of itself and
' all portions of another region.
' =====================================================================================
PRIVATE FUNCTION CGpRegion.Union_ (BYVAL pRegion AS CGpRegion PTR) AS GpStatus
   RETURN SetStatus(GdipCombineRegionRegion(m_pRegion, IIF(pRegion, pRegion->m_pRegion, NULL), CombineModeUnion))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Updates this region to the nonintersecting portions of itself and the specified rectangle's interior.
' =====================================================================================
PRIVATE FUNCTION CGpRegion.Xor_ (BYVAL rc AS GpRectF PTR) AS GpStatus
   RETURN SetStatus(GdipCombineRegionRect(m_pRegion, rc, CombineModeXor))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpRegion.Xor_ (BYVAL rc AS GpRect PTR) AS GpStatus
   RETURN SetStatus(GdipCombineRegionRectI(m_pRegion, rc, CombineModeXor))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpRegion.Xor_ (BYVAL pPath AS CGpGraphicsPath PTR) AS GpStatus
   RETURN SetStatus(GdipCombineRegionPath(m_pRegion, IIF(pPath, pPath->m_pPath, NULL), CombineModeXor))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpRegion.Xor_ (BYVAL pRegion AS CGpRegion PTR) AS GpStatus
   RETURN SetStatus(GdipCombineRegionRegion(m_pRegion, IIF(pRegion, pRegion->m_pRegion, NULL), CombineModeXor))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Updates this region to the portion of itself that does not intersect the specified rectangle's interior.
' =====================================================================================
PRIVATE FUNCTION CGpRegion.Exclude (BYVAL rc AS GpRectF PTR) AS GpStatus
   RETURN SetStatus(GdipCombineRegionRect(m_pRegion, rc, CombineModeExclude))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpRegion.Exclude (BYVAL rc AS GpRect PTR) AS GpStatus
   RETURN SetStatus(GdipCombineRegionRectI(m_pRegion, rc, CombineModeExclude))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Updates this region to the portion of itself that does not intersect the specified path's interior.
' =====================================================================================
PRIVATE FUNCTION CGpRegion.Exclude (BYVAL pPath AS CGpGraphicsPath PTR) AS GpStatus
   RETURN SetStatus(GdipCombineRegionPath(m_pRegion, IIF(pPath, pPath->m_pPath, NULL), CombineModeExclude))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Updates this region to the portion of itself that does not intersect another region.
' =====================================================================================
PRIVATE FUNCTION CGpRegion.Exclude (BYVAL pRegion AS CGpRegion PTR) AS GpStatus
   RETURN SetStatus(GdipCombineRegionRegion(m_pRegion, IIF (pRegion, pRegion->m_pRegion, NULL), CombineModeExclude))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Updates this region to the portion of itself that does not intersect the specified rectangle's interior.
' =====================================================================================
PRIVATE FUNCTION CGpRegion.Complement (BYVAL rcf AS GpRectF PTR) AS GpStatus
   RETURN SetStatus(GdipCombineRegionRect(m_pRegion, rcf, CombineModeComplement))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpRegion.Complement (BYVAL rc AS GpRect PTR) AS GpStatus
   RETURN SetStatus(GdipCombineRegionRectI(m_pRegion, rc, CombineModeComplement))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Updates this region to the portion of itself that does not intersect the specified path's interior.
' =====================================================================================
PRIVATE FUNCTION CGpRegion.Complement (BYVAL pPath AS CGpGraphicsPath PTR) AS GpStatus
   RETURN SetStatus(GdipCombineRegionPath(m_pRegion, IIF(pPath, pPath->m_pPath, NULL), CombineModeComplement))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Updates this region to the portion of itself that does not intersect another region.
' =====================================================================================
PRIVATE FUNCTION CGpRegion.Complement (BYVAL pRegion AS CGpRegion PTR) AS GpStatus
   RETURN SetStatus(GdipCombineRegionRegion(m_pRegion, IIF(pRegion, pRegion->m_pRegion, NULL), CombineModeComplement))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Offsets this region by specified amounts in the horizontal and vertical directions.
' =====================================================================================
PRIVATE FUNCTION CGpRegion.Translate (BYVAL dx AS SINGLE, BYVAL dy AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipTranslateRegion(m_pRegion, dx, dy))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpRegion.Translate (BYVAL dx AS INT_, BYVAL dy AS INT_) AS GpStatus
   RETURN SetStatus(GdipTranslateRegionI(m_pRegion, dx, dy))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Transforms this region by multiplying each of its data points by a specified matrix.
' =====================================================================================
PRIVATE FUNCTION CGpRegion.Transform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
   RETURN SetStatus(GdipTransformRegion(m_pRegion, IIF(pMatrix, pMatrix->m_pMatrix, NULL)))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Gets a rectangle that encloses this region.
' =====================================================================================
PRIVATE FUNCTION CGpRegion.GetBounds (BYVAL rc AS GpRectF PTR, BYVAL pGraphics AS CGpGraphics PTR) AS GpStatus
   RETURN SetStatus(GdipGetRegionBounds(m_pRegion, IIF(pGraphics, pGraphics->m_pGraphics, NULL), rc))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpRegion.GetBounds (BYVAL rc AS GpRect PTR, BYVAL pGraphics AS CGpGraphics PTR) AS GpStatus
   RETURN SetStatus(GdipGetRegionBoundsI(m_pRegion, IIF(pGraphics, pGraphics->m_pGraphics, NULL), rc))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Creates a GDI region from this region.
' =====================================================================================
PRIVATE FUNCTION CGpRegion.GetHRGN (BYVAL pGraphics AS CGpGraphics PTR) AS HRGN
   DIM hrgn AS HRGN
   m_Status = GdipGetRegionHRgn(m_pRegion, IIF(pGraphics, pGraphics->m_pGraphics, NULL), @hrgn)
   RETURN hrgn
END FUNCTION
' =====================================================================================

' =====================================================================================
' Determines whether this region is empty.
' =====================================================================================
PRIVATE FUNCTION CGpRegion.IsEmpty (BYVAL pGraphics AS CGpGraphics PTR) AS BOOLEAN
   DIM booln AS BOOL
   m_Status = GdipIsEmptyRegion(m_pRegion, IIF(pGraphics, pGraphics->m_pGraphics, NULL), @booln)
   RETURN booln
END FUNCTION
' =====================================================================================

' =====================================================================================
' Determines whether this region is infinite.
' =====================================================================================
PRIVATE FUNCTION CGpRegion.IsInfinite (BYVAL pGraphics AS CGpGraphics PTR) AS BOOLEAN
   DIM booln AS BOOL
   m_Status = GdipIsInfiniteRegion(m_pRegion, IIF(pGraphics, pGraphics->m_pGraphics, NULL), @booln)
   RETURN booln
END FUNCTION
' =====================================================================================

' =====================================================================================
' Determines whether this region is equal to a specified region.
' =====================================================================================
PRIVATE FUNCTION CGpRegion.Equals (BYVAL pRegion AS CGpRegion PTR, BYVAL pGraphics AS CGpGraphics PTR) AS BOOLEAN
   DIM booln AS BOOL
   m_Status = GdipIsEqualRegion(m_pRegion, IIF(pRegion, pRegion->m_pRegion, NULL), IIF(pGraphics, pGraphics->m_pGraphics, NULL), @booln)
   RETURN booln
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets number of bytes of data that describes this region.
' =====================================================================================
PRIVATE FUNCTION CGpRegion.GetDataSize () AS UINT
   DIM bufferSize AS UINT
   m_Status = GdipGetRegionDataSize(m_pRegion, @bufferSize)
   RETURN bufferSize
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets data that describes this region.
' =====================================================================================
PRIVATE FUNCTION CGpRegion.GetData (BYVAL buffer AS BYTE PTR, BYVAL bufferSize AS UINT, BYVAL sizeFilled AS UINT PTR) AS GpStatus
   RETURN SetStatus(GdipGetRegionData(m_pRegion, buffer, bufferSize, sizeFilled))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Determines whether a point is inside this region.
' =====================================================================================
PRIVATE FUNCTION CGpRegion.IsVisible (BYVAL pt AS GpPointF PTR, BYVAL pGraphics AS CGpGraphics PTR) AS GpStatus
   DIM booln AS BOOL
   m_Status = GdipIsVisibleRegionPoint(m_pRegion, pt->x, pt->y, IIF(pGraphics, pGraphics->m_pGraphics, NULL), @booln)
   RETURN booln
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpRegion.IsVisible (BYVAL pt AS GpPoint PTR, BYVAL pGraphics AS CGpGraphics PTR) AS GpStatus
   DIM booln AS BOOL
   m_Status = GdipIsVisibleRegionPointI(m_pRegion, pt->x, pt->y, IIF(pGraphics, pGraphics->m_pGraphics, NULL), @booln)
   RETURN booln
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpRegion.IsVisible (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL pGraphics AS CGpGraphics PTR) AS GpStatus
   DIM booln AS BOOL
   m_Status = GdipIsVisibleRegionPoint(m_pRegion, x, y, IIF(pGraphics, pGraphics->m_pGraphics, NULL), @booln)
   RETURN booln
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpRegion.IsVisible (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL pGraphics AS CGpGraphics PTR) AS GpStatus
   DIM booln AS BOOL
   m_Status = GdipIsVisibleRegionPointI(m_pRegion, x, y, IIF(pGraphics, pGraphics->m_pGraphics, NULL), @booln)
   RETURN booln
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpRegion.IsVisible (BYVAL rc AS GpRectF PTR, BYVAL pGraphics AS CGpGraphics PTR) AS GpStatus
   DIM booln AS BOOL
   m_Status = GdipIsVisibleRegionRect(m_pRegion, rc->x, rc->y, rc->Width, rc->Height, IIF(pGraphics, pGraphics->m_pGraphics, NULL), @booln)
   RETURN booln
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpRegion.IsVisible (BYVAL rc AS GpRect PTR, BYVAL pGraphics AS CGpGraphics PTR) AS GpStatus
   DIM booln AS BOOL
   m_Status = GdipIsVisibleRegionRectI(m_pRegion, rc->x, rc->y, rc->Width, rc->Height, IIF(pGraphics, pGraphics->m_pGraphics, NULL), @booln)
   RETURN booln
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpRegion.IsVisible (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE, BYVAL pGraphics AS CGpGraphics PTR) AS GpStatus
   DIM booln AS BOOL
   m_Status = GdipIsVisibleRegionRect(m_pRegion, x, y, nWidth, nHeight, IIF(pGraphics, pGraphics->m_pGraphics, NULL), @booln)
   RETURN booln
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpRegion.IsVisible (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL pGraphics AS CGpGraphics PTR) AS GpStatus
   DIM booln AS BOOL
   m_Status = GdipIsVisibleRegionRectI(m_pRegion, x, y, nWidth, nHeight, IIF (pGraphics, pGraphics->m_pGraphics, NULL), @booln)
   RETURN booln
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Gets the number of rectangles that approximate this region. The region is transformed
' by a specified matrix before the rectangles are calculated.
' =====================================================================================
PRIVATE FUNCTION CGpRegion.GetRegionScansCount (BYVAL pMatrix AS CGpMatrix PTR) AS UINT
   DIM nCount AS UINT
   m_Status = GdipGetRegionScansCount(m_pRegion, @nCount, IIF(pMatrix, pMatrix->m_pMatrix, NULL))
   RETURN nCount
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Gets the number of rectangles that approximate this region. The region is transformed
' by a specified matrix before the rectangles are calculated.
' =====================================================================================
PRIVATE FUNCTION CGpRegion.GetRegionScans (BYVAL pMatrix AS CGpMatrix PTR, BYVAL rects AS GpRectF PTR, BYVAL nCount AS INT_ PTR) AS GpStatus
   RETURN SetStatus(GdipGetRegionScans(m_pRegion, rects, nCount, IIF(pMatrix, pMatrix->m_pMatrix, NULL)))
END FUNCTION
' =====================================================================================
