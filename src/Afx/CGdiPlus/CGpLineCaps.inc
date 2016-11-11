' ########################################################################################
' Microsoft Windows
' File: CGpLineCaps.inc
' Contents: GDI+ line cap classes.
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

' ########################################################################################
' CGpCustomLineCap Class
' ########################################################################################

' ========================================================================================
' * Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CGpCustomLineCap
END CONSTRUCTOR
' ========================================================================================

' =====================================================================================
' * Creates a CustomLineCap object from a fill path and a stroke path.
' =====================================================================================
PRIVATE CONSTRUCTOR CGpCustomLineCap (BYVAL pFillPath AS CGpGraphicsPath PTR, BYVAL pStrokePath AS CGpGraphicsPath PTR, BYVAL baseCap AS LineCap = LinecapFLat, BYVAL baseInset AS SINGLE = 0.0)
   m_Status = GdipCreateCustomLineCap(IIF(pFillPath, pFillPath->m_pPath, NULL), IIF(pStrokePath, pStrokePath->m_pPath, NULL), baseCap, baseinset, @m_pCustomLineCap)
END CONSTRUCTOR
' =====================================================================================

' ========================================================================================
' * Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR CGpCustomLineCap
   IF m_pCustomLineCap THEN GdipDeleteCustomLineCap(m_pCustomLineCap)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last status code.
' ========================================================================================
PRIVATE FUNCTION CGpCustomLineCap.GetLastStatus () AS GpStatus
   RETURN m_Status
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last status code.
' ========================================================================================
PRIVATE FUNCTION CGpCustomLineCap.SetStatus (BYVAL status AS GpStatus) AS GpStatus
   m_Status = status
   RETURN m_Status
END FUNCTION
' ========================================================================================

' =====================================================================================
' * Sets the LineCap object used to start and end lines within the GraphicsPath object
' that defines this CustomLineCap object.
' =====================================================================================
PRIVATE FUNCTION CGpCustomLineCap.SetStrokeCaps (BYVAL startCap AS LineCap, BYVAL endCap AS LineCap) AS GpStatus
   RETURN SetStatus(GdipSetCustomLineCapStrokeCaps(m_pCustomLineCap, startCap, endCap))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Gets the end cap styles for both the start line cap and the end line cap. Line caps
' are LineCap objects that end the individual lines within a path.
' =====================================================================================
PRIVATE FUNCTION CGpCustomLineCap.GetStrokeCaps (BYVAL startCap AS LineCap PTR, BYVAL endCap AS LineCap PTR) AS GpStatus
   RETURN SetStatus(GdipGetCustomLineCapStrokeCaps(m_pCustomLineCap, startCap, endCap))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets the LineCap object used to start and end lines within the GraphicsPath object
' that defines this CustomLineCap object.
' =====================================================================================
PRIVATE FUNCTION CGpCustomLineCap.SetStrokeCap (BYVAL strokeCap AS LineCap) AS GpStatus
   RETURN SetStatus(GdipSetCustomLineCapStrokeCaps(m_pCustomLineCap, strokeCap, strokeCap))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets the style of line join for the stroke. The line join specifies how two lines that
' intersect within the GraphicsPath object that makes up the custom line cap are joined.
' =====================================================================================
PRIVATE FUNCTION CGpCustomLineCap.SetStrokeJoin (BYVAL nLineJoin AS LineJoin) AS GpStatus
   RETURN SetStatus(GdipSetCustomLineCapStrokeJoin(m_pCustomLineCap, nLineJoin))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Returns the style of LineJoin used to join multiple lines in the same GraphicsPath object.
' =====================================================================================
PRIVATE FUNCTION CGpCustomLineCap.GetStrokeJoin () AS LineJoin
   DIM nLineJoin AS LineJoin
   m_Status = GdipGetCustomLineCapStrokeJoin(m_pCustomLineCap, @nLineJoin)
   FUNCTION = nLineJoin
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets the LineCap that appears as part of this CustomLineCap at the end of a line.
' =====================================================================================
PRIVATE FUNCTION CGpCustomLineCap.SetBaseCap (BYVAL baseCap AS LineCap) AS GpStatus
   RETURN SetStatus(GdipSetCustomLineCapBaseCap(m_pCustomLineCap, baseCap))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Gets the style of the base cap. The base cap is a LineCap object used as a cap at the
' end of a line along with this CustomLineCap object.
' =====================================================================================
PRIVATE FUNCTION CGpCustomLineCap.GetBaseCap () AS LineCap
   DIM baseCap AS LineCap
   m_Status = GdipGetCustomLineCapBaseCap(m_pCustomLineCap, @baseCap)
   RETURN baseCap
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets the base inset value of this custom line cap. This is the distance between the
' end of a line and the base cap.
' =====================================================================================
PRIVATE FUNCTION CGpCustomLineCap.SetBaseInset (BYVAL inset AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipSetCustomLineCapBaseInset(m_pCustomLineCap, inset))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Returns the distance between the base cap to the start of the line.
' =====================================================================================
PRIVATE FUNCTION CGpCustomLineCap.GetBaseInset () AS SINGLE
   DIM inset AS SINGLE
   m_Status = GdipGetCustomLineCapBaseInset(m_pCustomLineCap, @inset)
   RETURN inset
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets the value of the scale width. This is the amount to scale the custom line cap
' relative to the width of the Pen used to draw lines. The default value of 1.0 does
' not scale the line cap.
' =====================================================================================
PRIVATE FUNCTION CGpCustomLineCap.SetWidthScale (BYVAL widthScale AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipSetCustomLineCapWidthScale(m_pCustomLineCap, widthScale))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Gets the value of the scale width. This is the amount to scale the custom line cap
' relative to the width of the Pen object used to draw a line. The default value of 1.0
' does not scale the line cap.
' =====================================================================================
PRIVATE FUNCTION CGpCustomLineCap.GetWidthScale () AS SINGLE
   DIM widthScale AS SINGLE
   m_Status = GdipGetCustomLineCapWidthScale(m_pCustomLineCap, @widthScale)
   RETURN widthScale
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Creates a new CustomLineCap object based on this CustomLineCap object.
' Note: This function differs from the C++ one, which returns a pointer to a
' CGpCustomLineCap created with NEW. Instead, we have added an empty constructor and
' we will pass the type created with this destructor to this function, e.g.
'   DIM secondCap AS CGpCustomLineCap
'   firstCap.Clone(@secondCap)
' The advantage is that we won't have to worry to free the returned CGpCustomLineCap,
' as we will have to do if we return a pointer to a CGpCustomLineCap created with NEW.
' =====================================================================================
PRIVATE FUNCTION CGpCustomLineCap.Clone (BYVAL pCustomLineCap AS CGpCustomLineCap PTR) AS GpStatus
   IF pCustomLineCap = NULL THEN RETURN SetStatus(InvalidParameter)
   IF pCustomLineCap->m_pCustomLineCap THEN GdipDeleteCustomLineCap(pCustomLineCap->m_pCustomLineCap)
   RETURN SetStatus(GdipCloneCustomLineCap(m_pCustomLineCap, @pCustomLineCap->m_pCustomLineCap))
END FUNCTION
' =====================================================================================


' #####################################################################################
' CGpAdjustableArrowCap class
' Inherits from CGpCustomLineCap.
' #####################################################################################

' =====================================================================================
' * Creates an adjustable arrow line cap with the specified height and width. The arrow
' line cap can be filled or nonfilled. The middle inset defaults to zero.
' =====================================================================================
PRIVATE CONSTRUCTOR CGpAdjustableArrowCap (BYVAL nHeight AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL bIsFilled AS BOOL = CTRUE)
   DIM pAdjustableArrowCap AS GpAdjustableArrowCap PTR
   m_Status = GdipCreateAdjustableArrowCap(nHeight, nWidth, bIsFilled, @pAdjustableArrowCap)
   m_pCustomLineCap = cast(GpCustomLineCap PTR, pAdjustableArrowCap)
END CONSTRUCTOR
' =====================================================================================

' =====================================================================================
' * Sets the height of the arrow cap. This is the distance from the base of the arrow to its vertex.
' =====================================================================================
PRIVATE FUNCTION CGpAdjustableArrowCap.SetHeight (BYVAL nHeight AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipSetAdjustableArrowCapHeight(cast(GpAdjustableArrowCap PTR, m_pCustomLineCap), nHeight))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Gets the height of the arrow cap. This is the distance from the base of the arrow to its vertex.
' =====================================================================================
PRIVATE FUNCTION CGpAdjustableArrowCap.GetHeight () AS SINGLE
   DIM nHeight AS SINGLE
   m_Status = GdipGetAdjustableArrowCapHeight(cast(GpAdjustableArrowCap PTR, m_pCustomLineCap), @nHeight)
   RETURN nHeight
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets the width of the arrow cap. The width is the distance between the endpoints of the base of the arrow.
' =====================================================================================
PRIVATE FUNCTION CGpAdjustableArrowCap.SetWidth (BYVAL nWidth AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipSetAdjustableArrowCapWidth(cast(GpAdjustableArrowCap PTR, m_pCustomLineCap), nWidth))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Gets the width of the arrow cap. This is the distance from the base of the arrow to its vertex.
' =====================================================================================
PRIVATE FUNCTION CGpAdjustableArrowCap.GetWidth () AS SINGLE
   DIM nWidth AS SINGLE
   m_Status = GdipGetAdjustableArrowCapWidth(cast(GpAdjustableArrowCap PTR, m_pCustomLineCap), @nWidth)
   FUNCTION = nWidth
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets the number of units that the midpoint of the base shifts towards the vertex.
' =====================================================================================
PRIVATE FUNCTION CGpAdjustableArrowCap.SetMiddleInset (BYVAL middleInset AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipSetAdjustableArrowCapMiddleInset(cast(GpAdjustableArrowCap PTR, m_pCustomLineCap), middleInset))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Gets the value of the inset. The middle inset is the number of units that the midpoint
' of the base shifts towards the vertex.
' =====================================================================================
PRIVATE FUNCTION CGpAdjustableArrowCap.GetMiddleInset () AS SINGLE
   DIM middleInset AS SINGLE
   m_Status = GdipGetAdjustableArrowCapMiddleInset(cast(GpAdjustableArrowCap PTR, m_pCustomLineCap), @middleInset)
   RETURN middleInset
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Sets the fill state of the arrow cap. If the arrow cap is not filled, only the outline is drawn.
' =====================================================================================
PRIVATE FUNCTION CGpAdjustableArrowCap.SetFillState (BYVAL bIsFilled AS BOOL) AS GpStatus
   RETURN SetStatus(GdipSetAdjustableArrowCapFillState(cast(GpAdjustableArrowCap PTR, m_pCustomLineCap), bIsFilled))
END FUNCTION
' =====================================================================================

' =====================================================================================
' * Determines whether the arrow cap is filled.
' =====================================================================================
PRIVATE FUNCTION CGpAdjustableArrowCap.IsFilled () AS BOOLEAN
   DIM bIsFilled AS BOOL
   m_Status = GdipGetAdjustableArrowCapFillState(cast(GpAdjustableArrowCap PTR, m_pCustomLineCap), @bIsFilled)
   RETURN bIsFilled
END FUNCTION
' =====================================================================================
