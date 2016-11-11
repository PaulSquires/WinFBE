' ########################################################################################
' Microsoft Windows
' File: CGpImageAttributes.inc
' Contents: GDI+ Image attributes class
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

' ########################################################################################
' CGpImageAttributes class
' ########################################################################################

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CGpImageAttributes
   GdipCreateImageAttributes(@m_pImgAttr)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR CGpImageAttributes
   IF m_pImgAttr THEN GdipDisposeImageAttributes(m_pImgAttr)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last status code.
' ========================================================================================
PRIVATE FUNCTION CGpImageAttributes.GetLastStatus () AS GpStatus
   RETURN m_Status
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last status code.
' ========================================================================================
PRIVATE FUNCTION CGpImageAttributes.SetStatus (BYVAL status AS GpStatus) AS GpStatus
   m_Status = status
   RETURN m_Status
END FUNCTION
' ========================================================================================

' =====================================================================================
' Creates a new ImageAttributes object based on this ImageAttributes object.
' =====================================================================================
PRIVATE FUNCTION CGpImageAttributes.Clone (BYVAL pCloneImgAttr AS CGpImageAttributes PTR) AS GpStatus
   IF pCloneImgAttr = NULL THEN RETURN SetStatus(InvalidParameter)
   IF pCloneImgAttr->m_pImgAttr THEN GdipDisposeImageAttributes(pCloneImgAttr->m_pImgAttr)
   RETURN SetStatus(GdipCloneImageAttributes(m_pImgAttr, @pCloneImgAttr->m_pImgAttr))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the color-adjustment matrix of a specified category to identity matrix.
' =====================================================================================
PRIVATE FUNCTION CGpImageAttributes.SetToIdentity (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   RETURN SetStatus(GdipSetImageAttributesToIdentity(m_pImgAttr, nType))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Clears all color- and grayscale-adjustment settings for a specified category.
' =====================================================================================
PRIVATE FUNCTION CGpImageAttributes.Reset (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   RETURN SetStatus(GdipResetImageAttributes(m_pImgAttr, nType))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the color-adjustment matrix for a specified category.
' =====================================================================================
PRIVATE FUNCTION CGpImageAttributes.SetColorMatrix (BYVAL pColorMatrix AS ColorMatrix PTR, BYVAL nMode AS ColorMatrixFlags = ColorMatrixFlagsDefault, BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   RETURN SetStatus(GdipSetImageAttributesColorMatrix(m_pImgAttr, nType, CTRUE, pColorMatrix, NULL, nMode))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Clears the color-adjustment matrix for a specified category.
' =====================================================================================
PRIVATE FUNCTION CGpImageAttributes.ClearColorMatrix (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   RETURN SetStatus(GdipSetImageAttributesColorMatrix(m_pImgAttr, nType, FALSE, NULL, NULL, ColorMatrixFlagsDefault))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the color-adjustment matrix and the grayscale-adjustment matrix for a specified category.
' =====================================================================================
PRIVATE FUNCTION CGpImageAttributes.SetColorMatrices (BYVAL pColorMatrix AS ColorMatrix PTR, BYVAL pGrayMatrix AS ColorMatrix PTR, BYVAL nMode AS ColorAdjustType = ColorMatrixFlagsDefault, BYVAL nType AS LONG = ColorAdjustTypeDefault) AS GpStatus
   RETURN SetStatus(GdipSetImageAttributesColorMatrix(m_pImgAttr, nType, CTRUE, pColorMatrix, pGrayMatrix, nMode))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Clears the color-adjustment matrix and the grayscale-adjustment matrix for a specified category.
' =====================================================================================
PRIVATE FUNCTION CGpImageAttributes.ClearColorMatrices (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   RETURN SetStatus(GdipSetImageAttributesColorMatrix(m_pImgAttr, nType, FALSE, NULL, NULL, ColorMatrixFlagsDefault))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the threshold (transparency range) for a specified category.
' =====================================================================================
PRIVATE FUNCTION CGpImageAttributes.SetThreshold (BYVAL threshold AS SINGLE, BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   RETURN SetStatus(GdipSetImageAttributesThreshold(m_pImgAttr, nType, CTRUE, threshold))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Clears the threshold value for a specified category.
' =====================================================================================
PRIVATE FUNCTION CGpImageAttributes.ClearThreshold (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   RETURN SetStatus(GdipSetImageAttributesThreshold(m_pImgAttr, nType, FALSE, 0.0))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the gamma value for a specified category.
' =====================================================================================
PRIVATE FUNCTION CGpImageAttributes.SetGamma (BYVAL gamma AS SINGLE, BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   RETURN SetStatus(GdipSetImageAttributesGamma(m_pImgAttr, nType, CTRUE, gamma))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Disables gamma correction for a specified category.
' =====================================================================================
PRIVATE FUNCTION CGpImageAttributes.ClearGamma (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   RETURN SetStatus(GdipSetImageAttributesGamma(m_pImgAttr, nType, FALSE, 0.0))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Turns off color adjustment for a specified category. You can call ClearNoOp to reinstate
' the color-adjustment settings that were in place before the call to SetNoOp.
' =====================================================================================
PRIVATE FUNCTION CGpImageAttributes.SetNoOp (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   RETURN SetStatus(GdipSetImageAttributesNoOp(m_pImgAttr, nType, CTRUE))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Clears the NoOp setting for a specified category.
' =====================================================================================
PRIVATE FUNCTION CGpImageAttributes.ClearNoOp (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   RETURN SetStatus(GdipSetImageAttributesNoOp(m_pImgAttr, nType, FALSE))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the color key (transparency range) for a specified category.
' =====================================================================================
PRIVATE FUNCTION CGpImageAttributes.SetColorKey (BYVAL colorLow AS ARGB, BYVAL colorHigh AS ARGB, BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   RETURN SetStatus(GdipSetImageAttributesColorKeys(m_pImgAttr, nType, CTRUE, colorLow, colorHigh))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Clears the color key (transparency range) for a specified category.
' =====================================================================================
PRIVATE FUNCTION CGpImageAttributes.ClearColorKey (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   RETURN SetStatus(GdipSetImageAttributesColorKeys(m_pImgAttr, nType, FALSE, NULL, NULL))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the CMYK output channel for a specified category.
' =====================================================================================
PRIVATE FUNCTION CGpImageAttributes.SetOutputChannel (BYVAL channelFlags AS LONG, BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   m_Status = GdipSetImageAttributesOutputChannel(m_pImgAttr, nType, CTRUE, channelFlags)
   RETURN m_Status
END FUNCTION
' =====================================================================================

' =====================================================================================
' Clears the cyan-magenta-yellow-black (CMYK) output channel setting for a specified category.
' =====================================================================================
PRIVATE FUNCTION CGpImageAttributes.ClearOutputChannel (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   RETURN SetStatus(GdipSetImageAttributesOutputChannel(m_pImgAttr, nType, FALSE, ColorChannelFlagsLast))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the output channel color-profile file for a specified category.
' =====================================================================================
PRIVATE FUNCTION CGpImageAttributes.SetOutputChannelColorProfile (BYVAL pwszColorProfileFilename AS WSTRING PTR, BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   RETURN SetStatus(GdipSetImageAttributesOutputChannelColorProfile(m_pImgAttr, nType, CTRUE, pwszColorProfileFilename))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Clears the output channel color profile setting for a specified category.
' =====================================================================================
PRIVATE FUNCTION CGpImageAttributes.ClearOutputChannelColorProfile (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   RETURN SetStatus(GdipSetImageAttributesOutputChannelColorProfile(m_pImgAttr, nType, FALSE, NULL))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the color-remap table for a specified category.
' // map parameter declared as ANY PTR because in the FB headers
' // it is declared as ColorMap PTR for 32 bit and ColorMap_ PTR for 64 bit
' =====================================================================================
PRIVATE FUNCTION CGpImageAttributes.SetRemapTable (BYVAL mapSize AS UINT, BYVAL map AS ANY PTR, BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   RETURN SetStatus(GdipSetImageAttributesRemapTable(m_pImgAttr, nType, CTRUE, mapSize, map))
END FUNCTION
' =====================================================================================
' =====================================================================================
' Sets the color remap table for the brush category.
' // map parameter declared as ANY PTR because in the FB headers
' // it is declared as ColorMap PTR for 32 bit and ColorMap_ PTR for 64 bit
' =====================================================================================
PRIVATE FUNCTION CGpImageAttributes.SetBrushRemapTable (BYVAL mapSize AS UINT, BYVAL map AS ANY PTR) AS GpStatus
   RETURN SetStatus(this.SetRemapTable(mapSize, map, ColorAdjustTypeBrush))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Clears the color-remap table for a specified category.
' =====================================================================================
PRIVATE FUNCTION CGpImageAttributes.ClearRemapTable (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   RETURN SetStatus(GdipSetImageAttributesRemapTable(m_pImgAttr, nType, FALSE, 0, NULL))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Clears the brush color-remap table of this ImageAttributes object.
' =====================================================================================
PRIVATE FUNCTION CGpImageAttributes.ClearBrushRemapTable () AS GpStatus
   RETURN SetStatus(this.ClearRemapTable(ColorAdjustTypeBrush))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the wrap mode of this ImageAttributes object.
' =====================================================================================
PRIVATE FUNCTION CGpImageAttributes.SetWrapMode (BYVAL nWrap AS WrapMode, BYVAL colour AS ARGB = &HFF000000, BYVAL clamp AS BOOL = FALSE) AS GpStatus
   RETURN SetStatus(GdipSetImageAttributesWrapMode(m_pImgAttr, nWrap, colour, clamp))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Adjusts the colors in a palette according to the adjustment settings of a specified category.
' =====================================================================================
PRIVATE FUNCTION CGpImageAttributes.GetAdjustedPalette (BYVAL pColorPalette AS ColorPalette PTR, BYVAL colorAdjustType AS LONG) AS GpStatus
   RETURN SetStatus(GdipGetImageAttributesAdjustedPalette(m_pImgAttr, pColorPalette, colorAdjustType))
END FUNCTION
' =====================================================================================
