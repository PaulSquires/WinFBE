' ########################################################################################
' Microsoft Windows
' File: CGpFont.inc
' Contents: GDI+ Font class
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

' ########################################################################################
' CGpFont Class
' ########################################################################################

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CGpFont
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a Font object based on the GDI font object that is currently selected into a
' specified device context. This constructor is provided for compatibility with GDI.
' ========================================================================================
PRIVATE CONSTRUCTOR CGpFont (BYVAL hdc AS HDC)
   m_Status = GdipCreateFontFromDC(hdc, @m_pFont)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a Font object indirectly from a Windows Graphics Device Interface (GDI)
' logical font by using a handle to a GDI LOGFONT structure.
' Note: In the FB headers for 32 bit, GdipCreateFontFromLogfontA and GdipCreateFontFromLogfontW
' use only the name GdipCreateFontFromLogfont.
' ========================================================================================
PRIVATE CONSTRUCTOR CGpFont (BYVAL hdc AS HDC, BYVAL hFont AS HFONT)
#ifdef __FB_64BIT__
   DIM lfA AS LOGFONTA
   IF hFont THEN
      IF GetObjectA(hFont, SIZEOF(LOGFONTA), @lfA) THEN
         m_Status = GdipCreateFontFromLogfontA(hdc, @lfA, @m_pFont)
      ELSE
         m_Status = GdipCreateFontFromDC(hdc, @m_pFont)
      END IF
   ELSE
      m_Status = GdipCreateFontFromDC(hdc, @m_pFont)
   END IF
#else
   DIM lf AS LOGFONT
   IF hFont THEN
      IF GetObject(hFont, SIZEOF(LOGFONT), @lf) THEN
         m_Status = GdipCreateFontFromLogfont(hdc, @lf, @m_pFont)
      ELSE
         m_Status = GdipCreateFontFromDC(hdc, @m_pFont)
      END IF
   ELSE
      m_Status = GdipCreateFontFromDC(hdc, @m_pFont)
   END IF
#endif
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a Font object directly from a Windows Graphics Device Interface (GDI) logical
' font. The GDI logical font is a LOGFONTW structure, which is the one-byte character
' version of a logical font. This constructor is provided for compatibility with GDI.
' Note: Because of a mess in the GdiPlus headers for 32 bit, it's not possible to call
' GdipCreateFontFromLogfontA and GdipCreateFontFromLogfontW, only GdipCreateFontFromLogfont.
' ========================================================================================
#ifdef __FB_64BIT__
PRIVATE CONSTRUCTOR CGpFont (BYVAL hdc AS HDC, BYVAL plogfont AS LOGFONTW PTR)
   IF plogfont <> NULL THEN
      m_Status = GdipCreateFontFromLogfontW(hdc, plogfont, @m_pFont)
   ELSE
      m_Status = GdipCreateFontFromDC(hdc, @m_pFont)
   END IF
END CONSTRUCTOR
#endif
' ========================================================================================

' ========================================================================================
' Creates a Font object directly from a Windows Graphics Device Interface (GDI) logical
' font. The GDI logical font is a LOGFONTA structure, which is the one-byte character
' version of a logical font. This constructor is provided for compatibility with GDI.
' ========================================================================================
#ifdef __FB_64BIT__
PRIVATE CONSTRUCTOR CGpFont (BYVAL hdc AS HDC, BYVAL plogfont AS LOGFONTA PTR)
   IF plogfont <> NULL THEN
      m_Status = GdipCreateFontFromLogfontA(hdc, plogfont, @m_pFont)
   ELSE
      m_Status = GdipCreateFontFromDC(hdc, @m_pFont)
   END IF
END CONSTRUCTOR
#endif
' ===========================================================================================

' ========================================================================================
' Creates a Font object based on a FontFamily object, a size, a font style, and a unit of measurement.
' ========================================================================================
PRIVATE CONSTRUCTOR CGpFont (BYVAL pFamily AS CGpFontFamily PTR, BYVAL emSize AS SINGLE, BYVAL nStyle AS INT_ = 0, BYVAL nUnit AS GpUnit = 0)
   m_Status = GdipCreateFont(IIF(pFamily, pFamily->m_pFontFamily, NULL), emSize, nStyle, nUnit, @m_pFont)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a Font object based on a font family, a size, a font style, a unit of
' measurement, and a FontCollection object.
' ========================================================================================
PRIVATE CONSTRUCTOR CGpFont (BYVAL pwszFamilyName AS WSTRING PTR, BYVAL emSize AS SINGLE, BYVAL nStyle AS INT_ = 0, BYVAL nUnit AS GpUnit = 0, BYVAL pFontCollection AS CGpFontCollection PTR = NULL)
   DIM pFamily AS CGpFontFamily = CGpFontFamily(pwszFamilyName, pFontCollection)
   IF pFamily.GetLastStatus = Ok THEN
      m_Status = GdipCreateFont(pFamily.m_pFontFamily, emSize, nStyle, nUnit, @m_pFont)
   ELSE
      m_Status = GdipGetGenericFontFamilySansSerif(@pFamily.m_pFontFamily)
      IF m_Status = Ok THEN m_Status = GdipCreateFont(pFamily.m_pFontFamily, emSize, nStyle, nUnit, @m_pFont)
   END IF
END CONSTRUCTOR
' ========================================================================================

' ===========================================================================================
' Cleanup
' ===========================================================================================
PRIVATE DESTRUCTOR CGpFont
   IF m_pFont THEN GdipDeleteFont(m_pFont)
END DESTRUCTOR
' ===========================================================================================

' ===========================================================================================
' Returns the last status code.
' ===========================================================================================
PRIVATE FUNCTION CGpFont.GetLastStatus () AS GpStatus
   RETURN m_Status
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Sets the last status code.
' ===========================================================================================
PRIVATE FUNCTION CGpFont.SetStatus (BYVAL status AS GpStatus) AS GpStatus
   m_Status = status
   RETURN m_Status
END FUNCTION
' ===========================================================================================

' ========================================================================================
' The GetLogFontA method uses a LOGFONTA structure to get the attributes of this Font object.
' Note: Because of a mess in the FB GDI+ headers for 32 bit, it's not possible to call
' GdipGetLogFontA and GdipGetLogFont, only GdipGetLogFont.
' ========================================================================================
#ifdef __FB_64BIT__
PRIVATE FUNCTION CGpFont.GetLogFontA (BYVAL pGraphics AS CGpGraphics PTR, BYVAL pfA AS LOGFONTA PTR) AS GpStatus
   RETURN SetStatus(GdipGetLogFontA(m_pFont, IIF(pGraphics, pGraphics->m_pGraphics, NULL), pfA))
END FUNCTION
' ========================================================================================
' ========================================================================================
' The GetLogFontW method uses a LOGFONTW structure to get the attributes of this Font object.
' ========================================================================================
PRIVATE FUNCTION CGpFont.GetLogFontW (BYVAL pGraphics AS CGpGraphics PTR, BYVAL pfW AS LOGFONTW PTR) AS GpStatus
   RETURN SetStatus(GdipGetLogFontW(m_pFont, IIF(pGraphics, pGraphics->m_pGraphics, NULL), pfW))
END FUNCTION
' ========================================================================================
#endif

' ========================================================================================
' Creates a new Font object based on this Font object.
' ========================================================================================
PRIVATE FUNCTION CGpFont.Clone (BYVAL pCloneFont AS CGpFont PTR) AS GpStatus
   IF pCloneFont = NULL THEN RETURN SetStatus(InvalidParameter)
   IF pCloneFont->m_pFont THEN GdipDeleteFont(pCloneFont->m_pFont)
   RETURN SetStatus(GdipCloneFont(m_pFont, @pCloneFont->m_pFont))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The IsAvailable method determines whether this Font object was created successfully.
' ========================================================================================
PRIVATE FUNCTION CGpFont.IsAvailable () AS BOOLEAN
   RETURN m_pFont <> NULL
END FUNCTION
' ========================================================================================

' ========================================================================================
' The GetFamily method gets the font family on which this font is based.
' ========================================================================================
PRIVATE FUNCTION CGpFont.GetFamily (BYVAL pFamily AS CGpFontFamily PTR) AS GpStatus
   IF pFamily = NULL THEN RETURN SetStatus(InvalidParameter)
   RETURN SetStatus(GdipGetFamily(m_pFont, @pFamily->m_pFontFamily))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The GetStyle method gets the style of this font's typeface.
' ========================================================================================
PRIVATE FUNCTION CGpFont.GetStyle () AS INT_
   DIM nStyle AS INT_
   m_Status = GdipGetFontStyle(m_pFont, @nStyle)
   RETURN nStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' The GetSize method returns the font size (commonly called the em size) of this
' Font object. The size is in the units of this Font object.
' ========================================================================================
PRIVATE FUNCTION CGpFont.GetSize () AS SINGLE
   DIM nSize AS SINGLE
   m_Status = GdipGetFontSize(m_pFont, @nSize)
   RETURN nSize
END FUNCTION
' ========================================================================================

' ========================================================================================
' The GetUnit method returns the unit of measure of this Font object.
' ========================================================================================
PRIVATE FUNCTION CGpFont.GetUnit () AS GpUnit
   DIM nUnit AS GpUnit
   m_Status = GdipGetFontUnit(m_pFont, @nUnit)
   RETURN nUnit
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the line spacing of this font in the current unit of a specified Graphics object.
' The line spacing is the vertical distance between the base lines of two consecutive
' lines of text. Thus, the line spacing includes the blank space between lines along
' with the height of the character itself.
' ========================================================================================
PRIVATE FUNCTION CGpFont.GetHeight (BYVAL pGraphics AS CGpGraphics PTR) AS SINGLE
   DIM nHeight AS SINGLE
   m_Status = GdipGetFontHeight(m_pFont, IIF(pGraphics, pGraphics->m_pGraphics, NULL), @nHeight)
   RETURN nHeight
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the line spacing, in pixels, of this font. The line spacing is the vertical
' distance between the base lines of two consecutive lines of text. Thus, the line
' spacing includes the blank space between lines along with the height of the character
' itself.
' ========================================================================================
PRIVATE FUNCTION CGpFont.GetHeight (BYVAL dpi AS SINGLE) AS SINGLE
   DIM nHeight AS SINGLE
   m_Status = GdipGetFontHeightGivenDPI(m_pFont, dpi, @nHeight)
   RETURN nHeight
END FUNCTION
' ========================================================================================

' ########################################################################################

' ########################################################################################
' CGpFontFamily Class
' ########################################################################################

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CGpFontFamily
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a FontFamily object based on a specified font collection.
' ========================================================================================
PRIVATE CONSTRUCTOR CGpFontFamily (BYVAL pwszName AS WSTRING PTR, BYVAL pFontCollection AS CGpFontCollection PTR = NULL)
   m_Status = GdipCreateFontFamilyFromName(pwszName, IIF(pFontCollection, pFontCollection->m_pFontCollection, NULL), @m_pFontFamily)
END CONSTRUCTOR
' ========================================================================================

' =====================================================================================
' Cleanup
' =====================================================================================
PRIVATE DESTRUCTOR CGpFontFamily
   IF m_pFontFamily THEN GdipDeleteFontFamily(m_pFontFamily)
END DESTRUCTOR
' =====================================================================================

' ===========================================================================================
' Returns the last status code.
' ===========================================================================================
PRIVATE FUNCTION CGpFontFamily.GetLastStatus () AS GpStatus
   RETURN m_Status
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Sets the last status code.
' ===========================================================================================
PRIVATE FUNCTION CGpFontFamily.SetStatus (BYVAL status AS GpStatus) AS GpStatus
   m_Status = status
   RETURN m_Status
END FUNCTION
' ===========================================================================================

' =====================================================================================
' * Creates a new FontFamily object based on this Brush object.
' =====================================================================================
PRIVATE FUNCTION CGpFontFamily.Clone (BYVAL pFontFamily AS CGpFontFamily PTR) AS GpStatus
   IF pFontFamily = NULL THEN RETURN SetStatus(InvalidParameter)
   IF pFontFamily->m_pFontFamily THEN GdipDeleteFontFamily(pFontFamily->m_pFontFamily)
   RETURN SetStatus(GdipCloneFontFamily(m_pFontFamily, @pFontFamily->m_pFontFamily))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets a FontFamily object that specifies a generic sans serif typeface.
' =====================================================================================
PRIVATE FUNCTION CGpFontFamily.GenericSansSerif (BYVAL pFontFamily AS CGpFontFamily PTR) AS GpStatus
   IF pFontFamily = NULL THEN RETURN SetStatus(InvalidParameter)
   IF pFontFamily->m_pFontFamily THEN GdipDeleteFontFamily(pFontFamily->m_pFontFamily)
   m_Status = GdipGetGenericFontFamilySansSerif(@pFontFamily->m_pFontFamily)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets a FontFamily object that specifies a generic serif typeface.
' =====================================================================================
PRIVATE FUNCTION CGpFontFamily.GenericSerif (BYVAL pFontFamily AS CGpFontFamily PTR) AS GpStatus
   IF pFontFamily = NULL THEN RETURN SetStatus(InvalidParameter)
   IF pFontFamily->m_pFontFamily THEN GdipDeleteFontFamily(pFontFamily->m_pFontFamily)
   m_Status = GdipGetGenericFontFamilySerif(@pFontFamily->m_pFontFamily)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets a FontFamily object that specifies a generic monospace typeface.
' =====================================================================================
PRIVATE FUNCTION CGpFontFamily.GenericMonospace (BYVAL pFontFamily AS CGpFontFamily PTR) AS GpStatus
   IF pFontFamily = NULL THEN RETURN SetStatus(InvalidParameter)
   IF pFontFamily->m_pFontFamily THEN GdipDeleteFontFamily(pFontFamily->m_pFontFamily)
   m_Status = GdipGetGenericFontFamilyMonospace(@pFontFamily->m_pFontFamily)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the name of this font family.
' =====================================================================================
PRIVATE FUNCTION CGpFontFamily.GetFamilyName (BYVAL pwszName AS WSTRING PTR, BYVAL language AS LANGID = LANG_NEUTRAL) AS GpStatus
   IF pwszName = NULL THEN RETURN SetStatus(InvalidParameter)
   RETURN SetStatus(GdipGetFamilyName(m_pFontFamily, pwszName, language))
END FUNCTION
' =====================================================================================

' ========================================================================================
' The IsAvailable method determines whether this Font object was created successfully.
' ========================================================================================
PRIVATE FUNCTION CGpFontFamily.IsAvailable () AS BOOLEAN
   RETURN m_pFontFamily <> NULL
END FUNCTION
' ========================================================================================

' =====================================================================================
' Determines whether the specified style is available for this font family.
' =====================================================================================
PRIVATE FUNCTION CGpFontFamily.IsStyleAvailable (BYVAL nStyle AS INT_) AS BOOLEAN
   DIM bStyleAvailable AS BOOL
   m_Status = GdipIsStyleAvailable(m_pFontFamily, nStyle, @bStyleAvailable)
   RETURN bStyleAvailable
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the size (commonly called em size or em height), in design units, of this font family.
' =====================================================================================
PRIVATE FUNCTION CGpFontFamily.GetEmHeight (BYVAL nStyle AS INT_) AS UINT16
   DIM EmHeight AS UINT16
   m_Status = GdipGetEmHeight(m_pFontFamily, nStyle, @EmHeight)
   RETURN EmHeight
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the cell ascent, in design units, of this font family for the specified style
' or style combination.
' =====================================================================================
PRIVATE FUNCTION CGpFontFamily.GetCellAscent (BYVAL nStyle AS INT_) AS UINT16
   DIM CellAscent AS UINT16
   m_Status = GdipGetCellAscent(m_pFontFamily, nStyle, @CellAscent)
   RETURN CellAscent
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the cell descent, in design units, of this font family for the specified style
' or style combination.
' =====================================================================================
PRIVATE FUNCTION CGpFontFamily.GetCellDescent (BYVAL nStyle AS INT_) AS UINT16
   DIM CellDescent AS UINT16
   m_Status = GdipGetCellDescent(m_pFontFamily, nStyle, @CellDescent)
   RETURN CellDescent
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the line spacing, in design units, of this font family for the specified style
' or style combination. The line spacing is the vertical distance between the base
' lines of two consecutive lines of text.
' =====================================================================================
PRIVATE FUNCTION CGpFontFamily.GetLineSpacing (BYVAL nStyle AS INT_) AS UINT16
   DIM LineSpacing AS UINT16
   m_Status = GdipGetLineSpacing(m_pFontFamily, nStyle, @LineSpacing)
   RETURN LineSpacing
END FUNCTION
' =====================================================================================

' ########################################################################################
' CGpFontCollection class
' Base interface to allow to use both an CGpPrivateFontCollection pointer or an
' CGpInstalledFontCollection pointer as a parameter.
' ########################################################################################

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CGpFontCollection
END CONSTRUCTOR
' ========================================================================================

' =====================================================================================
' Cleanup
' =====================================================================================
PRIVATE DESTRUCTOR CGpFontCollection
END DESTRUCTOR
' =====================================================================================

' ===========================================================================================
' Returns the last status code.
' ===========================================================================================
PRIVATE FUNCTION CGpFontCollection.GetLastStatus () AS GpStatus
   RETURN m_Status
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Sets the last status code.
' ===========================================================================================
PRIVATE FUNCTION CGpFontCollection.SetStatus (BYVAL status AS GpStatus) AS GpStatus
   m_Status = status
   RETURN m_Status
END FUNCTION
' ===========================================================================================

' =====================================================================================
' Gets the number of font families contained in this font collection.
' =====================================================================================
PRIVATE FUNCTION CGpFontCollection.GetFamilyCount () AS INT_
   DIM numFound AS INT_
   m_Status = GdipGetFontCollectionFamilyCount(m_pFontCollection, @numFound)
   RETURN numFound
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the font families contained in this font collection.
' =====================================================================================
PRIVATE FUNCTION CGpFontCollection.GetFamilies (BYVAL numSought AS INT_, BYVAL gpfamilies AS CGpFontFamily PTR, BYVAL numFound AS INT_ PTR) AS GpStatus
   IF numSought <= 0 OR gpfamilies = NULL OR numFound = NULL THEN RETURN SetStatus(InvalidParameter)
   *numFound = 0
   DIM rgfamilies(numSought - 1) AS GpFontFamily PTR
   m_Status = GdipGetFontCollectionFamilyList(m_pFontCollection, numSought, @rgfamilies(0), numFound)
   IF m_Status = Ok THEN
      FOR i AS INT_ = 0 TO *numFound - 1
         m_Status = GdipCloneFontFamily(rgfamilies(i), @gpfamilies[i].m_pFontFamily)
      NEXT
   END IF
   RETURN m_Status
END FUNCTION
' =====================================================================================

' ########################################################################################
' CGpInstalledFontCollection class
' ########################################################################################

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CGpInstalledFontCollection
   m_Status = GdipNewInstalledFontCollection(@m_pFontCollection)
END CONSTRUCTOR
' ========================================================================================

' =====================================================================================
' Cleanup
' =====================================================================================
PRIVATE DESTRUCTOR CGpInstalledFontCollection
   ' Don't free the pointer of an installed font
END DESTRUCTOR
' =====================================================================================

' ########################################################################################
' CGpPrivateFontCollection class
' ########################################################################################

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CGpPrivateFontCollection
   m_Status = GdipNewPrivateFontCollection(@m_pFontCollection)
END CONSTRUCTOR
' ========================================================================================

' =====================================================================================
' Cleanup
' =====================================================================================
PRIVATE DESTRUCTOR CGpPrivateFontCollection
   IF m_pFontCollection THEN GdipDeletePrivateFontCollection(@m_pFontCollection)
END DESTRUCTOR
' =====================================================================================

' =====================================================================================
' Adds a font file to this private font collection.
' =====================================================================================
PRIVATE FUNCTION CGpPrivateFontCollection.AddFontFile (BYVAL pwszFileName AS WSTRING PTR) AS GpStatus
   RETURN SetStatus(GdipPrivateAddFontFile(m_pFontCollection, pwszFileName))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Adds a font that is contained in system memory to a Windows GDI+ font collection.
' =====================================================================================
PRIVATE FUNCTION CGpPrivateFontCollection.AddMemoryFont (BYVAL pMemory AS ANY PTR, BYVAL length AS INT_) AS GpStatus
   RETURN SetSTatus(GdipPrivateAddMemoryFont(m_pFontCollection, pMemory, length))
END FUNCTION
' =====================================================================================
