' ########################################################################################
' Microsoft Windows
' File: CGpStringFormat.inc
' Contents: GDI+ string format class
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

' ########################################################################################
' CGpStringFormat Class
' ########################################################################################

' =====================================================================================
' Creates a StringFormat object based on string format flags and a language.
' =====================================================================================
PRIVATE CONSTRUCTOR CGpStringFormat (BYVAL formatFlags AS INT_ = 0, BYVAL language AS LANGID = LANG_NEUTRAL)
   m_Status = GdipCreateStringFormat(formatFlags, language, @m_pStringFormat)
END CONSTRUCTOR
' =====================================================================================

' ===========================================================================================
' Cleanup
' ===========================================================================================
PRIVATE DESTRUCTOR CGpStringFormat
   IF m_pStringFormat THEN GdipDeleteStringFormat(m_pStringFormat)
END DESTRUCTOR
' ===========================================================================================

' ===========================================================================================
' Returns the last status code.
' ===========================================================================================
PRIVATE FUNCTION CGpStringFormat.GetLastStatus () AS GpStatus
   RETURN m_Status
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Sets the last status code.
' ===========================================================================================
PRIVATE FUNCTION CGpStringFormat.SetStatus (BYVAL status AS GpStatus) AS GpStatus
   m_Status = status
   RETURN m_Status
END FUNCTION
' ===========================================================================================

' =====================================================================================
' Creates a new StringFormat object based on this StringFormat object.
' =====================================================================================
PRIVATE FUNCTION CGpStringFormat.Clone (BYVAL pCloneStringFormat AS CGpStringFormat PTR) AS GpStatus
   IF pCloneStringFormat = NULL THEN RETURN SetStatus(InvalidParameter)
   IF pCloneStringFormat->m_pStringFormat THEN GdipDeleteStringFormat(pCloneStringFormat->m_pStringFormat)
   RETURN SetStatus(GdipCloneStringFormat(m_pStringFormat, @pCloneStringFormat->m_pStringFormat))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Creates a generic, default StringFormat object.
' =====================================================================================
PRIVATE FUNCTION CGpStringFormat.GenericDefault (BYVAL pStringFormat AS CGpStringFormat PTR) AS GpStatus
   IF pStringFormat = NULL THEN RETURN SetStatus(InvalidParameter)
   IF pStringFormat->m_pStringFormat THEN GdipDeleteStringFormat(pStringFormat->m_pStringFormat)
   RETURN SetStatus(GdipStringFormatGetGenericDefault(@pStringFormat->m_pStringFormat))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Creates a generic, typographic StringFormat object.
' =====================================================================================
PRIVATE FUNCTION CGpStringFormat.GenericTypographic (BYVAL pStringFormat AS CGpStringFormat PTR) AS GpStatus
   IF pStringFormat = NULL THEN RETURN SetStatus(InvalidParameter)
   IF pStringFormat->m_pStringFormat THEN GdipDeleteStringFormat(pStringFormat->m_pStringFormat)
   RETURN SetStatus(GdipStringFormatGetGenericTypographic(@pStringFormat->m_pStringFormat))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the format flags for this StringFormat object. The format flags determine most
' of the characteristics of a StringFormat object.
' =====================================================================================
PRIVATE FUNCTION CGpStringFormat.SetFormatFlags (BYVAL flags AS INT_) AS GpStatus
   RETURN SetStatus(GdipSetStringFormatFlags(m_pStringFormat, flags))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the format flags for this StringFormat object. The format flags determine most
' of the characteristics of a StringFormat object.
' =====================================================================================
PRIVATE FUNCTION CGpStringFormat.GetFormatFlags () AS INT_
   DIM flags AS INT_
   m_Status = GdipGetStringFormatFlags(m_pStringFormat, @flags)
   RETURN flags
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the character alignment of this StringFormat object in relation to the origin of
' the layout rectangle. A layout rectangle is used to position the displayed string.
' =====================================================================================
PRIVATE FUNCTION CGpStringFormat.SetAlignment (BYVAL nAlign AS StringAlignment) AS GpStatus
   RETURN SetStatus(GdipSetStringFormatAlign(m_pStringFormat, nAlign))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the format flags for this StringFormat object. The format flags determine most
' of the characteristics of a StringFormat object.
' =====================================================================================
PRIVATE FUNCTION CGpStringFormat.GetAlignment () AS StringAlignment
   DIM alignment AS StringAlignment
   m_Status = GdipGetStringFormatFlags(m_pStringFormat, @alignment)
   RETURN alignment
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the line alignment of this StringFormat object in relation to the origin of the
' layout rectangle. The line alignment setting specifies how to align the string vertically
' in the layout rectangle. The layout rectangle is used to position the displayed string.
' =====================================================================================
PRIVATE FUNCTION CGpStringFormat.SetLineAlignment (BYVAL nAlign AS StringAlignment) AS GpStatus
   RETURN SetStatus(GdipSetStringFormatLineAlign(m_pStringFormat, nAlign))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the line alignment of this StringFormat object in relation to the origin of the
' layout rectangle. The line alignment setting specifies how to align the string vertically
' in the layout rectangle. The layout rectangle is used to position the displayed string.
' =====================================================================================
PRIVATE FUNCTION CGpStringFormat.GetLineAlignment () AS StringAlignment
   DIM alignment AS StringAlignment
   m_Status = GdipGetStringFormatLineAlign(m_pStringFormat, @alignment)
   FUNCTION = alignment
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the type of processing that is performed on a string when the hot key prefix, an
' ampersand (&), is encountered. The ampersand is called the hot key prefix and can be
' used to designate certain keys as hot keys.
' =====================================================================================
PRIVATE FUNCTION CGpStringFormat.SetHotkeyPrefix (BYVAL nHotkeyPrefix AS HotkeyPrefix) AS GpStatus
   RETURN SetStatus(GdipSetStringFormatHotkeyPrefix(m_pStringFormat, nHotkeyPrefix))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the type of processing that is performed on a string when the hot key prefix, an
' ampersand (&), is encountered. The ampersand is called the hot key prefix and can be
' used to designate certain keys as hot keys.
' =====================================================================================
PRIVATE FUNCTION CGpStringFormat.GetHotkeyPrefix () AS HotkeyPrefix
   DIM nHotkeyPrefix AS HotkeyPrefix
   m_Status = GdipGetStringFormatHotkeyPrefix(m_pStringFormat, @nHotkeyPrefix)
   RETURN nHotkeyPrefix
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the offsets for tab stops in this StringFormat object.
' =====================================================================================
PRIVATE FUNCTION CGpStringFormat.SetTabStops (BYVAL firstTabOffset AS SINGLE, BYVAL nCount AS INT_, BYVAL tabStops AS SINGLE PTR) AS GpStatus
   RETURN SetStatus(GdipSetStringFormatTabStops(m_pStringFormat, firstTabOffset, nCount, tabStops))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the number of tab-stop offsets in this StringFormat object.
' =====================================================================================
PRIVATE FUNCTION CGpStringFormat.GetTabStopCount () AS INT_
   DIM nCount AS INT_
   m_Status = GdipGetStringFormatTabStopCount(m_pStringFormat, @nCount)
   RETURN nCount
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets offsets of the tab stops in this StringFormat object.
' =====================================================================================
PRIVATE FUNCTION CGpStringFormat.GetTabStops (BYVAL nCount AS INT_, BYVAL firstTabOffset AS SINGLE PTR, BYVAL tabStops AS SINGLE PTR) AS GpStatus
   RETURN SetStatus(GdipGetStringFormatTabStops(m_pStringFormat, nCount, firstTabOffset, tabStops))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the digit substitution FUNCTION and the language that corresponds to the digit substitutes.
' =====================================================================================
PRIVATE FUNCTION CGpStringFormat.SetDigitSubstitution (BYVAL language AS LANGID, BYVAL substitute AS StringDigitSubstitute) AS GpStatus
   RETURN SetStatus(GdipSetStringFormatDigitSubstitution(m_pStringFormat, language, substitute))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the digit substitution FUNCTION and the language that corresponds to the digit substitutes.
' =====================================================================================
PRIVATE FUNCTION CGpStringFormat.GetDigitSubstitutionLanguage () AS LANGID
   DIM language AS LANGID
   m_Status = GdipGetStringFormatDigitSubstitution(m_pStringFormat, @language, NULL)
   FUNCTION = language
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets an element of the StringDigitSubstitute enumeration that indicates the digit
' substitution method that is used by this StringFormat object.
' =====================================================================================
PRIVATE FUNCTION CGpStringFormat.GetDigitSubstitutionMethod () AS StringDigitSubstitute
   DIM substitute AS StringDigitSubstitute
   m_Status = GdipGetStringFormatDigitSubstitution(m_pStringFormat, BYVAL NULL, @substitute)
   RETURN substitute
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the trimming style for this StringFormat object. The trimming style determines
' how to trim a string so that it fits into the layout rectangle.
' =====================================================================================
PRIVATE FUNCTION CGpStringFormat.SetTrimming (BYVAL trimming AS StringTrimming) AS GpStatus
   RETURN SetStatus(GdipSetStringFormatTrimming(m_pStringFormat, trimming))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets an element of the StringDigitSubstitute enumeration that indicates the digit
' substitution FUNCTION that is used by this StringFormat object.
' =====================================================================================
PRIVATE FUNCTION CGpStringFormat.GetTrimming () AS StringTrimming
   DIM trimming AS StringTrimming
   m_Status = GdipGetStringFormatTrimming(m_pStringFormat, @trimming)
   RETURN trimming
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets a series of character ranges for this StringFormat object that, when in a string,
' can be measured by the MeasureCharacterRanges function.
' =====================================================================================
PRIVATE FUNCTION CGpStringFormat.SetMeasurableCharacterRanges (BYVAL rangeCount AS INT_, BYVAL ranges AS CharacterRange PTR) AS GpStatus
   RETURN SetStatus(GdipSetStringFormatMeasurableCharacterRanges(m_pStringFormat, rangeCount, ranges))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets a series of character ranges for this StringFormat object that, when in a string,
' can be measured by the MeasureCharacterRanges function.
' =====================================================================================
PRIVATE FUNCTION CGpStringFormat.GetMeasurableCharacterRangeCount () AS INT_
   DIM nCount AS INT_
   m_Status = GdipGetStringFormatMeasurableCharacterRangeCount(m_pStringFormat, @nCount)
   RETURN nCount
END FUNCTION
' =====================================================================================
