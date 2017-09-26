' ########################################################################################
' Microsoft Windows
' File: AfxStr.inc
' Contents: String wrapper functions.
' Compiler: FreeBasic 32 & 64-bit, Unicode.
' Copyright (c) 2016 Paul Squires and José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "win/ole2.bi"
#include once "win/shlwapi.bi"
#include once "Afx/AfxWin.inc"
#include once "Afx/CWStr.inc"
USING Afx

#if _WIN32_WINNT = &h0602
' ========================================================================================
' Returns an uppercase version of a string.
' - pwszStr [in]:
'      The string to convert. Cannot have a size of 0.
' - pwszLocaleName [in, optional]:
'      Pointer to a locale name or one of these pre-defined values:
'         LOCALE_NAME_INVARIANT
'         LOCALE_NAME_SYSTEM_DEFAULT
'         LOCALE_NAME_USER_DEFAULT
'      For a table of language culture names see:
'         https://msdn.microsoft.com/es-es/library/ee825488(v=cs.20).aspx
' - dwMapFlags [in, optional]:
'      Flag specifying the type of transformation to use during string mapping or the type
'      of sort key to generate. This parameter can have the following values
'      For a complete list see:
'      https://msdn.microsoft.com/en-us/library/windows/desktop/dd318702(v=vs.85).aspx
' Return value:
'   The uppercased string.
' ========================================================================================
PRIVATE FUNCTION AfxStrUCase (BYVAL pwszStr AS WSTRING PTR, _
   BYVAL pwszLocaleName AS WSTRING PTR = LOCALE_NAME_USER_DEFAULT, _
   BYVAL dwMapFlags AS DWORD = 0) AS CWSTR

   ' // Check the vailidy of the passed parameters
   IF pwszStr = NULL THEN RETURN ""
   IF LEN(*pwszStr) = 0 THEN RETURN ""
   DIM cwsOut AS CWSTR = *pwszStr
   ' // Make sure that LCMAP_LOWERCASE is not being used
   dwMapFlags = dwMapFlags AND (NOT LCMAP_LOWERCASE)
   ' // Make sure that LCMAP_LINGUISTIC_CASING OR LCMAP_UPPERCASE is being used
   dwMapFlags OR= LCMAP_LINGUISTIC_CASING OR LCMAP_UPPERCASE
   ' // Ususally, the length of the converted string will be the same that the one of
   ' // the input string, so let's try it
   DIM cbLen AS LONG = LCMapStringEx(pwszLocaleName, dwMapFlags, pwszStr, LEN(*pwszStr), _
       cwsOut.vptr, LEN(cwsOut), NULL, NULL, 0)
   ' // If the output length is bigger than the input one, request the needed length,
   ' // mke the buffer bigger and try it again
   IF cbLen = 0 AND GetLastError = ERROR_INSUFFICIENT_BUFFER THEN
      cbLen = LCMapStringEx(pwszLocaleName, dwMapFlags, pwszStr, LEN(*pwszStr), NULL, 0, NULL, NULL, 0)
      IF cbLen THEN
         cwsOut += SPACE(cbLen)
         cbLen = LCMapStringEx(pwszLocaleName, dwMapFlags, pwszStr, LEN(*pwszStr), _
                 cwsOut.vptr, LEN(cwsOut), NULL, NULL, 0)
      END IF
   END IF
   RETURN cwsOut

END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a lowercase version of a string.
' - pwszStr [in]:
'      The string to convert. Cannot have a size of 0.
' - pwszLocaleName [in, optional]:
'      Pointer to a locale name or one of these pre-defined values:
'         LOCALE_NAME_INVARIANT
'         LOCALE_NAME_SYSTEM_DEFAULT
'         LOCALE_NAME_USER_DEFAULT
'      For a table of language culture names see:
'         https://msdn.microsoft.com/es-es/library/ee825488(v=cs.20).aspx
' - dwMapFlags [in, optional]:
'      Flag specifying the type of transformation to use during string mapping or the type
'      of sort key to generate. This parameter can have the following values
'      For a complete list see:
'      https://msdn.microsoft.com/en-us/library/windows/desktop/dd318702(v=vs.85).aspx
' Return value:
'   The lowercased string.
' ========================================================================================
PRIVATE FUNCTION AfxStrLCase (BYVAL pwszStr AS WSTRING PTR, _
   BYVAL pwszLocaleName AS WSTRING PTR = LOCALE_NAME_USER_DEFAULT, _
   BYVAL dwMapFlags AS DWORD = 0) AS CWSTR

   ' // Check the vailidy of the passed parameters
   IF pwszStr = NULL THEN RETURN ""
   IF LEN(*pwszStr) = 0 THEN RETURN ""
   DIM cwsOut AS CWSTR = *pwszStr
   ' // Make sure that LCMAP_UPPERCASE is not being used
   dwMapFlags = dwMapFlags AND (NOT LCMAP_UPPERCASE)
   ' // Make sure that LCMAP_LINGUISTIC_CASING OR LCMAP_LOWERCASE is being used
   dwMapFlags OR= LCMAP_LINGUISTIC_CASING OR LCMAP_LOWERCASE
   ' // Ususally, the length of the converted string will be the same that the one of
   ' // the input string, so let's try it
   DIM cbLen AS LONG = LCMapStringEx(pwszLocaleName, dwMapFlags, pwszStr, LEN(*pwszStr), _
       cwsOut.vptr, LEN(cwsOut), NULL, NULL, 0)
   ' // If the output length is bigger than the input one, request the needed length,
   ' // mke the buffer bigger and try it again
   IF cbLen = 0 AND GetLastError = ERROR_INSUFFICIENT_BUFFER THEN
      cbLen = LCMapStringEx(pwszLocaleName, dwMapFlags, pwszStr, LEN(*pwszStr), NULL, 0, NULL, NULL, 0)
      IF cbLen THEN
         cwsOut += SPACE(cbLen)
         cbLen = LCMapStringEx(pwszLocaleName, dwMapFlags, pwszStr, LEN(*pwszStr), _
                 cwsOut.vptr, LEN(cwsOut), NULL, NULL, 0)
      END IF
   END IF
   RETURN cwsOut

END FUNCTION
' ========================================================================================
#endif

' ========================================================================================
' Translates ansi bytes to unicode bytes.
' Parameters:
' - ansiStr = An ansi or UTF8 string.
' - nCodePage = The code page used in the conversion, e.g. 1251 for Russian.
'   If you specify CP_UTF8, the returned string will be UTF8 encoded.
'   If you don't pass an unicode page, the function will use CP_ACP (0), which is the
'   system default Windows ANSI code page.
' Return value:
'   The converted string as a CWSTR.
' ========================================================================================
PRIVATE FUNCTION AfxUcode (BYREF ansiStr AS CONST STRING, BYVAL nCodePage AS LONG = 0) AS CWSTR
   IF nCodePage = CP_UTF8 THEN
      DIM dwLen AS DWORD = MultiByteToWideChar(CP_UTF8, 0, STRPTR(ansiStr), LEN(ansiStr), NULL, 0)
      IF dwLen THEN
         DIM cws AS CWSTR = SPACE(dwLen)
         dwLen = MultiByteToWideChar(CP_UTF8, 0, STRPTR(ansiStr), LEN(ansiStr), *cws, dwLen * 2)
         IF dwLen THEN RETURN cws
      END IF
   ELSE
      DIM cws AS CWSTR = SPACE(LEN(ansiStr))
      DIM dwLen AS DWORD = MultiByteToWideChar(nCodePage, MB_PRECOMPOSED, STRPTR(ansiStr), LEN(ansiStr), *cws, LEN(ansiStr) * 2)
      IF dwLen THEN RETURN cws
   END IF
   RETURN ""
END FUNCTION
' ========================================================================================

' ========================================================================================
' Translates unicode bytes to ansi bytes.
' Parameters:
' - pwszStr = The unicode string to convert
' - nCodePage = The code page used in the conversion, e.g. 1251 for Russian.
'   If you specify CP_UT8, it is assumed that ansiStr contains an UTF8 encoded string.
'   If you don't pass an unicode page, the function will use CP_ACP (0), which is the
'   system default Windows ANSI code page.
' Return value:
'   The converted string.
' ========================================================================================
PRIVATE FUNCTION AfxAcode (BYVAL pwszStr AS WSTRING PTR, BYVAL nCodePage AS LONG = 0) AS STRING
   IF pwszStr = NULL THEN EXIT FUNCTION
   IF nCodePage = CP_UTF8 THEN
      DIM dwLen AS DWORD = WideCharToMultiByte(CP_UTF8, 0, pwszStr, LEN(*pwszStr), NULL, 0, NULL, NULL)
      IF dwLen THEN
         DIM ansiStr AS STRING = SPACE(dwLen)
         dwLen = WideCharToMultiByte(CP_UTF8, 0, pwszStr, LEN(*pwszStr), STRPTR(ansiStr), LEN(ansiStr), NULL, NULL)
         IF dwLen THEN RETURN ansiStr
      END IF
   ELSE
      DIM ansiStr AS STRING = SPACE(LEN(*pwszStr))
      DIM dwLen AS DWORD = WideCharToMultiByte(nCodePage, 0, pwszStr, LEN(*pwszStr), STRPTR(ansiStr), LEN(ansiStr), NULL, NULL)
      IF dwLen THEN RETURN ansiStr
   ENDIF
   RETURN ""
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a copy of a string with substrings removed.
' If wszMatchStr is not present in wszMainStr, all of wszMainStr is returned intact.
' This function is case sensitive.
' Example: AfxStrRemove("Hello World. Welcome to the Freebasic World", "World")
' ========================================================================================
PRIVATE FUNCTION AfxStrRemove (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING) AS CWSTR
   DIM cws AS CWSTR = wszMainStr
   DIM nLen AS LONG = LEN(wszMatchStr)
   DO
      DIM nPos AS LONG = INSTR(cws, wszMatchStr)
      IF nPos = 0 THEN EXIT DO
      cws.DelChars nPos, nLen
   LOOP
   RETURN cws
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Case insensitive version of AfxStrRemove.
' Example: AfxStrRemoveI("Hello World. Welcome to the Freebasic World", "world")
' ========================================================================================
PRIVATE FUNCTION AfxStrRemoveI (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING) AS CWSTR
   DIM cwsMainStr AS CWSTR = wszMainStr
   DIM cwsMatchStr AS CWSTR = UCASE(wszMatchStr)
   DIM nLen AS LONG = LEN(wszMatchStr)
   DO
      DIM nPos AS LONG = INSTR(UCASE(cwsMainStr), cwsMatchStr)
      IF nPos = 0 THEN EXIT DO
      cwsMainStr.DelChars nPos, nLen
   LOOP
   RETURN cwsMainStr
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a copy of a string with characters removed.
' If wszMatchStr is not present in wszMainStr, all of wszMainStr is returned intact.
' wszMatchStr specifies a list of single characters to be searched for individually,
' a match on any one of which will cause that character to be removed from the result.
' This function is case sensitive.
' Example: AfxStrRemoveAny("abacadabra", "bac")   ' -> "dr"
' ========================================================================================
PRIVATE FUNCTION AfxStrRemoveAny (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING) AS CWSTR
   DIM cwsMainStr AS CWSTR = wszMainStr
   FOR i AS LONG = 1 TO LEN(wszMatchStr)
      DO
         DIM nPos AS LONG = INSTR(cwsMainStr, MID(wszMatchStr, i, 1))
         IF nPos = 0 THEN EXIT DO
         cwsMainStr.DelChars nPos, 1
      LOOP
   NEXT
   RETURN cwsMainStr
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Case insensitive version of AfxStrRemoveAny.
' Example: AfxStrRemoveAnyI("abacadabra", "BaC")   ' -> "dr"
' ========================================================================================
PRIVATE FUNCTION AfxStrRemoveAnyI (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING) AS CWSTR
   DIM cwsMainStr AS CWSTR = wszMainStr
   DIM cwsMatchStr AS CWSTR = UCASE(wszMatchStr)
   FOR i AS LONG = 1 TO LEN(wszMatchStr)
      DO
         DIM nPos AS LONG = INSTR(UCASE(cwsMainStr), MID(cwsMatchStr, i, 1))
         IF nPos = 0 THEN EXIT DO
         cwsMainStr.DelChars nPos, 1
      LOOP
   NEXT
   RETURN cwsMainStr
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Within a specified string, replace all occurrences of one string with another string.
' Replaces all occurrences of wszMatchStr in wszMainStr with wszReplaceWith
' The replacement can cause wszMainStr to grow or condense in size.
' When a match is found, the scan for the next match begins at the position immediately
' following the prior match.
' This function is case sensitive.
' Example: AfxStrReplace("Hello World", "World", "Earth")   ' -> "Hello Earth"
' ========================================================================================
PRIVATE FUNCTION AfxStrReplace OVERLOAD (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING, BYREF wszReplaceWith AS WSTRING) AS CWSTR
   DIM cwsMainStr AS CWSTR = wszMainStr
   DIM nLenReplaceWith AS LONG = LEN(wszReplaceWith)
   DIM nLenMatchStr AS LONG = LEN(wszMatchStr)
   DIM nPos AS LONG = 1
   DO
      nPos = INSTR(nPos, cwsMainStr, wszMatchStr)
      IF nPos = 0 THEN EXIT DO
      cwsMainStr = MID(cwsMainStr, 1, nPos - 1) + wszReplaceWith + MID(cwsMainStr, nPos + nLenMatchStr)
      nPos += nLenReplaceWith
   LOOP
   RETURN cwsMainStr
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Case insensitive version of AfxStrReplace.
' Example: AfxStrReplaceI("Hello world", "World", "Earth")   ' -> "Hello Earth"
' ========================================================================================
PRIVATE FUNCTION AfxStrReplaceI (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING, BYREF wszReplaceWith AS WSTRING) AS CWSTR
   DIM cwsMainStr AS CWSTR = wszMainStr
   DIM cwsMatchStr AS CWSTR = UCASE(wszMatchStr)
   DIM nLenReplaceWith AS LONG = LEN(wszReplaceWith)
   DIM nLenMatchStr AS LONG = LEN(wszMatchStr)
   DIM nPos AS LONG = 1
   DO
      nPos = INSTR(nPos, UCASE(cwsMainStr), cwsMatchStr)
      IF nPos = 0 THEN EXIT DO
      cwsMainStr = MID(cwsMainStr, 1, nPos - 1) + wszReplaceWith + MID(cwsMainStr, nPos + nLenMatchStr)
      nPos += nLenReplaceWith
   LOOP
   RETURN cwsMainStr
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Within a specified string, replace all occurrences of any of the individual characters
' specified in the wszMainStr string.
' wszReplaceWith must be a single character. This function does not replace words therefore
' wszMatchStr will be the same size - it will not shrink or grow.
' This function is case-sensitive.
' Example: AfxStrReplaceAny("abacadabra", "bac", "*")   ' -> *****d**r*
' ========================================================================================
PRIVATE FUNCTION AfxStrReplaceAny (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING, BYREF wszReplaceWith AS WSTRING) AS CWSTR
   DIM cwsMainStr AS CWSTR = wszMainStr
   IF LEN(wszMatchStr) = 0 THEN RETURN cwsMainStr
   IF LEN(wszReplaceWith) = 0 THEN RETURN cwsMainStr
   FOR x AS LONG = 1 TO LEN(wszMatchStr)
      FOR i AS LONG = 1 TO LEN(wszMainStr)
         IF MID(wszMatchStr, x, 1) = MID(wszMainStr, i, 1) THEN
            MID(**cwsMainStr, i, 1) = wszReplaceWith
         END IF
      NEXT
   NEXT
   RETURN cwsMainStr
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Case insensitive version of AfxStrReplaceAny.
' Example: AfxStrReplaceAnyI("abacadabra", "BaC", "*")   ' -> *****d**r*
' ========================================================================================
PRIVATE FUNCTION AfxStrReplaceAnyI (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING, BYREF wszReplaceWith AS WSTRING) AS CWSTR
   DIM cwsMainStr AS CWSTR = wszMainStr
   IF LEN(wszMatchStr) = 0 THEN RETURN cwsMainStr
   IF LEN(wszReplaceWith) = 0 THEN RETURN cwsMainStr
   FOR x AS LONG = 1 TO LEN(wszMatchStr)
      FOR i AS LONG = 1 TO LEN(wszMainStr)
         IF MID(UCASE(wszMatchStr), x, 1) = MID(UCASE(wszMainStr), i, 1) THEN
            MID(**cwsMainStr, i, 1) = wszReplaceWith
         END IF
      NEXT
   NEXT
   RETURN cwsMainStr
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Reverses the contents of a string expression.
' Usage example: DIM cws AS CWSTR = AfxStrReverse("garden")
' ========================================================================================
PRIVATE FUNCTION AfxStrReverse (BYREF wszMainStr AS WSTRING) AS CWSTR
   DIM cwsMainStr AS CWSTR = wszMainStr
   DIM wszChar AS WSTRING * 2
   DIM nLen AS LONG = LEN(wszMainStr)
   FOR i AS LONG = 1 TO nLen \ 2
      wszChar = MID(cwsMainStr, i, 1)
      MID(**cwsMainStr, i, 1) = MID(**cwsMainStr, nLen - i + 1, 1)
      MID(**cwsMainStr, nLen - i + 1, 1) = wszChar
   NEXT
   RETURN cwsMainStr
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Extracts characters from a string up to a character or group of characters.
' Complement function to AfxStrRemain.
' Returns a substring of wszMainStr starting with its first character (or the character
' specified by nStart) and up to (but not including) the first occurrence of wszMatchStr
' If wszMatchStr is not present in wszMainStr (or is null) then all of wszMainStr is
' returned from the nStart position.
' This function is case-sensitive.
' The following line returns "aba" (match on "cad")
' DIM cws AS CWSTR = AfxStrExtract(1, "abacadabra","cad")
' ========================================================================================
PRIVATE FUNCTION AfxStrExtract (BYVAL nStart AS LONG = 1, BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING) AS CWSTR
   DIM nLen AS LONG = LEN(wszMainStr)
   IF (nStart = 0) OR (nStart > nLen) THEN RETURN ""
   IF nStart < 0 THEN nStart = nLen + nStart + 1
   DIM nPos AS LONG = INSTR(nStart, wszMainStr, wszMatchStr)
   IF nPos THEN RETURN MID(wszMainStr, nStart, nPos - nStart)
   RETURN MID(wszMainStr, nStart)
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Case insensitive version of AfxStrExtract.
' DIM cws AS CWSTR = AfxStrExtractI(1, "abacadabra","CaD")
' ========================================================================================
PRIVATE FUNCTION AfxStrExtractI (BYVAL nStart AS LONG = 1, BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING) AS CWSTR
   DIM cws AS CWSTR = wszMainStr
   DIM nLen AS LONG = LEN(wszMainStr)
   IF (nStart = 0) OR (nStart > nLen) THEN RETURN ""
   IF nStart < 0 THEN nStart = nLen + nStart + 1
   DIM nPos AS LONG = INSTR(nStart, UCASE(wszMainStr), UCASE(wszMatchStr))
   IF nPos THEN RETURN MID(wszMainStr, nStart, nPos - nStart )
   RETURN cws = MID(wszMainStr, nStart)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Extract characters from a string up to a specific character.
' Returns a substring of wszMainStr starting with its first character (or the character
' specified by nStart) and up to (but not including) the first occurrence of wszMatchStr.
' wszMatchStr specifies a list of single characters to be searched for individually, a
' match on any one of which will cause the extract operation to be performed up to that character.
' If wszMatchStr is not present in wszMainStr (or is null) then all of wszMainStr is returned.
' This function is case-sensitive.
' The following line returns "aba" (match on "c")
' Example: AfxStrExtractAny(1, "abacadabra","cd")
' ========================================================================================
PRIVATE FUNCTION AfxStrExtractAny (BYVAL nStart AS LONG = 1, BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING) AS CWSTR
   DIM cwsMainStr AS CWSTR = wszMainStr
   DIM nLen AS LONG = LEN(wszMainStr)
   IF (nStart = 0) OR (nStart > nLen) THEN RETURN ""
   IF nStart < 0 THEN nStart = nLen + nStart + 1
   FOR i AS LONG = nStart TO nLen
      FOR x AS LONG = 1 TO LEN(wszMatchStr)
         IF MID(wszMainStr, i, 1) = MID(wszMatchStr, x, 1) THEN
            cwsMainStr = MID(wszMainStr, nStart, i - nStart)
            RETURN cwsMainStr
         END IF
      NEXT
   NEXT
   RETURN ""
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Case insensitive version of AfxStrExtractAny.
' Example: AfxStrExtractAnyI(1, "abacadabra","CD")
' ========================================================================================
PRIVATE FUNCTION AfxStrExtractAnyI (BYVAL nStart AS LONG = 1, BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING) AS CWSTR
   DIM cwsMainStr AS CWSTR = wszMainStr
   DIM nLen AS LONG = LEN(wszMainStr)
   IF (nStart = 0) OR (nStart > nLen) THEN RETURN ""
   IF nStart < 0 THEN nStart = nLen + nStart + 1
   FOR i AS LONG = nStart TO nLen
      FOR x AS LONG = 1 TO LEN(wszMatchStr)
         IF MID(UCASE(wszMainStr), i, 1) = MID(UCASE(wszMatchStr), x, 1) THEN
            cwsMainStr = MID(wszMainStr, nStart, i - nStart)
            RETURN cwsMainStr
         END IF
      NEXT
   NEXT
   RETURN ""
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Complement to the AfxStrExtract function.
' Returns the portion of a string following the first occurrence of a substring.
' wszMainStr is searched for the string specified in wszMatchStr If found, all characters
' after cbMatchStr are returned. If wszMatchStr is not present in wszMainStr (or is null) then
' a zero-length empty string is returned.
' nStart is an optional starting position to begin searching. If nStart is not specified,
' position 1 will be used. If nStart is zero, a nul string is returned. If nStart is negative,
' the starting position is counted from right to left: if -1, the search begins at the last
' character; if -2, the second to last, and so forth.
' This function is case-sensitive.
' Example: AfxStrRemain("Brevity is the soul of wit", "is ")   ' -> "the soul of wit"
' ========================================================================================
PRIVATE FUNCTION AfxStrRemain (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING, BYVAL nStart AS LONG = 1) AS CWSTR
   IF LEN(wszMainStr) = 0 OR LEN(wszMatchStr) = 0 THEN RETURN ""
   IF nStart = 0 OR nStart > LEN(wszMainStr) THEN RETURN ""
   IF nStart < 0 THEN nStart = LEN(wszMainStr) + nStart + 1
   DIM nPos AS LONG = INSTR(nStart, wszMainStr, wszMatchStr)
   IF nPos = 0 THEN RETURN ""
   DIM cwsMainStr AS CWSTR = wszMainStr
   cwsMainStr = MID(**cwsMainStr, nPos + LEN(wszMatchStr))
   RETURN cwsMainStr
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Case insensitive version of AfxStrRemain.
' Example: AfxStrRemainI("Brevity is the soul of wit", "Is ")   ' -> "the soul of wit"
' ========================================================================================
PRIVATE FUNCTION AfxStrRemainI (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING, BYVAL nStart AS LONG = 1) AS CWSTR
   IF LEN(wszMainStr) = 0 OR LEN(wszMatchStr) = 0 THEN RETURN ""
   IF nStart = 0 OR nStart > LEN(wszMainStr) THEN RETURN ""
   IF nStart < 0 THEN nStart = LEN(wszMainStr) + nStart + 1
   DIM nPos AS LONG = INSTR(nStart, UCASE(wszMainStr), UCASE(wszMatchStr))
   IF nPos = 0 THEN RETURN ""
   DIM cwsMainStr AS CWSTR = wszMainStr
   cwsMainStr = MID(**cwsMainStr, nPos + LEN(wszMatchStr))
   RETURN cwsMainStr
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Complement to the AfxStrExtract function. Returns the portion of a string following the
' first occurrence of a character or group of characters.
' wszMainStr is searched for the string specified in wszMatchStr If found, all characters
' after wszMatchStr are returned. If wszMatchStr is not present in wszMainStr (or is null) then
' a zero-length empty string is returned.
' wszMatchStr specifies a list of single characters to be searched for individually. A match
' on any one of which will cause the extract operation be performed after that character.
' nStart is an optional starting position to begin searching. If nStart is not specified,
' position 1 will be used. If nStart is zero, a nul string is returned. If nStart is negative,
' the starting position is counted from right to left: if -1, the search begins at the last
' character; if -2, the second to last, and so forth.
' This function is case-sensitive.
' Example: AfxStrRemainAny("I think, therefore I am", ",")   ' -> " therefore I am"
' ========================================================================================
PRIVATE FUNCTION AfxStrRemainAny (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING, BYVAL nStart AS LONG = 1) AS CWSTR
   IF LEN(wszMainStr) = 0 OR LEN(wszMatchStr) = 0 THEN RETURN ""
   IF nStart = 0 OR nStart > LEN(wszMainStr) THEN RETURN ""
   IF nStart < 0 THEN nStart = LEN(wszMainStr) + nStart + 1
   DIM cwsMainStr AS CWSTR
   FOR i AS LONG = nStart TO LEN(wszMainStr)
      FOR x AS LONG = 1 TO LEN(wszMatchStr)
         IF MID(wszMainStr, i, 1) = MID(wszMatchStr, x, 1) THEN
            cwsMainStr = MID(wszMainStr, i + 1)
            RETURN cwsMainStr
         END IF
      NEXT
   NEXT
   RETURN ""
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Case insensitive version of AfxStrRemainAny.
' Example: AfxStrRemainAnyI("I think, therefore I am", "E")   ' -> "refore I am"
' ========================================================================================
PRIVATE FUNCTION AfxStrRemainAnyI (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING, BYVAL nStart AS LONG = 1) AS CWSTR
   IF LEN(wszMainStr) = 0 OR LEN(wszMatchStr) = 0 THEN RETURN ""
   IF nStart = 0 OR nStart > LEN(wszMainStr) THEN RETURN ""
   IF nStart < 0 THEN nStart = LEN(wszMainStr) + nStart + 1
   DIM cwsMainStr AS CWSTR
   FOR i AS LONG = nStart TO LEN(wszMainStr)
      FOR x AS LONG = 1 TO LEN(wszMatchStr)
         IF MID(UCASE(wszMainStr), i, 1) = MID(UCASE(wszMatchStr), x, 1) THEN
            cwsMainStr = MID(wszMainStr, i + 1)
            RETURN cwsMainStr
         END IF
      NEXT
   NEXT
   RETURN ""
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Count the number of occurrences of strings within a string.
' wszMainStr is the string expression in which to count characters.
' cbMatchStr is the string expression to count all occurrences of.
' If cbMatchStr is not present in wszMainStr, zero is returned.
' When a match is found, the scan for the next match begins at the position immediately
' following the prior match.
' This function is case-sensitive.
' Example: DIM nCount AS LONG = AfxStrTally("abacadabra", "ab")   ' -> 2
' ========================================================================================
PRIVATE FUNCTION AfxStrTally (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING) AS LONG
   DIM nCount AS LONG, nPos AS LONG = 1
   DIM nLen AS LONG = LEN(wszMatchStr)
   DO
      nPos = INSTR(nPos, wszMainStr, wszMatchStr)
      IF nPos = 0 THEN EXIT DO
      nCount += 1
      nPos += nLen
   LOOP
   RETURN nCount
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Case insensitive version of AfxStrTally.
' Example: DIM nCount AS LONG = AfxStrTallyI("abacadabra", "Ab")   ' -> 2
' ========================================================================================
PRIVATE FUNCTION AfxStrTallyI (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING) AS LONG
   DIM nCount AS LONG, nPos AS LONG = 1
   DIM cwsMainStr AS CWSTR = UCASE(wszMainStr)
   DIM cwsMatchStr AS CWSTR = UCASE (wszMatchStr)
   DIM nLen AS LONG = LEN(wszMatchStr)
   DO
      nPos = INSTR(nPos, cwsMainStr, cwsMatchStr)
      IF nPos = 0 THEN EXIT DO
      nCount += 1
      nPos += nLen
   LOOP
   RETURN nCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Count the number of occurrences of specified characters strings within a string.
' wszMainStr is the string expression in which to count characters.
' wszMatchStr is a list of single characters to be searched for individually. A match on
' any one of which will cause the count to be incremented for each occurrence of that
' character. Note that repeated characters in wszMatchStr will not increase the count.
' This function is case-sensitive.
' Example: DIM nCount AS LONG = AfxStrTallyAny("abacadabra", "bac")   ' -> 8
' ========================================================================================
PRIVATE FUNCTION AfxStrTallyAny (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING) AS LONG
   IF LEN(wszMainStr) = 0 OR LEN(wszMatchStr) = 0 THEN EXIT FUNCTION
   ' // Remove possible duplicates in the matches string
   DIM nPos AS LONG
   DIM cwsMatchStr AS CWSTR = wszMatchStr
   FOR i AS LONG = 1 TO LEN(cwsMatchStr)
      nPos = INSTR(cwsMatchStr, MID(wszMatchStr, i, 1))
      IF nPos = 0 THEN cwsMatchStr += MID(wszMatchStr, i, 1)
   NEXT
   ' // Do the count
   DIM nCount AS LONG
   FOR i AS LONG = 1 TO LEN(cwsMatchStr)
      nPos = 1
      DO
         nPos = INSTR(nPos, wszMainStr, MID(cwsMatchStr, i, 1))
         IF nPos = 0 THEN EXIT DO
         IF nPos THEN
            nCount += 1
            nPos += 1
         END IF
      LOOP
   NEXT
   RETURN nCount
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Case insensitive version of AfxStrTallyAny.
' Example: DIM nCount AS LONG = AfxStrTallyAnyI("abacadabra", "bAc")
' ========================================================================================
PRIVATE FUNCTION AfxStrTallyAnyI (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING) AS LONG
   IF LEN(wszMainStr) = 0 OR LEN(wszMatchStr) = 0 THEN EXIT FUNCTION
   ' // Remove possible duplicates in the matches string
   DIM nPos AS LONG
   DIM cwsMainStr AS CWSTR = UCASE(wszMainStr)
   DIM cwsMatchStr AS CWSTR = UCASE(wszMatchStr)
   FOR i AS LONG = 1 TO LEN(cwsMatchStr)
      nPos = INSTR(cwsMatchStr, MID(wszMatchStr, i, 1))
      IF nPos = 0 THEN cwsMatchStr += MID(wszMatchStr, i, 1)
   NEXT
   ' // Do the count
   DIM nCount AS LONG
   FOR i AS LONG = 1 TO LEN(cwsMatchStr)
      nPos = 1
      DO
         nPos = INSTR(nPos, cwsMainStr, MID(cwsMatchStr, i, 1))
         IF nPos = 0 THEN EXIT DO
         IF nPos THEN
            nCount += 1
            nPos += 1
         END IF
      LOOP
   NEXT
   RETURN nCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Determine whether each character of a string is present in another string.
' Returns zero if each character in wszMainStr is present in wszMatchStr
' If not, it returns the position of the first non-matching character in wszMainStr.
' This function is very useful for determining if a string contains only numeric digits, for example.
' This function is case-sensitive.
' If nStart evaluates to a position outside of the string, or if nStart is zero, then the
' function returns zero.
' Example: DIM nCount AS LONG = AfxStrVerify(5, "123.65,22.5", "0123456789")   ' -> 7
' Rreturns 7 since 5 starts it past the first non-digit ("." at position 4)
' ========================================================================================
PRIVATE FUNCTION AfxStrVerify (BYVAL nStart AS LONG, BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING) AS LONG
   IF nStart <= 0 OR nStart > LEN(wszMainStr) THEN RETURN 0
   ' // Get each character in wszMainStr and look for it in wszMatchStr
   DIM AS LONG nPos, idx
   FOR i AS LONG = nStart TO LEN(wszMainStr)
      nPos = INSTR(wszMatchStr, MID(wszMainStr, i, 1))
      IF nPos = 0 THEN
         idx = i
         EXIT FOR
      END IF
   NEXT
   RETURN  idx
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Case sinsentive version of AfxStrVerify.
' Example: AfxStrVerifyI(5, "123.65abcx22.5", "0123456789ABC")   ' -> 10
' ========================================================================================
PRIVATE FUNCTION AfxStrVerifyI (BYVAL nStart AS LONG, BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING) AS LONG
   IF nStart <= 0 OR nStart > LEN(wszMainStr) THEN RETURN 0
   ' // Get each character in wszMainStr and look for it in wszMatchStr
   DIM cwsMainStr AS CWSTR = UCASE(wszMainStr)
   DIM cwsMatchStr AS CWSTR = UCASE(wszMatchStr)
   DIM AS LONG nPos, idx
   FOR i AS LONG = nStart TO LEN(cwsMainStr)
      nPos = INSTR(cwsMatchStr, MID(cwsMainStr, i, 1))
      IF nPos = 0 THEN
         idx = i
         EXIT FOR
      END IF
   NEXT
   RETURN  idx
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a string containing a left-justified (padded) string.
' If the optional parameter wszPadCharacter not specified, the function pads the string with
' space characters to the left. Otherwise, the function pads the string with the first
' character of wszPadCharacter
' Example: DIM cws AS CWSTR = AfxStrLSet("FreeBasic", 20, "*")
' ========================================================================================
PRIVATE FUNCTION AfxStrLSet (BYREF wszMainStr AS WSTRING, BYVAL nStringLength AS LONG, BYREF wszPadCharacter AS WSTRING = " ") AS CWSTR
   DIM cws AS CWSTR = STRING(nStringLength, ASC(wszPadCharacter, 1))
   MID(**cws, 1, LEN(wszMainStr)) = wszMainStr
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a string containing a right-justified (padded) string.
' If the optional parameter wszPadCharacter not specified, the function pads the string with
' space characters to the left. Otherwise, the function pads the string with the first
' character of wszPadCharacter.
' Example: DIM cws AS CWSTR = AfxStrRSet("FreeBasic", 20, "*")
' ========================================================================================
PRIVATE FUNCTION AfxStrRSet (BYREF wszMainStr AS WSTRING, BYVAL nStringLength AS LONG, BYREF wszPadCharacter AS WSTRING = " ") AS CWSTR
   DIM cws AS CWSTR = STRING(nStringLength, ASC(wszPadCharacter, 1))
   MID(**cws, nStringLength - LEN(wszMainStr) + 1, LEN(wszMainStr)) = wszMainStr
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a string containing a centered (padded) string.
' If the optional parameter wszPadCharacter not specified, the function pads the string with
' space characters to the left. Otherwise, the function pads the string with the first
' character of wszPadCharacter.
' Example: DIM cws AS CWSTR = AfxStrCSet("FreeBasic", 20, "*")
' ========================================================================================
PRIVATE FUNCTION AfxStrCSet (BYREF wszMainStr AS WSTRING, BYVAL nStringLength AS LONG, BYREF wszPadCharacter AS WSTRING = " ") AS CWSTR
   DIM cws AS CWSTR = STRING(nStringLength, ASC(wszPadCharacter, 1))
   MID(**cws, (nStringLength - LEN(wszMainStr)) \ 2 + 1, LEN(wszMainStr)) = wszMainStr
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
'  Parses a path/file name to extract component parts.
'  This function evaluates a text path/file text name, and returns a requested part of the
'  name. The functionality is strictly one of string parsing alone.
'  wszOption is one of the following words which is used to specify the requested part:
'  PATH
'        Returns the path portion of the path/file Name. That is the text up to and
'        including the last backslash (\) or colon (:).
'  NAME
'        Returns the name portion of the path/file Name. That is the text to the right
'        of the last backslash (\) or colon (:), ending just before the last period (.).
'  EXTN
'        Returns the extension portion of the path/file name. That is the last
'        period (.) in the string plus the text to the right of it.
'  NAMEX
'        Returns the name and the EXTN parts combined.
' ========================================================================================
PRIVATE FUNCTION AfxStrPathName (BYREF wszOption AS WSTRING, BYREF wszFileSpec AS WSTRING) AS CWSTR
   DIM cws AS CWSTR = ""
   IF LEN(wszFileSpec) = 0 THEN RETURN cws
   SELECT CASE UCASE(wszOption)
      CASE "PATH"
         ' // Returns the path portion of file spec
         DIM nPos AS LONG = InstrRev(wszFileSpec, ANY ":/\")
         IF nPos THEN cws = MID(wszFileSpec, 1, nPos)
      CASE "NAME"
         ' // Retrieve the full filename
         cws = wszFileSpec
         DIM nPos AS LONG = InstrRev(wszFileSpec, ANY ":/\")
         IF nPos THEN cws = MID(wszFileSpec, nPos + 1)
         ' // Retrieve the filename
         nPos = InstrRev(cws, ".")
         IF nPos THEN cws = MID(cws, 1, nPos - 1)
      CASE "NAMEX"
         ' // Retrieve the name and extension combined
         DIM nPos AS LONG = InStrRev(wszFileSpec, ANY ":/\")
         IF nPos THEN cws = MID(wszFileSpec, nPos + 1) ELSE cws = wszFileSpec
      CASE "EXTN"
         ' // Retrieve the name and extension combined
         DIM nPos AS LONG = InstrRev(wszFileSpec, ANY ":/\")
         IF nPos THEN cws = MID(wszFileSpec, nPos + 1) ELSE cws = wszFileSpec
         ' // Retrieve the extension
         nPos = InStrRev(cws, ".")
         IF nPos THEN cws = MID(cws, nPos) ELSE cws = ""
   END SELECT
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a string consisting of multiple copies of the specified string.
' This function is very similar to STRING (which makes multiple copies of a single character).
' Example: DIM cws AS CWSTR = AfxStrRepeat(5, "Paul")
' ========================================================================================
PRIVATE FUNCTION AfxStrRepeat (BYVAL nCount AS LONG, BYREF wszMainStr AS WSTRING) AS CWSTR
   DIM cws AS CWSTR = ""
   IF nCount <= 0 THEN RETURN cws
   ' // Create the final full buffer and insert the strings into it
   ' // in order to avoid nCount concatenations.
   DIM nLen AS LONG = LEN(wszMainStr)
   cws = SPACE(nCount * nLen)
   FOR i AS LONG = 0 TO nCount - 1
      MID(**cws, (i * nLen) + 1, nLen) = wszMainStr
   NEXT
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a string with nCount characters removed from the left side of the string.
' If nCount is less than one then the entire string is returned.
' Example: DIM cws AS CWSTR = AfxStrClipLeft("1234567890", 3)
' ========================================================================================
PRIVATE FUNCTION AfxStrClipLeft (BYREF wszMainStr AS WSTRING, BYVAL nCount AS LONG) AS CWSTR
   DIM cws AS CWSTR = wszMainStr
   IF nCount <= 0 THEN RETURN cws
   DIM nLen AS LONG = LEN(wszMainStr)
   nCount = IIF(nLen < nCount, nLen, nCount)
   cws = MID(wszMainStr, nCount + 1)
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a string with nCount characters removed from the right side of the string.
' If nCount is less than one then the entire string is returned.
' DIM cws AS CWSTR = AfxStrClipRight("1234567890", 3)
' ========================================================================================
PRIVATE FUNCTION AfxStrClipRight (BYREF wszMainStr AS WSTRING, BYVAL nCount AS LONG) AS CWSTR
   DIM cws AS CWSTR = wszMainStr
   IF nCount <= 0 THEN RETURN cws
   DIM nLen AS LONG = LEN(wszMainStr)
   nCount = nLen - nCount
   nCount = IIF(nLen < nCount, nLen, nCount)
   cws = LEFT(wszMainStr, nCount)
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a string with nCount characters removed starting at position nStart. The first
' character is considered position 1, the second position 2, etc...
' If nCount or nStart is less than one then the entire string is returned.
' Usage example:
' DIM cws AS CWSTR = AfxStrClipMid("1234567890", 3, 4)
' ========================================================================================
PRIVATE FUNCTION AfxStrClipMid (BYREF wszMainStr AS WSTRING, BYVAL nStart AS LONG, BYVAL nCount AS LONG) AS CWSTR
   DIM cws AS CWSTR = wszMainStr
   IF (nCount <= 0) OR (nStart <= 0) THEN RETURN cws
   DIM nLen AS LONG = LEN(wszMainStr)
   cws = LEFT(wszMainStr, nStart - 1) + MID(wszMainStr, nStart + nCount)
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
'  Adds paired characters to the beginning and end of a string.
'  It is particularly useful for enclosing text with parenthesess, quotes, brackets, etc.
'  For example: AfxStrWrap("Paul", "<", ">") results in <Paul>
'  If only one wrap character/string is specified then that character or string is used
'  for both sides.
'  For example: AfxStrWrap("Paul", "'") results in 'Paul'
'  If no wrap character/string is specified then double quotes are used.
'  For example: AfxStrWrap("Paul") results in "Paul"
' ========================================================================================
PRIVATE FUNCTION AfxStrWrap OVERLOAD (BYREF wszMainStr AS WSTRING, BYREF wszLeftChar AS WSTRING, BYREF wszRightChar AS WSTRING) AS CWSTR
   DIM cws AS CWSTR = wszLeftChar + wszMainStr & wszRightChar
   RETURN cws
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxStrWrap OVERLOAD (BYREF wszMainStr AS WSTRING, BYREF wszChar AS WSTRING = CHR(34)) AS CWSTR
   DIM cws AS CWSTR = wszChar + wszMainStr + wszChar
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes paired characters to the beginning and end of a string.
' It is particularly useful for removing text with parenthesess, quotes, brackets, etc.
' For example: AfxStrUnWrap("<Paul>", "<", ">") results in Paul
' If only one unwrap character/string is specified then that character or string is used for both sides.
' For example: AfxStrUnWrap("'Paul'", "'") results in Paul
' If no wrap character/string is specified then double quotes are used.
' For example: AfxStrUnWrap("""Paul""") results in Paul
' ========================================================================================
PRIVATE FUNCTION AfxStrUnWrap OVERLOAD (BYREF wszMainStr AS WSTRING, BYREF wszLeftChar AS WSTRING, BYREF wszRightChar AS WSTRING) AS CWSTR
   DIM cws AS CWSTR = LTRIM(wszMainStr, wszLeftChar)
   cws = RTRIM(cws, wszRightChar)
   RETURN cws
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxStrUnWrap OVERLOAD (BYREF wszMainStr AS WSTRING, BYREF wszChar AS WSTRING = CHR(34)) AS CWSTR
   DIM cws AS CWSTR = LTRIM(wszMainStr, wszChar)
   cws = RTRIM(cws, wszChar)
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Deletes a specified number of characters from a string expression.
' Returns a string based on wszMainStr but with nCount characters deleted
' starting at position nStart. The first character in the string is position 1, etc.
' Usage example:
' DIM cws AS CWSTR = AfxStrDelete("1234567890", 4, 3)
' ========================================================================================
PRIVATE FUNCTION AfxStrDelete (BYREF wszMainStr AS WSTRING, BYVAL nStart AS LONG, BYVAL nCount AS LONG) AS CWSTR
   DIM cws AS CWSTR = wszMainStr
   DIM nLen AS LONG = LEN(wszMainStr)
   IF nLen = 0 OR nStart < 0 OR nCount <= 0 OR nStart > nLen THEN RETURN cws
'   cws = LEFT(wszMainStr, nStart) + MID(wszMainStr, nStart + 1 + nCount, nCount)   ' // wrong calculation
   cws.DelChars nStart, nCount
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Inserts a string at a specified position within another string expression.
' Returns a string consisting of wszMainStr with the string wszInsertString inserted
' at nPosition. If nPosition is greater than the length of wszMainStr or <= zero then
' wszInsertString is appended to wszMainStr. The first character in the string is position 1, etc.
' DIM cws AS CWSTR = AfxStrInsert("1234567890", "--", 6)
' ========================================================================================
PRIVATE FUNCTION AfxStrInsert (BYREF wszMainStr AS WSTRING, BYREF wszInsertString AS WSTRING, BYVAL nPosition AS LONG) AS CWSTR
   DIM cws AS CWSTR = wszMainStr
'   IF nPosition > LEN(wszMainStr) OR nPosition <= 0 THEN
'      cws += wszInsertString
'   ELSE
'      cws = MID(wszMainStr, 1, nPosition) + wszInsertString + MID(wszMainStr, nPosition + 1)
'   END IF
   IF nPosition <= 0 THEN RETURN CWS
   IF nPosition > LEN(wszMainStr) THEN
      cws += wszInsertString
   ELSEIF nPosition = 1 THEN
      cws = wszInsertString + MID(wszMainStr, 1)
   ELSE
      cws = MID(wszMainStr, 1, nPosition - 1) + wszInsertString + MID(wszMainStr, nPosition)
   END IF
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a string containing only the characters contained in a specified match string.
' All other characters are removed. If wszMatchStr is an empty string the function returns
' an empty string. This function is case-sensitive.
' Example: DIM cws AS CWSTR = AfxStrRetain("abacadabra","b")   ' -> "bb"
' ========================================================================================
PRIVATE FUNCTION AfxStrRetain (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING) AS CWSTR
   DIM cws AS CWSTR = ""
   IF LEN(wszMainStr) = 0 OR LEN(wszMatchStr) = 0 THEN RETURN cws
   DIM nLen AS LONG = LEN(wszMatchStr)
   DIM nPos AS LONG = 1
   DO
      nPos = INSTR(nPos, wszMainStr, wszMatchStr)
      IF nPos = 0 THEN EXIT DO
      cws += MID(wszMainStr, nPos, nLen)
      nPos += nLen
   LOOP
   RETURN cws
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Case insensitive version of AfxStrRetain.
' Example: DIM cws AS CWSTR = AfxStrRetainI("abacadabra","B")   ' -> "bb"
' ========================================================================================
PRIVATE FUNCTION AfxStrRetainI (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING) AS CWSTR
   DIM cws AS CWSTR = ""
   IF LEN(wszMainStr) = 0 OR LEN(wszMatchStr) = 0 THEN RETURN cws
   DIM cwsMainStr AS CWSTR = UCASE(wszMainStr)
   DIM cwsMatchStr AS CWSTR = UCASE(wszMatchStr)
   DIM nLen AS LONG = LEN(wszMatchStr)
   DIM nPos AS LONG = 1
   DO
      nPos = INSTR(nPos, cwsMainStr, cwsMatchStr)
      IF nPos = 0 THEN EXIT DO
      cws += MID(wszMainStr, nPos, nLen)
      nPos += nLen
   LOOP
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a string containing only the characters contained in a specified match string.
' All other characters are removed.
' If wszMatchStr is an empty string the function returns an empty string.
' wszMatchStr specifies a list of single characters to be searched for individually.
' A match on any one of which will cause that character to be removed from the result.
' This function is case-sensitive.
' Example: AfxStrRetainAny("<p>1234567890<ak;lk;l>1234567890</p>", "<;/p>")
' ========================================================================================
PRIVATE FUNCTION AfxStrRetainAny (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING) AS CWSTR
   DIM cws AS CWSTR = ""
   DIM nLen AS LONG = LEN(wszMainStr)
   IF nLen = 0 OR LEN(wszMatchStr) = 0 THEN RETURN cws
   DIM nPos AS LONG
   FOR i AS LONG = 1 TO nLen
      nPos = INSTR(wszMatchStr, MID(wszMainStr, i, 1))
      IF nPos THEN cws += MID(wszMainStr, i, 1)
   NEXT
   RETURN cws
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Case insensitive version of AfxStrRetainAny.
' Example: AfxStrRetainAnyI("<p>1234567890<ak;lk;l>1234567890</p>", "<;/P>")
' ========================================================================================
PRIVATE FUNCTION AfxStrRetainAnyI (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING) AS CWSTR
   DIM cws AS CWSTR = ""
   DIM nLen AS LONG = LEN(wszMainStr)
   IF nLen = 0 OR LEN(wszMatchStr) = 0 THEN RETURN cws
   DIM cwsMainStr AS CWSTR = UCASE(wszMainStr)
   DIM cwsMatchStr AS CWSTR = UCASE(wszMatchStr)
   DIM nPos AS LONG
   FOR i AS LONG = 1 TO nLen
      nPos = INSTR(cwsMatchStr, MID(cwsMainStr, i, 1))
      IF nPos THEN cws += MID(wszMainStr, i, 1)
   NEXT
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retuns TRUE if c is a number (0-9), a numeric sign (+-) or a decimal point (.).
' Note: Works both with single characters and strings.
' ========================================================================================
PRIVATE FUNCTION AfxIsNumeric (BYREF c AS WSTRING) AS BOOLEAN
   RETURN (AfxStrRetainAny(c, "+-.0123456789") = c) AND c <> ""
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Shrinks a string to use a consistent single character delimiter.
' The purpose of this function is to create a string with consecutive data items (words)
' separated by a consistent single character. This makes it very straightforward to parse
' the results as needed.
' If wszMask is not defined then all leading spaces and trailing spaces are removed entirely.
' All occurrences of two or more spaces are changed to a single space. Therefore, the new
' string returned consists of zero or more words, each separated by a single space character.
' If wszMask is specified, it defines one or more delimiter characters to shrink. All leading
' and trailing mask characters are removed entirely. All occurrences of one or more mask
' characters are replaced with the first character of wszMask The new string returned consists
' of zero or more words, each separated by the character found in the first position of wszMask.
' WhiteSpace is generally defined as the four common non-printing characters:
' Space, Tab, Carriage-Return, and Line-Feed. wszbMask = Chr(32,9,13,10)
' Example: DIM cws AS CWSTR = AfxStrShrink(",,, one , two     three, four,", " ,")
' ========================================================================================
PRIVATE FUNCTION AfxStrShrink (BYREF wszMainStr AS WSTRING, BYREF wszMask AS WSTRING = " ") AS CWSTR
   DIM cws AS CWSTR = ""
   IF LEN(wszMainStr) = 0 OR LEN(wszMask) = 0 THEN RETURN cws
   ' // Eliminate all leading and trailing cbMask characters
   cws = TRIM(wszMainStr, ANY wszMask)
   ' // Eliminate all duplicate wszMask characters within the string
   DIM wszReplace AS WSTRING * 2 = MID(wszMask, 1, 1)
   DIM wszDuplicate AS WSTRING * 3
   DIM nMaskLen AS LONG = LEN(wszMask)
   DIM nPos AS LONG
   FOR i AS LONG = 1 TO nMaskLen
      wszDuplicate = MID(wszMask, i, 1) + MID(wszMask, i, 1)   ' usually double spaces
      nPos = 1
      DO
         nPos = INSTR(cws, wszDuplicate)
         IF nPos = 0 THEN EXIT DO
         cws = MID(cws, 1, nPos - 1) + wszReplace + MID(cws, nPos + LEN(wszDuplicate))
      LOOP
   NEXT
   ' // Replace all single characters in the mask with the first character of the mask.
   nPos = 1
   DO
      nPos = INSTR(nPos, cws, ANY wszMask)
      IF nPos = 0 THEN EXIT DO
      ' Only do the replace if the character at the position found is
      ' different than the character we need to replace it with. This saves
      ' us from having to do an unneeded string concatenation.
      IF MID(cws, nPos, 1) <> wszReplace  THEN
         cws = MID(cws, 1, nPos - 1) + wszReplace + MID(cws, nPos + 1)
      END IF
      nPos += 1
   LOOP
   ' Finally, do a pass to ensure that there are no duplicates of the
   ' first mask character because of the replacements in the step above.
   wszDuplicate = MID(wszMask, 1, 1) + MID(wszMask, 1, 1)
   nPos = 1
   DO
      nPos = INSTR(cws, wszDuplicate)
      IF nPos = 0 THEN EXIT DO
      cws = MID(cws, 1, nPos - 1) + wszReplace + MID(cws, nPos + LEN(wszDuplicate))
   LOOP
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the count of delimited fields from a string expression.
' If wszMainStr is empty (a null string) or contains no delimiter character(s), the string
' is considered to contain exactly one sub-field. In this case, AfxStrParseCount returns the value 1.
' Delimiter contains a string (one or more characters) that must be fully matched.
' Delimiters are case-sensitive.
' Example: DIM nCount AS LONG = AfxStrParseCount("one,two,three", ",")
' ========================================================================================
PRIVATE FUNCTION AfxStrParseCount (BYREF wszMainStr AS WSTRING, BYREF wszDelimiter AS WSTRING = ",") AS LONG
   DIM nCount AS LONG = 1
   DIM nPos AS LONG = 1
   DO
      nPos = INSTR(nPos, wszMainStr, wszDelimiter)
      IF nPos = 0 THEN EXIT DO
      nCount += 1
      nPos += LEN(wszDelimiter)
   LOOP
   RETURN nCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Return the count of delimited fields from a string expression.
' If wszMainStr is empty (a null string) or contains no delimiter character(s), the string
' is considered to contain exactly one sub-field. In this case, AfxStrParseCountAny returns the value 1.
' Delimiter contains a set of characters (one or more), any of which may act as a delimiter character.
' Delimiters are case-sensitive.
' Example: DIM nCount AS LONG = AfxStrParseCountAny("1;2,3", ",;")
' ========================================================================================
PRIVATE FUNCTION AfxStrParseCountAny (BYREF wszMainStr AS WSTRING, BYREF wszDelimiter AS WSTRING = ",") AS LONG
   DIM nCount AS LONG = 1
   DIM nPos AS LONG = 1
   DO
      nPos = INSTR(nPos, wszMainStr, ANY wszDelimiter)
      IF nPos = 0 THEN EXIT DO
      nCount += 1
      nPos += LEN(wszDelimiter)
   LOOP
   RETURN nCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the nPosition-th substring in a string wszMainStr with separations wszDelimiter
' (one or more charcters), beginning with nPosition = 1.
' ========================================================================================
PRIVATE FUNCTION AfxStrParse OVERLOAD (BYREF wszMainStr AS WSTRING, BYREF wszDelimiter AS WSTRING, BYVAL nPosition AS LONG, BYVAL bIsAny AS BOOLEAN, BYVAL nLenDelimiter AS LONG) AS CWSTR
   DIM nCount AS LONG, nStart AS LONG
   nPosition = ABS(nPosition)
   DIM nPos AS LONG = 1
   DIM fReverse AS BOOLEAN = IIF&(nPosition < 0, TRUE, FALSE)
   DIM cws AS CWSTR = ""
   IF fReverse THEN
      ' Reverse search
      ' Get the start of the token (j) by searching in reverse
      IF bIsAny THEN
         nPos = InstrRev(wszMainStr, ANY wszDelimiter)
      ELSE
         nPos = InstrRev(wszMainStr, wszDelimiter)
      END IF
      DO WHILE nPos > 0        ' if not found loop will be skipped
         nStart = nPos + nLenDelimiter
         nCount += 1
         nPos = nPos - nLenDelimiter
         IF nCount = nPosition THEN EXIT DO
         IF bIsAny THEN
            nPos = InStrRev(wszMainStr, ANY wszDelimiter, nPos)
         ELSE
             nPos = InStrRev(wszMainStr, wszDelimiter, nPos)
         END IF
      LOOP
      IF nPos = 0 THEN nStart = 1
      ' Now continue forward to get the end of the token
      IF bIsAny THEN
         nPos = INSTR(nStart, wszMainStr, ANY wszDelimiter)
      ELSE
         nPos = INSTR(nStart, wszMainStr, wszDelimiter)
      END IF
      IF nPos > 0 OR nCount = nPosition THEN
         IF nPos = 0 THEN
            cws = MID(wszMainStr, nStart)
         ELSE
            cws = MID(wszMainStr, nStart, nPos - nStart)
         END IF
      END IF
   ELSE
      ' Forward search
      DO
         nStart = nPos
         IF bIsAny THEN
            nPos = INSTR(nPos, wszMainStr, ANY wszDelimiter)
         ELSE
            nPos = INSTR(nPos, wszMainStr, wszDelimiter)
         END IF
         IF nPos THEN
            nCount += 1
            nPos += nLenDelimiter
         END IF
      LOOP UNTIL nPos = 0 OR nCount = nPosition
      IF nPos > 0 OR nCount = nPosition - 1 THEN
         IF nPos = 0 THEN
            cws = MID(wszMainStr, nStart)
         ELSE
            cws = MID(wszMainStr, nStart, nPos - nLenDelimiter - nStart)
         END IF
      END IF
   END IF
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a delimited field from a string expression.
' wszDelimiter contains a string of one or more characters that must be fully matched to be successful.
' If nPosition evaluates to zero or is outside of the actual field count, an empty string is returned.
' If nPosition is negative then fields are searched from the right to left of the wszMainStr
' Delimiters are case-sensitive.
' Example: DIM cws AS CWSTR = AfxStrParse("one,two,three", 2)
' Example: DIM cws AS CWSTR = AfxStrParse("one;two,three", 1, ";")
' ========================================================================================
PRIVATE FUNCTION AfxStrParse OVERLOAD (BYREF wszMainStr AS WSTRING, BYVAL nPosition AS LONG = 1, BYREF wszDelimiter AS WSTRING = ",") AS CWSTR
   ' The parse must match the entire deliminter string
   RETURN AfxStrParse(wszMainStr, wszDelimiter, nPosition, FALSE, Len(wszDelimiter))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Return a delimited field from a string expression.
' Delimiter contains a set of characters (one or more), any of which may act as a delimiter character.
' If nPosition evaluates to zero or is outside of the actual field count, an empty string is returned.
' If nPosition is negative then fields are searched from the right to left of the MainString.
' Delimiters are case-sensitive.
' Example: DIM cws AS CWSTR = AfxStrParseAny("1;2,3", 2, ",;")
' ========================================================================================
PRIVATE FUNCTION AfxStrParseAny (BYREF wszMainStr AS WSTRING, BYVAL nPosition AS LONG = 1, BYREF wszDelimiter AS WSTRING = ",") AS CWSTR
   ' The parse must match one character (len = 1) in the delimiter string
   RETURN AfxStrParse(wszMainStr, wszDelimiter, nPosition, TRUE, 1)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the length of the initial portion of a string which consists only of characters
' that are part of a specified set of characters.
' Example:
' DIM wszText AS WSTRING * 260 = "129th"
' DIM wszSet AS WSTRING * 260 = "1234567890"
' DIM n AS LONG = StrSpnW(@wszText, @wszSet)
' printf(!"The initial number has %d digits.\n", n)
' ========================================================================================
PRIVATE FUNCTION AfxStrSpn (BYREF wszText AS WSTRING, BYREF wszSet AS WSTRING) AS LONG
   RETURN StrSpnW(@wszText, @wszSet)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a numeric value into a string that represents the number expressed as a size
' value in bytes, kilobytes, megabytes, or gigabytes, depending on the size.
' ========================================================================================
PRIVATE FUNCTION AfxStrFormatByteSize (BYVAL ull AS ULONGLONG) AS CWSTR
   DIM cws AS CWSTR = SPACE(260)
   StrFormatByteSizeW(ull, cws.vptr, 260)
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a numeric value into a string that represents the number expressed as a size
' value in kilobytes.
' ========================================================================================
PRIVATE FUNCTION AfxStrFormatKBSize (BYVAL ull AS ULONGLONG) AS CWSTR
   DIM cws AS CWSTR = SPACE(260)
   StrFormatKBSizeW(ull, cws.vptr, 260)
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a time interval, specified in milliseconds, to a string.
' Parameters:
' - dwTimeMS: The time interval, in milliseconds.
' - digits  : The maximum number of significant digits to be represented in the output
'             string. Some examples are:
'             dwTimeMS digits   cwsOut
'             -------- ------ -----------
'                34000    3        34 sec
'                34000    2        34 sec
'                34000    1        30 sec
'                74000    3  1 min 14 sec
'                74000    2  1 min 10 sec
'                74000    1  1 min
' ========================================================================================
PRIVATE FUNCTION AfxStrFromTimeInterval (BYVAL dwTimeMS AS DWORD, BYVAL digits AS LONG) AS CWSTR
   DIM cws AS CWSTR = SPACE(260)
   StrFromTimeIntervalW(cws.vptr, 260, dwTimeMS, digits)
   RETURN cws
END FUNCTION
' ========================================================================================
