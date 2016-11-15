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
#include once "Afx/AfxWin.inc"
#include once "Afx/CWStr.inc"
#include once "Afx/CWStrArray.inc"
USING Afx

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
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a copy of a string with strings removed.
' If wszMatchStr is not present in wszMainStr, all of wszMainStr is returned intact.
' This function is case sensitive.
' Example: DIM cws AS CWSTR = AfxStrRemove("[]Hello[]", "[]")
' ========================================================================================
PRIVATE FUNCTION AfxStrRemove (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING) AS CWSTR
   DIM cws AS CWSTR = wszMainStr
   DO
      DIM nPos AS LONG = INSTR(cws, wszMatchStr)
      IF nPos = 0 THEN EXIT DO
      cws.DelChars nPos, LEN(wszMatchStr)
   LOOP
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a copy of a string with characters removed.
' If wszMatchStr is not present in wszMainStr, all of wszMainStr is returned intact.
' wszMatchStr specifies a list of single characters to be searched for individually,
' a match on any one of which will cause that character to be removed from the result.
' This function is case sensitive.
' Example: DIM cws AS CWSTR = AfxStrRemoveAny("abacadabra", "bac")
' ========================================================================================
PRIVATE FUNCTION AfxStrRemoveAny (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING) AS CWSTR
   DIM cws AS CWSTR = wszMainStr
   DIM i AS LONG
   FOR i = 1 TO LEN(wszMatchStr)
      cws = AfxStrRemove(cws, MID(wszMatchStr, i, 1))
   NEXT
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Within a specified string, replace all occurrences of one string with another string.
' Replaces all occurrences of wszMatchStr in wszMainStr with wszReplaceWith
' The replacement can cause wszMainStr to grow or condense in size.
' When a match is found, the scan for the next match begins at the position immediately
' following the prior match.
' This function is case sensitive.
' Example: DIM cws AS CWSTR = AfxStrReplace("Hello World", "World", "Earth")
' ========================================================================================
PRIVATE FUNCTION AfxStrReplace OVERLOAD (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING, BYREF wszReplaceWith AS WSTRING) AS CWSTR
   DIM cws AS CWSTR = wszMainStr
   DIM nPos AS LONG = 1
   DO
      nPos = INSTR(nPos, cws, wszMatchStr)
      IF nPos = 0 THEN EXIT DO
      cws = MID(cws, 1, nPos - 1) & wszReplaceWith & MID(cws, nPos + LEN(wszMatchStr))
      nPos += LEN(wszReplaceWith)
   LOOP
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Within a specified string, replace all occurrences of any of the individual characters
' specified in the wszMainStr string.
' wszReplaceWith must be a single character. This function does not replace words therefore
' wszMatchStr will be the same size - it will not shrink or grow.
' This function is case-sensitive.
' Example: DIM cws AS CWSTR = AfxStrReplaceAny("abacadabra", "bac", "*")
' ========================================================================================
PRIVATE FUNCTION AfxStrReplaceAny (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING, BYREF wszReplaceWith AS WSTRING) AS CWSTR
   DIM cws AS CWSTR = wszMainStr
   IF LEN(wszMatchStr) = 0 THEN RETURN cws
   IF LEN(wszReplaceWith) = 0 THEN RETURN cws
   DIM i AS LONG, x AS LONG
   FOR x = 1 TO LEN(wszMatchStr)
      FOR i = 1 TO LEN(wszMainStr)
         IF MID(wszMatchStr, x, 1) = MID(wszMainStr, i, 1) THEN
            MID(**cws, i, 1) = wszReplaceWith
         END IF
      NEXT
   NEXT
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Reverses the contents of a string expression.
' Usage example: DIM cws AS CWSTR = AfxStrReverse("garden")
' ========================================================================================
PRIVATE FUNCTION AfxStrReverse (BYREF wszMainStr AS WSTRING) AS CWSTR
   DIM cws AS CWSTR = wszMainStr
   DIM wszChar AS WSTRING * 2
   DIM i AS LONG, nLen AS LONG = LEN(wszMainStr)
   FOR i = 1 TO nLen \ 2
      wszChar = MID(cws, i, 1)
      MID(**cws, i, 1) = MID(**cws, nLen - i + 1, 1)
      MID(**cws, nLen - i + 1, 1) = wszChar
   NEXT
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Complement to the AfxStrExtract function.
' Returns the portion of a string following the first occurrence of a substring.
' wszMainStr is searched for the string specified in wszMatchStr If found, all characters
' after cbMatchStr are returned. If wszMatchStr is not present in wszMainStr (or is null) then
' a zero-length empty string is returned.
' nStart is an optional starting position to begin searching. If nStart is not specified,
' position 1 will be used. If nStart is zero, a nul string is returned. If nStart is negative,
' the starting position is counted from right to left: if -1, the search begins at the last
' character; if -2, the second to last, and so forth.
' This function is case-sensitive.
' Example: DIM cws AS CWSTR = AfxStrRemain("Brevity is the soul of wit", "is ")
' ========================================================================================
PRIVATE FUNCTION AfxStrRemain (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING, BYVAL nStart AS LONG = 1) AS CWSTR
   DIM cws AS CWSTR = ""
   IF LEN(wszMainStr) = 0 OR LEN(wszMatchStr) = 0 THEN RETURN cws
   IF nStart = 0 OR nStart > LEN(wszMainStr) THEN RETURN cws
   IF nStart < 0 THEN nStart = LEN(wszMainStr) + nStart + 1
   DIM nPos AS LONG = INSTR(nStart, wszMainStr, wszMatchStr)
   IF nPos = 0 THEN RETURN cws
   cws = wszMainStr
   cws = MID(**cws, nPos + LEN(wszMatchStr))
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Complement to the AfxStrExtract function. Returns the portion of a string following the
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
' Example: DIM cws AS CWSTR = AfxStrRemainAny("I think, therefore I am", ",")
' ========================================================================================
PRIVATE FUNCTION AfxStrRemainAny (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING, BYVAL nStart AS LONG = 1) AS CWSTR
   DIM cws AS CWSTR = ""
   IF LEN(wszMainStr) = 0 OR LEN(wszMatchStr) = 0 THEN RETURN cws
   IF nStart = 0 OR nStart > LEN(wszMainStr) THEN RETURN cws
   IF nStart < 0 THEN nStart = LEN(wszMainStr) + nStart + 1
   DIM i AS LONG, x AS LONG
   FOR i = nStart TO LEN(wszMainStr)
      FOR x = 1 TO LEN(wszMatchStr)
         IF MID(wszMainStr, i, 1) = MID(wszMatchStr, x, 1) THEN
            cws = MID(wszMainStr, i + 1)
            RETURN cws
         END IF
      NEXT
   NEXT
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Count the number of occurrences of strings within a string.
' wszMainStr is the string expression in which to count characters.
' cbMatchStr is the string expression to count all occurrences of.
' If cbMatchStr is not present in wszMainStr, zero is returned.
' When a match is found, the scan for the next match begins at the position immediately
' following the prior match.
' This function is case-sensitive.
' Example: DIM nCount AS LONG = AfxStrTally("abacadabra", "ab")
' ========================================================================================
PRIVATE FUNCTION AfxStrTally (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING) AS LONG
   DIM nCount AS LONG, nPos AS LONG = 1
   DO
      nPos = INSTR(nPos, wszMainStr, wszMatchStr)
      IF nPos = 0 THEN EXIT DO
      nCount += 1
      nPos += LEN(wszMatchStr)
   LOOP
   RETURN nCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Count the number of occurrences of specified characters strings within a string.
' wszMainStr is the string expression in which to count characters.
' wszMatchStr is a list of single characters to be searched for individually. A match on
' any one of which will cause the count to be incremented for each occurrence of that
' character. Note that repeated characters in wszMatchStr will not increase the count.
' This function is case-sensitive.
' Example: DIM nCount AS LONG = AfxStrTallyAny("abacadabra", "bac")
' ========================================================================================
PRIVATE FUNCTION AfxStrTallyAny (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING) AS LONG
   IF LEN(wszMainStr) = 0 OR LEN(wszMatchStr) = 0 THEN EXIT FUNCTION
   ' // Remove possible duplicates in the matches string
   DIM AS LONG i, nPos
   DIM cwsMatchStr AS CWSTR = wszMatchStr
   FOR i = 1 TO LEN(cwsMatchStr)
      nPos = INSTR(cwsMatchStr, MID(wszMatchStr, i, 1))
      IF nPos = 0 THEN cwsMatchStr = cwsMatchStr & MID(wszMatchStr, i, 1)
   NEXT
   ' // Do the count
   DIM nCount AS LONG
   FOR i = 1 TO LEN(cwsMatchStr)
      nPos = 1
      DO
         nPos = Instr(nPos, wszMainStr, MID(cwsMatchStr, i, 1))
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
' Determine whether each character of a string is present in another string.
' Returns zero if each character in wszMainStr is present in wszMatchStr
' If not, it returns the position of the first non-matching character in wszMainStr.
' This function is very useful for determining if a string contains only numeric digits, for example.
' This function is case-sensitive.
' If nStart evaluates to a position outside of the string, or if nStart is zero, then the
' function returns zero.
' Example: DIM nCount AS LONG = AfxStrVerify(5, "123.65,22.5", "0123456789")
' ========================================================================================
PRIVATE FUNCTION AfxStrVerify (BYVAL nStart AS LONG, BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING) AS LONG
   IF nStart <= 0 OR nStart > LEN(wszMainStr) THEN RETURN 0
   ' // Get each character in wszMainStr and look for it in wszMatchStr
   DIM AS LONG i, nPos, idx
   For i = nStart TO LEN(wszMainStr)
      nPos = Instr(wszMatchStr, MID(wszMainStr, i, 1))
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
         RETURN cws
      CASE "NAME"
         ' // Retrieve the full filename
         cws = wszFileSpec
         DIM nPos AS LONG = InstrRev(wszFileSpec, ANY ":/\")
         IF nPos THEN cws = MID(wszFileSpec, nPos + 1)
         ' // Retrieve the filename
         nPos = InstrRev(cws, ".")
         IF nPos THEN cws = MID(cws, 1, nPos - 1)
         RETURN cws
      CASE "NAMEX"
         ' // Retrieve the name and extension combined
         DIM nPos AS LONG = InStrRev(wszFileSpec, ANY ":/\")
         IF nPos THEN cws = MID(wszFileSpec, nPos + 1) ELSE cws = wszFileSpec
         RETURN cws
      CASE "EXTN"
         ' // Retrieve the name and extension combined
         DIM nPos AS LONG = InstrRev(wszFileSpec, ANY ":/\")
         IF nPos THEN cws = MID(wszFileSpec, nPos + 1) ELSE cws = wszFileSpec
         ' // Retrieve the extension
         nPos = InStrRev(cws, ".")
         IF nPos THEN cws = MID(cws, nPos) ELSE cws = ""
         RETURN cws
   END SELECT
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
   DIM i AS LONG
   FOR i = 0 TO nCount - 1
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
   cws = LEFT(wszMainStr, nStart - 1) & MID(wszMainStr, nStart + nCount)
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
   DIM cws AS CWSTR = wszLeftChar & wszMainStr & wszRightChar
   RETURN cws
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxStrWrap OVERLOAD (BYREF wszMainStr AS WSTRING, BYREF wszChar AS WSTRING = CHR(34)) AS CWSTR
   DIM cws AS CWSTR = wszChar & wszMainStr & wszChar
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
' Deletes a specified number of characters from a string expression.
' Returns a string based on wszMainStr but with nCount characters deleted
' starting at position nStart. The first character in the string is position 1, etc.
' Usage example:
' DIM cws AS CWSTR = AfxStrDelete("1234567890", 4, 3)
' ========================================================================================
PRIVATE FUNCTION AfxStrDelete (BYREF wszMainStr AS WSTRING, BYVAL nStart AS LONG, BYVAL nCount AS LONG) AS CWSTR
   DIM cws AS CWSTR = wszMainStr
   DIM nLen AS LONG = LEN(wszMainStr)
   IF nLen = 0 OR nStart < 0 OR nCount <= 0 OR nStart > nLen THEN RETURN cws
   cws = LEFT(wszMainStr, nStart) & MID(wszMainStr, nStart + 1 + nCount, nCount)
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts a string at a specified position within another string expression.
' Returns a string consisting of wszMainStr with the string wszInsertString inserted
' at nPosition. If nPosition is greater than the length of wszMainStr or <= zero then
' wszInsertString is appended to wszMainStr. The first character in the string is position 1, etc.
' DIM cws AS CWSTR = AfxStrInsert("1234567890", "--", 6)
' ========================================================================================
PRIVATE FUNCTION AfxStrInsert (BYREF wszMainStr AS WSTRING, BYREF wszInsertString AS WSTRING, BYVAL nPosition AS LONG) AS CWSTR
   DIM cws AS CWSTR = wszMainStr
   IF nPosition > LEN(wszMainStr) OR nPosition <= 0 THEN
      cws = cws & wszInsertString
   ELSE
      cws = MID(wszMainStr, 1, nPosition) & wszInsertString & MID(wszMainStr, nPosition + 1)
   END IF
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts characters from a string up to a character or group of characters.
' Complement function to AfxStrRemain.
' Returns a substring of wszMainStr starting with its first character (or the character
' specified by nStart) and up to (but not including) the first occurrence of wszMatchStr
' If wszMatchStr is not present in wszMainStr (or is null) then all of wszMainStr is
' returned from the nStart position.
' This function is case-sensitive.
' The following line returns "aba" (match on "cad")
' DIM cws AS CWSTR = AfxStrExtract(1, "abacadabra","cad")
' ========================================================================================
PRIVATE FUNCTION AfxStrExtract (BYVAL nStart AS LONG, BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING) AS CWSTR
   DIM cws AS CWSTR = wszMainStr
   DIM nLen AS LONG = LEN(wszMainStr)
   IF (nStart = 0) OR (nStart > nLen) THEN
      cws = ""
      RETURN cws
   END IF
   IF nStart < 0 THEN nStart = nLen + nStart + 1
   DIM nPos AS LONG = Instr(nStart, wszMainStr, wszMatchStr)
   IF nPos THEN
      cws = MID(wszMainStr, nStart, nPos - nStart )
   ELSE
      cws = MID(wszMainStr, nStart)
   END IF
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extract characters from a string up to a specific character.
' Returns a substring of wszMainStr starting with its first character (or the character
' specified by nStart) and up to (but not including) the first occurrence of wszMatchStr.
' wszMatchStr specifies a list of single characters to be searched for individually, a
' match on any one of which will cause the extract operation to be performed up to that character.
' If wszMatchStr is not present in wszMainStr (or is null) then all of wszMainStr is returned.
' This function is case-sensitive.
' Example: DIM cws AS CWSTR = AfxStrExtractAny(1, "abacadabra","cd")
' ========================================================================================
PRIVATE FUNCTION AfxStrExtractAny (BYVAL nStart AS LONG, BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING) AS CWSTR
   DIM cws AS CWSTR = wszMainStr
   DIM nLen AS LONG = LEN(wszMainStr)
   IF (nStart = 0) OR (nStart > nLen) THEN
      cws = ""
      RETURN cws
   END IF
   IF nStart < 0 THEN nStart = nLen + nStart + 1
   DIM AS LONG i, x
   FOR i = nStart TO LEN(wszMainStr)
      FOR x = 1 TO LEN(wszMatchStr)
         IF MID(wszMainStr, i, 1) = MID(wszMatchStr, x, 1) THEN
            cws = MID(wszMainStr, nStart, i - nStart)
            RETURN cws
         END IF
      NEXT
   NEXT
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a string containing only the characters contained in a specified match string.
' All other characters are removed. If wszMatchStr is an empty string the function returns
' an empty string.
' This function is case-sensitive.
' Example: DIM cws AS CWSTR = AfxStrRetain("abacadabra","b")
' ========================================================================================
PRIVATE FUNCTION AfxStrRetain (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING) AS CWSTR
   DIM cws AS CWSTR = ""
   IF LEN(wszMainStr) = 0 OR LEN(wszMatchStr) = 0 THEN RETURN cws
   DIM nPos AS LONG = 1
   DO
      nPos = Instr(nPos, wszMainStr, wszMatchStr)
      IF nPos = 0 THEN EXIT DO
      cws = cws & MID(wszMainStr, nPos, LEN(wszMatchStr))
      nPos += LEN(wszMatchStr)
   LOOP
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a string containing only the characters contained in a specified match string.
' All other characters are removed.
' If wszMatchStr is an empty string the function returns an empty string.
' wszMatchStr specifies a list of single characters to be searched for individually.
' A match on any one of which will cause that character to be removed from the result.
' This function is case-sensitive.
' Example: DIM cws AS CWSTR = AfxStrRetainAny("<p>1234567890<ak;lk;l>1234567890</p>", "<;/p>")
' ========================================================================================
PRIVATE FUNCTION AfxStrRetainAny (BYREF wszMainStr AS WSTRING, BYREF wszMatchStr AS WSTRING) AS CWSTR
   DIM cws AS CWSTR = ""
   IF LEN(wszMainStr) = 0 OR LEN(wszMatchStr) = 0 THEN RETURN cws
   DIM AS LONG i, x, nPos
   FOR i = 1 TO LEN(wszMainStr)
      nPos = Instr(wszMatchStr, MID(wszMainStr, i, 1))
      IF nPos THEN cws = cws & MID(wszMainStr, i, 1)
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
' Shrinks a string to use a consistent single character delimiter.
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
   DIM AS LONG i, nPos
   FOR i = 1 TO nMaskLen
      wszDuplicate = MID(wszMask, i, 1) & MID(wszMask, i, 1)   ' usually double spaces
      nPos = 1
      DO
         nPos = Instr(cws, wszDuplicate)
         IF nPos = 0 THEN EXIT DO
         cws = MID(cws, 1, nPos - 1) & wszReplace & MID(cws, nPos + LEN(wszDuplicate))
      LOOP
   NEXT
   ' // Replace all single characters in the mask with the first character of the mask.
   nPos = 1
   DO
      nPos = Instr(nPos, cws, ANY wszMask)
      IF nPos = 0 THEN EXIT DO
      ' Only do the replace if the character at the position found is
      ' different than the character we need to replace it with. This saves
      ' us from having to do an unneeded string concatenation.
      IF MID(cws, nPos, 1) <> wszReplace  THEN
         cws = MID(cws, 1, nPos - 1) & wszReplace & MID(cws, nPos + 1)
      END IF
      nPos += 1
   LOOP
   ' Finally, do a pass to ensure that there are no duplicates of the
   ' first mask character because of the replacements in the step above.
   wszDuplicate = MID(wszMask, 1, 1) & MID(wszMask, 1, 1)
   nPos = 1
   DO
      nPos = Instr(cws, wszDuplicate)
      IF nPos = 0 THEN EXIT DO
      cws = MID(cws, 1, nPos - 1) & wszReplace & MID(cws, nPos + LEN(wszDuplicate))
   LOOP
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the count of delimited fields from a string expression.
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
      nPos = Instr(nPos, wszMainStr, wszDelimiter)
      IF nPos = 0 THEN EXIT DO
      nCount += 1
      nPos += LEN(wszDelimiter)
   LOOP
   RETURN nCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Return the count of delimited fields from a string expression.
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
      nPos = Instr(nPos, wszMainStr, ANY wszDelimiter)
      IF nPos = 0 THEN EXIT DO
      nCount += 1
      nPos += LEN(wszDelimiter)
   LOOP
   RETURN nCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the nPosition-th substring in a string wszMainStr with separations wszDelimiter
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
         nPos = Instr(nStart, wszMainStr, ANY wszDelimiter)
      ELSE
         nPos = Instr(nStart, wszMainStr, wszDelimiter)
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
            nPos = Instr(nPos, wszMainStr, ANY wszDelimiter)
         ELSE
            nPos = Instr(nPos, wszMainStr, wszDelimiter)
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
' Returns a delimited field from a string expression.
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
'==============================================================================

'==============================================================================
' Return a delimited field from a string expression.
' Delimiter contains a set of characters (one or more), any of which may act as a delimiter character.
' If nPosition evaluates to zero or is outside of the actual field count, an empty string is returned.
' If nPosition is negative then fields are searched from the right to left of the MainString.
' Delimiters are case-sensitive.
' Example: DIM cws AS CWSTR = AfxStrParseAny("1;2,3", 2, ",;")
'==============================================================================
PRIVATE FUNCTION AfxStrParseAny (BYREF wszMainStr AS WSTRING, BYVAL nPosition AS LONG = 1, BYREF wszDelimiter AS WSTRING = ",") AS CWSTR
   ' The parse must match one character (len = 1) in the delimiter string
   RETURN AfxStrParse(wszMainStr, wszDelimiter, nPosition, TRUE, 1)
END FUNCTION
'==============================================================================

' ========================================================================================
' Splits a string into tokens, which are sequences of contiguous characters separated by
' any of the characters that are part of delimiters.
' - wszStr = The string to split.
' - wszDelimiters = The delimiter characters.
' Return value: A CBSTRSA (safe array) containing a token in each element.
' Usage example:
' DIM cws AS CWSTR = "- This, a sample string."
' DIM cwsa AS CWstrArray = AfxStrSplit(cws, " ,.-")
' FOR i AS LONG = cwsa.LBound TO cwsa.UBound
'    PRINT cwsa.Item(i)
' NEXT
' ========================================================================================
PRIVATE FUNCTION AfxStrSplit (BYREF wszStr AS WSTRING, BYREF wszDelimiters AS WSTRING = " ") AS CWstrArray
   DIM cws AS CWSTR = wszStr
   DIM cwsa AS CWstrArray
   DIM pwsz AS WSTRING PTR = wcstok(cws, @wszDelimiters)
   WHILE pwsz <> NULL
      cwsa.Append(pwsz)
      pwsz = wcstok(NULL, @wszDelimiters)
   WEND
   RETURN cwsa
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a string consisting of all of the strings in an array, each separated by a delimiter.
' If the delimiter is a null (zero-length) string then no separators are inserted between
' the string sections. If the delimiter expression is the 3-byte value of "," which may be
' expressed in your source code as the string literal """,""" or as Chr(34,44,34) then a
' leading and trailing double-quote is added to each string section. This ensures that the
' returned string contains standard comma-delimited quoted fields that can be easily parsed.
' Usage example:
' DIM cwsa AS CWstrArray = CWstrArray(3, 1)
' cwsa.Item(1) = "One"
' cwsa.Item(2) = "Two"
' cwsa.Item(3) = "Three"
' DIM cws AS CWSTR = AfxStrJoin(cwsa, ",")
' PRINT cws   ' ouput: One,Two,Three
' ========================================================================================
PRIVATE FUNCTION AfxStrJoin (BYREF cwsa AS CWstrArray, BYREF wszDelimiter AS WSTRING = " ") AS CWSTR
   ' // Add a leading ""
   DIM cws AS CWSTR
   IF wszDelimiter = CHR(34, 44, 34) THEN cws = CHR(34)
   DIM nUBound AS LONG = cwsa.UBound
   FOR i AS LONG = cwsa.LBound TO cwsa.UBound
      cws += cwsa.Item(i)
      IF i <> nUBound AND wszDelimiter <> "" THEN cws += wszDelimiter
   NEXT
   ' // Add a trailing ""
   IF wszDelimiter = CHR(34, 44, 34) THEN cws += CHR(34)
   RETURN cws
END FUNCTION
' ========================================================================================
