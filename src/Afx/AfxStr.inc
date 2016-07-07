' ########################################################################################
' Microsoft Windows
' File: AfxStr.inc
' Contents: String wrapper functions.
' Compiler: FreeBasic 32 & 64-bit, Unicode.
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "win/ole2.bi"
#include once "Afx/CbStr.inc"
USING Afx.CBStrClass

' ========================================================================================
' Returns a copy of a string with characters or strings removed.
' If cbMatchStr is not present in cbMainStr, all of cbMainStr is returned intact.
' This function is case-sensitive.
' Usage example:
' DIM cbs AS CBSTR = AfxRemove("[]Hello[]", "[]")
' MessageBoxW 0, cbs, "", MB_OK
' ========================================================================================
PRIVATE FUNCTION AfxRemove (BYREF cbMainStr AS CBSTR, BYREF cbMatchStr AS CBSTR) AS CBSTR
   DIM cbOutStr AS CBSTR = TYPE<CBSTR>(cbMainStr)
   DO
      DIM nPos AS LONG = INSTR(cbOutStr, cbMatchStr)
      IF nPos = 0 THEN EXIT DO
      cbOutStr = LEFT(**cbOutStr, nPos - 1) & MID(**cbOutStr, nPos + LEN(cbMatchStr))
   LOOP
   FUNCTION = cbOutStr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a copy of a string with characters or strings removed.
' If cbMatchStr is not present in cbMainStr, all of cbMainStr is returned intact.
' cbMatchStr specifies a list of single characters to be searched for individually,
' a match on any one of which will cause that character to be removed from the result.
' This function is case-sensitive.
' Usage example:
' Removing all "b", "a", and "c"
' DIM cbs AS CBSTR = AfxRemoveAny("abacadabra", "bac")
' MessageBoxW 0, cbs, "", MB_OK
' ========================================================================================
PRIVATE FUNCTION AfxRemoveAny (BYREF cbMainStr AS CBSTR, BYREF cbMatchStr AS CBSTR) AS CBSTR
   DIM cbOutStr AS CBSTR = TYPE<CBSTR>(cbMainStr)
   DIM nLen AS LONG =  LEN(cbMatchStr)
   DIM i AS LONG
   FOR i = 1 TO nLen
      cbOutStr = AfxRemove(cbOutStr, MID(**cbMatchStr, i, 1))
   NEXT
   FUNCTION = cbOutStr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Within a specified string, replace all occurrences of one string with another string.
' Replaces all occurrences of cbMatchStr in cbMainStr with cbReplaceWith
' The replacement can cause cbMainStr to grow or condense in size.
' When a match is found, the scan for the next match begins at the position immediately
' following the prior match.
' This function is case-sensitive.
' Usage example:
' DIM cbs AS CBSTR = AfxReplace("Hello World", "World", "Earth")
' MessageBoxW 0, cbs, "", MB_OK
' ========================================================================================
PRIVATE FUNCTION AfxReplace (BYREF cbMainStr AS CBSTR, BYREF cbMatchStr AS CBSTR, BYREF cbReplaceWith AS CBSTR) AS CBSTR
   DIM cbOutStr AS CBSTR = TYPE<CBSTR>(cbMainStr)
   DIM nPos AS LONG = 1
   DO
      nPos = INSTR(nPos, cbOutStr, cbMatchStr)
      IF nPos = 0 THEN EXIT DO
      cbOutStr = LEFT(**cbOutStr, nPos - 1) & **cbReplaceWith & MID(**cbOutStr, nPos + LEN(cbMatchStr))
      nPos += LEN(cbReplaceWith)
   LOOP
   RETURN cbOutStr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Within a specified string, replace all occurrences of any of the individual characters
' specified in the cbMatchStr string.
' cbReplaceWith must be a single character. This function does not replace words therefore
' cbMainStr will be the same size - it will not shrink or grow.
' This function is case-sensitive.
' ========================================================================================
PRIVATE FUNCTION AfxReplaceAny (BYREF cbMainStr AS CBSTR, BYREF cbMatchStr AS CBSTR, BYREF cbReplaceWith AS CBSTR) AS CBSTR
   DIM cbOutStr AS CBSTR = TYPE<CBSTR>(cbMainStr)
   IF LEN(cbMainStr) = 0 THEN RETURN cbOutStr
   IF LEN(cbMatchStr) = 0 THEN RETURN cbOutStr
   IF LEN(cbReplaceWith) <> 1 THEN RETURN cbOutStr
   DIM i AS LONG, x AS LONG
   FOR x = 1 TO LEN(cbMatchStr)
      FOR i = 1 TO LEN(cbMainStr)
         IF MID(**cbMatchStr, x, 1) = MID(**cbMainStr, i, 1) THEN
            MID(**cbOutStr, i, 1) = **cbReplaceWith
         END IF
      NEXT
   NEXT
   RETURN cbOutStr
END FUNCTION
' ========================================================================================

' ========================================================================================
'  Count the number of occurrences of specified characters strings within a string.
'  cbMainStr is the string expression in which to count characters.
'  cbMatchStr is a list of single characters to be searched for individually. A match on
'  any one of which will cause the count to be incremented for each occurrence of that
'  character. Note that repeated characters in cbMatchStr will not increase the count.
'  This function is case-sensitive.
' ========================================================================================
PRIVATE FUNCTION AfxTallyAny (BYREF cbMainStr AS CBSTR, BYREF cbMatchStr AS CBSTR) AS LONG

   IF LEN(cbMainStr) = 0 OR LEN(cbMatchStr) = 0 THEN EXIT FUNCTION

   ' // Remove possible duplicates in the matches string
   DIM AS LONG i, nPos
   DIM cbMatches AS CBSTR = TYPE<CBSTR>("")
   FOR i = 1 TO LEN(cbMatchStr)
      nPos = INSTR(cbMatches, MID(**cbMatchStr, i, 1))
      IF nPos = 0 THEN cbMatches = **cbMatches & MID(**cbMatchStr, i, 1)
   NEXT

   ' // Do the count
   DIM AS LONG nCount
   FOR i = 1 TO LEN(cbMatchStr)
      nPos = 1
      DO
         nPos = Instr(nPos, cbMainStr, MID(cbMatches, i, 1))
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
' Reverse the contents of a string expression.
' Usage example:
' DIM cbs AS CBSTR = AfxStrReverse("garden")
' MessageBoxW 0, cbs, "", MB_OK
' ========================================================================================
PRIVATE FUNCTION AfxStrReverse (BYREF cbMainStr AS CBSTR) AS CBSTR
   DIM cbOutStr AS CBSTR = TYPE<CBSTR>(cbMainStr)
   IF LEN(cbMainStr) = 0 THEN RETURN cbOutSTr
   DIM AS LONG i, nLen
   nLen = LEN(cbOutStr)
   DIM cbChar AS CBSTR = TYPE<CBSTR>("")
   FOR i = 1 TO nLen \ 2
      cbChar = MID(**cbOutStr, i, 1)
      MID(**cbOutStr, i, 1) = MID(**cbOutStr, nLen - i + 1, 1)
      MID(**cbOutStr, nLen - i + 1, 1) = **cbChar
   NEXT
   RETURN cbOutStr
END FUNCTION
' ========================================================================================


' ========================================================================================
' Complement to the AfxExtract function. Returns the portion of a string following the
' first occurrence of a character or group of characters.
' cbMainStr is searched for the string specified in cbMatchStr If found, all characters
' after cbMatchStr are returned. If cbMatchStr is not present in cbMainStr (or is null) then
' a zero-length empty string is returned.
' nStart is an optional starting position to begin searching. If nStart is not specified,
' position 1 will be used. If nStart is zero, a nul string is returned. If nStart is negative,
' the starting position is counted from right to left: if -1, the search begins at the last
' character; if -2, the second to last, and so forth.
' This function is case-sensitive.
' Usage example:
' DIM cbs AS CBSTR = AfxRemain("Brevity is the soul of wit", "is ")
' MessageBoxW 0, cbs, "", MB_OK
' ========================================================================================
PRIVATE FUNCTION AfxRemain (BYREF cbMainStr AS CBSTR, BYREF cbMatchStr AS CBSTR, BYVAL nStart AS LONG = 1) AS CBSTR
   DIM cbOutStr AS CBSTR = TYPE<CBSTR>("")
   IF LEN(cbMainStr) = 0 OR LEN(cbMatchStr) = 0 THEN RETURN cbOutStr
   IF nStart = 0 OR nStart > LEN(cbMainStr) THEN RETURN cbOutStr
   IF nStart < 0 THEN nStart = LEN(cbMainStr) + nStart + 1
   DIM nPos AS LONG = INSTR(nStart, **cbMainStr, **cbMatchStr)
   IF nPos = 0 THEN RETURN cbOutStr
   cbOutStr = MID(**cbMainStr, nPos + LEN(cbMatchStr))
   RETURN cbOutStr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Complement to the AfxExtract function. Returns the portion of a string following the
' first occurrence of a character or group of characters.
' cbMainStr is searched for the string specified in cbMatchStr If found, all characters
' after cbMatchStr are returned. If cbMatchStr is not present in cbMainStr (or is null) then
' a zero-length empty string is returned.
' cbMatchStr specifies a list of single characters to be searched for individually. A match
' on any one of which will cause the extract operation be performed after that character.
' nStart is an optional starting position to begin searching. If nStart is not specified,
' position 1 will be used. If nStart is zero, a nul string is returned. If nStart is negative,
' the starting position is counted from right to left: if -1, the search begins at the last
' character; if -2, the second to last, and so forth.
' This function is case-sensitive.
' Usage example:
' DIM cbs AS CBSTR = AfxRemain("Brevity is the soul of wit", "is ")
' MessageBoxW 0, cbs, "", MB_OK
' ========================================================================================
PRIVATE FUNCTION AfxRemainAny (BYREF cbMainStr AS CBSTR, BYREF cbMatchStr AS CBSTR, BYVAL nStart AS LONG = 1) AS CBSTR
   DIM cbOutStr AS CBSTR = TYPE<CBSTR>("")
   IF LEN(cbMainStr) = 0 OR LEN(cbMatchStr) = 0 THEN RETURN cbOutStr
   IF nStart = 0 OR nStart > LEN(cbMainStr) THEN RETURN cbOutStr
   IF nStart < 0 THEN nStart = LEN(cbMainStr) + nStart + 1
   DIM i AS LONG, x AS LONG
   FOR i = nStart TO LEN(cbMainStr)
      FOR x = 1 TO LEN(cbMatchStr)
         IF MID(**cbMainStr, i, 1) = MID(**cbMatchStr, x, 1) THEN
            cbOutStr = MID(**cbMainStr, i + 1)
            RETURN cbOutStr
         END IF
      NEXT
   NEXT
   RETURN cbOutStr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Count the number of occurrences of strings within a string.
' cbMainStr is the string expression in which to count characters.
' cbMatchStr is the string expression to count all occurrences of.
' If cbMatchStr is not present in cbMainStr, zero is returned.
' When a match is found, the scan for the next match begins at the position immediately
' following the prior match.
' This function is case-sensitive.
' Usage example:
' DIM nCount AS LONG = AfxTally("abacadabra", "ab")
' MessageBoxW 0, WSTR(nCount), "", MB_OK
' ========================================================================================
PRIVATE FUNCTION AfxTally (BYREF cbMainStr AS CBSTR, BYREF cbMatchStr AS CBSTR) AS LONG
   DIM nCount AS LONG
   DIM nPos AS LONG = 1
   DO
      nPos = INSTR(nPos, cbMainStr, cbMatchStr)
      IF nPos = 0 THEN EXIT DO
      nCount += 1
      nPos += LEN(cbMatchStr)
   LOOP
   RETURN nCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determine whether each character of a string is present in another string.
' Returns zero if each character in cbMainStr is present in cbMatchStr
' If not, it returns the position of the first non-matching character in MainString.
' This function is very useful for determining if a string contains only numeric
' digits, for example.
' This function is case-sensitive.
' If nStart evaluates to a position outside of the string, or if nStart is zero,
' then the function returns zero.
' ========================================================================================
PRIVATE FUNCTION AfxVerify (BYVAL nStart AS LONG, BYREF cbMainStr AS CBSTR, BYREF cbMatchStr AS CBSTR) AS LONG

   IF VARPTR(nStart) = NULL THEN EXIT FUNCTION
   IF nStart <= 0 OR nStart > LEN(cbMainStr) THEN RETURN 0

   ' Get each character in cbMainStr and look for it in cbMatchStr
   DIM AS LONG i, nPos, idx
   For i = nStart To LEN(cbMainStr)
      nPos = Instr(cbMatchStr, MID(cbMainStr, i, 1))
      IF nPos = 0 THEN
         idx = i
         EXIT FOR
      END IF
   NEXT
   FUNCTION =  idx

END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a string containing a left-justified (padded) string.
' If the optional parameter cbPadCharacter not specified, the function pads the string with
' space characters to the left. Otherwise, the function pads the string with the first
' character of cbPadCharacter.
' Usage example:
' DIM cbs AS CBSTR = AfxRSet("FreeBasic", 20, "*")
' MessageBoxW 0, cbs, "", MB_OK
' ========================================================================================
PRIVATE FUNCTION AfxLSet (BYREF cbMainStr AS CBSTR, BYVAL nStringLength AS LONG, BYREF cbPadCharacter AS CBSTR = " ") AS CBSTR
   DIM cbOutStr AS CBSTR = TYPE<CBSTR>("")
   cbOutStr = STRING(nStringLength, ASC(**cbPadCharacter, 1))
   MID(**cbOutStr, 1, LEN(cbMainStr)) = **cbMainStr
   RETURN cbOutStr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a string containing a right-justified (padded) string.
' If the optional parameter cbPadCharacter not specified, the function pads the string with
' space characters to the left. Otherwise, the function pads the string with the first
' character of cbPadCharacter.
' Usage example:
' DIM cbs AS CBSTR = AfxRSet("FreeBasic", 20, "*")
' MessageBoxW 0, cbs, "", MB_OK
' ========================================================================================
PRIVATE FUNCTION AfxRSet (BYREF cbMainStr AS CBSTR, BYVAL nStringLength AS LONG, BYREF cbPadCharacter AS CBSTR = " ") AS CBSTR
   DIM cbOutStr AS CBSTR = TYPE<CBSTR>("")
   cbOutStr = STRING(nStringLength, ASC(**cbPadCharacter, 1))
   MID(**cbOutStr, nStringLength - LEN(cbMainStr) + 1, LEN(cbMainStr)) = **cbMainStr
   RETURN cbOutStr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a string containing a right-justified (padded) string.
' If the optional parameter cbPadCharacter not specified, the function pads the string with
' space characters to the left. Otherwise, the function pads the string with the first
' character of cbPadCharacter.
' Usage example:
' DIM cbs AS CBSTR = AfxCSet("FreeBasic", 20, "*")
' MessageBoxW 0, cbs, "", MB_OK
' ========================================================================================
PRIVATE FUNCTION AfxCSet (BYREF cbMainStr AS CBSTR, BYVAL nStringLength AS LONG, BYREF cbPadCharacter AS CBSTR = " ") AS CBSTR
   DIM cbOutStr AS CBSTR = TYPE<CBSTR>("")
   cbOutStr = STRING(nStringLength, ASC(**cbPadCharacter, 1))
   MID(**cbOutStr, (nStringLength - LEN(cbMainStr)) \ 2 + 1, LEN(cbMainStr)) = **cbMainStr
   RETURN cbOutStr
END FUNCTION
' ========================================================================================

' ========================================================================================
'  Parses a path/file name to extract component parts
'  This function evaluates a text path/file text name, and returns a requested part of the
'  name. The functionality is strictly one of string parsing alone.
'  cbOption is one of the following words which is used to specify the requested part:
'  PATH
'        Returns the path portion of the path/file Name. That is the text up to and
'        including the last backslash (\) or colon (:).
'  NAME
'        Return the name portion of the path/file Name. That is the text To the right
'        of the last backslash (\) or colon (:), ending just before the last period (.).
'  EXTN
'        Returns the extension portion of the path/file name. That is the last
'        period (.) in the string plus the text to the right of it.
'  NAMEX
'        Returns the name and the EXTN parts combined.
' ========================================================================================
PRIVATE FUNCTION AfxPathName (BYREF cbOption AS CBSTR, BYREF cbFileSpec AS CBSTR) AS CBSTR

   DIM cbOutStr AS CBSTR = TYPE<CBSTR>(cbFileSpec)
   IF LEN(cbFileSpec) = 0 THEN RETURN cbOutStr

   SELECT CASE UCASE(cbOption)

      CASE "PATH"
         DIM nPos AS LONG = InstrRev(cbFileSpec, ANY ":/\")
         IF nPos THEN cbOutStr = LEFT(**cbFileSpec, nPos - 1)
         RETURN cbOutStr

      CASE "NAME"
         ' // Retrieve the full filename
         DIM nPos AS LONG = InstrRev(cbFileSpec, ANY ":/\")
         IF nPos THEN cbOutStr = MID(**cbFileSpec, nPos + 1)
         ' // Retrieve the filename
         nPos = InstrRev(cbOutStr, ".")
         IF nPos THEN cbOutStr = LEFT(**cbOutStr, nPos - 1)
         RETURN cbOutStr

      CASE "NAMEX"
         DIM nPos AS LONG = InStrRev(cbFileSpec, ANY ":/\")
         IF nPos THEN cbOutStr = MID(**cbFileSpec, nPos + 1) ELSE cbOutStr = ""
         RETURN cbOutStr

      CASE "EXTN"
         ' // Retrieve the full filename
         DIM nPos AS LONG = InstrRev(cbFileSpec, ANY ":/\")
         IF nPos THEN cbOutStr = MID(**cbFileSpec, nPos + 1)
         ' // Retrieve the extension
         nPos = InStrRev(cbOutStr, ".")
         IF nPos THEN cbOutStr = MID(**cbOutStr, nPos) ELSE cbOutStr = ""
         RETURN cbOutStr

   END SELECT

END FUNCTION
' ========================================================================================

' ========================================================================================
'  Adds paired characters to the beginning and end of a string.
'  It is particularly useful for enclosing text with parenthesess, quotes, brackets, etc.
'  For example:  AfxWrap("Paul", "<", ">") results in <Paul>
'  If only one wrap character/string is specified then that character or string is used
'  for both sides.
'  For example:  AfxWrap("Paul", "'") results in 'Paul'
'  If no wrap character/string is specified then double quotes are used.
'  For example: AfxWrap("Paul") results in "Paul"
' ========================================================================================
Private Function AfxWrap Overload (ByRef cbMainStr As CBSTR, ByRef cbLeftChar As CBSTR, ByRef cbRightChar As CBSTR) As CBSTR
   Dim cbOutStr As CBSTR = Type<CBSTR>("")
   cbOutStr = **cbLeftChar & **cbMainStr & **cbRightChar
   Function = cbOutStr
End Function
' ========================================================================================
' ========================================================================================
Private Function AfxWrap Overload (ByRef cbMainStr As CBSTR, ByRef cbChar As CBSTR = Chr(34)) As CBSTR
   Dim cbOutStr As CBSTR = Type<CBSTR>("")
   cbOutStr = **cbChar & **cbMainStr & **cbChar
   Function = cbOutStr
End Function
' ========================================================================================

' ========================================================================================
' Returns the count of delimited fields from a string expression.
' If cbMainStr is empty (a null string) or contains no delimiter character(s), the string 
' is considered to contain exactly one sub-field. In this case, AfxParseCount returns the value 1.
' Delimiter contains a string (one or more characters) that must be fully matched.
' Delimiters are case-sensitive.
' Usage example:
' DIM nCount AS LONG = AfxParseCount("one,two,three", ",")
' MessageBoxW 0, WSTR(nCount), "", MB_OK
' ========================================================================================
Private Function AfxParseCount (ByRef cbMainStr As CBSTR, ByRef cbDelimiter As CBSTR = ",") As Long
   Dim nCount As Long = 1
   Dim nPos As Long = 1
   Do
      nPos = Instr(nPos, cbMainStr, cbDelimiter)
      If nPos = 0 Then Exit Do
      nCount += 1
      nPos += Len(cbDelimiter)
   Loop
   Return nCount
End Function
' ========================================================================================

' ========================================================================================
' Return the count of delimited fields from a string expression.
' If MainString is empty (a null string) or contains no delimiter character(s), the string 
' is considered to contain exactly one sub-field. In this case, AfxParseCountAny returns the value 1.
' Delimiter contains a set of characters (one or more), any of which may act as a delimiter character.
' Delimiters are case-sensitive.
' Usage example:
' DIM nCount AS LONG = AfxParseCountAny("1;2,3", ",;")
' MessageBoxW 0, WSTR(nCount), "", MB_OK
' ========================================================================================
Private Function AfxParseCountAny (ByRef cbMainStr As CBSTR, ByRef cbDelimiter As CBSTR) As Long
   Dim nCount As Long = 1
   Dim nPos As Long = 1
   Do
      nPos = Instr(nPos, cbMainStr, Any cbDelimiter)
      If nPos = 0 Then Exit Do
      nCount += 1
      nPos += Len(cbDelimiter)
   Loop
   Return nCount
End Function
' ========================================================================================

' ========================================================================================
' Return a string consisting of multiple copies of the specified string.
' This function is very similar to STRING (which makes multiple copies of a single character).
' Usage example:
' DIM cbs AS CBSTR = AfxRepeat(5, "Paul")
' MessageBoxW 0, cbs, "", MB_OK
' ========================================================================================
Private Function AfxRepeat (ByVal nCount As Long, ByRef cbMainStr As CBSTR) As CBSTR

   Dim cbOutStr As CBSTR = Type<CBSTR>("")
   If nCount <= 0 Then Return cbOutStr
    
   ' Create the final full buffer and insert the
   ' strings into it in order to avoid nCount concatenations.
   Dim nLen As Long = Len(cbMainStr)
   cbOutStr = Space(nCount * nLen)
   
   Dim i As Long
   For i = 0 To nCount - 1
      Mid(**cbOutStr, (i * nLen) + 1, nLen) = **cbMainStr
   Next
    
   Return cbOutStr

End Function
' ========================================================================================

' ========================================================================================
' Returns a string with nCount characters removed from the left side of the string.
' If nCount is less than one then the entire string is returned. 
' ========================================================================================
Private Function AfxClipLeft (ByRef cbMainStr As CBSTR, ByVal nCount As Long) As CBSTR
   Dim cbOutStr As CBSTR = Type<CBSTR>(cbMainStr)
   If nCount <= 0 Then Return cbOutStr
   Dim nLen As Long = Len(cbMainStr)
   nCount = Iif(nLen < nCount, nLen, nCount)
   cbOutStr = Mid(**cbMainStr, nCount + 1)
   Return cbOutStr
End Function
' ========================================================================================

' ========================================================================================
' Returns a string with nCount characters removed from the right side of the string.
' If nCount is less than one then the entire string is returned. 
' ========================================================================================
Private Function AfxClipRight (ByRef cbMainStr As CBSTR, ByVal nCount As Long) As CBSTR
   Dim cbOutStr As CBSTR = Type<CBSTR>(cbMainStr)
   If nCount <= 0 Then Return cbOutStr
   Dim nLen As Long = Len(cbMainStr)
   nCount = nLen - nCount
   nCount = Iif(nLen < nCount, nLen, nCount)
   cbOutStr = Left(**cbMainStr, nCount)
   Return cbOutStr
End Function
' ========================================================================================

' ========================================================================================
' Returns a string with nCount characters removed starting at position nStart. The first
' character is considered position 1, the second position 2, etc...
' If nCount or nStart is less than one then the entire string is returned. 
' ========================================================================================
Private Function AfxClipMid (ByRef cbMainStr As CBSTR, ByVal nStart As Long, ByVal nCount As Long) As CBSTR
   Dim cbOutStr As CBSTR = Type<CBSTR>(cbMainStr)
   If (nCount <= 0) Or (nStart <= 0) Then Return cbOutStr
   Dim nLen As Long = Len(cbMainStr)
   cbOutStr = Left(**cbMainStr, nStart - 1) & Mid(**cbMainSTr, nStart + nCount)
   Return cbOutStr
End Function
' ========================================================================================

' ========================================================================================
' Remove paired characters to the beginning and end of a string.
' It is particularly useful for removing text with parenthesess, quotes, brackets, etc.
' For example: AfxUnWrap("<Paul>", "<", ">") results in Paul
' If only one unwrap character/string is specified then that character or string is used for both sides.
' For example: AfxUnWrap("'Paul'", "'") results in Paul
' If no wrap character/string is specified then double quotes are used.
' For example: AfxUnWrap("""Paul""") results in Paul
' ========================================================================================
Private Function AfxUnWrap Overload (ByRef cbMainStr As CBSTR, ByRef cbLeftChar As CBSTR, ByRef cbRightChar As CBSTR) As CBSTR
   Dim cbOutStr As CBSTR = Type<CBSTR>("")
   cbOutStr = LTrim(**cbMainStr, **cbLeftChar)
   cbOutStr = Rtrim(**cbOutStr, **cbRightChar)
   Function = cbOutStr
End Function
' ========================================================================================
' ========================================================================================
Private Function AfxUnWrap Overload (ByRef cbMainStr As CBSTR, ByRef cbChar As CBSTR = Chr(34)) As CBSTR
   Dim cbOutStr As CBSTR = Type<CBSTR>("")
   cbOutStr = LTrim(**cbMainStr, **cbChar)
   cbOutStr = Rtrim(**cbOutStr, **cbChar)
   Function = cbOutStr
End Function
' ========================================================================================

' ========================================================================================
' Delete a specified number of characters from a string expression.
' Returns a string based on MainString but with nCount characters deleted 
' starting at position nStart. The first character in the string is position 1, etc.
' ========================================================================================
Private Function AfxDelete (ByRef cbMainStr As CBSTR, ByVal nStart As Long, ByVal nCount As Long) As CBSTR
   Dim cbOutStr As CBSTR = Type<CBSTR>(cbMainStr)
   Dim nLen As Long = Len(cbMainStr)
   If nLen = 0 Or nStart < 0 Or nCount <= 0 Or nStart > nLen Then Return cbOutStr
   cbOutStr = Left(**cbMainStr, nStart) & Mid(**cbMainStr, nStart + 1 + nCount, nCount)
   Return cbOutStr
End Function
' ========================================================================================

' ========================================================================================
' Inserts a string at a specified position within another string expression.
' Returns a string consisting of cbMainStr with the string cbInsertString inserted 
' at nPosition. If nPosition is greater than the length of cbMainStr or <= zero then 
' cbInsertString is appended to cbMainStr The first character in the string is position 1, etc.
' ========================================================================================
Private Function AfxStrInsert (ByRef cbMainStr As CBSTR, ByRef cbInsertString As CBSTR, ByVal nPosition As Long) As CBSTR
   Dim cbOutStr As CBSTR = Type<CBSTR>(cbMainStr)
   If nPosition > Len(cbMainStr) Or nPosition <= 0 Then
      cbOutStr = **cbOutStr & **cbInsertString
   Else
      cbOutStr = Left(**cbMainStr, nPosition) & **cbInsertString & Mid(**cbMainStr, nPosition)
   End If
   Return cbOutStr
End Function
' ========================================================================================

' ========================================================================================
' Extract characters from a string up to a character or group of characters.
' Complement function to AfxRemain.
' Returns a substring of cbMainStr starting with its first character (or the character
' specified by nStart) and up to (but not including) the first occurrence of cbMatchStr
' If cbMatchStr is not present in cbMainStr (or is null) then all of MainString is
' returned from the nStart position.
' This function is case-sensitive.
' The following line returns "aba" (match on "cad")
' DIM cbs AS CBSTR = AfxExtract(1, "abacadabra","cad")
' MessageBoxW 0, cbs, "", MB_OK
' ========================================================================================
Private Function AfxExtract (ByVal nStart As Long, ByRef cbMainStr As CBSTR, ByRef cbMatchStr As CBSTR) As CBSTR
   Dim cbOutStr As CBSTR = Type<CBSTR>(cbMainStr)
   Dim nLen As Long = Len(cbMainStr)
   If (nStart = 0) Or (nStart > nLen) Then
      cbOutStr = ""
      Return cbOutStr
   End If
   If nStart < 0 Then nStart = nLen + nStart + 1
   Dim nPos As Long = Instr(nStart, cbMainStr, cbMatchStr)
   If nPos Then
      cbOutStr = Mid(cbMainStr, nStart, nPos - nStart )
   Else
      cbOutStr = Mid(cbMainStr, nStart)
   End If
   Return cbOutStr
End Function
' ========================================================================================

' ========================================================================================
' Extract characters from a string up to a specific character.
' Returns a substring of cbMainStr starting with its first character (or the character
' specified by nStart) and up to (but not including) the first occurrence of cbMatchStr
' cbMatchStr specifies a list of single characters to be searched for individually, a
' match on any one of which will cause the extract operation to be performed up to that character.
' If cbMatchStr is not present in MainString (or is null) then all of cbMainStr is returned.
' This function is case-sensitive.
' Usage example:
' DIM cbs AS CBSTR = AfxExtractAny(1, "abacadabra","cd")
' MessageBoxW 0, cbs, "", MB_OK
' ========================================================================================
Private Function AfxExtractAny (ByVal nStart As Long, ByRef cbMainStr As CBSTR, ByRef cbMatchStr As CBSTR) As CBSTR
   Dim cbOutStr As CBSTR = Type<CBSTR>(cbMainStr)
   Dim nLen As Long = Len(cbMainStr)
   If (nStart = 0) Or (nStart > nLen) Then
      cbOutStr = ""
      Return cbOutStr
   End If
   If nStart < 0 Then nStart = nLen + nStart + 1
   Dim As Long i, x
   For i = nStart To Len(cbMainStr)
      For x = 1 To Len(cbMatchStr)
         If Mid(cbMainStr, i, 1) = Mid(cbMatchStr, x, 1) Then
            cbOutStr = Mid(cbMainStr, nStart, i - nStart)
            Return cbOutStr
         End If
      Next
   Next
End Function
' ========================================================================================

' ========================================================================================
' Returns a string containing only the characters contained in a specified match string.
' All other characters are removed. If cbMatchStr is an empty string the function returns
' an empty string.
' This function is case-sensitive.
' Usage example:
' DIM cbs AS CBSTR = AfxRetain("abacadabra","b")
' MessageBoxW 0, cbs, "", MB_OK
' ========================================================================================
Private Function AfxRetain (ByRef cbMainStr As CBSTR, ByRef cbMatchStr As CBSTR) As CBSTR
   Dim cbOutStr As CBSTR = Type<CBSTR>("")
   If Len(cbMainStr) = 0 Or Len(cbMatchStr) = 0 Then Return cbOutStr
   Dim nPos As Long = 1
   Do
      nPos = Instr(nPos, cbMainStr, cbMatchStr)
      If nPos = 0 Then Exit Do
      cbOutStr = **cbOutStr & Mid(**cbMainStr, nPos, Len(cbMatchStr))
      nPos += Len(cbMatchStr)
   Loop
   Return cbOutStr
End Function
' ========================================================================================

' ========================================================================================
' Returns a string containing only the characters contained in a specified match string.
' All other characters are removed.
' If cbMatchStr is an empty string the function returns an empty string.
' cbMatchStr specifies a list of single characters to be searched for individually.
' A match on any one of which will cause that character to be removed from the result.
' This function is case-sensitive.
' Usage example:
' DIM cbs AS CBSTR = AfxRetainAny("<p>1234567890<ak;lk;l>1234567890</p>", "<;/p>")
' MessageBoxW 0, cbs, "", MB_OK
' ========================================================================================
Private Function AfxRetainAny (ByRef cbMainStr As CBSTR, ByRef cbMatchStr As CBSTR) As CBSTR
   Dim cbOutStr As CBSTR = Type<CBSTR>("")
   If Len(cbMainStr) = 0 Or Len(cbMatchStr) = 0 Then Return cbOutStr
   Dim As Long i, x, nPos
   For i = 1 To Len(cbMainStr)
      nPos = Instr(cbMatchStr, Mid(cbMainStr, i, 1))
      If nPos Then cbOutStr = **cbOutStr & Mid(cbMainStr, i, 1)
   Next
   Return cbOutStr
End Function
'==============================================================================

'==============================================================================
' Shrink a string to use a consistent single character delimiter.
' The purpose of this function is to create a string with consecutive data
' items (words) separated by a consistent single character. This makes it very
' straightforward to parse the results as needed.
' If cbMask is not defined then all leading spaces and trailing spaces are
' removed entirely. All occurrences of two or more spaces are changed to a
' single space. Therefore, the new string returned consists of zer' or more
' words, each separated by a single space character.
' If cbMask is specified, it defines one or more delimiter characters to shrink.
' All leading and trailing mask characters are removed entirely.
' All occurrences of one or more mask characters are replaced with the first
' character of cbMask. The new string returned consists of zero or more words,
' each separated by the character found in the first position of cbMask.
' WhiteSpace is generally defined as the four common non-printing characters:
' Space, Tab, Carriage-Return, and Line-Feed. cbMask = Chr(32,9,13,10)
' Usage example:
' DIM cbs AS CBSTR = AfxShrink(",,, one , two     three, four,", " ,")
' MessageBoxW 0, cbs, "", MB_OK
'==============================================================================
Private Function AfxShrink (ByRef cbMainStr As CBSTR, ByRef cbMask As CBSTR = " ") As CBSTR

   Dim cbOutStr As CBSTR = Type<CBSTR>("")
   If Len(cbMainStr) = 0 Or Len(cbMask) = 0 Then Return cbOutStr

   ' Eliminate all leading and trailing cbMask characters
   cbOutStr = Trim(**cbMainStr, Any **cbMask)

   ' Eliminate all duplicate sMask characters within the string
   Dim cbReplace As CBSTR = Type<CBSTR>(Left(**cbMask, 1))
   Dim cbDuplicate As CBSTR = Type<CBSTR>("")
   Dim nMaskLen As Long = Len(cbMask)
   Dim As Long i, nPos
   For i = 1 To nMaskLen
      cbDuplicate = Mid(**cbMask, i, 1) & Mid(**cbMask, i, 1)   ' usually double spaces
      nPos = 1
      Do
         nPos = Instr(cbOutStr, cbDuplicate)
         If nPos = 0 Then Exit Do
         cbOutStr = Left(**cbOutStr, nPos - 1) & **cbReplace & Mid(**cbOutStr, nPos + Len(cbDuplicate))
      Loop
   Next

   ' Replace all single characters in the mask with the first character of the mask.
   nPos = 1
   Do
      nPos = Instr(nPos, cbOutStr, Any cbMask)
      If nPos = 0 Then Exit Do
      ' Only do the replace if the character at the position found is
      ' different than the character we need to replace it with. This saves
      ' us from having to do an unneeded string concatenation.
      If Mid(**cbOutStr, nPos, 1) <> **cbReplace  Then
         cbOutStr = Left(**cbOutStr, nPos - 1) & **cbReplace & Mid(**cbOutStr, nPos + 1)
      End If
      nPos += 1
   Loop

   ' Finally, do a pass to ensure that there are no duplicates of the
   ' first mask character because of the replacements in the step above.
   cbDuplicate = Left(**cbMask, 1) & Left(**cbMask, 1)
   nPos = 1
   Do
      nPos = Instr(cbOutStr, cbDuplicate)
      If nPos = 0 Then Exit Do
      cbOutStr = Left(**cbOutStr, nPos - 1) & **cbReplace & Mid(**cbOutStr, nPos + Len(cbDuplicate))
   Loop

   Return cbOutStr

End Function
'==============================================================================

