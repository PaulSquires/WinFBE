' ########################################################################################
' Microsoft Windows
' Implements a dynamic data type for null terminated unicode strings.
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2017 José Roca and Paul Squires.
' Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma ONCE
#INCLUDE ONCE "windows.bi"
#INCLUDE ONCE "win/ole2.bi"
#INCLUDE ONCE "/crt/string.bi"
#INCLUDE ONCE "/crt/wchar.bi"
#INCLUDE ONCE "utf_conv.bi"
#INCLUDE ONCE "Afx/CWSTR.inc"

' // The definition for BSTR in the FreeBASIC headers was inconveniently changed to WCHAR
#ifndef AFX_BSTR
   #define AFX_BSTR WSTRING PTR
#ENDIF

NAMESPACE Afx

' // Forward reference
TYPE CUTF16_ AS CUTF16
TYPE CUTF16_ AS CUTF16

' ========================================================================================
' Macro for debug
' To allow debugging, define _CUTF16_DEBUG_ 1 in your application before including this file.
' ========================================================================================
#ifndef _CUTF16_DEBUG_
   #define _CUTF16_DEBUG_ 0
#ENDIF
#ifndef _CUTF16_DP_
   #define _CUTF16_DP_ 1
   #MACRO CUTF16_DP(st)
      #IF (_CUTF16_DEBUG_ = 1)
         OutputDebugStringW(st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

' ========================================================================================
' // Checks if the passed pointer is a BSTR.
' // Will return FALSE if it is a null pointer.
' // If it is an OLE string it must have a descriptor; otherwise, don't.
' // Get the length in bytes looking at the descriptor and divide by 2 to get the number of
' // unicode characters, that is the value returned by the FreeBASIC LEN operator.
' // If the retrieved length if the same that the returned by LEN, then it must be an OLE string.
' ========================================================================================
FUNCTION AfxIsUtf16 (BYVAL pv aS ANY PTR) AS BOOLEAN
   IF pv = NULL THEN RETURN FALSE
   DIM res AS DWORD = PEEK(DWORD, pv - 4) \ 2
   IF res = LEN(*cast(WSTRING PTR, pv)) THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ########################################################################################
'                                  *** CUTF16 CLASS ***
' ########################################################################################
TYPE CUTF16

   Private:
      m_Capacity AS UINT         ' The total size of the buffer in UTF16 characters
      m_GrowSize AS LONG = 260   ' How much to grow the buffer by when required
      m_CodePage AS UINT         ' Unicode code page

   Public:
      m_pBuffer AS WSTRING PTR   ' Pointer to the buffer
      m_BufferLen AS UINT        ' Length in characters of the string

      DECLARE CONSTRUCTOR
      DECLARE CONSTRUCTOR (BYVAL nChars AS UINT, BYVAL nCodePage AS UINT)
      DECLARE CONSTRUCTOR (BYVAL pwszStr AS WSTRING PTR)
      DECLARE CONSTRUCTOR (BYREF ansiStr AS STRING, BYVAL nCodePage AS UINT = 0)
      DECLARE CONSTRUCTOR (BYREF cws AS CUTF16)
      DECLARE CONSTRUCTOR (BYREF cws AS CWSTR)
      DECLARE CONSTRUCTOR (BYREF cbs AS CBSTR)
      DECLARE CONSTRUCTOR (BYVAL n AS LONGINT)
      DECLARE CONSTRUCTOR (BYVAL n AS DOUBLE)
      DECLARE DESTRUCTOR
      DECLARE PROPERTY GrowSize () AS LONG
      DECLARE PROPERTY GrowSize (BYVAL nValue AS LONG)
      DECLARE PROPERTY Capacity () AS UINT
      DECLARE PROPERTY Capacity (BYVAL nValue AS UINT)
      DECLARE PROPERTY SizeAlloc (BYVAL nChars AS UINT)
      DECLARE PROPERTY SizeOf () AS UINT
      DECLARE PROPERTY CodePage () AS UINT
      DECLARE PROPERTY CodePage (BYVAL nCodePage AS UINT)
      DECLARE FUNCTION ResizeBuffer (BYVAL nChars AS UINT) AS WSTRING PTR
      DECLARE FUNCTION AppendBuffer (BYVAL memAddr AS ANY PTR, BYVAL nChars AS UINT) AS BOOLEAN
      DECLARE FUNCTION InsertBuffer (BYVAL memAddr AS ANY PTR, BYVAL nIndex AS UINT, BYVAL nChars AS UINT) AS BOOLEAN
      DECLARE SUB Clear
      DECLARE FUNCTION Add (BYVAL pwszStr AS WSTRING PTR) AS BOOLEAN
      DECLARE FUNCTION Add (BYREF ansiStr AS STRING, BYVAL nCodePage AS UINT = 0) AS BOOLEAN
      DECLARE FUNCTION Add (BYREF cws AS CUTF16) AS BOOLEAN
      DECLARE FUNCTION Add (BYREF cbs AS CWSTR) AS BOOLEAN
      DECLARE FUNCTION Add (BYREF cbs AS CBSTR_) AS BOOLEAN
      DECLARE PROPERTY Char(BYVAL nIndex AS UINT) AS USHORT
      DECLARE PROPERTY Char(BYVAL nIndex AS UINT, BYVAL nValue AS USHORT)
      DECLARE OPERATOR [] (BYVAL nIndex AS UINT) AS USHORT
      DECLARE FUNCTION DelChars (BYVAL nIndex AS UINT, BYVAL nCount AS UINT) AS BOOLEAN
      DECLARE FUNCTION Insert (BYVAL pwszStr AS WSTRING PTR, BYVAL nIndex AS UINT) AS BOOLEAN
      DECLARE FUNCTION Insert (BYREF ansiStr AS STRING, BYVAL nIndex AS UINT, BYVAL nCodePage AS UINT = 0) AS BOOLEAN
		DECLARE FUNCTION Insert (BYREF cws AS CUTF16, BYVAL nIndex AS UINT) AS BOOLEAN
		DECLARE FUNCTION Insert (BYREF cws AS CWSTR, BYVAL nIndex AS UINT) AS BOOLEAN
      DECLARE FUNCTION Insert (BYREF cbs AS CBSTR_, BYVAL nIndex AS LONG) AS BOOLEAN
      DECLARE OPERATOR CAST () BYREF AS WSTRING
      DECLARE OPERATOR CAST () AS ANY PTR
      DECLARE OPERATOR LET (BYREF wszStr AS CONST WSTRING)
      DECLARE OPERATOR LET (BYREF pwszStr AS WSTRING PTR)
      DECLARE OPERATOR LET (BYREF ansiStr AS STRING)
      DECLARE OPERATOR LET (BYREF cws AS CUTF16)
      DECLARE OPERATOR LET (BYREF cws AS CWSTR)
      DECLARE OPERATOR LET (BYREF cbs AS CBSTR_)
      DECLARE OPERATOR LET (BYVAL n AS LONGINT)
      DECLARE OPERATOR LET (BYVAL n AS DOUBLE)
'      DECLARE OPERATOR @ () AS WSTRING PTR
      DECLARE OPERATOR += (BYREF ansiStr AS STRING)
      DECLARE OPERATOR += (BYREF wszStr AS WSTRING)
      DECLARE OPERATOR += (BYREF cws AS CUTF16)
      DECLARE OPERATOR += (BYREF cws AS CWSTR)
      DECLARE OPERATOR += (BYREF cbs AS CBSTR_)
      DECLARE OPERATOR += (BYVAL n AS LONGINT)
      DECLARE OPERATOR += (BYVAL n AS DOUBLE)
      DECLARE OPERATOR &= (BYREF ansiStr AS STRING)
      DECLARE OPERATOR &= (BYREF wszStr AS WSTRING)
      DECLARE OPERATOR &= (BYREF cws AS CUTF16)
      DECLARE OPERATOR &= (BYREF cws AS CWSTR)
      DECLARE OPERATOR &= (BYREF cbs AS CBSTR_)
      DECLARE OPERATOR &= (BYVAL n AS LONGINT)
      DECLARE OPERATOR &= (BYVAL n AS DOUBLE)
      DECLARE FUNCTION vptr () AS WSTRING PTR
      DECLARE FUNCTION sptr () AS WSTRING PTR
      DECLARE FUNCTION wstr () BYREF AS WSTRING
      DECLARE FUNCTION wchar () AS WSTRING PTR
      DECLARE PROPERTY utf8 () AS STRING
      DECLARE PROPERTY utf8 (BYREF ansiStr AS STRING)
      DECLARE FUNCTION cbstr () AS CBSTR_
      DECLARE FUNCTION bstr () AS AFX_BSTR
      DECLARE FUNCTION LeftChars (BYVAL nChars AS LONG) AS CUTF16
      DECLARE FUNCTION RightChars (BYVAL nChars AS LONG) AS CUTF16
      DECLARE FUNCTION MidChars (BYVAL nStart AS LONG, BYVAL nChars AS LONG = 0) AS CUTF16
      DECLARE FUNCTION ValLong () AS LONG
      DECLARE FUNCTION ValInt () AS LONG
      DECLARE FUNCTION ValULong () AS ULONG
      DECLARE FUNCTION ValUInt () AS ULONG
      DECLARE FUNCTION ValLongInt () AS LONGINT
      DECLARE FUNCTION ValULongInt () AS ULONGINT
      DECLARE FUNCTION ValDouble () AS DOUBLE
      DECLARE FUNCTION Value () AS DOUBLE

END TYPE
' ########################################################################################

' ========================================================================================
' CUTF16 constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CUTF16
   CUTF16_DP("CUTF16 CONSTRUCTOR - Default")
   this.ResizeBuffer(m_GrowSize)   ' Create the initial buffer
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CUTF16 (BYVAL nChars AS UINT, BYVAL nCodePage AS UINT)
   CUTF16_DP("CUTF16 CONSTRUCTOR - Default")
   m_CodePage = nCodePage
   IF nChars = 0 THEN nChars = m_GrowSize
   this.ResizeBuffer(nChars)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CUTF16 (BYVAL pwszStr AS WSTRING PTR)
   CUTF16_DP("CUTF16 CONSTRUCTOR - WSTRING PTR")
   IF pwszStr = NULL THEN this.ResizeBuffer(m_GrowSize) : EXIT CONSTRUCTOR
   this.Add(pwszStr)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CUTF16 (BYREF ansiStr AS STRING, BYVAL nCodePage AS UINT = 0)
   CUTF16_DP("CUTF16 CONSTRUCTOR - STRING")
   m_CodePage = nCodePage
   IF .LEN(ansiStr) = 0 THEN this.ResizeBuffer(m_GrowSize) : EXIT CONSTRUCTOR
   this.Add(ansiStr, nCodePage)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CUTF16 (BYREF cws AS CUTF16)
   CUTF16_DP("CUTF16 CONSTRUCTOR - CUTF16")
   IF cws.m_BufferLen = 0 THEN this.ResizeBuffer(m_GrowSize) : EXIT CONSTRUCTOR
   this.Add(cws)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CUTF16 (BYREF cws AS CWSTR)
   CUTF16_DP("+++BEGIN- CUTF16 CONSTRUCTOR CWSTR")
   IF cws.m_BufferLen = 0 THEN this.ResizeBuffer(m_GrowSize) : EXIT CONSTRUCTOR
   this.Add(cws)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CUTF16 (BYREF cbs AS CBSTR_)
   CUTF16_DP("+++BEGIN- CUTF16 CONSTRUCTOR CBSTR")
   IF SysStringLen(cbs.m_bstr) = 0 THEN this.ResizeBuffer(m_GrowSize) : EXIT CONSTRUCTOR
   this.AppendBuffer(cbs.m_bstr, SysStringLen(cbs.m_bstr))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CUTF16 (BYVAL n AS LONGINT)
   CUTF16_DP("+++BEGIN- CUTF16 CONSTRUCTOR LONGINT")
   DIM wsz AS WSTRING * 260 = .WSTR(n)
   this.Add(wsz)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CUTF16 (BYVAL n AS DOUBLE)
   CUTF16_DP("+++BEGIN- CUTF16 CONSTRUCTOR DOUBLE")
   DIM wsz AS WSTRING * 260 = .WSTR(n)
   this.Add(wsz)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR CUTF16
   CUTF16_DP("***CUTF16 DESTRUCTOR - buffer: " & .WSTR(m_pBuffer))
   IF m_pBuffer THEN Deallocate(m_pBuffer)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Number of characters to preallocate to minimize multiple allocations when doing multiple
' concatenations. A value of less than 0 indicates that it must double the capacity each
' time that the buffer needs to be resized.
' ========================================================================================
PRIVATE PROPERTY CUTF16.GrowSize() AS LONG
   CUTF16_DP("CUTF16 PROPERTY GET GrowSize")
   PROPERTY = m_GrowSize
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CUTF16.GrowSize (BYVAL nChars AS LONG)
   CUTF16_DP("CUTF16 PROPERTY SET Growsize")
   m_GrowSize = nChars
END PROPERTY
' ========================================================================================

' ========================================================================================
' The size of the internal string buffer is retrieved and returned to the caller. The size
' is the number of characters which can be stored without further expansion.
' ========================================================================================
PRIVATE PROPERTY CUTF16.Capacity() AS UINT
   CUTF16_DP("CUTF16 PROPERTY GET Capacity")
   PROPERTY = m_Capacity
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CUTF16.SizeOf() AS UINT
   CUTF16_DP("CUTF16 PROPERTY GET SizeOf")
   PROPERTY = m_Capacity
END PROPERTY
' ========================================================================================
' ========================================================================================
' The internal string buffer is expanded to the specified number of characters. If the new
' capacity is smaller or equal to the current capacity, no operation is performed. If it is
' smaller, the buffer is shortened and the contents that exceed the new capacity are lost.
' ========================================================================================
PRIVATE PROPERTY CUTF16.Capacity (BYVAL nValue AS UINT)
   CUTF16_DP("CUTF16 PROPERTY SET Capacity")
   IF nValue = m_Capacity THEN EXIT PROPERTY
   this.ResizeBuffer(nValue)
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CUTF16.SizeAlloc (BYVAL nValue AS UINT)
   CUTF16_DP("CUTF16 PROPERTY SET SizeAlloc")
   IF nValue = m_Capacity THEN EXIT PROPERTY
   this.ResizeBuffer(nValue)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets/sets the code page used to ansi to unicode translations
' ========================================================================================
PRIVATE PROPERTY CUTF16.CodePage () AS UINT
   CUTF16_DP("CUTF16 PROPERTY GET CodePage - " & .WSTR(m_CodePage))
   PROPERTY = m_CodePage
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CUTF16.CodePage (BYVAL nCodePage AS UINT)
   CUTF16_DP("CUTF16 PROPERTY SET CodePage - " & .WSTR(nCodePage))
   m_CodePage = nCodePage
END PROPERTY
' ========================================================================================

' ========================================================================================
' Returns a pointer to the CWSTR buffer.
' ========================================================================================
PRIVATE OPERATOR CUTF16.CAST () AS ANY PTR
   CUTF16_DP("CUTF16 CAST ANY PTR - buffer: " & .WSTR(m_pBuffer))
   OPERATOR = m_pBuffer
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the string data (same as **).
' ========================================================================================
PRIVATE OPERATOR CUTF16.CAST () BYREF AS WSTRING
   CUTF16_DP("CUTF16 CAST BYREF AS WSTRING - buffer: " & .WSTR(m_pBuffer))
   OPERATOR = *m_pBuffer
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the string data (same as **).
' ========================================================================================
PRIVATE FUNCTION CUTF16.wstr () BYREF AS WSTRING
   CUTF16_DP("CUTF16 StrAddr - buffer: " & .WSTR(m_pBuffer))
   RETURN *m_pBuffer
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns the address of the CUTF16 buffer.
' ========================================================================================
PRIVATE FUNCTION CUTF16.vptr () AS WSTRING PTR
   CUTF16_DP("CUTF16 vptr - buffer: " & .WSTR(m_pBuffer))
   RETURN m_pBuffer
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns the address of the CUTF16 buffer.
' Same as vptr for this kind of variable.
' ========================================================================================
PRIVATE FUNCTION CUTF16.sptr () AS WSTRING PTR
   CUTF16_DP("CUTF16 sptr - buffer: " & .WSTR(m_pBuffer))
   RETURN m_pBuffer
END FUNCTION
' ========================================================================================

' =====================================================================================
' Returns the contents of the CWSTR as a WSTRING allocated with CoTaskMemAlloc.
' Free the returned string later with CoTaskMemFree.
' Note: This is useful when we need to pass a pointer to a null terminated wide string to a
' function or method that will release it. If we pass a WSTRING it will GPF.
' If the length of the input string is 0, CoTaskMemAlloc allocates a zero-length item and
' returns a valid pointer to that item. If there is insufficient memory available,
' CoTaskMemAlloc returns NULL.
' =====================================================================================
PRIVATE FUNCTION CUTF16.wchar () AS WSTRING PTR
   CUTF16_DP("CUTF16 wchar")
   DIM pwchar AS WSTRING PTR
   pwchar = CoTaskMemAlloc((m_BufferLen + 1) * 2)
   IF pwchar = NULL THEN RETURN NULL
   IF m_BufferLen THEN wmemmove pwchar, m_pBuffer, m_BufferLen
   IF m_BufferLen = 0 THEN *pwchar = CHR(0)
   RETURN pwchar
END FUNCTION
' =====================================================================================

' ========================================================================================
' Returns the corresponding unicode integer representation of the character at the position
' specified by the nIndex parameter (1 for the first character, 2 for the second, etc.).
' If nIndex is beyond the current length of the string, a 0 is returned.
' ========================================================================================
PRIVATE PROPERTY CUTF16.Char (BYVAL nIndex AS UINT) AS USHORT
   CUTF16_DP("CUTF16 PROPERTY GET Char")
   IF nIndex < 1 OR nIndex > m_BufferLen THEN EXIT PROPERTY
   ' Get the numeric character code at position nIndex
   nIndex -= 1
   PROPERTY = PEEK(USHORT, m_pBuffer + nIndex)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Changes the corresponding unicode integer representation of the character at the position
' specified by the nIndex parameter (1 for the first character, 2 for the second, etc.).
' If nIndex is beyond the current length of the string, nothing is changed.
' ========================================================================================
PRIVATE PROPERTY CUTF16.Char (BYVAL nIndex AS UINT, BYVAL nValue AS USHORT)
   CUTF16_DP("CUTF16 PROPERTY SET Char")
   IF nIndex < 1 OR nIndex > m_BufferLen THEN EXIT PROPERTY
   ' Set the numeric character code at position nIndex (zero based)
   nIndex -= 1
   POKE USHORT, m_pBuffer + nIndex, nValue
END PROPERTY
' ========================================================================================
' ========================================================================================
' Returns the corresponding ASCII or Unicode integer representation of the character at
' the position specified by the nIndex parameter. Like the get Char property, but allowing
' to use the [] syntax, e.g. value = cws[1]. Can't be used to change a value.
' ========================================================================================
PRIVATE OPERATOR CUTF16.[] (BYVAL nIndex AS UINT) AS USHORT
   IF nIndex < 1 OR nIndex > m_BufferLen THEN EXIT OPERATOR
   ' Get the numeric character code at position nIndex
   nIndex -= 1
   OPERATOR = PEEK(USHORT, m_pBuffer + nIndex)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Converts the CUTF16 to UTF8.
' ========================================================================================
PRIVATE PROPERTY CUTF16.Utf8 () AS STRING
   CUTF16_DP("CUTF16 Utf8 GET PROPERTY")
   DIM cbLen AS INTEGER
   IF m_BufferLen = 0 THEN RETURN ""
   DIM buffer AS STRING = STRING(m_BufferLen * 5 + 1, 0)
   PROPERTY = *cast(ZSTRING PTR, WCharToUTF(1, m_pBuffer, m_BufferLen * 2, STRPTR(buffer), @cbLen))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Converts UTF8 to unicode and assigns it to the CUTF16.
' ========================================================================================
PRIVATE PROPERTY CUTF16.Utf8 (BYREF utf8String AS STRING)
   CUTF16_DP("CUTF16 Utf8 SET PROPERTY")
   this.Clear
   this.Add(utf8String, CP_UTF8)
END PROPERTY
' ========================================================================================

' =====================================================================================
' Returns the contents of the CWSTR as a CBSTR.
' Useful to pass a CWSTR to a function that expects a BYVAL IN BSTR parameter.
' =====================================================================================
PRIVATE FUNCTION CUTF16.cbstr () AS CBSTR_
   RETURN SysAllocString(m_pBuffer)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the contents of the CWSTR as a BSTR.
' =====================================================================================
PRIVATE FUNCTION CUTF16.bstr () AS AFX_BSTR
   RETURN SysAllocString(m_pBuffer)
END FUNCTION
' =====================================================================================

' ========================================================================================
' Resizes the internal buffer capacity
' ========================================================================================
PRIVATE FUNCTION CUTF16.ResizeBuffer (BYVAL nChars AS UINT) AS WSTRING PTR
   CUTF16_DP("CUTF16 ResizeBuffer")
   ' // Allocate a buffer of n utf16 characters + 1 for the terminating null
   DIM pNewBuffer AS WSTRING PTR = Allocate((nChars + 1) * 2)
   ' // Copy the old buffer in the new one
   IF nChars < m_BufferLen THEN m_BufferLen = nChars
   IF m_pBuffer THEN
      ' // Copy the old buffer in the new
      wmemmove(pNewBuffer, m_pBuffer, m_BufferLen)
      ' // Mark the end of the string with a double null
      m_pBuffer[m_BufferLen] = 0
      ' // Deallocate the old buffer
      Deallocate m_pBuffer
   END IF
   ' // Update the capacity
   m_Capacity = nChars
   ' // Store the new pointer
   m_pBuffer = pNewBuffer
   RETURN m_pBuffer
END FUNCTION
' ========================================================================================

' ========================================================================================
' Appends the specified number of characters from the specified memory address to the end of the buffer.
' ========================================================================================
PRIVATE FUNCTION CUTF16.AppendBuffer (BYVAL memAddr AS ANY PTR, BYVAL nChars AS UINT) AS BOOLEAN
   CUTF16_DP("CUTF16 AppendBuffer")
   IF memAddr = NULL OR nCHars = 0 THEN RETURN FALSE
   ' // Number of characters + extra space to avoid multiple memory allocations
   DIM nSize AS UINT = m_BufferLen + nChars + m_GrowSize
   ' // If m_GrowSize = -1 THEN double the current capacity
   IF m_GrowSize < 0 THEN nSize = (m_BufferLen + nChars) * 2
   ' // there is not enough capacity, resize the buffer
   IF m_BufferLen + nChars > m_Capacity THEN this.ResizeBuffer(nSize)
   ' // Copy the passed buffer
   IF m_pBuffer = NULL THEN RETURN FALSE
   wmemmove(m_pBuffer + m_BufferLen, memAddr, nChars)
   ' // Update the length of the buffer
   m_BufferLen += nChars
   ' // Mark the end of the string with a double null
   m_pBuffer[m_BufferLen] = 0
   RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' nCount characters are removed starting at the position given by nIndex.
' nIndex = 1 for the first character, 2 for the second, etc.
' Return value: If the function succeeds, it returns TRUE; otherwise, FALSE.
' Remarks: If nCount is bigger that the number of characters available to delete, the
' function deletes all the characters from nIndex to the end of the string.
' ========================================================================================
PRIVATE FUNCTION CUTF16.DelChars (BYVAL nIndex AS UINT, BYVAL nCount AS UINT) AS BOOLEAN
   CUTF16_DP("CUTF16 DelChars")
   IF nIndex < 1 OR nIndex > m_BufferLen OR nCount < 1 THEN RETURN FALSE
   DIM numChars AS UINT = m_BufferLen
   IF nCount > m_BufferLen - nIndex + 1 THEN nCount = m_BufferLen - nIndex + 1
   wmemmove(m_pBuffer + nIndex - 1, m_pBuffer + (nIndex + nCount) - 1, (m_BufferLen - nIndex - nCount + 1))
   m_BufferLen -= nCount
   m_pBuffer[m_BufferLen] = 0
   RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts the specified number of bytes from the specified memory address into the buffer.
' ========================================================================================
PRIVATE FUNCTION CUTF16.InsertBuffer (BYVAL memAddr AS ANY PTR, BYVAL nIndex AS UINT, BYVAL nChars AS UINT) AS BOOLEAN
   CUTF16_DP("CUTF16 InsertBuffer")
   IF nIndex < 1 OR nIndex > m_BufferLen THEN RETURN FALSE
   ' // Determine the size of the new buffer
   IF m_BufferLen + nChars > m_Capacity THEN m_Capacity = m_BufferLen + nChars
   DIM pNewBuffer AS WSTRING PTR = Allocate((m_Capacity + 1) * 2)
   IF m_pBuffer THEN
      ' // Copy the existing data into the new buffer up to nIndex
      nIndex -= 1   ' // Buffer memory is zero-based
      wmemmove(pNewBuffer, m_pBuffer, nIndex)
      DIM nOffset AS UINT = nIndex * 2
      ' // Insert characters
      wmemmove(pNewBuffer + nIndex, memAddr, nChars)
      ' // Copy characters from nIndex
      wmemmove(pNewBuffer + nIndex + nChars, m_pBuffer + nIndex, m_BufferLen - nIndex)
      Deallocate m_pBuffer
   END IF
   m_pBuffer = pNewBuffer
   m_BufferLen += nChars
   m_pBuffer[m_BufferLen] = 0
   RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' The incoming string parameter is inserted in the string starting at the position
' given by nIndex. nIndex = 1 for the first character, 2 For the second, etc.
' If nIndex is beyond the current length of the string + 1, no operation is performed.
' ========================================================================================
PRIVATE FUNCTION CUTF16.Insert (BYREF ansiStr AS STRING, BYVAL nIndex AS UINT, BYVAL nCodePage AS UINT = 0) AS BOOLEAN
   CUTF16_DP("CUTF16 Insert STRING")
   IF nIndex < 1 OR nIndex > m_BufferLen OR .LEN(ansiStr) = 0 THEN RETURN FALSE
   ' // Create the wide string from the incoming ansi string
   DIM dwLen AS UINT, pbuffer AS ANY PTR
   IF nCodePage = 0 THEN nCodePage = m_CodePage
   IF nCodePage = CP_UTF8 THEN
      dwLen = MultiByteToWideChar(CP_UTF8, 0, STRPTR(ansiStr), LEN(ansiStr), NULL, 0)
      IF dwLen = 0 THEN RETURN FALSE
      dwLen *= 2
      pbuffer = Allocate(dwLen)
      dwLen = MultiByteToWideChar(CP_UTF8, 0, STRPTR(ansiStr), LEN(ansiStr), pbuffer, dwLen)
      IF dwLen = 0 THEN RETURN FALSE
   ELSE
      dwLen = .LEN(ansiStr)
      dwLen *= 2
      pbuffer = Allocate(dwLen)
      dwLen = MultiByteToWideChar(m_CodePage, MB_PRECOMPOSED, STRPTR(ansiStr), .LEN(ansiStr), pbuffer, dwLen)
      IF dwLen = 0 THEN RETURN FALSE
   END IF
   ' // Copy the string into the buffer
   DIM bRes AS BOOLEAN = this.InsertBuffer(pbuffer, nIndex, dwLen)
   Deallocate(pbuffer)
   RETURN bRes
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CUTF16.Insert (BYVAL pwszStr AS WSTRING PTR, BYVAL nIndex AS UINT) AS BOOLEAN
   CUTF16_DP("CUTF16 Insert WSTRING")
   IF nIndex < 1 OR nIndex > m_BufferLen THEN RETURN FALSE
   RETURN this.InsertBuffer(pwszStr, nIndex, .LEN(*pwszStr))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CUTF16.Insert (BYREF cws AS CUTF16, BYVAL nIndex AS UINT) AS BOOLEAN
   CUTF16_DP("CUTF16 Insert CWSTR")
   IF (nIndex < 1) OR nIndex > m_BufferLen THEN RETURN FALSE
   IF cws.m_BufferLen = 0 THEN RETURN FALSE
   RETURN this.InsertBuffer(cws.m_pBuffer, nIndex, cws.m_BufferLen)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CUTF16.Insert (BYREF cws AS CWSTR, BYVAL nIndex AS UINT) AS BOOLEAN
   CUTF16_DP("CUTF16 Insert CWSTR")
   IF (nIndex < 1) OR nIndex > m_BufferLen THEN RETURN FALSE
   RETURN this.InsertBuffer(cws.m_pBuffer, nIndex, cws.m_BufferLen \ 2)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CUTF16.Insert (BYREF cbs AS CBSTR_, BYVAL nIndex AS LONG) AS BOOLEAN
   CUTF16_DP("CUTF16 Insert CBSTR")
   IF (nIndex < 1) OR nIndex > m_BufferLen THEN RETURN 0
   RETURN this.InsertBuffer(cbs.m_bstr, nIndex, SysStringLen(cbs.m_bstr))
END FUNCTION
' ========================================================================================

' ========================================================================================
' All data in the class object is erased. Actually, we only set the buffer length to zero,
' indicating no string in the buffer. The allocated memory for the buffer is deallocated
' when the class is destroyed.
' ========================================================================================
PRIVATE SUB CUTF16.Clear
   CUTF16_DP("CUTF16 Clear")
   m_BufferLen = 0 : m_pBuffer[m_BufferLen] = 0
END SUB
' ========================================================================================

' ========================================================================================
' The string parameter is appended to the string held in the class. If the internal string
' buffer overflows, the class will automatically extend it to an appropriate size.
' ========================================================================================
PRIVATE FUNCTION CUTF16.Add (BYVAL pwszStr AS WSTRING PTR) AS BOOLEAN
   CUTF16_DP("CUTF16 Add - WSTRING PTR")
   RETURN this.AppendBuffer(pwszStr, .LEN(*pwszStr))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CUTF16.Add (BYREF ansiStr AS STRING, BYVAL nCodePage AS UINT = 0) AS BOOLEAN
   CUTF16_DP("CUTF16 Add - STRING")
   IF LEN(ansiStr) = 0 THEN RETURN FALSE
   ' // Create the wide string from the incoming ansi string
   DIM dwLen AS UINT, pbuffer AS ANY PTR
   IF nCodePage = 0 THEN nCodePage = m_CodePage
   IF nCodePage = CP_UTF8 THEN
      dwLen = MultiByteToWideChar(CP_UTF8, 0, STRPTR(ansiStr), LEN(ansiStr), NULL, 0)
      IF dwLen = 0 THEN RETURN FALSE
      dwLen *= 2
      pbuffer = Allocate(dwLen)
      dwLen = MultiByteToWideChar(CP_UTF8, 0, STRPTR(ansiStr), LEN(ansiStr), pbuffer, dwLen)
      IF dwLen = 0 THEN RETURN FALSE
   ELSE
      dwLen = .LEN(ansiStr)
      dwLen *= 2
      pbuffer = Allocate(dwLen)
      dwLen = MultiByteToWideChar(m_CodePage, MB_PRECOMPOSED, STRPTR(ansiStr), .LEN(ansiStr), pbuffer, dwLen)
      IF dwLen = 0 THEN RETURN FALSE
   END IF
   ' // Copy the string into the buffer
   DIM bRes AS BOOLEAN = this.AppendBuffer(pbuffer, dwLen)
   Deallocate(pbuffer)
   RETURN bRes
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CUTF16.Add (BYREF cws AS CUTF16) AS BOOLEAN
   CUTF16_DP("CUTF16 Add - CUTF16")
   RETURN this.AppendBuffer(cws.m_pBuffer, cws.m_BufferLen)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CUTF16.Add (BYREF cws AS CWSTR) AS BOOLEAN
   CUTF16_DP("CUTF16 Add CWSTR - buffer = " & .WSTR(cws.m_pBuffer) & " - LEN = " & .WSTR(LEN(*cast(WSTRING PTR, @cws))))
   RETURN this.AppendBuffer(cws.m_pBuffer, cws.m_BufferLen \ 2)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CUTF16.Add (BYREF cbs AS CBSTR_) AS BOOLEAN
   CUTF16_DP("CUTF16 Add CBSTR")
   RETURN this.AppendBuffer(cbs.m_bstr, SysStringLen(cbs.m_bstr))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Appends a string to the CUTF16
' ========================================================================================
PRIVATE OPERATOR CUTF16.+= (BYREF ansiStr AS STRING)
   CUTF16_DP("CUTF16 OPERATOR += STRING")
   this.Add(ansiStr)
END OPERATOR
' ========================================================================================
' ========================================================================================
' Appends a CWSTR to the CUTF16
' ========================================================================================
PRIVATE OPERATOR CUTF16.+= (BYREF cws AS CWSTR)
   CUTF16_DP("CUTF16 OPERATOR += CWSTR")
   this.Add(cws)
END OPERATOR
' ========================================================================================
' ========================================================================================
' Appends a CBSTR to the CUTF16.
' ========================================================================================
PRIVATE OPERATOR CUTF16.+= (BYREF cbs AS CBSTR_)
   CUTF16_DP("CUTF16 OPERATOR += CBSTR")
   this.Add(cbs.m_bstr)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CUTF16.+= (BYVAL n AS LONGINT)
   CUTF16_DP("CUTF16 OPERATOR += LONGINT")
   DIM wsz AS WSTRING * 260 = .WSTR(n) : this.Add(wsz)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CUTF16.+= (BYVAL n AS DOUBLE)
   CUTF16_DP("CUTF16 OPERATOR += DOUBLE")
   DIM wsz AS WSTRING * 260 = .WSTR(n) : this.Add(wsz)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Appends a WSTRING to the CWSTR
' ========================================================================================
PRIVATE OPERATOR CUTF16.&= (BYREF wszStr AS WSTRING)
   CUTF16_DP("CUTF16 OPERATOR &= WSTRING")
   this.Add(wszStr)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Appends a string to the CWSTR
' ========================================================================================
PRIVATE OPERATOR CUTF16.&= (BYREF ansiStr AS STRING)
   CUTF16_DP("CUTF16 OPERATOR &= STRING")
   this.Add(ansiStr)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Appends a CWSTR to the CWSTR
' ========================================================================================
PRIVATE OPERATOR CUTF16.&= (BYREF cws AS CWSTR)
   CUTF16_DP("CUTF16 OPERATOR &= CWSTR")
   this.Add(cws)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Appends a CBSTR to the CWSTR.
' ========================================================================================
PRIVATE OPERATOR CUTF16.&= (BYREF cbs AS CBSTR_)
   CUTF16_DP("CUTF16 OPERATOR &= CBSTR")
   this.Add(cbs.m_bstr)
END OPERATOR
' ========================================================================================

' ========================================================================================
PRIVATE OPERATOR CUTF16.&= (BYVAL n AS LONGINT)
   CUTF16_DP("CUTF16 OPERATOR &= LONGINT")
   DIM wsz AS WSTRING * 260 = .WSTR(n) : this.Add(wsz)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CUTF16.&= (BYVAL n AS DOUBLE)
   CUTF16_DP("CUTF16 OPERATOR &= DOUBLE")
   DIM wsz AS WSTRING * 260 = .WSTR(n) : this.Add(wsz)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns new text to the CUTF16
' ========================================================================================
PRIVATE OPERATOR CUTF16.Let (BYREF wszStr AS CONST WSTRING)
   CUTF16_DP("CUTF16 Add - WSTRING")
   this.Clear : this.Add(wszStr)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CUTF16.Let (BYREF pwszStr AS WSTRING PTR)
   CUTF16_DP("CUTF16 Add - WSTRING PTR")
   this.Clear : this.Add(*pwszStr)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CUTF16.Let (BYREF ansiStr AS STRING)
   CUTF16_DP("CWSTR LET STRING")
   this.Clear : this.Add(ansiStr)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CUTF16.Let (BYREF cws AS CUTF16)
   CUTF16_DP("CUTF16 LET CUTF16")
   IF m_pBuffer = cws.m_pBuffer THEN EXIT OPERATOR   ' // Ignore cws = cws
   this.Clear : this.Add(cws)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CUTF16.Let (BYREF cws AS CWSTR)
   CUTF16_DP("CUTF16 LET CWSTR")
   IF m_pBuffer = cws.m_pBuffer THEN EXIT OPERATOR   ' // Ignore cws = cws
   this.Clear : this.Add(cws)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CUTF16.Let (BYREF cbs AS CBSTR_)
   CUTF16_DP("CUTF16 LET CBSTR")
   this.Clear : this.AppendBuffer(cbs.m_bstr, SysStringLen(cbs.m_bstr))
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CUTF16.Let (BYVAL n AS LONGINT)
   CUTF16_DP("CUTF16 OPERATOR Let LONGINT")
   this.Clear : DIM wsz AS WSTRING * 260 = .WSTR(n) : this.Add(wsz)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CUTF16.Let (BYVAL n AS DOUBLE)
   CUTF16_DP("CUTF16 OPERATOR Let DOUBLE")
   this.Clear : DIM wsz AS WSTRING * 260 = .WSTR(n) : this.Add(wsz)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns the leftmost substring of the string.
' ========================================================================================
PRIVATE FUNCTION CUTF16.LeftChars (BYVAL nChars AS LONG) AS CUTF16
   CUTF16_DP("CUTF16 LeftChars")
   RETURN LEFT(*m_pBuffer, nChars)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the rightmost substring of the string.
' ========================================================================================
PRIVATE FUNCTION CUTF16.RightChars (BYVAL nChars AS LONG) AS CUTF16
   CUTF16_DP("CUTF16 LeftChars")
   RETURN RIGHT(*m_pBuffer, nChars)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a substring of the string.
' ========================================================================================
PRIVATE FUNCTION CUTF16.MidChars (BYVAL nStart AS LONG, BYVAL nChars AS LONG = 0) AS CUTF16
   CUTF16_DP("CUTF16 MidChars")
   IF nChars = 0 THEN RETURN MID(*m_pBuffer, nStart)
   RETURN MID(*m_pBuffer, nStart, nChars)
END FUNCTION
' ========================================================================================

' =====================================================================================
' Converts the string to a 32bit integer
' =====================================================================================
PRIVATE FUNCTION CUTF16.ValLong () AS LONG
   RETURN .ValInt(*m_pBuffer)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CUTF16.ValInt () AS LONG
   RETURN .ValInt(*m_pBuffer)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts the string to an unsigned 32bit integer
' =====================================================================================
PRIVATE FUNCTION CUTF16.ValULong () AS ULONG
   RETURN .ValUInt(*m_pBuffer)
END FUNCTION
' =====================================================================================
PRIVATE FUNCTION CUTF16.ValUInt () AS ULONG
   RETURN .ValUInt(*m_pBuffer)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts the string to a 64bit integer
' =====================================================================================
PRIVATE FUNCTION CUTF16.ValLongInt () AS LONGINT
   RETURN .ValLng(*m_pBuffer)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts the string to an unsigned 64bit integer
' =====================================================================================
PRIVATE FUNCTION CUTF16.ValULongInt () AS ULONGINT
   RETURN .ValULng(*m_pBuffer)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts the string to a floating point number (DOUBLE)
' =====================================================================================
PRIVATE FUNCTION CUTF16.ValDouble () AS DOUBLE
   RETURN .VAL(*m_pBuffer)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CUTF16.Value () AS DOUBLE
   RETURN .VAL(*m_pBuffer)
END FUNCTION
' =====================================================================================

END NAMESPACE

' ########################################################################################
'                               *** GLOBAL OPERATORS ***
' ########################################################################################

' // Outside a namespace because they are global
using Afx

' ========================================================================================
' Returns the length, in characters, of the CUTF16.
' ========================================================================================
PRIVATE OPERATOR LEN (BYREF cws AS CUTF16) AS UINT
   OPERATOR = cws.m_BufferLen
END OPERATOR
' ========================================================================================

' ========================================================================================
' One * returns the address of the CWSTR buffer.
' Two ** deferences the string data.
' Needed because LEFT and RIGHT (cws) fail with an ambiguous call error.
' We have to use **cws (notice the double indirection) with these functions.
' ========================================================================================
PRIVATE OPERATOR * (BYREF cws AS CUTF16) AS WSTRING PTR
   OPERATOR = cws.m_pBuffer
END OPERATOR
' ========================================================================================

' ========================================================================================
PRIVATE OPERATOR & (BYREF cws1 AS CUTF16, BYREF cws2 AS CUTF16) AS CUTF16
   DIM cwsRes AS CUTF16 = cws1
   cwsRes.Add(cws2)
   OPERATOR = cwsRes
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Left OVERLOAD (BYREF cws AS CUTF16, BYVAL nChars AS INTEGER) AS CUTF16
   RETURN LEFT(*cws.m_pBuffer, nChars)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Right OVERLOAD (BYREF cws AS CUTF16, BYVAL nChars AS INTEGER) AS CUTF16
   RETURN RIGHT(*cws.m_pBuffer, nChars)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Val OVERLOAD (BYREF cws AS CUTF16) AS DOUBLE
   RETURN .VAL(*cws.m_pBuffer)
END FUNCTION
' ========================================================================================
