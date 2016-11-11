' ########################################################################################
' Microsoft Windows
' Implements a dynamic data type for null terminated unicode strings, offering the ability
' to concatenate many string sections at a very high level of performance.
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2016 Paul Squires & José Roca, with the collaboration of Marc Pons.
' Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma ONCE
#INCLUDE ONCE "windows.bi"
#INCLUDE ONCE "win/ole2.bi"
#INCLUDE ONCE "/crt/string.bi"

NAMESPACE Afx

' ========================================================================================
' Macro for debug
' To allow debugging, define _CWSTR_DEBUG_ 1 in your application before including this file.
' ========================================================================================
#ifndef _CWSTR_DEBUG_
   #define _CWSTR_DEBUG_ 0
#ENDIF
#ifndef _CWSTR_DP_
   #define _CWSTR_DP_ 1
   #MACRO CWSTR_DP(st)
      #IF (_CWSTR_DEBUG_ = 1)
         OutputDebugStringW(st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

' ########################################################################################
'                                  *** CWSTR CLASS ***
' ########################################################################################
TYPE CWSTR

   Private:
      m_Capacity AS LONG            ' The total size of the buffer
      m_GrowSize AS LONG = 260 * 2  ' How much to grow the buffer by when required
      m_CodePage AS UINT            ' Unicode code page

   Public:
      m_pBuffer AS UBYTE PTR   ' Pointer to the buffer
      m_BufferLen AS LONG      ' Length in bytes of the current string in the buffer
      DECLARE CONSTRUCTOR
      DECLARE CONSTRUCTOR (BYVAL nCapacity AS UINT)
      DECLARE CONSTRUCTOR (BYVAL pwszStr AS WSTRING PTR, BYVAL fAttach AS BOOLEAN = FALSE)
      DECLARE CONSTRUCTOR (BYREF ansiStr AS STRING, BYVAL nCodePage AS UINT = 0)
      DECLARE CONSTRUCTOR (BYREF cws AS CWSTR)
      DECLARE DESTRUCTOR
      DECLARE SUB ResizeBuffer (BYVAL nValue AS DWORD)
      DECLARE FUNCTION AppendBuffer (BYVAL addrMemory AS ANY PTR, BYVAL nNumBytes AS LONG) AS BOOLEAN
      DECLARE FUNCTION InsertBuffer (BYVAL addrMemory AS ANY PTR, BYVAL nIndex AS LONG, BYVAL nNumBytes AS LONG) AS BOOLEAN
      DECLARE PROPERTY Capacity () AS LONG
      DECLARE PROPERTY Capacity (nValue AS LONG)
      DECLARE PROPERTY CodePage () AS UINT
      DECLARE PROPERTY CodePage (BYVAL nCodePage AS UINT)
      DECLARE SUB Clear
      DECLARE SUB Add (BYREF cws AS CWSTR)
      DECLARE SUB Add (BYVAL pwszStr AS WSTRING PTR)
      DECLARE SUB Add (BYREF ansiStr AS STRING, BYVAL nCodePage AS UINT = 0)
      DECLARE PROPERTY Char(BYVAL nIndex AS LONG) AS USHORT
      DECLARE PROPERTY Char(BYVAL nIndex AS LONG, BYVAL nValue AS USHORT)
      DECLARE OPERATOR [] (BYVAL nIndex AS LONG) AS USHORT
      DECLARE FUNCTION DelChars (BYVAL nIndex AS LONG, BYVAL nCount AS LONG) AS BOOLEAN
      DECLARE FUNCTION Insert (BYVAL pwszStr AS WSTRING PTR, BYVAL nIndex AS LONG) AS BOOLEAN
      DECLARE FUNCTION Insert (BYREF ansiStr AS STRING, BYVAL nIndex AS LONG, BYVAL nCodePage AS UINT = 0) AS BOOLEAN
      DECLARE OPERATOR CAST () BYREF AS WSTRING
      DECLARE OPERATOR CAST () AS ANY PTR
      DECLARE OPERATOR LET (BYREF ansiStr AS STRING)
      DECLARE OPERATOR LET (BYREF wszStr AS CONST WSTRING)
      DECLARE OPERATOR LET (BYREF pwszStr AS WSTRING PTR)
      DECLARE OPERATOR LET (BYREF cws AS CWSTR)
      DECLARE OPERATOR @ () AS WSTRING PTR
      DECLARE FUNCTION pthis () AS ANY PTR
      DECLARE OPERATOR += (BYREF wszStr AS WSTRING)
      DECLARE OPERATOR += (BYREF cws AS CWSTR)
      DECLARE OPERATOR += (BYREF ansiStr AS STRING)
      DECLARE OPERATOR &= (BYREF wszStr AS WSTRING)
      DECLARE OPERATOR &= (BYREF cws AS CWSTR)
      DECLARE OPERATOR &= (BYREF ansiStr AS STRING)

END TYPE
' ########################################################################################

' ========================================================================================
' CWSTR constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CWSTR
   CWSTR_DP("+++BEGIN- CWSTR CONSTRUCTOR Default")
   this.ResizeBuffer(m_GrowSize)   ' Create the initial buffer
   CWSTR_DP("END CWSTR CONSTRUCTOR Default - " & WSTR(m_pBuffer))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CWSTR (BYVAL nCapacity AS UINT)
   CWSTR_DP("+++BEGIN- CWSTR CONSTRUCTOR Capacity: " & WSTR(nCapacity))
   this.ResizeBuffer(nCapacity)   ' Create the initial buffer
   CWSTR_DP("-END- CWSTR CONSTRUCTOR Capacity - " & WSTR(m_pBuffer))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CWstr (BYVAL pwszStr AS WSTRING PTR, BYVAL fAttach AS BOOLEAN = FALSE)
   CWSTR_DP("+++BEGIN- CWSTR CONSTRUCTOR WSTRING - " & WSTR(pwszStr))
   IF pwszStr = NULL THEN
      ' Create the initial buffer
      this.ResizeBuffer(m_GrowSize)
   ELSEIF .LEN(*pwszStr) THEN
      IF fAttach THEN
         CWSTR_DP("+++CWSTR CONSTRUCTOR WSTRING - Attach")
         ' Attach the passed WSTRING handle
         m_pBuffer = cast(UBYTE PTR, pwszStr)
         m_BufferLen = .LEN(*pwszStr) * 2
         m_Capacity = m_BufferLen
      ELSE
         ' Add the contents of the passed WSTRING
         this.Add(pwszStr)
      END IF
   ELSE
      ' Create the initial buffer
      this.ResizeBuffer(m_GrowSize)
   END IF
   CWSTR_DP("-END- CWSTR CONSTRUCTOR WSTRING - " & WSTR(m_pBuffer))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CWstr (BYREF ansiStr AS STRING, BYVAL nCodePage AS UINT = 0)
   CWSTR_DP("+++BEGIN- CWSTR CONSTRUCTOR STRING - " & WSTR(VARPTR(ansiStr)))
   m_CodePage = nCodePage             ' Store the code page
   IF .LEN(ansiStr) THEN
      this.Add(ansiStr, nCodePage)    ' Add the passed ansi string
   ELSE
      this.ResizeBuffer(m_GrowSize)   ' Create the initial buffer
   END IF
   CWSTR_DP("-END- CWSTR CONSTRUCTOR STRING - " & WSTR(m_pBuffer))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CWstr (BYREF cws AS CWSTR)
   CWSTR_DP("+++BEGIN- CWSTR CONSTRUCTOR CWSTR - " & WSTR(@cws))
   IF cws.m_BufferLen THEN
      this.Add(cws)                   ' Add the passed CWSTR
   ELSE
      this.ResizeBuffer(m_GrowSize)   ' Create the initial buffer
   END IF
   CWSTR_DP("-END- CWSTR CONSTRUCTOR CWSTR - " & WSTR(m_pBuffer))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR CWstr
   CWSTR_DP("***CWSTR DESTRUCTOR - buffer: " & WSTR(m_pBuffer))
   IF m_pBuffer THEN Deallocate(m_pBuffer)
   m_pBuffer = NULL
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the address of the CWSTR buffer.
' ========================================================================================
PRIVATE OPERATOR CWstr.@ () AS WSTRING PTR
   CWSTR_DP("CWSTR OPERATOR @ - buffer: " & WSTR(m_pBuffer))
   OPERATOR = cast(WSTRING PTR, m_pBuffer)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns the address of the CWSTR class.
' ========================================================================================
PRIVATE FUNCTION CWstr.pthis () AS ANY PTR
   CWSTR_DP("CWSTR FUNCTION pthis - address: " & WSTR(@this))
   RETURN @this
END FUNCTION
' ========================================================================================

' ========================================================================================
' One * returns the value of the CWSTR pointer.
' Two ** returns the adress of the start of the string data.
' Needed because LEFT and RIGHT (cws) fail with an ambiguous call error.
' We have to use **cws (notice the double indirection) with these functions.
' ========================================================================================
PRIVATE OPERATOR * (BYREF cws AS CWSTR) AS WSTRING PTR
   CWSTR_DP("CWSTR OPERATOR * buffer: " & WSTR(cws.m_pBuffer))
   OPERATOR = cast(WSTRING PTR, cws.m_pBuffer)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns the length of the CWSTR.
' ========================================================================================
PRIVATE OPERATOR LEN (BYREF cws AS CWSTR) AS LONG
   CWSTR_DP("CWSTR OPERATOR LEN - len: " & WSTR(.LEN(**cws)))
   OPERATOR = .LEN(**cws)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns a pointer to the CWSTR buffer.
' ========================================================================================
PRIVATE OPERATOR CWstr.CAST () BYREF AS WSTRING
   CWSTR_DP("CWSTR CAST BYREF AS WSTRING - buffer: " & WSTR(m_pBuffer))
   OPERATOR = *cast(WSTRING PTR, m_pBuffer)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CWstr.CAST () AS ANY PTR
   CWSTR_DP("CWSTR CAST ANY PTR - buffer: " & WSTR(m_pBuffer))
   OPERATOR = cast(ANY PTR, m_pBuffer)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Gets/sets the code page used to ansi to unicode translations
' ========================================================================================
PRIVATE PROPERTY CWStr.CodePage () AS UINT
   CWSTR_DP("CWSTR PROPERTY GET CodePage - " & WSTR(m_CodePage))
   PROPERTY = m_CodePage
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWStr.CodePage (BYVAL nCodePage AS UINT)
   CWSTR_DP("CWSTR PROPERTY SET CodePage - " & WSTR(nCodePage))
   m_CodePage = nCodePage
END PROPERTY
' ========================================================================================

' ========================================================================================
' Assigns new text to the CWSTR.
' ========================================================================================
PRIVATE OPERATOR CWstr.Let (BYREF wszStr AS CONST WSTRING)
   CWSTR_DP("CWSTR LET WSTRING")
   this.Clear
   this.Add(wszStr)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CWstr.Let (BYREF ansiStr AS STRING)
   CWSTR_DP("CWSTR LET STRING")
   this.Clear
   this.Add(ansiStr)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CWstr.Let (BYREF pwszStr AS WSTRING PTR)
   CWSTR_DP("CWSTR LET WSTRING PTR = " & WSTR(pwszStr))
   IF pwszStr = NULL THEN EXIT OPERATOR
   this.Clear
   this.Add(*pwszStr)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CWstr.Let (BYREF cws AS CWSTR)
   CWSTR_DP("CWSTR LET CWSTR - m_pBuffer = " & WSTR(m_pBuffer) & " - IN buffer = " & WSTR(cws.m_pBuffer))
   IF m_pBuffer = cws.m_pBuffer THEN EXIT OPERATOR   ' // Ignore cws = cws
   this.Clear
   this.Add(cws)
END OPERATOR
' ========================================================================================

' ========================================================================================
' ResizeBuffer
' Increase the size of the internal buffer capacity
' ========================================================================================
PRIVATE SUB CWstr.ResizeBuffer (BYVAL nValue AS DWORD)
   CWSTR_DP("CWSTR ResizeBuffer - Value = " & WSTR(nValue))
   ' Increase the size of the existing buffer by creating a new buffer copying
   ' the existing data into it and then finally deleting the original buffer.
   DIM pNewBuffer AS UBYTE PTR = Allocate(nValue + 1 * SIZEOF(UBYTE) * 2)
   CWSTR_DP("CWSTR ResizeBuffer - pNewBuffer = " & WSTR(pNewBuffer) & " - old buffer = " & (WSTR(m_pBuffer)))
   IF m_pBuffer THEN
      memcpy(pNewBuffer, m_pBuffer, m_BufferLen)
      Deallocate m_pBuffer
   END IF
   m_pBuffer = pNewBuffer
   m_Capacity = nValue
   ' Mark the end of the string with a double null
   m_pBuffer[m_BufferLen] = 0
   m_pBuffer[m_BufferLen + 1] = 0
END SUB
' ========================================================================================

' ========================================================================================
' Append the number of bytes from the specified memory address to the end of the buffer.
' ========================================================================================
PRIVATE FUNCTION CWstr.AppendBuffer (BYVAL addrMemory AS ANY PTR, BYVAL nNumBytes AS LONG) AS BOOLEAN
   CWSTR_DP("CWSTR AppendBuffer " & WSTR(m_BufferLen) & " " & WSTR(nNumBytes))
   IF (m_BufferLen + nNumBytes) > m_Capacity THEN this.ResizeBuffer((m_BufferLen + nNumBytes) * 2)
   IF m_pBuffer = NULL THEN RETURN FALSE
   memcpy(m_pBuffer + m_BufferLen, addrMemory, nNumBytes)
   m_BufferLen += nNumBytes
   ' Mark the end of the string with a double null
   m_pBuffer[m_BufferLen] = 0
   m_pBuffer[m_BufferLen + 1] = 0
   RETURN TRUE
   CWSTR_DP("--END - CWSTR AppendBuffer " & WSTR(m_BufferLen))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The string parameter is appended to the string held in the class. If the internal string
' buffer overflows, the class will automatically extend it to an appropriate size.
' ========================================================================================
PRIVATE SUB CWstr.Add (BYREF cws AS CWSTR)
   CWSTR_DP("CWSTR Add CWSTR - buffer = " & WSTR(cws.m_pBuffer) & " - LEN = " & WSTR(LEN(*cast(WSTRING PTR, @cws))))
   ' Incoming string is already in wide format, simply copy it to the buffer.
   DIM AS LONG nLenString = LEN(*cast(WSTRING PTR, @cws))
   IF nLenString = 0 THEN RETURN
   ' Copy the string into the buffer and update the length
   this.AppendBuffer(cast(ANY PTR, cws), nLenString * 2)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CWstr.Add (BYVAL pwszStr AS WSTRING PTR)
   CWSTR_DP("CWSTR Add WSTRING")
   IF pwszStr = NULL THEN RETURN
   ' Incoming string is already in wide format
   DIM AS LONG nLenString = .LEN(*pwszStr)
   IF nLenString = 0 THEN RETURN
   ' Copy the string into the buffer and update the length
   this.AppendBuffer(cast(ANY PTR, pwszStr), nLenString * 2)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CWstr.Add (BYREF ansiStr AS STRING, BYVAL nCodePage AS UINT = 0)
   CWSTR_DP("CWSTR Add STRING Code page = " & WSTR (nCodePage))
   IF LEN(ansiStr) = 0 THEN RETURN
   ' Create the wide string from the incoming ansi string
   DIM dwLen AS DWORD, pbuffer AS ANY PTR
   IF nCodePage = 0 THEN nCodePage = m_CodePage
   IF nCodePage = CP_UTF8 THEN
      dwLen = MultiByteToWideChar(CP_UTF8, 0, STRPTR(ansiStr), LEN(ansiStr), NULL, 0)
      IF dwLen = 0 THEN RETURN
      dwLen *= 2
      pbuffer = Allocate(dwLen)
      MultiByteToWideChar(CP_UTF8, 0, STRPTR(ansiStr), LEN(ansiStr), pbuffer, dwLen)
   ELSE
      dwLen = .LEN(ansiStr)
      dwLen *= 2
      pbuffer = Allocate(dwLen)
      MultiByteToWideChar(m_CodePage, MB_PRECOMPOSED, STRPTR(ansiStr), .LEN(ansiStr), pbuffer, dwLen)
   END IF
   IF pbuffer THEN
      ' Copy the string into the buffer and update the length
      this.AppendBuffer(pbuffer, dwLen)
      Deallocate(pbuffer)
   END IF
END SUB
' ========================================================================================

' ========================================================================================
' Appends a string to the CWSTR
' ========================================================================================
PRIVATE OPERATOR CWStr.+= (BYREF wszStr AS WSTRING)
   CWSTR_DP("CWSTR OPERATOR += WSTRING")
   this.Add(wszStr)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Appends a string to the CWSTR
' ========================================================================================
PRIVATE OPERATOR CWStr.+= (BYREF ansiStr AS STRING)
   CWSTR_DP("CWSTR OPERATOR += STRING")
   this.Add(ansiStr)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Appends a CWSTR to the CWSTR
' ========================================================================================
PRIVATE OPERATOR CWStr.+= (BYREF cws AS CWStr)
   CWSTR_DP("CWSTR OPERATOR += CWSTR")
   this.Add(cws)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Appends a WSTRING to the CWSTR
' ========================================================================================
PRIVATE OPERATOR CWStr.&= (BYREF wszStr AS WSTRING)
   CWSTR_DP("CWSTR OPERATOR &= WSTRING")
   this.Add(wszStr)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Appends a string to the CWSTR
' ========================================================================================
PRIVATE OPERATOR CWStr.&= (BYREF ansiStr AS STRING)
   CWSTR_DP("CWSTR OPERATOR &= STRING")
   this.Add(ansiStr)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Appends a CWSTR to the CWSTR
' ========================================================================================
PRIVATE OPERATOR CWStr.&= (BYREF cws AS CWStr)
   CWSTR_DP("CWSTR OPERATOR &= CWSTR")
   this.Add(cws)
END OPERATOR
' ========================================================================================

' ========================================================================================
' The size of the internal string buffer is retrieved and returned to the caller. The size
' is the number of characters which can be stored without further expansion.
' ========================================================================================
PRIVATE PROPERTY CWstr.Capacity() AS LONG
   CWSTR_DP("CWSTR PROPERTY GET Capacity")
   PROPERTY = m_Capacity
END PROPERTY
' ========================================================================================
' ========================================================================================
' The internal string buffer is expanded to the specified number of characters. If the new
' capacity is smaller or equal to the current capacity, no operation is performed.
' ========================================================================================
PRIVATE PROPERTY CWstr.Capacity (BYVAL nValue AS LONG)
   CWSTR_DP("CWSTR PROPERTY SET Capacity")
   IF nValue <= m_Capacity THEN EXIT PROPERTY
   this.ResizeBuffer(nValue)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Returns the corresponding unicode integer representation of the character at the position
' specified by the nIndex parameter (1 for the first character, 2 for the second, etc.).
' If nIndex is beyond the current length of the string, a 0 is returned.
' ========================================================================================
PRIVATE PROPERTY CWstr.Char (BYVAL nIndex AS LONG) AS USHORT
   CWSTR_DP("CWSTR PROPERTY GET Char")
   IF nIndex < 1 OR nIndex > m_BufferLen \ 2 THEN EXIT PROPERTY
   ' Get the numeric character code at position nIndex
   nIndex -= 1
   PROPERTY = PEEK(USHORT, m_pBuffer + (nIndex * 2))
END PROPERTY
' ========================================================================================
' ========================================================================================
' Changes the corresponding unicode integer representation of the character at the position
' specified by the nIndex parameter (1 for the first character, 2 for the second, etc.).
' If nIndex is beyond the current length of the string, nothing is changed.
' ========================================================================================
PRIVATE PROPERTY CWstr.Char (BYVAL nIndex AS LONG, BYVAL nValue AS USHORT)
   CWSTR_DP("CWSTR PROPERTY SET Char")
   IF nIndex < 1 OR nIndex > m_BufferLen \ 2 THEN EXIT PROPERTY
   ' Set the numeric character code at position nIndex (zero based)
   nIndex -= 1
   POKE USHORT, m_pBuffer + (nIndex * 2), nValue
END PROPERTY
' ========================================================================================
' ========================================================================================
' Returns the corresponding ASCII or Unicode integer representation of the character at
' the position specified by the nIndex parameter. Like the get Char property, but allowing
' to use the [] syntax, e.g. value = cws[1]. Can't be used to change a value.
' ========================================================================================
PRIVATE OPERATOR CWStr.[] (BYVAL nIndex AS LONG) AS USHORT
   IF nIndex < 1 OR nIndex > m_BufferLen \ 2 THEN EXIT OPERATOR
   ' Get the numeric character code at position nIndex
   nIndex -= 1
   OPERATOR = PEEK(USHORT, m_pBuffer + (nIndex * 2))
END OPERATOR
' ========================================================================================

' ========================================================================================
' All data in the class object is erased. Actually, we only set the buffer length to zero,
' indicating no string in the buffer. The allocated memory for the buffer is deallocated
' when the class is destroyed.
' ========================================================================================
PRIVATE SUB CWstr.Clear
   CWSTR_DP("CWSTR Clear")
   m_BufferLen = 0
   ' Mark the end of the string with a double null
   m_pBuffer[m_BufferLen] = 0
   m_pBuffer[m_BufferLen + 1] = 0
END SUB
' ========================================================================================

' ========================================================================================
' nCount characters are removed starting at the position given by nIndex.
' nIndex = 1 for the first character, 2 for the second, etc.
' Return value: If the function succeeds, it returns TRUE; otherwise, FALSE.
' Remarks: If nCount is bigger that the number of characters available to delete, the
' function deletes all the characters from nIndex to the end of the string.
' ========================================================================================
PRIVATE FUNCTION CWstr.DelChars (BYVAL nIndex AS LONG, BYVAL nCount AS LONG) AS BOOLEAN
   CWSTR_DP("CWSTR DelChars")
   IF nIndex < 1 OR nIndex > m_BufferLen \ 2 OR nCount < 1 THEN RETURN FALSE
   DIM numChars AS LONG = m_BufferLen \ 2
   IF nCount > numChars - nIndex + 1 THEN nCount = numChars - nIndex + 1
   DIM AS LONG nOffset = (nIndex - 1 + nCount) * 2
   memcpy(m_pBuffer + (nOffset - nCount * 2), m_pBuffer + nOffset, m_BufferLen - (nCount * 2))
   m_BufferLen -= (nCount * 2)
   ' Mark the end of the string with a double null
   m_pBuffer[m_BufferLen] = 0
   m_pBuffer[m_BufferLen + 1] = 0
   RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Insert the number of bytes from the specified memory address into the buffer.
' ========================================================================================
PRIVATE FUNCTION CWstr.InsertBuffer (BYVAL addrMemory AS ANY PTR, BYVAL nIndex AS LONG, BYVAL nNumBytes AS LONG) AS BOOLEAN
   CWSTR_DP("CWSTR InsertBuffer")
   IF nIndex < 1 OR nIndex > m_BufferLen \ 2 THEN RETURN FALSE
   ' Determine the size of the new buffer
   IF m_BufferLen + nNumBytes > m_Capacity THEN m_Capacity = m_BufferLen + nNumBytes
   DIM pNewBuffer AS UBYTE PTR = Allocate(m_Capacity * SIZEOF(UBYTE) * 2)
   IF m_pBuffer THEN
      nIndex -= 1
      ' Copy the existing data into the new buffer
      memcpy(pNewBuffer, m_pBuffer, nIndex * 2)
      DIM nOffset AS LONG = nIndex * 2
      memcpy(pNewBuffer + nOffset, addrMemory, nNumBytes)
      nOffset = nOffset + nNumBytes
      memcpy(pNewBuffer + nOffset, m_pBuffer + (nIndex * 2), (m_BufferLen - nIndex) * 2)
      Deallocate m_pBuffer
   END IF
   m_pBuffer = pNewBuffer
   m_BufferLen += nNumBytes
   ' Mark the end of the string with a double null
   m_pBuffer[m_BufferLen] = 0
   m_pBuffer[m_BufferLen + 1] = 0
   RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' The incoming string parameter is inserted in the string starting at the position
' given by nIndex. nIndex = 1 for the first character, 2 For the second, etc.
' If nIndex is beyond the current length of the string + 1, no operation is performed.
' ========================================================================================
PRIVATE FUNCTION CWstr.Insert (BYVAL pwszStr AS WSTRING PTR, BYVAL nIndex AS LONG) AS BOOLEAN
   CWSTR_DP("CWSTR Insert WSTRING")
   IF nIndex < 1 OR nIndex > m_BufferLen \ 2 THEN RETURN FALSE
   DIM AS LONG nLenString = .LEN(*pwszStr)
   IF nLenString = 0 THEN RETURN FALSE
   RETURN this.InsertBuffer(cast(ANY PTR, pwszStr), nIndex, nLenString * 2)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CWstr.Insert (BYREF ansiStr AS STRING, BYVAL nIndex AS LONG, BYVAL nCodePage AS UINT = 0) AS BOOLEAN
   CWSTR_DP("CWSTR Insert STRING")
   IF nIndex < 1 OR nIndex > m_BufferLen \ 2 OR .LEN(ansiStr) = 0 THEN RETURN FALSE
   ' Create the wide string from the incoming ansi string
   DIM dwLen AS DWORD, pbuffer AS ANY PTR
   IF nCodePage = 0 THEN nCodePage = m_CodePage
   IF nCodePage = CP_UTF8 THEN
      dwLen = MultiByteToWideChar(CP_UTF8, 0, STRPTR(ansiStr), LEN(ansiStr), NULL, 0)
      IF dwLen = 0 THEN RETURN FALSE
      dwLen *= 2
      pbuffer = Allocate(dwLen)
      MultiByteToWideChar(CP_UTF8, 0, STRPTR(ansiStr), LEN(ansiStr), pbuffer, dwLen)
   ELSE
      dwLen = .LEN(ansiStr)
      dwLen *= 2
      pbuffer = Allocate(dwLen)
      MultiByteToWideChar(m_CodePage, MB_PRECOMPOSED, STRPTR(ansiStr), .LEN(ansiStr), pbuffer, dwLen)
   END IF
   ' Copy the string into the buffer and update the length
   FUNCTION = this.InsertBuffer(pbuffer, nIndex, dwLen)
   Deallocate(pbuffer)
END FUNCTION
' ========================================================================================

END NAMESPACE

' ########################################################################################
'            *** Overloading of some intrinsic functions of Free Basic ***
' ########################################################################################

USING Afx

' ========================================================================================
PRIVATE OPERATOR & (BYREF cws1 AS CWSTR, BYREF cws2 AS CWSTR) AS CWSTR
   DIM nLen1 AS LONG = .LEN(**cws1) * 2
   DIM nLen2 AS LONG = .LEN(**cws2) * 2
   DIM pBuffer AS UBYTE PTR = Allocate(nLen1 + nLen2 + 2)
   CWSTR_DP("CWSTR OPERATOR & - CWSTR & CWSTR - pBuffer = " & WSTR(pBuffer))
   IF pBuffer THEN
      memcpy(pBuffer, @cws1, nLen1)
      memcpy(pBuffer + nLen1, @cws2, nLen2)
      pBuffer[nLen1 + nLen2] = 0
      pBuffer[nLen1 + nLen2 + 1] = 0
      RETURN CWSTR(cast(WSTRING PTR, pBuffer), TRUE)
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR & (BYREF cws AS CWSTR, BYREF wsz AS WSTRING) AS CWSTR
   DIM nLen1 AS LONG = .LEN(**cws) * 2
   DIM nLen2 AS LONG = .LEN(wsz) * 2
   DIM pBuffer AS UBYTE PTR = Allocate(nLen1 + nLen2 + 2)
   CWSTR_DP("CWSTR OPERATOR & - CWSTR & WSTRING - pBuffer = " & WSTR(pBuffer))
   IF pBuffer THEN
      memcpy(pBuffer, @cws, nLen1)
      memcpy(pBuffer + nLen1, @wsz, nLen2)
      pBuffer[nLen1 + nLen2] = 0
      pBuffer[nLen1 + nLen2 + 1] = 0
      RETURN CWSTR(cast(WSTRING PTR, pBuffer), TRUE)
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR & (BYREF wsz AS WSTRING, BYREF cws AS CWSTR) AS CWSTR
   DIM nLen1 AS LONG = .LEN(wsz) * 2
   DIM nLen2 AS LONG = .LEN(**cws) * 2
   DIM pBuffer AS UBYTE PTR = Allocate(nLen1 + nLen2 + 2)
   CWSTR_DP("CWSTR OPERATOR & - WSTRING & CWSTR - pBuffer = " & WSTR(pBuffer))
   IF pBuffer THEN
      memcpy(pBuffer, @wsz, nLen1)
      memcpy(pBuffer + nLen1, @cws, nLen2)
      pBuffer[nLen1 + nLen2] = 0
      pBuffer[nLen1 + nLen2 + 1] = 0
      RETURN CWSTR(cast(WSTRING PTR, pBuffer), TRUE)
   END IF
END OPERATOR
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION VAL (BYREF cws AS CWSTR) AS DOUBLE
   CWSTR_DP("CWSTR VAL")
   RETURN VAL(**cws)
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION RIGHT (BYREF cws AS CWSTR, BYREF n AS LONG) AS CWSTR
   CWSTR_DP("CWSTR RIGHT")
   IF n <= 0 THEN RETURN ""
   IF n > cws.m_BufferLen \ 2 THEN RETURN cws
   RETURN cast(WSTRING PTR, cws.m_pBuffer + cws.m_BufferLen - (n * 2))
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION LEFT (BYREF cws AS CWSTR, BYREF n AS LONG) AS CWSTR
   CWSTR_DP("CWSTR LEFT")
   IF n <= 0 THEN EXIT FUNCTION
   IF n > cws.m_BufferLen \ 2 THEN FUNCTION = cws : EXIT FUNCTION
   ' Cast the buffer as a WSTRING
   DIM pwsz AS WSTRING PTR = cast(WSTRING PTR, cws.m_pBuffer)
   ' Store the value of the character at position n
   DIM nChar AS USHORT = pwsz[n]
   ' Change it to null to shorten the buffer to n characters
   pwsz[n] = 0
   ' Return the shortened buffer
   FUNCTION = pwsz
   ' Restore the value of the character to its original value
   pwsz[n] = nChar
END FUNCTION
' ========================================================================================
