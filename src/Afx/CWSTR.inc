' ########################################################################################
' Microsoft Windows
' Implements a dynamic data type for null terminated unicode strings.
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
#INCLUDE ONCE "utf_conv.bi"

' // The definition for BSTR in the FreeBASIC headers was inconveniently changed to WCHAR
#ifndef AFX_BSTR
   #define AFX_BSTR WSTRING PTR
#ENDIF

NAMESPACE Afx

' // Forward reference
TYPE CWSTR_ AS CWSTR
TYPE CBSTR_ AS CBSTR

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

' ========================================================================================
' Macro for debug
' To allow debugging, define _CBSTR_DEBUG_ 1 in your application before including this file.
' ========================================================================================
#ifndef _CBSTR_DEBUG_
   #define _CBSTR_DEBUG_ 0
#ENDIF
#ifndef _CBSTR_DP_
   #define _CBSTR_DP_ 1
   #MACRO CBSTR_DP(st)
      #IF (_CBSTR_DEBUG_ = 1)
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
FUNCTION AfxIsBstr (BYVAL pv aS ANY PTR) AS BOOLEAN
   IF pv = NULL THEN RETURN FALSE
   DIM res AS DWORD = PEEK(DWORD, pv - 4) \ 2
   IF res = LEN(*cast(WSTRING PTR, pv)) THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' CBStr - OLE strings class
' ========================================================================================
TYPE CBSTR

   Public:
      m_bstr AS AFX_BSTR
      m_CodePage AS UINT

   Public:
      DECLARE CONSTRUCTOR
      DECLARE CONSTRUCTOR (BYVAL nCopePage AS UINT)
      DECLARE CONSTRUCTOR (BYREF wszStr AS CONST WSTRING = "")
      DECLARE CONSTRUCTOR (BYREF ansiStr AS STRING = "", BYVAL nCodePage AS UINT = 0)
      DECLARE CONSTRUCTOR (BYREF cbs AS CBStr)
      DECLARE CONSTRUCTOR (BYREF cws AS CWStr_)
      DECLARE CONSTRUCTOR (BYREF bstrHandle AS AFX_BSTR = NULL, BYVAL fAttach AS LONG = TRUE)
      DECLARE CONSTRUCTOR (BYVAL n AS LONGINT)
      DECLARE CONSTRUCTOR (BYVAL n AS DOUBLE)
      DECLARE DESTRUCTOR
'      DECLARE OPERATOR @ () AS AFX_BSTR PTR
      DECLARE FUNCTION bptr () AS AFX_BSTR
      DECLARE FUNCTION vptr () AS AFX_BSTR PTR
      DECLARE FUNCTION sptr () AS WSTRING PTR
      DECLARE FUNCTION wstr () BYREF AS WSTRING
      DECLARE FUNCTION wchar () AS WSTRING PTR
      DECLARE OPERATOR LET (BYREF ansiStr AS STRING)
      DECLARE OPERATOR LET (BYREF wszStr AS CONST WSTRING)
      DECLARE OPERATOR LET (BYREF cbs AS CBStr)
      DECLARE OPERATOR LET (BYREF cws AS CWStr_)
      DECLARE OPERATOR LET (BYREF bstrHandle AS AFX_BSTR)
      DECLARE OPERATOR LET (BYVAL n AS LONGINT)
      DECLARE OPERATOR LET (BYVAL n AS DOUBLE)
      DECLARE OPERATOR CAST () BYREF AS WSTRING
      DECLARE OPERATOR CAST () AS ANY PTR
      DECLARE PROPERTY CodePage () AS UINT
      DECLARE PROPERTY CodePage (BYVAL nCodePage AS UINT)
      DECLARE SUB Append (BYREF wszStr AS CONST WSTRING)
      DECLARE SUB Empty
      DECLARE SUB Clear
      DECLARE SUB Attach (BYVAL pbstrSrc AS AFX_BSTR)
      DECLARE FUNCTION Detach () AS AFX_BSTR
      DECLARE FUNCTION Copy () AS AFX_BSTR
      DECLARE OPERATOR += (BYREF wszStr AS CONST WSTRING)
      DECLARE OPERATOR += (BYREF cbs AS CBStr)
      DECLARE OPERATOR += (BYREF cws AS CWStr_)
      DECLARE OPERATOR += (BYVAL n AS LONGINT)
      DECLARE OPERATOR += (BYVAL n AS DOUBLE)
      DECLARE OPERATOR &= (BYREF wszStr AS CONST WSTRING)
      DECLARE OPERATOR &= (BYREF cbs AS CBStr)
      DECLARE OPERATOR &= (BYREF cws AS CWStr_)
      DECLARE OPERATOR &= (BYVAL n AS LONGINT)
      DECLARE OPERATOR &= (BYVAL n AS DOUBLE)
      DECLARE FUNCTION LeftChars (BYVAL nChars AS LONG) AS CBSTR
      DECLARE FUNCTION RightChars (BYVAL nChars AS LONG) AS CBSTR
      DECLARE FUNCTION MidChars (BYVAL nStart AS LONG, BYVAL nChars AS LONG = 0) AS CBSTR
      DECLARE FUNCTION ValLong () AS LONG
      DECLARE FUNCTION ValInt () AS LONG
      DECLARE FUNCTION ValULong () AS ULONG
      DECLARE FUNCTION ValUInt () AS ULONG
      DECLARE FUNCTION ValLongInt () AS LONGINT
      DECLARE FUNCTION ValULongInt () AS ULONGINT
      DECLARE FUNCTION ValDouble () AS DOUBLE
      DECLARE FUNCTION Value () AS DOUBLE
      DECLARE PROPERTY Utf8 () AS STRING
      DECLARE PROPERTY Utf8 (BYREF utf8String AS STRING)

END TYPE
' ========================================================================================

' ########################################################################################
'                                  *** CWSTR CLASS ***
' ########################################################################################
TYPE CWSTR

   Private:
      m_Capacity AS UINT            ' The total size of the buffer
      m_GrowSize AS LONG = 260 * 2  ' How much to grow the buffer by when required
      m_CodePage AS UINT            ' Unicode code page

   Public:
      m_pBuffer AS UBYTE PTR        ' Pointer to the buffer
      m_BufferLen AS UINT           ' Length in bytes of the current string in the buffer

      DECLARE CONSTRUCTOR
      DECLARE CONSTRUCTOR (BYVAL nChars AS UINT, BYVAL nCodePage AS UINT)
      DECLARE CONSTRUCTOR (BYVAL pwszStr AS WSTRING PTR)
      DECLARE CONSTRUCTOR (BYREF ansiStr AS STRING, BYVAL nCodePage AS UINT = 0)
      DECLARE CONSTRUCTOR (BYREF cws AS CWSTR)
      DECLARE CONSTRUCTOR (BYREF cbs AS CBSTR_)
      DECLARE CONSTRUCTOR (BYVAL n AS LONGINT)
      DECLARE CONSTRUCTOR (BYVAL n AS DOUBLE)
      DECLARE DESTRUCTOR
      DECLARE SUB ResizeBuffer (BYVAL nValue AS UINT)
      DECLARE FUNCTION AppendBuffer (BYVAL addrMemory AS ANY PTR, BYVAL nNumBytes AS UINT) AS BOOLEAN
      DECLARE FUNCTION InsertBuffer (BYVAL addrMemory AS ANY PTR, BYVAL nIndex AS UINT, BYVAL nNumBytes AS UINT) AS BOOLEAN
      DECLARE PROPERTY GrowSize () AS LONG
      DECLARE PROPERTY GrowSize (BYVAL nValue AS LONG)
      DECLARE PROPERTY Capacity () AS UINT
      DECLARE PROPERTY Capacity (BYVAL nValue AS UINT)
      DECLARE PROPERTY SizeAlloc (BYVAL nChars AS UINT)
      DECLARE PROPERTY SizeOf () AS UINT
      DECLARE PROPERTY CodePage () AS UINT
      DECLARE PROPERTY CodePage (BYVAL nCodePage AS UINT)
      DECLARE SUB Clear
      DECLARE SUB Add (BYREF cws AS CWSTR)
      DECLARE SUB Add (BYREF cbs AS CBSTR_)
      DECLARE SUB Add (BYVAL pwszStr AS WSTRING PTR)
      DECLARE SUB Add (BYREF ansiStr AS STRING, BYVAL nCodePage AS UINT = 0)
      DECLARE PROPERTY Char(BYVAL nIndex AS UINT) AS USHORT
      DECLARE PROPERTY Char(BYVAL nIndex AS UINT, BYVAL nValue AS USHORT)
      DECLARE OPERATOR [] (BYVAL nIndex AS UINT) AS USHORT
      DECLARE FUNCTION DelChars (BYVAL nIndex AS UINT, BYVAL nCount AS UINT) AS BOOLEAN
		DECLARE FUNCTION Insert (BYREF cws AS CWSTR, BYVAL nIndex AS UINT) AS BOOLEAN
      DECLARE FUNCTION Insert (BYREF cbs AS CBSTR_, BYVAL nIndex AS LONG) AS BOOLEAN
      DECLARE FUNCTION Insert (BYVAL pwszStr AS WSTRING PTR, BYVAL nIndex AS UINT) AS BOOLEAN
      DECLARE FUNCTION Insert (BYREF ansiStr AS STRING, BYVAL nIndex AS UINT, BYVAL nCodePage AS UINT = 0) AS BOOLEAN
      DECLARE OPERATOR CAST () BYREF AS WSTRING
      DECLARE OPERATOR CAST () AS ANY PTR
      DECLARE OPERATOR LET (BYREF ansiStr AS STRING)
      DECLARE OPERATOR LET (BYREF wszStr AS CONST WSTRING)
      DECLARE OPERATOR LET (BYREF pwszStr AS WSTRING PTR)
      DECLARE OPERATOR LET (BYREF cws AS CWSTR)
      DECLARE OPERATOR LET (BYREF cbs AS CBSTR_)
      DECLARE OPERATOR LET (BYVAL n AS LONGINT)
      DECLARE OPERATOR LET (BYVAL n AS DOUBLE)
'      DECLARE OPERATOR @ () AS WSTRING PTR
      DECLARE OPERATOR += (BYREF wszStr AS WSTRING)
      DECLARE OPERATOR += (BYREF cws AS CWSTR)
      DECLARE OPERATOR += (BYREF cbs AS CBSTR_)
      DECLARE OPERATOR += (BYREF ansiStr AS STRING)
      DECLARE OPERATOR += (BYVAL n AS LONGINT)
      DECLARE OPERATOR += (BYVAL n AS DOUBLE)
      DECLARE OPERATOR &= (BYREF wszStr AS WSTRING)
      DECLARE OPERATOR &= (BYREF cws AS CWSTR)
      DECLARE OPERATOR &= (BYREF cbs AS CBSTR_)
      DECLARE OPERATOR &= (BYREF ansiStr AS STRING)
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
      DECLARE FUNCTION LeftChars (BYVAL nChars AS LONG) AS CWSTR
      DECLARE FUNCTION RightChars (BYVAL nChars AS LONG) AS CWSTR
      DECLARE FUNCTION MidChars (BYVAL nStart AS LONG, BYVAL nChars AS LONG = 0) AS CWSTR
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
' CWSTR constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CWSTR
   CWSTR_DP("+++BEGIN- CWSTR CONSTRUCTOR Default")
   this.ResizeBuffer(m_GrowSize)   ' Create the initial buffer
   CWSTR_DP("END CWSTR CONSTRUCTOR Default - " & .WSTR(m_pBuffer))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CWSTR (BYVAL nChars AS UINT, BYVAL nCodePage AS UINT)
   CWSTR_DP("+++BEGIN- CWSTR CONSTRUCTOR nChars " & .WSTR(nChars))
   m_CodePage = nCodePage          ' Store the code page
   IF nChars = 0 THEN nChars = m_GrowSize \ 2
   this.ResizeBuffer(nChars * 2)   ' Create the initial buffer
   CWSTR_DP("-END- CWSTR CONSTRUCTOR nChars - " & .WSTR(m_pBuffer))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CWstr (BYVAL pwszStr AS WSTRING PTR)
   CWSTR_DP("+++BEGIN- CWSTR CONSTRUCTOR WSTRING - " & .WSTR(pwszStr))
   IF pwszStr = NULL THEN
      this.ResizeBuffer(m_GrowSize)   ' Create the initial buffer
   ELSE
      this.Add(pwszStr)               ' Add the passed WSTRING
   END IF
   CWSTR_DP("-END- CWSTR CONSTRUCTOR WSTRING - " & .WSTR(m_pBuffer))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CWstr (BYREF ansiStr AS STRING, BYVAL nCodePage AS UINT = 0)
   CWSTR_DP("+++BEGIN- CWSTR CONSTRUCTOR STRING - " & .WSTR(VARPTR(ansiStr)))
   m_CodePage = nCodePage             ' Store the code page
   IF .LEN(ansiStr) THEN
      this.Add(ansiStr, nCodePage)    ' Add the passed ansi string
   ELSE
      this.ResizeBuffer(m_GrowSize)   ' Create the initial buffer
   END IF
   CWSTR_DP("-END- CWSTR CONSTRUCTOR STRING - " & .WSTR(m_pBuffer))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CWstr (BYREF cws AS CWSTR)
   CWSTR_DP("+++BEGIN- CWSTR CONSTRUCTOR CWSTR")
   IF cws.m_BufferLen THEN
      this.Add(cws)                   ' Add the passed CWSTR
   ELSE
      this.ResizeBuffer(m_GrowSize)   ' Create the initial buffer
   END IF
   CWSTR_DP("-END- CWSTR CONSTRUCTOR CWSTR - " & .WSTR(m_pBuffer))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CWstr (BYREF cbs AS CBStr_)
   CWSTR_DP("+++BEGIN- CWSTR CONSTRUCTOR CBSTR")
   this.ResizeBuffer(m_GrowSize)   ' Create the initial buffer
   ' Copy the string into the buffer and update the length
   this.AppendBuffer(cast(ANY PTR, cbs), SysStringLen(cbs) * 2)
   CWSTR_DP("-END- CWSTR CONSTRUCTOR CBSTR - " & .WSTR(m_pBuffer))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CWstr (BYVAL n AS LONGINT)
   CWSTR_DP("+++BEGIN- CWSTR CONSTRUCTOR LONGINT")
   DIM wsz AS WSTRING * 260 = .WSTR(n)
   this.Add(wsz)
   CWSTR_DP("-END- CWSTR CONSTRUCTOR LONGINT - " & .WSTR(m_pBuffer))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CWstr (BYVAL n AS DOUBLE)
   CWSTR_DP("+++BEGIN- CWSTR CONSTRUCTOR DOUBLE")
   DIM wsz AS WSTRING * 260 = .WSTR(n)
   this.Add(wsz)
   CWSTR_DP("-END- CWSTR CONSTRUCTOR DOUBLE - " & .WSTR(m_pBuffer))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR CWstr
   CWSTR_DP("***CWSTR DESTRUCTOR - buffer: " & .WSTR(m_pBuffer))
   IF m_pBuffer THEN Deallocate(m_pBuffer)
   m_pBuffer = NULL
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the address of the CWSTR buffer.
' Removed to allow to use @ to get the address of the class.
' ========================================================================================
'PRIVATE OPERATOR CWstr.@ () AS WSTRING PTR
'   CWSTR_DP("CWSTR OPERATOR @ - buffer: " & .WSTR(m_pBuffer))
'   OPERATOR = cast(WSTRING PTR, m_pBuffer)
'END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CWstr.vptr () AS WSTRING PTR
   CWSTR_DP("CWSTR vptr - buffer: " & .WSTR(m_pBuffer))
   RETURN cast(WSTRING PTR, m_pBuffer)
END FUNCTION
' ========================================================================================

' ========================================================================================
' One * returns the address of the CWSTR buffer.
' Two ** deferences the string data.
' Needed because LEFT and RIGHT (cws) fail with an ambiguous call error.
' We have to use **cws (notice the double indirection) with these functions.
' ========================================================================================
PRIVATE OPERATOR * (BYREF cws AS CWSTR) AS WSTRING PTR
   CWSTR_DP("CWSTR OPERATOR * buffer: " & .WSTR(cws.m_pBuffer))
   OPERATOR = cast(WSTRING PTR, cws.m_pBuffer)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns the address of the CWSTR buffer.
' ========================================================================================
PRIVATE FUNCTION CWstr.sptr () AS WSTRING PTR
   CWSTR_DP("CWSTR sptr - buffer: " & .WSTR(m_pBuffer))
   RETURN cast(WSTRING PTR, m_pBuffer)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the length, in characters, of the CWSTR.
' ========================================================================================
'PRIVATE OPERATOR LEN (BYREF cws AS CWSTR) AS UINT
'   CWSTR_DP("CWSTR OPERATOR LEN - len: " & .WSTR(.LEN(**cws)))
'   OPERATOR = .LEN(**cws)
'END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR LEN (BYREF cws AS CWSTR) AS UINT
   CWSTR_DP("CWSTR OPERATOR LEN - len: " & .WSTR(cws.m_BufferLen \ 2))
   OPERATOR = cws.m_BufferLen \ 2
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns a pointer to the CWSTR buffer.
' ========================================================================================
PRIVATE OPERATOR CWstr.CAST () AS ANY PTR
   CWSTR_DP("CWSTR CAST ANY PTR - buffer: " & .WSTR(m_pBuffer))
   OPERATOR = cast(ANY PTR, m_pBuffer)
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the string data (same as **).
' ========================================================================================
PRIVATE OPERATOR CWstr.CAST () BYREF AS WSTRING
   CWSTR_DP("CWSTR CAST BYREF AS WSTRING - buffer: " & .WSTR(m_pBuffer))
   OPERATOR = *cast(WSTRING PTR, m_pBuffer)
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the string data (same as **).
' ========================================================================================
PRIVATE FUNCTION CWstr.wstr () BYREF AS WSTRING
   CWSTR_DP("CWSTR wstr - buffer: " & .WSTR(m_pBuffer))
   RETURN *cast(WSTRING PTR, m_pBuffer)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets/sets the code page used to ansi to unicode translations
' ========================================================================================
PRIVATE PROPERTY CWStr.CodePage () AS UINT
   CWSTR_DP("CWSTR PROPERTY GET CodePage - " & .WSTR(m_CodePage))
   PROPERTY = m_CodePage
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWStr.CodePage (BYVAL nCodePage AS UINT)
   CWSTR_DP("CWSTR PROPERTY SET CodePage - " & .WSTR(nCodePage))
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
   CWSTR_DP("CWSTR LET WSTRING PTR = " & .WSTR(pwszStr))
   IF pwszStr = NULL THEN EXIT OPERATOR
   this.Clear
   this.Add(*pwszStr)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CWstr.Let (BYREF cws AS CWSTR)
   CWSTR_DP("CWSTR LET CWSTR - m_pBuffer = " & .WSTR(m_pBuffer) & " - IN buffer = " & .WSTR(cws.m_pBuffer))
   IF m_pBuffer = cws.m_pBuffer THEN EXIT OPERATOR   ' // Ignore cws = cws
   this.Clear
   this.Add(cws)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CWstr.Let (BYREF cbs AS CBStr_)
   CWSTR_DP("CWSTR LET CBSTR")
   this.Clear
   DIM AS LONG nLenString = SysStringLen(cbs)
   this.ResizeBuffer(nLenString)
   ' Copy the string into the buffer and update the length
   this.AppendBuffer(cast(ANY PTR, cbs), nLenString * 2)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CWstr.Let (BYVAL n AS LONGINT)
   CWSTR_DP("CWSTR OPERATOR Let LONGINT")
   this.Clear
   DIM wsz AS WSTRING * 260 = .WSTR(n)
   this.Add(wsz)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CWstr.Let (BYVAL n AS DOUBLE)
   CWSTR_DP("CWSTR OPERATOR Let DOUBLE")
   this.Clear
   DIM wsz AS WSTRING * 260 = .WSTR(n)
   this.Add(wsz)
END OPERATOR
' ========================================================================================

' ========================================================================================
' ResizeBuffer
' Increases the size of the internal buffer capacity
' ========================================================================================
PRIVATE SUB CWstr.ResizeBuffer (BYVAL nValue AS UINT)
   CWSTR_DP("CWSTR ResizeBuffer - Value = " & .WSTR(nValue))
   ' // If it is an odd value, make it even.
   IF (nValue MOD 2) <> 0 THEN nValue += 1
   ' Increase the size of the existing buffer by creating a new buffer copying
   ' the existing data into it and then finally deleting the original buffer.
   DIM pNewBuffer AS UBYTE PTR = Allocate(nValue + 2)   ' // +2 to make room for the double null terminator.
   CWSTR_DP("CWSTR ResizeBuffer - pNewBuffer = " & .WSTR(pNewBuffer) & " - old buffer = " & (.WSTR(m_pBuffer)))
   IF m_pBuffer THEN
      IF nValue < m_BufferLen THEN m_BufferLen = nValue
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
' Appends the specified number of bytes from the specified memory address to the end of the buffer.
' ========================================================================================
PRIVATE FUNCTION CWstr.AppendBuffer (BYVAL addrMemory AS ANY PTR, BYVAL nNumBytes AS UINT) AS BOOLEAN
   CWSTR_DP("CWSTR AppendBuffer " & .WSTR(m_BufferLen) & " " & .WSTR(nNumBytes))
'   IF (m_BufferLen + nNumBytes) > m_Capacity THEN this.ResizeBuffer(m_BufferLen + nNumBytes)
   ' // Add extra capacity to minimize multiple allocations when doing multiple concatenations
'   IF (m_BufferLen + nNumBytes) > m_Capacity THEN this.ResizeBuffer((m_BufferLen + nNumBytes) * 2)
   IF m_GrowSize < 0 THEN
      IF (m_BufferLen + nNumBytes) > m_Capacity THEN this.ResizeBuffer((m_BufferLen + nNumBytes) * 2)
   ELSE
      IF (m_BufferLen + nNumBytes) > m_Capacity THEN this.ResizeBuffer(m_BufferLen + nNumBytes + m_GrowSize)
   END IF
   IF m_pBuffer = NULL THEN RETURN FALSE
   memcpy(m_pBuffer + m_BufferLen, addrMemory, nNumBytes)
   m_BufferLen += nNumBytes
   ' Mark the end of the string with a double null
   m_pBuffer[m_BufferLen] = 0
   m_pBuffer[m_BufferLen + 1] = 0
   RETURN TRUE
   CWSTR_DP("--END - CWSTR AppendBuffer " & .WSTR(m_BufferLen))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The string parameter is appended to the string held in the class. If the internal string
' buffer overflows, the class will automatically extend it to an appropriate size.
' ========================================================================================
PRIVATE SUB CWstr.Add (BYREF cws AS CWSTR)
   CWSTR_DP("CWSTR Add CWSTR - buffer = " & .WSTR(cws.m_pBuffer) & " - LEN = " & .WSTR(LEN(*cast(WSTRING PTR, @cws))))
   ' Incoming string is already in wide format, simply copy it to the buffer.
   DIM nLenString AS UINT = cws.m_BufferLen   ' // Length in bytes
   IF nLenString = 0 THEN RETURN
   ' Copy the string into the buffer and update the length
   this.AppendBuffer(cast(ANY PTR, cws), nLenString)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CWstr.Add (BYREF cbs AS CBSTR_)
   CWSTR_DP("CWSTR Add CBSTR")
   ' Incoming string is already in wide format, simply copy it to the buffer.
   DIM AS LONG nLenString = SysStringLen(cbs)
   IF nLenString = 0 THEN RETURN
   ' Copy the string into the buffer and update the length
   this.AppendBuffer(cast(ANY PTR, cbs), nLenString * 2)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CWstr.Add (BYVAL pwszStr AS WSTRING PTR)
   CWSTR_DP("CWSTR Add WSTRING")
   IF pwszStr = NULL THEN RETURN
   ' Incoming string is already in wide format
   DIM nLenString AS UINT = .LEN(*pwszStr)   ' // Length in characters
   IF nLenString = 0 THEN RETURN
   ' Copy the string into the buffer and update the length
   this.AppendBuffer(cast(ANY PTR, pwszStr), nLenString * 2)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CWstr.Add (BYREF ansiStr AS STRING, BYVAL nCodePage AS UINT = 0)
   CWSTR_DP("CWSTR Add STRING Code page = " & .WSTR (nCodePage))
   IF LEN(ansiStr) = 0 THEN RETURN
   ' Create the wide string from the incoming ansi string
   DIM dwLen AS UINT, pbuffer AS ANY PTR
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
' Appends a CBSTR to the CWSTR.
' ========================================================================================
PRIVATE OPERATOR CWStr.+= (BYREF cbs AS CBStr_)
   CWSTR_DP("CWSTR OPERATOR += CBSTR")
   this.Add(cbs)
END OPERATOR
' ========================================================================================

' ========================================================================================
PRIVATE OPERATOR CWStr.+= (BYVAL n AS LONGINT)
   CWSTR_DP("CWSTR OPERATOR += LONGINT")
   DIM wsz AS WSTRING * 260 = .WSTR(n)
   this.Add(wsz)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CWStr.+= (BYVAL n AS DOUBLE)
   CWSTR_DP("CWSTR OPERATOR += DOUBLE")
   DIM wsz AS WSTRING * 260 = .WSTR(n)
   this.Add(wsz)
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
' Appends a CBSTR to the CWSTR.
' ========================================================================================
PRIVATE OPERATOR CWStr.&= (BYREF cbs AS CBStr_)
   CWSTR_DP("CWSTR OPERATOR &= CBSTR")
   this.Add(cbs)
END OPERATOR
' ========================================================================================

' ========================================================================================
PRIVATE OPERATOR CWStr.&= (BYVAL n AS LONGINT)
   CWSTR_DP("CWSTR OPERATOR &= LONGINT")
   DIM wsz AS WSTRING * 260 = .WSTR(n)
   this.Add(wsz)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CWStr.&= (BYVAL n AS DOUBLE)
   CWSTR_DP("CWSTR OPERATOR &= DOUBLE")
   DIM wsz AS WSTRING * 260 = .WSTR(n)
   this.Add(wsz)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Number of characters to preallocate to minimize multiple allocations when doing multiple
' concatenations. A value of less than 0 indicates that it must double the capacity each
' time that the buffer needs to be resized.
' ========================================================================================
PRIVATE PROPERTY CWstr.GrowSize() AS LONG
   CWSTR_DP("CWSTR PROPERTY GET GrowSize")
   IF m_GrowSize > -1 THEN PROPERTY = m_GrowSize \ 2 ELSE PROPERTY = m_GrowSize
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWstr.GrowSize (BYVAL nChars AS LONG)
   CWSTR_DP("CWSTR PROPERTY SET Growsize")
   IF nChars > -1 THEN m_GrowSize = nChars * 2 ELSE m_GrowSize = -1
END PROPERTY
' ========================================================================================

' ========================================================================================
' The size of the internal string buffer is retrieved and returned to the caller. The size
' is the number of bytes which can be stored without further expansion.
' ========================================================================================
PRIVATE PROPERTY CWstr.Capacity() AS UINT
   CWSTR_DP("CWSTR PROPERTY GET Capacity")
   PROPERTY = m_Capacity
END PROPERTY
' ========================================================================================
' ========================================================================================
' The internal string buffer is expanded to the specified number of bytes. If the new
' capacity is smaller or equal to the current capacity, no operation is performed. If it is
' smaller, the buffer is shortened and the contents that exceed the new capacity are lost.
' ========================================================================================
PRIVATE PROPERTY CWstr.Capacity (BYVAL nValue AS UINT)
   CWSTR_DP("CWSTR PROPERTY SET Capacity")
   ' // If the new capacity is the same that the current capacity, do nothing.
   IF nValue = m_Capacity THEN EXIT PROPERTY
   ' // Make sure that the number is odd (ResizeBuffer already does it)
'   IF (nValue MOD 2) <> 0 THEN nValue += 1
   this.ResizeBuffer(nValue)
END PROPERTY
' ========================================================================================
' ========================================================================================
' The internal string buffer is expanded to the specified number of byyes.
' Sets the capacity of the buffer in characters. If the new capacity is equal to the
' current capacity, no operation is performed. If it is smaller, the buffer is shortened
' and the contents that exceed the new capacity are lost.
' ========================================================================================
PRIVATE PROPERTY CWstr.SizeAlloc (BYVAL nChars AS UINT)
   CWSTR_DP("CWSTR PROPERTY SET SizeAlloc")
   ' // If the new capacity is the same that the current capacity, do nothing.
   IF nChars = m_Capacity \ 2 THEN EXIT PROPERTY
   this.ResizeBuffer(nChars * 2)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Returns the capacity of the buffer in characters.
' ========================================================================================
PRIVATE PROPERTY CWstr.SizeOf() AS UINT
   CWSTR_DP("CWSTR PROPERTY GET SizeOf")
   PROPERTY = m_Capacity \ 2
END PROPERTY
' ========================================================================================

' ========================================================================================
' Returns the corresponding unicode integer representation of the character at the position
' specified by the nIndex parameter (1 for the first character, 2 for the second, etc.).
' If nIndex is beyond the current length of the string, a 0 is returned.
' ========================================================================================
PRIVATE PROPERTY CWstr.Char (BYVAL nIndex AS UINT) AS USHORT
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
PRIVATE PROPERTY CWstr.Char (BYVAL nIndex AS UINT, BYVAL nValue AS USHORT)
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
PRIVATE OPERATOR CWStr.[] (BYVAL nIndex AS UINT) AS USHORT
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
PRIVATE FUNCTION CWstr.DelChars (BYVAL nIndex AS UINT, BYVAL nCount AS UINT) AS BOOLEAN
   CWSTR_DP("CWSTR DelChars")
   IF nIndex < 1 OR nIndex > m_BufferLen \ 2 OR nCount < 1 THEN RETURN FALSE
   DIM numChars AS UINT = m_BufferLen \ 2
   IF nCount > numChars - nIndex + 1 THEN nCount = numChars - nIndex + 1
   DIM nOffset AS UINT = (nIndex - 1 + nCount) * 2
   memcpy(m_pBuffer + (nOffset - nCount * 2), m_pBuffer + nOffset, m_BufferLen - (nCount * 2))
   m_BufferLen -= (nCount * 2)
   ' Mark the end of the string with a double null
   m_pBuffer[m_BufferLen] = 0
   m_pBuffer[m_BufferLen + 1] = 0
   RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Inserts the specified number of bytes from the specified memory address into the buffer.
' ========================================================================================
PRIVATE FUNCTION CWstr.InsertBuffer (BYVAL addrMemory AS ANY PTR, BYVAL nIndex AS UINT, BYVAL nNumBytes AS UINT) AS BOOLEAN
   CWSTR_DP("CWSTR InsertBuffer")
   IF nIndex < 1 OR nIndex > m_BufferLen \ 2 THEN RETURN FALSE
   ' Determine the size of the new buffer
   IF m_BufferLen + nNumBytes > m_Capacity THEN m_Capacity = m_BufferLen + nNumBytes
   DIM pNewBuffer AS UBYTE PTR = Allocate((m_Capacity + 1) * 2)
   IF m_pBuffer THEN
      nIndex -= 1
      ' Copy the existing data into the new buffer
      memcpy(pNewBuffer, m_pBuffer, nIndex * 2)
      DIM nOffset AS UINT = nIndex * 2
      memcpy(pNewBuffer + nOffset, addrMemory, nNumBytes)
      nOffset += nNumBytes
      memcpy(pNewBuffer + nOffset, m_pBuffer + (nIndex * 2), m_BufferLen - (nIndex * 2))
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
PRIVATE FUNCTION CWstr.Insert (BYREF cws AS CWSTR, BYVAL nIndex AS UINT) AS BOOLEAN
   CWSTR_DP("CWSTR Insert CWSTR")
   IF (nIndex < 1) OR nIndex > m_BufferLen \ 2 THEN RETURN FALSE
   IF cws.m_BufferLen = 0 THEN RETURN FALSE
   RETURN this.InsertBuffer(cast(ANY PTR,cws.m_pBuffer), nIndex, cws.m_BufferLen)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CWstr.Insert (BYREF cbs AS CBSTR_, BYVAL nIndex AS LONG) AS BOOLEAN
   CWSTR_DP("CWSTR Insert CBSTR")
   IF (nIndex < 1) OR nIndex > m_BufferLen \ 2 THEN RETURN 0
   DIM AS LONG nLenString = SysStringLen(cbs)
   IF nLenString = 0 THEN RETURN 0
   RETURN this.InsertBuffer(cast(ANY PTR, cbs), nIndex, nLenString * 2)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CWstr.Insert (BYVAL pwszStr AS WSTRING PTR, BYVAL nIndex AS UINT) AS BOOLEAN
   CWSTR_DP("CWSTR Insert WSTRING")
   IF nIndex < 1 OR nIndex > m_BufferLen \ 2 THEN RETURN FALSE
   DIM nLenString AS UINT = .LEN(*pwszStr)
   IF nLenString = 0 THEN RETURN FALSE
   RETURN this.InsertBuffer(cast(ANY PTR, pwszStr), nIndex, nLenString * 2)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CWstr.Insert (BYREF ansiStr AS STRING, BYVAL nIndex AS UINT, BYVAL nCodePage AS UINT = 0) AS BOOLEAN
   CWSTR_DP("CWSTR Insert STRING")
   IF nIndex < 1 OR nIndex > m_BufferLen \ 2 OR .LEN(ansiStr) = 0 THEN RETURN FALSE
   ' Create the wide string from the incoming ansi string
   DIM dwLen AS UINT, pbuffer AS ANY PTR
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

' ========================================================================================
' Converts the CWSTR to UTF8.
' ========================================================================================
PRIVATE PROPERTY CWstr.Utf8 () AS STRING
   CWSTR_DP("CWSTR Utf8 GET PROPERTY")
   DIM cbLen AS INTEGER
   IF m_BufferLen = 0 THEN RETURN ""
   DIM buffer AS STRING = STRING(m_BufferLen * 5 + 1, 0)
   PROPERTY = *cast(ZSTRING PTR, WCharToUTF(1, cast(WSTRING PTR, m_pBuffer), m_BufferLen, STRPTR(buffer), @cbLen))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Converts UTF8 to unicode and assigns it to the CWSTR.
' ========================================================================================
PRIVATE PROPERTY CWstr.Utf8 (BYREF utf8String AS STRING)
   CWSTR_DP("CWSTR Utf8 SET PROPERTY")
   this.Clear
   this.Add(utf8String, CP_UTF8)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Returns the leftmost substring of the string.
' ========================================================================================
PRIVATE FUNCTION CWStr.LeftChars (BYVAL nChars AS LONG) AS CWSTR
   CWSTR_DP("CWSTR LeftChars")
   RETURN LEFT(*cast(WSTRING PTR, m_pBuffer), nChars)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the rightmost substring of the string.
' ========================================================================================
PRIVATE FUNCTION CWStr.RightChars (BYVAL nChars AS LONG) AS CWSTR
   CWSTR_DP("CWBSTR LeftChars")
   RETURN RIGHT(*cast(WSTRING PTR, m_pBuffer), nChars)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a substring of the string.
' ========================================================================================
PRIVATE FUNCTION CWStr.MidChars (BYVAL nStart AS LONG, BYVAL nChars AS LONG = 0) AS CWSTR
   CWSTR_DP("CBSTR MidChars")
   IF nChars = 0 THEN RETURN MID(*cast(WSTRING PTR, m_pBuffer), nStart)
   RETURN MID(*cast(WSTRING PTR, m_pBuffer), nStart, nChars)
END FUNCTION
' ========================================================================================

' =====================================================================================
' Converts the string to a 32bit integer
' =====================================================================================
PRIVATE FUNCTION CWStr.ValLong () AS LONG
   RETURN .ValInt(*cast(WSTRING PTR, m_pBuffer))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CWStr.ValInt () AS LONG
   RETURN .ValInt(*cast(WSTRING PTR, m_pBuffer))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts the string to an unsigned 32bit integer
' =====================================================================================
PRIVATE FUNCTION CWStr.ValULong () AS ULONG
   RETURN .ValUInt(*cast(WSTRING PTR, m_pBuffer))
END FUNCTION
' =====================================================================================
PRIVATE FUNCTION CWStr.ValUInt () AS ULONG
   RETURN .ValUInt(*cast(WSTRING PTR, m_pBuffer))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts the string to a 64bit integer
' =====================================================================================
PRIVATE FUNCTION CWStr.ValLongInt () AS LONGINT
   RETURN .ValLng(*cast(WSTRING PTR, m_pBuffer))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts the string to an unsigned 64bit integer
' =====================================================================================
PRIVATE FUNCTION CWStr.ValULongInt () AS ULONGINT
   RETURN .ValULng(*cast(WSTRING PTR, m_pBuffer))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts the string to a floating point number (DOUBLE)
' =====================================================================================
PRIVATE FUNCTION CWStr.ValDouble () AS DOUBLE
   RETURN .VAL(*cast(WSTRING PTR, m_pBuffer))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CWStr.Value () AS DOUBLE
   RETURN .VAL(*cast(WSTRING PTR, m_pBuffer))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the contents of the CWSTR as a CBSTR.
' Useful to pass a CWSTR to a function that expects a BYVAL IN BSTR parameter.
' =====================================================================================
PRIVATE FUNCTION CWStr.cbstr () AS CBStr_
   RETURN SysAllocString(cast(WSTRING PTR, m_pBuffer))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the contents of the CWSTR as a BSTR.
' =====================================================================================
PRIVATE FUNCTION CWStr.bstr () AS AFX_BSTR
   RETURN SysAllocString(cast(WSTRING PTR, m_pBuffer))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the contents of the CWSTR as a WSTRING allocated with CoTaskMemAlloc.
' Free the returned string later with CoTaskMemFree.
' Note: This is useful when we need to pass a pointer to a null terminated wide string to a
' function or method that will release it. If we pass a WSTRING it will GPF.
' If the length of the input string is 0, CoTaskMemAlloc allocates a zero-length item and
' returns a valid pointer to that item. If there is insufficient memory available,
' CoTaskMemAlloc returns NULL.
' =====================================================================================
PRIVATE FUNCTION CWStr.wchar () AS WSTRING PTR
   DIM pwchar AS WSTRING PTR
   pwchar = CoTaskMemAlloc((m_BufferLen + 1) * 2)
   IF pwchar = NULL THEN RETURN NULL
   IF m_BufferLen THEN memcpy pwchar, m_pBuffer, m_BufferLen
   IF m_BufferLen = 0 THEN *pwchar = CHR(0)
   RETURN pwchar
END FUNCTION
' =====================================================================================

' ########################################################################################
'                                  *** CBSTR CLASS ***
' ########################################################################################

' ========================================================================================
' CBStr class constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CBStr
'   m_bstr = SysAllocString("")   ' // warning: don't initialize it to an empty string
   CBSTR_DP("CBSTR CONSTRUCTOR Default - " & .WSTR(m_bstr))
END CONSTRUCTOR
' ========================================================================================
PRIVATE CONSTRUCTOR CBStr (BYVAL nCodePage AS UINT)
'   m_bstr = SysAllocString("")   ' // warning: don't initialize it to an empty string
   m_CodePage = nCodePage
   CBSTR_DP("CBSTR CONSTRUCTOR CodePage - " & .WSTR(m_bstr))
END CONSTRUCTOR
' ========================================================================================
PRIVATE CONSTRUCTOR CBStr (BYREF wszStr AS CONST WSTRING = "")
   m_bstr = SysAllocString(wszStr)
   CBSTR_DP("CBSTR CONSTRUCTOR WSTRING - " & .WSTR(m_bstr))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CBStr (BYREF ansiStr AS STRING = "", BYVAL nCodePage AS UINT = 0)
   IF nCodePage <> 0 THEN m_CodePage = nCodePage
   IF m_CodePage = CP_UTF8 THEN
      DIM dwLen AS DWORD = MultiByteToWideChar(CP_UTF8, 0, STRPTR(ansiStr), LEN(ansiStr), NULL, 0)
      IF dwLen THEN
         m_bstr = SysAllocString(.WSTR(SPACE(dwLen)))
         MultiByteToWideChar(CP_UTF8, 0, STRPTR(ansiStr), LEN(ansiStr), m_bstr, dwLen * 2)
      ELSE
         m_bstr = SysAllocString("")
      END IF
   ELSE
      IF LEN(ansiStr) THEN
         m_bstr = SysAllocString(.WSTR(ansiStr))
         MultiByteToWideChar(m_CodePage, MB_PRECOMPOSED, STRPTR(ansiStr), -1, m_bstr, LEN(ansiStr) * 2)
      ELSE
         m_bstr = SysAllocString("")
      END IF
   END IF
   CBSTR_DP("CBSTR CONSTRUCTOR STRING - " & .WSTR(m_bstr))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CBStr (BYREF cbs AS CBStr)
   m_bstr = SysAllocString(cbs)
   CBSTR_DP("CBSTR CONSTRUCTOR CBSTR - " & .WSTR(m_bstr))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CBStr (BYREF cws AS CWStr)
   m_bstr = SysAllocString(cws)
   CBSTR_DP("CBSTR CONSTRUCTOR CBSTR - " & .WSTR(m_bstr))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CBStr (BYREF bstrHandle AS AFX_BSTR = NULL, BYVAL fAttach AS LONG = TRUE)
   CBSTR_DP("--BEGIN CBSTR CONSTRUCTOR AFX_BSTR - handle: " & .WSTR(bstrHandle) & " - Attach: " & .WSTR(fAttach))
   IF bstrHandle = NULL THEN
      m_bstr = SysAllocString("")
      CBSTR_DP("CBSTR CONSTRUCTOR SysAllocString - " & .WSTR(m_bstr))
   ELSE
      ' Detect if the passed handle is an OLE string
      ' If it is an OLE string it must have a descriptor; otherwise, don't
      ' Get the length in bytes looking at the descriptor and divide by 2 to get the number of
      ' unicode characters, that is the value returned by the FreeBASIC LEN operator.
      DIM Res AS INTEGER = PEEK(DWORD, CAST(ANY PTR, bstrHandle) - 4) \ 2
      ' If the retrieved length if the same that the returned by LEN, then it must be an OLE string
      IF Res = .LEN(*bstrHandle) AND fAttach <> FALSE THEN
         CBSTR_DP("CBSTR CONSTRUCTOR AFX_BSTR - Attach handle: " & .WSTR(bstrHandle))
         ' Attach the passed handle to the class
         m_bstr = bstrHandle
      ELSE
         CBSTR_DP("CBSTR CONSTRUCTOR AFX_BSTR - Alloc handle: " & .WSTR(bstrHandle))
         ' Allocate an OLE string with the contents of the string pointer by bstrHandle
         m_bstr = SysAllocString(*bstrHandle)
      END IF
   END IF
   CBSTR_DP("--END CBSTR CONSTRUCTOR AFX_BSTR - " & .WSTR(m_bstr))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CBstr (BYVAL n AS LONGINT)
   m_bstr = SysAllocString(.WSTR(n))
   CBSTR_DP("CBSTR CONSTRUCTOR LONGINT - " & .WSTR(m_bstr))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CBstr (BYVAL n AS DOUBLE)
   m_bstr = SysAllocString(.WSTR(n))
   CBSTR_DP("CBSTR CONSTRUCTOR DOUBLE - " & .WSTR(m_bstr))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CBStr class destructor
' ========================================================================================
PRIVATE DESTRUCTOR CBStr
   CBSTR_DP("CBSTR DESTRUCTOR - " & .WSTR(m_bstr))
   IF m_bstr THEN SysFreeString m_bstr
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the address of the BSTR
' Removed to allow to use @ to get the address of the class.
' ========================================================================================
'PRIVATE OPERATOR CBStr.@ () AS AFX_BSTR PTR
'   OPERATOR = @m_bstr
'END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns the underlying BSTR pointer.
' ========================================================================================
PRIVATE FUNCTION CBStr.bptr () AS AFX_BSTR
   CBSTR_DP("CBSTR bptr")
   RETURN m_bstr
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Frees the underlying BSTR and returns the BSTR pointer.
' To pass the underlying BSTR to an OUT BYVAL BSTR PTR parameter.
' If we pass a CBSTR to a function with an OUT BSTR parameter without first freeing it
' we will have a memory leak.
' ========================================================================================
PRIVATE FUNCTION CBStr.vptr () AS AFX_BSTR PTR
   CBSTR_DP("CBSTR vptr")
   IF m_bstr THEN
      SysFreeString(m_bstr)
      m_bstr = NULL
   END IF
   RETURN @m_bstr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the address of the CBSTR string data (same as **)
' ========================================================================================
PRIVATE FUNCTION CBStr.sptr () AS WSTRING PTR
   CBSTR_DP("CBSTR sptr")
   RETURN cast(WSTRING PTR, m_bstr)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Assigns new text to the BSTR
' ========================================================================================
PRIVATE OPERATOR CBStr.Let (BYREF wszStr AS CONST WSTRING)
   CBSTR_DP("CBSTR LET WSTRING")
   IF m_bstr THEN SysFreeString(m_bstr)
   m_bstr = SysAllocString(wszStr)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CBStr.Let (BYREF ansiStr AS STRING)
   CBSTR_DP("CBSTR LET STRING")
   IF m_bstr THEN SysFreeString(m_bstr)
   IF m_CodePage = CP_UTF8 THEN
      DIM dwLen AS DWORD = MultiByteToWideChar(CP_UTF8, 0, STRPTR(ansiStr), LEN(ansiStr), NULL, 0)
      IF dwLen THEN
         m_bstr = SysAllocString(.WSTR(SPACE(dwLen)))
         MultiByteToWideChar(CP_UTF8, 0, STRPTR(ansiStr), LEN(ansiStr), m_bstr, dwLen * 2)
      ELSE
         m_bstr = SysAllocString("")
      END IF
   ELSE
      IF LEN(ansiStr) THEN
         m_bstr = SysAllocString(.WSTR(ansiStr))
         MultiByteToWideChar(m_CodePage, MB_PRECOMPOSED, STRPTR(ansiStr), -1, m_bstr, LEN(ansiStr) * 2)
      ELSE
         m_bstr = SysAllocString("")
      END IF
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CBStr.Let (BYREF cbs AS CBStr)
   CBSTR_DP("CBSTR LET CBStr")
   IF m_bstr <> cbs.m_bstr THEN   ' // If the user has done cbs = cbs, ignore it
      IF m_bstr THEN SysFreeString(m_bstr)
      m_bstr = SysAllocString(cbs)
   END IF
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CBStr.Let (BYREF cws AS CWStr)
   CBSTR_DP("CBSTR LET CWSTR")
   IF m_bstr THEN SysFreeString(m_bstr)
   m_bstr = SysAllocString(cws)
END OPERATOR
' ========================================================================================
' ========================================================================================
' WARNING: If you pass the handle of a BSTR, don't free it, because it will be attached.
' ========================================================================================
PRIVATE OPERATOR CBStr.Let (BYREF bstrHandle AS AFX_BSTR)
   IF bstrHandle = NULL THEN EXIT OPERATOR
   ' Free the current OLE string
   IF m_bstr THEN SysFreeString(m_bstr)
   ' Detect if the passed handle is an OLE string.
   ' If it is an OLE string it must have a descriptor; otherwise, don't.
   ' Get the length in bytes looking at the descriptor and divide by 2 to get the number of
   ' unicode characters, that is the value returned by the FreeBASIC LEN operator.
   DIM Res AS DWORD = PEEK(DWORD, CAST(ANY PTR, bstrHandle) - 4) \ 2
   ' If the retrieved length is the same that the returned by LEN, then it must be an OLE string
   IF Res = .LEN(*bstrHandle) THEN
      CBSTR_DP("CBSTR LET AFX_BSTR - Attach handle: " & .WSTR(bstrHandle))
      ' Attach the passed handle to the class
      m_bstr = bstrHandle
   ELSE
      CBSTR_DP("CBSTR LET AFX_BSTR - Alloc handle: " & .WSTR(bstrHandle))
      ' Allocate an OLE string with the contents of the string pointed by bstrHandle
      m_bstr = SysAllocString(*bstrHandle)
   END IF
   CBSTR_DP("-END CBSTR LET AFX_BSTR - m_bstr: " & .WSTR(m_bstr))
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CBSTR.Let (BYVAL n AS LONGINT)
   CBSTR_DP("CBSTR OPERATOR Let LONGINT")
   this.Clear
   DIM wsz AS WSTRING * 260 = .WSTR(n)
   m_bstr = SysAllocString(wsz)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CBSTR.Let (BYVAL n AS DOUBLE)
   CBSTR_DP("CBSTR OPERATOR Let DOUBLE")
   this.Clear
   DIM wsz AS WSTRING * 260 = .WSTR(n)
   m_bstr = SysAllocString(wsz)
END OPERATOR
' ========================================================================================

' ========================================================================================
' One * returns the BSTR pointer.
' Two ** returns the adress of the start of the string data.
' Needed because LEFT and RIGHT (cbs) fail with an ambiguous call error.
' We have to use **cbs (notice the double indirection) with these functions.
' ========================================================================================
PRIVATE OPERATOR * (BYREF cbs AS CBStr) AS AFX_BSTR
   CBSTR_DP("CBSTR OPERATOR *")
   OPERATOR = cbs.m_bstr
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Returns a pointer to the string data (same as **)
' ========================================================================================
PRIVATE OPERATOR CBStr.CAST () BYREF AS WSTRING
   CBSTR_DP("CBSTR OPERATOR CAST BYREF AS WSTRING " & .WSTR(m_bstr))
   OPERATOR =  *CAST(WSTRING PTR, m_bstr)
END OPERATOR
' ========================================================================================
' ========================================================================================
' * Returns the BSTR pointer
' ========================================================================================
PRIVATE OPERATOR CBStr.CAST () AS ANY PTR
   CBSTR_DP("CBSTR OPERATOR CAST ANY PTR " & .WSTR(m_bstr))
   OPERATOR =  CAST(ANY PTR, m_bstr)
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Returns a pointer to the string data (same as **)
' ========================================================================================
PRIVATE FUNCTION CBStr.wstr () BYREF AS WSTRING
   CBSTR_DP("CBSTR wstr " & .WSTR(m_bstr))
   RETURN *CAST(WSTRING PTR, m_bstr)
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
PRIVATE FUNCTION CBStr.wchar () AS WSTRING PTR
   DIM pwchar AS WSTRING PTR
   DIM nLen AS LONG = SysStringLen(m_bstr) * 2
   pwchar = CoTaskMemAlloc(nLen)
   IF pwchar = NULL THEN RETURN NULL
   IF nLen THEN memcpy pwchar, m_bstr, nLen
   IF nLen = 0 THEN *pwchar = CHR(0)
   RETURN pwchar
END FUNCTION
' =====================================================================================

' ========================================================================================
' * Returns the length of the BSTR in characters.
' Needed because FB's LEN operator does not work with BSTRs.
' ========================================================================================
PRIVATE OPERATOR LEN (BYREF cbs AS CBStr) AS INTEGER
   CBSTR_DP("CBSTR OPERATOR LEN")
   OPERATOR = SysStringLen(cbs)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Gets/sets the code page used to ansi to unicode translations
' ========================================================================================
PRIVATE PROPERTY CBStr.CodePage () AS UINT
   CBSTR_DP("CBSTR PROPERTY GET CodePage")
   PROPERTY = m_CodePage
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CBStr.CodePage (BYVAL nCodePage AS UINT)
   CBSTR_DP("CBSTR PROPERTY SET CodePage")
   m_CodePage = nCodePage
END PROPERTY
' ========================================================================================

' ========================================================================================
' * Appends a string to the CBSTR. The string can be a literal or a FB STRING, a WSTRING,
' a CBSTR or a CWSTR variable.
' ========================================================================================
PRIVATE SUB CBStr.Append (BYREF wszStr AS CONST WSTRING)
   CBSTR_DP("CBSTR Append - WSTRING")
   DIM n1 AS UINT = SysStringLen(m_bstr)
   DIM nLen AS UINT = .LEN(wszStr)
   IF nLen = 0 THEN EXIT SUB
   DIM b AS AFX_BSTR = SysAllocStringLen(NULL, n1 + nLen)
   IF b = NULL THEN EXIT SUB
   memcpy(b, m_bstr, n1 * SIZEOF(WSTRING))
   memcpy(b + n1, @wszStr, nLen * SIZEOF(WSTRING))
   IF m_bstr THEN SysFreeString(m_bstr)
   m_bstr = b
END SUB
' ========================================================================================

' ========================================================================================
' * Frees the m_bstr member.
' ========================================================================================
PRIVATE SUB CBstr.Empty
   IF m_bstr THEN SysFreeString(m_bstr)
   m_bstr = NULL
END SUB
' ========================================================================================

' ========================================================================================
' * Frees the m_bstr member.
' ========================================================================================
PRIVATE SUB CBstr.Clear
   IF m_bstr THEN SysFreeString(m_bstr)
   m_bstr = NULL
END SUB
' ========================================================================================

' ========================================================================================
' * Attaches a BSTR to the CBSTR object by setting the m_bstr member to pbstrSrc.
' WARNING: Don't attach a CBSTR to another CBSTR because each CBSTR will try to free the
' same BSTR when they are destroyed.
' ========================================================================================
PRIVATE SUB CBstr.Attach (BYVAL pbstrSrc AS AFX_BSTR)
   IF m_bstr THEN SysFreeString(m_bstr)
   IF AfxIsBstr(pbstrSrc) THEN m_bstr = pbstrSrc
END SUB
' ========================================================================================

' ========================================================================================
' * Detaches m_bstr from the CBSTR object and sets m_bstr to NULL.
' ========================================================================================
PRIVATE FUNCTION CBstr.Detach () AS AFX_BSTR
   DIM pbstr AS AFX_BSTR = m_bstr
   m_bstr = NULL
   RETURN pbstr
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Allocates and returns a copy of m_bstr.
' ========================================================================================
PRIVATE FUNCTION CBStr.Copy () AS AFX_BSTR
   RETURN SysAllocStringLen(m_bstr, SysStringLen(m_bstr))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Appends a string to the BSTR. The string can be a literal or a FB STRING or WSTRING variable.
' ========================================================================================
PRIVATE OPERATOR CBStr.+= (BYREF wszStr AS CONST WSTRING)
   CBSTR_DP("CBSTR OPERATOR += WSTRING")
   this.Append(wszStr)
END OPERATOR
' ========================================================================================
' ========================================================================================
' * Appends a CBSTR to the CBSTR.
' ========================================================================================
PRIVATE OPERATOR CBStr.+= (BYREF cbs AS CBStr)
   CBSTR_DP("CBSTR OPERATOR += CBSTR")
   this.Append(**cbs)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Appends a CBSTR to the CBSTR.
' ========================================================================================
PRIVATE OPERATOR CBStr.+= (BYREF cbs AS CWStr)
   CBSTR_DP("CBSTR OPERATOR += CWSTR")
   this.Append(**cbs)
END OPERATOR
' ========================================================================================

' ========================================================================================
PRIVATE OPERATOR CBStr.+= (BYVAL n AS LONGINT)
   CBSTR_DP("CBSTR OPERATOR += LONGINT")
   DIM wsz AS WSTRING * 260 = .WSTR(n)
   this.Append(wsz)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CBStr.+= (BYVAL n AS DOUBLE)
   CBSTR_DP("CBSTR OPERATOR += DOUBLE")
   DIM wsz AS WSTRING * 260 = .WSTR(n)
   this.Append(wsz)
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Appends a string to the CBSTR. The string can be a literal or a FB STRING or WSTRING variable.
' ========================================================================================
PRIVATE OPERATOR CBStr.&= (BYREF wszStr AS CONST WSTRING)
   CBSTR_DP("CBSTR OPERATOR &= WSTRING")
   this.Append(wszStr)
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Appends a CBSTR to the CBSTR.
' ========================================================================================
PRIVATE OPERATOR CBStr.&= (BYREF cbs AS CBStr)
   CBSTR_DP("CBSTR OPERATOR &= CBSTR")
   this.Append(**cbs)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Appends a CBSTR to the CBSTR.
' ========================================================================================
PRIVATE OPERATOR CBStr.&= (BYREF cbs AS CWStr)
   CBSTR_DP("CBSTR OPERATOR &= CWSTR")
   this.Append(**cbs)
END OPERATOR
' ========================================================================================

' ========================================================================================
PRIVATE OPERATOR CBStr.&= (BYVAL n AS LONGINT)
   CBSTR_DP("CBSTR OPERATOR &= LONGINT")
   DIM wsz AS WSTRING * 260 = .WSTR(n)
   this.Append(wsz)
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CBStr.&= (BYVAL n AS DOUBLE)
   CBSTR_DP("CBSTR OPERATOR &= DOUBLE")
   DIM wsz AS WSTRING * 260 = .WSTR(n)
   this.Append(wsz)
END OPERATOR
' ========================================================================================

' ========================================================================================
' * Returns the leftmost substring of the string.
' ========================================================================================
PRIVATE FUNCTION CBStr.LeftChars (BYVAL nChars AS LONG) AS CBSTR
   CBSTR_DP("CBSTR LeftChars")
   DIM pbstr AS BSTR = SysAllocString(LEFT(*m_bstr, nChars))
   RETURN pbstr
   ' // Don't free psbtr because it will be attached to the returned CBSTR
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the rightmost substring of the string.
' ========================================================================================
PRIVATE FUNCTION CBStr.RightChars (BYVAL nChars AS LONG) AS CBSTR
   CBSTR_DP("CBSTR LeftChars")
   DIM pbstr AS BSTR = SysAllocString(RIGHT(*m_bstr, nChars))
   RETURN pbstr
   ' // Don't free psbtr because it will be attached to the returned CBSTR
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a substring of the string.
' ========================================================================================
PRIVATE FUNCTION CBStr.MidChars (BYVAL nStart AS LONG, BYVAL nChars AS LONG = 0) AS CBSTR
   CBSTR_DP("CBSTR LeftChars")
   DIM pbstr AS BSTR
   IF nChars = 0 THEN
      pbstr = SysAllocString(MID(*m_bstr, nStart))
   ELSE
      pbstr = SysAllocString(MID(*m_bstr, nStart, nChars))
   END IF
   RETURN pbstr
   ' // Don't free psbtr because it will be attached to the returned CBSTR
END FUNCTION
' ========================================================================================

' =====================================================================================
' Converts the string to a 32bit integer
' =====================================================================================
PRIVATE FUNCTION CBStr.ValLong () AS LONG
   RETURN .ValInt(*m_bstr)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CBStr.ValInt () AS LONG
   RETURN .ValInt(*m_bstr)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts the string to an unsigned 32bit integer
' =====================================================================================
PRIVATE FUNCTION CBStr.ValULong () AS ULONG
   RETURN .ValUInt(*m_bstr)
END FUNCTION
' =====================================================================================
PRIVATE FUNCTION CBStr.ValUInt () AS ULONG
   RETURN .ValUInt(*m_bstr)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts the string to a 64bit integer
' =====================================================================================
PRIVATE FUNCTION CBStr.ValLongInt () AS LONGINT
   RETURN .ValLng(*m_bstr)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts the string to an unsigned 64bit integer
' =====================================================================================
PRIVATE FUNCTION CBStr.ValULongInt () AS ULONGINT
   RETURN .ValULng(*m_bstr)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts the string to a floating point number (DOUBLE)
' =====================================================================================
PRIVATE FUNCTION CBStr.ValDouble () AS DOUBLE
   RETURN .VAL(*m_bstr)
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CBStr.Value () AS DOUBLE
   RETURN .VAL(*m_bstr)
END FUNCTION
' =====================================================================================

' ========================================================================================
' Converts the CBSTR to UTF8.
' ========================================================================================
PRIVATE PROPERTY CBStr.Utf8 () AS STRING
   CBSTR_DP("CBSTR Utf8 GET PROPERTY")
   DIM cbLen AS INTEGER, bstrLen AS LONG
   IF m_bstr = NULL THEN RETURN ""
   bstrLen = SysStringLen(m_bstr)
   DIM buffer AS STRING = STRING(bstrLen * 5 + 1, 0)
   PROPERTY = *cast(ZSTRING PTR, WCharToUTF(1, cast(WSTRING PTR, m_bstr), bstrLen, STRPTR(buffer), @cbLen))
END PROPERTY
' ========================================================================================

' ========================================================================================
' Converts UTF8 to unicode and assigns it to the CWSTR.
' ========================================================================================
PRIVATE PROPERTY CBStr.Utf8 (BYREF utf8String AS STRING)
   CBSTR_DP("CWSTR Utf8 SET PROPERTY")
   IF m_bstr THEN SysFreeString(m_bstr)
   DIM dwLen AS DWORD = MultiByteToWideChar(CP_UTF8, 0, STRPTR(utf8String), LEN(utf8String), NULL, 0)
   IF dwLen THEN
      m_bstr = SysAllocString(.WSTR(SPACE(dwLen)))
      MultiByteToWideChar(CP_UTF8, 0, STRPTR(utf8String), LEN(utf8String), m_bstr, dwLen * 2)
   ELSE
      m_bstr = SysAllocString("")
   END IF
END PROPERTY
' ========================================================================================

END NAMESPACE

' ########################################################################################
'                               *** GLOBAL OPERATORS ***
' ########################################################################################

' // Outside a namespace because they are global
using Afx

' ========================================================================================
PRIVATE OPERATOR & (BYREF cws1 AS CWSTR, BYREF cws2 AS CWSTR) AS CWSTR
   DIM cwsRes AS CWSTR = cws1
   cwsRes.Add(cws2)
   OPERATOR = cwsRes
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Left OVERLOAD (BYREF cws AS CWSTR, BYVAL nChars AS INTEGER) AS CWSTR
   RETURN LEFT(*cast(WSTRING PTR, cws.m_pBuffer), nChars)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Right OVERLOAD (BYREF cws AS CWSTR, BYVAL nChars AS INTEGER) AS CWSTR
   RETURN RIGHT(*cast(WSTRING PTR, cws.m_pBuffer), nChars)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Val OVERLOAD (BYREF cws AS CWSTR) AS DOUBLE
   RETURN .VAL(*cast(WSTRING PTR, cws.m_pBuffer))
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE OPERATOR & (BYREF cbs1 AS CBSTR, BYREF cbs2 AS CBSTR) AS CBSTR
   OPERATOR = cbs1 + cbs2
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Left OVERLOAD (BYREF cbs AS CBSTR, BYVAL nChars AS INTEGER) AS CBSTR
   RETURN LEFT(**cbs, nChars)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Right OVERLOAD (BYREF cbs AS CBSTR, BYVAL nChars AS INTEGER) AS CBSTR
   RETURN RIGHT(**cbs, nChars)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Val OVERLOAD (BYREF cbs AS CBSTR) AS DOUBLE
   RETURN .VAL(**cbs)
END FUNCTION
' ========================================================================================

' ########################################################################################
'                                *** HELPER FUNCTIONS ***
' ########################################################################################

' ========================================================================================
' qsort CWstr comparison function
' ========================================================================================
PRIVATE FUNCTION AfxCWstrArrayCompare CDECL (BYVAL a AS CWSTR PTR, BYVAL b AS CWSTR PTR) AS LONG
   FUNCTION = wcscmp(cast(WSTRING PTR, a->m_pBuffer), cast(WSTRING PTR, b->m_pBuffer))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Reverse qsort CWstr comparison function
' ========================================================================================
PRIVATE FUNCTION AfxCWStrArrayReverseCompare CDECL (BYVAL a AS CWSTR PTR, BYVAL b AS CWSTR PTR) AS LONG
   DIM r AS LONG = wcscmp(cast(WSTRING PTR, a->m_pBuffer), cast(WSTRING PTR, b->m_pBuffer))
   IF r = 1 THEN r = -1 ELSE IF r = -1 THEN r = 1
   RETURN r
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sorts a one-dimensional CWSTR array calling the C qsort function.
' Parameters:
' - rgwstr : Start of target array.
' - numElm : Number of elements in the array.
' - bAscend: TRUE for sorting in ascending order; FALSE for sorting in descending order.
' Example:
' DIM rg(1 TO 10) AS CWSTR
' FOR i AS LONG = 1 TO 10
'    rg(i) = "string " & i
' NEXT
' FOR i AS LONG = 1 TO 10
'   print rg(i)
' NEXT
' print "---- after sorting ----"
' AfxCWstrSort @rg(1), 10, TRUE
' FOR i AS LONG = 1 TO 10
'    print rg(i)
' NEXT
' ========================================================================================
PRIVATE SUB AfxCWstrSort (BYREF rgwstr AS ANY PTR, BYVAL numElm AS LONG, BYVAL bAscend AS BOOLEAN = TRUE)
   IF rgwstr = NULL OR numElm < 2 THEN EXIT SUB
   IF bAscend THEN
      qsort rgwstr, numElm, SIZEOF(CWSTR), CPTR(ANY PTR, @AfxCWstrArrayCompare)
   ELSE
      qsort rgwstr, numElm, SIZEOF(CWSTR) , CPTR(ANY PTR, @AfxCWStrArrayReverseCompare)
   END IF
END SUB
' ========================================================================================
