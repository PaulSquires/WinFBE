' ########################################################################################
' Microsoft Windows
' File: CBStr.inc
' Contents: Windows wrapper functions.
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

#include once "windows.bi"
#include once "win/ole2.bi"
#include once "win/shlwapi.bi"

#ifndef AFX_BSTR
   #define AFX_BSTR WSTRING PTR
#endif

NAMESPACE Afx.CBStrClass

' ========================================================================================
' CBStr - OLE strings class
' ========================================================================================
TYPE CBStr

   Private:
      m_bstr AS AFX_BSTR
      m_CodePage AS LONG

   Public:
      DECLARE CONSTRUCTOR
      DECLARE CONSTRUCTOR (BYVAL nCopePage AS LONG)
      DECLARE CONSTRUCTOR (BYREF wszStr AS CONST WSTRING = "")
      DECLARE CONSTRUCTOR (BYREF ansiStr AS STRING = "", BYVAL nCodePage AS LONG = 0)
      DECLARE CONSTRUCTOR (BYREF pCBStr AS CBStr)
      DECLARE CONSTRUCTOR (BYREF bstrHandle AS AFX_BSTR = NULL)
      DECLARE DESTRUCTOR
      DECLARE OPERATOR Let (BYREF ansiStr AS STRING)
      DECLARE OPERATOR Let (BYREF wszStr AS CONST WSTRING)
      DECLARE OPERATOR Let (BYREF pCBStr AS CBStr)
      DECLARE OPERATOR Let (BYREF bstrHandle AS AFX_BSTR)
      DECLARE OPERATOR += (BYREF wszStr AS CONST WSTRING)
      DECLARE OPERATOR += (BYREF pCBStr AS CBStr)
      DECLARE OPERATOR &= (BYREF wszStr AS CONST WSTRING)
      DECLARE OPERATOR &= (BYREF pCBStr AS CBStr)
      DECLARE OPERATOR @ () AS AFX_BSTR PTR
      DECLARE OPERATOR CAST () BYREF AS WSTRING
      DECLARE OPERATOR CAST () AS ANY PTR
      DECLARE PROPERTY Handle () AS AFX_BSTR
      DECLARE PROPERTY CodePage () AS LONG
      DECLARE PROPERTY CodePage (BYVAL nCodePage AS LONG)
      DECLARE SUB Append (BYREF wszStr AS CONST WSTRING)

END TYPE
' ========================================================================================

' ========================================================================================
' CBStr class constructors
' ========================================================================================
CONSTRUCTOR CBStr
   m_bstr = SysAllocString("")
END CONSTRUCTOR
' ========================================================================================
CONSTRUCTOR CBStr (BYVAL nCodePage AS LONG)
   m_CodePage = nCodePage
END CONSTRUCTOR
' ========================================================================================
CONSTRUCTOR CBStr (BYREF wszStr AS CONST WSTRING = "")
   m_bstr = SysAllocString(wszStr)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
CONSTRUCTOR CBStr (BYREF ansiStr AS STRING = "", BYVAL nCodePage AS LONG = 0)
   m_bstr = SysAllocString(WSTR(ansiStr))
   IF nCodePage <> 0 THEN
      MultiByteToWideChar(nCodePage, MB_PRECOMPOSED, STRPTR(ansiStr), -1, m_bstr, LEN(ansiStr) * 2)
   ELSEIF m_CodePage <> 0 THEN
      MultiByteToWideChar(m_CodePage, MB_PRECOMPOSED, STRPTR(ansiStr), -1, m_bstr, LEN(ansiStr) * 2)
   END IF
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
CONSTRUCTOR CBStr (BYREF pCBStr AS CBStr)
   m_bstr = SysAllocString(*pCBStr.Handle)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
CONSTRUCTOR CBStr (BYREF bstrHandle AS AFX_BSTR = NULL)
   IF bstrHandle = NULL THEN
      m_bstr = SysAllocString("")
   ELSE
      ' Detect if the passed handle is an OLE string
      ' If it is an OLE string it must have a descriptor; otherwise, don't
      ' Get the length looking at the descriptor
      DIM res AS INTEGER = PEEK(DWORD, CAST(ANY PTR, bstrHandle) - 4) \ 2
      ' If the retrieved length if the same that the returned by LEN, then it must be an OLE string
      IF res = .LEN(*bstrHandle) THEN
         ' Attach the passed handle to the class
         m_bstr = bstrHandle
      ELSE
         ' Allocate an OLE string with the contents of the string pointer by bstrHandle
         m_bstr = SysAllocString(*bstrHandle)
      END IF
   END IF
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CBStr class destructor
' ========================================================================================
DESTRUCTOR CBStr
   IF m_bstr THEN SysFreeString m_bstr
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns a pointer to the BSTR
' ========================================================================================
OPERATOR CBStr.CAST () BYREF AS WSTRING
   OPERATOR =  *CAST(WSTRING PTR, m_bstr)
END OPERATOR
' ========================================================================================
' ========================================================================================
OPERATOR CBStr.CAST () AS ANY PTR
   OPERATOR =  CAST(ANY PTR, m_bstr)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns new text to the BSTR
' ========================================================================================
OPERATOR CBStr.Let (BYREF wszStr AS CONST WSTRING)
   IF m_bstr THEN SysFreeString(m_bstr)
   m_bstr = SysAllocString(wszStr)
END OPERATOR
' ========================================================================================
' ========================================================================================
OPERATOR CBStr.Let (BYREF ansiStr AS STRING)
   IF m_bstr THEN SysFreeString(m_bstr)
   m_bstr = SysAllocString(WSTR(ansiStr))
   IF m_CodePage <> 0 THEN MultiByteToWideChar(m_CodePage, MB_PRECOMPOSED, STRPTR(ansiStr), -1, m_bstr, LEN(ansiStr) * 2)
END OPERATOR
' ========================================================================================
' ========================================================================================
OPERATOR CBStr.Let (BYREF pCBStr AS CBStr)
   IF m_bstr THEN SysFreeString(m_bstr)
   m_bstr = SysAllocString(*pCBStr.Handle)
END OPERATOR
' ========================================================================================
' ========================================================================================
OPERATOR CBStr.Let (BYREF bstrHandle AS AFX_BSTR)
   IF bstrHandle = NULL THEN EXIT OPERATOR
   ' Free the current OLE string
   IF m_bstr THEN SysFreeString(m_bstr)
   ' Detect if the passed handle is an OLE string
   ' If it is an OLE string it must have a descriptor; otherwise, don't
   ' Get the length looking at the descriptor
   DIM res AS INTEGER = PEEK(DWORD, CAST(ANY PTR, bstrHandle) - 4) \ 2
   ' If the retrieved length if the same that the returned by LEN, then it must be an OLE string
   IF res = .LEN(*bstrHandle) THEN
      ' Attach the passed handle to the class
      m_bstr = bstrHandle
   ELSE
      ' Allocate an OLE string with the contents of the string pointer by bstrHandle
      m_bstr = SysAllocString(*bstrHandle)
   END IF
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns the handle of the BSTR
' ========================================================================================
PROPERTY CBStr.Handle () AS AFX_BSTR
   PROPERTY = m_bstr
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets/sets the code page used to ansi to unicode translations
' ========================================================================================
PROPERTY CBStr.CodePage () AS LONG
   PROPERTY = m_CodePage
END PROPERTY
' ========================================================================================
PROPERTY CBStr.CodePage (BYVAL nCodePage AS LONG)
   m_CodePage = nCodePage
END PROPERTY
' ========================================================================================

' ========================================================================================
' Appends a string to the BSTR. The string can be a literal or a FB STRING or WSTRING variable.
' To append another BSTR:
' DIM bs1 AS CBStr = "1st string"
' DIM bs2 AS CBStr = "2nd string"
' bs1.Append bs2
' -or-
' bs1.Append *bs2.Handle
' -or-
' bs1.Append **bs2
' ========================================================================================
SUB CBStr.Append (BYREF wszStr AS CONST WSTRING)
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
' Returns the address of the BSTR
' ========================================================================================
OPERATOR CBStr.@ () AS AFX_BSTR PTR
   OPERATOR = @m_bstr
END OPERATOR
' ========================================================================================

' ========================================================================================
' Appends a string to the BSTR. The string can be a literal or a FB STRING or WSTRING variable.
' ========================================================================================
OPERATOR CBStr.+= (BYREF wszStr AS CONST WSTRING)
   this.Append(wszStr)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Appends a BSTR to the BSTR.
' ========================================================================================
OPERATOR CBStr.+= (BYREF pCBStr AS CBStr)
   this.Append(*pCBStr.Handle)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Appends a string to the BSTR. The string can be a literal or a FB STRING or WSTRING variable.
' ========================================================================================
OPERATOR CBStr.&= (BYREF wszStr AS CONST WSTRING)
   this.Append(wszStr)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Appends a BSTR to the BSTR.
' ========================================================================================
OPERATOR CBStr.&= (BYREF pCBStr AS CBStr)
   this.Append(*pCBStr.Handle)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns vbTrue (-1) if the two BSTRings are equal; FALSE, otherwise.
' ========================================================================================
OPERATOR = (BYREF pCBStr1 AS CBStr, BYREF pCBStr2 AS CBStr) AS INTEGER
   OPERATOR = StrCmpW(*pCBStr1.Handle, *pCBStr2.Handle) = 0
END OPERATOR
' ========================================================================================
' ========================================================================================
OPERATOR = (BYREF pCBStr AS CBStr, BYREF wszStr AS CONST WSTRING) AS INTEGER
   OPERATOR = StrCmpW(*pCBStr.Handle, wszStr) = 0
END OPERATOR
' ========================================================================================
OPERATOR = (BYREF wszStr AS CONST WSTRING, BYREF pCBStr AS CBStr) AS INTEGER
   OPERATOR = StrCmpW(wszStr, *pCBStr.Handle) = 0
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns vbTrue (-1) if the two BSTRings are not equal; FALSE, otherwise.
' ========================================================================================
OPERATOR <> (BYREF pCBStr1 AS CBStr, BYREF pCBStr2 AS CBStr) AS INTEGER
   OPERATOR = StrCmpW(*pCBStr1.Handle, *pCBStr2.Handle) <> 0
END OPERATOR
' ========================================================================================
' ========================================================================================
OPERATOR <> (BYREF pCBStr AS CBStr, BYREF wszStr AS CONST WSTRING) AS INTEGER
   OPERATOR = StrCmpW(*pCBStr.Handle, wszStr) <> 0
END OPERATOR
' ========================================================================================
OPERATOR <> (BYREF wszStr AS CONST WSTRING, BYREF pCBStr AS CBStr) AS INTEGER
   OPERATOR = StrCmpW(wszStr, *pCBStr.Handle) <> 0
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns vbTrue (-1) if the the first BSTR is greater than the 2nd BSTR; FALSE, otherwise.
' ========================================================================================
OPERATOR > (BYREF pCBStr1 AS CBStr, BYREF pCBStr2 AS CBStr) AS INTEGER
   OPERATOR = StrCmpW(*pCBStr1.Handle, *pCBStr2.Handle) > 0
END OPERATOR
' ========================================================================================
' ========================================================================================
OPERATOR > (BYREF pCBStr AS CBStr, BYREF wszStr AS CONST WSTRING) AS INTEGER
   OPERATOR = StrCmpW(*pCBStr.Handle, wszStr) > 0
END OPERATOR
' ========================================================================================
OPERATOR > (BYREF wszStr AS CONST WSTRING, BYREF pCBStr AS CBStr) AS INTEGER
   OPERATOR = StrCmpW(wszStr, *pCBStr.Handle) > 0
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns vbTrue (-1) if the the first BSTR is less than the 2nd BSTR; FALSE, otherwise.
' ========================================================================================
OPERATOR < (BYREF pCBStr1 AS CBStr, BYREF pCBStr2 AS CBStr) AS INTEGER
   OPERATOR = StrCmpW(*pCBStr1.Handle, *pCBStr2.Handle) < 0
END OPERATOR
' ========================================================================================
' ========================================================================================
OPERATOR < (BYREF pCBStr AS CBStr, BYREF wszStr AS CONST WSTRING) AS INTEGER
   OPERATOR = StrCmpW(*pCBStr.Handle, wszStr) < 0
END OPERATOR
' ========================================================================================
OPERATOR < (BYREF wszStr AS CONST WSTRING, BYREF pCBStr AS CBStr) AS INTEGER
   OPERATOR = StrCmpW(wszStr, *pCBStr.Handle) < 0
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns vbTrue (-1) if the the first BSTR is greater or equal than the 2nd BSTR; FALSE, otherwise.
' ========================================================================================
OPERATOR >= (BYREF pCBStr1 AS CBStr, BYREF pCBStr2 AS CBStr) AS INTEGER
   DIM nResult AS LONG
   nResult = StrCmpW(*pCBStr1.Handle, *pCBStr2.Handle)
   IF nResult > 0 OR nResult = 0 THEN OPERATOR = -1 ELSE OPERATOR = 0
END OPERATOR
' ========================================================================================
' ========================================================================================
OPERATOR >= (BYREF pCBStr AS CBStr, BYREF wszStr AS CONST WSTRING) AS INTEGER
   DIM nResult AS LONG
   nResult = StrCmpW(*pCBStr.Handle, wszStr)
   IF nResult > 0 OR nResult = 0 THEN OPERATOR = -1 ELSE OPERATOR = 0
END OPERATOR
' ========================================================================================
OPERATOR >= (BYREF wszStr AS CONST WSTRING, BYREF pCBStr AS CBStr) AS INTEGER
   DIM nResult AS LONG
   nResult = StrCmpW(wszStr, *pCBStr.Handle)
   IF nResult > 0 OR nResult = 0 THEN OPERATOR = -1 ELSE OPERATOR = 0
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns vbTrue (-1) if the the first BSTR is less or equal than the 2nd BSTR; FALSE, otherwise.
' ========================================================================================
OPERATOR <= (BYREF pCBStr1 AS CBStr, BYREF pCBStr2 AS CBStr) AS INTEGER
   DIM nResult AS LONG
   nResult = StrCmpW(*pCBStr1.Handle, *pCBStr2.Handle)
   IF nResult < 0 OR nResult = 0 THEN OPERATOR = -1 ELSE OPERATOR = 0
END OPERATOR
' ========================================================================================
' ========================================================================================
OPERATOR <= (BYREF pCBStr AS CBStr, BYREF wszStr AS CONST WSTRING) AS INTEGER
   DIM nResult AS LONG
   nResult = StrCmpW(*pCBStr.Handle, wszStr)
   IF nResult < 0 OR nResult = 0 THEN OPERATOR = -1 ELSE OPERATOR = 0
END OPERATOR
' ========================================================================================
OPERATOR <= (BYREF wszStr AS CONST WSTRING, BYREF pCBStr AS CBStr) AS INTEGER
   DIM nResult AS LONG
   nResult = StrCmpW(wszStr, *pCBStr.Handle)
   IF nResult < 0 OR nResult = 0 THEN OPERATOR = -1 ELSE OPERATOR = 0
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns the address of the BSTR handle.
' Usage:
' DIM bs AS CBStr = "Test string"
' PRINT bs
' -or-
' PRINT **bs (notice the double indirection)
' ========================================================================================
OPERATOR * (BYREF pCBStr AS CBStr) AS AFX_BSTR
   OPERATOR = pCBStr.Handle
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns the length of the BSTR in characters
' ========================================================================================
OPERATOR Len (BYREF pCBStr AS CBStr) AS INTEGER
   OPERATOR = SysStringLen(pCBStr.Handle)
END OPERATOR
' ========================================================================================

END NAMESPACE

