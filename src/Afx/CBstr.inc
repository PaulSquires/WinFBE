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

#ifndef AFX_BSTR
   #define AFX_BSTR WSTRING PTR
#endif

NAMESPACE Afx.CBStrClass

' ========================================================================================
' CBStr - OLE strings class
' ========================================================================================
TYPE CBStr

   Public:
      m_bstr AS AFX_BSTR
      m_CodePage AS LONG

   Public:
      DECLARE CONSTRUCTOR
      DECLARE CONSTRUCTOR (BYVAL nCopePage AS LONG)
      DECLARE CONSTRUCTOR (BYREF wszStr AS CONST WSTRING = "")
      DECLARE CONSTRUCTOR (BYREF ansiStr AS STRING = "", BYVAL nCodePage AS LONG = 0)
      DECLARE CONSTRUCTOR (BYREF cbs AS CBStr)
      DECLARE CONSTRUCTOR (BYREF bstrHandle AS AFX_BSTR = NULL)
      DECLARE DESTRUCTOR
      DECLARE OPERATOR Let (BYREF ansiStr AS STRING)
      DECLARE OPERATOR Let (BYREF wszStr AS CONST WSTRING)
      DECLARE OPERATOR Let (BYREF cbs AS CBStr)
      DECLARE OPERATOR Let (BYREF bstrHandle AS AFX_BSTR)
      DECLARE OPERATOR CAST () BYREF AS WSTRING
      DECLARE OPERATOR CAST () AS ANY PTR
      DECLARE PROPERTY CodePage () AS LONG
      DECLARE PROPERTY CodePage (BYVAL nCodePage AS LONG)

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
CONSTRUCTOR CBStr (BYREF cbs AS CBStr)
   m_bstr = SysAllocString(cbs)
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
OPERATOR CBStr.Let (BYREF cbs AS CBStr)
   IF m_bstr THEN SysFreeString(m_bstr)
   m_bstr = SysAllocString(cbs)
END OPERATOR
' ========================================================================================
' ========================================================================================
OPERATOR CBStr.Let (BYREF bstrHandle AS AFX_BSTR)
   IF bstrHandle = NULL THEN EXIT OPERATOR
   ' Free the current OLE string
   IF m_bstr THEN SysFreeString(m_bstr)
   ' Detect if the passed handle is an OLE string.
   ' If it is an OLE string it must have a descriptor; otherwise, don't.
   ' Get the length in bytes looking at the descriptor and divide by 2 to get the number of
   ' unicode characters, that is the value returned by the FreeBASIC LEN operator.
   DIM res AS DWORD = PEEK(DWORD, CAST(ANY PTR, bstrHandle) - 4) \ 2
   ' If the retrieved length is the same that the returned by LEN, then it must be an OLE string
   IF res = .LEN(*bstrHandle) THEN
      ' Attach the passed handle to the class
      m_bstr = bstrHandle
   ELSE
      ' Allocate an OLE string with the contents of the string pointed by bstrHandle
      m_bstr = SysAllocString(*bstrHandle)
   END IF
END OPERATOR
' ========================================================================================

' ========================================================================================
' One * returns the value of the BSTR pointer.
' Two ** returns the adress of the start of the string data.
' Needed because LEFT and RIGHT (cbs) fail with an ambiguous call error.
' We have to use **cbs (notice the double indirection) with these functions.
' ========================================================================================
OPERATOR * (BYREF cbs AS CBStr) AS AFX_BSTR
   OPERATOR = cbs.m_bstr
END OPERATOR
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
' Returns the length of the BSTR in characters.
' Needed because FB's LEN operator does not work with BSTRs.
' ========================================================================================
OPERATOR Len (BYREF cbs AS CBStr) AS INTEGER
   OPERATOR = SysStringLen(cbs)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Gets/sets the code page used to ansi to unicode translations
' ========================================================================================
PROPERTY CBStr.CodePage () AS LONG
   PROPERTY = m_CodePage
END PROPERTY
' ========================================================================================
' ========================================================================================
PROPERTY CBStr.CodePage (BYVAL nCodePage AS LONG)
   m_CodePage = nCodePage
END PROPERTY
' ========================================================================================

END NAMESPACE


