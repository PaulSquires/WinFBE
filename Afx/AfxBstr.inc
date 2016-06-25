' ########################################################################################
' Microsoft Windows
' File: AfxBstr.inc
' Contents: Windows wrapper functions.
' Copyright (c) 2015 José Roca
' Compiler: FreeBasic
' All Rights Reserved.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

#include once "windows.bi"
#include once "win/ole2.bi"

' ========================================================================================
' AfxBstr - OLE strings class
' ========================================================================================
TYPE AfxBstr

   Private:
      m_bstr AS BSTR

   Public:
      DECLARE CONSTRUCTOR (BYREF wszStr AS CONST WSTRING = "")
      DECLARE DESTRUCTOR
      DECLARE PROPERTY Attach (BYVAL bstrHandle AS BSTR)
      DECLARE PROPERTY Assign (BYREF wszStr AS CONST WSTRING)
      DECLARE PROPERTY Handle () AS BSTR
      DECLARE FUNCTION Len () AS LONG
      DECLARE FUNCTION ToAnsi () AS STRING
      DECLARE PROPERTY FromAnsi (BYREF ansiText AS CONST STRING)
      DECLARE SUB MakeUpper
      DECLARE SUB MakeLower

END TYPE
' ========================================================================================

' ========================================================================================
' AfxBstr class constructor
' ========================================================================================
CONSTRUCTOR AfxBstr (BYREF wszBstr AS CONST WSTRING = "")
   this.Assign = wszBstr
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' AfxBstr class destructor
' ========================================================================================
DESTRUCTOR AfxBstr
   IF m_bstr THEN SysFreeString m_bstr
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Attaches a BSTR
' ========================================================================================
PROPERTY AfxBstr.Attach (BYVAL bstrHandle AS BSTR)
   IF m_bstr THEN SysFreeString(m_bstr)
   m_bstr = bstrHandle
END PROPERTY
' ========================================================================================

' ========================================================================================
' Assigns new text to the BSTR
' ========================================================================================
PROPERTY AfxBstr.Assign (BYREF wszStr AS CONST WSTRING)
   IF m_bstr THEN SysFreeString(m_bstr)
   m_bstr = SysAllocString(wszStr)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Returns the handle of the BSTR
' ========================================================================================
PROPERTY AfxBstr.Handle () AS BSTR
   PROPERTY = m_bstr
END PROPERTY
' ========================================================================================

' ========================================================================================
' Returns the length of the BSTR in characters
' ========================================================================================
FUNCTION AfxBstr.Len () AS LONG
   IF m_bstr THEN FUNCTION = SysStringLen(m_bstr)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the text of the string converted to ansi
' ========================================================================================
FUNCTION AfxBstr.ToAnsi () AS STRING
   IF m_bstr THEN FUNCTION = STR(*m_bstr)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Assigns ansi text to the BSTR after being converted to unicode.
' ========================================================================================
PROPERTY AfxBstr.FromAnsi (BYREF ansiText AS CONST STRING)
   this.Assign = WSTR(ansiText)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Converts the string to upper case
' ========================================================================================
SUB AfxBstr.MakeUpper
   *m_bstr = UCASE(*m_bstr)
END SUB
' ========================================================================================

' ========================================================================================
' Converts the string to lower case
' ========================================================================================
SUB AfxBstr.MakeLower
   *m_bstr = LCASE(*m_bstr)
END SUB
' ========================================================================================
