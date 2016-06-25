' ########################################################################################
' Microsoft Windows
' File: AfxStr.inc
' Contents: String wrapper functions.
' Copyright (c) 2015 José Roca
' Compiler: FreeBasic
' All Rights Reserved.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

#include once "windows.bi"

' ========================================================================================
' Return the portion of a string following the first occurrence of a character.
' ========================================================================================
FUNCTION AfxRemain OVERLOAD (BYREF MainStr AS CONST STRING, BYREF MatchStr AS CONST STRING, BYVAL Start AS INTEGER = 1) AS STRING
   DIM nPos AS INTEGER
   IF LEN(MainStr) = 0 OR LEN(MatchStr) = 0 THEN EXIT FUNCTION
   nPos = INSTR(Start, MainStr, MatchStr)
   IF nPos = 0 THEN EXIT FUNCTION
   FUNCTION = MID(MainStr, nPos + LEN(MatchStr))
END FUNCTION
' ========================================================================================
' ====================================================s====================================
'FUNCTION AfxRemain OVERLOAD (BYREF MainStr AS CONST WSTRING, BYREF MatchStr AS CONST WSTRING, BYVAL Start AS INTEGER = 1) AS WSTRING PTR
'   DIM nPos AS INTEGER
'   IF LEN(MainStr) = 0 OR LEN(MatchStr) = 0 THEN EXIT FUNCTION
'   nPos = INSTR(Start, MainStr, MatchStr)
'   IF nPos = 0 THEN EXIT FUNCTION
'   FUNCTION = CPTR(WSTRING PTR, MID(MainStr, nPos + LEN(MatchStr)))
'END FUNCTION
' ========================================================================================


' ========================================================================================
' Converts an unicode string to an UTF8 encoded string.
' ========================================================================================
Function AfxWideCharToUtf8 (ByVal lpWideCharStr As LPCWCH) As String
   Dim hr As Long, buffer As String
   ' // Get the number of bytes required for the buffer for ANSI string
   hr = WideCharToMultiByte(CP_UTF8, 0, lpWideCharStr, Lstrlenw(lpWideCharStr), Null, 0, Null, Null)
   If hr = 0 Then Exit Function
   ' // Allocate a string of the needed size
   buffer = Space(hr)
   ' // Convert to UTF8 and return the result
   hr = WideCharToMultiByte(CP_UTF8, 0, lpWideCharStr, Lstrlenw(lpWideCharStr), Strptr(buffer), Len(buffer), Null, Null)
   If hr = 0 Then Exit Function
   Function = buffer
End Function
' ========================================================================================

' ========================================================================================
' Converts an unicode string to an UTF8 encoded string.
' ========================================================================================
Function AfxUtf8ToAnsi (ByVal strUtf8 As String) As String
   Dim hr As Long, buffer As String
   ' // Get the number of bytes required for the buffer for WIDE string
   hr = MultiByteToWideChar(CP_UTF8, 0, Strptr(strUtf8), Len(strUtf8), Null, 0)
   If hr = 0 Then Exit Function
   ' // Convert to Unicode
   buffer = Space(hr * 2)
   Dim bstrHandle As BSTR = SysAllocStringByteLen(Strptr(buffer), Len(buffer))
   hr = MultiByteToWideChar(CP_UTF8, 0, Strptr(strUtf8), Len(strUtf8), bstrHandle, hr * 2)
   ' // Convert to to ansi
   If hr Then
      buffer = Space(hr)
      hr = WideCharToMultiByte(CP_ACP, 0, bstrHandle, SysStringLen(bstrHandle), Strptr(buffer), Len(buffer), Null, Null)
   End If
   If bstrHandle Then SysFreeString bstrHandle
   Function = buffer
End Function
' ========================================================================================

' ========================================================================================
' Converts an UTF8 encoded string to an unicode string.
' ========================================================================================
Function AfxUtf8ToWideChar (ByVal strUtf8 As String) As BSTR
   Dim hr As Long, buffer As String
   ' // Get the number of bytes required for the buffer for WIDE string
   hr = MultiByteToWideChar(CP_UTF8, 0, Strptr(strUtf8), Len(strUtf8), Null, 0)
   If hr = 0 Then Exit Function
   ' // Convert to Unicode
   buffer = Space(hr * 2)
   Dim bstrHandle As WString Ptr = SysAllocStringByteLen(Strptr(buffer), Len(buffer))
   hr = MultiByteToWideChar(CP_UTF8, 0, Strptr(strUtf8), Len(strUtf8), bstrHandle, hr * 2)
   Function = bstrHandle
'   // Alternate way
'   DIM pwBuffer AS WSTRING PTR
'   pwBuffer = CAllocate(hr * 2)
'   IF pwBuffer = NULL THEN EXIT FUNCTION
'   hr = MultiByteToWideChar(CP_UTF8, 0, STRPTR(strUtf8), LEN(strUtf8), pwBuffer, hr * 2)
'   FUNCTION = pwBuffer

End Function
' ========================================================================================

