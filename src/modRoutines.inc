'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2023 Paul Squires, PlanetSquires Software
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT any WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS for A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.

#include once "modRoutines.bi"
#include once "frmUserTools.bi"
#include once "frmOutput.bi"
#include once "frmOptionsCompiler.bi"
#include once "frmOptionsLocal.bi"
#include once "frmProjectOptions.bi"
#include once "frmMain.bi"


' ========================================================================================
' Determine if filename is a *.FRM form file
' ========================================================================================
function IsFormFilename( byref wszName as wstring ) as boolean
   dim as CWSTR ext = ucase(AfxStrPathName( "EXT", wszName ))
   if ext = "FRM" then 
      return true
   else
      return false   
   end if
end function


' ========================================================================================
' Enclose JSON string (both sides) in double quotes
' ========================================================================================
function qstr( byval wst as CWSTR ) as CWSTR
   if len(wst) = 0 then return wst
   dim as CWSTR stLeft = AfxStrParse(wst, 1, ":")
   dim as CWSTR stRight = AfxStrParse(wst, 2, ":")
   dim as CWSTR newSt = AfxStrWrap(stLeft) & ":" & AfxStrWrap(stRight)
   return newSt
end function

' ========================================================================================
' Enclose JSON string (left side) in double quotes
' This is used mostly for creating JSON numbers/booleans/arrays.
' ========================================================================================
function qnum( byval wst as CWSTR ) as CWSTR
   if len(wst) = 0 then return wst
   dim as CWSTR stLeft = AfxStrParse(wst, 1, ":")
   dim as CWSTR stRight = AfxStrParse(wst, 2, ":")
   dim as CWSTR newSt = AfxStrWrap(stLeft) & ":" & stRight
   return newSt
end function


' ========================================================================================
' Get the width of the text in unscaled pixels because we feed this value
' to the pWindow create control function that will then scale up the value.
' ========================================================================================
function getTextWidth( _
            byval hWnd as HWND, _
            byref wszText as WSTRING, _
            byval _hFont as HFONT, _
            byval nPadding as long _
            ) as long

   dim size AS SIZEL
   dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
   
   dim as HDC hDC = GetDC(hWnd)
   if _hFont then SelectObject(hDC, _hFont)
   GetTextExtentPoint32( hDC, @wszText, len(wszText), @size )  'get text size
   ReleaseDC hWnd, hDC
   
   function = pWindow->UnScaleX(size.cx) + pWindow->ScaleX(nPadding)  ' add enough padding

end function

' ========================================================================================
' Determine if the mouse cursor is currently over the incoming window
' ========================================================================================
function isMouseOverWindow( byval hChild as HWND ) as boolean
   dim as POINT pt
   GetCursorPos( @pt )
   'dim as RECT rc = AfxGetWindowRect(hChild)
   'if PtInRect( @rc, pt ) then return true
   if WindowFromPoint( pt ) = hChild then return true
end function

' ========================================================================================
' Determine if the mouse cursor is currently over client RECT area
' ========================================================================================
function isMouseOverRECT( byval hWin as HWND, byval rc as RECT ) as boolean
   dim as POINT pt
   GetCursorPos( @pt )
   MapWindowPoints( hWin, HWND_DESKTOP, cast( POINT ptr, @rc), 2 )
   if PtInRect( @rc, pt ) then return true
end function

' ========================================================================================
' Return the pixel width of the incoming text string
' ========================================================================================
function GetTextWidthPixels( byval hWin as HWND, byref wszText as WString ) as Long
   dim as HDC _hdc = GetDC(hWin)
   dim as SIZE _size 
   dim as HFONT oldfont = SelectObject( _hdc, AfxGetWindowFont(hWin) )
   GetTextExtentPoint32( _hdc, wszText, len(wszText), @_size)
   SelectObject( _hdc, oldfont )
   ReleaseDC( hWin, _hdc )
   function = _size.cx 
end function


' ========================================================================================
' Calculate the Julian date for today's date. Used by CheckForUpdates.
' ========================================================================================
function JulianDateNow() as long
   function = AfxGregorianToJulian( AfxSystemDay, AfxSystemMonth, AfxSystemYear )
end function


' ========================================================================================
' Convert WinFBE version to whole number for comparison to file versions being loaded.
' ========================================================================================
function ConvertWinFBEversion( byref wszVersion as wstring ) as long
   dim as CWSTR wszPart1, wszPart2, wszPart3
   
   wszPart1 = AfxStrLSet( AfxStrParse(wszVersion, 1, "."), 3, "0" )
   wszPart2 = AfxStrLSet( AfxStrParse(wszVersion, 2, "."), 3, "0" )
   wszPart3 = AfxStrRSet( AfxStrParse(wszVersion, 3, "."), 3, "0" )
   
   function = val(wszPart1 + wszPart2 + wszPart3) 
end function


' ========================================================================================
' Disable all modeless windows belonging to frmMain so that the popup modal is truly modal.
' ========================================================================================
function DisableAllModeless() as long
   ' No need to enable/disable the modeless Help form.
   if IsWindowVisible(HWND_FRMFINDREPLACE)  then EnableWindow(HWND_FRMFINDREPLACE, false)
   if IsWindowVisible(HWND_FRMVDTOOLBOX)    then EnableWindow(HWND_FRMVDTOOLBOX, false)
   EnableWindow(HWND_FRMMAIN, false)
   function = 0
end function


' ========================================================================================
' Enable all modeless windows belonging to frmMain.
' ========================================================================================
function EnableAllModeless() as long
   ' No need to enable/disable the modeless Help form.
   if IsWindowVisible(HWND_FRMFINDREPLACE)  then EnableWindow(HWND_FRMFINDREPLACE, true)
   if IsWindowVisible(HWND_FRMVDTOOLBOX)    then EnableWindow(HWND_FRMVDTOOLBOX, true)
   EnableWindow(HWND_FRMMAIN, true)
   function = 0
end function


' ========================================================================================
' Return temporary file name
' ========================================================================================
FUNCTION GetTemporaryFilename( _
            byref wszFolder as wstring, _
            BYREF wszExtension AS wSTRING _
            ) AS string

   dim wszTempFilename as wstring * MAX_PATH
   if GetTempFileName( @wszFolder, "TMP", 0, @wszTempFilename ) then
      ' Delete the temp file that gets created b/c we will create it ourselves based on the 
      ' returned filename.
      AfxDeleteFile( wszTempFilename )
      IF LEN(wszExtension) THEN 
         wszTempFilename = LEFT(wszTempFilename, LEN(wszTempFilename) -  3) & wszExtension
      end if
   end if      
   function = wszTempFilename
END FUNCTION


' ========================================================================================
' Replace a string in a combobox
' ========================================================================================
FUNCTION ComboBox_ReplaceString( _
            BYVAL hComboBox AS HWND, _
            BYVAL index AS LONG, _
            BYVAL pwszNewText AS WSTRING PTR, _
            BYVAL pNewData AS LONG_PTR = 0 _
            ) AS LONG
   ' Delete the string
   DIM lRes AS LRESULT = SendMessage(hComboBox, CB_DELETESTRING, index, 0)
   IF lRes = LB_ERR THEN RETURN lRes
   ' Insert the new string
   index = SendMessage(hComboBox, CB_INSERTSTRING, index, CAST(LPARAM, pwszNewText))
   IF index = LB_ERR OR index = LB_ERRSPACE THEN Return index
   lRes = SendMessage(hComboBox, CB_SETITEMDATA, index, CAST(LPARAM, pNewData))
   IF lRes = LB_ERR THEN Return lRes
   FUNCTION = SendMessage(hComboBox, CB_SETCURSEL, index, 0)
END FUNCTION


' ========================================================================================
' Get the Scintilla value for a character sets
' ========================================================================================
Function GetFontCharSetID(ByREF wzCharsetName As CWSTR ) As Long

   If Len(wzCharsetName) = 0 Then Return SC_CHARSET_DEFAULT
   
   Select Case wzCharsetName
      Case "Default"       : Function = SC_CHARSET_DEFAULT
      Case "Ansi"          : Function = SC_CHARSET_ANSI
      Case "Arabic"        : Function = SC_CHARSET_ARABIC
      Case "Baltic"        : Function = SC_CHARSET_BALTIC
      Case "Chinese Big 5" : Function = SC_CHARSET_CHINESEBIG5
      Case "East Europe"   : Function = SC_CHARSET_EASTEUROPE
      Case "GB 2312"       : Function = SC_CHARSET_GB2312
      Case "Greek"         : Function = SC_CHARSET_GREEK
      Case "Hangul"        : Function = SC_CHARSET_HANGUL
      Case "Hebrew"        : Function = SC_CHARSET_HEBREW
      Case "Johab"         : Function = SC_CHARSET_JOHAB
      Case "Mac"           : Function = SC_CHARSET_MAC
      Case "OEM"           : Function = SC_CHARSET_OEM
      Case "Russian"       : Function = SC_CHARSET_RUSSIAN
      Case "Shiftjis"      : Function = SC_CHARSET_SHIFTJIS
      Case "Symbol"        : Function = SC_CHARSET_SYMBOL
      Case "Thai"          : Function = SC_CHARSET_THAI
      Case "Turkish"       : Function = SC_CHARSET_TURKISH
      Case "Vietnamese"    : Function = SC_CHARSET_VIETNAMESE
   End Select

End Function


' ========================================================================================
' Remove duplicate spaces from the incoming line.
' ========================================================================================
function RemoveDuplicateSpaces( byref sText as const string) as string
   dim as string st = sText
   do until instr(st, "  ") = 0
      st = AfxStrReplace(st, "  ", " ")
   loop   
   function = st
end function


' ========================================================================================
' Convert incoming text to proper case based on config setting. Used for autocomplete.
' ========================================================================================
function ConvertCase( byval sText as string) as string

   Select Case gConfig.KeywordCase
      Case 0:  return lcase(sText)
      Case 1:  return ucase(sText)
      Case 2   ' Mixed case
         ' Loop through each character. If the previous character was an alphabet letter
         ' then make the character lowercase otherwise make it uppercase.
         Dim As String sChar, sPrevChar
         For i As Long = 1 To Len(sText)
            sChar = Mid(sText, i, 1)
            sPrevChar = Mid(sText, i-1, 1)
            If (sPrevChar = " ") OrElse (sPrevChar = "") Then
               Mid(sText, i, 1) = Ucase(sChar)
            Else   
               Mid(sText, i, 1) = LCase(sChar)
            End If   
         Next
         Return sText   
   End Select            
end function


' ========================================================================================
' Maps UTF-8 string to Ansi string.
' ========================================================================================
FUNCTION Utf8ToAscii(byref strUtf8 AS STRING) AS STRING

   dim i AS LONG                ' // Loop counter
   dim strAscii AS STRING       ' // Ascii string
   dim idx AS LONG              ' // Position in the string
   dim c AS LONG                ' // ASCII code
   dim b2 AS LONG               ' // Second byte
   dim fSkipChar AS boolean     ' // Flag

   IF LEN(strUtf8) = 0 THEN EXIT FUNCTION
   
   ' // The maximum length of the translated string will be
   ' // the same as the length of the original string.
   ' // We are pre-allocating the buffer for faster operation
   ' // than concatenating each character one by one.
   strAscii = SPACE(LEN(strUtf8))

   ' // Intialize index position in the string buffer
   ' // used to store the converted Ascii string
   idx = 1
   
   ' // Examine the contents of each character in the UTF-8 encoded string
   FOR i = 1 TO LEN(strUtf8)
      ' // If fSkipChar is set we have to skip this character
      IF fSkipChar THEN
         fSkipChar = 0
         continue FOR
      END IF
      ' // Get the Ascii code of the character
      c = ASC(MID(strUtf8, i, 1))
      ' // If it is betwen 0 and 127...
      IF c < 128 THEN 
         ' // ...we simply copy it to the string buffer...
         MID(strAscii, idx, 1) = MID(strUtf8, idx, 1)
         ' // ...and increase the position by 1.
         idx = idx + 1
      ELSEIF c < 224 THEN
         ' // We need to join this byte and the next byte.
         b2 = ASC(MID(strUtf8, i + 1, 1))
         IF b2 > 127 THEN
            c = (c - 192) * 64 + (b2 - 128)
            MID(strAscii, idx, 1) = CHR(c)
            ' // Set the flag to skip the next character
            fSkipChar = TRUE
            ' // Increase the position by 1.
            idx = idx + 1
         END IF
      END IF
   NEXT

   ' // Return the string
   FUNCTION = LEFT(strAscii, idx - 1)

END FUNCTION


' ========================================================================================
' Maps Ansi character string to a UTF-8 string.
' ========================================================================================
FUNCTION AnsiToUtf8( BYREF sAnsi AS STRING ) AS STRING
 dim sUnicode AS STRING
 dim sUtf8    AS STRING

 'Maps Ansi character string to a UTF-8 string.

 'Step one, convert to UNICODE
 sUnicode = string(LEN(sAnsi) * 2, 0)
 MultiByteToWideChar(CP_ACP, _                  'System default Windows ANSI code page
                     MB_PRECOMPOSED, _          'Conversion type
                     cast(LPCSTR, STRPTR(sAnsi)), _     'ANSI string to convert
                     LEN(sAnsi), _              'Lenght of ANSI string
                     cast(LPWSTR, STRPTR(sUnicode)), _  'Unicode string
                     LEN(sUnicode))             'Lenght of Unicode buffer

 'Step two, convert to UTF-8
 sUtf8 = string(LEN(sAnsi), 0)
 WideCharToMultiByte(CP_UTF8, _                 'Set to UTF-8
                     0, _                       'Conversion type
                     cast(LPCWSTR, STRPTR(sUnicode)), _  'Unicode string to convert
                     LEN(sUnicode) / 2, _       'Lenght of Unicode string
                     cast(LPSTR, STRPTR(sUtf8)), _     'UTF-8 string
                     LEN(sUtf8), _              'Length of UTF-8 buffer
                     BYVAL 0, _                 'Invalid character replacement
                     BYVAL 0)                   'Replacement was used flag
 FUNCTION = sUtf8

END FUNCTION


' ========================================================================================
' Maps UTF-8 string to Unicode character string 
' ========================================================================================
FUNCTION Utf8ToUnicode( BYREF ansiStr AS CONST STRING ) AS STRING
'*** This conversion does not appear to be reliable.
'*** Better using the CWSTR.Utf8 method instead.
   DIM dwLen AS DWORD = MultiByteToWideChar(CP_UTF8, 0, STRPTR(ansiStr), LEN(ansiStr), NULL, 0)
   IF dwLen THEN
      DIM s AS STRING = SPACE(dwLen * 2)
      dwLen = MultiByteToWideChar(CP_UTF8, 0, STRPTR(ansiStr), LEN(ansiStr), CAST(WSTRING PTR, STRPTR(s)), dwLen * 2)
      IF dwLen THEN RETURN s
   END IF
end function

   
' ========================================================================================
' Maps Unicode character string to a UTF-8 string.
' ========================================================================================
FUNCTION UnicodeToUtf8( byval wzUnicode as CWSTR ) AS STRING
 dim sUtf8 AS STRING

 ' Maps Unicode character string to a UTF-8 string.
 sUtf8 = string(len(wzUnicode) * 2, 0)
 
 dim as long bytesWritten = _
      WideCharToMultiByte ( _
      CP_UTF8, _                        'Set to UTF-8
      0, _                              'Conversion type
      cast(LPCWSTR, wzUnicode.vptr), _  'Unicode string to convert
      LEN(wzUnicode), _                 'Length of Unicode string
      cast(LPSTR, STRPTR(sUtf8)), _     'UTF-8 string
      LEN(sUtf8), _                     'Length of UTF-8 buffer
      BYVAL 0, _                        'Invalid character replacement
      BYVAL 0)                          'Replacement was used flag

 FUNCTION = left(sUtf8, bytesWritten)

END FUNCTION

' ========================================================================================
' Parse a string to a string array and return the number of lines.
' ========================================================================================
function GetStringToArray( _
            byref txtBuffer as string, _
            txtArray() as string _
            ) as long

   ' Load the lines into the string array. This is MUCH faster than having
   ' to use AfxStrParse to retrieve each line.
   dim as string st
   dim as long iLineStart = 1
   dim as long iLineEnd, nNextLine
   dim as Long nNextBufferArrayLine = 1

   ' Convert the string into an array that can be parsed.
   do until iLineStart >= len( txtBuffer )
      iLineEnd = instr(iLineStart, txtBuffer, vbcrlf)
      if iLineEnd = 0 then iLineEnd = len( txtBuffer )  ' cr/lf not found
      st = mid( txtBuffer, iLineStart, iLineEnd - iLineStart )
      iLineStart = iLineStart + len(st) + len(vbcrlf)
      if nNextBufferArrayLine >= ubound(txtArray) THEN
         redim preserve txtArray( ubound(txtArray) + 5000 )
      END IF
      txtArray(nNextBufferArrayLine) = st   
      nNextBufferArrayLine = nNextBufferArrayLine + 1
   loop

   function = nNextBufferArrayLine - 1
end function


' ========================================================================================
' Open a disk file and read it into a string (ANSI or UTF8)
' ========================================================================================
function GetFileToString( _
            byref wszFilename as const wstring, _
            byref txtBuffer as string, _
            byval pDoc as clsDocument ptr _
            ) as boolean
   
   if pDoc = 0 then return true
   if AfxFileExists(wszFilename) = false then return true
   
   ' Load the entire file into a string
   DIM dwCount AS DWORD, dwFileSize AS DWORD, dwHighSize AS DWORD, dwBytesRead AS DWORD
   DIM hFile AS HANDLE = CreateFileW(@wszFileName, GENERIC_READ, FILE_SHARE_READ, NULL, _
                         OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, NULL)
   IF hFile = INVALID_HANDLE_VALUE THEN return true
   dwFileSize = GetFileSize(hFile, @dwHighSize)
   txtBuffer = String(dwFileSize, 0)
   DIM bSuccess AS LONG = ReadFile(hFile, strptr(txtBuffer), dwFileSize, @dwBytesRead, NULL)
   CloseHandle(hFile)
   IF bSuccess = FALSE THEN return true

   ' A JSON Design file will not have a BOM signature but it is UTF8
   dim as CWSTR wst = wszFilename
   if ucase(AfxStrPathName("EXTN", wst)) = ".DESIGN" then
      pDoc->FileEncoding = FILE_ENCODING_UTF8_BOM
   else
      ' Check for BOM signatures
      if left(txtBuffer, 3) = chr(&HEF, &HBB, &HBF) THEN
         ' UTF8 BOM encoded 
         pDoc->FileEncoding = FILE_ENCODING_UTF8_BOM
         txtBuffer = mid(txtBuffer, 4)   ' bypass the BOM
      elseif left(txtBuffer, 2) = chr(&HFF, &HFE) THEN
         ' UTF16 BOM (little endian) encoded
         pDoc->FileEncoding = FILE_ENCODING_UTF16_BOM 
         txtBuffer = mid(txtBuffer, 3)   ' bypass the BOM
      else
         pDoc->FileEncoding = FILE_ENCODING_ANSI
      end if
   end if

   select case pDoc->FileEncoding
      case FILE_ENCODING_ANSI 
         ' No conversion needed. clsDocument ApplyProperties will *not*
         ' set the editor to UTF8 code.
      
      case FILE_ENCODING_UTF8_BOM   
         ' No conversion needed. clsDocument ApplyProperties will set
         ' the editor to UTF8 code.
         
      case FILE_ENCODING_UTF16_BOM
         ' Convert the whole buffer to UTF-16 unicode string
         dim as CWSTR wszText = string(len(txtBuffer),0)
         MemCpy( CAST(any PTR, wszText.m_pBuffer), strptr(txtBuffer), len(txtBuffer))

         ' Need to parse wszText to remove/process any possible visual designer code. Only
         ' the non-VD text is returned. We also skip over any codegen code. This is only
         ' used in pre-version 3.02 files. New 3.02+ files will call LoadFormJSONdata during
         ' the codewindow creation.
         wszText = pDoc->ParseFormMetaData(HWND_FRMMAIN, wszText)
         
         ' Convert to UTF8 so it can display in the editor
         txtBuffer = UnicodeToUtf8(wszText)

   end select
      
   function = false
end function


' ========================================================================================
' Convert the current text buffer to the specified encoding and redisplay the text.
' ========================================================================================
function ConvertTextBuffer( _
            byval pDoc as clsDocument ptr, _
            byval FileEncoding as long _
            ) as Long
                            
   if pDoc = 0 then exit function

   dim as hwnd hEdit = pDoc->hWndActiveScintilla
   ' Save the current file position and first visible line
   dim nFirstLine as long = SciExec( hEdit, SCI_GETFIRSTVISIBLELINE, 0, 0) 
   Dim nPos As Long = SciExec(hEdit, SCI_GETCURRENTPOS, 0, 0)

   pDoc->FileEncoding = FileEncoding
   Dim As ZString Ptr psz = Cast( ZString Ptr, SciExec(hEdit, SCI_GETCHARACTERPOINTER, 0, 0) )
   dim as long sciCodePage = SciExec(hEdit, SCI_GETCODEPAGE, 0, 0)   ' 0 or SC_CP_UTF8 
   dim as string txtBuffer 
   
   ' Convert buffer to specified file encoding
   select CASE FileEncoding
      case FILE_ENCODING_ANSI
         if sciCodePage = 0 THEN  ' already in ANSI format  
            exit function
         else   
            ' need to convert from UTF8 to ANSI
            txtBuffer = Utf8ToAscii(*psz)
            SciExec(hEdit, SCI_SETCODEPAGE, 0, 0 )
         end if    

      case FILE_ENCODING_UTF8_BOM, FILE_ENCODING_UTF16_BOM
         if sciCodePage = SC_CP_UTF8 THEN  ' already in unicode format
            exit function
         else
            ' need to convert from ANSI to UTF8
            txtBuffer = AnsiToUtf8(*psz)
            SciExec(hEdit, SCI_SETCODEPAGE, SC_CP_UTF8, 0 )
         end if    

   END SELECT
   
   ' Set the new buffer
   pDoc->SetText(txtBuffer)
    
   SciExec(hEdit, SCI_SETFIRSTVISIBLELINE, nFirstLine, 0) 
   SciExec(hEdit, SCI_GOTOPOS, nPos, 0)
   
   function = 0
end function   


' ========================================================================================
' Determine if current line is a valid #Include filename
' ========================================================================================
function IsCurrentLineIncludeFilename() as Boolean
   ' Determine if the text under the current line is a valid #Include filename
   ' and return TRUE if it is. If F6 was pressed then the calling program can
   ' simply open/load the gApp.IncludeFilename. If the right click popup menu
   ' is to be shown then simply add the option to open this file.

   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   If pDoc = 0 Then Exit Function
   
   Dim wszPath         As WString * MAX_PATH
   Dim wszCompilerPath As WString * MAX_PATH
   Dim wszText         As WString * MAX_PATH
   Dim sFilename       As String 
   Dim sLine           As String 
   Dim nLine           As Long   
   Dim i               As Long

   dim as long idxBuild = frmBuildConfig_getActiveBuildIndex()
   if idxBuild = -1 then idxBuild = 0
   if gConfig.Builds(idxBuild).Is32bit then wszCompilerPath = gConfig.FBWINcompiler32
   if gConfig.Builds(idxBuild).Is64bit then wszCompilerPath = gConfig.FBWINcompiler64
   wszCompilerPath = ProcessFromCurdriveApp(wszCompilerPath)
   wszCompilerPath = AfxStrPathname( "PATH", wszCompilerPath ) & "inc"

   ' Convert relative path to absolute path if needed.
   if AfxPathIsRelative(wszCompilerPath) then
      wszCompilerPath = AfxPathCombine(AfxGetExePathName, wszCompilerPath)
   END IF

   nLine = pDoc->GetCurrentLineNumber()
   sLine = LTrim(pDoc->GetLine(nLine))
   
   If Left(Ucase(sLine), 9) = "#INCLUDE " Then sFilename = Mid(sLine, 10)
   If Left(Ucase(sLine), 14) = "#INCLUDE ONCE " Then sFilename = Mid(sLine, 15)

   gApp.IncludeFilename = ""   
   If Len(sFilename) Then
      ' remove any comments at the end of the line
      i = Instr(sFilename, "'")
      If i Then sFilename = Left(sFilename, i-1)
      sFilename = Trim(sFilename, Any Chr(32,34))  ' remove spaces and double quotes
      wszPath = AfxStrPathname( "PATH", pDoc->DiskFilename )

      If AfxFileExists(wszPath & sFilename) Then 
         gApp.IncludeFilename = wszPath & sFilename
      ElseIf AfxFileExists(sFilename) Then 
         gApp.IncludeFilename = sFilename
      ElseIf AfxFileExists(AfxGetCurDir & "\" & sFilename) Then 
         gApp.IncludeFilename = AfxGetCurDir & "\" & sFilename
      ElseIf AfxFileExists(AfxGetExePathName & sFilename) Then 
         gApp.IncludeFilename = AfxGetExePathName & sFilename
      ElseIf AfxFileExists(Str(wszCompilerPath) & "\" & sFilename) Then 
         gApp.IncludeFilename = Str(wszCompilerPath) & "\" & sFilename
      End If
      gApp.IncludeFilename = AfxStrReplace(gApp.IncludeFilename, "/", "\")
      gApp.IncludeFilename = AfxStrReplace(gApp.IncludeFilename, "\\", "\")
   End If
   
   function = AfxFileExists( gApp.IncludeFilename ) 

end function


' ========================================================================================
' Generic open document handler for when Function ListBox item selected or Explorer Treeview
' ========================================================================================
function OpenSelectedDocument( _
            byref wszFilename as wstring, _
            byref wszFunctionName as WSTRING = "", _
            byval nLineNumber as long = -1 _
            ) as clsDocument ptr

   ' This function is called in the following situations:
   '   1. When a selection is made in the Function List.
   '   2. When a selection is made through the Explorer treeview.
   '   3. When a Find In Files line is selected.
   '   4. When a Goto Definition word is clicked on.
   '   5. When OnActivateApp needs to reload a document.
   '   6. When a compile error occurs and need to position to the error line.
   '   7. When right-click select #Include file to open. 
   
   ' If incoming FunctionName then search for filename and line number.
   dim pData as DB2_DATA ptr    
   if len( wszFunctionName ) AndAlso nLineNumber = -1 then
      ' Search for function, sub, or property (get/set)
      pData = gdb2.dbFindFunction( wszFunctionName, wszFilename) 
      if pData then
         wszFilename = pData->fileName
         nLineNumber = pData->nLineStart
      end if
   end if
   
   ' Not all documents exist on disk file. For example, a QuickRun file will exist in
   ' the editor but may never have a disk footprint. We need to search the project to
   ' determine if the filename has a pDoc already associated with it. If it does, then
   ' pass that pDoc rather than looking for a disk filename.
   dim pDoc as clsDocument ptr
   pDoc = gApp.GetDocumentPtrByFilename( wszFilename )

   if pDoc then
      pDoc = frmMain_OpenFileSafely( _
                  HWND_FRMMAIN, _
                  False, _    ' bIsNewFile
                  False, _    ' bIsTemplate
                  true, _     ' bShowInTab
                  false, _    ' bIsInclude
                  "", _       ' wszFileName
                  pDoc )      ' pDocIn
   else   
      if AfxFileExists(wszFilename) = false THEN exit function
      ' Display the document containing the selected sub/function       
      pDoc = frmMain_OpenFileSafely(HWND_FRMMAIN, _
                              False, _    ' bIsNewFile
                              False, _    ' bIsTemplate
                              true, _     ' bShowInTab
                              false, _    ' bIsInclude
                              wszFilename, _  ' wszFileName
                              0 )         ' pDocIn
   end if

   ' Set the top line to display in the editor. I chose to start 3 lines before the
   ' function just to make it visually more appealing.
   if pDoc THEN
      ' Make sure that the code editor is selected for visual designer forms
      if (pDoc->IsDesigner = true) andalso (IsDesignerView(pDoc) = true) then
         if nLineNumber <> -1 then
            pDoc->DesignTabsCurSel = 1  ' 1 = code tab, 0 = Designer
            frmMain_PositionWindows
         end if   
      end if

      ' Do not reposition if incoming LineNumber is -1 because that value represents
      ' the caller specifically not wanting a repositioning.
      if nLineNumber <> -1 then 
         dim as hwnd hEdit = pDoc->hWndActiveScintilla
         ' ensure that the line is visible (not folded)
         if SciExec( hEdit, SCI_GETLINEVISIBLE, nLineNumber, 0) = false then
            ' unfold the block that contains this hidden line
            pDoc->FoldToggle( nLineNumber )
         end if
         SciExec( hEdit, SCI_SETFIRSTVISIBLELINE, Max(nLineNumber - 3, 0), 0) 
         SciExec( hEdit, SCI_GOTOLINE, nLineNumber, 0) 
         pDoc->CenterCurrentLine
      end if
   END IF 
   
   function = pDoc
end function


' ========================================================================================
' Process prefix {CURDRIVE} and convert to current drive letter.
' ========================================================================================
Function ProcessToCurdriveApp( Byval wszFilename As CWSTR ) As CWSTR
   ' For each folder location determine if it resides on the same drive as
   ' the WinFBE app. If it does then substitute the replaceable parameter
   ' {CURDRIVE} for the drive letter. This allows you to easily run the editor
   ' on different media (eg. thumb drive) that may be assigned a different
   ' drive letter.
   dim as CWSTR wszText = AfxGetExePathName

   Dim wszCurDrive As CWSTR = LCase(Left(wszText, 3))  ' eg. D:\

   ' If the incoming filename is a relative file name then the following test
   ' will have no effect.
   If LCase(Left(wszFilename, 3)) = wszCurDrive Then 
      wszFilename = WSTR("{CURDRIVE}") & Mid(wszFilename, 2)
   End If

   Return wszFilename
End Function


' ========================================================================================
' Process current drive to prefix {CURDRIVE} 
' ========================================================================================
Function ProcessFromCurdriveApp( Byval wszFilename As CWSTR ) As CWSTR
   ' For each folder location determine if it resides on the same drive as
   ' the WinFBE app. If it does then substitute the replaceable parameter
   ' {CURDRIVE} for the drive letter. This allows you to easily run the editor
   ' on different media (eg. thumb drive) that may be assigned a different
   ' drive letter.
   dim as CWSTR wszText = AfxGetExePathName

   If Ucase(Left(wszFilename, 10)) = WSTR("{CURDRIVE}") Then 
      wszFilename = Left(wszText, 1) & Mid(wszFilename, 11)
   End If

   Return wszFilename
End Function


' ========================================================================================
' Process prefix {CURDRIVE} and convert to current drive letter.
' ========================================================================================
Function ProcessToCurdriveProject( Byval wszFilename As CWSTR ) As CWSTR
   ' For each folder location determine if it resides on the same drive as
   ' the project file. If it does then substitute the replaceable parameter
   ' {CURDRIVE} for the drive letter. This allows you to easily run the editor
   ' on different media (eg. thumb drive) that may be assigned a different
   ' drive letter.
   dim wszText as CWSTR = gApp.ProjectFilename

   Dim wszCurDrive As CWSTR = LCase(Left(wszText, 3))  ' eg. D:\

   ' If the incoming filename is a relative file name then the following test
   ' will have no effect.
   If LCase(Left(wszFilename, 3)) = wszCurDrive Then 
      wszFilename = WSTR("{CURDRIVE}") & Mid(wszFilename, 2)
   End If

   Return wszFilename
End Function


' ========================================================================================
' Process current drive to prefix {CURDRIVE} 
' ========================================================================================
Function ProcessFromCurdriveProject( Byval wszFilename As CWSTR ) As CWSTR
   ' For each folder location determine if it resides on the same drive as
   ' the project file. If it does then substitute the replaceable parameter
   ' {CURDRIVE} for the drive letter. This allows you to easily run the editor
   ' on different media (eg. thumb drive) that may be assigned a different
   ' drive letter.
   dim wszText as CWSTR = gApp.ProjectFilename
   
   If Ucase(Left(wszFilename, 10)) = WSTR("{CURDRIVE}") Then 
      wszFilename = Left(wszText, 1) & Mid(wszFilename, 11)
   End If

   Return wszFilename
End Function


' ========================================================================================
' Displays the FileOpenDialog.
' The returned pointer must be freed with CoTaskMemFree
' ========================================================================================
Function AfxIFileOpenDialogW( _
            ByVal hwndOwner As HWnd, _
            ByVal idButton As Long _
            ) As WString Ptr

   Dim hr As Long
   Dim CLSID_FileOpenDialog As CLSID = (&hDC1C5A9C, &hE88A, &h4DDE, {&hA5, &hA1, &h60, &hF8, &h2A, &h20, &hAE, &hF7})
   Dim IID_IFileOpenDialog As GUID   = (&hD57C7288, &hD4AD, &h4768, {&hBE, &h02, &h9D, &h96, &h95, &h32, &hD9, &h60})

   ' Create an instance of the FileOpenDialog object
   Dim pofd As IFileOpenDialog Ptr
   hr = CoCreateInstance(@CLSID_FileOpenDialog, Null, CLSCTX_INPROC_SERVER, @IID_IFileOpenDialog, @pofd)
   If pofd = Null Then Return Null

   ' Set the file types depending on the button pushed that calls this open dialog
   Dim rgFileTypes(1 To 5) As COMDLG_FILTERSPEC

   Select Case idButton
      case IDM_LOADSESSION  
         rgFileTypes(1).pszName = @wstr("Session files")
         rgFileTypes(1).pszSpec = @WSTR("*.session")
         rgFileTypes(2).pszName = @L(79,"All files")
         rgFileTypes(2).pszSpec = @WSTR("*.*")
         pofd->lpVtbl->SetFileTypes(pofd, 2, @rgFileTypes(1))
         ' Set the title of the dialog
         hr = pofd->lpVtbl->SetTitle(pofd, L(426,"Load Session"))

      case IDC_FRMUSERTOOLS_CMDBROWSEEXE  
         rgFileTypes(1).pszName = @L(79,"All files")
         rgFileTypes(1).pszSpec = @WSTR("*.*")
         pofd->lpVtbl->SetFileTypes(pofd, 1, @rgFileTypes(1))
         ' Set the title of the dialog
         hr = pofd->lpVtbl->SetTitle(pofd, L(291,"Command:"))

      Case IDM_PROJECTOPEN
         rgFileTypes(1).pszName = @L(216,"Project files")
         rgFileTypes(1).pszSpec = @WSTR("*.wfbe")
         rgFileTypes(2).pszName = @L(79,"All files")
         rgFileTypes(2).pszSpec = @WSTR("*.*")
         pofd->lpVtbl->SetFileTypes(pofd, 2, @rgFileTypes(1))
         ' Set the title of the dialog
         hr = pofd->lpVtbl->SetTitle(pofd, L(216,"Project files"))
      
      Case IDM_INSERTFILE
         rgFileTypes(1).pszName = @L(106,"Open source file")
         rgFileTypes(2).pszName = @L(77,"Code files")
         rgFileTypes(3).pszName = @L(78,"Header files")
         rgFileTypes(4).pszName = @L(209,"Resource files")
         rgFileTypes(5).pszName = @L(79,"All files")
         rgFileTypes(1).pszSpec = @WSTR("*.bas;*.bi;*.inc;*.rc")
         rgFileTypes(2).pszSpec = @WSTR("*.bas;*.inc")
         rgFileTypes(3).pszSpec = @WSTR("*.bi")
         rgFileTypes(4).pszSpec = @WSTR("*.rc")
         rgFileTypes(5).pszSpec = @WSTR("*.*")
         pofd->lpVtbl->SetFileTypes(pofd, 5, @rgFileTypes(1))
         ' Set the title of the dialog
         hr = pofd->lpVtbl->SetTitle(pofd, L(80,"Insert File"))

      Case IDC_FRMOPTIONSLOCAL_CMDLOCALIZATION  '1012
         rgFileTypes(1).pszName = @L(102,"Localization files")
         rgFileTypes(1).pszSpec = @WSTR("*.lang")
         rgFileTypes(2).pszName = @L(79,"All files")
         rgFileTypes(2).pszSpec = @WSTR("*.*")
         pofd->lpVtbl->SetFileTypes(pofd, 2, @rgFileTypes(1))
         ' Set the title of the dialog
         hr = pofd->lpVtbl->SetTitle(pofd, L(103,"Open Localization File"))

      Case IDC_FRMOPTIONSCOMPILER_CMDFBHELPFILE 
         rgFileTypes(1).pszName = @L(104,"Help file")
         rgFileTypes(1).pszSpec = @WSTR("*.chm")
         rgFileTypes(2).pszName = @L(79,"All files")
         rgFileTypes(2).pszSpec = @WSTR("*.*")
         pofd->lpVtbl->SetFileTypes(pofd, 2, @rgFileTypes(1))
         ' Set the title of the dialog
         hr = pofd->lpVtbl->SetTitle(pofd, L(105,"Find Help File"))
   End Select
   
   ' Display the dialog
   hr = pofd->lpVtbl->Show(pofd, hwndOwner)
   hr = pofd->lpVtbl->SetOptions(pofd, FOS_NOCHANGEDIR)

   ' Get the result
   Dim pItem As IShellItem Ptr
   Dim pwszName As WString Ptr
   If SUCCEEDED(hr) Then
      hr = pofd->lpVtbl->GetResult(pofd, @pItem)
      If SUCCEEDED(hr) Then
         hr = pItem->lpVtbl->GetDisplayName(pItem, SIGDN_FILESYSPATH, @pwszName)
         Function = pwszName
      End If
   End If

   ' Cleanup
   If pItem Then pItem->lpVtbl->Release(pItem)
   If pofd Then pofd->lpVtbl->Release(pofd)

End Function


' ========================================================================================
' Displays the FileOpenDialog (multiple selection)
' Returns a pointer to the IShellItemArray collection.
' ========================================================================================
Function AfxIFileOpenDialogMultiple( _
            ByVal hwndOwner As HWnd, _
            ByVal idButton As Long _
            ) As IShellItemArray Ptr

   ' Create an instance of the FileOpenDialog interface
   Dim hr As Long
   Dim pofd As IFileOpenDialog Ptr
   hr = CoCreateInstance( @CLSID_FileOpenDialog, Null, CLSCTX_INPROC_SERVER, _
                             @IID_IFileOpenDialog, @pofd)
   If pofd = Null Then Return Null

   ' Set the file types
   Dim rgFileTypes(1 To 5) As COMDLG_FILTERSPEC

   select case idButton
      case IDM_FILEOPEN
         rgFileTypes(1).pszName = @L(106,"Open source file")
         rgFileTypes(2).pszName = @L(77,"Code files")
         rgFileTypes(3).pszName = @L(78,"Header files")
         rgFileTypes(4).pszName = @L(209,"Resource files")
         rgFileTypes(5).pszName = @L(79,"All files")
         rgFileTypes(1).pszSpec = @WSTR("*.bas;*.bi;*.inc;*.rc")
         rgFileTypes(2).pszSpec = @WSTR("*.bas;*.inc")
         rgFileTypes(3).pszSpec = @WSTR("*.bi")
         rgFileTypes(4).pszSpec = @WSTR("*.rc")
         rgFileTypes(5).pszSpec = @WSTR("*.*")
         pofd->lpVtbl->SetFileTypes(pofd, 5, @rgFileTypes(1))
      
      case IDM_ADDIMAGE
         rgFileTypes(1).pszName = @L(378,"Images")
         rgFileTypes(2).pszName = @L(79,"All files")
         rgFileTypes(1).pszSpec = @WSTR("*.ico;*.bmp;*.jpg;*.gif;*.wmf;*.png;*.tiff;*.cur")
         rgFileTypes(2).pszSpec = @WSTR("*.*")
         pofd->lpVtbl->SetFileTypes(pofd, 2, @rgFileTypes(1))
         
   end select

   ' Set the title of the dialog
   hr = pofd->lpVtbl->SetTitle(pofd, L(248,"Open file"))

   ' Set the default folder to display in the open dialog
   dim wszDefaultFolder as wstring * MAX_PATH
   if gApp.wszLastOpenFolder = "" then 
      if AfxFileExists( gApp.ProjectFilename ) then 
         wszDefaultFolder = AfxStrPathName( "PATH", gApp.ProjectFilename )
      else
         wszDefaultFolder = AfxGetCurDir
      end if
   else
      wszDefaultFolder = gApp.wszLastOpenFolder
   end if
   
   Dim pFolder As IShellItem Ptr
   SHCreateItemFromParsingName (wszDefaultFolder, Null, @IID_IShellItem, @pFolder)
   If pFolder Then
      hr = pofd->lpVtbl->SetFolder(pofd, pFolder)
      If SUCCEEDED(hr) Then
         pFolder->lpVtbl->Release(pFolder)
      end if      
   End If
   
   ' Allow multiselection
   hr = pofd->lpVtbl->SetOptions(pofd, FOS_ALLOWMULTISELECT Or FOS_NOCHANGEDIR or FOS_FILEMUSTEXIST)
   ' Display the dialog
   hr = pofd->lpVtbl->Show(pofd, hwndOwner)

   ' Get the result
   Dim pItemArray As IShellItemArray Ptr
   If SUCCEEDED(hr) Then
      hr = pofd->lpVtbl->GetResults(pofd, @pItemArray)
      Function = pItemArray
   End If

   If pofd Then pofd->lpVtbl->Release(pofd)

End Function


' ========================================================================================
' Displays the FileSaveDialog
' The returned pointer must be freed with CoTaskMemFree
' ========================================================================================
Function AfxIFileSaveDialog( _
            ByVal hwndOwner As HWnd, _
            ByVal pwszFileName As WString Ptr, _    ' full path and filename
            ByVal pwszDefExt As WString Ptr, _
            ByVal id As Long = 0, _
            ByVal sigdnName As SIGDN = SIGDN_FILESYSPATH _
            ) As WString Ptr

   ' // Create an instance of the IFileSaveDialog interface
   Dim rgFileTypes(1 To 4) As COMDLG_FILTERSPEC 
   Dim hr As Long
   Dim psfd As IFileSaveDialog Ptr
   hr = CoCreateInstance(@CLSID_FileSaveDialog, Null, CLSCTX_INPROC_SERVER, @IID_IFileSaveDialog, @psfd)
   If psfd = Null Then Return Null

   dim as CWSTR wszFilename, wszFilePath 
   
   ' Add extensions if it does not already exist as part of the filename
   wszFilename = AfxStrPathname( "NAMEX", *pwszFileName )
   if len(wszFilename) then 
      if AfxStrPathname( "EXTN", wszFilename ) = "" then
         if len(*pwszDefExt) then
            wszFilename = wszFilename & "." & *pwszDefExt
         end if   
      end if   
   end if
   
   if AfxFileExists( *pwszFileName ) then
      wszFilePath = AfxStrPathName( "PATH", *pwszFileName )
   else   
      ' Set the default folder to save the file
      if gApp.wszLastOpenFolder then 
         wszFilePath = gApp.wszLastOpenFolder
      else
         ' New file being saved try to default to the project folder
         if AfxFileExists( gApp.ProjectFilename ) then
            wszFilePath = AfxStrPathname( "PATH", gApp.ProjectFilename )
         end if
      end if
      gApp.wszLastOpenFolder = wszFilePath
   end if

   ' Set the file types
   Select Case id
      Case IDM_SAVESESSION
         rgFileTypes(1).pszName = @wstr("Session files")
         rgFileTypes(1).pszSpec = @WSTR("*.session")
         rgFileTypes(2).pszName = @L(79,"All files")
         rgFileTypes(2).pszSpec = @WSTR("*.*")
         psfd->lpVtbl->SetFileTypes(psfd, 2, @rgFileTypes(1))
         ' // Set the title of the dialog
         hr = psfd->lpVtbl->SetTitle(psfd, L(425,"Save Session"))

      Case IDC_FRMPROJECTOPTIONS_CMDSELECT, IDM_PROJECTSAVE, IDM_PROJECTSAVEAS
         rgFileTypes(1).pszName = @L(216,"Project files")
         rgFileTypes(1).pszSpec = @WSTR("*.wfbe")
         rgFileTypes(2).pszName = @L(79,"All files")
         rgFileTypes(2).pszSpec = @WSTR("*.*")
         psfd->lpVtbl->SetFileTypes(psfd, 2, @rgFileTypes(1))
         ' // Set the title of the dialog
         hr = psfd->lpVtbl->SetTitle(psfd, L(185,"Save Project As..."))

      Case IDC_FRMOPTIONSLOCAL_CMDNEW
         rgFileTypes(1).pszName = @L(102,"Localization files")
         rgFileTypes(1).pszSpec = @WSTR("*.lang")
         rgFileTypes(2).pszName = @L(79,"All files")
         rgFileTypes(2).pszSpec = @WSTR("*.*")
         psfd->lpVtbl->SetFileTypes(psfd, 2, @rgFileTypes(1))
         psfd->lpVtbl->SetTitle(psfd, L(8,"Save As..."))

      Case Else
         rgFileTypes(1).pszName = @L(77,"FB code files")
         rgFileTypes(1).pszSpec = @WSTR("*.bas")
         rgFileTypes(2).pszName = @L(78,"FB Include files")
         rgFileTypes(2).pszSpec = @WSTR("*.bi;*.inc")
         rgFileTypes(3).pszName = @L(209,"Resource files")
         rgFileTypes(3).pszSpec = @WSTR("*.rc")
         rgFileTypes(4).pszName = @L(79,"All files")
         rgFileTypes(4).pszSpec = @WSTR("*.*")
         psfd->lpVtbl->SetFileTypes(psfd, 4, @rgFileTypes(1))
         psfd->lpVtbl->SetTitle(psfd, L(8,"Save As..."))
         if pwszDefExt then 
            if *pwszDefExt = "inc" then
               psfd->lpVtbl->SetFileTypeIndex(psfd, 2)
            end if
         end if
   End Select
   
   ' // Set the file name
   hr = psfd->lpVtbl->SetFileName(psfd, wszFileName)
   ' // Set the extension
   hr = psfd->lpVtbl->SetDefaultExtension(psfd, pwszDefExt)

   ' // Set the default folder to display in the save dialog
   if len(wszFilePath ) then
      Dim pFolder As IShellItem Ptr
      SHCreateItemFromParsingName (wszFilePath, Null, @IID_IShellItem, @pFolder)
      If pFolder Then
         hr = psfd->lpVtbl->SetFolder(psfd, pFolder)
         If SUCCEEDED(hr) Then
            pFolder->lpVtbl->Release(pFolder)
         end if      
      End If
   end if
   
   ' // Display the dialog
   hr = psfd->lpVtbl->Show(psfd, hwndOwner)

   ' // Get the result
   Dim pItem As IShellItem Ptr
   Dim pwszName As WString Ptr
   If SUCCEEDED(hr) Then
      hr = psfd->lpVtbl->GetResult(psfd, @pItem)
      If SUCCEEDED(hr) Then
         hr = pItem->lpVtbl->GetDisplayName(pItem, sigdnName, @pwszName)
         Function = pwszName
      End If
   End If
   ' // Cleanup
   If pItem Then pItem->lpVtbl->Release(pItem)
   If psfd Then psfd->lpVtbl->Release(psfd)

End Function


' ========================================================================================
' Inserts an item at a specific location in the ListView.
' ========================================================================================
Function FF_ListView_InsertItem( _
            ByVal hWndControl As HWnd, _
            ByVal iRow        As Long, _         
            ByVal iColumn     As Long, _
            ByVal pwszText    As WString Ptr, _
            ByVal lParam      As LPARAM = 0 _
            ) As BOOLEAN

   Dim lvi As LVITEMW
   lvi.iItem     = iRow
   lvi.iSubItem  = iColumn 
   lvi.pszText   = pwszText
   lvi.lParam    = lParam
   If iColumn = 0 Then
      lvi.mask = LVIF_TEXT Or LVIF_PARAM Or LVIF_IMAGE 
      Function = SendMessage( hWndControl, LVM_INSERTITEM, 0, Cast(LPARAM, @lvi) )
   Else 
      lvi.mask = LVIF_TEXT Or LVIF_IMAGE
      Function = SendMessage( hWndControl, LVM_SETITEM, 0, Cast(LPARAM, @lvi) )
   End If
End Function


' ========================================================================================
' Retrieves the text of a ListView item.
' ========================================================================================
Function FF_ListView_GetItemText( _
            ByVal hWndControl As HWnd, _
            ByVal iRow As Long, _
            ByVal iColumn As Long, _
            ByVal pwszText As WString Ptr, _
            ByVal nTextMax As Long _
            ) As BOOLEAN

   If pwszText = 0 Then Return False
   If nTextMax = 0 Then Return False
   Dim lvi As LVITEMW

   lvi.mask       = LVIF_TEXT
   lvi.iItem      = iRow
   lvi.iSubItem   = iColumn 
   lvi.pszText    = pwszText
   lvi.cchTextMax = nTextMax
       
   Function = SendMessage( hWndControl, LVM_GETITEM, 0, Cast(LPARAM, @lvi) )
End Function


' ========================================================================================
' Set the text for the specified row and col item
' ========================================================================================
Function FF_ListView_SetItemText( _
            ByVal hWndControl As HWnd, _
            ByVal iRow As Long, _
            ByVal iColumn As Long, _
            ByVal pwszText As WString Ptr, _
            ByVal nTextMax As Long _
            ) As Long

  Dim li As LV_ITEM
  li.mask       = LVIF_TEXT
  li.iItem      = iRow
  li.iSubItem   = iColumn 
  li.pszText    = pwszText
  li.cchTextMax = nTextMax
  Function = SendMessage( hWndControl, LVM_SETITEM, 0, Cast(LPARAM, @li) )
End Function


' ========================================================================================
' Retrieve the lParam value from a Listview line
' ========================================================================================
Function FF_ListView_GetlParam( _
            ByVal hWndControl As HWnd, _
            ByVal iRow As Long _
            ) As LPARAM

   Dim li As LV_ITEM
   li.mask       = LVIF_PARAM
   li.iItem      = iRow
   li.iSubItem   = 0 
   If SendMessage( hWndControl, LVM_GETITEM, 0, Cast(LPARAM, @li) ) Then
      Function = li.lParam
   End If   
End Function


' ========================================================================================
' Set the lParam value for a Listview line
' ========================================================================================
Function FF_ListView_SetlParam( _
            ByVal hWndControl As HWnd, _
            ByVal iRow As Long, _
            ByVal ilParam As LPARAM _
            ) As long
   Dim li As LV_ITEM
   li.mask       = LVIF_PARAM
   li.iItem      = iRow
   li.iSubItem   = 0
   li.lParam     = ilParam
   function = SendMessage( hWndControl, LVM_SETITEM, 0, Cast(LPARAM, @li) ) 
End Function


' ========================================================================================
' Load a .lang localization file from disk and populate the localization array
' The IsEnglish parameter is used when we want to populate the gLangEnglish global
' array that is used in the WinFBE.bas startup code.
' ========================================================================================
Function LoadLocalizationFile( _
            Byref wszFileName As CWSTR, _
            byval IsEnglish as boolean = false _
            ) As BOOLEAN

   ' default that the file failed to load
   Function = False
   If AfxFileExists( wszFileName ) = 0 Then Exit Function

   Dim as CBSTR wst, wKey, wData
   Dim nKey  As Long
   Dim nData As Long  
   Dim i     As Long
    
   dim pStream AS CTextStream
   if pStream.OpenUnicode( wszFileName ) <> S_OK then exit function
   
   do until pStream.EOS
      wst = pStream.ReadLine
      
      If Len(wst) = 0 Then Continue Do
      If Left(wst, 1) = "'" Then Continue Do
      
      i = Instr(wst, ":")
      If i = 0 Then Continue Do
      
      wKey = "": wData = "": nData = 0

      wKey  = Left(wst, i-1)
      wData = Mid(**wst, i+1)    ' MID causes problems with Chinese data so ** is used.
      
      nKey  = Val(wKey)
      nData = Val(wData)

      If Ucase(wKey) = "MAXIMUM" Then
         ' resize the global dynamic array
         if IsEnglish then
            ReDim gLangEnglish(nData) As WString * MAX_PATH
         else   
            ReDim LL(nData) As WString * MAX_PATH
         end if
      Else
         ' this should be a key/value pair line in the format:
         ' 00001:value
         ' Ensure that we add the value to the array within the valid
         ' boundaries of the array.
         if IsEnglish then
            If (nKey >= LBound(gLangEnglish)) AndAlso (nKey <= Ubound(gLangEnglish)) Then
               ' Use ** to ensure that cyrillic langauge gets converted correctly. FB intrinsic
               ' functions (RTRIM) automatically convert those incorrectly when using CBSTR or CWSTR.
               gLangEnglish(nKey) = rtrim(**AfxStrParse(wData, 1, ";"), any chr(9,32))
            end if
         else
            If (nKey >= LBound(LL)) AndAlso (nKey <= Ubound(LL)) Then
               ' Remove any comments from end of the line. Comments begin with
               ' a semicolon character.
               ' Use ** to ensure that cyrillic langauge gets converted correctly. FB intrinsic
               ' functions (RTRIM) automatically convert those incorrectly when using CBSTR or CWSTR.
               LL(nKey) = rtrim(**AfxStrParse(wData, 1, ";"), any chr(9,32))
               ' If the local phrase is empty then fill it using the English version.
               if len(LL(nKey)) = 0 then 
                  If (nKey >= LBound(gLangEnglish)) AndAlso (nKey <= Ubound(gLangEnglish)) Then
                     LL(nKey) = gLangEnglish(nKey)
                  end if
               end if   
            End If
         end if
      End If   
         
   Loop
   pStream.Close
 
   Function = True
End Function


' ========================================================================================
' Get the full process image name
' ========================================================================================
Function GetProcessImageName( _
            ByVal pe32w As PROCESSENTRY32W Ptr, _
            ByVal pwszExeName As WString Ptr _
            ) As Long

   Dim dwSize As Long
   Dim hProcess As HANDLE 
   hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 1, pe32w->th32ProcessID)
   If hProcess Then
      dwSize = MAX_PATH
      
'      Without using dynamic loading...      
'      QueryFullProcessImageNameW( hProcess, 0, pwszExeName, @dwSize ) 
'      CloseHandle hProcess

      ' QueryFullProcessImageNameW is only available in Vista or higher. Try to dynamically load the 
      ' function because statically linking to it will cause a runtime error if WinFBE is run using WinXP.
      Dim As Any Ptr hLib = DyLibLoad("Kernel32")
      If hLib then
         dim MyQueryFullProcessImageName as function( byval hProcess as HANDLE, byval dwFlags as DWORD, byval lpExeName as LPWSTR, byval lpdwSize as PDWORD) as WINBOOL
         MyQueryFullProcessImageName = DyLibSymbol( hLib, "QueryFullProcessImageNameW" )
         If MyQueryFullProcessImageName Then
            MyQueryFullProcessImageName( hProcess, 0, pwszExeName, @dwSize ) 
            CloseHandle hProcess
         end if
         DyLibFree(hLib)
      End If

   End If
   Function = 0
End Function


' ========================================================================================
' Checks if the program that we are going to compile is already running
' ========================================================================================
Function IsProcessRunning( ByVal pwszExeFileName As WString Ptr ) As BOOLEAN

   Dim hSnapShot As HANDLE
   Dim pe32w As PROCESSENTRY32W

   Dim wszExeFileName As WString * MAX_PATH = Ucase(*pwszExeFileName)
   Dim wszExeProcessName As WString * MAX_PATH

   pe32w.dwSize = Sizeof(pe32w)
   hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)

   If hSnapShot <> INVALID_HANDLE_VALUE Then
      If Process32First(hSnapShot, @pe32w) Then
         GetProcessImageName( @pe32w, @wszExeProcessName )
         If Ucase(wszExeProcessName) = wszExeFileName Then
            Function = True
         Else
            Do While Process32Next(hSnapShot, @pe32w) > 0
               GetProcessImageName( @pe32w, @wszExeProcessName )
               If Ucase(wszExeProcessName) = wszExeFileName Then
                  Function = True
                  Exit Do
               End If
            Loop
         End If
      End If
      CloseHandle hSnapShot
   End If

End Function


' ========================================================================================
' Determine the EXE name of the currently active document or project.
' ========================================================================================
function GetRunExecutableFilename() as CWSTR
   ' Used by the top menu to determine if the Run Executable option is available
   dim wszFilename as CWSTR
   dim wszOptions as CWSTR
   Dim pDocMain As clsDocument Ptr 

   dim as long idxBuild = frmBuildConfig_getActiveBuildIndex()
   dim as Boolean bInString = false
   
   pDocMain = iIf( gApp.IsProjectActive, gApp.GetMainDocumentPtr, gTTabCtl.GetActiveDocumentPtr)
   If pDocMain = 0 Then return ""
   
   if idxBuild > -1 then 
      wszOptions = gConfig.Builds(idxBuild).wszOptions 

      if gApp.IsProjectActive THEN
         if gConfig.Builds(idxBuild).Is32bit then
            wszOptions = wszOptions + " " + gApp.ProjectOther32
         end if   
         if gConfig.Builds(idxBuild).Is64bit then
            wszOptions = wszOptions + " " + gApp.ProjectOther64
         end if   
      else
         wszOptions = wszOptions + " " + gConfig.CompilerSwitches
      END IF
      wszOptions = " " + ucase(wszOptions) + " "
      
      If Instr(wszOptions, wstr(" -DLL "))   Then return ""
      If Instr(wszOptions, wstr(" -DYLIB ")) Then return ""
      If Instr(wszOptions, wstr(" -LIB "))   then return ""
      
   end if
   
   ' Need to check the compiler options to see if the -x switch is used. That switch specifically
   ' names the output file.
   dim y as long = Instr(wszOptions, wstr(" -X "))
   if y THEN
      y = y + 4  ' skip over the switch itself
      for i as long = y to len(wszOptions)
         ' iterate to the beginning of the next switch or end of the string
         ' skip over spaces that are part of a filename string
         if wszOptions[i-1] = 34 THEN bInString = not bInString  ' double quotes
         if bInString THEN continue for
         if wszOptions[i-1] = 45 then  ' dash   
            wszFilename = mid(wszOptions, y, i-y-1)
            exit for
         end if   
      NEXT
      ' There was no other switch so we made it to the end of the string
      if len(wszFilename) = 0 then wszFilename = mid(wszOptions, y)
   END IF
   
   ' Has a filename been determined yet? If it has then it possibly does not have
   ' a path assigned to it so we should add one.
   if len(wszFilename) THEN
      if len(AfxStrPathname("PATH", wszFilename)) = 0 THEN
         wszFilename = AfxStrPathname("PATH", pDocMain->DiskFilename) + wszFilename
      END IF
   else
      ' Default 
      wszFilename = AfxStrPathname("PATH", pDocMain->DiskFilename) + _
                    AfxStrPathname("NAME", pDocMain->DiskFilename) + _
                    wstr(".exe")
   END IF

   return wszFilename

END FUNCTION



' ========================================================================================
' Loads an image from a resource, converts it to an icon or bitmap and returns a pointer
' to the raw PNG pixels. This is needed by the scintilla autocomplete popup.
' Memory is ALLOCATE so it needs to be DEALLOCATE.
' Parameters:
' - hInstance     = [in] A handle to the module whose portable executable file or an accompanying
'                   MUI file contains the resource. If this parameter is NULL, the function searches
'                   the module used to create the current process.
' - wszImageName  = [in] Name of the image in the resource file (.RES). If the image resource uses
'                   an integral identifier, wszImage should begin with a number symbol (#)
'                   followed by the identifier in an ASCII format, e.g., "#998". Otherwise,
'                   use the text identifier name for the image. Only images embedded as raw data
'                   (type RCDATA) are valid. These must be icons in format .png, .jpg, .gif, .tiff.
' ========================================================================================
TYPE MYBITMAPINFO
    bmiHeader AS BITMAPINFOHEADER
    bmiColors(256) AS RGBQUAD 
END TYPE
             

' ========================================================================================
' Get a raw PNG from the Resource - needed to load images into Scintilla popup list.
' ========================================================================================
FUNCTION LoadPNGfromRes( _
            BYVAL hInstance AS HINSTANCE, _
            BYREF wszImageName AS WSTRING _
            ) as any ptr

   dim as HBITMAP hBitmap = AfxGdipImageFromRes(hInstance, wszImageName, 0, false, IMAGE_BITMAP) 
        
   dim bi AS MYBITMAPINFO
   dim bm AS BITMAP
   dim dwp AS DWORD PTR
   dim as HDC hIC, hDC 
   dim as any ptr pPixel
   
   IF hBitmap THEN
      hIC = GetWindowDC(0)
      hDC = CreateCompatibleDC(hIC)
      SelectObject(hDC, hBitmap)

      GetObject(hBitmap, SIZEOF(bm), @bm)
      bi.bmiHeader.biSize        = SIZEOF(bi.bmiHeader)
      bi.bmiHeader.biWidth       = bm.bmWidth
      bi.bmiHeader.biHeight      = -bm.bmHeight ' Put top in TOP instead on bottom!
      bi.bmiHeader.biPlanes      = 1
      bi.bmiHeader.biBitCount    = 32
      bi.bmiHeader.biCompression = BI_RGB
                      
      pPixel = ALLOCATE((bm.bmWidth * bm.bmHeight) * sizeof(DWORD))
      dwp = cast(dword ptr, pPixel)
      GetDIBits( hDC, hBitmap, 0, bm.bmHeight, _
                        BYVAL dwp, cast(LPBITMAPINFO, @bi), DIB_RGB_COLORS)
     
      DeleteDC(hIC)
      DeleteDC(hDC)
   end if

   ' // Return the handle to the PNG raw pixels
   DeleteObject(hBitmap)
   return pPixel
END FUNCTION


' ========================================================================================
' Check PlanetSquires server for latest WinFBE version.
' ========================================================================================
function DoCheckForUpdates( _
            byval hWndParent as hwnd, _
            byval bSilentCheck as Boolean = false _
            ) as long
   
   ' Check safeguard for when Beta versions are issued.
   if PREVENT_UPDATE_CHECK then exit function
   
   ''  Contact the PlanetSquires server and download the text file containing
   ''  the latest WinFBE version number. If an update exists then ask the user
   ''  if they wish to navigate to the Releases page on GitHub.
   ''
   dim as CWSTR wszServerFile = "winfbe_version.txt"
   dim as CWSTR wszFilename
   dim as CWSTR wszLatestVersion
   dim as CWSTR wszMsg
   dim as string st
   dim as boolean bFailedCheck = false

   if AfxWinHttpCheckPlatform() = false then
      ' This platform does not support the Windows HTTP Services (WinHTTP)
      'print "Windows HTTP Services (WinHTTP) not supported on this platform"
      exit function
   end if
   
   DIM pWHttp AS CWinHttpRequest
   
   wszFilename = AfxGetExePathName & wszServerFile

   ' Open an HTTP connection to an HTTP resource (synchronous mode)
   if pWHttp.Open( "GET", "https://www.planetsquires.com/" & wszServerFile ) = S_OK then

      ' Send HTTP request and wait 5 seconds for response 
      pWHttp.Send
    
      if pWHttp.WaitForResponse(5) then
         st = pWHttp.GetResponseBody
         ' Open a file stream and save the downloaded file
         DIM pFileStream AS CFileStream
         IF pFileStream.Open(wszFilename, STGM_CREATE OR STGM_WRITE) = S_OK then
            pFileStream.WriteTextA(st)
            pFileStream.Close
         else
            bFailedCheck = true
         END IF
      end if
   
   else
      bFailedCheck = true
      'print "Error GET update version file: "; pWHttp.GetErrorInfo()
   end if
   
   ' Open the downloaded file and check for the version number. It is possible that the socket
   ' will open but fail and send request but security parameters of the system will prevent data
   ' from being received resulting in a zero byte file. 
   if bFailedCheck = false then
      if AfxFileExists( wszFilename ) then
         dim pStream AS CTextStream 
         if pStream.Open(wszFilename, IOMODE_FORREADING) = S_OK then 
            Do Until pStream.EOS
               st = pStream.ReadLine
               if left( st, 15 ) = "latest_version=" then
                  wszLatestVersion = trim(mid( st, 16 ))
                  if len(wszLatestVersion) = 0 then
                     bFailedCheck = true
                  end if   
                  exit do
               end if
            loop
            pStream.Close
         end if
      end if
   end if
      
   ' Check the installed vs. available version numbers   
   dim as long latestVersion = ConvertWinFBEversion(wszLatestVersion)
   dim as long installedVersion = ConvertWinFBEversion(APPVERSION)

   if (bFailedCheck = true) or (latestVersion = 0) then
      if bSilentCheck = false then      
         MessageBox( hWndParent, _
                     L(92,"Failed to retrieve update information"), _
                     L(94,"Software Update"), _
                     MB_ICONINFORMATION Or MB_OK )
      end if
      AfxDeleteFile( wszFilename ) 
      exit function
   end if
      
      
   ' Save the config file so that other editor instances will not also do update checks again
   gConfig.LastUpdateCheck = JulianDateNow
   gConfig.SaveConfigFile

   AfxDeleteFile( wszFilename ) 

   if bSilentCheck = false then      
      if installedVersion >= latestVersion then
         wszMsg = L(96,"You are up to date!") & vbcrlf & _
                  "WinFBE v" & APPVERSION & " " & L(97,"is currently the newest version available.")
         MessageBox( hWndParent, wszMsg,L(94,"Software Update"), MB_ICONINFORMATION Or MB_OK )
      
      elseif installedVersion < latestVersion then
         wszMsg = APPNAME & vbcrlf & _
                  L(98,"A new version is available.") & vbcrlf & vbcrlf & _
                  L(99,"Current") & ": " & APPVERSION & vbcrlf & _
                  L(107,"Available") & ": " & wszLatestVersion & vbcrlf & vbcrlf & _
                  L(137,"Do you wish to visit the download website?")
         If MessageBox( hWndParent, wszMsg, L(94,"Software Update"), MB_ICONQUESTION Or MB_YESNOCANCEL ) = IDYES Then 
            ShellExecute( NULL, "open", "https://github.com/PaulSquires/WinFBE/releases", Null, Null, SW_SHOWNORMAL )
         end if
      end if
   end if
   
   function = 0
end function


' ========================================================================================
' Calcluate the client area at bottom of Listbox not covered by a row (needed to manually
' paint the unused area in WM_ERASEBKGRD messages to avoid flicker.
' ========================================================================================
function GetListBoxEmptyClientArea( _
            byval hListBox as HWND, _
            byval nLineHeight as long _
            ) as RECT
   
   dim as RECT rc: GetClientRect( hListBox, @rc )
   ' If the number of lines in the listbox is less than the number per page then 
   ' calculate from last item to bottom of listbox, otherwise calculate based on
   ' the mod of the lineheight to listbox height so we can color the partial line
   ' that won't be displayed at the bottom of the list.
   dim as RECT rcItem
   SendMessage( hListBox, LB_GETITEMRECT, 0, cast(LPARAM, @rcItem) )
   dim as long itemHeight = rcItem.bottom - rcItem.top
   dim as long NumItems = ListBox_GetCount(hListBox)
   dim as long ItemsPerPage = ( rc.bottom \ itemHeight )
   dim as long nTopIndex = SendMessage( hListBox, LB_GETTOPINDEX, 0, 0 ) 
   dim as long visible_rows = 0

   if NumItems > 0 then
      ItemsPerPage = (rc.bottom - rc.top) / itemHeight
      dim as long bottom_index = (nTopIndex + ItemsPerPage)
      if bottom_index >= NumItems then bottom_index = NumItems - 1
      visible_rows = (bottom_index - nTopIndex) + 1
   end if

   rc.top = visible_rows * itemHeight 
   if rc.top > rc.bottom then rc.top = rc.bottom
   
   return rc
end function