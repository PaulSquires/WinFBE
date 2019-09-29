'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2019 Paul Squires, PlanetSquires Software
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
' Change everything between quotes to # symbols
' ========================================================================================
public function MaskStringCharacters( byval st as string) as string
   ' Iterate the line and change everything between quotes to # symbols. This
   ' ensures that we correctly deal with strings that have embedded single 
   ' quote characters.
   dim as long i 
   dim as Boolean bInString = false
   for i = 0 to len(st) - 1
      if st[i] = 34 THEN bInString = not(bInString)
      if bInString THEN  
         if st[i] <> 34 THEN st[i] = 35   ' # symbol
      END IF
   NEXT
   function = st
end function


' ========================================================================================
' Removes a single line comment.
' ========================================================================================
public function RemoveComments( byval st as string) as string
   function = AfxStrExtract( 1, st, "'")
end function


' ========================================================================================
' Calculate the Julian date for today's date. Used by CheckForUpdates.
' ========================================================================================
public function JulianDateNow() as long
   function = AfxGregorianToJulian( AfxSystemDay, AfxSystemMonth, AfxSystemYear )
end function


' ========================================================================================
' Convert WinFBE version to whole number for comparison to file versions being loaded.
' ========================================================================================
public function ConvertWinFBEversion( byref wszVersion as wstring ) as long
   dim as CWSTR wszPart1, wszPart2, wszPart3
   
   wszPart1 = AfxStrLSet( AfxStrParse(wszVersion, 1, "."), 3, "0" )
   wszPart2 = AfxStrLSet( AfxStrParse(wszVersion, 2, "."), 3, "0" )
   wszPart3 = AfxStrRSet( AfxStrParse(wszVersion, 3, "."), 3, "0" )
   
   function = val(wszPart1 + wszPart2 + wszPart3) 
end function


' ========================================================================================
' Disable all modeless windows belonging to frmMain so that the popup modal is truly modal.
' ========================================================================================
public function DisableAllModeless() as long
   ' No need to enable/disable the modeless Help form.
   if IsWindowVisible(HWND_FRMFUNCTIONLIST) then EnableWindow(HWND_FRMFUNCTIONLIST, false)
   if IsWindowVisible(HWND_FRMFINDREPLACE)  then EnableWindow(HWND_FRMFINDREPLACE, false)
   if IsWindowVisible(HWND_FRMVDTOOLBOX)    then EnableWindow(HWND_FRMVDTOOLBOX, false)
   EnableWindow(HWND_FRMMAIN, false)
   function = 0
end function


' ========================================================================================
' Enable all modeless windows belonging to frmMain.
' ========================================================================================
public function EnableAllModeless() as long
   ' No need to enable/disable the modeless Help form.
   if IsWindowVisible(HWND_FRMFUNCTIONLIST) then EnableWindow(HWND_FRMFUNCTIONLIST, true)
   if IsWindowVisible(HWND_FRMFINDREPLACE)  then EnableWindow(HWND_FRMFINDREPLACE, true)
   if IsWindowVisible(HWND_FRMVDTOOLBOX)    then EnableWindow(HWND_FRMVDTOOLBOX, true)
   EnableWindow(HWND_FRMMAIN, true)
   function = 0
end function


' ========================================================================================
' Return temporary file name
' ========================================================================================
public FUNCTION GetTemporaryFilename( byref wszFolder as wstring, _
                                      BYREF wszExtension AS wSTRING _
                                      ) AS string
   dim wszTempFilename as wstring * MAX_PATH
   if GetTempFileName(@wszFolder, "TMP", 0, @wszTempFilename) then
      ' Delete the temp file that gets created b/c we will create it ourselves based on the 
      ' returned filename.
      kill wszTempFilename
      IF LEN(wszExtension) THEN 
         wszTempFilename = LEFT(wszTempFilename, LEN(wszTempFilename) -  3) & wszExtension
      end if
   end if      
   function = wszTempFilename
END FUNCTION


' ========================================================================================
' Replace a string in a combobox
' ========================================================================================
public FUNCTION ComboBox_ReplaceString( BYVAL hComboBox AS HWND, _
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
public Function GetFontCharSetID(ByREF wzCharsetName As CWSTR ) As Long

   If Len(wzCharsetName) = 0 Then Return SC_CHARSET_DEFAULT
   
   Select Case **wzCharsetName
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
public function RemoveDuplicateSpaces( byref sText as const string) as string
   dim as string st = sText
   do until instr(st, "  ") = 0
      st = AfxStrReplace(st, "  ", " ")
   loop   
   function = st
end function


' ========================================================================================
' Convert incoming text to proper case based on config setting. Used for autocomplete.
' ========================================================================================
public function ConvertCase( byval sText as string) as string

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
public FUNCTION Utf8ToAscii(byref strUtf8 AS STRING) AS STRING

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
public FUNCTION AnsiToUtf8(BYREF sAnsi AS STRING) AS STRING
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
public FUNCTION Utf8ToUnicode(BYREF ansiStr AS CONST STRING) AS STRING
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
public FUNCTION UnicodeToUtf8(byval pswzUnicode as wstring ptr) AS STRING
 dim sUtf8 AS STRING

 'Maps Unicode character string to a UTF-8 string.
 sUtf8 = string(LEN(*pswzUnicode) * 2, 0)
 WideCharToMultiByte(CP_UTF8, _                 'Set to UTF-8
                     0, _                       'Conversion type
                     cast(LPCWSTR, pswzUnicode), _  'Unicode string to convert
                     LEN(*pswzUnicode), _       'Length of Unicode string
                     cast(LPSTR, STRPTR(sUtf8)), _     'UTF-8 string
                     LEN(sUtf8), _              'Length of UTF-8 buffer
                     BYVAL 0, _                 'Invalid character replacement
                     BYVAL 0)                   'Replacement was used flag
 FUNCTION = rtrim(sUtf8, chr(0))

END FUNCTION


' ========================================================================================
' Open a disk file and read it into a string (ANSI or UTF8)
' ========================================================================================
public function GetFileToString( byref wszFilename as const wstring, _
                                 byref txtBuffer as string, _
                                 byval pDoc as clsDocument ptr _
                                 ) as boolean
   
   if pDoc = 0 then return true
   
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


   ' Check for BOM signatures
   if left(txtBuffer, 3) = chr(&HEF, &HBB, &HBF) THEN
      ' UTF8 BOM encoded 
      if pDoc then pDoc->FileEncoding = FILE_ENCODING_UTF8_BOM
      txtBuffer = mid(txtBuffer, 4)   ' bypass the BOM
   elseif left(txtBuffer, 2) = chr(&HFF, &HFE) THEN
      ' UTF16 BOM (little endian) encoded
      if pDoc then pDoc->FileEncoding = FILE_ENCODING_UTF16_BOM 
      txtBuffer = mid(txtBuffer, 3)   ' bypass the BOM
   else
      if pDoc then pDoc->FileEncoding = FILE_ENCODING_ANSI
   END IF

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
         ' the non-VD text is returned. We also skip over any codegen code.
         if pDoc then wszText = pDoc->ParseFormMetaData(HWND_FRMMAIN, wszText)

         ' Convert to UTF8 so it can display in the editor
         ' Need to pass a WSTRING pointer to the conversion function.
         txtBuffer = wszText.utf8

   END select
      
   function = false
END FUNCTION


' ========================================================================================
' Convert the current text buffer to the specified encoding and redisplay the text.
' ========================================================================================
public function ConvertTextBuffer( byval pDoc as clsDocument ptr, _
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
public function IsCurrentLineIncludeFilename() as Boolean
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

   dim as long idxBuild = ComboBox_GetCurSel(HWND_FRMMAIN_COMBOBUILDS)
   if idxBuild = -1 then idxBuild = 0
   if gConfig.Builds(idxBuild).Is32bit then wszCompilerPath = gConfig.FBWINcompiler32
   if gConfig.Builds(idxBuild).Is64bit then wszCompilerPath = gConfig.FBWINcompiler64
   wszCompilerPath = ProcessFromCurdrive(wszCompilerPath)
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
public function OpenSelectedDocument( byref wszFilename as wstring, _
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
   if len( wszFunctionName ) then
      ' Search for function, sub, or property (get/set)
      pData = gdb2.dbFindFunction( wszFunctionName, wszFilename) 
      if pData = 0 then pData = gdb2.dbFindSub( wszFunctionName, wszFilename) 
      if pData = 0 then pData = gdb2.dbFindProperty( wszFunctionName, wszFilename) 
      if pData then
         wszFilename = pData->fileName
         nLineNumber = pData->nLineNum
      end if
   end if
   
   ' Not all documents exist on disk file. For example, a QuickRun file will exist in
   ' the editor but may never have a disk footprint. We need to search the project to
   ' determine if the filename has a pDoc already associated with it. If it does, then
   ' pass that pDoc rather than looking for a disk filename.
   dim pDoc as clsDocument ptr
   pDoc = gApp.GetDocumentPtrByFilename( wszFilename )
   
   if pDoc then
      pDoc = frmMain_OpenFileSafely(HWND_FRMMAIN, _
                              False, _    ' bIsNewFile
                              False, _    ' bIsTemplate
                              true, _     ' bShowInTab
                              false, _    ' bIsInclude
                              "", _       ' pwszName
                              pDoc )      ' pDocIn
   else   
      if AfxFileExists(wszFilename) = 0 THEN exit function
      ' Display the document containing the selected sub/function       
      pDoc = frmMain_OpenFileSafely(HWND_FRMMAIN, _
                              False, _    ' bIsNewFile
                              False, _    ' bIsTemplate
                              true, _     ' bShowInTab
                              false, _    ' bIsInclude
                              wszFilename, _  ' pwszName
                              0 )         ' pDocIn
   end if

   ' Set the top line to display in the editor. I chose to start 3 lines before the
   ' function just to make it visually more appealing.
   if pDoc THEN
     
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
      end if
   END IF 
   
   function = pDoc
end function


' ========================================================================================
' Process prefix {CURDRIVE} and convert to current drive letter.
' ========================================================================================
public Function ProcessToCurdrive( Byval wzFilename As CWSTR ) As CWSTR
   ' For each folder location determine if it resides on the same drive as
   ' the project file. If it does then substitute the replaceable parameter
   ' {CURDRIVE} for the drive letter. This allows you to easily run the editor
   ' on different media (eg. thumb drive) that may be assigned a different
   ' drive letter.
   dim zText as CWSTR 
   if AfxFileExists(gApp.ProjectFilename) Then
      zText = gApp.ProjectFilename
   else
      zText = AfxGetExePathName
   end if      
   Dim wzCurDrive As CWSTR = LCase(Left(zText, 3))  ' eg. D:\

   If LCase(Left(wzFilename, 3)) = wzCurDrive Then 
      wzFilename = WSTR("{CURDRIVE}") & Mid(**wzFilename, 2)
   End If

   Return wzFilename
End Function


' ========================================================================================
' Process current drive to prefix {CURDRIVE} 
' ========================================================================================
public Function ProcessFromCurdrive( Byval wzFilename As CWSTR ) As CWSTR
   ' For each folder location determine if it resides on the same drive as
   ' the project file. If it does then substitute the replaceable parameter
   ' {CURDRIVE} for the drive letter. This allows you to easily run the editor
   ' on different media (eg. thumb drive) that may be assigned a different
   ' drive letter.
   dim zText as CWSTR 
   if AfxFileExists(gApp.ProjectFilename) Then
      zText = gApp.ProjectFilename
   else
      zText = AfxGetExePathName
   end if      
   If Ucase(Left(wzFilename, 10)) = WSTR("{CURDRIVE}") Then 
      wzFilename = Left(zText, 1) & Mid(wzFilename, 11)
   End If

   Return wzFilename
End Function


' ========================================================================================
' Adds an item to a Treeview
' ========================================================================================
public Function FF_TreeView_InsertItem( ByVal hWndControl As HWnd, _
                                        ByVal hParent As HANDLE, _
                                        ByRef TheText As WString, _
                                        ByVal lParam As LPARAM = 0, _
                                        ByVal iImage As Long = 0, _
                                        ByVal iSelectedImage As Long = 0 _
                                        ) As HANDLE
   Dim uInsert As TV_INSERTSTRUCT

   uInsert.hInsertAfter            = TVI_LAST
   uInsert.item.mask               = TVIF_TEXT Or TVIF_PARAM Or TVIF_SELECTEDIMAGE Or TVIF_IMAGE 
   uInsert.hParent                 = hParent
   uInsert.item.pszText            = Cast(WString Ptr, Strptr(TheText))
   uInsert.item.iImage             = iImage 
   uInsert.item.iSelectedImage     = iSelectedImage
   uInsert.item.lParam             = lParam 
   Function = TreeView_InsertItem( hWndControl, @uInsert )
    
End Function


' ========================================================================================
' Retrieve the lParam value from a Treeview
' ========================================================================================
public Function FF_TreeView_GetlParam( ByVal hWndControl As HWnd, _
                                       ByVal hItem As HANDLE _
                                       ) As Long
    Dim ti As TV_ITEM
    ti.hItem = hItem                                                 
    ti.mask  = TVIF_HANDLE Or TVIF_PARAM  
    TreeView_GetItem(hWndControl, Varptr(ti))
    Function = ti.lParam
End Function


' ========================================================================================
' Set the lParam value for a Treeview Item
' ========================================================================================
public Function FF_TreeView_SetlParam( ByVal hWndControl as HWnd, _
                                       ByVal hItem as HANDLE, _
                                       ByVal lParam as Long _
                                       ) as Long
   Dim ti as TV_ITEM
   ti.hItem  = hItem                                        
   ti.mask   = TVIF_PARAM
   ti.lParam = lParam
   If TreeView_SetItem( hWndControl, @ti ) = 0 Then Function = True
End Function


' ========================================================================================
' Displays the FileOpenDialog.
' The returned pointer must be freed with CoTaskMemFree
' ========================================================================================
public Function AfxIFileOpenDialogW( ByVal hwndOwner As HWnd, _
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
      case IDC_FRMUSERTOOLS_CMDBROWSEEXE   '1012
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

      Case IDC_FRMOPTIONSCOMPILER_CMDFBWIN32, IDC_FRMOPTIONSCOMPILER_CMDFBWIN64  '1003,1008
         rgFileTypes(1).pszName = @L(100,"FB Compiler")
         rgFileTypes(1).pszSpec = @WSTR("fbc*.exe")
         rgFileTypes(2).pszName = @L(230,"Executable")
         rgFileTypes(2).pszSpec = @WSTR("*.exe")
         rgFileTypes(3).pszName = @L(79,"All files")
         rgFileTypes(3).pszSpec = @WSTR("*.*")
         pofd->lpVtbl->SetFileTypes(pofd, 3, @rgFileTypes(1))
         ' Set the title of the dialog
         hr = pofd->lpVtbl->SetTitle(pofd, L(101,"Find FreeBASIC Compiler"))
   
      Case IDC_FRMOPTIONSLOCAL_CMDLOCALIZATION  '1012
         rgFileTypes(1).pszName = @L(102,"Localization files")
         rgFileTypes(1).pszSpec = @WSTR("*.lang")
         rgFileTypes(2).pszName = @L(79,"All files")
         rgFileTypes(2).pszSpec = @WSTR("*.*")
         pofd->lpVtbl->SetFileTypes(pofd, 2, @rgFileTypes(1))
         ' Set the title of the dialog
         hr = pofd->lpVtbl->SetTitle(pofd, L(103,"Open Localization File"))

      Case IDC_FRMOPTIONSCOMPILER_CMDFBHELPFILE, IDC_FRMOPTIONSCOMPILER_CMDAPIHELPPATH '1007,1011
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
public Function AfxIFileOpenDialogMultiple( ByVal hwndOwner As HWnd, _
                                            ByVal idButton As Long _
                                            ) As IShellItemArray Ptr

   ' Create an instance of the FileOpenDialog interface
   Dim hr As Long
   Dim pofd As IFileOpenDialog Ptr
   hr = CoCreateInstance(@CLSID_FileOpenDialog, Null, CLSCTX_INPROC_SERVER, @IID_IFileOpenDialog, @pofd)
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
   ' Set the default folder
   Dim pFolder As IShellItem Ptr
   SHCreateItemFromParsingName (AfxGetExePathName, Null, @IID_IShellItem, @pFolder)
   If pFolder Then
      pofd->lpVtbl->SetDefaultFolder(pofd, pFolder)
      pFolder->lpVtbl->Release(pFolder)
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
public Function AfxIFileSaveDialog( ByVal hwndOwner As HWnd, _
                                    ByVal pwszFileName As WString Ptr, _    ' full path and filename
                                    ByVal pwszDefExt As WString Ptr, _
                                    ByVal id As Long = 0, _
                                    ByVal sigdnName As SIGDN = SIGDN_FILESYSPATH _
                                    ) As WString Ptr

   ' // Create an instance of the IFileSaveDialog interface
   Dim hr As Long
   Dim psfd As IFileSaveDialog Ptr
   hr = CoCreateInstance(@CLSID_FileSaveDialog, Null, CLSCTX_INPROC_SERVER, @IID_IFileSaveDialog, @psfd)
   If psfd = Null Then Return Null

   dim as CWSTR wszFilename, wszFilePath 
   
   wszFilename = AfxStrPathname( "NAMEX", *pwszFileName )
   if AfxFileExists( *pwszFileName ) then
      wszFilePath = AfxStrPathName( "PATH", *pwszFileName ) 
   end if

   ' // Set the file types
   Dim rgFileTypes(1 To 4) As COMDLG_FILTERSPEC 
   
   Select Case id
      Case IDC_FRMPROJECTOPTIONS_CMDSELECT, IDM_PROJECTSAVE, IDM_PROJECTSAVEAS
         rgFileTypes(1).pszName = @L(216,"Project files")
         rgFileTypes(2).pszName = @L(79,"All files")
         rgFileTypes(1).pszSpec = @WSTR("*.wfbe")
         rgFileTypes(2).pszSpec = @WSTR("*.*")
         psfd->lpVtbl->SetFileTypes(psfd, 2, @rgFileTypes(1))
         ' // Set the title of the dialog
         hr = psfd->lpVtbl->SetTitle(psfd, L(185,"Save Project As..."))

      case IDM_FILESAVEDECLARES 
         rgFileTypes(1).pszName = @L(78,"FB Include files")
         rgFileTypes(2).pszName = @L(77,"FB code files")
         rgFileTypes(3).pszName = @L(79,"All files")
         rgFileTypes(1).pszSpec = @WSTR("*.bi;*.inc")
         rgFileTypes(2).pszSpec = @WSTR("*.bas")
         rgFileTypes(3).pszSpec = @WSTR("*.*")
         psfd->lpVtbl->SetFileTypes(psfd, 3, @rgFileTypes(1))
         psfd->lpVtbl->SetTitle(psfd, L(265,"Save Declares File"))

      Case IDC_FRMOPTIONSLOCAL_CMDNEW
         rgFileTypes(1).pszName = @L(102,"Localization files")
         rgFileTypes(2).pszName = @L(79,"All files")
         rgFileTypes(1).pszSpec = @WSTR("*.lang")
         rgFileTypes(2).pszSpec = @WSTR("*.*")
         psfd->lpVtbl->SetFileTypes(psfd, 2, @rgFileTypes(1))
         psfd->lpVtbl->SetTitle(psfd, L(8,"Save As..."))

      Case Else
         rgFileTypes(1).pszName = @L(77,"FB code files")
         rgFileTypes(2).pszName = @L(78,"FB Include files")
         rgFileTypes(3).pszName = @L(209,"Resource files")
         rgFileTypes(4).pszName = @L(79,"All files")
         rgFileTypes(1).pszSpec = @WSTR("*.bas")
         rgFileTypes(2).pszSpec = @WSTR("*.bi;*.inc")
         rgFileTypes(3).pszSpec = @WSTR("*.rc")
         rgFileTypes(4).pszSpec = @WSTR("*.*")
         psfd->lpVtbl->SetFileTypes(psfd, 4, @rgFileTypes(1))
         psfd->lpVtbl->SetTitle(psfd, L(8,"Save As..."))
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
' Enables the specified button in a toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
public Function FF_Toolbar_EnableButton( ByVal hToolBar As HWnd, _
                                         ByVal idButton As Long _
                                         ) As BOOLEAN
   ' Only enable the button if it is not already enabled. This should help reduce flicker.
   If SendMessage(hToolBar, TB_ISBUTTONENABLED, Cast(WPARAM, idButton), 0 ) = 0 Then
      Function = SendMessage(hToolBar, TB_ENABLEBUTTON, Cast(WPARAM, idButton), Cast(LPARAM, MAKELONG(CTRUE, 0)))
   End If   
End Function


' ========================================================================================
' Disables the specified button in a toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
public Function FF_Toolbar_DisableButton( ByVal hToolBar As HWnd, _
                                          ByVal idButton As Long _
                                          ) As BOOLEAN
   ' Only disable the button if it is not already disabled. This should help reduce flicker.
   If SendMessage(hToolBar, TB_ISBUTTONENABLED, Cast(WPARAM, idButton), 0 ) Then
      Function = SendMessage(hToolBar, TB_ENABLEBUTTON, Cast(WPARAM, idButton), Cast(LPARAM, MAKELONG(False, 0)))
   End If
End Function



' ========================================================================================
' Inserts an item at a specific location in the ListView.
' ========================================================================================
public Function FF_ListView_InsertItem( ByVal hWndControl As HWnd, _
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
public Function FF_ListView_GetItemText( ByVal hWndControl As HWnd, _
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
public Function FF_ListView_SetItemText( ByVal hWndControl As HWnd, _
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
public Function FF_ListView_GetlParam( ByVal hWndControl As HWnd, _
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
public Function FF_ListView_SetlParam( ByVal hWndControl As HWnd, _
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
public Function LoadLocalizationFile( Byref wszFileName As CWSTR, _
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
public Function GetProcessImageName( ByVal pe32w As PROCESSENTRY32W Ptr, _
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
public Function IsProcessRunning( ByVal pwszExeFileName As WString Ptr ) As BOOLEAN

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
public function GetRunExecutableFilename() as CWSTR
   ' Used by the top menu and toolbar to determine if the Run Executable options
   ' should be grayed out.
   dim wszFilename as CWSTR
   dim wszOptions as CWSTR
   Dim pDocMain As clsDocument Ptr 

   dim as long idxBuild = ComboBox_GetCurSel(HWND_FRMMAIN_COMBOBUILDS)
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
            wszFilename = mid(**wszOptions, y, i-y-1)
            exit for
         end if   
      NEXT
      ' There was no other switch so we made it to the end of the string
      if len(wszFilename) = 0 then wszFilename = mid(**wszOptions, y)
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
public FUNCTION LoadPNGfromRes( BYVAL hInstance AS HINSTANCE, _
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
' Parse the incoming pDoc clsDocument.
' ========================================================================================
public Function ParseDocument( byval pDoc as clsDocument ptr ) As Long

   If pDoc = 0 Then exit function
   if pDoc->bNeedsParsing = false then exit function


   dim as long nLineCount, nLineNum, nPos, nStartParseLine
   Dim As String st, st_ucase, sLine
   
   dim wszFilename as wstring * MAX_PATH
   wszFilename = pDoc->DiskFilename 


   ' The incoming document is already loaded into a clsDocument so we can simply 
   ' use that pDoc to retrieve the lines. We can use that approach for the actual
   ' editor lines but in order to parse any Form CodeGen lines then we need to 
   ' load them into an array.
   
   ' Dynamic string array to hold all lines loaded from the file. It
   ' is sized at a reasonable beginning capacity in order to avoid
   ' too many memory reallocations if constantly redim'd. 
   redim sLineArray( any ) as string 
   
   ' Boolean to ensure that we check to parse any Form codegen code. Once it is
   ' checked and parsed then we set the boolean to true to exit the loop.
   dim as boolean bParsingCodeGen = false
   
   do

      if (gApp.IsProjectLoading) orelse (gApp.IsFileLoading) then
         gApp.wszPanelText = wszFilename
         gApp.hIconPanel = 0
         frmMain_SetStatusbar
      END IF

      ' Create a parse TYPE that handles the state information throughout parsing this document.
      dim parser as clsParser
      
      with parser
         .action   = ACTION_NONE
         .fileName = wszFilename
         ' Flag whether we are parsing the codegen code. We do nit want to save any
         ' of the Form Types sub/function/ctor/dtor because those will then display
         ' in the Explorer Treeview. The user would click on them expecting to be
         ' able to then see them in the editor. We can't allow that.
         .bParsingCodeGen = bParsingCodeGen
      end with

        
      if bParsingCodeGen then
         ' We put the codegen code lines after the scintilla lines because this
         ' ensures that nStartParseLine is > 0 then the codegen code will get
         ' reparsed when the main scintilla document is modified. Otherwise, 
         ' we'd have to gdb2.dbDelete all entries in the database and have to
         ' reparse the whole file when at times we only need to parse from the
         ' line currently being modified onwards to the end of the file.
         nLineCount = ubound( sLineArray )
         nStartParseLine = lbound( sLineArray )
      else
         nLineCount = pDoc->GetLineCount - 1  
         nStartParseLine = pDoc->ParseStartLine
      
         ' Before starting to parse this document we must remove any previously saved
         ' data that exists in the gdb2 database for this file. This will also delete
         ' any form codegen parsed structures because line numbers for those structures
         ' are after the main scintilla code.
         gdb2.dbDelete( wszFilename, nStartParseLine )      
      end if
      
      ' Analyze all lines in the file
      for nLineNum = nStartParseLine to nLineCount 
         parser.st       = ""
         parser.st_ucase = ""
         sLine           = ""   ' reset our full line string
         
         ' IMPORTANT: Get the full line as it could be across multiple lines via the line continuation underscore.
         do
            if bParsingCodeGen THEN
               ' Parsing the codegen array
               st = sLineArray( nLineNum )
            else
               ' Parsing the lines in the scintilla control
               st = pDoc->GetLine( nLineNum )
            END IF
            
            parser.lineNum = nLineNum 
                
            ' Check for a TODO entry. Do not normalize line until after this check is done because
            ' TODO's are found as part of a comment.
            if parser.parseToDoItem(st) THEN continue for
            
            ' Mask out all characters inside of a string. This enables us to correctly
            ' test for embedded comment characters and multiple line ":" separators.
            st = trim(MaskStringCharacters(st))   ' important to remove all pre/post spaces

            ' Check for multiline comments 
            if parser.IsMultilineComment(st) THEN continue for
                   
            ' Remove any single line comments
            st = trim(RemoveComments(st))
            
            if right(st, 2) = " _" THEN
               sLine = sLine & left(st, len(st) - 1) & " "
               nLineNum = nLineNum + 1
            else
               ' no line continuation character so we now have all the characters for the current line.   
               sLine = sLine & st 
               exit do
            end if
         loop
         
         parser.lineNum = nLineNum 
       
         ' If the line contains ":" characters then the line needs to be broken up
         ' into as many sublines to be processed. Don't use AfxStrParse because we
         ' need to manually check whether the substring was actually created because
         ' of a trailing ":". Using AfxStrParse and adding a trailing ":" will cause
         ' parse problems.
         for nPos = 1 to len(sLine)
            ' Get the subline but ensure that we add back the ":" because it 
            ' could be part of a Label or identifier like PRIVATE: or PUBLIC:
            dim AS LONG i = instr(nPos, sLine, ":")
            parser.st = trim(AfxStrExtract(nPos, sLine, ":"))
            if i THEN parser.st = parser.st & ":"
            nPos = iif(i, i + 1, len(sLine) + 1)
               
            ' Skip blank lines for speed
            if len(parser.st) = 0 THEN continue for
            parser.st_ucase = Ucase(parser.st)

            ' If normalized line is len = 0 then returns TRUE to continue to next line.
            ' This removes double spaces and also stores the UCASE version of the line.
            if parser.NormalizeLine() then continue for 
            
            ' Continue to parse a TYPE if applicable
            if parser.parseTYPE() then continue for

            ' Continue to parse an ENUM if applicable
            if parser.parseENUM() then continue for

            ' Test the line to see if we have encountered any start or end points
            ' for a parse. For example, starting to parse a TYPE or FUNCTION, or
            ' ending an existing parse of a TYPE or FUNCTION. If we have completed/ended
            ' a parse action then function returns TRUE and we can advance to the next line.
            if parser.InspectLine() then continue for
            
            ' Parse the line for any DIM'd variables. This call parses for variables 
            ' regardless whether they are global/local, inside a sub/function, or as
            ' part of a TYPE structure. The current parser state will dictate how the
            ' variable information is saved to the database.
            if parser.parseVariableDefinitions() then continue for

         NEXT
      next

      ' Check to see if this is a visual design form. If it is, then we need
      ' to parse the codegen string.
      if bParsingCodeGen = false then
         bParsingCodeGen = true
         if pDoc->IsDesigner then
            ' Load the codegen lines into the sLineArray
            dim as CWSTR wst
            dim as long iLineStart = 1
            dim as long iLineEnd, nNextLine

            ' Convert the CodeGen string into an array that can be parsed.
            nNextLine = nLineCount
            do until iLineStart >= len( pDoc->wszFormCodeGen )
               iLineEnd = instr(iLineStart, pDoc->wszFormCodeGen, vbcrlf)
               if iLineEnd = 0 then iLineEnd = len( pDoc->wszFormCodeGen )  ' cr/lf not found
               wst = mid( pDoc->wszFormCodeGen.wstr, iLineStart, iLineEnd - iLineStart )
               iLineStart = iLineStart + len(wst) + len(vbcrlf)
               if nNextLine >= ubound(sLineArray) THEN
                  redim preserve sLineArray( nLineCount to nLineCount + 5000 )
               END IF
               sLineArray(nNextLine) = wst   
               nNextLine = nNextLine + 1
            loop
            continue do
         end if
      end if
   
      exit do
   loop 
   

   ' Document has now been fully parsed
   pDoc->bNeedsParsing = false
   pDoc->ParseStartLine = 0

   ' Update the Explorer treeview nodes to ensure that any new nodes are added
   ' and any old nodes are deleted.
   if pDoc->hNodeExplorer then
      AddFunctionsToExplorerTreeview( pDoc, true )
   end if
   
   ' Update the TODO listview in the frmOutput form
   frmOutput_UpdateToDoListview
   
   Function = 0
 
End Function


' ========================================================================================
' Check PlanetSquires server for latest WinFBE version.
' ========================================================================================
public function DoCheckForUpdates( byval hWndParent as hwnd, _
                                   byval bSilentCheck as Boolean = false _
                                   ) as long
   
   ''  Contact the PlanetSquires server and download the text file containing
   ''  the latest WinFBE version number. If an update exists then ask the user
   ''  if they wish to navigate to the Releases page on GitHub.
   ''
   dim as CWSTR wszServerFile = "winfbe_version.txt"
   dim as CWSTR wszFilename
   dim as CWSTR wszLatestVersion
   dim as CWSTR wszMsg
   dim as string st
    
   DIM pWHttp AS CWinHttpRequest
   
   wszFilename = AfxGetExePathName + wszServerFile

   ' Open an HTTP connection to an HTTP resource
   if pWHttp.Open( "GET", "https://www.planetsquires.com/" & wszServerFile ) = S_OK then

      ' Send HTTP request and wait 5 seconds for response 
      pWHttp.Send
    
      if pWHttp.WaitForResponse(5) then
         st = pWHttp.GetResponseBody
          
         ' Open a file stream and save the downloaded file
         DIM pFileStream AS CFileStream
         IF pFileStream.Open(wszFilename, STGM_CREATE OR STGM_WRITE) = S_OK then
            pFileStream.WriteTextA(st)
         END IF
      end if
   end if
   
   ' Open the downloaded file and check for the version number. It is possible that the socket
   ' will open but fail and send request but security parameters of the system will prevent data
   ' from being received resulting in a zero byte file. 
   if AfxFileExists( wszFilename ) then
      dim as long f = freefile
      open wszFilename for input as #f
       
      do until eof(f)
         line input #f, st
         if left( st, 15 ) = "latest_version=" then
            wszLatestVersion = trim(mid( st, 16 ))
            exit do
         end if
      loop
          
      close #f
      AfxDeleteFile( wszFilename ) 
   end if
   
   if len(wszLatestVersion) = 0 then
      if bSilentCheck = false then      
         MessageBox( hWndParent, _
                     L(92,"Failed to retrieve update information"), _
                     L(94,"Software Update"), _
                     MB_ICONINFORMATION Or MB_OK )
      end if
      exit function
   end if
      
      
   ' Check the installed vs. available version numbers   
   dim as long latestVersion = ConvertWinFBEversion(wszLatestVersion)
   dim as long installedVersion = ConvertWinFBEversion(APPVERSION)

   ' Save the config file so that other editor instances will not also do update checks again
   gConfig.LastUpdateCheck = JulianDateNow
   gConfig.SaveConfigFile
   
   if bSilentCheck = false then      
      if installedVersion = latestVersion then
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
public function GetListBoxEmptyClientArea( byval hListBox as HWND, _
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
   
   return rc
end function