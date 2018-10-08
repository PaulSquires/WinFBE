'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2018 Paul Squires, PlanetSquires Software
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


' ========================================================================================
' Return temporary file name
' ========================================================================================
FUNCTION GetTemporaryFilename( byref wszFolder as wstring, BYREF wszExtension AS wSTRING) AS string
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
FUNCTION ComboBox_ReplaceString (BYVAL hComboBox AS HWND, BYVAL index AS LONG, BYVAL pwszNewText AS WSTRING PTR, BYVAL pNewData AS LONG_PTR = 0) AS LONG
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


'' ========================================================================================
'' Simple Bubblesort to sort function names before displaying them
'' ========================================================================================
'Sub BubbleSort( MyArray() As String )
'   dim as long i, j
'   For i = 1 To UBound(MyArray)-1
'      For j = i To UBound(MyArray)
'         If MyArray(i) > MyArray(j) Then Swap MyArray(i), MyArray(j) 
'      Next 
'   Next 
'End Sub

'' ========================================================================================
'' Simple Insertion sort to sort function names before displaying them (case insensitive)
'' ========================================================================================
'Sub InsertionSort( MyArray() As String )
'   dim as long lb = LBound(myArray)
'   dim as long ub = Ubound(myArray)
'   for startIndex as long = lb + 1 to ub 'establish starting index of each pass
'      for index as long = startIndex to lb + 1 step -1
'         if (ucase(myArray(index)) < ucase(myArray(index-1))) then   'REVERSE TO SORT DECENDING
'            swap myArray(index), myArray(index-1)
'         else
'            exit for
'         end if
'      next 
'   next 
'end sub


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
   DIM sUnicode AS STRING 
   DIM sAnsi AS STRING = ""
   DIM dwLen AS DWORD = MultiByteToWideChar(CP_UTF8, 0, STRPTR(strUtf8),LEN(strUtf8), NULL, 0)
   IF dwLen THEN
      sUnicode = string(dwLen * 2, 0)
      dwLen=MultiByteToWideChar(CP_UTF8, _                         'System default Windows UTF8 code page
                           0, _                                    'Conversion type
                           STRPTR(strUtf8), _                      'UTF8 string to convert
                           LEN(strUtf8), _                         'Lenght of UTF8 string
                           cast(WSTRING PTR, STRPTR(sUnicode)), _  'Unicode string
                           dwLen*2)                                'Lenght of Unicode buffer
      IF dwLen THEN
         dwLen=WideCharToMultiByte(CP_ACP, 0,cast(WSTRING PTR, STRPTR(sUnicode)), _
                                   -1, NULL, 0 , BYVAL 0, BYVAL 0)
         IF dwLen THEN
            sAnsi = string(dwLen, 0)
            dwLen=WideCharToMultiByte(CP_ACP, _                          'Set to ANSI
                                 0, _                                    'Conversion type
                                 cast(WSTRING PTR, STRPTR(sUnicode)), _  'Unicode string to convert
                                 LEN(sUnicode) / 2, _                    'Lenght of Unicode string
                                 STRPTR(sAnsi), _                        'ANSI string
                                 dwLen, _                                'Length of ANSI buffer
                                 BYVAL 0, _                              'Invalid character replacement
                                 BYVAL 0)                                'Replacement was used flag
         END IF
      END IF
   END IF
   FUNCTION = rtrim(sAnsi, chr(0))
END FUNCTION

' ========================================================================================
' Maps Ansi character string to a UTF-8 string.
' ========================================================================================
FUNCTION AnsiToUtf8(BYREF sAnsi AS STRING) AS STRING
 dim sUnicode AS STRING
 dim sUtf8    AS STRING = ""

 'Maps Ansi character string to a UTF-8 string.

 'Step one, convert to UNICODE
   DIM dwLen AS DWORD = MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, STRPTR(sAnsi), LEN(sAnsi), NULL, 0)
   IF dwLen THEN
      sUnicode = string(dwLen*2, 0)
      dwLen=MultiByteToWideChar(CP_ACP, _                     'System default Windows ANSI code page
                           MB_PRECOMPOSED, _                  'Conversion type
                           cast(LPCSTR, STRPTR(sAnsi)), _     'ANSI string to convert
                           LEN(sAnsi), _                      'Lenght of ANSI string
                           cast(LPWSTR, STRPTR(sUnicode)), _  'Unicode string
                           dwLen*2)                           'Lenght of Unicode buffer

      'Step two, convert to UTF-8
      IF dwLen THEN
         dwLen=WideCharToMultiByte(CP_UTF8, 0, cast(WSTRING PTR, STRPTR(sUnicode)), -1, NULL, 0, BYVAL 0, BYVAL 0)
         IF dwLen THEN
            sUtf8 = string(dwLen, 0)
            dwLen=WideCharToMultiByte(CP_UTF8, _                     'Set to UTF-8
                                 0, _                                'Conversion type
                                 cast(LPCWSTR, STRPTR(sUnicode)), _  'Unicode string to convert
                                 LEN(sUnicode) / 2, _                'Lenght of Unicode string
                                 cast(LPSTR, STRPTR(sUtf8)), _       'UTF-8 string
                                 dwLen, _                            'Length of UTF-8 buffer
                                 BYVAL 0, _                          'Invalid character replacement
                                 BYVAL 0)                            'Replacement was used flag
         END IF
      END IF
   END IF
   FUNCTION = rtrim(sUtf8, chr(0))

END FUNCTION


' ========================================================================================
' Maps UTF-8 string to Unicode character string 
' ========================================================================================
FUNCTION Utf8ToUnicode(BYREF ansiStr AS CONST STRING) AS STRING
   DIM sUnicode AS STRING = ""
   DIM dwLen AS DWORD = MultiByteToWideChar(CP_UTF8, 0, STRPTR(ansiStr), LEN(ansiStr), NULL, 0)
   IF dwLen THEN
      sUnicode = string(dwLen * 2, 0)
      dwLen = MultiByteToWideChar(CP_UTF8, 0, STRPTR(ansiStr), LEN(ansiStr), CAST(WSTRING PTR, STRPTR(sUnicode)), dwLen * 2)
   END IF
   FUNCTION = rtrim(sUnicode, chr(0))
end function

   
' ========================================================================================
' Maps Unicode character string to a UTF-8 string.
' ========================================================================================
FUNCTION UnicodeToUtf8(byval pswzUnicode as wstring ptr) AS STRING
 dim sUtf8 AS STRING = ""

 'Maps Unicode character string to a UTF-8 string.
   DIM dwLen AS DWORD = WideCharToMultiByte(CP_UTF8, 0,cast(WSTRING PTR, pswzUnicode), LEN(*pswzUnicode), NULL, 0 , BYVAL NULL, BYVAL NULL)
   IF dwLen THEN
       sUtf8 = string(dwLen * 2, 0)
       WideCharToMultiByte(CP_UTF8, _                     'Set to UTF-8
                           0, _                           'Conversion type
                           cast(LPCWSTR, pswzUnicode), _  'Unicode string to convert
                           LEN(*pswzUnicode), _           'Length of Unicode string
                           cast(LPSTR, STRPTR(sUtf8)), _  'UTF-8 string
                           dwLen*2, _                     'Length of UTF-8 buffer
                           BYVAL 0, _                     'Invalid character replacement
                           BYVAL 0)                       'Replacement was used flag
   END IF
   FUNCTION = rtrim(sUtf8, chr(0))

END FUNCTION

' ========================================================================================
' Return if string is UTF8 encoded
' ========================================================================================
'function isutf8( byval _string as  zstring ptr ) as boolean
'     if _string = 0 then return 0
'    if left(*_string,2)= wchr(&hFF,&hFE) or left(*_string,2)=wchr(&hFE,&hFF) then print "utf16":exit function 'utf-16 boms
'     if left(*_string,3)=wchr(&hEF,&hBB,&hBF) then return 1  'utf-8 bom
'   dim bytes as const ubyte ptr = cptr(const ubyte ptr, _string)
'   while *bytes
'      if (((bytes[0] = &h09) orelse (bytes[0] = &h0A)) orelse (bytes[0] = &h0D)) orelse ((&h20 <= bytes[0]) andalso (bytes[0] <= &h7E)) then
'         bytes += 1
'         continue while
'      end if
'      if ((&hC2 <= bytes[0]) andalso (bytes[0] <= &hDF)) andalso ((&h80 <= bytes[1]) andalso (bytes[1] <= &hBF)) then
'         bytes += 2
'         continue while
'      end if
'      if ((((bytes[0] = &hE0) andalso ((&hA0 <= bytes[1]) andalso (bytes[1] <= &hBF))) andalso ((&h80 <= bytes[2]) andalso (bytes[2] <= &hBF))) orelse ((((((&hE1 <= bytes[0]) andalso (bytes[0] <= &hEC)) orelse (bytes[0] = &hEE)) orelse (bytes[0] = &hEF)) andalso ((&h80 <= bytes[1]) andalso (bytes[1] <= &hBF))) andalso ((&h80 <= bytes[2]) andalso (bytes[2] <= &hBF)))) orelse (((bytes[0] = &hED) andalso ((&h80 <= bytes[1]) andalso (bytes[1] <= &h9F))) andalso ((&h80 <= bytes[2]) andalso (bytes[2] <= &hBF))) then
'         bytes += 3
'         continue while
'      end if
'      if (((((bytes[0] = &hF0) andalso ((&h90 <= bytes[1]) andalso (bytes[1] <= &hBF))) andalso ((&h80 <= bytes[2]) andalso (bytes[2] <= &hBF))) andalso ((&h80 <= bytes[3]) andalso (bytes[3] <= &hBF))) orelse (((((&hF1 <= bytes[0]) andalso (bytes[0] <= &hF3)) andalso ((&h80 <= bytes[1]) andalso (bytes[1] <= &hBF))) andalso ((&h80 <= bytes[2]) andalso (bytes[2] <= &hBF))) andalso ((&h80 <= bytes[3]) andalso (bytes[3] <= &hBF)))) orelse ((((bytes[0] = &hF4) andalso ((&h80 <= bytes[1]) andalso (bytes[1] <= &h8F))) andalso ((&h80 <= bytes[2]) andalso (bytes[2] <= &hBF))) andalso ((&h80 <= bytes[3]) andalso (bytes[3] <= &hBF))) then
'         bytes += 4
'         continue while
'      end if
'      return 0
'   wend
'   return 1
'end function

' ========================================================================================
' Return text description of the specified file encoding
' ========================================================================================
function FileEncodingTextDescription(byval FileEncoding as long) as CWSTR

   select case FileEncoding
      case FILE_ENCODING_UTF8_BOM:  RETURN "UTF-8 (BOM)"   
      case FILE_ENCODING_UTF16_BOM: RETURN "UTF-16 (BOM)"   
      case else
         return "ANSI"    ' includes FILE_ENCODING_ANSI
   END select

end function


' ========================================================================================
' Open a disk file and read it into a string (ANSI or UTF8)
' ========================================================================================
function GetFileToString( byref wszFilename as const wstring, _
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
      pDoc->FileEncoding = FILE_ENCODING_UTF8_BOM
      txtBuffer = mid(txtBuffer, 4)   ' bypass the BOM
   elseif left(txtBuffer, 2) = chr(&HFF, &HFE) THEN
      ' UTF16 BOM (little endian) encoded
      pDoc->FileEncoding = FILE_ENCODING_UTF16_BOM 
      txtBuffer = mid(txtBuffer, 3)   ' bypass the BOM
   else
      pDoc->FileEncoding = FILE_ENCODING_ANSI
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
         wszText = pDoc->ParseFormMetaData(HWND_FRMMAIN, wszText)

         ' Convert to UTF8 so it can display in the editor
         ' Need to pass a WSTRING pointer to the conversion function.
         'txtBuffer = UnicodeToUtf8( wszText )
         txtBuffer = wszText.utf8

   END select
      
   function = false
END FUNCTION


' ========================================================================================
' Convert the current text buffer to the specified encoding and redisplay the text.
' ========================================================================================
function ConvertTextBuffer( byval pDoc as clsDocument ptr, _
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
' Generic open document handler for when Function ListBox item selected or Explorer Treeview
' ========================================================================================
function OpenSelectedDocument( byval pDoc as clsDocument ptr, _
                               byref wszName as WSTRING, _
                               ByVal hCurrentNode As HTREEITEM =0) as long

   ' Open the document based on the filename. If the Header line is clicked
   ' then simply open the document and point to the first line. If a function 
   ' is clicked then search the gFunctionDeclares hash to get the filename
   ' and line number position within the file.
   dim wszText as wstring * MAX_PATH
   dim nLineNumber as LONG
   dim as long idx = gApp.GetActiveProjectIndex()
   if pDoc THEN  ' Header node
      wszText = pDoc->DiskFilename
      nLineNumber = 0
      idx=pDoc->ProjectIndex
   else 
      dim pData as DB2_DATA ptr
      if hCurrentNode then
         Dim As Hwnd hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
         if CheckIsDoc4ExplorerTreeview(hTree,hCurrentNode)=false then
            Dim As LPARAM lParam = FF_TreeView_GetlParam(hTree,hCurrentNode)
            pData = cast(DB2_DATA ptr , lParam)
            if pData=0 then
               idx=GetProjectIndex4ExplorerTreeview(hTree,hCurrentNode)
               if idx<0 then idx = gApp.GetActiveProjectIndex()
               gdb2.dbRewind4Proj(idx)
               pData=gdb2.dbFindTYPE(wszName,ELF_CurrentProject,1)
            end if
         end if
      else
         gdb2.dbRewind4Proj(idx) 
         pData = gdb2.dbFindFunction( wszName,ELF_CurrentProject,1 )
         if pData = 0 then pData = gdb2.dbFindSub( wszName,ELF_CurrentProject,1 )
      end if 
      if pData then
         IF pData->deleted=FALSE THEN
            if pData->pFileDATA then
               wszText = pData->pFileDATA->fileName
            endif
            nLineNumber = pData->nStartLineNum
         ENDIF
      end if
   end if   

   ' Not all documents exist on disk file. For example, a QuickRun file will exist in
   ' the editor but may never have a disk footprint. We need to search the project to
   ' determine if the filename has a pDoc already associated with it. If it does, then
   ' pass that pDoc rather than looking for a disk filename.
   pDoc = gApp.Projects(idx).GetDocumentPtrByFilename(wszText)
   gdb2.dbRewind4File(idx,wszText)
   if pDoc then
      pDoc = frmMain_OpenFileSafely(HWND_FRMMAIN, _
                              False, _    ' bIsNewFile
                              False, _    ' bIsTemplate
                              true, _     ' bShowInTab
                              false, _    ' bIsInclude
                              "", _       ' pwszName
                              pDoc )      ' pDocIn
   else   
      if AfxFileExists(wszText) = 0 THEN exit function
      ' Display the document containing the selected sub/function       
      pDoc = frmMain_OpenFileSafely(HWND_FRMMAIN, _
                              False, _    ' bIsNewFile
                              False, _    ' bIsTemplate
                              true, _     ' bShowInTab
                              false, _    ' bIsInclude
                              wszText, _  ' pwszName
                              0 )         ' pDocIn
   end if

   ' Set the top line to display in the editor. I chose to start 3 lines before the
   ' function just to make it visually more appealing.
   if pDoc THEN
      dim as hwnd hEdit = pDoc->hWndActiveScintilla
      SciExec( hEdit, SCI_SETFIRSTVISIBLELINE, nLineNumber - 3, 0) 
      SciExec( hEdit, SCI_GOTOLINE, nLineNumber, 0) 
   END IF
   
   function = 0
end function


' ========================================================================================
' Process prefix {CURDRIVE} and convert to current drive letter.
' ========================================================================================
Function ProcessToCurdrive( ByRef wzFilename As CWSTR ) As CWSTR
   ' For each folder location determine if it resides on the same drive as
   ' the project file. If it does then substitute the replaceable parameter
   ' {CURDRIVE} for the drive letter. This allows you to easily run the editor
   ' on different media (eg. thumb drive) that may be assigned a different
   ' drive letter.
   dim zText as CWSTR 
   dim idx as long = gApp.GetActiveProjectIndex
   if AfxFileExists(gApp.Projects(idx).ProjectFilename) Then
      zText = gApp.Projects(idx).ProjectFilename
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
Function ProcessFromCurdrive( ByRef wzFilename As CWSTR ) As CWSTR
   ' For each folder location determine if it resides on the same drive as
   ' the project file. If it does then substitute the replaceable parameter
   ' {CURDRIVE} for the drive letter. This allows you to easily run the editor
   ' on different media (eg. thumb drive) that may be assigned a different
   ' drive letter.
   dim zText as CWSTR 
   dim idx as long = gApp.GetActiveProjectIndex
   if AfxFileExists(gApp.Projects(idx).ProjectFilename) Then
      zText = gApp.Projects(idx).ProjectFilename
   else
      zText = AfxGetExePathName
   end if      
   If Ucase(Left(wzFilename, 10)) = WSTR("{CURDRIVE}") Then 
      wzFilename = Left(zText, 1) & Mid(wzFilename, 11)
   End If

   Return wzFilename
End Function


' ========================================================================================
' Removes checkbox from a node in a Treeview
' ========================================================================================
function Treeview_RemoveCheckBox( byval hTree as hwnd, byval hNode as HTREEITEM) as long 
   
   dim as TVITEM tvi
   tvi.hItem = hNode
   tvi.mask = TVIF_STATE
   tvi.stateMask = TVIS_STATEIMAGEMASK
   tvi.state = 0
   function = TreeView_SetItem(hTree, @tvi)

END FUNCTION


' ========================================================================================
' Adds an item to a Treeview
' ========================================================================================
Function FF_TreeView_InsertItem( ByVal hWndControl As HWnd, _
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
Function FF_TreeView_GetlParam( ByVal hWndControl As HWnd, _
                                ByVal hItem As HANDLE _
                                ) As Long

    Dim ti As TV_ITEM

    ti.hItem = hItem                                                 
    ti.mask  = TVIF_HANDLE Or TVIF_PARAM  
    TreeView_GetItem(hWndControl, Varptr(ti))

    Function = ti.lParam
End Function

' ========================================================================================
' Retrieve the iImage value from a Treeview
' ========================================================================================
Function FF_TreeView_GetiImage( ByVal hWndControl As HWnd, _
                                ByVal hItem As HANDLE _
                                ) As Long

    Dim ti As TV_ITEM

    with ti
       .hItem = hItem                                                 
       .mask  = TVIF_IMAGE  
       TreeView_GetItem(hWndControl, Varptr(ti))

       Function = .iImage 
    end with
End Function

' ========================================================================================
' Returns True/False state for whether node is Expanded
' ========================================================================================
Function FF_TreeView_SetCheckState( ByVal hWndControl As HWnd, _
                                    ByVal hItem As HANDLE, _
                                    ByVal fCheck As Boolean _
                                    ) As BOOLEAN

   Dim ti As TV_ITEM

   ti.hItem = hItem                                                 
   ti.mask  = TVIF_HANDLE Or TVIF_STATE
   ti.stateMask = TVIS_STATEIMAGEMASK
   
   ' Image 0: Blank image
   ' Image 1: Closed image
   ' Image 2: Open image
 
   ti.state = INDEXTOSTATEIMAGEMASK( Iif(fCheck, 2, 1)  )
                    
   ' Returns a nonzero value if successful, or zero otherwise.                  
   Function = TreeView_SetItem(hWndControl, @ti)

End Function

' ========================================================================================
' Returns True/False state for whether node is Expanded
' ========================================================================================
Function FF_TreeView_GetCheckState( ByVal hWndControl As HWnd, _
                                    ByVal hItem As HANDLE _
                                    ) As BOOLEAN

   Dim ti As TV_ITEM

   ti.hItem = hItem                                                 
   ti.mask  = TVIF_HANDLE Or TVIF_STATE
   ti.stateMask = TVIS_STATEIMAGEMASK
   
   TreeView_GetItem(hWndControl, @ti)

   Return (ti.state Shr 12) - 1
End Function


' ========================================================================================
' Set the lParam value for a Treeview Item
' ========================================================================================
Function FF_TreeView_SetlParam (ByVal hWndControl as HWnd, _
                                ByVal hItem as HANDLE, _
                                ByVal lParam as Long) as Long

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
Function AfxIFileOpenDialogW( ByVal hwndOwner As HWnd, ByVal idButton As Long) As WString Ptr

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
         rgFileTypes(2).pszName = @L(77,"FB code files")
         rgFileTypes(3).pszName = @L(78,"FB Include files")
         rgFileTypes(4).pszName = @L(209,"Resource files")
         rgFileTypes(5).pszName = @L(79,"All files")
         rgFileTypes(1).pszSpec = @WSTR("*.bas;*.bi;*.inc;*.rc")
         rgFileTypes(2).pszSpec = @WSTR("*.bas")
         rgFileTypes(3).pszSpec = @WSTR("*.bi;*.inc")
         rgFileTypes(4).pszSpec = @WSTR("*.rc")
         rgFileTypes(5).pszSpec = @WSTR("*.*")
         pofd->lpVtbl->SetFileTypes(pofd, 5, @rgFileTypes(1))
         ' Set the title of the dialog
         hr = pofd->lpVtbl->SetTitle(pofd, L(80,"Insert File"))

      Case IDC_FRMOPTIONSCOMPILER_CMDFBWIN32, IDC_FRMOPTIONSCOMPILER_CMDFBWIN64
         rgFileTypes(1).pszName = @L(100,"FB Compiler")
         rgFileTypes(1).pszSpec = @WSTR("fbc*.exe")
         rgFileTypes(2).pszName = @L(230,"Executable")
         rgFileTypes(2).pszSpec = @WSTR("*.exe")
         rgFileTypes(3).pszName = @L(79,"All files")
         rgFileTypes(3).pszSpec = @WSTR("*.*")
         pofd->lpVtbl->SetFileTypes(pofd, 3, @rgFileTypes(1))
         ' Set the title of the dialog
         hr = pofd->lpVtbl->SetTitle(pofd, L(101,"Find FreeBASIC Compiler"))
   
      Case IDC_FRMOPTIONSLOCAL_CMDLOCALIZATION
         rgFileTypes(1).pszName = @L(102,"Localization files")
         rgFileTypes(1).pszSpec = @WSTR("*.lang")
         rgFileTypes(2).pszName = @L(79,"All files")
         rgFileTypes(2).pszSpec = @WSTR("*.*")
         pofd->lpVtbl->SetFileTypes(pofd, 2, @rgFileTypes(1))
         ' Set the title of the dialog
         hr = pofd->lpVtbl->SetTitle(pofd, L(103,"Open Localization File"))

      Case IDC_FRMOPTIONSCOMPILER_CMDFBHELPFILE, IDC_FRMOPTIONSCOMPILER_CMDAPIHELPPATH
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
Function AfxIFileOpenDialogMultiple( ByVal hwndOwner As HWnd, ByVal sigdnName As SIGDN = SIGDN_FILESYSPATH) As IShellItemArray Ptr

   ' Create an instance of the FileOpenDialog interface
   Dim hr As Long
   Dim pofd As IFileOpenDialog Ptr
   hr = CoCreateInstance(@CLSID_FileOpenDialog, Null, CLSCTX_INPROC_SERVER, @IID_IFileOpenDialog, @pofd)
   If pofd = Null Then Return Null

   ' Set the file types
   Dim rgFileTypes(1 To 5) As COMDLG_FILTERSPEC
   rgFileTypes(1).pszName = @L(106,"Open source file")
   rgFileTypes(2).pszName = @L(77,"FB code files")
   rgFileTypes(3).pszName = @L(78,"FB Include files")
   rgFileTypes(4).pszName = @L(209,"Resource files")
   rgFileTypes(5).pszName = @L(79,"All files")
   rgFileTypes(1).pszSpec = @WSTR("*.bas;*.bi;*.inc;*.rc")
   rgFileTypes(2).pszSpec = @WSTR("*.bas")
   rgFileTypes(3).pszSpec = @WSTR("*.bi;*.inc")
   rgFileTypes(4).pszSpec = @WSTR("*.rc")
   rgFileTypes(5).pszSpec = @WSTR("*.*")
   pofd->lpVtbl->SetFileTypes(pofd, 5, @rgFileTypes(1))

   ' Set the title of the dialog
   hr = pofd->lpVtbl->SetTitle(pofd, L(106,"Open source file"))
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
Function AfxIFileSaveDialog( ByVal hwndOwner As HWnd, _
                             ByVal pwszFileName As WString Ptr, _
                             ByVal pwszDefExt As WString Ptr, _
                             ByVal id As Long = 0, _
                             ByVal sigdnName As SIGDN = SIGDN_FILESYSPATH _
                             ) As WString Ptr

   ' // Create an instance of the IFileSaveDialog interface
   Dim hr As Long
   Dim psfd As IFileSaveDialog Ptr
   hr = CoCreateInstance(@CLSID_FileSaveDialog, Null, CLSCTX_INPROC_SERVER, @IID_IFileSaveDialog, @psfd)
   If psfd = Null Then Return Null

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
   hr = psfd->lpVtbl->SetFileName(psfd, pwszFileName)
   ' // Set the extension
   hr = psfd->lpVtbl->SetDefaultExtension(psfd, pwszDefExt)
   ' // Display the dialog
   hr = psfd->lpVtbl->Show(psfd, hwndOwner)

   ' // Set the default folder
   Dim pFolder As IShellItem Ptr
   SHCreateItemFromParsingName (AfxGetExePathName, Null, @IID_IShellItem, @pFolder)
   If pFolder Then
      psfd->lpVtbl->SetDefaultFolder(psfd, pFolder)
      pFolder->lpVtbl->Release(pFolder)
   End If

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
Function FF_Toolbar_EnableButton (ByVal hToolBar As HWnd, ByVal idButton As Long) As BOOLEAN
   ' Only enable the button if it is not already enabled. This should help reduce flicker.
   If SendMessage(hToolBar, TB_ISBUTTONENABLED, Cast(WPARAM, idButton), 0 ) = 0 Then
      Function = SendMessage(hToolBar, TB_ENABLEBUTTON, Cast(WPARAM, idButton), Cast(LPARAM, MAKELONG(CTRUE, 0)))
   End If   
End Function
' ========================================================================================

' ========================================================================================
' Disables the specified button in a toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
Function FF_Toolbar_DisableButton (ByVal hToolBar As HWnd, ByVal idButton As Long) As BOOLEAN
   ' Only disable the button if it is not already disabled. This should help reduce flicker.
   If SendMessage(hToolBar, TB_ISBUTTONENABLED, Cast(WPARAM, idButton), 0 ) Then
      Function = SendMessage(hToolBar, TB_ENABLEBUTTON, Cast(WPARAM, idButton), Cast(LPARAM, MAKELONG(False, 0)))
   End If
End Function


'   Inserts an item at a specific location in the ListView.
'
'   hWndControl:    Handle of control (e.g. HWND_FORM1_LISTVIEW1)   
'   iRow:           The zero-based row to insert this item into.
'   iColumn:        The zero-based column to insert this item into.
'   TheText:        A string representing the text to display.   
'   lParam:         A 32-bit value to be stored with the item. (optional. Only for Col0)
'            
'   Returns:        If adding an item at Column 0 then the function
'                   will return the index of the new item. If inserting
'                   to other than Column 0 then the function returns
'                   TRUE if successful, FALSE if otherwise.
'
Function FF_ListView_InsertItem( ByVal hWndControl As HWnd, _
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


'   Retrieves the text of a ListView item.
'
'   hWndControl:  Handle of control (e.g. HWND_FORM1_LISTVIEW1)   
'   iRow:         The zero-based row to retrieve.
'   iColumn:      The zero-based column to retrieve.
'   pwszText:     Pointer to wString to hold the retrieved text,
'   nTextMax:     Size of text buffer pointed to by pwszText.           
'   Returns:      Returns TRUE if successful, FALSE otherwise.
'
Function FF_ListView_GetItemText( ByVal hWndControl As HWnd, _
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
Function FF_ListView_SetItemText( ByVal hWndControl As HWnd, _
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
' Retrieve the lParam value from a Listview
' ========================================================================================
Function FF_ListView_GetlParam( ByVal hWndControl As HWnd, _
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
' Load a .lang localization file from disk and populate the localization array
' ========================================================================================
Function LoadLocalizationFile( Byref wszFileName As CWSTR ) As BOOLEAN

   ' default that the file failed to load
   Function = False
   If AfxFileExists(wszFileName) = 0 Then Exit Function

   Dim as CBSTR wst, wKey, wData
   Dim nKey  As Long
   Dim nData As Long  
   Dim i     As Long
    
   dim pStream AS CTextStream
   if pStream.OpenUnicode(wszFileName) <> S_OK then exit function
   
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
         ReDim LL(nData) As WString * MAX_PATH
      Else
         ' this should be a key/value pair line in the format:
         ' 00001:value
         ' Ensure that we add the value to the array within the valid
         ' boundaries of the array.
         If (nKey >= LBound(LL)) AndAlso (nKey <= Ubound(LL)) Then
            ' Remove any comments from end of the line. Comments begin with
            ' a semicolon character.
            LL(nKey) = rtrim(AfxStrParse(wData, 1, ";"), any chr(9,32))
         End If
      End If   
         
   Loop
   pStream.Close
   
   Function = True
End Function


' ========================================================================================
' Get the full process image name
' ========================================================================================
Function GetProcessImageName( ByVal pe32w As PROCESSENTRY32W Ptr, _
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

''
''
function GetRunExecutableFilename() as CWSTR
   ' Determine the EXE name of the currently active document or project. Used
   ' by the top menu and toolbar to determine if the Run Executable options
   ' should be grayed out.
   dim wszFilename as CWSTR
   dim wszOptions as CWSTR
   Dim pDocMain As clsDocument Ptr 

   dim as long idx = gApp.GetActiveProjectIndex()
   dim as long idxBuild = ComboBox_GetCurSel(HWND_FRMMAIN_COMBOBUILDS)
   dim as Boolean bInString = false
   
   pDocMain = iIf( gApp.IsProjectActive, gApp.Projects(idx).GetMainDocumentPtr, gTTabCtl.GetActiveDocumentPtr)
   If pDocMain = 0 Then return ""
   
   if idxBuild > -1 then 
      wszOptions = gConfig.Builds(idxBuild).wszOptions 

      if gApp.IsProjectActive THEN
         if gConfig.Builds(idxBuild).Is32bit then
            wszOptions = wszOptions + " " + gApp.Projects(idx).ProjectOther32
         end if   
         if gConfig.Builds(idxBuild).Is64bit then
            wszOptions = wszOptions + " " + gApp.Projects(idx).ProjectOther64
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
      wszFilename= AfxStrMid(wszOptions,"""","""",y)
   END IF
   ' Has a filename been determined yet? If it has then it possibly does not have
   ' a path assigned to it so we should add one.
   if len(wszFilename) THEN      
      if gApp.IsProjectActive THEN   
         gApp.Projects(idx).Format2FullPath(wszFilename)
      ELSEif len(AfxStrPathname("PATH", wszFilename)) = 0 THEN
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
             
''
''
FUNCTION LoadPNGfromRes(BYVAL hInstance AS HINSTANCE, BYREF wszImageName AS WSTRING) as any ptr
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

Function WStrIsEqual(Byref lpWord As wString,Byref lpList As wString,ByVal fMatchCase As BOOLEAN=FALSE) As BOOLEAN
   Function = false
   If fMatchCase Then
      If strcmpW(lpWord,lpList)=0 Then 
         Function = TRUE
      end if
   ElseIf strcmpiW(lpWord,lpList)=0 Then 
      Function = TRUE
   EndIf
End Function

Function WStrNIsEqual(Byref lpWord As wString,Byref lpList As wString,Byref Size As long,ByVal fMatchCase As BOOLEAN=FALSE) As BOOLEAN
   Function = false
   If fMatchCase Then
      If wcsncmp(@lpWord,@lpList,Size)=0 Then 
         Function = TRUE
      end if
   ElseIf _wcsnicmp(@lpWord,@lpList,Size)=0 Then 
      Function = TRUE
   EndIf
End Function

Function StrIsEqual(Byref lpWord As String,Byref lpList As String,ByVal fMatchCase As BOOLEAN=FALSE) As BOOLEAN
   Function = false
   If fMatchCase Then
      If strcmp(lpWord,lpList)=0 Then 
         Function = TRUE
      end if
   ElseIf strcmpi(lpWord,lpList)=0 Then 
      Function = TRUE
   EndIf
End Function

Function StrNIsEqual(Byref lpWord As String,Byref lpList As String,Byref Size As long,ByVal fMatchCase As BOOLEAN=FALSE) As BOOLEAN
   Function = false
   If fMatchCase Then
      If strncmp(lpWord,lpList,Size)=0 Then 
         Function = TRUE
      end if
   ElseIf strncmpi(lpWord,lpList,Size)=0 Then 
      Function = TRUE
   EndIf
End Function   
   
FUNCTION AfxStrRemoveWithMark (BYREF wszMainStr AS CONST WSTRING, BYREF wszDelim1 AS CONST WSTRING, BYREF wszDelim2 AS CONST WSTRING, BYREF MarkKeys AS CONST WSTRING, BYVAL fRemoveAll AS BOOLEAN = FALSE, BYVAL IsInstrRev AS BOOLEAN = FALSE) AS CWSTR
   DIM nPos1 AS LONG =iif(IsInstrRev=true, INSTRRev(wszMainStr, wszDelim1), INSTR(wszMainStr, wszDelim1))
   IF nPos1 = 0 THEN RETURN wszMainStr
   DIM nPos2 AS LONG = INSTR(nPos1 + LEN(wszDelim1), wszMainStr, wszDelim2)
   IF nPos2 = 0 THEN RETURN wszMainStr
   nPos2 += LEN(wszDelim2)
   DIM nLen AS LONG = nPos2 - nPos1
   IF fRemoveAll = FALSE THEN 
      RETURN MID(wszMainStr, 1, nPos1 - 1)& MarkKeys & MID(wszMainStr, nPos2)
   else
      RETURN AfxStrRemoveWithMark(MID(wszMainStr, 1, nPos1 - 1) & MarkKeys & MID(wszMainStr, nPos2), wszDelim1, wszDelim2,MarkKeys, fRemoveAll,IsInstrRev)  
   end if
END FUNCTION
   
FUNCTION AfxStrMid (BYREF wszMainStr AS CONST WSTRING, BYREF wszDelim1 AS CONST WSTRING, BYREF wszDelim2 AS CONST WSTRING, BYVAL Index AS long = 1) AS CWSTR
   DIM nPos1 AS LONG = INSTR(Index,wszMainStr, wszDelim1)
   IF nPos1   THEN
      nPos1 += LEN(wszDelim1)
      DIM nPos2 AS LONG = INSTR(nPos1, wszMainStr, wszDelim2)
      IF nPos2 THEN 
         RETURN MID(wszMainStr,nPos1, nPos2-nPos1)
      end if
   end if 
   RETURN ""
END FUNCTION