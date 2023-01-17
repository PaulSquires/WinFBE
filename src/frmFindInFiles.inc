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

#include once "frmFindInFiles.bi"
#include once "frmOutput.bi"
#include once "clsDocument.bi"

''
''  FINDSTRING.BAS
''     Search for string in one or more files or folders.
''

#define unicode
#include once "windows.bi"
#include once "Afx/CFindFile.inc"
#include once "Afx/CTextStream.inc"

USING Afx


''
''
function FindInFiles_ProcessFile( _
            byref pStreamResults as CTextStream, _
            byref wszFindText as wstring, _
            byref wszFilename as wstring, _
            byval bMatchCase  as Boolean, _
            byval bWholeWords as Boolean _
            ) as Long

   if AfxFileExists( wszFilename ) = false then exit function
   
   dim as long nLineNum = 0
   dim as long nFoundPos, nStartPos 
   dim as Boolean bUnicode = AfxIsFileUnicode( wszFilename )
   dim as Boolean bReadingForm = false
   
   dim pStream as CTextStream  

   if pStream.Open( wszFilename, IOMODE_FORREADING, false, bUnicode ) = S_OK then
   
      dim as CWSTR wszLine, wszResult, wszLookLine
      
      if bMatchCase = false then wszFindText = ucase(wszFindText)
      
      Do Until pStream.EOS
         wszLine = pStream.ReadLine

         ' Only unicode files can be Form files. We need to bypass all Form meta data
         ' and code generation.
         if (bUnicode = true) and (nLineNum = 0) then
            if left(wszLine, 13) = "' WINFBE FORM" then
               bReadingForm = true
            end if
         end if
         if bReadingForm = true then  
            if left(wszLine, 20) = "' WINFBE_CODEGEN_END" then
               bReadingForm = false
               continue do
            end if
         end if
            
         if bReadingForm then continue do

         nLineNum = nLineNum + 1
         
         if len(wszLine) >= len(wszFindText) then
            wszLookLine = iif(bMatchCase, wszLine, ucase(wszLine))

            nFoundPos = instr( 1, wszLookLine, wszFindText ) 
            do until nFoundPos = 0    
               nStartPos = nFoundPos + len(wszFindText)
               if bWholeWords then
                  ' Only give a positive match if the character before and after the match
                  ' position is not an alphanumeric or space.
                  dim as long chLeftChar  = asc(wszLine, nFoundPos - 1)
                  dim as long chRightChar = asc(wszLine, nFoundPos + len(wszFindText))
                  
                  ' If word is enclosed by spaces then it is okay. 
                  if ( chLeftChar = 32 ) andalso ( chRightChar = 32 ) then 
                     ' okay
                  elseif IsCharAlphaNumeric( chLeftChar ) orelse _
                         IsCharAlphaNumeric( chRightChar ) then _
                         nFoundPos = 0
                  end if   
               end if
            
               ' If we still have a valid result after doing the bWholeWords test then
               ' output the result to the results text file.
               if nFoundPos > 0 then   
                  wszResult = wszFilename & ":" & nLineNum & ":" & wszLine
                  pStreamResults.WriteLine wszResult
                  ' We have found a match within the line. No need to continue searching
                  ' on this line because we don't want multiple copies outputted to the file.
                  exit do
               end if   

               nFoundPos = instr( nStartPos, wszLookLine, wszFindText ) 
            loop   
         end if
         
      Loop                        
      pStream.Close

   else
      ' ? "ERROR: "; wszFilename, bUnicode, pStream.GetErrorInfo
   end if
   
   function = 0
end function


''
''
function FindInFiles_RecursiveFileSearch( _
            byref pStreamResults as CTextStream, _
            byref wszFindText as wstring, _
            byref wszFilePath as wstring, _
            wszFileSpecs() as CWSTR, _
            byval bSubFolders as Boolean, _
            byval bMatchCase  as Boolean, _
            byval bWholeWords as Boolean _
            ) as Long
                              
   DIM pFinder AS CFindFile
   
   IF pFinder.FindFile( wszFilePath & "\*.*" ) = S_OK THEN
      DO
         IF pFinder.IsDots THEN   
            ' skip
         elseif pFinder.IsFolder then
            if bSubFolders then
               FindInFiles_RecursiveFileSearch( _
                           pStreamResults, _
                           wszFindText, _
                           pFinder.FilePath, _
                           wszFileSpecs(), _ 
                           bSubFolders, _
                           bMatchCase, _
                           bWholeWords _
                           )
            end if
         else
            dim as Boolean bProcessFile = false
            dim as CWSTR wszExt
            
            for i as long = lbound(wszFileSpecs) to ubound(wszFileSpecs)
               bProcessFile = false
               ' Check if the file spec is a wildcard or a matching filename.
               if instr( wszFileSpecs(i), "*" ) then
                  wszExt = ucase(ltrim(AfxStrPathName("EXTN", wszFileSpecs(i)), "."))
                  IF (UCASE(pFinder.FileExt) = wszExt) orelse (wszExt = "*") THEN
                     ' bypass know binary files
                     select case UCASE(pFinder.FileExt)
                        case "EXE", "DLL", "GIF", "PNG", "JPG", "JPEG", "BMP", "ICO", _
                             "TIFF", "CHM", "CHW", "CUR", "XLS", "XLSX", "PDF", "ZIP", _
                             "SVG", "RAR"
                           bProcessFile = false
                        case else   
                           bProcessFile = true
                     end select
                  end if
               else
                  ' Check if the filename matches.
                  if ucase(wszFileSpecs(i)) = ucase(pFinder.FilePath) then
                     bProcessFile = true
                  end if
               end if
               
               if bProcessFile then
                  FindInFiles_ProcessFile( _
                                 pStreamResults, _
                                 wszFindText, _
                                 pFinder.FilePath, _
                                 bMatchCase, _
                                 bWholeWords _
                                 )
                  exit for
               END IF
            next
         
         END IF

         IF pFinder.FindNext = 0 THEN EXIT DO
      LOOP
   END IF
   pFinder.Close

   function = 0
end function



' ========================================================================================
' Set the states of the various controls
' ========================================================================================
Function frmFindInFiles_SetControlStates() As Long

   Button_SetCheck( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFINDINFILES_CHKWHOLEWORDS),  gFindInFiles.nWholeWord )
   Button_SetCheck( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFINDINFILES_CHKMATCHCASE),   gFindInFiles.nMatchCase )
   Button_SetCheck( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFINDINFILES_CHKSUBFOLDERS),  gFindInFiles.nSearchSubFolders )

   Button_SetCheck( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFINDINFILES_CHKCURRENTDOC),  gFindInFiles.nSearchCurrentDoc )
   Button_SetCheck( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFINDINFILES_CHKALLOPENDOCS), gFindInFiles.nSearchAllOpenDocs )
   Button_SetCheck( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFINDINFILES_CHKPROJECT),     gFindInFiles.nSearchProject )
   
   dim as long bEnable = true
   
   if (gFindInFiles.nSearchCurrentDoc = BST_CHECKED) orelse _
      (gFindInFiles.nSearchAllOpenDocs = BST_CHECKED) orelse _
      (gFindInFiles.nSearchProject = BST_CHECKED ) then
      bEnable = false   
   end if
   
   EnableWindow( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFINDINFILES_COMBOFILES), bEnable )    
   EnableWindow( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFINDINFILES_COMBOFOLDER), bEnable )    
   EnableWindow( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFINDINFILES_CMDFILES), bEnable )    
   EnableWindow( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFINDINFILES_CMDFOLDERS), bEnable )    
   EnableWindow( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFINDINFILES_CHKSUBFOLDERS), bEnable )    
   
   EnableWindow( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFINDINFILES_CHKCURRENTDOC), _
                 iif( gTTabCtl.GetActiveDocumentPtr, true, false) )     
   EnableWindow( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFINDINFILES_CHKALLOPENDOCS), _
                 iif( gTTabCtl.GetItemCount, true, false) )     
   EnableWindow( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFINDINFILES_CHKPROJECT), _
                 gApp.IsProjectActive )     

   function = 0
end function


' ========================================================================================
' Add a string to the Find combobox (ensure no duplicates)
' ========================================================================================
Function frmFindInFiles_AddToFindCombo( ByRef wszText As String ) As Long
   Dim as long i
   dim as long nFound = -1
   
   wszText = trim(wszText)
   if len(wszText) = 0 THEN exit function
   ' See if the string already exists in the combobox
   for i = 0 to ubound(gFindInFiles.txtFindCombo)
      if gFindInFiles.txtFindCombo(i) = wszText THEN
         nFound = i: exit for
      END IF
   NEXT
   if nFound = -1 THEN  ' was not found
      for i = ubound(gFindInFiles.txtFindCombo) to 1 step -1
         gFindInFiles.txtFindCombo(i) = gFindInFiles.txtFindCombo(i-1)
      next   
      gFindInFiles.txtFindCombo(0) = wszText
   END IF
   gFindInFiles.txtFind = wszText
   
   Function = 0
End Function


' ========================================================================================
' Add a string to the Files combobox (ensure no duplicates)
' ========================================================================================
Function frmFindInFiles_AddToFilesCombo( ByRef wszText As wString ) As Long
   Dim as long i
   dim as long nFound = -1
   
   ' Array is current 0 to 10 dimension
   wszText = trim(wszText)
   if len(wszText) = 0 THEN exit function
   ' See if the string already exists in the combobox
   for i = 0 to ubound(gFindInFiles.txtFilesCombo)
      if gFindInFiles.txtFilesCombo(i) = wszText THEN
         nFound = i: exit for
      END IF
   NEXT
   if nFound = -1 THEN  ' was not found
      ' Move all entries down onw and add the text at the beginning of the list.
      for i = ubound(gFindInFiles.txtFilesCombo) to 1 step -1
         gFindInFiles.txtFilesCombo(i) = gFindInFiles.txtFilesCombo(i-1)
      next   
      gFindInFiles.txtFilesCombo(0) = wszText
   END IF
   gFindInFiles.txtFiles = wszText

   Function = 0
End Function


' ========================================================================================
' Add a string to the Folders combobox (ensure no duplicates)
' ========================================================================================
Function frmFindInFiles_AddToFolderCombo( ByRef wszText As wString ) As Long
   Dim as long i
   dim as long nFound = -1
   
   wszText = trim(wszText)
   if len(wszText) = 0 THEN exit function
   ' See if the string already exists in the combobox
   for i = 0 to ubound(gFindInFiles.txtFolderCombo)
      if gFindInFiles.txtFolderCombo(i) = wszText THEN
         nFound = i: exit for
      END IF
   NEXT
   if nFound = -1 THEN  ' was not found
      for i = ubound(gFindInFiles.txtFolderCombo) to 1 step -1
         gFindInFiles.txtFolderCombo(i) = gFindInFiles.txtFolderCombo(i-1)
      next   
      gFindInFiles.txtFolderCombo(0) = wszText
   END IF
   gFindInFiles.txtFolder = wszText

   Function = 0
End Function


' ========================================================================================
' File search procedure (shell out to builtin Windows "findstr" utility
' ========================================================================================
function frmFindInFiles_DoFindInFilesEx() as LONG
   Dim ShExecInfo As SHELLEXECUTEINFOW  
   dim as CWSTR DQ = wchr(34)
   
   dim as HCURSOR hCurSave = GetCursor()
   SetCursor( LoadCursor(0, IDC_WAIT) )

   ' Ensure that the text in the find textbox has been added to the combobox. Need to do
   ' this to ensure that manually typed in text has been captured.
   Dim wszFindText As WString * MAX_PATH
   wszFindText = AfxGetWindowText( GetDlgItem(HWnd_FRMFINDINFILES, IDC_FRMFINDINFILES_COMBOFIND))
   if len(wszFindText) = 0 THEN
      SetFocus( GetDlgItem(HWnd_FRMFINDINFILES, IDC_FRMFINDINFILES_COMBOFIND) )
      exit function      
   END IF
   gFindInFiles.txtFind = wszFindText
   frmFindInFiles_AddToFindCombo(gFindInFiles.txtFind)

   
   ' What folder to start searching in
   gFindInFiles.txtFolder = trim(AfxGetWindowText( GetDlgItem(HWND_FRMFINDINFILES,IDC_FRMFINDINFILES_COMBOFOLDER) ))
   if len(gFindInFiles.txtFolder) = 0 THEN 
      gFindInFiles.txtFolder = AfxGetExePathName
      AfxSetWindowText( GetDlgItem(HWND_FRMFINDINFILES,IDC_FRMFINDINFILES_COMBOFOLDER), gFindInFiles.txtFolder )
   end if
   frmFindInFiles_AddToFolderCombo(gFindInFiles.txtFolder)
   if right(gFindInFiles.txtFolder,1) <> "\" then gFindInFiles.txtFolder = gFindInFiles.txtFolder & "\"
   dim wszFolder as CWSTR = gFindInFiles.txtFolder 


   ' What files are we searching in
   dim as BOOLEAN bInString
   dim as long ub
   dim as CWSTR wszFiles, wszSearchFiles, wszFilename
   
   wszFiles = trim(AfxGetWindowText( GetDlgItem(HWnd_FRMFINDINFILES, IDC_FRMFINDINFILES_COMBOFILES)))
   gFindInFiles.txtFiles = wszFiles
   frmFindInFiles_AddToFilesCombo( gFindInFiles.txtFiles )
   
   dim wszFileSpecs(any) as CWSTR
   ' Examples:
   '     dim wszFileSpecs(2) as CWSTR
   '     wszFileSpecs(0) = "*.bas"
   '     wszFileSpecs(1) = "*.bi"
   '     wszFileSpecs(2) = "*.inc"
   '   
   '     dim wszFileSpecs(0) as CWSTR
   '     wszFileSpecs(0) = "D:\WinFBE_Suite-Editor\sample.bas"

   ' Look at the files to determine if multiple files or file types need to be searched for.
   if instr( wszFiles, DQ ) THEN
      ' In order to better parse the string for the filenames, replace any spaces within the
      ' filenames with special character (Ascii(8)); then parse; and finally replace the special character.
      for i as long = 0 to len(wszFiles) - 1
         select case wszFiles[i]
            case 34      ' double quote
               bInString = not bInString 
            case 32      ' space
               if bInString THEN wszFiles[i] = 8   
         end select
      NEXT
      
      ' Clean the string to remove any duplicate remaining spaces
      wszFiles = AfxStrShrink(wszFiles)

      ' This is a list of double quoted filenames to search
      dim as long nCount = AfxStrParseCount(wszFiles, " ")

      for i as long = 1 to nCount
         wszFilename = trim( AfxStrParse(wszFiles, i, " "), DQ )
         ' Put back any special characters to spaces again
         wszFilename = AfxStrReplace(wszFilename, wchr(8), " " )
         ' Add it to our list of wszFileSpecs
         ub = ubound(wszFileSpecs) + 1
         redim preserve wszFileSpecs( ub )
         wszFileSpecs(ub) = wszFolder & wszFilename
      NEXT   

   else
      ' This is a list of patterns to use (*.bas *.bi) etc
      dim as long nCount = AfxStrParseCount( wszFiles, " " )
      for i as long = 1 to nCount
         ub = ubound(wszFileSpecs) + 1
         redim preserve wszFileSpecs( ub )
         wszFileSpecs(ub) = wszFolder & AfxStrParse( wszFiles, i, " " )
      NEXT   
   END IF
      

   ' Create a results file to capture the results
   dim pStreamResults as CTextStream  ' (create)(utf16)
   dim as CWSTR wszSearchResults = AfxGetExePathName & "_searchresults.txt"
   if pStreamResults.Create( wszSearchResults, true, true) = S_OK then 

      dim as Boolean bWholeWord  = iif( gFindInFiles.nWholeWord = BST_CHECKED, true, false )
      dim as Boolean bMatchCase  = iif( gFindInFiles.nMatchCase = BST_CHECKED, true, false )
      dim as Boolean bSubFolders = iif( gFindInFiles.nSearchSubFolders = BST_CHECKED, true, false )
      
      if (gFindInFiles.nSearchProject = BST_CHECKED) or _
         (gFindInFiles.nSearchCurrentDoc = BST_CHECKED) or _
         (gFindInFiles.nSearchAllOpenDocs = BST_CHECKED) then 

         Dim pDoc As clsDocument Ptr = gApp.pDocList
         do until pDoc = 0
            wszFilename = ""
            
            if gFindInFiles.nSearchProject = BST_CHECKED then
               wszFilename = pDoc->DiskFilename
            end if   
            
            if gFindInFiles.nSearchAllOpenDocs = BST_CHECKED then 
               if gTTabCtl.GetTabIndexFromFilename( pDoc->DiskFilename ) <> -1 then
                  wszFilename = pDoc->DiskFilename
               end if   
            end if
            
            if gFindInFiles.nSearchCurrentDoc = BST_CHECKED then
               if pDoc = gTTabCtl.GetActiveDocumentPtr() then
                  wszFilename = pDoc->DiskFilename
               end if
            end if

            FindInFiles_ProcessFile( _
                           pStreamResults, _
                           gFindInFiles.txtFind, _
                           wszFilename, _
                           bMatchCase, _
                           bWholeWord _
                           )
                  
            pDoc = pDoc->pDocNext
         loop   

      else
         FindInFiles_RecursiveFileSearch( _
                        pStreamResults, _
                        gFindInFiles.txtFind, _
                        gFindInFiles.txtFolder, _
                        wszFileSpecs(), _
                        bSubFolders, _      ' Search subfolders
                        bMatchCase, _       ' Match case
                        bWholeWord _        ' Whole words
                        )
      end if
      
      pStreamResults.Close
   else
      ' ? "Error creating _FindResults.txt file."
   end if   

   frmOutput_UpdateSearchListview( wszSearchResults )
   
   SetCursor( hCurSave )
      
   function = 0
end function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmFindInFiles
' ========================================================================================
Function frmFindInFiles_OnCommand( _
            ByVal HWnd As HWnd, _
            ByVal id As Long, _
            ByVal hwndCtl As HWnd, _
            ByVal codeNotify As UINT _
            ) As LRESULT

   Select Case id
      Case IDOK  ' Find Next
         If codeNotify = BN_CLICKED Then
            EnableWindow(hwndCtl, False)
            frmFindInFiles_DoFindInFilesEx()
            EnableWindow(hwndCtl, true)
            Exit Function
         End If  
         
      Case IDCANCEL  ' Close
         If codeNotify = BN_CLICKED Then
            SendMessage( HWnd, WM_CLOSE, 0, 0)
            Exit Function
         End If

      Case IDC_FRMFINDINFILES_COMBOFIND
         ' Catch the dropdown of the Find combobox so we can populate it
         If codeNotify = CBN_DROPDOWN Then
            ' Add any previous search terms to the combobox
            Dim As Long i
            Dim wszText As WString * MAX_PATH
            wszText = AfxGetWindowText(hwndCtl)
            frmFindInFiles_AddToFindCombo(Str(wszText))
            ComboBox_ResetContent(hwndCtl)
            For i = 0 To ubound(gFindInFiles.txtFindCombo)
               If Len(gFindInFiles.txtFindCombo(i)) Then 
                  wszText = gFindInFiles.txtFindCombo(i)
                  ComboBox_AddString( hwndCtl, @wszText )
               end if   
            Next
            wszText = gFindInFiles.txtFind
            i = ComboBox_FindStringExact( hwndCtl, -1, @wszText )
            ComboBox_SetCurSel(hwndCtl, i)
         End If

      Case IDC_FRMFINDINFILES_COMBOFILES
         ' Catch the dropdown of the Files combobox so we can populate it
         If codeNotify = CBN_DROPDOWN Then
            ' Add any previous search terms to the combobox
            Dim As Long i
            Dim wszText As WString * MAX_PATH
            wszText = AfxGetWindowText(hwndCtl)
            frmFindInFiles_AddToFilesCombo(Str(wszText))
            ComboBox_ResetContent(hwndCtl)
            For i = 0 To ubound(gFindInFiles.txtFilesCombo)
               If Len(gFindInFiles.txtFilesCombo(i)) Then 
                  wszText = gFindInFiles.txtFilesCombo(i)
                  ComboBox_AddString( hwndCtl, @wszText )
               end if   
            Next
            wszText = gFindInFiles.txtFiles
            i = ComboBox_FindStringExact( hwndCtl, -1, @wszText )
            ComboBox_SetCurSel(hwndCtl, i)
         End If
      
      Case IDC_FRMFINDINFILES_COMBOFOLDER
         ' Catch the dropdown of the Files combobox so we can populate it
         If codeNotify = CBN_DROPDOWN Then
            ' Add any previous search terms to the combobox
            Dim As Long i
            Dim wszText As WString * MAX_PATH
            wszText = AfxGetWindowText(hwndCtl)
            frmFindInFiles_AddToFolderCombo(Str(wszText))
            ComboBox_ResetContent(hwndCtl)
            For i = 0 To ubound(gFindInFiles.txtFolderCombo)
               If Len(gFindInFiles.txtFolderCombo(i)) Then 
                  wszText = gFindInFiles.txtFolderCombo(i)
                  ComboBox_AddString( hwndCtl, @wszText )
               end if   
            Next
            wszText = gFindInFiles.txtFolder
            i = ComboBox_FindStringExact( hwndCtl, -1, @wszText )
            ComboBox_SetCurSel(hwndCtl, i)
         End If

      Case IDC_FRMFINDINFILES_CMDFILES
         If codeNotify = BN_CLICKED Then
            ' Display the Open File Dialog
            Dim pItems As IShellItemArray Ptr = AfxIFileOpenDialogMultiple(HWnd, IDM_FILEOPEN)
            If pItems = Null Then Exit Function
            Dim dwItemCount As Long, i As Long, idx As Long, pItem As IShellItem Ptr, pwszName As WString Ptr
            pItems->lpVtbl->GetCount(pItems, @dwItemCount)
            
            dim wszFolder as CWSTR
            dim wszFiles as CWSTR

            For i = 0 To dwItemCount - 1
               pItems->lpVtbl->GetItemAt(pItems, i, @pItem)
               If pItem Then
                  pItem->lpVtbl->GetDisplayName(pItem, SIGDN_FILESYSPATH, @pwszName)
                  If pwszName Then 
                     wszFolder = AfxStrPathName("PATH", *pwszName)
                     AfxSetWindowText( GetDlgItem(HWND_FRMFINDINFILES,IDC_FRMFINDINFILES_COMBOFOLDER), wszFolder)
                     wszFiles = wszFiles + Chr(34) + AfxStrPathName("NAMEX", *pwszName) + chr(34, 32)
                     CoTaskMemFree(pwszName)
                     pwszName = Null
                  End If
                  pItem->lpVtbl->Release(pItem)
                  pItem = Null
               End If
            Next
            AfxSetWindowText( GetDlgItem(HWND_FRMFINDINFILES,IDC_FRMFINDINFILES_COMBOFILES), wszFiles)
            gFindInFiles.txtFiles = str(wszFiles)
            pItems->lpVtbl->Release(pItems)
            Exit Function
         End If

      Case IDC_FRMFINDINFILES_CMDFOLDERS
         If codeNotify = BN_CLICKED Then
            static wszStartFolder as CWSTR
            dim wszFolder as CWSTR
            if len(wszStartFolder) = 0 then wszStartFolder = AfxGetExePathName
            wszFolder = AfxBrowseForFolder( HWND_FRMFINDINFILES, L(261,"Choose the folder to search in:"), wszStartFolder )
            if len(wszFolder) then
               AfxSetWindowText( GetDlgItem(HWND_FRMFINDINFILES,IDC_FRMFINDINFILES_COMBOFOLDER), wszFolder)
               gFindInFiles.txtFolder = str(wszFolder)
            end if
            Exit Function
         End If

      Case IDC_FRMFINDINFILES_CHKMATCHCASE
         If codeNotify = BN_CLICKED Then
            gFindInFiles.nMatchCase = Button_GetCheck(hwndCtl)
            Exit Function
         End If

      Case IDC_FRMFINDINFILES_CHKWHOLEWORDS
         If codeNotify = BN_CLICKED Then
            gFindInFiles.nWholeWord = Button_GetCheck(hwndCtl)
            Exit Function
         End If

      Case IDC_FRMFINDINFILES_CHKSUBFOLDERS
         If codeNotify = BN_CLICKED Then
            gFindInFiles.nSearchSubFolders = Button_GetCheck(hwndCtl)
            Exit Function
         End If

      Case IDC_FRMFINDINFILES_CHKCURRENTDOC
         If codeNotify = BN_CLICKED Then
            gFindInFiles.nSearchCurrentDoc = Button_GetCheck(hwndCtl)
            gFindInFiles.nSearchAllOpenDocs = BST_UNCHECKED
            gFindInFiles.nSearchProject = BST_UNCHECKED
            frmFindInFiles_SetControlStates
            Exit Function
         End If

      Case IDC_FRMFINDINFILES_CHKALLOPENDOCS
         If codeNotify = BN_CLICKED Then
            gFindInFiles.nSearchAllOpenDocs = Button_GetCheck(hwndCtl)
            gFindInFiles.nSearchCurrentDoc = BST_UNCHECKED
            gFindInFiles.nSearchProject = BST_UNCHECKED
            frmFindInFiles_SetControlStates
            Exit Function
         End If

      Case IDC_FRMFINDINFILES_CHKPROJECT
         If codeNotify = BN_CLICKED Then
            gFindInFiles.nSearchProject = Button_GetCheck(hwndCtl)
            gFindInFiles.nSearchAllOpenDocs = BST_UNCHECKED
            gFindInFiles.nSearchCurrentDoc = BST_UNCHECKED
            frmFindInFiles_SetControlStates
            Exit Function
         End If

   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmFindInFiles
' ========================================================================================
Function frmFindInFiles_OnClose( byval HWnd As HWnd ) As LRESULT
   ' Enables parent window keeping parent's zorder
   EnableAllModeless()
   DestroyWindow HWnd
   Function = 0
End Function


' ========================================================================================
' Process WM_CREATE message for window/dialog: frmFindInFiles
' ========================================================================================
Function frmFindInFiles_OnCreate( _
            ByVal HWnd As HWnd, _
            ByVal lpCreateStructPtr As LPCREATESTRUCT _
            ) As BOOLEAN

   ' This is a modal popup window so disable the parent window
   DisableAllModeless()

   '  Message cracker macro expects a True to be returned for a successful
   '  OnCreate handler even though returning -1 from a standard WM_CREATE
   '  call would stop creating the window. This is just one of those Windows
   '  inconsistencies.
   Return True
End Function

' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmFindInFiles
' ========================================================================================
Function frmFindInFiles_OnDestroy( byval HWnd As HWnd ) As LRESULT 
   ' Save the search parameters for later instances of the find/replace dialog
   gFindInFiles.nWholeWord = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMFINDINFILES_CHKWHOLEWORDS) )  
   gFindInFiles.nMatchCase = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMFINDINFILES_CHKMATCHCASE) )    
   gFindInFiles.nSearchSubFolders = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMFINDINFILES_CHKSUBFOLDERS) )    
   gFindInFiles.nSearchCurrentDoc = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMFINDINFILES_CHKCURRENTDOC) )
   gFindInFiles.nSearchAllOpenDocs = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMFINDINFILES_CHKALLOPENDOCS) )
   gFindInFiles.nSearchProject = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMFINDINFILES_CHKPROJECT) )
   gFindInFiles.txtFind = AfxGetWindowText( GetDlgItem(HWnd_FRMFINDINFILES, IDC_FRMFINDINFILES_COMBOFIND))
   gFindInFiles.txtLastFind = gFindInFiles.txtFind 
   gFindInFiles.txtFiles = AfxGetWindowText( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFINDINFILES_COMBOFILES))
   gFindInFiles.txtFolder = AfxGetWindowText( GetDlgItem(HWND_FRMFINDINFILES,IDC_FRMFINDINFILES_COMBOFOLDER))

   HWND_FRMFINDINFILES = 0

   PostQuitMessage(0)

   Function = 0
End Function


' ========================================================================================
' frmFindInFiles Window procedure
' ========================================================================================
Function frmFindInFiles_WndProc( _
            ByVal HWnd   As HWnd, _
            ByVal uMsg   As UINT, _
            ByVal wParam As WPARAM, _
            ByVal lParam As LPARAM _
            ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CREATE,   frmFindInFiles_OnCreate)
      HANDLE_MSG (HWnd, WM_CLOSE,    frmFindInFiles_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,  frmFindInFiles_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,  frmFindInFiles_OnCommand)
   End Select

   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmFindInFiles_Show
' ========================================================================================
Function frmFindInFiles_Show( ByVal hWndParent As HWnd ) As LRESULT
   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   HWND_FRMFINDINFILES = _
   pWindow->Create(hwndParent, L(256,"Find In Files"), @frmFindInFiles_WndProc, 0, 0, 0, 0, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->SetClientSize(515, 222)
   pWindow->Center
   
   pWindow->AddControl("LABEL", , IDC_FRMFINDINFILES_LBLFINDWHAT, L(159,"Find What") & ":", 7, 19, 70, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_RIGHT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", , IDC_FRMFINDINFILES_LABEL1, L(259,"In Files") & ":", 7, 48, 70, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_RIGHT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", , IDC_FRMFINDINFILES_LABEL1, L(260,"In Folder") & ":", 7, 78, 70, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_RIGHT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("COMBOBOX", , IDC_FRMFINDINFILES_COMBOFIND, "", 85, 17, 317, 22, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWN Or CBS_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("COMBOBOX", , IDC_FRMFINDINFILES_COMBOFILES, "", 85, 46, 317, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWN Or CBS_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("COMBOBOX", , IDC_FRMFINDINFILES_COMBOFOLDER, "", 85, 75, 317, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWN Or CBS_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("BUTTON", , IDOK, L(158,"Find"), 414, 16, 90, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_DEFPUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMFINDINFILES_CMDFILES, "...", 414, 45, 24, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMFINDINFILES_CMDFOLDERS, "...", 414, 74, 24, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDCANCEL, L(161,"Close"), 414, 105, 90, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("GROUPBOX", , IDC_FRMFINDINFILES_FRAMEOPTIONS, L(167,"Options"), 16, 105, 386, 107, _
        WS_CHILD Or WS_VISIBLE Or BS_TEXT Or BS_LEFT Or BS_NOTIFY Or BS_GROUPBOX, _
        WS_EX_TRANSPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMFINDINFILES_CHKWHOLEWORDS, L(162,"Match Whole Words"), 26, 123, 167, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMFINDINFILES_CHKMATCHCASE, L(163,"Match Case"), 26, 143, 167, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMFINDINFILES_CHKSUBFOLDERS, L(258,"Search Subfolders"), 26, 163, 167, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   pWindow->AddControl("CHECKBOX", , IDC_FRMFINDINFILES_CHKCURRENTDOC, L(408,"Current Document"), 195, 123, 167, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMFINDINFILES_CHKALLOPENDOCS, L(409,"All Open Documents"), 195, 143, 167, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMFINDINFILES_CHKPROJECT, L(410,"Current Project"), 195, 163, 167, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   Dim As Long i, nCount, startPos, endPos, startLine, endLine
   Dim As String buffer, sFind, sFiles, sFolder
   
   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   If pDoc Then
      ' Fills the search box with the selected word.
      ' If there are carriage returns or/and line feeds, this mean that
      ' there is a block selected, instead of a word, so avoid it.
      buffer = pDoc->GetSelText()
      If Len(buffer) Then
         if pDoc->IsMultilineSelection = false then
            sFind = buffer
            gFindInFiles.txtLastFind = buffer
         end if
      else
         sFind = gFindInFiles.txtLastFind
      End If
   End If
   frmFindInFiles_AddToFindCombo(sFind)

   if len(gFindInFiles.txtFilesCombo(0)) = 0 THEN sFiles = "*.bas *.bi *.inc"
   frmFindInFiles_AddToFilesCombo(sFiles)
   
   if len(gFindInFiles.txtFolderCombo(0)) = 0 THEN 
      if gApp.IsProjectActive THEN
         sFolder = AfxStrPathName("PATH", gApp.ProjectFilename)
      else
         sFolder = AfxGetExePathName 
      END IF
   end if
   frmFindInFiles_AddToFolderCombo(sFolder)

   AfxSetWindowText( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFINDINFILES_COMBOFIND), gFindInFiles.txtFind )    
   AfxSetWindowText( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFINDINFILES_COMBOFILES), gFindInFiles.txtFiles )    
   AfxSetWindowText( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFINDINFILES_COMBOFOLDER), gFindInFiles.txtFolder )    
   EnableWindow( GetDlgItem(HWND_FRMFINDINFILES, IDOK), True )
 
   frmFindInFiles_SetControlStates

   ShowWindow HWND_FRMFINDINFILES, SW_SHOW
   SetFocus( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFINDINFILES_COMBOFIND) )

   ' Process Windows messages(modal)
   Function = pWindow->DoEvents( SW_SHOW )
   
   ' Delete the CWindow class manually allocated memory 
   Delete pWindow

   Function = 0
End Function

