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
' Add a string to the Find combobox (ensure no duplicates)
' ========================================================================================
Function frmFindInFiles_AddToFindCombo( ByRef sText As Const String ) As Long
   Dim as long i
   dim as long nFound = -1
   
   if len(sText) = 0 THEN exit function
   ' See if the string already exists in the combobox
   for i = 0 to ubound(gFindInFiles.txtFindCombo)
      if gFindInFiles.txtFindCombo(i) = sText THEN
         nFound = i: exit for
      END IF
   NEXT
   if nFound = -1 THEN  ' was not found
      for i = ubound(gFindInFiles.txtFindCombo) to 1 step -1
         gFindInFiles.txtFindCombo(i) = gFindInFiles.txtFindCombo(i-1)
      next   
      gFindInFiles.txtFindCombo(0) = sText
   END IF
   gFindInFiles.txtFind = sText
   
   Function = 0
End Function

' ========================================================================================
' Add a string to the Files combobox (ensure no duplicates)
' ========================================================================================
Function frmFindInFiles_AddToFilesCombo( ByRef sText As Const String ) As Long
   Dim as long i
   dim as long nFound = -1
   
   if len(sText) = 0 THEN exit function
   ' See if the string already exists in the combobox
   for i = 0 to ubound(gFindInFiles.txtFilesCombo)
      if gFindInFiles.txtFilesCombo(i) = sText THEN
         nFound = i: exit for
      END IF
   NEXT
   if nFound = -1 THEN  ' was not found
      for i = ubound(gFindInFiles.txtFilesCombo) to 1 step -1
         gFindInFiles.txtFilesCombo(i) = gFindInFiles.txtFilesCombo(i-1)
      next   
      gFindInFiles.txtFilesCombo(0) = sText
   END IF
   gFindInFiles.txtFiles = sText

   Function = 0
End Function

' ========================================================================================
' Add a string to the Folders combobox (ensure no duplicates)
' ========================================================================================
Function frmFindInFiles_AddToFolderCombo( ByRef sText As Const String ) As Long
   Dim as long i
   dim as long nFound = -1
   
   if len(sText) = 0 THEN exit function
   ' See if the string already exists in the combobox
   for i = 0 to ubound(gFindInFiles.txtFolderCombo)
      if gFindInFiles.txtFolderCombo(i) = sText THEN
         nFound = i: exit for
      END IF
   NEXT
   if nFound = -1 THEN  ' was not found
      for i = ubound(gFindInFiles.txtFolderCombo) to 1 step -1
         gFindInFiles.txtFolderCombo(i) = gFindInFiles.txtFolderCombo(i-1)
      next   
      gFindInFiles.txtFolderCombo(0) = sText
   END IF
   gFindInFiles.txtFolder = sText

   Function = 0
End Function

' ========================================================================================
' File search procedure (shell out to builtin Windows "findstr" utility
' ========================================================================================
function ParseSearchResults( byref wszResultFile as wstring ) as long 
   
   dim hLB as hwnd = GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_LISTSEARCH)
   ListBox_ResetContent( hLB )

   wszResultFile = AfxStrRemove(wszResultFile, chr(34))   ' remove quotes
   if AfxFileExists(wszResultFile) = 0 THEN exit function
   
   dim as CWSTR cwzStr
   dim as string st
   dim as long f1, f2

   dim pStream as CTextStream
   if pStream.Open(wszResultFile) <> S_OK then exit function
   
   cwzStr = L(262, "Search Results") & ": " & gFindInFiles.txtFind
   ListBox_AddString(hLB, cwzStr)
   cwzStr = ""
   ListBox_AddString(hLB, cwzStr)
   
   do until pStream.EOS
      st = pStream.ReadLine
      
      st = trim(st)
      if len(st) = 0 THEN continue do

      ' Format the string to display in the ListBox
      
      ' Original as seen in the output file
      'X:\FB\WinFBE - Editor\license.txt:1:WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
      ' Format that we want to transfor it into
      'X:\FB\WinFBE - Editor\license.txt(1): WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
      
      ' Search for the 2nd semicolon
      f1 = instr(3, st, ":")
      if f1 THEN
         mid(st, f1, 1) = "("
         f2 = instr(f1, st, ":")
         if f2 THEN
            st = left(st, f2-1) & "): " & mid(st, f2+1)
         end if
      END IF

      cwzStr = st   
      ListBox_AddString(hLB, cwzStr)
   loop
   pStream.Close
   
   kill wszResultFile 
   
   ' Show the ListBox search results
   TabCtrl_SetCurSel( GetDlgItem( HWND_FRMOUTPUT, IDC_FRMOUTPUT_TABCONTROL), 2 )
   ShowWindow HWND_FRMOUTPUT, SW_SHOW
   PositionOutputWindows(HWND_FRMOUTPUT)
   frmMain_PositionWindows(HWND_FRMMAIN)
      
   function = 0
END FUNCTION


' ========================================================================================
' File search procedure (shell out to builtin Windows "findstr" utility
' ========================================================================================
function DoFindInFilesEx() as LONG
   Dim ShExecInfo As SHELLEXECUTEINFOW  

'  /L         Uses search strings literally.
'  /P         Skip files with non-printable characters.
'  /N         Prints the line number before each line that matches.
   'Dim sParam as string = " /L /P /N"
   Dim sParam as string = " /P /N"
   
'  /R         Whole word searching needs regular expression
   If gFindInFiles.nWholeWord Then sParam = sParam & " /R "
'  /I         Specifies that the search is not to be case-sensitive.
   If gFindInFiles.nMatchCase = 0 Then sParam = sParam & " /I "     ' case insensitive search
   
'  /S         Searches for matching files in the current directory and all
'             subdirectories.
   ' If /S is used then resulting filenames will include full path, otherwise
   ' only the filename is returned.
   If gFindInFiles.nSearchSubFolders Then sParam = sParam & " /S "

   ' Ensure that the text in the find textbox has been added to the combobox. Need to do
   ' this to ensure that manually typed in text has been captured.
   Dim wszText As WString * MAX_PATH
   wszText = AfxGetWindowText( GetDlgItem(HWnd_FRMFINDINFILES, IDC_FRMFINDINFILES_COMBOFIND))
   if len(wszText) = 0 THEN
      SetFocus(GetDlgItem(HWnd_FRMFINDINFILES, IDC_FRMFINDINFILES_COMBOFIND))
      exit function      
   END IF
   gFindInFiles.txtFind = Str(wszText)
   frmFindInFiles_AddToFindCombo(gFindInFiles.txtFind)
   If gFindInFiles.nWholeWord Then 
      wszText = "\<" & wszText & "\>"
   end if   
   
   ' What folder to start searching in
   gFindInFiles.txtFolder = str(trim(AfxGetWindowText( GetDlgItem(HWND_FRMFINDINFILES,IDC_FRMFINDINFILES_COMBOFOLDER) )))
   if len(gFindInFiles.txtFolder) = 0 THEN 
      gFindInFiles.txtFolder = AfxGetExePathName
      AfxSetWindowText(GetDlgItem(HWND_FRMFINDINFILES,IDC_FRMFINDINFILES_COMBOFOLDER), gFindInFiles.txtFolder)
   end if
   frmFindInFiles_AddToFolderCombo(gFindInFiles.txtFolder)
   if right(gFindInFiles.txtFolder,1) <> "\" then gFindInFiles.txtFolder = gFindInFiles.txtFolder & "\"
   dim sFolder as CWSTR = gFindInFiles.txtFolder 

   ' What files are we searching in
   dim as long nCount, i
   dim as BOOLEAN InString
   dim as CWSTR sFiles, sSearchFiles, sFilename, sChar
   sFiles = trim(AfxGetWindowText( GetDlgItem(HWnd_FRMFINDINFILES, IDC_FRMFINDINFILES_COMBOFILES)))
   gFindInFiles.txtFiles = str(sFiles)
   frmFindInFiles_AddToFilesCombo(gFindInFiles.txtFiles)
   ' Parse the files to determine if multiple files or file types need to be searched for
   if instr(sFiles, wchr(34)) THEN
      ' In order to better parse the string for the filenames, replace any spaces in the
      ' filenames with special character; then parse; and finally remove the special character.
      for i = 1 to len(sFiles)
         sChar = mid(sFiles, i, 1)
         if sChar = Wchr(34) THEN
            InString = iif( InString = true, false, true) 
         elseif sChar = wstr(" ") then
            if InString = true THEN
               ' special character replaces the space
               sFiles = left(**sFiles, i-1) & wchr(8) & mid(sFiles, i+1)
            END IF         
         END IF
      NEXT
      
      ' Clean the string to remove any duplicate remaining spaces
      sFiles = AfxStrShrink(sFiles)

      ' This is a list of double quoted filenames to search
      nCount = AfxStrParseCount(sFiles, " ")
      for i = 1 to nCount
         sFilename = trim( AfxStrParse(sFiles, i, " "), Wchr(34) )
         ' Put back any special characters to spaces again
         sFilename = AfxStrReplace(sFilename, wchr(8), wchr(32) )
         sSearchFiles = sSearchFiles + Wchr(34) + sFolder + sFilename + Wchr(34, 32)
      NEXT   
   else
      ' This is a list of patterns to use (*.bas *.bi) etc
      nCount = AfxStrParseCount(sFiles, " ")
      for i = 1 to nCount
         sSearchFiles = sSearchFiles + wchr(34) + sFolder + AfxStrParse(sFiles, i, " ") + wchr(34, 32)
      NEXT   
   END IF
      
   Dim wszFileExe   As WString * MAX_PATH
   Dim wszParam     As WString * MAX_PATH  
   Dim wszBatchFile As WString * MAX_PATH  
   Dim wszResults   As WString * MAX_PATH  

   'FINDSTR [/B] [/E] [/L] [/R] [/S] [/I] [/X] [/V] [/N] [/M] [/O] [/P] [/F:file]
   '        [/C:string] [/G:file] [/D:dir list] [/A:color attributes] [/OFF[LINE]]
   '        strings [[drive:][path]filename[ ...]]
   wszResults = AfxGetExePathName & "\_search.txt"
   if AfxFileExists(wszResults) then kill wszResults
   if instr(wszResults, " ") THEN wszResults = chr(34) & wszResults & chr(34)
   wszFileExe = "FindStr.exe"  
   wszParam   = WSTR(sParam & " /C:") & chr(34) & wszText & chr(34) & _
                 " " & **sSearchFiles & " >> " & wszResults

   wszBatchFile = AfxGetExePathName & "\_search.bat"
   ' Open filename does not use enclosing quotes
   dim pStream as CTextStream
   if pStream.Create(wszBatchFile) = S_OK then
      pStream.WriteLine wszFileExe & " " & wszParam
   end if
   pStream.Close   
 
   wszFileExe = Environ("COMSPEC")
   ' Batch filename must be enclosed in quotes b/c might have spaces in name
   wszParam   = wstr(" /C " & chr(34) & wszBatchFile & chr(34))
   With ShExecInfo
      .cbSize       = Len(SHELLEXECUTEINFOW)
      .fMask        = SEE_MASK_NOCLOSEPROCESS
      .HWnd         = 0
      .lpVerb       = Null
      .lpFile       = @wszFileExe
      .lpParameters = @wszParam   
      .lpDirectory  = 0
      .nShow        = SW_HIDE
      .hInstApp     = 0 
   End With

   ShellExecuteEx(@ShExecInfo)

   ' Give the process 33 ms (you can use a different value here depending on 
   ' how responsive you wish your app to be)
   Dim uMsg As MSG
   while WaitForSingleObject(ShExecInfo.hProcess, 33) = WAIT_TIMEOUT
   '  if after 33 ms the object's handle is not signaled..       
   '  we examine the message queue and if there are any waiting..
   '  Note:  see PeekMessage documentation for details on how to limit
   '         the types of messages to look for
      while(PeekMessage(@uMsg, HWND_FRMMAIN,  0, 0, PM_NOREMOVE))
        ' we process them..
        if(GetMessage(@uMsg, NULL, 0, 0) > 0) then
            TranslateMessage(@uMsg)
            DispatchMessage(@uMsg)
         end if
      wend
   wend

   kill wszBatchFile
   
   ParseSearchResults(wszResults)

'C:\>findstr /?
'Searches for strings in files.
'
'FINDSTR [/B] [/E] [/L] [/R] [/S] [/I] [/X] [/V] [/N] [/M] [/O] [/P] [/F:file]
'        [/C:string] [/G:file] [/D:dir list] [/A:color attributes] [/OFF[LINE]]
'        strings [[drive:][path]filename[ ...]]
'
'  /B         Matches pattern if at the beginning of a line.
'  /E         Matches pattern if at the end of a line.
'  /L         Uses search strings literally.
'  /R         Uses search strings as regular expressions.
'  /S         Searches for matching files in the current directory and all
'             subdirectories.
'  /I         Specifies that the search is not to be case-sensitive.
'  /X         Prints lines that match exactly.
'  /V         Prints only lines that do not contain a match.
'  /N         Prints the line number before each line that matches.
'  /M         Prints only the filename if a file contains a match.
'  /O         Prints character offset before each matching line.
'  /P         Skip files with non-printable characters.
'  /OFF[LINE] Do not skip files with offline attribute set.
'  /A:attr    Specifies color attribute with two hex digits. See "color /?"
'  /F:file    Reads file list from the specified file(/ stands for console).
'  /C:string  Uses specified string as a literal search string.
'  /G:file    Gets search strings from the specified file(/ stands for console).
'  /D:dir     Search a semicolon delimited list of directories
'  strings    Text to be searched for.
'  [drive:][path]filename
'             Specifies a file or files to search.
'
' Use spaces to separate multiple search strings unless the argument is prefixed
' with /C.  For example, 'FINDSTR "hello there" x.y' searches for "hello" or
' "there" in file x.y.  'FINDSTR /C:"hello there" x.y' searches for
' "hello there" in file x.y.
'
' Regular expression quick reference:
'  .        Wildcard: any character
'  *        Repeat: zero or more occurrences of previous character or class
'  ^        Line position: beginning of line
'  $        Line position: end of line
'  [class]  Character class: any one character in set
'  [^class] Inverse class: any one character not in set
'  [x-y]    Range: any characters within the specified range
'  \x       Escape: literal use of metacharacter x
'  \<xyz    Word position: beginning of word
'  xyz\>    Word position: end of word

   function = 0
end function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmFindInFiles
' ========================================================================================
Function frmFindInFiles_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   Select Case id
      Case IDOK  ' Find Next
         If codeNotify = BN_CLICKED Then
            EnableWindow(hwndCtl, False)
            DoFindInFilesEx()
            EnableWindow(hwndCtl, true)
            Exit Function
         End If  
         
      Case IDCANCEL  ' Close
         If codeNotify = BN_CLICKED Then
            SendMessageW HWnd, WM_CLOSE, 0, 0
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
                  wszText = wstr(gFindInFiles.txtFindCombo(i))
                  ComboBox_AddString( hwndCtl, @wszText )
               end if   
            Next
            wszText = wstr(gFindInFiles.txtFind)
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
                  wszText = wstr(gFindInFiles.txtFilesCombo(i))
                  ComboBox_AddString( hwndCtl, @wszText )
               end if   
            Next
            wszText = wstr(gFindInFiles.txtFiles)
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
                  wszText = wstr(gFindInFiles.txtFolderCombo(i))
                  ComboBox_AddString( hwndCtl, @wszText )
               end if   
            Next
            wszText = wstr(gFindInFiles.txtFolder)
            i = ComboBox_FindStringExact( hwndCtl, -1, @wszText )
            ComboBox_SetCurSel(hwndCtl, i)
         End If

      Case IDC_FRMFINDINFILES_CMDFILES
         If codeNotify = BN_CLICKED Then
            ' Display the Open File Dialog
            Dim pItems As IShellItemArray Ptr = AfxIFileOpenDialogMultiple(HWnd)
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

         
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmFindInFiles
' ========================================================================================
Function frmFindInFiles_OnClose(HWnd As HWnd) As LRESULT
   DestroyWindow HWnd
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmFindInFiles
' ========================================================================================
Function frmFindInFiles_OnDestroy(HWnd As HWnd) As LRESULT 
   ' Save the search parameters for later instances of the find/replace dialog
   gFindInFiles.nWholeWord = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMFINDINFILES_CHKWHOLEWORDS) )  
   gFindInFiles.nMatchCase = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMFINDINFILES_CHKMATCHCASE) )    
   gFindInFiles.nSearchSubFolders = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMFINDINFILES_CHKSUBFOLDERS) )    
   gFindInFiles.txtFind = str(AfxGetWindowText( GetDlgItem(HWnd_FRMFINDINFILES, IDC_FRMFINDINFILES_COMBOFIND)))
   gFindInFiles.txtLastFind = gFindInFiles.txtFind 
   gFindInFiles.txtFiles = str(AfxGetWindowText( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFINDINFILES_COMBOFILES)))
   gFindInFiles.txtFolder = str(AfxGetWindowText( GetDlgItem(HWND_FRMFINDINFILES,IDC_FRMFINDINFILES_COMBOFOLDER)))

   HWND_FRMFINDINFILES = 0
   Function = 0
End Function


' ========================================================================================
' frmFindInFiles Window procedure
' ========================================================================================
Function frmFindInFiles_WndProc( ByVal HWnd   As HWnd, _
                                 ByVal uMsg   As UINT, _
                                 ByVal wParam As WPARAM, _
                                 ByVal lParam As LPARAM _
                                 ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CLOSE,    frmFindInFiles_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,  frmFindInFiles_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,  frmFindInFiles_OnCommand)
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmFindInFiles_Show
' ========================================================================================
Function frmFindInFiles_Show( ByVal hWndParent As HWnd ) As Long

   ' If the FindInFiles window already exists then simply set focus to it and exit.
   If IsWindow(HWND_FRMFINDINFILES) Then
      SetForegroundWindow(HWND_FRMFINDINFILES)
      SetFocus( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFINDINFILES_COMBOFIND) )
      Exit Function
   End If

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   HWND_FRMFINDINFILES = _
   pWindow->Create(hwndParent, L(256,"Find In Files"), @frmFindInFiles_WndProc, 0, 0, 0, 0, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->SetClientSize(510, 205)
   pWindow->Center
   
   pWindow->AddControl("LABEL", , IDC_FRMFINDINFILES_LBLFINDWHAT, L(159,"Find What:"), 7, 20, 70, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_RIGHT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", , IDC_FRMREPLACE_LBLREPLACEWITH, L(259,"In Files:"), 7, 50, 70, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_RIGHT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", , IDC_FRMREPLACE_LBLREPLACEWITH, L(260,"In Folder:"), 7, 80, 70, 16, _
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
   pWindow->AddControl("GROUPBOX", , IDC_FRMFINDINFILES_FRAMEOPTIONS, L(167,"Options"), 16, 105, 187, 90, _
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

   if len(gFindInFiles.txtFilesCombo(0)) = 0 THEN sFiles = "*.bas *.bi *inc"
   frmFindInFiles_AddToFilesCombo(sFiles)
   
   if len(gFindInFiles.txtFolderCombo(0)) = 0 THEN 
      if gApp.IsProjectActive THEN
         sFolder = AfxStrPathName("PATH", gApp.Projects(gApp.GetActiveProjectIndex).ProjectFilename)
      else
         sFolder = AfxGetExePathName 
      END IF
   end if
   frmFindInFiles_AddToFolderCombo(sFolder)

   SetWindowText( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFINDINFILES_COMBOFIND), WStr(gFindInFiles.txtFind) )    
   SetWindowText( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFINDINFILES_COMBOFILES), WStr(gFindInFiles.txtFiles) )    
   SetWindowText( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFINDINFILES_COMBOFOLDER), WStr(gFindInFiles.txtFolder) )    
   EnableWindow( GetDlgItem(HWND_FRMFINDINFILES, IDOK), True )
 
   Button_SetCheck( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFINDINFILES_CHKWHOLEWORDS), gFindInFiles.nWholeWord )
   Button_SetCheck( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFINDINFILES_CHKMATCHCASE), gFindInFiles.nMatchCase )
   Button_SetCheck( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFINDINFILES_CHKSUBFOLDERS), gFindInFiles.nSearchSubFolders )
   
   ShowWindow HWND_FRMFINDINFILES, SW_SHOW
   SetFocus( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFIND_COMBOFIND) )

   Function = 0
End Function



