'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2017 Paul Squires, PlanetSquires Software
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


' ========================================================================================
' Set the mouse pointer for a specific window
' ========================================================================================
function SetMousePtr( byval HWND as hWnd, byval hDesiredCursor as HCURSOR ) as HCURSOR
 SetClassLongPtr(hWnd, GCLP_HCURSOR, cast(long_PTR, hDesiredCursor))
 return SetCursor(hDesiredCursor)
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
   if AfxFileExists(gApp.ProjectFilename) Then
      zText = gApp.ProjectFilename
   else
      zText = exepath
   end if      
   Dim wzCurDrive As CWSTR = LCase(Left(**zText, 3))  ' eg. D:\

   If LCase(Left(**wzFilename, 3)) = wzCurDrive Then 
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
   if AfxFileExists(gApp.ProjectFilename) Then
      zText = gApp.ProjectFilename
   else
      zText = exepath
   end if      
   If Ucase(Left(**wzFilename, 10)) = WSTR("{CURDRIVE}") Then 
      wzFilename = Left(**zText, 1) & Mid(wzFilename, 11)
   End If

   Return wzFilename
End Function


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
   uInsert.item.mask               = TVIF_TEXT Or TVIF_IMAGE Or TVIF_SELECTEDIMAGE Or TVIF_PARAM
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
   Dim rgFileTypes(1 To 3) As COMDLG_FILTERSPEC

   Select Case idButton
      Case IDM_PROJECTOPEN
         rgFileTypes(1).pszName = @L(216,"Project files")
         rgFileTypes(1).pszSpec = @WSTR("*.wfbe")
         rgFileTypes(2).pszName = @L(79,"All files")
         rgFileTypes(2).pszSpec = @WSTR("*.*")
         pofd->lpVtbl->SetFileTypes(pofd, 2, @rgFileTypes(1))
         ' Set the title of the dialog
         hr = pofd->lpVtbl->SetTitle(pofd, L(216,"Project files"))
      
      Case IDM_INSERTFILE
         rgFileTypes(1).pszName = @L(77,"FB code files")
         rgFileTypes(2).pszName = @L(79,"All files")
         rgFileTypes(1).pszSpec = @WSTR("*.bas;*.inc;*.bi;*.rc")
         rgFileTypes(2).pszSpec = @WSTR("*.*")
         pofd->lpVtbl->SetFileTypes(pofd, 2, @rgFileTypes(1))
         ' Set the title of the dialog
         hr = pofd->lpVtbl->SetTitle(pofd, L(80,"Insert File"))

      Case IDC_FRMOPTIONSCOMPILER_CMDFBWIN32, IDC_FRMOPTIONSCOMPILER_CMDFBWIN64
         rgFileTypes(1).pszName = @L(100,"FB Compiler")
         rgFileTypes(1).pszSpec = @WSTR("fbc.exe")
         rgFileTypes(2).pszName = @L(79,"All files")
         rgFileTypes(2).pszSpec = @WSTR("*.*")
         pofd->lpVtbl->SetFileTypes(pofd, 2, @rgFileTypes(1))
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
   Dim rgFileTypes(1 To 3) As COMDLG_FILTERSPEC
   rgFileTypes(1).pszName = @L(77,"FB code files")
   rgFileTypes(2).pszName = @L(79,"All files")
   rgFileTypes(1).pszSpec = @WSTR("*.bas;*.inc;*.bi;*.rc")
   rgFileTypes(2).pszSpec = @WSTR("*.*")
   pofd->lpVtbl->SetFileTypes(pofd, 2, @rgFileTypes(1))

   ' Set the title of the dialog
   hr = pofd->lpVtbl->SetTitle(pofd, L(106,"Open source file"))
   ' Set the default folder
   Dim pFolder As IShellItem Ptr
   SHCreateItemFromParsingName (Exepath, Null, @IID_IShellItem, @pFolder)
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
   Dim rgFileTypes(1 To 3) As COMDLG_FILTERSPEC 
   
   Select Case id
      Case IDC_FRMPROJECTOPTIONS_CMDSELECT, IDM_PROJECTSAVE
         rgFileTypes(1).pszName = @L(216,"Project files")
         rgFileTypes(2).pszName = @L(79,"All files")
         rgFileTypes(1).pszSpec = @WSTR("*.wfbe")
         rgFileTypes(2).pszSpec = @WSTR("*.*")
         psfd->lpVtbl->SetFileTypes(psfd, 2, @rgFileTypes(1))
      Case Else
         rgFileTypes(1).pszName = @L(77,"FB code files")
         rgFileTypes(2).pszName = @L(79,"All files")
         rgFileTypes(1).pszSpec = @WSTR("*.bas;*.inc;*.bi;*.rc")
         rgFileTypes(2).pszSpec = @WSTR("*.*")
         psfd->lpVtbl->SetFileTypes(psfd, 2, @rgFileTypes(1))
   End Select
   
   ' // Set the title of the dialog
   hr = psfd->lpVtbl->SetTitle(psfd, L(107,"Save source file"))
   ' // Set the file name
   hr = psfd->lpVtbl->SetFileName(psfd, pwszFileName)
   ' // Set the extension
   hr = psfd->lpVtbl->SetDefaultExtension(psfd, pwszDefExt)
   ' // Display the dialog
   hr = psfd->lpVtbl->Show(psfd, hwndOwner)

   ' // Set the default folder
   Dim pFolder As IShellItem Ptr
   SHCreateItemFromParsingName (Exepath, Null, @IID_IShellItem, @pFolder)
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
   If SendMessageW(hToolBar, TB_ISBUTTONENABLED, Cast(WPARAM, idButton), 0 ) = 0 Then
      Function = SendMessageW(hToolBar, TB_ENABLEBUTTON, Cast(WPARAM, idButton), Cast(LPARAM, MAKELONG(CTRUE, 0)))
   End If   
End Function
' ========================================================================================

' ========================================================================================
' Disables the specified button in a toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
Function FF_Toolbar_DisableButton (ByVal hToolBar As HWnd, ByVal idButton As Long) As BOOLEAN
   ' Only disable the button if it is not already disabled. This should help reduce flicker.
   If SendMessageW(hToolBar, TB_ISBUTTONENABLED, Cast(WPARAM, idButton), 0 ) Then
      Function = SendMessageW(hToolBar, TB_ENABLEBUTTON, Cast(WPARAM, idButton), Cast(LPARAM, MAKELONG(False, 0)))
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
      Function = SendMessageW( hWndControl, LVM_INSERTITEM, 0, Cast(LPARAM, @lvi) )
   Else 
      lvi.mask = LVIF_TEXT Or LVIF_IMAGE
      Function = SendMessageW( hWndControl, LVM_SETITEM, 0, Cast(LPARAM, @lvi) )
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
       
   Function = SendMessageW( hWndControl, LVM_GETITEM, 0, Cast(LPARAM, @lvi) )
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
Function LoadLocalizationFile( ByVal pwszName As WString Ptr ) As BOOLEAN

   ' default that the file failed to load
   Function = False
   If pwszName = 0 Then Exit Function
   If Dir(*pwszName) = "" Then Exit Function

   Dim st    As WString * MAX_PATH
   Dim wKey  As WString * MAX_PATH
   Dim wData As WString * MAX_PATH
   Dim nKey  As Long
   Dim nData As Long  
   Dim i     As Long
    
   Dim f As Long = Freefile
   If Open( *pwszName For Input Encoding "utf16" As #f ) <> 0 Then Exit Function

   Do Until Eof(f)
      Line Input #f, st
      If Len(st) = 0 Then Continue Do
      If Left(st, 1) = "'" Then Continue Do
      
      i = Instr(st, ":")
      If i = 0 Then Continue Do
      
      wKey  = Left(st, i-1)
      wData = Mid(st, i+1)
      
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
   Close #f

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






