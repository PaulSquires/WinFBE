
''
''
Function GetMRUMenuHandle() as HMENU
   Dim hFileMenu as HMENU = GetSubMenu( HWND_FRMMAIN_TOPMENU, 0 ) 
   Dim hSubMenu  As HMENU = GetSubMenu( hFileMenu, 10 )
   Function = hSubMenu
End Function


''
''
Function OpenMRUFile( ByVal HWnd As HWnd, ByVal wID As Long ) As Long
   Dim hSubMenu as HMENU = GetMRUMenuHandle()
   Dim wzFile As WString * MAX_PATH
   Dim idx As Long
   
   GetMenuStringW(hSubMenu, wID, wzFile, MAX_PATH, MF_BYCOMMAND)

   idx = gTTabCtl.GetTabIndexFromFilename(@wzFile)
   If idx >= 0 Then
      gTTabCtl.SetFocusTab(idx)
   Else
      ' Create the new Scintilla window and load the file
      Dim pDoc As clsDocument Ptr = New clsDocument
      pDoc->CreateCodeWindow(HWnd, False, False, wzFile)
      gpApp->AddDocument(pDoc)        ' Add the new document to the global app            
      idx = gTTabCtl.AddTab(pDoc)   ' Add the new document to the top tabcontrol
      gTTabCtl.SetFocusTab(idx)
   End If
   
   ' Position all of the controls into place
   frmMain_PositionWindows(HWnd)

   Function = 0
End Function

''
''
Function UpdateMRUMenu( ByVal hMenu As HMENU ) As Long
   ' Modify the actual topmenu display items based on the MRU strings
   Dim hSubMenu As HMENU = GetMRUMenuHandle()
   Dim i        As Long 
   
   If hMenu <> hSubMenu Then Exit Function
   
   ' Remove the existing items
   Do Until GetMenuItemCount(hSubMenu) = 0
      RemoveMenu( hSubMenu, 0, MF_BYPOSITION )
   Loop
   
   ' Add the new items
   For i = 0 To 9
      ' If the file no longer exists then remove it from the MRU list
      ProcessFromCurdrive(@gConfig.MRU(i))
      If Dir(gConfig.MRU(i)) = "" Then
         gConfig.MRU(i) = ""
      Else
         AppendMenuW( hSubMenu, MF_STRING, IDC_MRUBASE + i, @gConfig.MRU(i) )
      End If
   Next
   
   If GetMenuItemCount(hSubMenu) = 0 Then
      AppendMenuW( hSubMenu, MF_STRING Or MF_GRAYED Or MF_DISABLED, 0, L(11,"(Empty)") )
   End If
   
   Function = 0
End Function
   
''
''
Function UpdateMRUList( ByVal wzFilename As WString Ptr ) As Long

   Dim idxExistsAt as Long = -1   ' index of where filename already exists 
   Dim i           as Long
   Dim wzFile      As WString * MAX_PATH
   
   ' Search the 10 MRU items to see if the filename already exists and
   ' also for an available entry to put the filename in.
   For i = 0 To 9
       wzFile = gConfig.MRU(i)
       If Ucase(wzFile) = Ucase(*wzFilename) Then
          idxExistsAt = i:  Exit For
       End If   
   Next
   
   ' If does not already exist then we simply need to move all entries
   ' down and add at the beginning of the list. The original 10th entry
   ' will now drop off the list.
   If idxExistsAt = -1 Then idxExistsAt = 9
   
   For i = idxExistsAt To 1 Step -1
      gConfig.MRU(i) = gConfig.MRU(i-1)
   Next
   gConfig.MRU(0) = *wzFilename 
   
   ' Save the configuration file
   gConfig.SaveToFile

   Function = 0
End Function  

