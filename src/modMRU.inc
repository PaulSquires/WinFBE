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

''
''
Function GetMRUMenuHandle() as HMENU
   Dim hFileMenu as HMENU = GetSubMenu( HWND_FRMMAIN_TOPMENU, 0 ) 
   Dim hSubMenu  As HMENU = GetSubMenu( hFileMenu, 11 )
   Function = hSubMenu
End Function


''
''
Function OpenMRUFile( ByVal HWnd As HWnd, ByVal wID As Long ) As Long
   Dim wzFile As WString * MAX_PATH
   
   wzFile = gConfig.MRU(wID - IDC_MRUBASE)
   frmMain_OpenFileSafely(HWND_FRMMAIN, _
                           False, _    ' bIsNewFile
                           False, _    ' bIsTemplate
                           true, _     ' bShowInTab
                           false, _    ' bIsInclude
                           @wzFile, _  ' pwszName
                           0 )  ' pDocIn
   
   Function = 0
End Function


''
''
Function ClearMRUlist( ByVal wID As Long ) As Long

   Dim As Long i 
   
   Select Case wID
      Case IDM_MRUCLEAR
         For i = 0 To 9
            gConfig.MRU(i) = ""
         Next
         LoadRecentFilesTreeview(HWND_FRMRECENT)
      Case IDM_MRUPROJECTCLEAR
         For i = 0 To 9
            gConfig.MRUPROJECT(i) = ""
         Next
         LoadRecentFilesTreeview(HWND_FRMRECENT)
   End Select
   
   PositionRecentWindows(HWND_FRMRECENT)

   Function = 0
End Function


''
''
Function CreateMRUpopup() As HMENU
   ' Popup menu that is displayed when downarrown on toolbar is pressed.
   Dim As HMENU hPopupMenu = CreatePopupMenu()
   
   ' Display MRU files
   For i As Long = 0 To 9
      ' If the file no longer exists then remove it from the MRU list
      gConfig.MRU(i) = ProcessFromCurdrive(gConfig.MRU(i))
      If AfxFileExists(gConfig.MRU(i)) = 0 Then
         gConfig.MRU(i) = ""
      Else
         AppendMenuW( hPopupMenu, MF_STRING, IDC_MRUBASE + i, gConfig.MRU(i) )
      End If
   Next
   
   ' Display MRU projects
   For i As Long = 0 To 9
      ' If the file no longer exists then remove it from the MRU list
      gConfig.MRUPROJECT(i) = ProcessFromCurdrive(gConfig.MRUPROJECT(i))
      If AfxFileExists(gConfig.MRUPROJECT(i)) = 0 Then
         gConfig.MRUPROJECT(i) = ""
      Else               
         If i = 0 Then AppendMenuW( hPopUpMenu, MF_SEPARATOR, 0, "" )
         AppendMenuW( hPopupMenu, MF_STRING, IDC_MRUPROJECTBASE + i, gConfig.MRUPROJECT(i) )
      End If
   Next

   If GetMenuItemCount(hPopupMenu) = 0 Then
      AppendMenuW( hPopupMenu, MF_STRING Or MF_GRAYED Or MF_DISABLED, 0, L(11,"(Empty)") )
   End If
   
   Function = hPopupMenu
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
      gConfig.MRU(i) = ProcessFromCurdrive(gConfig.MRU(i))
      If AfxFileExists(gConfig.MRU(i)) = 0 Then
         gConfig.MRU(i) = ""
      Else
         AppendMenu( hSubMenu, MF_STRING, IDC_MRUBASE + i, gConfig.MRU(i) )
      End If
   Next
   
   If GetMenuItemCount(hSubMenu) = 0 Then
      AppendMenu( hSubMenu, MF_STRING Or MF_GRAYED Or MF_DISABLED, 0, L(11,"(Empty)") )
   Else
      AppendMenu( hSubMenu, MF_SEPARATOR, 0, "" )
      AppendMenu( hSubMenu, MF_STRING, IDM_MRUCLEAR, L(225, "Clear this list") )
   End If
   
   Function = 0
End Function
   
''
''
Function UpdateMRUList( ByVal wzFilename As WString Ptr ) As Long

   Dim idxExistsAt as Long = -1   ' index of where filename already exists 
   Dim i           as Long
   Dim wzFile      As WString * MAX_PATH
   
   if gProjectLoading then exit function

   ' Search the 10 MRU items to see if the filename already exists and
   ' also for an available entry to put the filename in.
   For i = 0 To 9
      wzFile = gConfig.MRU(i)
      wzFile = ProcessFromCurdrive(wzFile)
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
   gConfig.SaveConfigFile

   Function = 0
End Function  



''  PROJECTS  ''

''
''
Function GetMRUProjectMenuHandle() As HMENU
   Dim hFileMenu As HMENU = GetSubMenu( HWND_FRMMAIN_TOPMENU, 4 ) 
   Dim hSubMenu  As HMENU = GetSubMenu( hFileMenu, 3 )
   Function = hSubMenu
End Function


''
''
Function OpenMRUProjectFile( ByVal HWnd As HWnd, ByVal wID As Long, ByVal pwszFilename As WString Ptr = 0 ) As Long
   Dim wzFile As WString * MAX_PATH
   Dim idx As Long
   
   If pwszFilename Then
      wzFile = *pwszFilename
   Else   
      wzFile = gConfig.MRUPROJECT(wID - IDC_MRUPROJECTBASE)
   End If
   
   ' If the project is already open then simply set the project as active
   ' rather than opening an additional copy of the project.
   idx = gApp.GetProjectIndexByFilename(wzFile)
   if idx > 0 THEN 
      gApp.SetActiveProject( gApp.Projects(idx).hExplorerRootNode )
      exit function
   END IF
   
   ' If a project is not yet open then close any files that may be open in the editor.
   If gApp.IsProjectActive = false Then 
      ' If the Find/Replace window is open then close it.
      DestroyWindow(HWND_FRMFINDREPLACE)
      if OnCommand_FileClose(HWnd, EFC_CLOSEALL) = 0 then
         ' save was cancelled
         exit function
      end if   
      gApp.NonProjectNotes = AfxGetWindowText(GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTNOTES))
      ' Clear the FILES explorer node
      if gApp.Projects(0).hExplorerRootNode THEN
         dim as hwnd hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
         Treeview_DeleteItem( hTree, gApp.Projects(0).hExplorerRootNode )
         gApp.Projects(0).hExplorerRootNode = 0
      END IF
   end if
         
   ' Clear any previous info from the Output windows
   AfxSetWindowText( GetDlgItem( HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTLOGFILE), "" )
   ListView_DeleteAllItems( GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_LISTVIEW) )
   ListView_DeleteAllItems( GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_LVTODO) )
   AfxSetWindowText(GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTNOTES), "")
   
   ' Open the project
   If gConfig.ProjectLoadFromFile(wzFile) Then
      ' Update the most recently used project list 
      UpdateMRUProjectList(@wzFile)
   End If

   ' Position all of the controls into place
   PositionExplorerWindows(HWND_FRMEXPLORER)
   frmMain_PositionWindows(HWnd)

   Function = 0
End Function

''
''
Function UpdateMRUProjectMenu( ByVal hMenu As HMENU ) As Long
   ' Modify the actual topmenu display items based on the MRU strings
   Dim hSubMenu As HMENU = GetMRUProjectMenuHandle()
   Dim i        As Long 
 
   If hMenu <> hSubMenu Then Exit Function
   
   ' Remove the existing items
   Do Until GetMenuItemCount(hSubMenu) = 0
      RemoveMenu( hSubMenu, 0, MF_BYPOSITION )
   Loop
   
   ' Add the new items
   For i = 0 To 9
      ' If the file no longer exists then remove it from the MRU list
      gConfig.MRUProject(i) = ProcessFromCurdrive(gConfig.MRUProject(i))
      If AfxFileExists(gConfig.MRUProject(i)) = 0 Then
         gConfig.MRUProject(i) = ""
      Else
         AppendMenu( hSubMenu, MF_STRING, IDC_MRUPROJECTBASE + i, gConfig.MRUProject(i) )
      End If
   Next
   
   If GetMenuItemCount(hSubMenu) = 0 Then
      AppendMenu( hSubMenu, MF_STRING Or MF_GRAYED Or MF_DISABLED, 0, L(11,"(Empty)") )
   Else
      AppendMenu( hSubMenu, MF_SEPARATOR, 0, "" )
      AppendMenu( hSubMenu, MF_STRING, IDM_MRUPROJECTCLEAR, L(225, "Clear this list") )
   End If
   
   Function = 0
End Function
   
''
''
Function UpdateMRUProjectList( ByVal wzFilename As WString Ptr ) As Long

   Dim idxExistsAt As Long = -1   ' index of where filename already exists 
   Dim i           As Long
   
   ' Search the 10 MRU items to see if the filename already exists and
   ' also for an available entry to put the filename in.
   For i = 0 To 9
      gConfig.MRUProject(i) = ProcessFromCurdrive(gConfig.MRUProject(i))
      If Ucase(gConfig.MRUProject(i)) = Ucase(*wzFilename) Then
         idxExistsAt = i:  Exit For
      End If   
   Next
   
   ' If does not already exist then we simply need to move all entries
   ' down and add at the beginning of the list. The original 10th entry
   ' will now drop off the list.
   If idxExistsAt = -1 Then idxExistsAt = 9

   For i = idxExistsAt To 1 Step -1
      gConfig.MRUProject(i) = gConfig.MRUProject(i-1)
   Next
   gConfig.MRUProject(0) = *wzFilename 

   ' Save the configuration file
   gConfig.SaveConfigFile

   Function = 0
End Function  












