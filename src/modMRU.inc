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

#include once "modMRU.bi"
#include once "frmRecent.bi"



''
''
private Function GetMRUMenuHandle() as HMENU
   Dim hFileMenu as HMENU = GetSubMenu( HWND_FRMMAIN_TOPMENU, FILES_MENU_POSITION ) 
   Dim hSubMenu  As HMENU = GetSubMenu( hFileMenu, MRUFILES_MENU_POSITION )
   Function = hSubMenu
End Function


''
''
public Function OpenMRUFile( ByVal HWnd As HWnd, _
                             ByVal wID As Long _
                             ) As Long
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
public Function ClearMRUlist( ByVal wID As Long ) As Long

   Dim As Long i 
   
   Select Case wID
      Case IDM_MRUCLEAR
         For i = 0 To 9
            gConfig.MRU(i) = ""
         Next
         frmRecent_LoadTreeview
      Case IDM_MRUPROJECTCLEAR
         For i = 0 To 9
            gConfig.MRUPROJECT(i) = ""
         Next
         frmRecent_LoadTreeview
   End Select
   
   frmRecent_PositionWindows

   Function = 0
End Function


''
''
public Function CreateMRUpopup() As HMENU
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
public Function UpdateMRUMenu( ByVal hMenu As HMENU ) As Long
   ' Modify the actual topmenu display items based on the MRU strings
   Dim hSubMenu As HMENU = GetMRUMenuHandle()
   Dim i        As Long 
   
   If hMenu <> hSubMenu Then Exit Function
   
   ' Remove the existing items
   Do Until GetMenuItemCount( hSubMenu ) = 0
      RemoveMenu( hSubMenu, 0, MF_BYPOSITION )
   Loop
   
   ' Add the new items
   For i = 0 To 9
      ' If the file no longer exists then remove it from the MRU list
      gConfig.MRU(i) = ProcessFromCurdrive(gConfig.MRU(i))
      If AfxFileExists( gConfig.MRU(i) ) Then
         AppendMenu( hSubMenu, MF_STRING, IDC_MRUBASE + i, gConfig.MRU(i) )
      Else
         gConfig.MRU(i) = ""
      End If
   Next
   
   If GetMenuItemCount( hSubMenu ) = 0 Then
      AppendMenu( hSubMenu, MF_STRING Or MF_GRAYED Or MF_DISABLED, 0, L(11,"(Empty)") )
   Else
      AppendMenu( hSubMenu, MF_SEPARATOR, 0, "" )
      AppendMenu( hSubMenu, MF_STRING, IDM_MRUCLEAR, L(225, "Clear this list") )
   End If
   
   Function = 0
End Function
   
''
''
public Function UpdateMRUList( Byref wzFilename As WString ) As Long

   Dim idxExistsAt as Long = -1   ' index of where filename already exists 
   Dim i           as Long
   Dim wzFile      As WString * MAX_PATH
   
   if gApp.IsProjectLoading then exit function

   ' Search the 10 MRU items to see if the filename already exists and
   ' also for an available entry to put the filename in.
   For i = 0 To 9
      wzFile = gConfig.MRU(i)
      wzFile = ProcessFromCurdrive(wzFile)
      If Ucase(wzFile) = Ucase(wzFilename) Then
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
   gConfig.MRU(0) = wzFilename 
   
   ' Save the MRU to the configuration file. Only write the MRU items
   ' the ini file rather than overwriting the whole file. This enables
   ' WinFBE to work better with external tools that manually modify the
   ' WinFBE.ini file.
   gConfig.WriteMRU

   Function = 0
End Function  



''  PROJECTS  ''

''
''
private Function GetMRUProjectMenuHandle() As HMENU
   Dim hFileMenu As HMENU = GetSubMenu( HWND_FRMMAIN_TOPMENU, PROJECT_MENU_POSITION ) 
   Dim hSubMenu  As HMENU = GetSubMenu( hFileMenu, MRUPROJECTS_MENU_POSITION )
   Function = hSubMenu
End Function


''
''
public Function OpenMRUProjectFile( ByVal HWnd As HWnd, _
                                    ByVal wID As Long _
                                    ) As Long
   Dim wzFile As WString * MAX_PATH
   wzFile = gConfig.MRUPROJECT(wID - IDC_MRUPROJECTBASE)
   frmMain_OpenProjectSafely(HWnd, wzFile) 
   Function = 0
End Function

''
''
public Function UpdateMRUProjectMenu( ByVal hMenu As HMENU ) As Long
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
public Function UpdateMRUProjectList( Byval wszFilename As CWSTR ) As Long

   Dim idxExistsAt As Long = -1   ' index of where filename already exists 
   Dim i           As Long
   
   ' Search the 10 MRU items to see if the filename already exists and
   ' also for an available entry to put the filename in.
   For i = 0 To 9
      gConfig.MRUProject(i) = ProcessFromCurdrive(gConfig.MRUProject(i))
      If Ucase(gConfig.MRUProject(i)) = Ucase(wszFilename) Then
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
   gConfig.MRUProject(0) = wszFilename 

   ' Save the MRU to the configuration file. Only write the MRU items
   ' the ini file rather than overwriting the whole file. This enables
   ' WinFBE to work better with external tools that manually modify the
   ' WinFBE.ini file.
   gConfig.WriteMRUProjects

   Function = 0
End Function  












