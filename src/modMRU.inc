'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2025 Paul Squires, PlanetSquires Software
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

''
''
private function clearMRUFilesItems( byval nParentID as long ) as long
   ' clear out any existing items that match the nParentID so that they can be
   ' reused when the MRU popup menu is created and shown. This function is used
   ' for Files and Projects and UserTools MRU lists.
   for i as long = lbound(gTopMenu) to ubound(gTopMenu)
      if gTopMenu(i).nParentID = nParentID then
         gTopMenu(i).nParentID = 0
         gTopMenu(i).nID = 0
         gTopMenu(i).nChildID = 0
      end if
   next
   function = 0
end function

''
''
private function getNextFreeMRUindex() as long
   ' gets the first available free MRU index in the gTopMenu array. If there are
   ' none free then the gTopMenu array is extended and that index is returned. This
   ' function is used for both Files and Projects MRU lists.
   dim as long nFoundAt = -1
   for i as long = lbound(gTopMenu) to ubound(gTopMenu)
      if gTopMenu(i).nParentID = 0 then return i
   next
   if nFoundAt = -1 then
      nFoundAt = ubound(gTopMenu) + 1
      redim preserve gTopMenu(nFoundAt) as TOPMENU_TYPE
   end if
   function = nFoundAt
end function


''
''
public Function updateMRUFilesItems() as long
   ' clear MRU items already existing in the gTopMenu array. We overwrite and extend
   ' the gTopMenu array rather than erase it because existing menus depend on the 
   ' array index that aready exist. The function also returns the width to use for
   ' the resulting popup menu (based on text width metrics of each filename).
   clearMRUFilesItems( IDM_MRU )
   
   dim wszText as WSTRING * 256
   dim as boolean hasMRUfiles = false
   dim as long txtWidth = 0
   dim as long nMenuWidth = 0
   dim as long padding = AfxScaleX(12)

   For i As Long = 0 To 9
      ' If the file no longer exists then remove it from the MRU list
      gConfig.MRU(i) = ProcessFromCurdriveApp( gConfig.MRU(i) )
      If AfxFileExists( gConfig.MRU(i) ) = 0 Then
         gConfig.MRU(i) = ""
      Else
         hasMRUfiles = true
         setTopMenuMRUItem( getNextFreeMRUindex(), IDM_MRU, IDM_MRUBASE + i, 0, false, false )
         wszText = gConfig.MRU(i)
         txtWidth = getTextWidth( HWND_FRMMAIN_MENUBAR, wszText, ghMenuBar.hFontMenuBar, 30 )
         if txtWidth > nMenuWidth then nMenuWidth = txtWidth
      End If
   Next
   
   if hasMRUfiles = false then
      setTopMenuMRUItem( getNextFreeMRUindex(), IDM_MRU, IDM_MRUFILES, 0, false, false )  ' (empty)
   else
      setTopMenuMRUItem( getNextFreeMRUindex(), IDM_MRU, IDM_MRUFILES, 0, false, true ) 
      setTopMenuMRUItem( getNextFreeMRUindex(), IDM_MRU, IDM_MRUCLEAR, 0, false, false )
   end if   
   
   function = nMenuWidth + padding
end function


''
''
public Function updateMRUProjectFilesItems() as long
   ' clear MRU items already existing in the gTopMenu array. We overwrite and extend
   ' the gTopMenu array rather than erase it because existing menus depend on the 
   ' array index that aready eist.
   clearMRUFilesItems( IDM_MRUPROJECT )
   
   dim wszText as WSTRING * 256
   dim as boolean hasMRUfiles = false
   dim as long txtWidth = 0
   dim as long nMenuWidth = 0
   dim as long padding = AfxScaleX(12)

   For i As Long = 0 To 9
      ' If the file no longer exists then remove it from the MRU list
      gConfig.MRUPROJECT(i) = ProcessFromCurdriveApp( gConfig.MRUPROJECT(i) )
      If AfxFileExists( gConfig.MRUPROJECT(i) ) = 0 Then
         gConfig.MRUPROJECT(i) = ""
      Else
         hasMRUfiles = true
         setTopMenuMRUItem( getNextFreeMRUindex(), IDM_MRUPROJECT, IDM_MRUPROJECTBASE + i, 0, false, false )
         wszText = gConfig.MRUPROJECT(i)
         txtWidth = getTextWidth( HWND_FRMMAIN_MENUBAR, wszText, ghMenuBar.hFontMenuBar, 30 )
         if txtWidth > nMenuWidth then nMenuWidth = txtWidth
      End If
   Next
   
   if hasMRUfiles = false then
      setTopMenuMRUItem( getNextFreeMRUindex(), IDM_MRUPROJECT, IDM_MRUPROJECTFILES, 0, false, false )  ' (empty)
   else
      setTopMenuMRUItem( getNextFreeMRUindex(), IDM_MRUPROJECT, IDM_MRUPROJECTFILES, 0, false, true ) 
      setTopMenuMRUItem( getNextFreeMRUindex(), IDM_MRUPROJECT, IDM_MRUPROJECTCLEAR, 0, false, false )
   end if   
   
   function = nMenuWidth + padding
end function


''
''
public function OpenMRUFile( byval HWnd As HWnd, _
                             byval wID As long _
                             ) as long
   dim wszFile as WString * MAX_PATH
   dim pDoc as clsDocument ptr
   dim pDocIn as clsDocument ptr
   wszFile = gConfig.MRU(wID - IDM_MRUBASE)
   
   ' If file is already loaded then simply switch to that file
   ' within the editor.
   pDocIn = gApp.GetDocumentPtrByFilename( wszFile )

   pDoc = frmMain_OpenFileSafely( _
                  HWND_FRMMAIN, _
                  false, _    ' bIsNewFile
                  false, _    ' bIsTemplate
                  true, _     ' bShowInTab
                  false, _    ' bIsInclude
                  wszFile, _  ' wszName
                  pDocIn, _   ' pDocIn
                  IsFormFilename(wszFile) _
                  )

   ' Give this document a default project type depending on its file extension
   if (pDoc->IsNewFlag = false) andalso (pDoc->ProjectFileType = FILETYPE_UNDEFINED) then
      if ( gApp.IsProjectActive = true ) orelse ( gApp.IsProjectLoading = true ) then
         if pDoc->IsDesigner then 
            pDoc->ProjectFileType = FILETYPE_NORMAL 
         else
            gApp.ProjectSetFileType( pDoc, pDoc->ProjectFileType )
         end if
      end if   
   end if

   LoadExplorerFiles()
   LoadFunctionsFiles()

   function = 0
end function


''
''
public Function ClearMRUlist( ByVal wID As Long ) As Long
   Select Case wID
      Case IDM_MRUCLEAR
         For i as long = 0 To 9
            gConfig.MRU(i) = ""
         Next
      Case IDM_MRUPROJECTCLEAR
         For i as long = 0 To 9
            gConfig.MRUPROJECT(i) = ""
         Next
   End Select
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
      wzFile = ProcessFromCurdriveApp(wzFile)
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
public Function OpenMRUProjectFile( ByVal HWnd As HWnd, _
                                    ByVal wID As Long _
                                    ) As Long
   Dim wzFile As WString * MAX_PATH
   wzFile = gConfig.MRUPROJECT(wID - IDM_MRUPROJECTBASE)
   frmMain_OpenProjectSafely(HWnd, wzFile) 
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
      gConfig.MRUProject(i) = ProcessFromCurdriveApp(gConfig.MRUProject(i))
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



