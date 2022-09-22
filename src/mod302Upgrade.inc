'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2022 Paul Squires, PlanetSquires Software
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


' WinFBE Version 3.0.2+ requires that older version form file formats get upgraded
' to the new json format and also separated out from the source code file. There
' will now be two files: (1) *.frm for the json form definitions, and (2) *.inc/bas
' for the actual form code.

#include once "mod302Upgrade.bi"

' ========================================================================================
' Upgrade pre-version 3.02 form files to new format
' ========================================================================================
public function FormUpgrade302Format( byval pDoc as clsDocument ptr ) as boolean

   if pDoc = 0 then exit function
   
   ' WinFBE Version 3.0.2+ requires that older version form file formats get upgraded
   ' to the new json format and also separated out from the source code file. There
   ' will now be two files: (1) *.frm for the json form definitions, and (2) *.inc/bas
   ' for the actual form code.
   if (ConvertWinFBEversion(pDoc->wszFormVersion) >= ConvertWinFBEversion("3.0.2")) orelse _
      (ConvertWinFBEversion(pDoc->wszFormVersion) = 0) then  'b/c upgraded form files do not have version loaded yet
      ' We are already using new form file format so simply exit
      return false
   end if
   
   ' Commence conversion to the new json file format
   dim as CWSTR wszCodeFilename = pDoc->DiskFilename

   dim as CWSTR wszFormFilename = _
      AfxStrPathname( "PATH", gApp.ProjectFilename ) & _
      AfxStrPathName( "NAME", pDoc->DiskFilename ) & ".frm"
   
   if AfxFileExists(wszFormFilename) then
      AfxMsg( "Failed to upgrade form file to new format." & vbCrLf & _
              "REASON: File already exists. Please remove or rename the following file:" & vbCrLf & vbCrLf &_
              wszFormFilename ) 
      return false        
   end if
   
   dim pStream as CTextStream  
   if pStream.Create(wszFormFilename, true, false) <> S_OK then return true   ' error

   dim pCtrl as clsControl ptr
   dim as CWSTR wst 
   dim as CWSTR wcomma = ","
   
   wst = _
   qstr("'form':true")        & wcomma & _
   qstr("'version':'3.0.2'")  & wcomma & _
   qstr("'lockcontrols':"     & iif(pDoc->bLockControls, "true", "false")) & wcomma & _
   qstr("'snaplines':"        & iif(pDoc->bSnapLines, "true", "false")) 

   ' Save Images(if applicable)
   dim as long numImageItems = ubound(pDoc->AllImages) - lbound(pDoc->AllImages) + 1
   if numImageItems > 0 then
      wst = wst & wcomma & _
      qstr("'images': [")
      for ii as long = lbound(pDoc->AllImages) to ubound(pDoc->AllImages)
         dim as CWSTR wszRelative
         dim as CWSTR wszImageFilename = pDoc->AllImages(ii).wszFilename
         
         ' Attempt to convert the image file name to relative path
         if AfxFileExists( pDoc->DiskFilename ) then
            wszRelative = AfxPathRelativePathTo( pDoc->DiskFilename, FILE_ATTRIBUTE_NORMAL, wszImageFilename, FILE_ATTRIBUTE_NORMAL)
            if AfxPathIsRelative(wszRelative) then wszImageFilename = wszRelative
         end if
         
         wst = wst & _
         "{" & _
         qstr("'imagename':'"        & pDoc->AllImages(ii).wszImageName & "'") & wcomma & _ 
         qstr("'filename':'"         & AfxStrReplace(ProcessToCurdriveProject(wszImageFilename), "\", "\\") & "'") & wcomma & _ 
         qstr("'resourcetype':'"     & pDoc->AllImages(ii).wszFormat & "'") & _ 
         "},"
      next
      wst = rtrim(wst, ",") & "]"
   end if

   ' Save MainMenu (if applicable)
   dim as long numMenuItems = ubound(pDoc->MenuItems) - lbound(pDoc->MenuItems) + 1
   if numMenuItems > 0 then
      wst = wst & wcomma & _
      qstr("'mainmenu':{") & _
      qstr("'display':" & iif(pDoc->GenerateMenu, "true", "false")) & wcomma 

      wst = wst & _
      qstr("'items': [")
      for ii as long = lbound(pDoc->MenuItems) to ubound(pDoc->MenuItems)
         wst = wst & _
         "{" & _
         qstr("'name':'"     & pDoc->MenuItems(ii).wszName & "'") & wcomma & _
         qstr("'caption':'"  & pDoc->MenuItems(ii).wszCaption & "'") & wcomma & _
         qstr("'indent':"    & pDoc->MenuItems(ii).nIndent) & wcomma & _
         qstr("'alt':"       & pDoc->MenuItems(ii).chkAlt) & wcomma & _
         qstr("'shift':"     & pDoc->MenuItems(ii).chkShift) & wcomma & _
         qstr("'ctrl':"      & pDoc->MenuItems(ii).chkCtrl) & wcomma & _
         qstr("'shortcut':'" & pDoc->MenuItems(ii).wszShortcut & "'") & wcomma & _
         qstr("'checked':"   & pDoc->MenuItems(ii).chkChecked) & wcomma & _
         qstr("'grayed':"    & pDoc->MenuItems(ii).chkGrayed) & _
         "}," 
      next
      wst = rtrim(wst, ",") & "]}"
   end if

   ' Save ToolBar items (if applicable)
   dim as long numToolBarItems = Ubound(pDoc->ToolBarItems) - lbound(pDoc->ToolBarItems) + 1
   if numToolBarItems > 0 then
      wst = wst & wcomma & _
      qstr("'toolbar':{") & _
      qstr("'display':" & iif(pDoc->GenerateToolBar, "true", "false")) & wcomma & _
      qstr("'size':'" & pDoc->wszToolBarSize & "'") & wcomma 

      wst = wst & _
      qstr("'items': [")
      for ii as long = lbound(pDoc->ToolBarItems) to ubound(pDoc->ToolBarItems)
         wst = wst & _
         "{" & _
         qstr("'name':'"           & pDoc->ToolBarItems(ii).wszName & "'") & wcomma & _
         qstr("'type':'"           & pDoc->ToolBarItems(ii).wszButtonType & "'") & wcomma & _
         qstr("'tooltip':'"        & pDoc->ToolBarItems(ii).wszTooltip & "'") & wcomma & _
         qstr("'hotimage':'"       & pDoc->ToolBarItems(ii).pPropHotImage.wszPropValue & "'") & wcomma & _
         qstr("'normalimage':'"    & pDoc->ToolBarItems(ii).pPropNormalImage.wszPropValue & "'") & wcomma & _
         qstr("'disabledimage':'"  & pDoc->ToolBarItems(ii).pPropDisabledImage.wszPropValue & "'") & _
         "}," 
      next
      wst = rtrim(wst, ",") & "]}"
   end if


   ' Save StatusBar Panels (if applicable)
   dim as long numPanelItems = Ubound(pDoc->PanelItems) - lbound(pDoc->PanelItems) + 1
   if numPanelItems > 0 then
      wst = wst & wcomma & _
      qstr("'statusbar':{") & _
      qstr("'display':" & iif(pDoc->GenerateStatusBar, "true", "false")) & wcomma 

      wst = wst & _
      qstr("'items': [")
      for ii as long = lbound(pDoc->PanelItems) to ubound(pDoc->PanelItems)
         wst = wst & _
         "{" & _
         qstr("'name':'"           & pDoc->PanelItems(ii).wszName & "'") & wcomma & _
         qstr("'text':'"           & pDoc->PanelItems(ii).wszText & "'") & wcomma & _
         qstr("'tooltip':'"        & pDoc->PanelItems(ii).wszTooltip & "'") & wcomma & _
         qstr("'alignment':'"      & pDoc->PanelItems(ii).wszAlignment & "'") & wcomma & _
         qstr("'autosize':'"       & pDoc->PanelItems(ii).wszAutoSize & "'") & wcomma & _
         qstr("'width':"           & pDoc->PanelItems(ii).wszWidth) & wcomma & _
         qstr("'minwidth':"        & pDoc->PanelItems(ii).wszMinWidth) & wcomma & _
         qstr("'image':'"          & pDoc->PanelItems(ii).pProp.wszPropValue & "'") & wcomma & _
         qstr("'backcolor':'"      & pDoc->PanelItems(ii).wszBackColor & "'") & wcomma & _
         qstr("'backcolorhot':'"   & pDoc->PanelItems(ii).wszBackColorHot & "'") & wcomma & _
         qstr("'forecolor':'"      & pDoc->PanelItems(ii).wszForeColor & "'") & wcomma & _
         qstr("'forecolorhot':'"   & pDoc->PanelItems(ii).wszForeColorHot & "'") & _
         "}," 
              ' BorderStyle is deprecated as of v2.0.4 as it has no effect
              ' in WinFBE programs where Windows Themes are enabled.
              '"'     PANELBORDERSTYLE=" & pDoc->PanelItems(ii).wszBorderStyle & vbcrlf & _
      next
      wst = rtrim(wst, ",") & "]}"
   end if
   
   
   ' Iterate all of the controls on the form
   if pDoc->Controls.Count then
      wst = wst & wcomma & _
      qstr("'controls':[") 
            
      for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
         pCtrl = pDoc->Controls.ItemAt(i)
         if pCtrl then

            wst = wst & _
            "{" & _
            qstr("'type':'" & GetToolBoxName(pCtrl->ControlType) & "'") & wcomma & _
            qstr("'properties': [") 

            for ii as long = lbound(pCtrl->Properties) to ubound(pCtrl->Properties)
               wst = wst & "{" & _
               qstr("'name':'" & pCtrl->Properties(ii).wszPropName & "'") & wcomma & _
               qstr("'value':'" & pCtrl->Properties(ii).wszPropValue & "'") & _
               "},"
            next
            
            wst = rtrim(wst, ",") & "]," & _
            qstr("'events': [") 

            for ii as long = lbound(pCtrl->Events) to ubound(pCtrl->Events)
               ' Only need to output the names of the Events that have been ticked as being in use.
               if pCtrl->Events(ii).bIsSelected then
                  wst = wst & _
                  qstr("'" & pCtrl->Events(ii).wszEventName & "'") & wcomma 
               end if   
            next
            wst = rtrim(wst, ",") & "]"
      
         end if
         
         wst = wst & "},"
      next

      wst = rtrim(wst, ",") & "]"
   end if

   ' Convert UTF-16 to UTF-8 encoding
   wst = "{" & wst & "}"
   pStream.WriteLine wst.utf8

   pStream.Close
   
   ' The new frm Form file has now been created so we can remove all code in the
   ' existing code file.
   dim as long iTab = gTTabCtl.GetTabIndexFromFilename( wszCodeFilename )
   dim as boolean bShowInTab = iif( iTab = -1, false, true )

   ' We need to keep the Isdesigner to true because the close/remove routines
   ' depend on this flag in order to destroy windows.
   pDoc->IsDesigner = false
   pDoc->wszFormMetaData = ""
   pDoc->wszFormCodeGen = ""
   pDoc->SaveFile
  
   ' Close the current file in order to destroy and allow re-creation
   ' We need to keep the Isdesigner to true because the close/remove routines
   ' depend on this flag in order to destroy windows.
   pDoc->IsDesigner = true
   if gApp.IsProjectActive then
      OnCommand_ProjectRemove( 0, pDoc )

dim p as clsDocument ptr = gApp.pDocList
do until p = 0
   ? p->DiskFilename
   p = p->pDocnext
loop


   else
      if iTab > -1 then
         OnCommand_FileClose( HWND_FRMMAIN, EFC_CLOSECURRENT, iTab )
      end if
   end if
   
   ' Open the two new files
   pDoc = frmMain_OpenFileSafely(HWND_FRMMAIN, _
                                 False, _      ' bIsNewFile
                                 False, _      ' bIsTemplate
                                 bShowInTab, _ ' bShowInTab
                                 false, _      ' bIsInclude
                                 wszCodeFilename, _ ' pwszName
                                 0, _          ' pDocIn
                                 false _       ' bIsDesigner
                                 )     
   pDoc->ProjectFileType = FILETYPE_NORMAL 
   
   ' Add the new frm Form file to the Project/Explorer
   pDoc = frmMain_OpenFileSafely(HWND_FRMMAIN, _
                                 False, _      ' bIsNewFile
                                 False, _      ' bIsTemplate
                                 bShowInTab, _ ' bShowInTab
                                 false, _      ' bIsInclude
                                 wszFormFilename, _ ' pwszName
                                 0, _          ' pDocIn
                                 true _        ' bIsDesigner
                                 )     
   pDoc->ProjectFileType = FILETYPE_NORMAL 

   LoadExplorerFiles()
   LoadFunctionsFiles()
               
   return true  ' new format was created
end function


