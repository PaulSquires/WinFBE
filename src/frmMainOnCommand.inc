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

#include once "frmMainOnCommand.bi"

#include once "frmOutput.bi"
#include once "clsApp.bi"
#include once "clsDocument.bi"


' ========================================================================================
' Save project file to disk
' ========================================================================================
private Function OnCommand_ProjectSave( byval HWnd As HWnd, _
                                        ByVal bSaveAs As BOOLEAN = False _
                                        ) As LRESULT
   ' Save all dirty files
   OnCommand_FileSaveAll( HWnd )
   
   ' Do the actual saving to disk of the Project file
   gApp.SaveProject(bSaveAs)

   frmMain_SetFocusToCurrentCodeWindow

   Function = 0
End Function


' ========================================================================================
' Close currently active project
' ========================================================================================
private Function OnCommand_ProjectClose( ByVal HWnd As HWnd ) As LRESULT

   ' Save the project configuration to disk and any dirty files
   ' Do the actual saving to disk of the Project file
   gApp.SaveProject(false)
   if OnCommand_FileClose( HWnd, EFC_CLOSEALL ) = false then exit function
   
   ' If the entire app is closing down then we do not have to process the
   ' remaining project cleanup commands. 
   if gApp.IsShutdown then return true
      
   ' If we get this far then we can safely remove all of the open pDocs
   gApp.RemoveAllDocuments
   
   ' Clear the Files and Function comboboxes
   ComboBox_ResetContent( GetDlgItem(HWND_FRMMAIN, IDC_FRMMAIN_COMBOFILES) )
   ComboBox_ResetContent( GetDlgItem(HWND_FRMMAIN, IDC_FRMMAIN_COMBOFUNCTIONS) )

   ' Delete all nodes in the Explorer treeview and reset it to "FILES" mode.
   dim as hwnd hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
   Treeview_DeleteAllItems(hTree)
   gApp.hExplorerRootNode = FF_TreeView_InsertItem(hTree, TVI_ROOT, wstr("FILES") )
   Treeview_SetBold(hTree, gApp.hExplorerRootNode, TRUE)

   ' If the Find/Replace window is open then close it.
   DestroyWindow(HWND_FRMFINDREPLACE)
   DestroyWindow( HWND_FRMFUNCTIONLIST )
   DestroyWindow(HWND_FRMVDTOOLBOX)

   ' Reset the Project variables
   gApp.IsProjectActive      = false
   gApp.ProjectName          = ""
   gApp.ProjectFilename      = ""
   gApp.ProjectBuild         = ""
   gApp.ProjectOther32       = ""
   gApp.ProjectOther64       = ""
   gApp.ProjectManifest      = false
   gApp.ProjectNotes         = ""
   gApp.ProjectCommandLine   = ""
       
   ' Clear any previous info from the Output windows
   frmOutput_ResetAllControls
   
   ' Reset the ProjectNotes text box with the generic non-project related notes
   AfxSetWindowText(GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTNOTES), gApp.NonProjectNotes)

   frmExplorer_PositionWindows
   frmMain_PositionWindows
    
   frmMain_SetStatusbar
   frmMain_SetFocusToCurrentCodeWindow

   Function = true
End Function


' ========================================================================================
' Create a new project
' ========================================================================================
private Function OnCommand_ProjectNew( ByVal HWnd As HWnd ) As LRESULT

   ' Ensure that the Project Manager window has been created (show it now as well)
   frmProjectOptions_Show( HWnd, true ) 
   
   frmExplorer_PositionWindows
   frmMain_PositionWindows
   
   ' This will update the main window to show the project name in the window caption
   frmMain_SetFocusToCurrentCodeWindow
   
   Function = 0
End Function


' ========================================================================================
' Open a Project
' ========================================================================================
private Function OnCommand_ProjectOpen( ByVal HWnd As HWnd ) As LRESULT

   ' Display the Open File Dialog
   Dim pwszName As WString Ptr = AfxIFileOpenDialogW(HWnd, IDM_PROJECTOPEN)
   If pwszName Then 
      ' Pass the info to our generic project open function to handle everything.
      frmMain_OpenProjectSafely(HWND_FRMMAIN, *pwszName) 
      CoTaskMemFree(pwszName)
   End If
   frmMain_SetFocusToCurrentCodeWindow

   Function = 0
End Function

' ========================================================================================
' Create a new visual designer window
' ========================================================================================
private Function OnCommand_DesignerNewForm( ByVal HWnd As HWnd ) As LRESULT
   frmMain_OpenFileSafely(HWnd, _
                           True,  _   ' bIsNewFile
                           False, _   ' bIsTemplate
                           True,  _   ' bShowInTab
                           false, _   ' bIsInclude
                           0, _       ' pwszName
                           0, _       ' pDocIn
                           true )     ' Create a visual designer   
   Function = 0
End Function


' ========================================================================================
' Align controls on the form
' ========================================================================================
private function OnCommand_DesignerAlign( byval HWND as HWND, _
                                          byval id as long _
                                          ) as LRESULT

   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
   if pDoc = 0 THEN exit function

   dim pCtrlActive as clsControl ptr = pDoc->Controls.GetActiveControl()
   dim pCtrl as clsControl ptr
   dim as RECT rcCtrl, rcActive

   dim as long nWidth, nHeight
   dim as long nWidthActive, nHeightActive
   
   if pCtrlActive = 0 THEN exit function
   rcActive = GetControlRECT(pCtrlActive)
   nWidthActive  = rcActive.Right  - rcActive.Left
   nHeightActive = rcActive.Bottom - rcActive.Top 
   
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)
      if pCtrl->IsSelected THEN
         if IsControlLocked(pDoc, pCtrl) then continue for
         ' Skip processing the active control because that is unnecessary.
         if pCtrl = pCtrlActive THEN continue for
         rcCtrl = GetControlRECT(pCtrl)
         nWidth  = rcCtrl.Right  - rcCtrl.Left
         nHeight = rcCtrl.Bottom - rcCtrl.Top 
         select CASE id
            case IDM_ALIGNLEFTS
               SetControlProperty(pCtrl, "LEFT", str(rcActive.left))
               ApplyControlProperties(pDoc, pCtrl)
            case IDM_ALIGNTOPS
               SetControlProperty(pCtrl, "TOP", str(rcActive.top))
               ApplyControlProperties(pDoc, pCtrl)
            case IDM_ALIGNRIGHTS
               SetControlProperty(pCtrl, "LEFT", str(rcctrl.left + (rcActive.right-rcCtrl.right)) )
               ApplyControlProperties(pDoc, pCtrl)
            case IDM_ALIGNBOTTOMS 
               SetControlProperty(pCtrl, "TOP", str(rcCtrl.top + (rcActive.bottom-rcCtrl.bottom)) )
               ApplyControlProperties(pDoc, pCtrl)
            case IDM_ALIGNCENTERS  
               SetControlProperty(pCtrl, "LEFT", _
                  str(rcActive.Left + (nWidthActive / 2) - ( nWidth / 2 )) )
               ApplyControlProperties(pDoc, pCtrl)
            case IDM_ALIGNMIDDLES  
               SetControlProperty(pCtrl, "TOP", _
                  str(rcActive.Top + (nHeightActive / 2) - ( nHeight / 2 )) )
               ApplyControlProperties(pDoc, pCtrl)
            case IDM_SAMEWIDTHS  
               SetControlProperty(pCtrl, "WIDTH", str(nWidthActive))
               ApplyControlProperties(pDoc, pCtrl)
            case IDM_SAMEHEIGHTS
               SetControlProperty(pCtrl, "HEIGHT", str(nHeightActive))
               ApplyControlProperties(pDoc, pCtrl)
            case IDM_SAMEBOTH
               pCtrl->SuspendLayout = true
               SetControlProperty(pCtrl, "WIDTH", str(nWidthActive))
               SetControlProperty(pCtrl, "HEIGHT", str(nHeightActive))
               ApplyControlProperties(pDoc, pCtrl)
               pCtrl->SuspendLayout = false
         END SELECT
      END IF
   next

   ' Ensure the grab handles are redrawn
   AfxRedrawWindow(pDoc->hWndForm)

   function = 0
end function


' ========================================================================================
' Center controls on the form
' ========================================================================================
private function OnCommand_DesignerCenter( byval HWND as HWND, _
                                           byval id as long _
                                           ) as LRESULT

   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)
   If pWindow = 0 Then Exit Function

   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
   if pDoc = 0 THEN exit function

   dim pCtrl as clsControl ptr

   dim as RECT rcCtrl, rcForm
   dim as long nLeft, nTop, nMaxRight, nMaxBottom
   dim as long nHorizSpacing, nVertSpacing, nFormWidth, nFormHeight
   dim as long nMinLeft = 999999999
   dim as long nMinTop  = 999999999
   
   ' Calculate the total distance from the far left (or top) control to the far
   ' right (or bottom) control and subtract the widths of each control. This will
   ' give us the amount of 'white' space. Divide that white space by
   ' half to determine the amount to put on the left/top and right/bottom of the
   ' group of controls.
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)
      if pCtrl->ControlType = CTRL_FORM THEN continue for
      if pCtrl->IsSelected = false THEN continue for
      rcCtrl     = GetControlRECT(pCtrl)
      nMinLeft   = Min( nMinLeft, rcCtrl.left )
      nMaxRight  = Max( nMaxRight, rcCtrl.right )
      nMinTop    = Min( nMinTop, rcCtrl.top )
      nMaxBottom = Max( nMaxBottom, rcCtrl.bottom )
   next

   ' Get the client area of the design form
   GetClientRect(pDoc->hWndForm, @rcForm)
   nFormWidth  = (rcForm.right - rcForm.left) / pWindow->rxRatio
   nFormHeight = (rcForm.bottom - rcForm.top) / pWindow->ryRatio
   
   ' Determine the value by which the group of controls needs to be shifted by.  
   nHorizSpacing = nMinLeft - ( (nFormWidth - (nMaxRight - nMinLeft)) / 2) 
   nVertSpacing  = nMinTop - ( (nFormHeight - (nMaxBottom - nMinTop) ) / 2) 

   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)
      if pCtrl->ControlType = CTRL_FORM THEN continue for
      if pCtrl->IsSelected = false THEN continue for
      if IsControlLocked(pDoc, pCtrl) then continue for
      select CASE id
         case IDM_CENTERHORIZ
            nLeft = val(GetControlProperty(pCtrl, "LEFT"))
            SetControlProperty(pCtrl, "LEFT", str(nLeft-nHorizSpacing))
            ApplyControlProperties(pDoc, pCtrl)
         case IDM_CENTERVERT
            nTop  = val(GetControlProperty(pCtrl, "TOP"))
            SetControlProperty(pCtrl, "TOP", str(nTop-nVertSpacing))
            ApplyControlProperties(pDoc, pCtrl)
         case IDM_CENTERBOTH   
            nLeft = val(GetControlProperty(pCtrl, "LEFT"))
            nTop  = val(GetControlProperty(pCtrl, "TOP"))
            pCtrl->SuspendLayout = true
            SetControlProperty(pCtrl, "LEFT", str(nLeft-nHorizSpacing))
            SetControlProperty(pCtrl, "TOP", str(nTop-nVertSpacing))
            ApplyControlProperties(pDoc, pCtrl)
            pCtrl->SuspendLayout = false
      END SELECT
   next

   ' Ensure the grab handles are redrawn
   AfxRedrawWindow(pDoc->hWndForm)

   function = 0
end function

' ========================================================================================
' Create a new editing window
' ========================================================================================
private Function OnCommand_FileNew( ByVal HWnd As HWnd ) As clsDocument ptr
   function = frmMain_OpenFileSafely(HWnd, _
                                     True,  _   ' bIsNewFile
                                     False, _   ' bIsTemplate
                                     True,  _   ' bShowInTab
                                     false, _   ' bIsInclude
                                     0, _       ' pwszName
                                     0 )        ' pDocIn

End Function


' ========================================================================================
' Open one or more files for editing
' ========================================================================================
private Function OnCommand_FileOpen( ByVal HWnd As HWnd, _
                                     byval bShowInTab as Boolean = true _
                                     ) As LRESULT

   dim pDoc as clsDocument ptr
   
   ' Display the Open File Dialog
   Dim pItems As IShellItemArray Ptr = AfxIFileOpenDialogMultiple(HWnd, IDM_FILEOPEN)
   If pItems = Null Then Exit Function
   Dim dwItemCount As Long, i As Long, pItem As IShellItem Ptr, pwszName As WString Ptr
   pItems->lpVtbl->GetCount(pItems, @dwItemCount)
   
   ' Variable length array to hold sequence of TabCtrl tabs to open. We save the sequence
   ' here and open them only after all over documents have been loaded. This will look
   ' visually more appealing for thos files that will display in the top Tab Control.
   Dim pDocTabs(dwItemCount - 1) As clsDocument Ptr

   dim as HCURSOR hCurSave = GetCursor()
   SetCursor( LoadCursor(0, IDC_WAIT) )
   pDoc = gTTabCtl.GetActiveDocumentPtr()
   if pDoc THEN SciExec( pDoc->hWindow(0), SCI_SETCURSOR, SC_CURSORWAIT, 0 )

   ' Prevent the constant updating of the app caption bar with filename
   gApp.IsFileLoading = true  

   ' Prevent Treeview flickering during adding
   dim as hwnd hTV = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
   LockWindowUpdate(hTV)

   gApp.FileLoadingCount = 0
   
   For i = 0 To dwItemCount - 1
      pItems->lpVtbl->GetItemAt(pItems, i, @pItem)
      If pItem Then
         pItem->lpVtbl->GetDisplayName(pItem, SIGDN_FILESYSPATH, @pwszName)
         If pwszName Then 
            ' Test to see if the file is already loaded in the editor. If it is, then
            ' bypass loading it again thereby creating multiple ghost instances.
            pDoc = gApp.GetDocumentPtrByFilename(pwszName) 
            if pDoc = 0 then
               pDoc = frmMain_OpenFileSafely(HWnd, _
                                             False, _    ' bIsNewFile
                                             False, _    ' bIsTemplate
                                             false,  _   ' bShowInTab
                                             false, _    ' bIsInclude
                                             pwszName, _ ' pwszName
                                             0 )         ' pDocIn
            end if
            if bShowInTab then pDocTabs(i) = pDoc
            CoTaskMemFree(pwszName)
            pwszName = Null
         End If
         pItem->lpVtbl->Release(pItem)
         pItem = Null
      End If
   Next
   pItems->lpVtbl->Release(pItems)
   
   ' Display all of the tabs
   dim as long iTab
   For i = LBound(pDocTabs) To UBound(pDocTabs)
      if pDocTabs(i) = 0 then continue for
      iTab = gTTabCtl.GetTabIndexByDocumentPtr(pDocTabs(i))
      if iTab = -1 then iTab = gTTabCtl.AddTab(pDocTabs(i))  ' Add the new document to the top tabcontrol
   Next   
   gTTabCtl.SetFocusTab(iTab)

   gApp.IsFileLoading = false
   LockWindowUpdate(0)
   gApp.wszPanelText = ""   ' reset filename text that displays in StatusBar panel
   frmMain_SetStatusbar

   ' Update the MRU list. Do this after the tabs are loaded because IsProjectLoading
   ' will prevent the list from updating.
   if gApp.IsProjectActive = false then
      For i = LBound(pDocTabs) To UBound(pDocTabs)
         if pDocTabs(i) then UpdateMRUList(pDocTabs(i)->DiskFilename) 
      Next     
   end if
   
   SetCursor( hCurSave )
   ResetScintillaCursors
   frmMain_PositionWindows
   frmMain_SetFocusToCurrentCodeWindow

   Function = 0

End Function


' ========================================================================================
' Save the incoming pDoc file to disk
' ========================================================================================
public Function OnCommand_FileSave( ByVal HWnd As HWnd, _
                                    byval pDoc as clsDocument ptr, _
                                    ByVal bSaveAs As BOOLEAN _
                                    ) As LRESULT
   If pDoc = 0 Then Return 0

   if pDoc->IsDesigner then
      ' Parse the code to ensure up to date list of functions so that compiling will not
      ' fail by generating a Declare without a body function.
      ParseDocument(pDoc)
      ' Regenerate code to ensure that any missing body functions are now present
      pDoc->bRegenerateCode = true
      SendMessage(hwnd, MSG_USER_GENERATECODE, 0, cast(LPARAM, pDoc))
   end if
   
   ' Do the actual saving to disk
   If pDoc->SaveFile(bSaveAs) Then
      ' Apply document properties to this file because the file extension may have
      ' changed. For example from Untitled to *.bas
      If gApp.IsProjectActive Then
         gApp.ProjectSetFileType( pDoc, pDoc->ProjectFileType )    
      End If
      pDoc->ApplyProperties
      ParseDocument( pDoc )
      AddFunctionsToExplorerTreeview( pDoc, true )
   EndIf   

   ' Ensure that the Tab displays the correct filename in case it was changed
   gTTabCtl.SetTabText(-1)
   frmMain_SetFocusToCurrentCodeWindow
   frmFunctionList_UpdateListBox()

   Function = 0
End Function


' ========================================================================================
' Save all open files to disk
' ========================================================================================
public Function OnCommand_FileSaveAll( ByVal HWnd As HWnd ) As LRESULT

   Dim As Long nCount = TabCtrl_GetItemCount(gTTabCtl.hWindow) 
   If nCount = 0 Then Exit Function
                
   Dim pDoc As clsDocument Ptr 
   Dim i    As Long               
   Dim TB   As TC_ITEM 
   
   For i = 0 To nCount - 1
      ' Get the document pointer and then save file to disk
      TB.Mask = TCIF_PARAM
      TabCtrl_GetItem( gTTabCtl.hWindow, i, @TB )
      pDoc = Cast(clsDocument Ptr, TB.lParam)

      if pDoc->IsDesigner then
         ' Parse the code to ensure up to date list of functions so that compiling will not
         ' fail by generating a Declare without a body function.
         ParseDocument(pDoc)
         ' Regenerate code to ensure that any missing body functions are now present
         pDoc->bRegenerateCode = true
         SendMessage(hwnd, MSG_USER_GENERATECODE, 0, cast(LPARAM, pDoc))
      end if

      If cbool(SciExec( pDoc->hWindow(0), SCI_GETMODIFY, 0, 0 )) or pDoc->UserModified Then
         ' If this is a new Untitled document that is also dirty then we
         ' need to set the focus to that tab because the SaveAs dialog will be shown.
         If pDoc->IsNewFlag Then
            gTTabCtl.SetFocusTab(i)
         End If
         pDoc->SaveFile()
         gTTabCtl.SetTabText(i)  ' make sure "*" updates
         ' Parse the code for sub/function listing and create list
         ParseDocument(pDoc)
         AddFunctionsToExplorerTreeview(pDoc, true)
      End If   
   Next
   frmMain_SetFocusToCurrentCodeWindow
   frmFunctionList_UpdateListBox()

   Function = 0
End Function


' ========================================================================================
' Close current (or all) open file(s)
' ========================================================================================
private Function OnCommand_FileClose( ByVal HWnd As HWnd, _
                                      ByVal veFileClose As eFileClose = EFC_CLOSECURRENT _
                                      ) As LRESULT

   Dim wText As WString * MAX_PATH  
   Dim pDoc As clsDocument Ptr
   Dim As Long r
   
   ' Build a list of files to check
   dim as long nCurSel = TabCtrl_GetCurSel(gTTabCtl.hWindow)
   dim as long nCount = TabCtrl_GetItemCount(gTTabCtl.hWindow) 

   if nCount = 0 then return true
   
   redim bCloseIndex(nCount - 1) as boolean

   for i as long = 0 to nCount - 1 
      SELECT CASE veFileClose
         case EFC_CLOSEALL:         bCloseIndex(i) = true
         case EFC_CLOSECURRENT:     bCloseIndex(i) = iif(i = nCurSel, true, false)
         case EFC_CLOSEALLOTHERS:   bCloseIndex(i) = iif(i <> nCurSel, true, false)
         case EFC_CLOSEALLFORWARD:  bCloseIndex(i) = iif(i > nCurSel, true, false) 
         case EFC_CLOSEALLBACKWARD: bCloseIndex(i) = iif(i < nCurSel, true, false)
      end select
   next
   
   ' Must do everything in reverse order because when Tabs are removed
   ' then the indexes would be out of sync if done in ascending order.
   for i as long = nCount - 1 to 0 step -1
      if bCloseIndex(i) = false then continue for

      pDoc = gTTabCtl.GetDocumentPtr(i)
      If pDoc = 0 Then Return true
   
      If cbool(SciExec( pDoc->hWindow(0), SCI_GETMODIFY, 0, 0 )) or pDoc->UserModified Then
         TabCtrl_SetCurSel(gTTabCtl.hWindow, i)
         gTTabCtl.SetFocusTab(i)
         wText = AfxStrPathname( "NAMEX", pDoc->Diskfilename )
         r = MessageBox( HWnd, L(76,"Save current changes?") & WStr(" ") & wText, @WStr(APPNAME), _
                           MB_YESNOCANCEL Or MB_ICONQUESTION)
         If r = IDCANCEL Then Exit Function
         If r = IDYES Then 
            r = pDoc->SaveFile()
            If r = False Then Exit Function   ' save was cancelled
         End If
      End If
      
      ' Hide the Scintilla editing windows
      for ii as long = lbound(pDoc->hWindow) to ubound(pDoc->hWindow)
         ShowWindow( pDoc->hWindow(ii), SW_HIDE)
      NEXT
       ' Hide visual designer window
      if pDoc->IsDesigner THEN 
         ShowWindow(pDoc->hWndDesigner, SW_HIDE)
         ShowWindow(pDoc->hDesignTabCtrl, SW_HIDE)
         ShowWindow(HWND_FRMVDTOOLBOX, SW_HIDE)
      END IF
     
      TabCtrl_DeleteItem(gTTabCtl.hWindow, i)
      
      ' Only remove this document from the global collection if it is not
      ' part of any active Project or if it is new/unsaved file. Files that are part of a project are
      ' all closed (tabs/nodes) automatically in one shot rather than individually.
      if (gApp.IsProjectActive = false) or (pDoc->IsNewFlag = true) then 
         ' Remove it from the Explorer
         Dim As Hwnd hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
         Treeview_DeleteItem( hTree, pDoc->hNodeExplorer)
      end if   
      if gApp.IsProjectActive = false then
         ' Remove all references from the gdb2 database. Only remove the reference
         ' if the file is not part of an active project, otherwise, we will lose
         ' references to function names, etc.
         gdb2.dbDelete( pDoc->DiskFilename, 0 )
         frmMain_RemoveFromComboFiles( pDoc )
         gApp.RemoveDocument( pDoc )
      end if   
   next
      
   ' Set the active tab to the closest tab to the one just removed
   nCurSel = TabCtrl_GetCurSel(gTTabCtl.hWindow)
   gTTabCtl.SetFocusTab(nCurSel)
      
   ' Update the TODO listview in the frmOutput form
   frmOutput_UpdateToDoListview
   
   frmMain_PositionWindows
   frmExplorer_PositionWindows
   
   Function = true
End Function


