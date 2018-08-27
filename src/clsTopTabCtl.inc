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
Function clsTopTabCtl.GetTabIndexFromFilename( ByVal pwszName As WString Ptr ) As Long
   ' Get the index of tab that holds the incoming filename. This is used
   ' mostly for positioning the editor to the correct document when an
   ' error occurs during compiling. 
   ' Returns -1 if not found, otherwise zero based index of tab.
   
   If TabCtrl_GetItemCount(this.hWindow) = 0 Then Return -1 
   If pwszName = 0 Then Return -1
   
   Dim As Long i
   Dim pDoc As clsDocument Ptr 
   Dim As Long nCount = TabCtrl_GetItemCount(this.hWindow) 

   Dim TB As TC_ITEM 
   TB.Mask = TCIF_PARAM
   
   For i = 0 To nCount - 1
      TabCtrl_GetItem( this.hWindow, i, @TB )
      pDoc = Cast(clsDocument Ptr, TB.lParam)
      If pDoc Then 
         If Ucase(pDoc->DiskFilename) = Ucase(*pwszName) Then Return i
      End If   
   Next

   Function = -1  ' if not found
End Function


''
''
Function clsTopTabCtl.GetTabIndexByDocumentPtr( ByVal pDocIn As clsDocument Ptr ) As Long
   ' Returns the tav index that holds the incoming pDocIn ptr.
   If TabCtrl_GetItemCount(this.hWindow) = 0 Then Return -1 
   If pDocIn = 0 Then Return -1
   
   Dim As Long i
   Dim pDoc As clsDocument Ptr 
   Dim As Long nCount = TabCtrl_GetItemCount(this.hWindow) 

   Dim TB As TC_ITEM 
   TB.Mask = TCIF_PARAM
   
   For i = 0 To nCount - 1
      TabCtrl_GetItem( this.hWindow, i, @TB )
      pDoc = Cast(clsDocument Ptr, TB.lParam)
      If pDoc = pDocIn Then Return i
   Next

   Function = -1  ' if not found
End Function


''
''
Function clsTopTabCtl.SetTabIndexByDocumentPtr( ByVal pDocIn As clsDocument Ptr ) As Long
   ' Sets to tab that holds the incoming pDocIn ptr.
   If TabCtrl_GetItemCount(this.hWindow) = 0 Then Return -1 
   If pDocIn = 0 Then Return -1
   
   Dim As Long i
   Dim pDoc As clsDocument Ptr 
   Dim As Long nCount = TabCtrl_GetItemCount(this.hWindow) 

   Dim TB As TC_ITEM 
   TB.Mask = TCIF_PARAM
   
   For i = 0 To nCount - 1
      TabCtrl_GetItem( this.hWindow, i, @TB )
      pDoc = Cast(clsDocument Ptr, TB.lParam)
      If pDoc = pDocIn Then 
         this.SetFocusTab(i)
         Exit Function
      End If   
   Next

   Function = -1  ' if not found
End Function


''
''
Function clsTopTabCtl.SetFocusTab( ByVal idx As Long ) As Long
   ' Set the tab with focus (this sends the TCN_SELCHANGING and TCN_SELCHANGE 
   ' notification codes to its parent window). 
   Dim As Long nCount = TabCtrl_GetItemCount(this.hWindow) 
   If idx < 0 Then idx = nCount - 1
   If TabCtrl_GetCurSel(this.hWindow) <> idx Then
      Function = TabCtrl_SetCurFocus(this.hWindow, idx)
      if IsWindowVisible(HWND_FRMFINDREPLACE) THEN
         frmFindReplace_HighlightSearches
         frmFindReplace_SetButtonStates
      END IF
      SetFocusScintilla
   End If
End Function


''
''
Function clsTopTabCtl.NextTab() As Long
   ' Invoked by Ctl+TAB
   ' Set the tab with focus (this sends the TCN_SELCHANGING and TCN_SELCHANGE 
   ' notification codes to its parent window). 
   Dim As Long nCount = TabCtrl_GetItemCount(this.hWindow) 
   Dim As Long idx = TabCtrl_GetCurSel(this.hWindow)
   If nCount = 0 Then Exit Function
   idx += 1
   If idx > nCount - 1 Then idx = 0
   Function = TabCtrl_SetCurFocus(this.hWindow, idx)
End Function

''
''
Function clsTopTabCtl.PrevTab() As Long
   ' Invoked by Ctl+Shift+TAB
   ' Set the tab with focus (this sends the TCN_SELCHANGING and TCN_SELCHANGE 
   ' notification codes to its parent window). 
   Dim As Long nCount = TabCtrl_GetItemCount(this.hWindow) 
   Dim As Long idx = TabCtrl_GetCurSel(this.hWindow)
   If nCount = 0 Then Exit Function
   idx -= 1
   If idx < 0 Then idx = nCount - 1 
   Function = TabCtrl_SetCurFocus(this.hWindow, idx)
End Function

''
''
Function clsTopTabCtl.CloseTab() As Long
   ' Invoked by Ctl+W or by clicking "X" on tab
   OnCommand_FileClose( HWND_FRMMAIN, EFC_CLOSECURRENT )
   Function = 0
End Function


''
''
Function clsTopTabCtl.GetActiveDocumentPtr() As clsDocument Ptr
   ' Return pointer to clsDocument class for the current active tab
   If TabCtrl_GetItemCount(this.hWindow) = 0 Then Return 0 
      
   Dim TB As TC_ITEM 
   TB.Mask = TCIF_PARAM
   
   Dim As Long idx = TabCtrl_GetCurSel(this.hWindow)
   TabCtrl_GetItem( this.hWindow, idx, @TB )
   
   Function = Cast(clsDocument Ptr, TB.lParam)
End Function

''
''
Function clsTopTabCtl.GetDocumentPtr( ByVal idx As Long ) As clsDocument Ptr
   ' Return pointer to clsDocument class for the current active tab
   Dim TB As TC_ITEM 
   TB.Mask = TCIF_PARAM
   
   TabCtrl_GetItem( this.hWindow, idx, @TB )
   
   Function = Cast(clsDocument Ptr, TB.lParam)
End Function

''
''
Function clsTopTabCtl.DisplayScintilla( ByVal idx As Long, ByVal bShow As BOOLEAN ) As Long
   ' Show/Hide the Scintilla editing window (or visual designer window) for the incoming tab index
   Dim TB As TC_ITEM 
   TB.Mask = TCIF_PARAM
   
   TabCtrl_GetItem( this.hWindow, idx, @TB )
   
   Dim pDocShow As clsDocument Ptr = Cast(clsDocument Ptr, TB.lParam)
   If pDocShow = 0 Then exit function
   
   ' Hide all documents and their associated scrollbars
   Dim pDoc As clsDocument Ptr 
   dim as long idxProject = gApp.GetActiveProjectIndex()
   
   for i as long = 0 to gApp.Projects(idxProject).GetDocumentCount - 1
      pDoc = gApp.Projects(idxProject).GetDocumentPtr(i)
      if pDoc = 0 THEN continue for
      if pDoc->ProjectIndex <> idxProject THEN continue for
      ' Scintilla windows and scrollbars
      ShowWindow(pDoc->hWindow(0), SW_HIDE)
      ShowWindow(pDoc->hWindow(1), SW_HIDE)
      ShowWindow(pDoc->hScrollbar, SW_HIDE)
      if pDoc->IsDesigner THEN 
         ShowWindow(pDoc->hWndDesigner, SW_HIDE)
         ShowWindow(pDoc->hDesignTabCtrl, SW_HIDE)
      END IF
   next
   
   if (pDocShow->IsDesigner) andalso (IsDesignerView(pDocShow)) THEN
      ShowWindow(pDocShow->hWndDesigner, Iif(bShow, SW_SHOWNORMAL, SW_HIDE))
      ShowWindow(pDocShow->hDesignTabCtrl, Iif(bShow, SW_SHOWNORMAL, SW_HIDE))
      ShowWindow(HWND_FRMVDTOOLBOX, Iif(bShow, SW_SHOWNORMAL, SW_HIDE))
   else
      ' Show/Hide our current active Scintilla window and scrollbar
      ShowWindow(pDocShow->hWindow(0), Iif(bShow, SW_SHOWNORMAL, SW_HIDE))
      ShowWindow(pDocShow->hWindow(1), Iif(bShow, SW_SHOWNORMAL, SW_HIDE))
      ShowWindow(pDocShow->hScrollbar, Iif(bShow, SW_SHOWNORMAL, SW_HIDE))
   end if
   
   Function = 0
End Function
   
''
''
Function clsTopTabCtl.AddTab( ByVal pDoc As clsDocument Ptr ) As Long
   Dim idx As Long
   Dim TB  As TC_ITEM 
   
   TB.Mask   = TCIF_PARAM Or TCIF_IMAGE
   TB.lParam = Cast(LPARAM, pDoc)
   TB.iImage = 0
   
   Dim As Long NumItems = TabCtrl_GetItemCount(this.hWindow)             
   
   ' Return the index of the added tab
   idx = SendMessage( this.hWindow, TCM_INSERTITEM, NumItems, Cast(LPARAM, @TB) )     
   
   this.SetTabText(idx)
   
   Function = idx
End Function

''
''
Function clsTopTabCtl.SetTabText( ByVal idx As Long ) As Long
   Dim TB    As TC_ITEM 
   Dim wText As WString * MAX_PATH  
   Dim wTemp As WString * MAX_PATH  

   If idx = -1 Then idx = TabCtrl_GetCurSel(this.hWindow)

   ' Get the pointer to the clsDocument stored in lParam and also retrieve
   ' the existing text. 
   TB.Mask       = TCIF_PARAM Or TCIF_TEXT
   TB.pszText    = @wTemp
   TB.cchTextMax = MAX_PATH
   TabCtrl_GetItem( this.hWindow, idx, @TB )
   
   Dim pDoc As clsDocument Ptr = Cast(clsDocument Ptr, TB.lParam)
   If pDoc Then 
      ' Set the text that displays on the tab
      wText = AfxStrPathname( "NAMEX", pDoc->DiskFilename )
      If cbool(SciExec( pDoc->hWindow(0), SCI_GETMODIFY, 0, 0 )) or pDoc->UserModified Then
         wText = wText & "*"
      End If   
            
      ' We only update the text if it has changed in order to prevent flicker.
      If wText <> wTemp Then
         TB.Mask    = TCIF_TEXT 
         TB.pszText = @wText  
         Function = TabCtrl_SetItem( this.hWindow, idx, @TB )
         ' Also update the Explorer node's filename
         if pDoc->hNodeExplorer then
            Dim As Hwnd hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
            ' We don't want any * modififer to show
            wText = rtrim(wText, "*")
            TreeView_SetItemText( hTree, pDoc->hNodeExplorer, @wText )
         end if   
      End If
   End If
   
End Function

