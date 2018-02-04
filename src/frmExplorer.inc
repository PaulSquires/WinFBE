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
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
Function PositionExplorerWindows( ByVal HWnd As HWnd ) As LRESULT
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
   If pWindow = 0 Then Exit Function
   
   ' Get the entire client area
   Dim As Rect rc
   GetClientRect(HWnd, @rc)

   Dim As HWnd hCtl, hTree
   Dim As Long nLeft, nTop, nWidth, nHeight
   
   ' The Explorer label was painted in the top area so start our treeviews immediately after that
   nWidth  = pWindow->ScaleX(24)
   nHeight = pWindow->ScaleY(24)
   nTop    = pWindow->ScaleY(8)
   
   hCtl = GetDlgItem(HWND, IDC_FRMEXPLORER_BTNCLOSE)
   SetWindowPos( hCtl, 0, rc.Right - (nWidth*1.2), nTop, nWidth, nHeight, SWP_SHOWWINDOW Or SWP_NOZORDER )

   nTop = nTop + nHeight + pWindow->ScaleY(12)

   hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
   
   If gApp.IsProjectActive Then
      ' Project is active to continue to show the Explorer tree regardless of whether 
      ' any files are actually open.
      ShowWindow HWND_FRMRECENT, SW_HIDE
      ShowWindow hTree, SW_SHOW
   else
      ' If no project open then show Recent dialog
      if gApp.Projects(0).GetDocumentCount = 0 then
         ' Ensure that the Recent lists are the most current
         If IsWindowVisible(HWND_FRMRECENT) = 0 Then LoadRecentFilesTreeview(HWND_FRMRECENT)
         ShowWindow hTree, SW_HIDE
         ShowWindow HWND_FRMRECENT, SW_SHOW
      else
         ShowWindow HWND_FRMRECENT, SW_HIDE
         ShowWindow hTree, SW_SHOW
      end if
   End If   

   ' Allow a small space at the right edge to ensure area for splitter
   SetWindowPos( IIF(IsWindowVisible(hTree), hTree, HWND_FRMRECENT), _
                 0, 0, nTop, rc.Right - pWindow->ScaleX(2), rc.bottom-nTop, SWP_SHOWWINDOW Or SWP_NOZORDER )

   AfxRedrawWindow(HWND_FRMEXPLORER)
   
   Function = 0
End Function
 

' ========================================================================================
' Add the names of all subs/functions belonging to this document to the Explorer treeview
' ========================================================================================
Function AddFunctionsToExplorerTreeview( ByVal pDoc As clsDocument Ptr, ByVal fUpdateNodes As BOOLEAN ) As Long
 
   If pDoc = 0 Then Exit Function
   
   Dim As CWSTR wzFile = AfxStrPathname("NAMEX", pDoc->DiskFilename)
   Dim As Hwnd hTree   = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
   Dim As HTREEITEM hChild, hRootNode
   Dim As BOOLEAN fAddChildren = False
   dim as long idx
   
   ' Determine what root node to add or update the files from
   idx = gApp.GetActiveProjectIndex()
   hRootNode = gApp.Projects(idx).hExplorerRootNode
   
   ' If hNodeExplorer does not exist then we must be adding a new document. If it does
   ' exist then we simply need to update the child nodes.
   If fUpdateNodes = False Then
      If pDoc->hNodeExplorer = 0 Then
         pDoc->hNodeExplorer = FF_TreeView_InsertItem (hTree, hRootNode, wzFile, cast(LPARAM, pDoc) )
         fAddChildren = True
         Treeview_RemoveCheckBox(hTree, pDoc->hNodeExplorer)
      End If
   Else
      ' Delete any existing children sub/function nodes and then update with new pointer values
      hChild = TreeView_GetChild( hTree, pDoc->hNodeExplorer )
      Do Until hChild = 0
         TreeView_DeleteItem( hTree, hChild )
         hChild = TreeView_GetChild( hTree, pDoc->hNodeExplorer )
      Loop
      fAddChildren = True
   End If
 
   ' Need to ensure the Explorer treeview is sorted because the filename may have changed.
   Treeview_SortChildren( hTree, hRootNode, 0 )
      
   If fAddChildren Then
      dim wszText as wstring * MAX_PATH
      dim as HTREEITEM hNode
      dim pData as DB2_DATA ptr
      dim as string sFilename = ucase(pDoc->DiskFilename)
      
      gdb2.dbRewind()
      do 
         pData = gdb2.dbGetNext
         if pData = 0 THEN exit do
         if pData->id <> DB2_FUNCTION THEN continue do
         if sFilename = ucase(pData->fileName) THEN
            hNode = FF_TreeView_InsertItem(hTree, pDoc->hNodeExplorer, pData->ElementName, 0 )
            Treeview_RemoveCheckBox(hTree, hNode)
         END IF
      loop      
      Treeview_SortChildren(hTree, pDoc->hNodeExplorer, 0)
   End If
   
   Function = 0
End Function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmExplorer
' ========================================================================================
Function frmExplorer_OnSize(ByVal HWnd As HWnd, ByVal state As UINT, ByVal cx As Long, ByVal cy As Long) As LRESULT
   If state <> SIZE_MINIMIZED Then
      ' Position all of the child windows
      PositionExplorerWindows(HWnd)
   End If
   Function = 0
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmExplorer
' ========================================================================================
Function frmExplorer_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   Select case codeNotify
      case BN_CLICKED
         If id = IDC_FRMEXPLORER_BTNCLOSE Then
            ShowWindow HWND_FRMEXPLORER, SW_HIDE
            frmMain_PositionWindows(HWND_FRMMAIN)
         end if
   end select
                       
   Function = 0
End Function


' ========================================================================================
' Process WM_NOTIFY message for window/dialog: frmExplorer
' ========================================================================================
Function frmExplorer_OnNotify(ByVal HWnd As HWnd, ByVal id As Long, ByVal pNMHDR As NMHDR Ptr) As LRESULT
 
   If (id <> IDC_FRMEXPLORER_TREE) Then Return 0
   
   Dim lpNMTV As NM_TREEVIEW Ptr = Cast(NM_TREEVIEW Ptr, pNMHDR)
   Dim As HTREEITEM hItem
   Dim As HWND hTree 
   Dim As LPARAM lParam
   Dim pDoc As clsDocument Ptr
   Dim ht As TVHITTESTINFO
   Dim As POINT pt
   
   ' Determine what node was clicked on
   Select Case pNMHDR->code 
      
      case TVN_GETINFOTIP
         ' Display a tooltip when a treeview node is hovered over. This will normally
         ' be the full filename for the project or document.
         dim pnmTip as NMTVGETINFOTIPW ptr = cast(NMTVGETINFOTIPW ptr, pNMHDR)
         ' First, search if this tooltip belongs to a project root node
         for i as long = 1 to ubound(gApp.Projects)
            if gApp.Projects(i).InUse = true THEN
               if gApp.Projects(i).hExplorerRootNode = pnmTip->hItem then
                  pnmTip->pszText = cast(WSTRING Ptr, @**gApp.Projects(i).ProjectFilename)
                  pnmTip->cchTextMax = len(gApp.Projects(i).ProjectFilename)
                  exit function
               end if   
            END IF
         NEXT
         ' Second, search for a matching document
         pDoc = Cast(clsDocument ptr, pnmTip->lParam)
         if pDoc THEN
            pnmTip->pszText = @pDoc->DiskFilename
            pnmTip->cchTextMax = len(pDoc->DiskFilename)
         END IF
         exit function   

      Case NM_CLICK, NM_RCLICK
      ' If we have clicked on a folder item then issue an expand. Must use HitTest because the 
      ' the current node has not been selected yet.
         hTree = lpNMTV->hdr.hWndFrom
         GetCursorPos(@pt)  
         ScreenToClient(hTree, @pt)
         
         ht.pt = pt 
         Treeview_HitTest(hTree, @ht )
         hItem = ht.hItem
         If hItem Then Treeview_Select( hTree, hItem, TVGN_CARET)
         
         ' Was the checkbox clicked on. If yes then post user messge to select/deselect
         ' the other checkboxes for the other projects (only one can be active).
         if (ht.flags and TVHT_ONITEMSTATEICON) THEN
            PostMessage( HWND, MSG_USER_TOGGLE_TVCHECKBOXES, 0, Cast(LPARAM, ht.hItem))
         END IF

         If pNMHDR->code = NM_CLICK Then
            ' Ensure that the correct Project name is showing
            frmMain_PositionWindows(HWND_FRMMAIN)
         end if
         
         If pNMHDR->code = NM_RCLICK Then
            ' Create the popup menu
            lParam = FF_TreeView_GetlParam(hTree, hItem)
            pDoc = Cast(clsDocument ptr, lParam)
            ' Select the node and show the document
            if pDoc then SendMessage(hTree, WM_LBUTTONDBLCLK, 0, 0)
            Dim As HMENU hPopupMenu = CreateExplorerContextMenu(pDoc)
            GetCursorPos @pt
            TrackPopupMenu(hPopUpMenu, 0, pt.x, pt.y, 0, HWND_FRMMAIN, ByVal Null)
            DestroyMenu hPopUpMenu
            Function = True
         End If

         
'      Case NM_CUSTOMDRAW
'         Dim tvcd As NMTVCUSTOMDRAW Ptr = Cast(NMTVCUSTOMDRAW Ptr, pNMHDR)

'         Select Case tvcd->nmcd.dwDrawStage
'            Case CDDS_PREPAINT
'               Return CDRF_NOTIFYITEMDRAW 

'            Case CDDS_ITEMPREPAINT 
'               Dim As HFONT hFont = AfxGetWindowFont(GetDlgItem(hWnd, id))
'               SelectFont(tvcd->nmcd.hdc, hFont)
''               If (tvcd->nmcd.uItemState And ODS_SELECTED) And (GetFocus() = tvcd->nmcd.hdr.hwndFrom) Then
''                  tvcd->clrTextBk = gTheme.exlistbox_selected_active_bg
''                  tvcd->clrText   = gTheme.exlistbox_selected_active_fg
''               ElseIf (tvcd->nmcd.uItemState And ODS_SELECTED) Then
''                  tvcd->clrTextBk = gTheme.exlistbox_selected_inactive_bg
''                  tvcd->clrText   = gTheme.exlistbox_selected_inactive_fg
''               ElseIf (tvcd->nmcd.uItemState And CDIS_HOT) Then
''                  tvcd->clrTextBk = gTheme.exlistbox_hot_bg
''                  tvcd->clrText   = gTheme.exlistbox_hot_fg
''               Else
'                  tvcd->clrText   = gConfig.Themes(gConfig.idxTheme).colors(CLR_PANELTEXT).nFg
'                  tvcd->clrTextBk = gConfig.Themes(gConfig.idxTheme).colors(CLR_PANELTEXT).nBg
''               End If
                                    
'               Return CDRF_NEWFONT 
'
'         End Select

   End Select
 
   Function = 0
End Function


' ========================================================================================
' Processes messages for the subclassed frmExplorer treeview control.
' ========================================================================================
Function frmExplorer_Tree_SubclassProc ( _
                  ByVal HWnd   As HWnd, _                 ' // Control window handle
                  ByVal uMsg   As UINT, _                 ' // Type of message
                  ByVal wParam As WPARAM, _               ' // First message parameter
                  ByVal lParam As LPARAM, _               ' // Second message parameter
                  ByVal uIdSubclass As UINT_PTR, _        ' // The subclass ID
                  ByVal dwRefData As DWORD_PTR _          ' // Pointer to reference data
                  ) As LRESULT

   Dim As HWND hTree 
   Dim As POINT pt
   Dim As HTREEITEM hItem
   Dim As BOOLEAN bIsFolder
   Dim As BOOLEAN bIsExpanded

   ' Convert our ENTER key presses into LBUTTONDBLCLK to process them similarly
   If (uMsg = WM_KEYUP) And (Loword(wParam) = VK_RETURN) Then uMsg = WM_LBUTTONDBLCLK

      
   Select Case uMsg

      Case WM_GETDLGCODE
         ' All keyboard input
         Function = DLGC_WANTALLKEYS
         Exit Function

      Case WM_LBUTTONDBLCLK
         hTree = hWnd
         hItem = Treeview_GetSelection(hTree)
         lParam = FF_TreeView_GetlParam(hTree, hItem)

         dim wszText as wstring * MAX_PATH
         Dim pDoc As clsDocument Ptr
         pDoc = Cast(clsDocument Ptr, lParam)
         TreeView_GetItemText(hTree, hItem, @wszText, MAX_PATH)
         OpenSelectedDocument(pDoc, wszText)
         Exit Function
         
      Case WM_KEYUP
         Select Case Loword(wParam)
            Case VK_RETURN  ' already processed in WM_LBUTTONDBLCLK 
         End Select
         Exit Function

      Case WM_CHAR   ' prevent the annoying beep!
         If wParam = VK_RETURN Then Return 0

      Case WM_DESTROY
         ' REQUIRED: Remove control subclassing
         RemoveWindowSubclass HWnd, @frmExplorer_Tree_SubclassProc, uIdSubclass

   End Select

   ' Default processing of Windows messages
   Function = DefSubclassProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' Process WM_PAINT message for window/dialog: frmExplorer
' ========================================================================================
Function frmExplorer_OnPaint( ByVal HWnd As HWnd) As LRESULT
            
   Dim As PAINTSTRUCT ps
   Dim As HPEN hPen
   Dim As HDC hDc
   Dim As Rect rc
   
   Dim pWindow As CWindow Ptr = AfxCWindowOwnerPtr(HWND)
   if pWindow = 0 THEN exit function

   hDC = BeginPaint(hWnd, @ps)
   GetClientRect HWnd, @rc

   SaveDC hDC

   FillRect( hDC, @rc, GetSysColorBrush( COLOR_BTNFACE) )
   
   ' Outline the right edge of dialog with a dull gray pen
   hPen = CreatePen( PS_SOLID, 1, GetSysColor(COLOR_ACTIVEBORDER) )
   SelectPen( hDC, hPen )
   MoveToEx( hDC, rc.right - 1, rc.top, null)
   LineTo( hDC, rc.right - 1, rc.bottom-1 )
   RestoreDC hDC, -1 
   DeletePen(hPen)
   
   EndPaint hWnd, @ps
   
   Function = 0
End Function


' ========================================================================================
' frmExplorer Window procedure
' ========================================================================================
Function frmExplorer_WndProc( ByVal HWnd   As HWnd, _
                              ByVal uMsg   As UINT, _
                              ByVal wParam As WPARAM, _
                              ByVal lParam As LPARAM _
                              ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_SIZE,    frmExplorer_OnSize)
      HANDLE_MSG (HWnd, WM_COMMAND, frmExplorer_OnCommand)
      HANDLE_MSG (HWnd, WM_NOTIFY,  frmExplorer_OnNotify)
      HANDLE_MSG (HWnd, WM_PAINT,   frmExplorer_OnPaint)
   
      case WM_DESTROY
         Dim As HWnd hTree = GetDlgItem( hWnd, IDC_FRMEXPLORER_TREE)
         dim as HFONT hFont = AfxGetWindowFont(hTree)
         DeleteFont(hFont)
         
      case WM_ERASEBKGND
         return true

      case WM_CTLCOLORSTATIC
         ' Background color for the Explorer label
         SetBkColor Cast(HDC,wParam), GetSysColor(COLOR_BTNFACE)
         return Cast( LRESULT, GetSysColorBrush(COLOR_BTNFACE))

      case MSG_USER_TOGGLE_TVCHECKBOXES      
         ' If this is a project root node then ensure that only one of the projects
         ' root node has the checkmark box ticked. If there is only one project
         ' then do not allow it to be unchecked.
         dim as hwnd hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
         dim as HTREEITEM hNode = Cast(HTREEITEM, lParam)
         if gApp.GetProjectCount = 1 THEN
            gApp.SetActiveProject(hNode)
         else         
            if TreeView_IsItemChecked(hTree, hNode) then
               gApp.SetActiveProject(hNode)
            else
               gApp.EnsureDefaultActiveProject(hNode)
            end if   
         END IF

   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmExplorer_Show
' ========================================================================================
Function frmExplorer_Show( ByVal hWndParent As HWnd ) As LRESULT

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   ' Only make the Explorer initially visable if it was already visible
   ' when the most previous instance of the program closed. Also, set the width of
   ' the window to the last used visible width.
   ' Because pWindow->Create scales the control according to the DPI, you must 
   ' unscale the saved width:   
   dim as long nWidth = iif(gConfig.ShowExplorer, pWindow->UnscaleX(gConfig.ShowExplorerWidth), 250)
   
   HWND_FRMEXPLORER = pWindow->Create( hWndParent, "Explorer Window", @frmExplorer_WndProc, 0, 0, nWidth, 0, _
        WS_CHILD Or iif(gConfig.ShowExplorer, WS_VISIBLE, 0) Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   ' Disable background erasing by only assigning the one style
   pWindow->ClassStyle = CS_DBLCLKS
   pWindow->SetFont( "Segoe UI Semibold", 9)
   'pWindow->Brush = GetSysColorBrush(COLOR_BTNFACE)

   Dim As HWnd hBtn = _
   pWindow->AddControl("LABEL", , IDC_FRMEXPLORER_LBLEXPLORER, L(247,"EXPLORER"), 10, 12, 75, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   frmHotImgBtn( HWND_FRMEXPLORER, IDC_FRMEXPLORER_BTNCLOSE, "IMAGE_CLOSE", "IMAGE_CLOSE",  _
                  L(161, "Close"), 24, 24, _
                  GetSysColor(COLOR_BTNFACE), _
                  GetSysColor(COLOR_WINDOW), _
                  0, 0, 24, 24, true, true, true )

   Dim As HWnd hTree = _
        pWindow->AddControl("TREEVIEW", , IDC_FRMEXPLORER_TREE, "", 0, 0, 0, 0, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP or _
        TVS_INFOTIP or TVS_SHOWSELALWAYS Or TVS_FULLROWSELECT Or TVS_TRACKSELECT or _
        TVS_HASBUTTONS or TVS_HASLINES or TVS_LINESATROOT, _   
        WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR, , _
        Cast(WNDPROC, @frmExplorer_Tree_SubclassProc), IDC_FRMEXPLORER_TREE, Cast(DWORD_PTR, @pWindow))

      ' items that are initially checked will appear checked (must add TVS_CHECKBOXES after control is created)
      SetWindowLongPtr( hTree, GWL_STYLE, GetWindowLongPtr(hTree, GWL_STYLE) OR TVS_CHECKBOXES )
  
   Dim As HFONT hFont = pWindow->CreateFont("Segoe UI", 9) 
   SendMessage( hTree, WM_SETFONT, Cast(WPARAM, hFont), CTRUE )
   SendMessage( hTree, TVM_SETBKCOLOR, 0, Cast(LPARAM, GetSysColor(COLOR_BTNFACE)) )
'   SendMessage( hTree, TVM_SETITEMHEIGHT, pWindow->ScaleY(20), 0 )
   SendMessage( hTree, TVM_SETEXTENDEDSTYLE, TVS_EX_DOUBLEBUFFER, TVS_EX_DOUBLEBUFFER)   

   PositionExplorerWindows(HWND_FRMEXPLORER)

   Function = 0
   
End Function

