'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016 Paul Squires, PlanetSquires Software
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
' Load the Recent Files treeview
' ========================================================================================
Function LoadRecentFilesTreeview( ByVal HWnd As HWnd ) As LRESULT

   Dim As CWSTR wzStr
   Dim As HTREEITEM hNode
   Dim As BOOLEAN fEmpty = True

   Dim As HWnd hTree = GetDlgItem(HWnd, IDC_FRMRECENT_TREEFILES)
   
   Treeview_DeleteAllItems(hTree)

   ' Display MRU files
   wzStr = UCase(AfxStrRemove( L(10, "&Recent Files"), "&"))
   hNode = FF_TreeView_InsertItem( hTree, TVI_ROOT, wzStr, 0, gidxImageOpened, gidxImageOpened )
   FF_TreeView_SetCheckState( hTree, hNode, True)
   For i As Long = 0 To 9
      ' If the file no longer exists then remove it from the MRU list
      gConfig.MRU(i) = ProcessFromCurdrive(gConfig.MRU(i))
      If AfxFileExists(gConfig.MRU(i)) Then
         FF_TreeView_InsertItem( hTree, hNode, gConfig.MRU(i), 0, gidxImageBlank, gidxImageBlank)
         FF_TreeView_SetCheckState( hTree, hNode, False)
         fEmpty = False
      End If
   Next
   If fEmpty Then
      wzStr = @L(11,"(Empty)")
      FF_TreeView_InsertItem( hTree, hNode, wzStr, 0, gidxImageBlank, gidxImageBlank)
   End If
   Treeview_Expand(hTree, hNode, TVE_EXPAND) 

   Function = 0
      
End Function

' ========================================================================================
' Load the Recent Projects treeview
' ========================================================================================
Function LoadRecentProjectsTreeview( ByVal HWnd As HWnd ) As LRESULT

   Dim As CWSTR wzStr
   Dim As HTREEITEM hNode
   Dim As BOOLEAN fEmpty = True

   Dim As HWnd hTree = GetDlgItem(HWnd, IDC_FRMRECENT_TREEPROJECTS)
   
   Treeview_DeleteAllItems(hTree)

   ' Display MRU files
   wzStr = UCase(AfxStrRemove( L(219, "&Recent Projects"), "&"))
   hNode = FF_TreeView_InsertItem( hTree, TVI_ROOT, wzStr, 0, gidxImageOpened, gidxImageOpened )
   FF_TreeView_SetCheckState( hTree, hNode, True)
   For i As Long = 0 To 9
      ' If the file no longer exists then remove it from the MRU list
      gConfig.MRUPROJECT(i) = ProcessFromCurdrive(gConfig.MRUPROJECT(i))
      If AfxFileExists(gConfig.MRUPROJECT(i)) Then
         FF_TreeView_InsertItem( hTree, hNode, gConfig.MRUPROJECT(i), 0, gidxImageBlank, gidxImageBlank)
         FF_TreeView_SetCheckState( hTree, hNode, False)
         fEmpty = False
      End If
   Next
   If fEmpty Then
      wzStr = @L(11,"(Empty)")
      FF_TreeView_InsertItem( hTree, hNode, wzStr, 0, gidxImageBlank, gidxImageBlank)
   End If
   Treeview_Expand(hTree, hNode, TVE_EXPAND) 
   
   Function = 0
      
End Function


' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
Function PositionRecentWindows( ByVal HWnd As HWnd ) As LRESULT
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
   If pWindow = 0 Then Exit Function
   
   ' Get the entire client area
   Dim As Rect rc
   GetClientRect(HWnd, @rc)

   Dim As HWnd hBtn
   Dim As Long nLeft, nTop, nWidth, nHeight
   
   nTop = pWindow->ScaleY(12)
   hBtn = GetDlgItem(HWND_FRMRECENT, IDC_FRMRECENT_BTNOPENFILE)
   SetWindowPos( hBtn, 0, pWindow->ScaleX(28), nTop, _
                     rc.right-pWindow->ScaleX(56), pWindow->ScaleY(28),  _
                     SWP_SHOWWINDOW Or SWP_NOZORDER )
   AfxRedrawWindow(hBtn)
   
   nTop = nTop + pWindow->ScaleY(40)
   hBtn = GetDlgItem(HWND_FRMRECENT, IDC_FRMRECENT_BTNOPENPROJECT)
   SetWindowPos( hBtn, 0, pWindow->ScaleX(28), nTop, _
                     rc.right-pWindow->ScaleX(56), pWindow->ScaleY(28),  _
                     SWP_SHOWWINDOW Or SWP_NOZORDER )
   AfxRedrawWindow(hBtn)

   Dim As Long NumItems = 1  ' start with the header 
   For i As Long = 0 To 9
      ' If the file no longer exists then remove it from the MRU list
      gConfig.MRU(i) = ProcessFromCurdrive(gConfig.MRU(i))
      If AfxFileExists(gConfig.MRU(i)) Then NumItems = NumItems + 1
   Next
   ' The MRU list will consist of at least the Header and the (Empty) designation.
   If NumItems < 2 Then NumItems = 2
   
   nTop = nTop + pWindow->ScaleY(48)
   hBtn = GetDlgItem(HWND_FRMRECENT, IDC_FRMRECENT_TREEFILES)
   SetWindowPos( hBtn, 0, pWindow->ScaleX(4), nTop, _
                  rc.Right-rc.Left-pWindow->ScaleX(8), (NumItems * pWindow->ScaleY(24)), _
                  SWP_SHOWWINDOW Or SWP_NOZORDER )
   
   nTop = nTop + AfxGetWindowHeight(hBtn) + pWindow->ScaleY(18)
   hBtn = GetDlgItem(HWND_FRMRECENT, IDC_FRMRECENT_TREEPROJECTS)
   SetWindowPos( hBtn, 0, pWindow->ScaleX(4), nTop, _
                  rc.Right-rc.Left-pWindow->ScaleX(8), pWindow->ScaleY(264), _
                  SWP_SHOWWINDOW Or SWP_NOZORDER )

   Function = 0
End Function
 

' ========================================================================================
' Process WM_SIZE message for window/dialog: frmMain
' ========================================================================================
Function frmRecent_OnSize(ByVal HWnd As HWnd, ByVal state As UINT, ByVal cx As Long, ByVal cy As Long) As LRESULT
   If state <> SIZE_MINIMIZED Then
      ' Position all of the child windows
      PositionRecentWindows(HWnd)
   End If
   Function = 0
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmRecent
' ========================================================================================
Function frmRecent_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   Select Case id
      Case IDC_FRMRECENT_BTNOPENFILE
         If codeNotify = BN_CLICKED Then
            OnCommand_FileOpen(HWND_FRMMAIN) 
         End If
      Case IDC_FRMRECENT_BTNOPENPROJECT
         If codeNotify = BN_CLICKED Then
            OnCommand_ProjectOpen(HWND_FRMMAIN) 
         End If
   End Select
      
   Function = 0
End Function


' ========================================================================================
' Process WM_NOTIFY message for window/dialog: frmRecent
' ========================================================================================
Function frmRecent_OnNotify(ByVal HWnd As HWnd, ByVal id As Long, ByVal pNMHDR As NMHDR Ptr) As LRESULT
 
   Select Case pNMHDR->code 
      
      Case NM_CUSTOMDRAW
         Dim tvcd As NMTVCUSTOMDRAW Ptr = Cast(NMTVCUSTOMDRAW Ptr, pNMHDR)

         Select Case tvcd->nmcd.dwDrawStage
            Case CDDS_PREPAINT
               Return CDRF_NOTIFYITEMDRAW 

            Case CDDS_ITEMPREPAINT 
               Dim As HFONT hFont = AfxGetWindowFont(GetDlgItem(hWnd, id))
               SelectFont(tvcd->nmcd.hdc, hFont)
               If (tvcd->nmcd.uItemState And ODS_SELECTED) And (GetFocus() = tvcd->nmcd.hdr.hwndFrom) Then
                  tvcd->clrTextBk = BGR(9,71,113)
                  tvcd->clrText   = BGR(255,255,250)
               ElseIf (tvcd->nmcd.uItemState And ODS_SELECTED) Then
                  tvcd->clrTextBk = BGR(63,63,70)
                  tvcd->clrText   = BGR(204,204,204)
               ElseIf (tvcd->nmcd.uItemState And CDIS_HOT) Then
                  tvcd->clrTextBk = BGR(42,45,46)
                  tvcd->clrText   = BGR(204,204,204)
               Else
                  tvcd->clrTextBk = BGR(80,80,80)
                  tvcd->clrText   = BGR(204,204,204)
               End If                                   
               Return CDRF_NEWFONT 

            Case CDDS_PREPAINT
               Return CDRF_SKIPDEFAULT
                                            
         End Select
      
   End Select
 
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmRecent
' ========================================================================================
Function frmRecent_OnDestroy(HWnd As HWnd) As LRESULT
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)

   Dim As HFONT hFont
   Dim As HWnd hTree
   hTree = GetDlgItem( hWnd, IDC_FRMRECENT_TREEFILES)
   hFont = AfxGetWindowFont(hTree):  DeleteFont(hFont)
   hTree = GetDlgItem( hWnd, IDC_FRMRECENT_TREEPROJECTS)
   hFont = AfxGetWindowFont(hTree):  DeleteFont(hFont)
   
   Function = 0
End Function


' ========================================================================================
' Processes messages for the subclassed frmRecent treeview controls.
' ========================================================================================
Function frmRecent_Tree_SubclassProc ( _
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
         If hItem Then   
            ' Get the text for the line that was double clicked on.
            Dim As CWSTR wzStr = String(MAX_PATH, 0)
            TreeView_GetItemText( hTree, hItem, wzStr, MAX_PATH )
            If AfxFileExists(wzStr) Then
               If uIdSubclass = IDC_FRMRECENT_TREEPROJECTS Then
                  OpenMRUProjectFile( HWND_FRMMAIN, 0, wzStr)
               Else   
                  frmMain_OpenFileSafely(HWND_FRMMAIN, _
                                          False, _    ' bIsNewFile
                                          False, _    ' bIsTemplate
                                          true, _     ' bShowInTab
                                          false, _    ' bIsInclude
                                          wzStr, _  ' pwszName
                                          0 )  ' pDocIn
               End If
            End If
            frmMain_PositionWindows(HWND_FRMMAIN)
            SetFocusScintilla
         End If   
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
         RemoveWindowSubclass HWnd, @frmRecent_Tree_SubclassProc, uIdSubclass

   End Select

   ' Default processing of Windows messages
   Function = DefSubclassProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmRecent Window procedure
' ========================================================================================
Function frmRecent_WndProc( ByVal HWnd   As HWnd, _
                            ByVal uMsg   As UINT, _
                            ByVal wParam As WPARAM, _
                            ByVal lParam As LPARAM _
                            ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_SIZE,    frmRecent_OnSize)
      HANDLE_MSG (HWnd, WM_COMMAND, frmRecent_OnCommand)
      HANDLE_MSG (HWnd, WM_NOTIFY,  frmRecent_OnNotify)
      HANDLE_MSG (HWnd, WM_DESTROY, frmRecent_OnDestroy)
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProcW(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmRecent_Show
' ========================================================================================
Function frmRecent_Show( ByVal hWndParent As HWnd ) As LRESULT

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   HWND_FRMRECENT = pWindow->Create( hWndParent, "", @frmRecent_WndProc, 0, 0, 250, 0, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   ' Disable background erasing by only assigning the one style
   pWindow->ClassStyle = CS_DBLCLKS
   pWindow->SetFont( "Segoe UI Semibold", 9)
   pWindow->Brush = CreateSolidBrush(BGR(80,80,80))

   Dim as HWND hBtn = _
      frmHotTxtBtn( HWND_FRMRECENT, IDC_FRMRECENT_BTNOPENFILE, L(248, "Open File"), "", _
                       BGR(255,255,250), BGR(9,71,113), _
                       BGR(255,255,250), BGR(63,63,70), _
                       1 )            ' 0 Left, 1 Center, 2 Right
      AfxSetWindowFont( hBtn, pWindow->Font, True )

   hBtn = _   
      frmHotTxtBtn( HWND_FRMRECENT, IDC_FRMRECENT_BTNOPENPROJECT, L(249, "Open Project"), "", _
                       BGR(255,255,250), BGR(9,71,113), _
                       BGR(255,255,250), BGR(63,63,70), _
                       1 )            ' 0 Left, 1 Center, 2 Right
      AfxSetWindowFont( hBtn, pWindow->Font, True )

   
   Dim As HFONT hFont
   Dim As HWnd hTree 
   
   hTree = _
        pWindow->AddControl("TREEVIEW", , IDC_FRMRECENT_TREEFILES, "", 0, 0, 0, 0, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or _
        TVS_NOSCROLL Or TVS_SHOWSELALWAYS Or TVS_FULLROWSELECT Or TVS_TRACKSELECT, _ 
        WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR, , _
        Cast(WNDPROC, @frmRecent_Tree_SubclassProc), IDC_FRMRECENT_TREEFILES, Cast(DWORD_PTR, @pWindow))
  
   hFont = pWindow->CreateFont("Segoe UI", 9) 
   SendMessageW( hTree, WM_SETFONT, Cast(WPARAM, hFont), CTRUE )
   'SendMessageW( hTree, TVM_SETBKCOLOR, 0, Cast(LPARAM, BGR(37,37,37)) )
   SendMessageW( hTree, TVM_SETBKCOLOR, 0, Cast(LPARAM, BGR(80,80,80)) )
   SendMessageW( hTree, TVM_SETITEMHEIGHT, pWindow->ScaleY(24), 0 )
   SendMessageW( hTree, TVM_SETEXTENDEDSTYLE, TVS_EX_DOUBLEBUFFER, TVS_EX_DOUBLEBUFFER)   
   'TreeView_SetImageList( hTree, ghImageListNormal, TVSIL_NORMAL )

   
   hTree = _
        pWindow->AddControl("TREEVIEW", , IDC_FRMRECENT_TREEPROJECTS, "", 0, 0, 0, 0, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or _
        TVS_NOSCROLL Or TVS_SHOWSELALWAYS Or TVS_FULLROWSELECT Or TVS_TRACKSELECT, _  
        WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR, , _
        Cast(WNDPROC, @frmRecent_Tree_SubclassProc), IDC_FRMRECENT_TREEPROJECTS, Cast(DWORD_PTR, @pWindow))
   
   hFont = pWindow->CreateFont("Segoe UI", 9) 
   SendMessageW( hTree, WM_SETFONT, Cast(WPARAM, hFont), CTRUE )
   'SendMessageW( hTree, TVM_SETBKCOLOR, 0, Cast(LPARAM, BGR(37,37,37)) )
   SendMessageW( hTree, TVM_SETBKCOLOR, 0, Cast(LPARAM, BGR(80,80,80)) )
   SendMessageW( hTree, TVM_SETITEMHEIGHT, pWindow->ScaleY(24), 0 )
   SendMessageW( hTree, TVM_SETEXTENDEDSTYLE, TVS_EX_DOUBLEBUFFER, TVS_EX_DOUBLEBUFFER)   
   'TreeView_SetImageList( hTree, ghImageListNormal, TVSIL_NORMAL )

   LoadRecentFilesTreeview(HWND_FRMRECENT)
   LoadRecentProjectsTreeview(HWND_FRMRECENT)
   
   PositionRecentWindows(HWND_FRMRECENT)

   Function = 0
   
End Function










