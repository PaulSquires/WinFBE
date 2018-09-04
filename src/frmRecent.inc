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
' Load the Recent Files/Projects Treeview
' ========================================================================================
Function LoadRecentFilesTreeview( ByVal HWnd As HWnd ) As LRESULT

   Dim As CWSTR wzStr 
   dim As HTREEITEM hNode, hFirstNode
   Dim As BOOLEAN fEmpty = True 

   Dim As HWnd hTree = GetDlgItem(HWnd, IDC_FRMRECENT_TREEVIEW) 

   Treeview_DeleteAllItems(hTree) 

   ''  RECENT FILES
   wzStr = UCase(L(10, "Recent Files")) 
   hNode = FF_TreeView_InsertItem( hTree, TVI_ROOT, wzStr )
   hFirstNode = hNode
   gApp.hRecentFilesRootNode = hNode
   Treeview_SetBold( hTree, hNode, TRUE) 
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


   ''  RECENT PROJECTS
   wzStr = UCase(L(219, "Recent Projects")) 
   hNode = FF_TreeView_InsertItem( hTree, TVI_ROOT, wzStr ) 
   gApp.hRecentProjectsRootNode = hNode
   Treeview_SetBold( hTree, hNode, TRUE) 
   For i As Long = 0 To 9 
      ' If the file no longer exists then remove it from the MRU list 
      gConfig.MRUPROJECT(i) = ProcessFromCurdrive(gConfig.MRUPROJECT(i)) 
      If AfxFileExists(gConfig.MRUPROJECT(i)) Then 
         FF_TreeView_InsertItem( hTree, hNode, gConfig.MRUPROJECT(i) )
         fEmpty = False 
      End If 
   Next 
   If fEmpty Then 
      wzStr = @L(11,"(Empty)") 
      FF_TreeView_InsertItem( hTree, hNode, wzStr )
   End If 
   Treeview_Expand(hTree, hNode, TVE_EXPAND)  
 
   ' Set the view to the first line
   TreeView_EnsureVisible(hTree, hFirstNode)
   
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

   nTop = nTop + pWindow->ScaleY(48)
   hBtn = GetDlgItem(HWND_FRMRECENT, IDC_FRMRECENT_TREEVIEW)
   SetWindowPos( hBtn, 0, pWindow->ScaleX(4), nTop, _
                  rc.Right-rc.Left-pWindow->ScaleX(8), rc.bottom - nTop, _
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
' Process WM_DESTROY message for window/dialog: frmRecent
' ========================================================================================
Function frmRecent_OnDestroy(HWnd As HWnd) As LRESULT
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)

   Dim As HFONT hFont
   Dim As HWnd hTree
   hTree = GetDlgItem( hWnd, IDC_FRMRECENT_TREEVIEW)
   
   Function = 0
End Function


' ========================================================================================
' Processes messages for the subclassed frmRecent Treeview controls.
' ========================================================================================
Function frmRecent_Tree_SubclassProc ( _
                  ByVal HWnd   As HWnd, _                 ' // Control window handle
                  ByVal uMsg   As UINT, _                 ' // Type of message
                  ByVal wParam As WPARAM, _               ' // First message parameter
                  ByVal lParam As LPARAM, _               ' // Second message parameter
                  ByVal uIdSubclass As UINT_PTR, _        ' // The subclass ID
                  ByVal dwRefData As DWORD_PTR _          ' // Pointer to reference data
                  ) As LRESULT

   ' Convert our ENTER key presses into LBUTTONDBLCLK to process them similarly
   If (uMsg = WM_KEYUP) And (Loword(wParam) = VK_RETURN) Then uMsg = WM_LBUTTONDBLCLK
      
   Select Case uMsg

      Case WM_GETDLGCODE
         ' All keyboard input
         Function = DLGC_WANTALLKEYS
         Exit Function

      Case WM_LBUTTONDBLCLK
         dim as hwnd hTree = hWnd
         dim As HTREEITEM hItem = Treeview_GetSelection(hTree)
         If hItem Then   
            ' Get the text for the line that was double clicked on.
            Dim As CWSTR wzStr = String(MAX_PATH, 0)
            TreeView_GetItemText( hTree, hItem, wzStr, MAX_PATH )
            If AfxFileExists(wzStr) Then
               If Treeview_GetParent(hTree, hItem) = gApp.hRecentProjectsRootNode then 
                  OpenMRUProjectFile( HWND_FRMMAIN, 0, wzStr)
               Elseif Treeview_GetParent(hTree, hItem) = gApp.hRecentFilesRootNode then 
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
      HANDLE_MSG (HWnd, WM_DESTROY, frmRecent_OnDestroy)
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmRecent_Show
' ========================================================================================
Function frmRecent_Show( ByVal hWndParent As HWnd ) As LRESULT

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   HWND_FRMRECENT = pWindow->Create( hWndParent, "Recent Window", @frmRecent_WndProc, 0, 0, 250, 0, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   ' Disable background erasing by only assigning the one style
   pWindow->ClassStyle = CS_DBLCLKS
   pWindow->SetFont( "Segoe UI Semibold", 9)
   pWindow->Brush = GetSysColorBrush(COLOR_BTNFACE)

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

   Dim As HWnd hTree = _
        pWindow->AddControl("TREEVIEW", , IDC_FRMRECENT_TREEVIEW, "", 0, 0, 0, 0, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or _
        TVS_SHOWSELALWAYS Or TVS_FULLROWSELECT Or TVS_TRACKSELECT, _   
        WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR, , _
        Cast(SUBCLASSPROC, @frmRecent_Tree_SubclassProc), IDC_FRMRECENT_TREEVIEW, Cast(DWORD_PTR, @pWindow))

   Dim As HFONT hFont = pWindow->CreateFont("Segoe UI", 9) 
   SendMessage( hTree, WM_SETFONT, Cast(WPARAM, hFont), CTRUE )
   SendMessage( hTree, TVM_SETITEMHEIGHT, pWindow->ScaleY(24), 0 )
   SendMessage( hTree, TVM_SETEXTENDEDSTYLE, TVS_EX_DOUBLEBUFFER, TVS_EX_DOUBLEBUFFER)   
   SendMessage( hTree, TVM_SETBKCOLOR, 0, Cast(LPARAM, GetSysColor(COLOR_BTNFACE)) )
   
   LoadRecentFilesTreeview(HWND_FRMRECENT)
   PositionRecentWindows(HWND_FRMRECENT)

   Function = 0
   
End Function



