'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2023 Paul Squires, PlanetSquires Software
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
 
#include once "frmRecent.bi"
#include once "clsConfig.bi"


' ========================================================================================
' Load the Recent Files/Projects Treeview
' ========================================================================================
public Function frmRecent_LoadTreeview() As LRESULT

   Dim As CWSTR wzStr 
   dim As HTREEITEM hNode, hFirstNode
   Dim As BOOLEAN fEmpty = True 

   Dim As HWnd hTree = GetDlgItem(HWND_FRMRECENT, IDC_FRMRECENT_TREEVIEW) 

   Treeview_DeleteAllItems(hTree) 

   ''  RECENT FILES
   wzStr = UCase(L(10, "Recent Files")) 
   hNode = TreeView_AppendItem( hTree, TVI_ROOT, wzStr )
   hFirstNode = hNode
   gApp.hRecentFilesRootNode = hNode
   Treeview_SetBold( hTree, hNode, TRUE) 
   For i As Long = 0 To 9 
      ' If the file no longer exists then remove it from the MRU list 
      gConfig.MRU(i) = ProcessFromCurdriveApp(gConfig.MRU(i)) 
      If AfxFileExists(gConfig.MRU(i)) Then 
         TreeView_AppendItem( hTree, hNode, gConfig.MRU(i) )
         fEmpty = False 
      End If 
   Next 
   If fEmpty Then 
      wzStr = @L(11,"(Empty)") 
      TreeView_AppendItem( hTree, hNode, wzStr, 0 ) 
   End If 
   Treeview_Expand(hTree, hNode, TVE_EXPAND)  


   ''  RECENT PROJECTS
   wzStr = UCase(L(219, "Recent Projects")) 
   hNode = TreeView_AppendItem( hTree, TVI_ROOT, wzStr ) 
   gApp.hRecentProjectsRootNode = hNode
   Treeview_SetBold( hTree, hNode, TRUE) 
   For i As Long = 0 To 9 
      ' If the file no longer exists then remove it from the MRU list 
      gConfig.MRUPROJECT(i) = ProcessFromCurdriveApp(gConfig.MRUPROJECT(i)) 
      If AfxFileExists(gConfig.MRUPROJECT(i)) Then 
         TreeView_AppendItem( hTree, hNode, gConfig.MRUPROJECT(i) )
         fEmpty = False 
      End If 
   Next 
   If fEmpty Then 
      wzStr = @L(11,"(Empty)") 
      TreeView_AppendItem( hTree, hNode, wzStr )
   End If 
   Treeview_Expand(hTree, hNode, TVE_EXPAND)  
 
   ' Set the view to the first line
   TreeView_EnsureVisible(hTree, hFirstNode)
   
   Function = 0
      
End Function


' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
public Function frmRecent_PositionWindows() As LRESULT
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWND_FRMRECENT)
   If pWindow = 0 Then Exit Function
   
   ' Get the entire client area
   Dim As Rect rc
   GetClientRect(HWND_FRMRECENT, @rc)

   Dim As HWnd hBtn
   Dim As Long nLeft, nTop, nWidth, nHeight
   
   nTop = pWindow->ScaleY(12)
   hBtn = GetDlgItem(HWND_FRMRECENT, IDC_FRMRECENT_BTNNEWFILE)
   SetWindowPos( hBtn, 0, pWindow->ScaleX(28), nTop, _
                     rc.right-pWindow->ScaleX(56), pWindow->ScaleY(28),  _
                     SWP_SHOWWINDOW Or SWP_NOZORDER )
   AfxRedrawWindow(hBtn)

   nTop = nTop + pWindow->ScaleY(40)
   hBtn = GetDlgItem(HWND_FRMRECENT, IDC_FRMRECENT_BTNOPENFILE)
   SetWindowPos( hBtn, 0, pWindow->ScaleX(28), nTop, _
                     rc.right-pWindow->ScaleX(56), pWindow->ScaleY(28),  _
                     SWP_SHOWWINDOW Or SWP_NOZORDER )
   AfxRedrawWindow(hBtn)
   
   nTop = nTop + pWindow->ScaleY(40)
   hBtn = GetDlgItem(HWND_FRMRECENT, IDC_FRMRECENT_BTNNEWPROJECT)
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
private Function frmRecent_OnSize( ByVal HWnd As HWnd, _
                                   ByVal state As UINT, _
                                   ByVal cx As Long, _
                                   ByVal cy As Long _
                                   ) As LRESULT
   If state <> SIZE_MINIMIZED Then
      ' Position all of the child windows
      frmRecent_PositionWindows
   End If
   Function = 0
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmRecent
' ========================================================================================
private Function frmRecent_OnCommand( ByVal HWnd As HWnd, _
                                      ByVal id As Long, _
                                      ByVal hwndCtl As HWnd, _
                                      ByVal codeNotify As UINT _
                                      ) As LRESULT
   Select Case id
      Case IDC_FRMRECENT_BTNNEWFILE
         If codeNotify = BN_CLICKED Then
            OnCommand_FileNew(HWND_FRMMAIN) 
         End If
      Case IDC_FRMRECENT_BTNOPENFILE
         If codeNotify = BN_CLICKED Then
            OnCommand_FileOpen(HWND_FRMMAIN) 
         End If
      Case IDC_FRMRECENT_BTNNEWPROJECT
         If codeNotify = BN_CLICKED Then
            OnCommand_ProjectNew(HWND_FRMMAIN) 
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
private Function frmRecent_OnDestroy( byval HWnd As HWnd ) As LRESULT
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)

   DIM pButton AS CXpButton ptr 
   pButton = AfxCXpButtonPtr(hwnd, IDC_FRMRECENT_BTNNEWFILE)
   if pButton then delete pButton
   pButton = AfxCXpButtonPtr(hwnd, IDC_FRMRECENT_BTNOPENFILE)
   if pButton then delete pButton
   pButton = AfxCXpButtonPtr(hwnd, IDC_FRMRECENT_BTNNEWPROJECT)
   if pButton then delete pButton
   pButton = AfxCXpButtonPtr(hwnd, IDC_FRMRECENT_BTNOPENPROJECT)
   if pButton then delete pButton
   
   Dim As HFONT hFont
   Dim As HWnd hTree
   hTree = GetDlgItem( hWnd, IDC_FRMRECENT_TREEVIEW)
   
   Function = 0
End Function


' ========================================================================================
' Processes messages for the subclassed frmRecent Treeview controls.
' ========================================================================================
private Function frmRecent_Tree_SubclassProc ( _
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
                  frmMain_OpenProjectSafely(HWND_FRMMAIN, wzStr) 
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
            frmMain_PositionWindows
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
private Function frmRecent_WndProc( ByVal HWnd   As HWnd, _
                                    ByVal uMsg   As UINT, _
                                    ByVal wParam As WPARAM, _
                                    ByVal lParam As LPARAM _
                                    ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_SIZE,    frmRecent_OnSize)
      HANDLE_MSG (HWnd, WM_COMMAND, frmRecent_OnCommand)
      HANDLE_MSG (HWnd, WM_DESTROY, frmRecent_OnDestroy)
   End Select

   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmRecent_Show
' ========================================================================================
public Function frmRecent_Show( ByVal hWndParent As HWnd ) As LRESULT

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

   DIM pButton AS CXpButton ptr 
   pButton = new CXpButton(pWindow, IDC_FRMRECENT_BTNNEWFILE, L(411, "New File"), 0, 0, 0, 0)
   pButton->SetFont(pWindow->CreateFont( "Segoe UI Semibold", 10))
   pButton->DisableTheming
   pButton->SetTextForeColor(BGR(255,255,250))
   pButton->SetButtonBkColor(BGR(9,71,113))
   pButton->SetButtonBkColorHot(BGR(63,63,70))

   pButton = new CXpButton(pWindow, IDC_FRMRECENT_BTNOPENFILE, L(248, "Open File"), 0, 0, 0, 0)
   pButton->SetFont(pWindow->CreateFont( "Segoe UI Semibold", 10))
   pButton->DisableTheming
   pButton->SetTextForeColor(BGR(255,255,250))
   pButton->SetButtonBkColor(BGR(9,71,113))
   pButton->SetButtonBkColorHot(BGR(63,63,70))
    
   pButton = new CXpButton(pWindow, IDC_FRMRECENT_BTNOPENPROJECT, L(249, "Open Project"), 0, 0, 0, 0)
   pButton->SetFont(pWindow->CreateFont( "Segoe UI Semibold", 10))
   pButton->DisableTheming
   pButton->SetTextForeColor(BGR(255,255,250))
   pButton->SetButtonBkColor(BGR(9,71,113))
   pButton->SetButtonBkColorHot(BGR(63,63,70))

   pButton = new CXpButton(pWindow, IDC_FRMRECENT_BTNNEWPROJECT, L(180, "New Project"), 0, 0, 0, 0)
   pButton->SetFont(pWindow->CreateFont( "Segoe UI Semibold", 10))
   pButton->DisableTheming
   pButton->SetTextForeColor(BGR(255,255,250))
   pButton->SetButtonBkColor(BGR(9,71,113))
   pButton->SetButtonBkColorHot(BGR(63,63,70))

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

   frmRecent_LoadTreeview
   frmRecent_PositionWindows

   Function = 0
   
End Function



