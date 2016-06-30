



' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmProjectManager
' ========================================================================================
Function frmProjectManager_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   Dim As HWnd hLV = GetDlgItem(HWnd, IDC_FRMPROJECTMANAGER_LISTVIEW )
   Dim As Long nCurSel = ListView_GetSelectionMark(hLV)
   Dim pDoc As clsDocument Ptr 

   If nCurSel > -1 Then
      pDoc = Cast( clsDocument Ptr, FF_ListView_GetlParam(hLV, nCurSel) )
      gpApp->IncludeFilename = *pDoc->DiskFilename
   End If
   
   Select Case id
      Case IDM_OPENINCLUDE:           Function = OnCommand_OpenIncludeFile(HWND_FRMMAIN)
      Case IDM_REMOVEFILEFROMPROJECT: Function = gpApp->ProjectRemoveFile( pDoc )
      Case IDM_SETFILENORMAL:         Function = gpApp->ProjectSetFileType( pDoc, FILETYPE_NORMAL )
      Case IDM_SETFILEMODULE:         Function = gpApp->ProjectSetFileType( pDoc, FILETYPE_MODULE )
      Case IDM_SETFILEMAIN:           Function = gpApp->ProjectSetFileType( pDoc, FILETYPE_MAIN )
   End Select

   Function = 0
End Function


' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
Function frmProjectManager_PositionWindows( ByVal HWnd As HWnd ) As LRESULT
   
   ' Get the entire client area
   Dim As Rect rc
   GetClientRect(HWnd, @rc)
   
   Dim As HWnd hLV = GetDlgItem(HWnd, IDC_FRMPROJECTMANAGER_LISTVIEW )
   SetWindowPos hLV, 0, rc.Left, rc.Top, rc.Right, rc.Bottom, SWP_SHOWWINDOW Or SWP_NOZORDER

   Function = 0
End Function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmProjectManager
' ========================================================================================
Function frmProjectManager_OnSize(ByVal HWnd As HWnd, ByVal state As UINT, ByVal cx As Long, ByVal cy As Long) As LRESULT

   If state <> SIZE_MINIMIZED Then 
      frmProjectManager_PositionWindows(HWnd)
   End If   

   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmProjectManager
' ========================================================================================
Function frmProjectManager_OnClose(HWnd As HWnd) As LRESULT
   ' Never close the window; simply hide it.
   ShowWindow( HWnd, SW_HIDE )
   Function = 0
End Function


' ========================================================================================
' Process WM_NOTIFY message for window/dialog: frmProjectManager
' ========================================================================================
Function frmProjectManager_OnNotify(ByVal HWnd As HWnd, ByVal id As Long, ByVal pNMHDR As NMHDR Ptr) As LRESULT

   Dim nmi As NMITEMACTIVATE Ptr = Cast( NMITEMACTIVATE Ptr, pNMHDR)
   Dim hPopUpMenu As HMENU
   Dim pt As Point
   
   Select Case pNMHDR->code 
   
      Case NM_DBLCLK
         ' Load the selected file into the editor
         If id = IDC_FRMPROJECTMANAGER_LISTVIEW Then   
            ' Get the pDoc from the selected row                
            Dim pDoc As clsDocument Ptr = Cast( clsDocument Ptr, FF_ListView_GetlParam(nmi->hdr.hwndFrom, nmi->iItem) )
            frmMain_OpenFileSafely(HWND_FRMMAIN, False, False, True, pDoc->DiskFilename, pDoc )
         End If

      Case NM_RCLICK
         If id = IDC_FRMPROJECTMANAGER_LISTVIEW Then   
            ' Create the popup menu
            If nmi->iItem > -1 Then
              hPopupMenu = CreateProjectManagerContextMenu(nmi->hdr.hwndFrom, nmi->iItem)
              GetCursorPos @pt
              TrackPopupMenu(hPopUpMenu, 0, pt.x, pt.y, 0, HWnd, ByVal Null)
              DestroyMenu hPopUpMenu
            End If
            Return True   ' prevent further processing that leads to WM_CONTEXTMENU
         End If

   End Select

   Function = 0
End Function



' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmProjectManager
' ========================================================================================
Function frmProjectManager_OnDestroy(HWnd As HWnd) As LRESULT
   Function = 0
End Function


' ========================================================================================
' frmProjectManager Window procedure
' ========================================================================================
Function frmProjectManager_WndProc( ByVal HWnd   As HWnd, _
                                    ByVal uMsg   As UINT, _
                                    ByVal wParam As WPARAM, _
                                    ByVal lParam As LPARAM _
                                    ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_NOTIFY,   frmProjectManager_OnNotify)
      HANDLE_MSG (HWnd, WM_COMMAND,  frmProjectManager_OnCommand)
      HANDLE_MSG (HWnd, WM_SIZE,     frmProjectManager_OnSize)
      HANDLE_MSG (HWnd, WM_CLOSE,    frmProjectManager_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,  frmProjectManager_OnDestroy)
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProcW(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmProjectManager_Show
' ========================================================================================
Function frmProjectManager_Show( ByVal hWndParent As HWnd, _
                                 ByVal nCmdShow   As Long = 0 _
                                 ) As Long

   ' If project manager already exists then toggle its visibility
   ' unless nCmdShow is explicitly set to show it.
   If IsWindow(HWND_FRMPROJECTMANAGER) Then
      If nCmdShow <> SW_SHOW Then nCmdShow = Iif(IsWindowVisible(HWND_FRMPROJECTMANAGER), SW_HIDE, SW_SHOW)
      ShowWindow HWND_FRMPROJECTMANAGER, nCmdShow
      Exit Function
   End If
      

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   HWND_FRMPROJECTMANAGER = _
   pWindow->Create( hWndParent, L(188,"Project Manager"), @frmProjectManager_WndProc, 0, 0, 0, 0, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->SetClientSize(600, 350)
   pWindow->Center(pWindow->hWindow, hWndParent)
      
   ' Set the small and large icon for the main window (must be set after main window is created)
   pWindow->BigIcon   =  LoadImage( pWindow->InstanceHandle, "IMAGE_MAINICON", IMAGE_ICON, 32, 32, LR_SHARED)
   pWindow->SmallIcon =  LoadImage( pWindow->InstanceHandle, "IMAGE_MAINICON", IMAGE_ICON, 16, 16, LR_SHARED)

   Dim As HWnd hLV = _
        pWindow->AddControl("LISTVIEW", , IDC_FRMPROJECTMANAGER_LISTVIEW, "", 10, 10, 176, 335, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or LVS_REPORT Or LVS_SHOWSELALWAYS , _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR)

   ' Add some extended styles
   Dim dwExStyle As DWORD
   dwExStyle = ListView_GetExtendedListViewStyle(hLV)
   dwExStyle = dwExStyle Or LVS_EX_FULLROWSELECT Or LVS_EX_GRIDLINES
   ListView_SetExtendedListViewStyle(hLV, dwExStyle)

   ' Configure the ListView
   ListView_AddColumn( hLV, 0, WStr("File"), 150 * pWindow->rxRatio )
   ListView_AddColumn( hLV, 1, WStr("Type"), 100 * pWindow->rxRatio )
   ListView_AddColumn( hLV, 2, WStr("Location"), 500 * pWindow->rxRatio )
                                                                       
                                                                       
   Dim wzText As WString * MAX_PATH
   
   wzText = L(188,"Project Manager") & " - [" & *gpApp->ProjectName & "]"
   SetWindowTextW( HWND_FRMPROJECTMANAGER, wzText )                     
   
   
   frmProjectManager_PositionWindows(HWND_FRMPROJECTMANAGER)

   ShowWindow HWND_FRMPROJECTMANAGER, nCmdShow
   SetFocus hLV
   
   Function = 0
End Function
