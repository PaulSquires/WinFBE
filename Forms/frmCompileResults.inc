
' ========================================================================================
' Set the cursor to the error position based on the selected line in frmCompileResults
' ========================================================================================
Function SetDocumentErrorPosition( ByVal hLV As HWnd ) As Long

   Dim wszErrorLine As WString * MAX_PATH
   Dim wszErrorFile As WString * MAX_PATH
   
   Dim As Long idx
   Dim As Long nCurSel = ListView_GetSelection(hLV)
   If nCurSel < 0 Then Return 0 
   
   Dim pDoc As clsDocument Ptr 

   ' Get the line number and filename of the selected item
   FF_ListView_GetItemText(hLV, nCurSel, 0, @wszErrorLine, MAX_PATH)
   FF_ListView_GetItemText(hLV, nCurSel, 1, @wszErrorFile, MAX_PATH)

   If Dir(wszErrorFile) = "" Then Exit Function
   
   ' Position the editor to the Error line
   idx = gTTabCtl.GetTabIndexFromFilename( @wszErrorFile )
   If idx >= 0 Then
      gTTabCtl.SetFocusTab(idx)
      pDoc = gTTabCtl.GetActiveDocumentPtr()
   Else
      ' File is not already loaded. Load it now.
      pDoc = New clsDocument
      pDoc->CreateCodeWindow(HWND_FRMMAIN, False, False, @wszErrorFile)
   
      ' Add the new document to the global app
      gpApp->AddDocument(pDoc)                
      
      ' Add the new document to the top tabcontrol
      idx = gTTabCtl.AddTab( pDoc )
      gTTabCtl.SetFocusTab(idx)
      
      ' Position all of the controls into place
      frmMain_PositionWindows(HWND_FRMMAIN)
   End If 
   SendMessageW pDoc->hWindow, SCI_GOTOLINE, Val(wszErrorLine)-1, 0

End Function


' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
Function frmCompileResults_PositionWindows( ByVal HWnd As HWnd ) As LRESULT

   Dim As Rect rc
   
   ' Get the entire client area
   GetClientRect(HWnd, @rc)

   ' Ensure that the listview is the full size of the form
   Dim As HWnd hLV = GetDlgItem(HWnd, IDC_FRMCOMPILERESULTS_LISTVIEW)
   SetWindowPos hLV, 0, 0, 0, rc.Right, rc.Bottom, SWP_NOZORDER

   Function = 0
End Function


' ========================================================================================
' Process WM_CREATE message for window/dialog: frmCompileResults
' ========================================================================================
Function frmCompileResults_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN

   '  Message cracker macro expects a True to be returned for a successful
   '  OnCreate handler even though returning -1 from a standard WM_CREATE
   '  call would stop creating the window. This is just one of those Windows
   '  inconsistencies.
   Return True
End Function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmCompileResults
' ========================================================================================
Function frmCompileResults_OnSize(ByVal HWnd As HWnd, ByVal state As UINT, ByVal cx As Long, ByVal cy As Long) As LRESULT
   frmCompileResults_PositionWindows(HWnd)   
   Function = 0
End Function


' ========================================================================================
' Process WM_NOTIFY message for window/dialog: frmCompileResults
' ========================================================================================
Function frmCompileResults_OnNotify(ByVal HWnd As HWnd, ByVal id As Long, ByVal pNMHDR As NMHDR Ptr) As LRESULT

    
   Select Case id
      Case IDC_FRMCOMPILERESULTS_LISTVIEW
         
         Dim pNMLV As NM_LISTVIEW Ptr = Cast(NM_LISTVIEW Ptr, pNMHDR)
         Select Case pNMLV->hdr.Code
            Case NM_CLICK      ' left click                
            
            Case NM_RETURN, NM_DBLCLK     ' left double-click
               SetDocumentErrorPosition(pNMLV->hdr.hwndFrom)
            
            Case NM_RCLICK     ' right click
            Case NM_RDBLCLK    ' right double-click
         End Select

    End Select
             
   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmCompileResults
' ========================================================================================
Function frmCompileResults_OnClose(HWnd As HWnd) As LRESULT
   DestroyWindow HWnd
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmCompileResults
' ========================================================================================
Function frmCompileResults_OnDestroy(HWnd As HWnd) As LRESULT
   Function = 0
End Function


' ========================================================================================
' frmCompileResults Window procedure
' ========================================================================================
Function frmCompileResults_WndProc( ByVal HWnd   As HWnd, _
                                    ByVal uMsg   As UINT, _
                                    ByVal wParam As WPARAM, _
                                    ByVal lParam As LPARAM _
                                    ) As LRESULT
    
   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CREATE,   frmCompileResults_OnCreate)
      HANDLE_MSG (HWnd, WM_SIZE,     frmCompileResults_OnSize)
      HANDLE_MSG (HWnd, WM_CLOSE,    frmCompileResults_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,  frmCompileResults_OnDestroy)
      HANDLE_MSG (HWnd, WM_NOTIFY,   frmCompileResults_OnNotify)
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProcW(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmCompileResults_Show
' ========================================================================================
Function frmCompileResults_Show( ByVal hWndParent As HWnd, _
                                 ByVal nCmdShow   As Long = 0 _
                                 ) As Long

   '  Create the window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   HWND_FRMCOMPILERESULTS = _
       pWindow->Create( hWndParent, "Compile Results", @frmCompileResults_WndProc, 0, 0, 0, 0, _
            WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
            WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
       pWindow->SetClientSize(600, 375)
       pWindow->Center
       
   ' Set the small and large icon for the main window (must be set after main window is created)
   pWindow->BigIcon   =  LoadImage( pWindow->InstanceHandle, "IMAGE_MAINICON", IMAGE_ICON, 32, 32, LR_SHARED)
   pWindow->SmallIcon =  LoadImage( pWindow->InstanceHandle, "IMAGE_MAINICON", IMAGE_ICON, 16, 16, LR_SHARED)

   Dim As HWnd hLV = _
        pWindow->AddControl("LISTVIEW", , IDC_FRMCOMPILERESULTS_LISTVIEW, "", 10, 10, 176, 335, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or LVS_REPORT Or LVS_SHOWSELALWAYS , _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR)

   ' Add some extended styles
   Dim dwExStyle As DWORD
   dwExStyle = ListView_GetExtendedListViewStyle(hLV)
   dwExStyle = dwExStyle Or LVS_EX_FULLROWSELECT Or LVS_EX_GRIDLINES
   ListView_SetExtendedListViewStyle(hLV, dwExStyle)

   ' Configure the ListView
   ListView_AddColumn( hLV, 0, WStr("Line"), 75 * pWindow->rxRatio )
   ListView_AddColumn( hLV, 1, WStr("File"), 250 * pWindow->rxRatio )
   ListView_AddColumn( hLV, 2, WStr("Description"), 500 * pWindow->rxRatio )
   
   ' Ensure that ListView is full size of the form
   frmCompileResults_PositionWindows( HWND_FRMCOMPILERESULTS )
   
   SetFocus hLV
   
   Function = 0
End Function
