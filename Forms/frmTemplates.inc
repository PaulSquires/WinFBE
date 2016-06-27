
' ========================================================================================
' Process WM_CREATE message for window/dialog: frmTemplates
' ========================================================================================
Function frmTemplates_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN

   ' This is a modal popup window so disable the parent window
   EnableWindow GetParent(HWnd), False

   '  Message cracker macro expects a True to be returned for a successful
   '  OnCreate handler even though returning -1 from a standard WM_CREATE
   '  call would stop creating the window. This is just one of those Windows
   '  inconsistencies.
   Return True
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmTemplates
' ========================================================================================
Function frmTemplates_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   Select Case id
      Case IDOK  
         If codeNotify = BN_CLICKED Then
            ' Let the IDC_FRMTEMPLATES_LISTBOX message to process it
            SendMessageW HWnd, WM_COMMAND, MAKELONG(IDC_FRMTEMPLATES_LISTBOX, LBN_DBLCLK), 0
            Exit Function
         End If
      
      Case IDCANCEL ' button clicked or ESC pressed
         If codeNotify = BN_CLICKED Then
            SendMessageW HWnd, WM_CLOSE, 0, 0
            Exit Function
         End If

      Case IDC_FRMTEMPLATES_LISTBOX
         If codeNotify = LBN_DBLCLK Then
            ' Get the handle of the Listbox
            Dim hListBox As HWnd = GetDlgItem(HWnd, IDC_FRMTEMPLATES_LISTBOX)
            ' Get the current selection
            Dim curSel As Long = SendMessageW(hListBox, LB_GETCURSEL, 0, 0)
            If curSel = LB_ERR Then Exit Function
            ' Get the stored index
            Dim pwszPath As WString Ptr = Cast(WString Ptr, SendMessageW(hListBox, LB_GETITEMDATA, Cast(WPARAM, curSel), 0))
            If pwszPath = LB_ERR Then Exit Function
            ' Open the template
            If pwszPath Then
               ' Create the new Scintilla window and load the file
               Dim pDoc As clsDocument Ptr = New clsDocument
               pDoc->CreateCodeWindow(HWND_FRMMAIN, True, True, pwszPath)
            
               ' Add the new document to the global app
               gpApp->AddDocument(pDoc)                
               
               ' Add the new document to the top tabcontrol
               Dim As Long idx = gTTabCtl.AddTab( pDoc )
               gTTabCtl.SetFocusTab(idx)
               
               ' Position all of the controls into place
               frmMain_PositionWindows(HWND_FRMMAIN)
   
               ' Ensure that the code window scrolls to the correct location
               Dim As Long curPos = SendMessageW( pDoc->hWindow, SCI_GETCURRENTPOS, 0, 0)
               SendMessageW( pDoc->hWindow, SCI_GOTOPOS, curPos, 0 )

            End If
            ' Close the dialog
            SendMessageW HWnd, WM_CLOSE, 0, 0
            Exit Function
         End If
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmTemplates
' ========================================================================================
Function frmTemplates_OnClose(HWnd As HWnd) As LRESULT
   ' Enables parent window keeping parent's zorder
   EnableWindow GetParent(HWnd), True 
   DestroyWindow HWnd
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmTemplates
' ========================================================================================
Function frmTemplates_OnDestroy(HWnd As HWnd) As LRESULT
   PostQuitMessage(0)
   Function = 0
End Function


' ========================================================================================
' Templates window procedure
' ========================================================================================
Function frmTemplates_WndProc (ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CREATE,   frmTemplates_OnCreate)
      HANDLE_MSG (HWnd, WM_CLOSE,    frmTemplates_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,  frmTemplates_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,  frmTemplates_OnCommand)
   End Select

   ' Default processing of Windows messages
   Function = DefWindowProcW(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' Templates popup dialog
' ========================================================================================
Function frmTemplates_Show (ByVal hParent As HWnd, ByVal x As Long, ByVal y As Long) As Long

   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hParent)->DPI

   pWindow->Create(hParent, L(176,"Templates"), @frmTemplates_WndProc, _
               pWindow->UnScaleX(x), pWindow->UnScaleY(y), 0, 0, _
               WS_VISIBLE Or WS_CAPTION Or WS_POPUPWINDOW, WS_EX_WINDOWEDGE)
   pWindow->SetClientSize(420,395)
   
   ' Set the small and large icon for the main window (must be set after main window is created)
   pWindow->BigIcon   =  LoadImage( pWindow->InstanceHandle, "IMAGE_MAINICON", IMAGE_ICON, 32, 32, LR_SHARED)
   pWindow->SmallIcon =  LoadImage( pWindow->InstanceHandle, "IMAGE_MAINICON", IMAGE_ICON, 16, 16, LR_SHARED)

   ' Add a listbox
   Dim hListBox As HWnd = pWindow->AddControl("ListBox", , IDC_FRMTEMPLATES_LISTBOX, "", _
      WS_CHILD Or WS_VISIBLE Or WS_HSCROLL Or WS_VSCROLL Or WS_BORDER Or WS_TABSTOP Or _
      LBS_STANDARD Or LBS_HASSTRINGS Or LBS_SORT Or LBS_NOTIFY Or LBS_NOINTEGRALHEIGHT, WS_EX_CLIENTEDGE)
      
   pWindow->SetWindowPos hListBox, Null, 8, 8, 400, 337, SWP_NOZORDER
   SendMessageW hListBox, LB_SETHORIZONTALEXTENT, Cast(WPARAM, 600 * pWindow->rxRatio), 0

   ' Add the buttons
   pWindow->AddControl("Button", , IDOK, L(0,"&OK"), 245, 353, 74, 28, _
               WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_FLAT Or BS_DEFPUSHBUTTON, WS_EX_NOPARENTNOTIFY)
   pWindow->AddControl("Button", , IDCANCEL, L(1,"&Cancel"), 333, 353, 74, 28, _
               WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_FLAT, WS_EX_NOPARENTNOTIFY)

   ' Search templates
   Dim hSearch As HANDLE, WFD As WIN32_FIND_DATAW, FileNo As Long, nType As Long, nItem As Long
   Dim wszPath As WString * MAX_PATH, wszCurPath As WString * MAX_PATH, wszText As WString * 260
   Dim wszFullPath As WString * MAX_PATH * 2, idx As Long, nLin As Long
   wszPath = Exepath & "\Templates\"
   wszCurPath = wszPath & "*.fbtpl"

   ' Count the number of files and dimension the array
   ' REDIM PRESERVE causes problems
   Dim nCount As Long
   hSearch = FindFirstFileW(wszCurPath, @WFD)
   If hSearch <> INVALID_HANDLE_VALUE Then
      Do
         If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY Then
            nCount +=1
         End If
      Loop While FindNextFileW(hSearch, @WFD)
      FindClose(hSearch)
   End If
   ' Dimension the array
   If nCount = 0 Then Exit Function
   Dim rgwszPaths(nCount - 1) As WString * MAX_PATH
   idx = 0

   ' Find the files
   hSearch = FindFirstFileW(wszCurPath, @WFD)
   If hSearch <> INVALID_HANDLE_VALUE Then
      Do
         If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY Then
            wszFullPath = wszPath & WFD.cFileName
            ' Get the description
            FileNo = Freefile
            If Open( wszFullPath For Input As #FileNo) = 0 Then 
               nLin = 0
               Do Until Eof(FileNo)
                  Line Input #FileNo, wszText
                  nLin += 1
                  If nLin = 1 Then nType = Val(wszText)
                  If nType < 1 Or nType > 2 Then Exit Do
                  If nType = 1 And nLin = 3 Then Exit Do
                  If nType = 2 And nLin = 4 Then Exit Do
               Loop
               Close FileNo
               ' Display the description in the listbox
               nItem = SendMessageW(hListBox, LB_ADDSTRING, 0, Cast(LPARAM, @wszText))
               ' Store the full path in the array
               rgwszPaths(idx) = wszFullPath
               ' Store a pointer to the element of the array in the listbox item
               SendMessageW hListBox, LB_SETITEMDATA, nItem, Cast(LPARAM, Varptr(rgwszPaths(idx)))
               idx += 1
            End If
         End If
      Loop While FindNextFileW(hSearch, @WFD)
      FindClose(hSearch)
   End If

   ' Process Windows messages
   Function = pWindow->DoEvents

   ' Delete the frmTemplates CWindow class manually allocated memory 
   Delete pWindow

End Function
