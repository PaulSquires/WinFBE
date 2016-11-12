
' ========================================================================================
' Set the Listbox selected item to the currently selected open document (if applicable)
' ========================================================================================
Function frmFnList_SetListBoxPosition() As Long

   Dim pList   As FUNCTION_TYPE Ptr
   Dim nCurSel As Long = -1

   Dim As HWnd hList  = GetDlgItem(HWND_FRMFNLIST, IDC_FRMFNLIST_LISTBOX)
   Dim As Long nCount = Listbox_GetCount(hList)
   
   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr() 
   If pDoc Then 
      For i As Long = 0 To nCount - 1
         pList = Cast(FUNCTION_TYPE Ptr, Listbox_GetItemData(hList, i))
         If pList Then
            If pList->pDoc = pDoc Then
               nCurSel = i: Exit For
            End If
         End If
      Next
   End If            
   
   Listbox_SetCurSel(hList, nCurSel)
   
   Function = 0
End Function


' ========================================================================================
' Update the Listbox with the list of functions from all loaded documents
' ========================================================================================
Function frmFnList_UpdateListBox() As Long

   Dim As HWnd hList = GetDlgItem(HWND_FRMFNLIST, IDC_FRMFNLIST_LISTBOX)
   Dim szTemp As WString * MAX_PATH
   
   ' Clear the listbox and add the new items
   Listbox_ResetContent(hList)
   
   Dim pDoc As clsDocument Ptr
   Dim pList As FUNCTION_TYPE Ptr
   
   Dim As Long i, idx
   Dim As Long nCount = gApp.GetDocumentCount()
   
   For i = 0 To nCount - 1
      pDoc = gApp.GetDocumentPtr(i)
      If pDoc Then 
         ' Display all the sub/function names in this source file
         pList = pDoc->FnListPtr
         Do Until pList = 0
            ' Draw the function or header name and store pointer to pList
            idx = Listbox_AddString( hList, @pList->zFnName )
            ListBox_SetItemData( hList, idx, pList )    ' store pointer to type structure
            pList = pList->pListNext
         Loop      
      End If   
   Next
   
   Function = 0
End Function



' ========================================================================================
' Process WM_COMMAND message for window/dialog: FnList
' ========================================================================================
Function frmFnList_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   Select Case id
      Case IDC_FRMFNLIST_LISTBOX
         If codeNotify = LBN_DBLCLK Then
            Dim As Long nCurSel = Listbox_GetCurSel(hwndCtl)
            If nCurSel = -1 Then Exit Function
            
            ' Retrieve the pointer to the pList and attempt to load the document at the correct 
            ' location. If the pointer is Null then we have clicked on the header element which
            ' is the name of the file. 
            Dim pDoc As clsDocument Ptr 
            Dim pList As FUNCTION_TYPE Ptr = Cast(FUNCTION_TYPE Ptr, Listbox_GetItemData(hWndCtl, nCurSel))
            
            If pList = 0 Then  ' header row
               ' Get the filename and attempt to open the document
            Else
               ' Display the document containing the selected sub/function       
               pDoc = pList->pDoc
               frmMain_OpenFileSafely(HWND_FRMMAIN, _
                                       False, _    ' bIsNewFile
                                       False, _    ' bIsTemplate
                                       true, _     ' bShowInTab
                                       false, _    ' bIsInclude
                                       pDoc->DiskFilename, _  ' pwszName
                                       pDoc )  ' pDocIn
               ' Set the top line to display in the editor. I chose to start 3 lines before the
               ' function just to make it visually more appealing.
               SendMessageW( pDoc->hWindow, SCI_SETFIRSTVISIBLELINE, pList->nLineNumber - 3, 0) 
               SendMessageW( pDoc->hWindow, SCI_GOTOLINE, pList->nLineNumber, 0) 
            End If
            If gConfig.CloseFuncList Then ShowWindow( HWND_FRMFNLIST, SW_HIDE )
         End If
         
   End Select
   
   Function = 0
End Function


' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
Function frmFnList_PositionWindows( ByVal HWnd As HWnd ) As LRESULT
   
   ' Get the entire client area
   Dim As Rect rc
   GetClientRect(HWnd, @rc)
   
   Dim As HWnd hList = GetDlgItem(HWnd, IDC_FRMFNLIST_LISTBOX )
   SetWindowPos hList, 0, 0, 0, rc.Right-rc.Left, rc.Bottom-rc.top, SWP_SHOWWINDOW Or SWP_NOZORDER

   Function = 0
End Function


' ========================================================================================
' Process WM_SIZE message for window/dialog: FnList
' ========================================================================================
Function frmFnList_OnSize(ByVal HWnd As HWnd, ByVal state As UINT, ByVal cx As Long, ByVal cy As Long) As LRESULT

   If state <> SIZE_MINIMIZED Then 
      frmFnList_PositionWindows(HWnd)
   End If   

   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: FnList
' ========================================================================================
Function frmFnList_OnClose( ByVal HWnd As HWnd ) As LRESULT
   ' Never close the window; simply hide it.
   ShowWindow( HWnd, SW_HIDE )
   Function = 0
End Function


' ========================================================================================
' Process WM_MEASUREITEM message for window/dialog: FnList
' ========================================================================================
Function frmFnList_OnMeasureItem( ByVal HWnd As HWnd, ByVal lpmis As MEASUREITEMSTRUCT Ptr ) As Long
   ' Set the height of the List box items. 
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)

   If lpmis->CtlId <> IDC_FRMFNLIST_LISTBOX Then Exit Function
   lpmis->itemHeight = 18 * pWindow->ryRatio
      
   Function = 0
End Function


' ========================================================================================
' Process WM_DRAWITEM message for window/dialog: FnList
' ========================================================================================
Function frmFnList_OnDrawItem( ByVal HWnd As HWnd, ByVal lpdis As Const DRAWITEMSTRUCT Ptr ) As Long

   Dim As HBRUSH hBrush 
   Dim As HFONT hFont, hFontOld
   Dim As Rect rc 
   Dim wzText As WString * MAX_PATH
   
   If lpdis->itemID = -1 Then Exit Function
   
   Dim pWindow As CWindow Ptr = AfxCWindowOwnerPtr(HWnd)
   Dim pList As FUNCTION_TYPE Ptr
   
   Select Case lpdis->itemAction
      Case ODA_DRAWENTIRE, ODA_SELECT
   
         ' CLEAR BACKGROUND
         hBrush = GetSysColorBrush(COLOR_WINDOW) 
         SelectObject(lpdis->hDC, hBrush)      
         FillRect(lpdis->hDC, @lpdis->rcItem, hBrush)  
   
         ' DRAW TEXT
         SetBkColor(lpdis->hDC, GetSysColor(COLOR_WINDOW))   
         SetTextColor(lpdis->hDC, GetSysColor(COLOR_WINDOWTEXT))
         SendMessageW(lpdis->hwndItem, LB_GETTEXT, lpdis->itemID, Cast(LPARAM, @wzText))
         
         ' Determine what font to use. The pointer to the pList data is held in the ItemData
         ' for each row in the ListBox. If that value is zero then we must be dealing with a
         ' header row.
         pList = Cast(FUNCTION_TYPE Ptr, lpdis->itemData)
         hFont = pWindow->CreateFont("", 9, Iif(pList->bIsHeader, FW_BOLD, FW_NORMAL) )
         hFontOld = SelectObject(lpdis->hDC, hFont)
         DrawTextW( lpdis->hDC, wzText, -1, Cast(lpRect, @lpdis->rcItem), _
                        DT_LEFT Or DT_SINGLELINE Or Iif(pList->bIsHeader, DT_BOTTOM, DT_VCENTER) )
         DeleteObject(SelectObject(lpdis->hDC, hFontOld))
   
         ' SELECTED ITEM
         If (lpdis->itemState And ODS_SELECTED) Then     
            rc.Left   = lpdis->rcItem.Left + 1
            rc.Right  = lpdis->rcItem.Right 
            rc.top    = lpdis->rcItem.top
            rc.bottom = lpdis->rcItem.bottom
            InvertRect(lpdis->hDC, @rc) 
         End If
   
         Function = True : Exit Function
   
      Case ODA_FOCUS
         DrawFocusRect(lpdis->hDC, @lpdis->rcItem) 
   
   End Select

   Function = 0
End Function

   
' ========================================================================================
' Process WM_DESTROY message for window/dialog: FnList
' ========================================================================================
Function frmFnList_OnDestroy(HWnd As HWnd) As LRESULT
   HWND_FRMFNLIST = 0
   Function = 0
End Function


' ========================================================================================
' Processes messages for the subclassed ListBox window.
' ========================================================================================
Function frmFnList_ListBox_SubclassProc ( _
                  ByVal HWnd   As HWnd, _                 ' // Control window handle
                  ByVal uMsg   As UINT, _                 ' // Type of message
                  ByVal wParam As WPARAM, _               ' // First message parameter
                  ByVal lParam As LPARAM, _               ' // Second message parameter
                  ByVal uIdSubclass As UINT_PTR, _        ' // The subclass ID
                  ByVal dwRefData As DWORD_PTR _          ' // Pointer to reference data
                  ) As LRESULT

   Select Case uMsg

      Case WM_GETDLGCODE
         ' All keyboard input
         Function = DLGC_WANTALLKEYS
         Exit Function

      Case WM_KEYUP
         Select Case Loword(wParam)
            Case VK_RETURN
               ' Simulate the sending of a LBN_DBLCLK to the control.
               SendMessageW( GetParent(HWnd), WM_COMMAND, MAKEWPARAM(IDC_FRMFNLIST_LISTBOX, LBN_DBLCLK), Cast(LPARAM,HWnd) )
               Exit Function
            Case VK_ESCAPE
               SendMessageW( GetParent(hWnd), WM_CLOSE, 0, 0 ) 
         End Select

      Case WM_DESTROY
         ' REQUIRED: Remove control subclassing
         RemoveWindowSubclass HWnd, @frmFnList_ListBox_SubclassProc, uIdSubclass

   End Select

   ' Default processing of Windows messages
   Function = DefSubclassProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' FnList Window procedure
' ========================================================================================
Function frmFnList_WndProc( ByVal HWnd   As HWnd, _
                            ByVal uMsg   As UINT, _
                            ByVal wParam As WPARAM, _
                            ByVal lParam As LPARAM _
                            ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_COMMAND,     frmFnList_OnCommand)
      HANDLE_MSG (HWnd, WM_SIZE,        frmFnList_OnSize)
      HANDLE_MSG (HWnd, WM_CLOSE,       frmFnList_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,     frmFnList_OnDestroy)
      HANDLE_MSG (HWnd, WM_MEASUREITEM, frmFnList_OnMeasureItem)
      HANDLE_MSG (HWnd, WM_DRAWITEM,    frmFnList_OnDrawItem)
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProcW(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmFnList_Show
' ========================================================================================
Function frmFnList_Show( ByVal hWndParent As HWnd, _
                         ByVal nCmdShow   As Long = 0 _
                         ) As Long

   ' If project manager already exists then toggle its visibility
   ' unless nCmdShow is explicitly set to show it.
   If IsWindow(HWND_FRMFNLIST) Then
      If nCmdShow <> SW_SHOW Then nCmdShow = Iif(IsWindowVisible(HWND_FRMFNLIST), SW_HIDE, SW_SHOW)
      ShowWindow HWND_FRMFNLIST, nCmdShow
      If nCmdShow = SW_SHOW Then
         frmFnList_SetListBoxPosition
         SetFocus GetDlgItem(HWND_FRMFNLIST, IDC_FRMFNLIST_LISTBOX)
      End If
      Exit Function 
   Else 
      ' If the window does not exist yet then ensure that it becomes visible after creation.
      nCmdShow = SW_SHOW
   End If
      

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   Dim rcWork  As Rect = pWindow->GetWorkArea
   Dim nHeight As Long = (rcWork.Bottom - rcWork.Top) * .60
   Dim nWidth  As Long = (rcWork.Right - rcWork.Left) * .80

   HWND_FRMFNLIST = _
   pWindow->Create( hWndParent, L(223,"Function List"), @frmFnList_WndProc, 0, 0, nWidth, nHeight, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->Center(pWindow->hWindow, hWndParent)
   
   
   ' Set the small and large icon for the main window (must be set after main window is created)
   pWindow->BigIcon   =  LoadImage( pWindow->InstanceHandle, "IMAGE_AAA_MAINICON", IMAGE_ICON, 32, 32, LR_SHARED)
   pWindow->SmallIcon =  LoadImage( pWindow->InstanceHandle, "IMAGE_AAA_MAINICON", IMAGE_ICON, 16, 16, LR_SHARED)

   Dim As HWnd hList = _ 
        pWindow->AddControl("LISTBOX", , IDC_FRMFNLIST_LISTBOX, "", 0, 0, 0, 0, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_HSCROLL Or LBS_MULTICOLUMN Or LBS_NOINTEGRALHEIGHT Or _
        LBS_OWNERDRAWFIXED Or LBS_HASSTRINGS Or LBS_NOTIFY, WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR, , _
        Cast(WNDPROC, @frmFnList_ListBox_SubclassProc), IDC_FRMFNLIST_LISTBOX, Cast(DWORD_PTR, @pWindow))
   
   ' Can not use a variable height multicolumn listbox because Microsoft says the following:
   ' "  The LBS_MULTICOLUMN and LBS_OWNERDRAWVARIABLE styles cannot be combined. If both are "
   '    specified, LBS_OWNERDRAWVARIABLE is ignored.  "

   SendMessageW( hList, LB_SETCOLUMNWIDTH, 250 * pWindow->rxRatio, 0 )
   
   frmFnList_UpdateListBox()
   frmFnList_SetListBoxPosition()

   frmFnList_PositionWindows(HWND_FRMFNLIST)

   ShowWindow HWND_FRMFNLIST, nCmdShow
   SetFocus hList
   
   Function = 0
End Function





