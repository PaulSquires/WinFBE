
' ========================================================================================
' Listview callback compare function for sorting columns
' ========================================================================================
Function ListView_CompareFunc( ByVal index1 As Long, _
                               ByVal index2 As Long, _
                               ByVal lvs As LVSORT_TYPE Ptr _
                               ) As Long

   Dim wzItem1 As WString * MAX_PATH
   Dim wzItem2 As WString * MAX_PATH
   Dim lReturn As Long
        
   ' Get the values from the ListView     
   ListView_GetItemText( lvs->hLV, index1, lvs->nColumn, @wzItem1, MAX_PATH )
   ListView_GetItemText( lvs->hLV, index2, lvs->nColumn, @wzItem2, MAX_PATH )
    
   lReturn = Lstrcmpi(wzItem1, wzItem2)
    
   Function = Iif( lvs->nSortOrder = SO_ASCENDING, lReturn, -lReturn )                         

End Function


' ========================================================================================
' Draw the sort arrow for the listview column
' ========================================================================================
Function ListView_DrawSortArrow( ByVal hLV As HWnd, ByVal nColumn As Long, ByVal nSortOrder As Long) As Long

   Dim As HWnd hHeader = Cast(HWnd, SendMessage(hLV, LVM_GETHEADER, 0, 0))

   Dim HDI As HD_ITEM

   HDI.mask = HDI_FORMAT
   Header_GetItem( hHeader, nColumn, @HDI )
    
   Select Case nSortOrder
     Case SO_NONE      
        HDI.fmt = HDI.fmt And Not (HDF_SORTDOWN Or HDF_SORTUP) 
     Case SO_ASCENDING
        HDI.fmt = HDI.fmt And Not HDF_SORTDOWN
        HDI.fmt = HDI.fmt Or      HDF_SORTUP                       
     Case SO_DESCENDING
        HDI.fmt = HDI.fmt And Not HDF_SORTUP 
        HDI.fmt = HDI.fmt Or      HDF_SORTDOWN                             
   End Select       

   Header_SetItem( hHeader, nColumn, @HDI )
    
   Function = 0
End Function


' ========================================================================================
' Sort the listview based on the specified column and display the sort arrow
' ========================================================================================
Function Listview_SortColumn( ByVal hLV As HWnd, _
                              ByVal nColumn As Long, _
                              ByVal nSortOrder As Long _
                              ) As Long

   Dim lvs As LVSORT_TYPE
   lvs.hLV        = hLV
   lvs.nColumn    = nColumn
   lvs.nSortOrder = nSortOrder
   
   ' Remove any existing sort arrows and then set the current column arrrow
   ListView_DrawSortArrow( hLV, gLVProjSortCol, SO_NONE )
   
   ' set the sort arrow for the specified column
   ListView_DrawSortArrow( hLV, nColumn, nSortOrder )
   ListView_SortItemsEx( hLV, @ListView_CompareFunc, @lvs )
                       
   gLVProjSortCol = nColumn
   gLVProjSortDir = nSortOrder

   Function = 0
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmProjectManager
' ========================================================================================
Function frmProjectManager_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   Dim As HWnd hLV = GetDlgItem(HWnd, IDC_FRMPROJECTMANAGER_LISTVIEW )
   Dim pDoc As clsDocument Ptr 

   Select Case id
      Case IDM_OPENINCLUDE,IDM_REMOVEFILEFROMPROJECT, IDM_SETFILENORMAL, IDM_SETFILEMODULE, _
           IDM_SETFILEMAIN,IDM_SETFILERESOURCE
                  
         Dim As Long lIndex = -1 
         If ListView_GetSelectedCount(hLV) > 0 Then                
            Do
               ' Get the next selected index
               lIndex = ListView_GetNextItem(hLV, lIndex, LVNI_SELECTED)
               If lIndex <> -1 Then    
                  pDoc = Cast( clsDocument Ptr, FF_ListView_GetlParam(hLV, lIndex) )
                  Select Case id
                     Case IDM_OPENINCLUDE           
                          frmMain_OpenFileSafely(HWND_FRMMAIN, False, False, True, pDoc->DiskFilename, pDoc )
                     Case IDM_REMOVEFILEFROMPROJECT
                          Function = gpApp->ProjectRemoveFile( pDoc )
                     Case IDM_SETFILENORMAL:         Function = gpApp->ProjectSetFileType( pDoc, FILETYPE_NORMAL )
                     Case IDM_SETFILEMODULE:         Function = gpApp->ProjectSetFileType( pDoc, FILETYPE_MODULE )
                     Case IDM_SETFILEMAIN:           Function = gpApp->ProjectSetFileType( pDoc, FILETYPE_MAIN )
                     Case IDM_SETFILERESOURCE:       Function = gpApp->ProjectSetFileType( pDoc, FILETYPE_RESOURCE )
                  End Select
               End If
            Loop Until lIndex = -1
         End If
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

   Dim wzText As WString * MAX_PATH
   wzText = L(188,"Project Manager") & " - [" & *gpApp->ProjectName & "]"
   SetWindowTextW( HWND_FRMPROJECTMANAGER, wzText )                     
   
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
            If pDoc Then
               frmMain_OpenFileSafely(HWND_FRMMAIN, False, False, True, pDoc->DiskFilename, pDoc )
            End If
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
         
      Case LVN_COLUMNCLICK
         If id = IDC_FRMPROJECTMANAGER_LISTVIEW Then
            
            ' Sort the listview based on the clicked column header 
            Dim lpNMV As NM_LISTVIEW Ptr = Cast( NM_LISTVIEW Ptr, pNMHDR )

            Static fSortOrder As Long = SO_ASCENDING
            fSortOrder = Iif(fSortOrder = SO_ASCENDING, SO_DESCENDING, SO_ASCENDING)
            
            Listview_SortColumn( lpNMV->hdr.hwndFrom, lpNMV->iSubItem, fSortOrder )
            
         End If
         
      Case NM_CUSTOMDRAW
         Dim lpLvCd As NMLVCUSTOMDRAW Ptr = Cast(NMLVCUSTOMDRAW Ptr, pNMHDR)
         Dim pDoc As clsDocument Ptr
         Dim As Long lRow, lCol  
             
         If id = IDC_FRMPROJECTMANAGER_LISTVIEW  Then

            Select Case lpLvCd->nmcd.dwDrawStage
               Case CDDS_PREPAINT, CDDS_ITEMPREPAINT
                  Return CDRF_NOTIFYSUBITEMDRAW
                     
               Case CDDS_ITEMPREPAINT Or CDDS_SUBITEM
                  ' to retrieve the line and column number of the cell
                  lRow = lpLvCd->nmcd.dwItemSpec  ' Line number
                  lCol = lpLvCd->iSubItem         ' Column number                                                                
                  
                  If lCol = 1 Then   ' filetype column
                     pDoc = Cast( clsDocument Ptr, FF_ListView_GetlParam(pNMHDR->hwndFrom, lRow) )
                     If pDoc Then
                        Select Case pDoc->ProjectFileType
                           Case FILETYPE_NORMAL:    lpLvCd->clrText = BGR(0,0,0)        ' black
                           Case FILETYPE_MAIN:      lpLvCd->clrText = BGR(153,000,000)  ' dark red  
                           Case FILETYPE_MODULE:    lpLvCd->clrText = BGR(051,102,000)  ' dark green  
                           Case FILETYPE_RESOURCE:  lpLvCd->clrText = BGR(173,000,173)  ' dark purple
                        End Select   
                     End If       
                  Else
                     lpLvCd->clrText = BGR(0,0,0)  ' black  
                     'lpLvCd->clrTextBk = &H00D8E9EC  ' grey backcolor
                  End If
                  Return CDRF_NEWFONT
                                               
            End Select
         End If

   End Select

   Function = 0
End Function



' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmProjectManager
' ========================================================================================
Function frmProjectManager_OnDestroy(HWnd As HWnd) As LRESULT
   HWND_FRMPROJECTMANAGER = 0
   Function = 0
End Function


' ========================================================================================
' Processes messages for the subclassed Listview window.
' ========================================================================================
Function frmProjectManager_Listview_SubclassProc ( _
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
               ' Simulate the sending of a WM_COMMAND message to the control with the IDM_OPENINCLUDE
               ' value. This will invoke the code in WM_COMMAND that cycles through all selected rows
               ' in the listview and opens the selected files.
               SendMessage( GetParent(HWnd), WM_COMMAND, MAKEWPARAM(IDM_OPENINCLUDE, 0), 0 )
               Exit Function
         End Select

      Case WM_DESTROY
         ' REQUIRED: Remove control subclassing
         RemoveWindowSubclass HWnd, @frmProjectManager_Listview_SubclassProc, uIdSubclass

   End Select

   ' Default processing of Windows messages
   Function = DefSubclassProc(HWnd, uMsg, wParam, lParam)

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
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   Dim rcWork As Rect = pWindow->GetWorkArea
   Dim nHeight As Long = (rcWork.Bottom - rcWork.Top) * .80

   HWND_FRMPROJECTMANAGER = _
   pWindow->Create( hWndParent, L(188,"Project Manager"), @frmProjectManager_WndProc, 0, 0, 500, nHeight, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->Center(pWindow->hWindow, hWndParent)
   
   
   ' Set the small and large icon for the main window (must be set after main window is created)
   pWindow->BigIcon   =  LoadImage( pWindow->InstanceHandle, "IMAGE_MAINICON", IMAGE_ICON, 32, 32, LR_SHARED)
   pWindow->SmallIcon =  LoadImage( pWindow->InstanceHandle, "IMAGE_MAINICON", IMAGE_ICON, 16, 16, LR_SHARED)

   Dim As HWnd hLV = _   ' subclassed to catch ENTER key
        pWindow->AddControl("LISTVIEW", , IDC_FRMPROJECTMANAGER_LISTVIEW, "", 10, 10, 176, 335, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or LVS_REPORT Or LVS_SHOWSELALWAYS , _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR, , _
        Cast(WNDPROC, @frmProjectManager_Listview_SubclassProc), IDC_FRMPROJECTMANAGER_LISTVIEW, Cast(DWORD_PTR, @pWindow))

   ' Add some extended styles
   Dim dwExStyle As DWORD
   dwExStyle = ListView_GetExtendedListViewStyle(hLV)
   dwExStyle = dwExStyle Or LVS_EX_FULLROWSELECT Or LVS_EX_GRIDLINES
   ListView_SetExtendedListViewStyle(hLV, dwExStyle)

   ' Configure the ListView
   ListView_AddColumn( hLV, 0, WStr("File"), 150 * pWindow->rxRatio )
   ListView_AddColumn( hLV, 1, WStr("Type"), 100 * pWindow->rxRatio )
   ListView_AddColumn( hLV, 2, WStr("Location"), 500 * pWindow->rxRatio )
                                                                       
                                                                       
   frmProjectManager_PositionWindows(HWND_FRMPROJECTMANAGER)

   ShowWindow HWND_FRMPROJECTMANAGER, nCmdShow
   SetFocus hLV
   
   Function = 0
End Function
