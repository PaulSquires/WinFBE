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
' Retrieve the control/toolbox name that is currently actively selected in the Toolbox
' ========================================================================================
function GetActiveToolboxControlName() as CWSTR

   Dim As HWnd hList1 = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTTOOLBOX)
   dim as long nCurSel = ListBox_GetCurSel(hList1)
   
   if nCurSel = -1 THEN return CTRL_POINTER
   
   ' The index into the global gToolbox array is stored in the line's data area.
   dim as long idx = ListBox_GetItemData(hList1, nCurSel)
   return gToolbox(idx).wszName
end function

' ========================================================================================
' Retrieve the control class name that is currently actively selected in the Toolbox
' ========================================================================================
function GetActiveToolboxControlClassName() as CWSTR

   Dim As HWnd hList1 = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTTOOLBOX)
   dim as long nCurSel = ListBox_GetCurSel(hList1)
   
   if nCurSel = -1 THEN return CTRL_POINTER
   
   ' The index into the global gToolbox array is stored in the line's data area.
   dim as long idx = ListBox_GetItemData(hList1, nCurSel)
   return gToolbox(idx).wszClassName
end function

' ========================================================================================
' Retrieve the control type that is currently actively selected in the Toolbox
' ========================================================================================
function GetActiveToolboxControlType() as Long

   Dim As HWnd hList1 = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTTOOLBOX)
   dim as long nCurSel = ListBox_GetCurSel(hList1)
   
   if nCurSel = -1 THEN return CTRL_POINTER
   
   ' The index into the global gToolbox array is stored in the line's data area.
   dim as long idx = ListBox_GetItemData(hList1, nCurSel)
   function = gToolbox(idx).nToolType
end function


' ========================================================================================
' Set the Tools listbox to incoming control type.
' ========================================================================================
function SetActiveToolboxControl( byval ControlType as long ) as Long

   Dim As HWnd hList1 = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTTOOLBOX)
   dim as long NumItems = ListBox_GetCount(hList1)
   dim as long idx
   
   for i as long = 0 to NumItems
      idx = ListBox_GetItemData(hList1, i)
      if gToolbox(idx).nToolType = ControlType THEN 
         ListBox_SetCurSel(hList1, i)
         exit for
      end if   
   NEXT
   
   function = 0
end function


' ========================================================================================
' Display the properties for the current active control/form.
' ========================================================================================
Function DisplayPropertyList( byval pDoc as clsDocument ptr ) as Long
   if pDoc = 0 then exit function

   dim pCtrl as clsControl ptr = pDoc->Controls.GetActiveControl
   if pCtrl = 0 THEN exit function
   
   dim as long idx
   dim as hwnd hList = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTPROPERTIES)
   
   ' clear the listbox
   ListBox_ResetContent(hList)
   
   ' Iterate through all properties and add them to the listbox
   for i as long = lbound(pCtrl->Properties) to ubound(pCtrl->Properties)
      idx = Listbox_AddString(hList, @"")
      ListBox_SetItemData(hList, idx, i)   ' store array index in listbox item
   NEXT
      
   Function = 0
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmVDToolbox
' ========================================================================================
Function frmVDToolbox_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   Function = 0
End Function

' ========================================================================================
' Process WM_NOTIFY message for window/dialog: frmVDToolbox
' ========================================================================================
Function frmVDToolbox_OnNotify(ByVal HWnd As HWnd, ByVal id As Long, ByVal pNMHDR As NMHDR Ptr) As LRESULT

   SELECT CASE id
      CASE IDC_FRMVDTOOLBOX_TABCONTROL
         dim as long iPage = TabCtrl_GetCurSel(pNMHDR->hwndFrom)
         SELECT CASE pNMHDR->code
            CASE TCN_SELCHANGE
               ' Show the selected page controls
               if iPage = 0 then ShowWindow( GetDlgItem(HWND, IDC_FRMVDTOOLBOX_LSTTOOLBOX), SW_SHOW)
               if iPage = 1 then ShowWindow( GetDlgItem(HWND, IDC_FRMVDTOOLBOX_LSTPROPERTIES), SW_SHOW)
            CASE TCN_SELCHANGING
               ' Hide the current page
               if iPage = 0 then ShowWindow( GetDlgItem(HWND, IDC_FRMVDTOOLBOX_LSTTOOLBOX), SW_HIDE)
               if iPage = 1 then ShowWindow( GetDlgItem(HWND, IDC_FRMVDTOOLBOX_LSTPROPERTIES), SW_HIDE)
         END SELECT

   END SELECT

   function = 0
end function

' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
Function frmVDToolbox_PositionWindows( ByVal HWnd As HWnd ) As LRESULT
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
   if pWindow = 0 THEN exit function
      
   ' Get the entire client area
   Dim As Rect rc
   GetClientRect(HWnd, @rc)
   
   Dim As HWnd hList1 = GetDlgItem(HWnd, IDC_FRMVDTOOLBOX_LSTTOOLBOX )
   Dim As HWnd hList2 = GetDlgItem(HWnd, IDC_FRMVDTOOLBOX_LSTPROPERTIES )
   Dim As HWnd hTabCtl = GetDlgItem(HWnd, IDC_FRMVDTOOLBOX_TABCONTROL )
   
   pWindow->SetWindowPos( hTabCtl, 0, 0, 0, rc.Right-rc.Left, 24, SWP_SHOWWINDOW Or SWP_NOZORDER )
   pWindow->SetWindowPos( hList1, 0, 0, 24, rc.Right-rc.Left, rc.Bottom-rc.top-24, SWP_NOZORDER )
   pWindow->SetWindowPos( hList2, 0, 0, 24, rc.Right-rc.Left, rc.Bottom-rc.top-24, SWP_NOZORDER )

   Function = 0
End Function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmVDToolbox
' ========================================================================================
Function frmVDToolbox_OnSize(ByVal HWnd As HWnd, ByVal state As UINT, ByVal cx As Long, ByVal cy As Long) As LRESULT
   If state <> SIZE_MINIMIZED Then 
      frmVDToolbox_PositionWindows(HWnd)
   End If   
   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmVDToolbox
' ========================================================================================
Function frmVDToolbox_OnClose( ByVal HWnd As HWnd ) As LRESULT
   ' Never close the window; simply hide it.
   ShowWindow( HWnd, SW_HIDE )
   Function = 0
End Function


' ========================================================================================
' Process WM_MEASUREITEM message for window/dialog: frmVDToolbox
' ========================================================================================
Function frmVDToolbox_OnMeasureItem( ByVal HWnd As HWnd, ByVal lpmis As MEASUREITEMSTRUCT Ptr ) As Long
   ' Set the height of the List box items. 
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
   lpmis->itemHeight = pWindow->ScaleY(20)
   Function = 0
End Function


' ========================================================================================
' Process WM_DRAWITEM message for window/dialog: frmVDToolbox
' ========================================================================================
Function frmVDToolbox_OnDrawItem( ByVal HWnd As HWnd, ByVal lpdis As Const DRAWITEMSTRUCT Ptr ) As Long

   Dim As HBRUSH hBrush 
   dim as HANDLE hImage
   Dim As RECT rc, rc2 
   Dim wzText As WString * MAX_PATH
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
   if pWindow = 0 THEN exit function
      
   If lpdis->itemID = -1 Then Exit Function
   
   Select Case lpdis->itemAction
      Case ODA_DRAWENTIRE, ODA_SELECT
         
         if lpdis->CtlID = IDC_FRMVDTOOLBOX_LSTTOOLBOX then
            ' CLEAR BACKGROUND
            If (lpdis->itemState And ODS_SELECTED) Then     
               hBrush = GetSysColorBrush(COLOR_HIGHLIGHT) 
               SetBkColor(lpdis->hDC, GetSysColor(COLOR_HIGHLIGHT))   
               SetTextColor(lpdis->hDC, GetSysColor(COLOR_HIGHLIGHTTEXT))
            else
               hBrush = GetSysColorBrush(COLOR_WINDOW) 
               SetBkColor(lpdis->hDC, GetSysColor(COLOR_WINDOW))   
               SetTextColor(lpdis->hDC, GetSysColor(COLOR_WINDOWTEXT))
            end if
            SelectObject(lpdis->hDC, hBrush)      
            FillRect(lpdis->hDC, @lpdis->rcItem, hBrush)  
            ' The index into the gToolBox array is stored in the itemData of the line.
            
            ' DRAW IMAGE
            dim as HDC hdcMem = CreateCompatibleDC(lpdis->hDC)
            hImage = LoadImage(pWindow->InstanceHandle, gToolBox(lpdis->itemData).wszImage, IMAGE_BITMAP, _
                               pWindow->ScaleX(16), pWindow->ScaleX(16), LR_LOADTRANSPARENT)
            SelectObject(hdcMem, hImage)
            BitBlt( lpdis->hDC, _
                    lpdis->rcItem.left + pWindow->ScaleX(6), _
                    lpdis->rcItem.top, _
                    pWindow->ScaleX(16), pWindow->ScaleY(16), _
                    hdcMem, 0, 0, SRCCOPY)
            DeleteDC(hdcMem)                   

            ' DRAW TEXT
            rc = lpdis->rcItem: rc.left = rc.left + pWindow->ScaleX(30)
            DrawText( lpdis->hDC, gToolBox(lpdis->itemData).wszName, -1, Cast(lpRect, @rc), _
                      DT_LEFT Or DT_SINGLELINE Or DT_VCENTER )
         end if

         if lpdis->CtlID = IDC_FRMVDTOOLBOX_LSTPROPERTIES then
            dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
            dim pCtrl as clsControl ptr 
            if pDoc then
               pCtrl = pDoc->Controls.GetActiveControl
               if pCtrl THEN
                  rc  = lpdis->rcItem
                  rc2 = lpdis->rcItem
                  rc.right = rc.left + gPropDivPos
                  rc2.left = rc.right
                  
                  ' CLEAR BACKGROUND
                  If (lpdis->itemState And ODS_SELECTED) Then     
                     SetBkColor(lpdis->hDC, GetSysColor(COLOR_HIGHLIGHT))   
                     SetTextColor(lpdis->hDC, GetSysColor(COLOR_HIGHLIGHTTEXT))
                     hBrush = GetSysColorBrush(COLOR_HIGHLIGHT) 
                  else
                     SetBkColor(lpdis->hDC, GetSysColor(COLOR_WINDOW))   
                     SetTextColor(lpdis->hDC, GetSysColor(COLOR_WINDOWTEXT))
                     hBrush = GetSysColorBrush(COLOR_WINDOW) 
                  end if

                  ' Draw the property name 
                  SelectObject(lpdis->hDC, hBrush)      
                  FillRect(lpdis->hDC, @rc, hBrush)  
                  rc.left = rc.left + pWindow->ScaleX(4)
                  DrawText( lpdis->hDC, _
                            pCtrl->Properties(lpdis->itemData).wszPropName, _
                            -1, Cast(lpRect, @rc), _
                            DT_LEFT Or DT_SINGLELINE Or DT_VCENTER )

                  ' Draw the current value
                  hBrush = GetSysColorBrush(COLOR_WINDOW) 
                  SelectObject(lpdis->hDC, hBrush)      
                  FillRect(lpdis->hDC, @rc2, hBrush)  
                  SetBkColor(lpdis->hDC, GetSysColor(COLOR_WINDOW))   
                  SetTextColor(lpdis->hDC, GetSysColor(COLOR_WINDOWTEXT))
                  rc2.left = rc2.left + pWindow->ScaleX(4)
                  DrawText( lpdis->hDC, _
                            pCtrl->Properties(lpdis->itemData).wszPropValue, _
                            -1, Cast(lpRect, @rc2), _
                            DT_LEFT Or DT_SINGLELINE Or DT_VCENTER )

                  ' Draw the border edges
                  SetBkMode(lpdis->hDC, TRANSPARENT)   
                  rc  = lpdis->rcItem: rc.right = rc.left + gPropDivPos
                  rc2 = lpdis->rcItem: rc2.left = rc.right
                  DrawEdge( lpdis->hDC, @rc, EDGE_SUNKEN, BF_BOTTOMRIGHT)
                  DrawEdge( lpdis->hDC, @rc2, EDGE_SUNKEN, BF_BOTTOMRIGHT)
               end if
            END IF
         end if

         Function = True : Exit Function
   
      Case ODA_FOCUS
         DrawFocusRect(lpdis->hDC, @lpdis->rcItem) 
   
   End Select

   Function = 0
End Function

   
' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmVDToolbox
' ========================================================================================
Function frmVDToolbox_OnDestroy(HWnd As HWnd) As LRESULT
   HWND_FRMVDTOOLBOX = 0
   Function = 0
End Function

' ========================================================================================
' Processes messages for the subclassed ListBox window.
' ========================================================================================
Function frmVDToolbox_ListBox_SubclassProc ( _
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
               SendMessage( GetParent(HWnd), WM_COMMAND, MAKEWPARAM(IDC_FRMFNLIST_LISTBOX, LBN_DBLCLK), Cast(LPARAM,HWnd) )
               Exit Function
            Case VK_ESCAPE
               SendMessage( GetParent(hWnd), WM_CLOSE, 0, 0 ) 
         End Select

      
      case WM_LBUTTONDOWN
         ' Hit test to see if the mouse is over the vertical splitter bar
         if hwnd = GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTPROPERTIES) THEN
            dim as long hPos = loword(lParam)
            dim as RECT rc: GetClientRect(hwnd, @rc)
            if (hPos >= gPropDivPos - 2) andalso (hPos <= gPropDivPos + 2) THEN
               SetCursor( LoadCursor(0, ByVal IDC_SIZEWE) )
               'InvertLine(hwnd, gPropDivPos, rc.top, gPropDivPos, rc.bottom)
               gPropDivTracking = true
               SetCapture(hwnd)
            else
               SetCursor( LoadCursor(0, ByVal IDC_ARROW) )
            END IF
         else
            SetCursor( LoadCursor(0, ByVal IDC_ARROW) )
         END IF
         
      case WM_LBUTTONUP
         if gPropDivTracking = true THEN
            gPropDivTracking = false
            ReleaseCapture
            AfxRedrawWindow(hwnd)
         end if
         
      case WM_MOUSEMOVE
         ' Hit test to see if the mouse is over the vertical splitter bar
         dim as long hPos = loword(lParam)
         dim as RECT rc: GetClientRect(hwnd, @rc)
         if (hPos >= gPropDivPos - 2) andalso (hPos <= gPropDivPos + 2) THEN
            SetCursor( LoadCursor(0, ByVal IDC_SIZEWE) )
            if gPropDivTracking = true THEN
               gPropDivPos = loword(lParam)
               AfxRedrawWindow(hwnd)
            end if
         else
            SetCursor( LoadCursor(0, ByVal IDC_ARROW) )
         END IF
         
      Case WM_DESTROY
         ' REQUIRED: Remove control subclassing
         RemoveWindowSubclass HWnd, @frmFnList_ListBox_SubclassProc, uIdSubclass

   End Select

   ' Default processing of Windows messages
   Function = DefSubclassProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmVDToolbox Window procedure
' ========================================================================================
Function frmVDToolbox_WndProc( ByVal HWnd   As HWnd, _
                               ByVal uMsg   As UINT, _
                               ByVal wParam As WPARAM, _
                               ByVal lParam As LPARAM _
                               ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_COMMAND,     frmVDToolbox_OnCommand)
      HANDLE_MSG (HWnd, WM_NOTIFY,      frmVDToolbox_OnNotify)
      HANDLE_MSG (HWnd, WM_SIZE,        frmVDToolbox_OnSize)
      HANDLE_MSG (HWnd, WM_CLOSE,       frmVDToolbox_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,     frmVDToolbox_OnDestroy)
      HANDLE_MSG (HWnd, WM_MEASUREITEM, frmVDToolbox_OnMeasureItem)
      HANDLE_MSG (HWnd, WM_DRAWITEM,    frmVDToolbox_OnDrawItem)
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmVDToolbox_Show
' ========================================================================================
Function frmVDToolbox_Show( ByVal hWndParent As HWnd, _
                            ByVal nCmdShow   As Long = 0 _
                            ) As Long

   ' If Toolbox already exists then toggle its visibility
   ' unless nCmdShow is explicitly set to show it.
   If IsWindow(HWND_FRMVDTOOLBOX) Then
      If nCmdShow <> SW_SHOW Then nCmdShow = Iif(IsWindowVisible(HWND_FRMVDTOOLBOX), SW_HIDE, SW_SHOW)
      ShowWindow HWND_FRMVDTOOLBOX, nCmdShow
      If nCmdShow = SW_SHOW Then
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
   Dim nHeight As Long = (rcWork.Bottom - rcWork.Top) * .70
   dim nLeft   as long = (rcWork.Right - pWindow->ScaleY(310))
   
   ' We want the Toolbox to display just under the main window top tabcontrol
   Dim nTop As Long 
   dim rc as RECT
   GetWindowRect(gTTabCtl.hWindow, @rc)
   nTop = rc.bottom
   
   HWND_FRMVDTOOLBOX = _
   pWindow->Create( hWndParent, L(352,"Toolbox"), @frmVDToolbox_WndProc, nLeft, nTop, 300, nHeight, _
        WS_POPUP Or WS_CAPTION or WS_SYSMENU or WS_THICKFRAME or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CLIENTEDGE)
   ' We will set our own mouse pointer as needed
   SetClassLongPtr(HWND_FRMVDTOOLBOX, GCLP_HCURSOR, 0)   
   
   ' Set the small and large icon for the main window (must be set after main window is created)
   pWindow->BigIcon   =  LoadImage( pWindow->InstanceHandle, "IMAGE_AAA_MAINICON", IMAGE_ICON, 32, 32, LR_SHARED)
   pWindow->SmallIcon =  LoadImage( pWindow->InstanceHandle, "IMAGE_AAA_MAINICON", IMAGE_ICON, 16, 16, LR_SHARED)

   Dim As HWnd hTabCtl = _ 
        pWindow->AddControl("TAB", , IDC_FRMVDTOOLBOX_TABCONTROL, "", 0, 0, 0, 0)
        
   TabCtrl_AddTab(hTabCtl, 0, L(351,"Tools"))
   TabCtrl_AddTab(hTabCtl, 0, L(350,"Properties"))
   
   Dim As HWnd hList1 = _ 
        pWindow->AddControl("LISTBOX", , IDC_FRMVDTOOLBOX_LSTTOOLBOX, "", 0, 0, 0, 0, _
        WS_CHILD Or WS_VISIBLE or WS_TABSTOP Or LBS_NOINTEGRALHEIGHT Or WS_VSCROLL or  _
        LBS_OWNERDRAWFIXED Or LBS_HASSTRINGS Or LBS_NOTIFY, WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR, , _
        Cast(WNDPROC, @frmVDToolbox_ListBox_SubclassProc), IDC_FRMVDTOOLBOX_LSTTOOLBOX, Cast(DWORD_PTR, @pWindow))

   Dim As HWnd hList2 = _ 
        pWindow->AddControl("LISTBOX", , IDC_FRMVDTOOLBOX_LSTPROPERTIES, "", 0, 0, 0, 0, _
        WS_CHILD Or WS_TABSTOP Or LBS_NOINTEGRALHEIGHT Or WS_VSCROLL or _
        LBS_OWNERDRAWFIXED Or LBS_HASSTRINGS Or LBS_NOTIFY, WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR, , _
        Cast(WNDPROC, @frmVDToolbox_ListBox_SubclassProc), IDC_FRMVDTOOLBOX_LSTPROPERTIES, Cast(DWORD_PTR, @pWindow))
   
   ' Set the default position for the vertical sizing bar
   gPropDivPos = pWindow->ScaleY(0.3 * 300)
   
   ' Add the Tools to the toolbox
   dim as long ndx
   for i as long = lbound(gToolBox) to ubound(gToolBox)
      ndx = ListBox_AddString(hList1, gToolBox(i).wszName)
      ListBox_SetItemData(hList1, ndx, i)
   NEXT
   SetActiveToolboxControl(CTRL_POINTER)
   
   frmVDToolbox_PositionWindows(HWND_FRMVDTOOLBOX)

   TabCtrl_SetCurSel(hTabCtl, 0)
   ShowWindow HWND_FRMVDTOOLBOX, nCmdShow
   ListBox_SetCurSel(hList1, 0)
   SetFocus hList1
   
   Function = 0
End Function



