

' ========================================================================================
' Process WM_LBUTTONUP message for frmVScrollBar
' ========================================================================================
Function frmVScrollBar_OnLButtonUp(ByVal HWnd As HWnd, ByVal x As Long, ByVal y As Long, ByVal keyFlags As Long ) As LRESULT

   If gScroll.bScrollActive Then
      gScroll.bScrollActive = False
      gScroll.ptLastPos.y = 0 
      ReleaseCapture()
   End If
   
   Function = 0
End Function


' ========================================================================================
' Process WM_MOUSEMOVE message for frmVScrollBar
' ========================================================================================
Function frmVScrollBar_OnMouseMove(ByVal HWnd As HWnd, ByVal x As Long, ByVal y As Long, ByVal keyFlags As Long ) As LRESULT

   If gScroll.bScrollActive Then
      ' Get the current rect of the splitterbar thumb and compare it to the current cursor position.
      Dim As Long nDiff = (y - gScroll.ptLastPos.y) 
      If nDiff = 0 Then Exit Function
      
'      If Abs(nDiff) >= 3 Then
'         gScroll.ptLastPos.y = y
'         Dim pDoc As clsDocument Ptr = gApp.GetActivedoc()
'         If pDoc Then SendMessageW( pDoc->hWindow, SCI_LINESCROLL, 0, Cast(LPARAM, IIF(nDiff > 0, 1, -1)) ) 
'      End If
'SendMessage TV.hWnd, TVM_SELECTITEM, TVGN_FIRSTVISIBLE, ByVal hItem
   End If
   
   AfxRedrawWindow(hWnd)
   
   Function = 0
End Function


' ========================================================================================
' Process WM_LBUTTONDOWN message for window/dialog: frmVScrollBar
' ========================================================================================
Function frmVScrollBar_OnLButtonDown(ByVal HWnd As HWnd, ByVal fDoubleClick As BOOLEAN, _
                                     ByVal x As Long, ByVal y As Long, ByVal keyFlags As Long ) As LRESULT

   DIM As POINT pt
   
   GetCursorPos(@pt)
   ScreenToClient(hWnd, @pt)

   ' Determine if we have clicked on the Thumb or the client area. Clicking on the 
   If PtInRect(@gScroll.rcThumb, pt) Then
      gScroll.bScrollActive = True
      gScroll.ptLastPos = pt 
      SetCapture( HWnd )
   Else
'      Dim pDoc As clsDocument Ptr = gApp.GetActiveDoc()
'      Dim As HWnd hEdit 
'      Dim As Long nPageLines, nFirstLine
'      If pDoc Then
'         hEdit = pDoc->hWindow
'         nPageLines = SendMessageW( hEdit, SCI_LINESONSCREEN, 0, 0) 
'         ' The client area was clicked. Determine if above or below the thumb.
'         ' Scroll the Scintilla window up or down one page at a time
'         If pt.y < gScroll.rcThumb.top Then
'            SendMessageW( hEdit, SCI_LINESCROLL, 0, Cast(LPARAM, -nPageLines)) 
'         Else
'            SendMessageW( hEdit, SCI_LINESCROLL, 0, Cast(LPARAM, nPageLines)) 
'         End If
'      End If
   End If
   
   Function = 0
End Function


' ========================================================================================
' Process WM_PAINT message for window/dialog: frmVScrollBar
' ========================================================================================
Function frmVScrollBar_OnPaint( ByVal HWnd As HWnd) As LRESULT

   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)
   If pWindow = 0 Then Exit Function
   
   ' Paint the main application background
   Dim As PAINTSTRUCT ps
   Dim As HBITMAP hbit 
   
   Dim As HDC hDC, memDC 
   Dim As HBRUSH hBrush
   Dim As POINT pt
   Dim As RECT rc
   Dim As Long nLeft, nTop, nMargin

   hDC = BeginPaint(hWnd, @ps)

   SaveDC hDC
   
   GetCursorPos(@pt)
   ScreenToClient(hWnd, @pt)
   
   GetClientRect hWnd, @rc

   memDC = CreateCompatibleDC( ps.hDC )
   hbit  = CreateCompatibleBitmap( ps.hDC, rc.Right, rc.Bottom )
   If hbit Then hbit = SelectObject( memDC, hbit )

   ' Paint the entire client background
   hBrush = CreateSolidBrush( BGR(80,80,80) )
   FillRect(memDC, @rc, hBrush) 
   DeleteBrush(hBrush)

   Dim As HWnd hTV = GetDlgItem( HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE ) 
   Dim As Long nFirstLine, nPageLines, nThumbHeight
   nFirstLine = 1
   nPageLines = Treeview_GetVisibleCount(hTV)   ' number of nodes that can fit on the screen

   dim hItem as HTREEITEM
   dim nCountVisible as long
   
   ' Determine the total number of nodes that are expanded and visible
   hItem = Treeview_GetFirstVisible(hTV)
   do until hItem = null
      nCountVisible = nCountVisible + 1
      hItem = Treeview_GetNextVisible(hTV, hItem)
   loop   

   ' Do not create a scrollbar if the current number of visible nodes
   ' equals the total number of nodes in the treeview, therefore no
   ' need to scroll the treeview.
   If nPageLines < nCountVisible then
      ' Paint the scrollbar thumb
      nThumbHeight = (nPageLines / (nCountVisible - nPageLines))  * rc.bottom  
      ' Ensure our thumb is not too small for practical purposes.
      If nThumbHeight < pWindow->ScaleY(16) Then
         nthumbHeight = pWindow->ScaleY(16)
      End If   
      gScroll.rcThumb.left  = rc.left + 1
      gScroll.rcThumb.right = rc.right - 1
      gScroll.rcThumb.top   = (rc.bottom - rc.top) * (nFirstLine / nCountVisible)
      gScroll.rcThumb.bottom = gScroll.rcThumb.top + nThumbHeight
      ' Don't let the thumb extend past the bottom of the client area
      If gScroll.rcThumb.bottom > rc.bottom Then 
         gScroll.rcThumb.bottom = rc.bottom
         gScroll.rcThumb.top = rc.bottom - nThumbHeight
      End If   
      hBrush = CreateSolidBrush( BGR(66,66,66) )
      FillRect(memDC, @gScroll.rcThumb, hBrush) 
      DeleteBrush(hBrush)
   end if
   
   ' Copy the entire memory bitmap over to the visual display
   BitBlt ps.hDC, 0, 0, rc.Right, rc.Bottom, memDC, 0, 0, SRCCOPY 

   ' Cleanup
   If hbit  Then DeleteObject SelectObject(memDC, hbit)
   If memDC Then DeleteDC memDC

   RestoreDC hDC, -1 

   DeleteBrush(hBrush)
   
   EndPaint hWnd, @ps
   
   Function = 0
End Function


' ========================================================================================
' frmVScrollBar Window procedure
' ========================================================================================
Function frmVScrollBar_WndProc( ByVal HWnd   As HWnd, _
                             ByVal uMsg   As UINT, _
                             ByVal wParam As WPARAM, _
                             ByVal lParam As LPARAM _
                             ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_LBUTTONDOWN, frmVScrollBar_OnLButtonDown)
      HANDLE_MSG (HWnd, WM_LBUTTONUP,   frmVScrollBar_OnLButtonUp)
      HANDLE_MSG (HWnd, WM_MOUSEMOVE,   frmVScrollBar_OnMouseMove)
      HANDLE_MSG (HWnd, WM_PAINT,       frmVScrollBar_OnPaint)

      Case WM_ERASEBKGND
         Return True  ' prevent background erase using class brush

      Case WM_MOUSELEAVE
         RemovePropW HWND_FRMVSCROLLBAR, "HOT"
         ShowWindow HWND_FRMVSCROLLBAR, SW_HIDE
         PositionExplorerWindows(HWND_FRMEXPLORER)

'      Case MSG_USER_VSCROLL_UPDATE
'         AfxRedrawWindow( Hwnd )    ' handle calculations in WM_PAINT
         
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProcW(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmVScrollBar_Show
' ========================================================================================
Function frmVScrollBar_Show( ByVal hWndParent As HWnd ) As LRESULT

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   HWND_FRMVSCROLLBAR = pWindow->Create( hWndParent, "", @frmVScrollBar_WndProc, 0, 0, 14, 0, _
        WS_CHILD Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->ClassStyle = 0  ' to ensure that there are no double clciks
   pWindow->Brush = Null   ' we do the painting ourselves

   Function = 0
   
End Function























