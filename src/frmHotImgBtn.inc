
''
''  FRMHOTIMGBTN
''
''  Image or Text button with hot mouse over
''
''

Type HOTIMGBTN_TYPE
   hWndParent As HWND
   cID        As Long
   hImgNormal As HICON
   hImgHot    As HICON
   clrBg      As COLORREF
   nImgWidth  As Long
   nImgHeight As Long
   fSelected  As BOOLEAN  ' flag that button should remain displaying as selected
End Type



' ========================================================================================
' Set the Selected falg for the button. If Selected then the button will always 
' display using the Hot image regardless of whether the button is hot or not.
' ========================================================================================
Function frmHotImgBtn_SetSelected( ByVal HWnd As HWnd, ByVal fSelected As BOOLEAN) As Long
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)
   If pWindow = 0 Then Exit Function
   
   Dim pData As HOTIMGBTN_TYPE Ptr = Cast(HOTIMGBTN_TYPE Ptr, pWindow->UserData(0))
   If pData Then pData->fSelected = fSelected
   AfxRedrawWindow( hWnd )
   
   Function = 0
End Function

' ========================================================================================
' Gets the Selected falg from the button. If Selected then the button will always 
' display using the Hot image regardless of whether the button is hot or not.
' ========================================================================================
Function frmHotImgBtn_GetSelected( ByVal HWnd As HWnd ) As BOOLEAN
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)
   If pWindow = 0 Then Exit Function
   
   Dim pData As HOTIMGBTN_TYPE Ptr = Cast(HOTIMGBTN_TYPE Ptr, pWindow->UserData(0))
   If pData Then Return pData->fSelected

   Function = False
End Function


' ========================================================================================
' Process WM_LBUTTONUP message for frmHotImgBtn
' ========================================================================================
Function frmHotImgBtn_OnLButtonUp(ByVal HWnd As HWnd, ByVal x As Long, ByVal y As Long, ByVal keyFlags As Long ) As LRESULT
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)
   If pWindow = 0 Then Exit Function
   
   Dim pData As HOTIMGBTN_TYPE Ptr = Cast(HOTIMGBTN_TYPE Ptr, pWindow->UserData(0))

   ' Post a clicked message to the parent's WM_COMMAND handler.
   If pData Then
      PostMessageW( pData->hWndParent, WM_COMMAND, MAKEWPARAM(pData->cID, BN_CLICKED), Cast(LPARAM, hWnd) )
   End If
   
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmHotImgBtn
' ========================================================================================
Function frmHotImgBtn_OnDestroy(HWnd As HWnd) As LRESULT
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)
   If pWindow = 0 Then Exit Function
   
   Dim pData As HOTIMGBTN_TYPE Ptr = Cast(HOTIMGBTN_TYPE Ptr, pWindow->UserData(0))
   If pData Then
      DestroyIcon(pData->hImgNormal)
      DestroyIcon(pData->hImgHot)
      Delete pData
   End If
      
   Function = 0
End Function


' ========================================================================================
' Process WM_PAINT message for window/dialog: frmHotImgBtn
' ========================================================================================
Function frmHotImgBtn_OnPaint( ByVal HWnd As HWnd) As LRESULT

   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)
   If pWindow = 0 Then Exit Function
   
   Dim As PAINTSTRUCT ps
   Dim As HBITMAP hbit 
   
   Dim As HDC hDC, memDC 
   Dim As HBRUSH hBrush
   Dim As POINT pt
   Dim As RECT rc
   Dim As Long nHMargin, nVMargin
   Dim As HICON hIcon
   Dim As BOOLEAN bIsHot

   hDC = BeginPaint(hWnd, @ps)

   SaveDC hDC
   
   GetClientRect hWnd, @rc

   memDC = CreateCompatibleDC( ps.hDC )
   hbit  = CreateCompatibleBitmap( ps.hDC, rc.Right, rc.Bottom )
   If hbit Then hbit = SelectObject( memDC, hbit )

   Dim pData As HOTIMGBTN_TYPE Ptr = Cast(HOTIMGBTN_TYPE Ptr, pWindow->UserData(0))
   
   If pData Then
      bIsHot = IIf( GetPropW(hWnd, "HOT") = 0, False, True ) 

      ' Paint the entire client background
      hBrush = CreateSolidBrush( pData->clrBg )
      FillRect(memDC, @rc, hBrush) 
      DeleteBrush(hBrush)

      ' Center the image
      nHMargin = (rc.right - pData->nImgWidth) / 2
      nVMargin = (rc.bottom - pData->nImgHeight) / 2
      
      If ( bIsHot = True ) Or (pData->fSelected = True) Then
         hIcon = pData->hImgHot
      Else
         hIcon = pData->hImgNormal
      End If
      DrawStateW memDC, NULL, NULL, CAST(LPARAM, hIcon), 0, _
                  rc.Left + nHMargin, rc.Top + nVMargin, 0, 0, DST_ICON Or DSS_NORMAL
       
   End If
   
   ' Copy the entire memory bitmap over to the visual display
   BitBlt ps.hDC, 0, 0, rc.Right, rc.Bottom, memDC, 0, 0, SRCCOPY 

   ' Cleanup
   If hbit  Then DeleteObject SelectObject(memDC, hbit)
   If memDC Then DeleteDC memDC

   RestoreDC hDC, -1 
   
   EndPaint hWnd, @ps
   
   Function = 0
End Function


' ========================================================================================
' frmHotImgBtn Window procedure
' ========================================================================================
Function frmHotImgBtn_WndProc( ByVal HWnd   As HWnd, _
                               ByVal uMsg   As UINT, _
                               ByVal wParam As WPARAM, _
                               ByVal lParam As LPARAM _
                               ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_LBUTTONUP,   frmHotImgBtn_OnLButtonUp)
      HANDLE_MSG (HWnd, WM_PAINT,       frmHotImgBtn_OnPaint)
      HANDLE_MSG (HWnd, WM_DESTROY,     frmHotImgBtn_OnDestroy)

      CASE WM_MOUSEMOVE
         ' Need to record trackmouse event in order to get the WM_MOUSELEAVE message 
         ' when the mouse leaves the client area of the tabctl.
         DIM trackMouse AS TRACKMOUSEEVENT
         IF GetPropW(hwnd, "HOT") = 0 THEN
            trackMouse.cbSize = SIZEOF(trackMouse)
            trackMouse.dwFlags = TME_LEAVE
            trackMouse.hwndTrack = hwnd
            trackMouse.dwHoverTime = 1
            TrackMouseEvent(@trackMouse)
            SetPropW hwnd, "HOT", CAST(HANDLE, CTRUE)
            ' Ensure repainting as hot
            AfxRedrawWindow hWnd
         END IF
         Exit Function
         
      CASE WM_MOUSELEAVE
         ' Repaint to remove any Hot state tab
         RemovePropW hwnd, "HOT"
         AfxRedrawWindow hWnd
         Exit Function

      Case WM_ERASEBKGND
         Return True  ' prevent background erase using class brush

   End Select

   ' for messages that we don't deal with
   Function = DefWindowProcW(HWnd, uMsg, wParam, lParam)

End Function



' ========================================================================================
' frmHotImgBtn
' ========================================================================================
Function frmHotImgBtn( ByVal hWndParent   As HWnd, _
                       ByVal cID          As Long = 0, _
                       ByRef wszImgNormal As WString = "", _
                       ByRef wszImgHot    As WString = "",  _
                       ByRef wszTooltip   As WString = "",  _
                       ByVal nImgWidth    As Long = 24, _
                       ByVal nImgHeight   As Long = 24, _
                       ByVal clrBg        As COLORREF = 0, _
                       BYVAL x            AS LONG = 0, _
                       BYVAL y            AS LONG = 0, _
                       BYVAL nWidth       AS LONG = 0, _
                       BYVAL nHeight      AS LONG = 0 _
                       ) As HWND

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   Dim As HWND hWindow = _
   pWindow->Create( hWndParent, "", @frmHotImgBtn_WndProc, x, y, nWidth, nHeight, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->ClassStyle = 0  ' to ensure that there are no double clciks
   pWindow->Brush = Null
   
   SetWindowLongPtr( hWindow, GWLP_ID, cID )
   
   Dim pData As HOTIMGBTN_TYPE Ptr = New HOTIMGBTN_TYPE
   With *pData
      .hWndParent = hWndParent
      .cID        = cID
      .nImgWidth  = pWindow->ScaleX( nImgWidth )
      .nImgHeight = pWindow->ScaleY( nImgHeight )
      .hImgNormal = LoadImageW( pWindow->InstanceHandle, wszImgNormal, IMAGE_ICON, .nImgWidth, .nImgHeight, LR_DEFAULTCOLOR)
      .hImgHot    = LoadImageW( pWindow->InstanceHandle, wszImgHot, IMAGE_ICON, .nImgWidth, .nImgHeight, LR_DEFAULTCOLOR)
      .clrBg      = clrBg
   End With
   
   ' If the images failed to load then try loading them using GDI in cases of png format.
   If pData->hImgNormal = 0 Then 
      pData->hImgNormal = AfxGdipIconFromRes( pWindow->InstanceHandle, wszImgNormal )
   End If
   If pData->hImgHot = 0 Then 
      pData->hImgHot = AfxGdipIconFromRes( pWindow->InstanceHandle, wszImgHot )
   End If

   ' Save the control data elements
   pWindow->UserData(0) = Cast(LONG_PTR, pData)
   
   ' Register a tooltip for this window. Change the text of the tooltip to the full filename/path as
   ' the mouse changes from tab to tab.
   Dim As Hwnd hTooltip
   If Len(wszTooltip) Then
      hTooltip = AfxAddTooltip( hWindow, AfxStrRemove(wszTooltip, "&") )
      Tooltip_SetDelayTime( hTooltip, TTDT_INITIAL, 250 )
      Tooltip_SetDelayTime( hTooltip, TTDT_RESHOW, 1000 )
      Tooltip_SetDelayTime( hTooltip, TTDT_AUTOPOP, 2000 )
   End If
   
   ' Refresh the control
   AfxRedrawWindow(hWindow)
   
   Function = hWindow
   
End Function














