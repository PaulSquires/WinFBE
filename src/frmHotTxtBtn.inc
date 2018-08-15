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


''
''  FRMHOTTXTBTN
''
''  Image or Text button with hot mouse over
''
''

Type HOTTXTBTN_TYPE
   hWndParent   As HWND
   cID          As Long
   wzText       As WString * MAX_PATH
   clrFg        As COLORREF
   clrBg        As COLORREF
   clrFgHot     As COLORREF
   clrBgHot     As COLORREF
   nOrientation As Long
End Type


' ========================================================================================
' Process WM_LBUTTONUP message for frmHotTxtBtn
' ========================================================================================
Function frmHotTxtBtn_OnLButtonUp(ByVal HWnd As HWnd, ByVal x As Long, ByVal y As Long, ByVal keyFlags As Long ) As LRESULT
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)
   If pWindow = 0 Then Exit Function
   
   Dim pData As HOTTXTBTN_TYPE Ptr = Cast(HOTTXTBTN_TYPE Ptr, pWindow->UserData(0))

   ' Post a clicked message to the parent's WM_COMMAND handler.
   If pData Then
      PostMessageW( pData->hWndParent, WM_COMMAND, MAKEWPARAM(pData->cID, BN_CLICKED), Cast(LPARAM, hWnd) )
   End If
   
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmHotTxtBtn
' ========================================================================================
Function frmHotTxtBtn_OnDestroy(HWnd As HWnd) As LRESULT
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)
   If pWindow = 0 Then Exit Function
   
   Dim pData As HOTTXTBTN_TYPE Ptr = Cast(HOTTXTBTN_TYPE Ptr, pWindow->UserData(0))
   If pData Then delete pData
      
   Function = 0
End Function


' ========================================================================================
' Process WM_PAINT message for window/dialog: frmHotTxtBtn
' ========================================================================================
Function frmHotTxtBtn_OnPaint( ByVal HWnd As HWnd) As LRESULT

   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)
   If pWindow = 0 Then Exit Function
   
   Dim As PAINTSTRUCT ps
   Dim As HBITMAP hbit 
   
   Dim As HDC hDC, memDC 
   Dim As HBRUSH hBrush
   Dim As POINT pt
   Dim As RECT rc
   Dim As BOOLEAN bIsHot
   Dim As Long nOrientation
   
   hDC = BeginPaint(hWnd, @ps)

   SaveDC hDC
   
   GetClientRect hWnd, @rc

   memDC = CreateCompatibleDC( ps.hDC )
   hbit  = CreateCompatibleBitmap( ps.hDC, rc.Right, rc.Bottom )
   If hbit Then hbit = SelectObject( memDC, hbit )

   Dim pData As HOTTXTBTN_TYPE Ptr = Cast(HOTTXTBTN_TYPE Ptr, pWindow->UserData(0))
   
   If pData Then
      bIsHot = IIf( GetPropW(hWnd, "HOT") = 0, False, True ) 

      ' Paint the entire client background
      hBrush = CreateSolidBrush( Iif(bIsHot, pData->clrBgHot, pData->clrBg) )
      FillRect(memDC, @rc, hBrush) 
      DeleteBrush(hBrush)

      SelectFont( memDC, pWindow->Font )
      
      SetBkColor   memDC, Iif(bIsHot, pData->clrBgHot, pData->clrBg)
      SetTextColor memDC, Iif(bIsHot, pData->clrFgHot, pData->clrFg )
      Select Case pData->nOrientation 
         Case 0: nOrientation = DT_LEFT
         Case 1: nOrientation = DT_CENTER
         Case 2: nOrientation = DT_RIGHT
         Case Else: nOrientation = DT_CENTER
      End Select         
      DrawTextW memDC, @pData->wzText, Len(pData->wzText), @rc, nOrientation Or DT_SINGLELINE or DT_VCENTER
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
' frmHotTxtBtn Window procedure
' ========================================================================================
Function frmHotTxtBtn_WndProc( ByVal HWnd   As HWnd, _
                               ByVal uMsg   As UINT, _
                               ByVal wParam As WPARAM, _
                               ByVal lParam As LPARAM _
                               ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_LBUTTONUP,   frmHotTxtBtn_OnLButtonUp)
      HANDLE_MSG (HWnd, WM_PAINT,       frmHotTxtBtn_OnPaint)
      HANDLE_MSG (HWnd, WM_DESTROY,     frmHotTxtBtn_OnDestroy)

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
' frmHotTxtBtn
' ========================================================================================
Function frmHotTxtBtn( ByVal hWndParent   As HWnd, _
                       ByVal cID          As Long = 0, _
                       ByRef wszText      As WString = "", _
                       ByRef wszTooltip   As WString = "", _
                       ByVal clrFg        As COLORREF = 0, _
                       ByVal clrBg        As COLORREF = 0, _
                       ByVal clrFgHot     As COLORREF = 0, _
                       ByVal clrBgHot     As COLORREF = 0, _
                       ByVal nOrientation As Long = 1, _            ' 0 Left, 1 Center, 2 Right
                       BYVAL x            AS LONG = 0, _
                       BYVAL y            AS LONG = 0, _
                       BYVAL nWidth       AS LONG = 0, _
                       BYVAL nHeight      AS LONG = 0 _
                       ) As HWND

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   Dim As HWND hWindow = _
   pWindow->Create( hWndParent, "", @frmHotTxtBtn_WndProc, x, y, nWidth, nHeight, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->ClassStyle = 0  ' to ensure that there are no double clciks
   pWindow->Brush = Null
   pWindow->SetFont( "Segoe UI Semibold", 10)
   
   SetWindowLongPtr( hWindow, GWLP_ID, cID )
   
   Dim pData As HOTTXTBTN_TYPE Ptr = New HOTTXTBTN_TYPE
   With *pData
      .hWndParent = hWndParent
      .cID        = cID
      .wzText     = AfxStrRemove(wszText, "&")
      .clrFg      = clrFg
      .clrBg      = clrBg
      .clrFgHot   = clrFgHot
      .clrBgHot   = clrBgHot
      .nOrientation = nOrientation
   End With
   
   ' Save the control data elements
   pWindow->UserData(0) = Cast(LONG_PTR, pData)
   
   ' Register a tooltip for this window. Change the text of the tooltip to the full filename/path as
   ' the mouse changes from tab to tab.
   Dim As Hwnd hTooltip
   If Len(wszTooltip) Then
      hTooltip = AfxAddTooltip( hWindow, wszTooltip )
      Tooltip_SetDelayTime( hTooltip, TTDT_INITIAL, 250 )
      Tooltip_SetDelayTime( hTooltip, TTDT_RESHOW, 1000 )
      Tooltip_SetDelayTime( hTooltip, TTDT_AUTOPOP, 2000 )
   End If
   
   ' Refresh the control
   AfxRedrawWindow(hWindow)
   
   Function = hWindow
   
End Function












