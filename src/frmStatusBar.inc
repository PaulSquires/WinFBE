'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2022 Paul Squires, PlanetSquires Software
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
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
Function frmStatusBar_PositionWindows() As LRESULT

   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWND_FRMMAIN_STATUSBAR)
   If pWindow = 0 Then Exit Function
   
   ' Get the entire client area
   Dim As Rect rc
   GetClientRect(HWND_FRMMAIN_STATUSBAR, @rc)

   dim as long i = 0
   dim as long nLeft = rc.left
   dim as long nSpacer = pWindow->ScaleX(4)
   
   ' the text for each panel is set in frmMain_SetStatusbar which then calls this
   ' function to determine proper sizing and placement.
   
   '''' LEFT ALIGNED STATUSBAR PANELS ''''
   i = 0  ' Ln 23, Col 20
   gSBPanels(i).rc.Left   = nLeft
   gSBPanels(i).rc.top    = rc.top
   gSBPanels(i).rc.right  = nLeft + _
      pWindow->ScaleX( getTextWidth( HWND_FRMMAIN_STATUSBAR, gSBPanels(i).wszText, ghStatusBar.hFontStatusBar, 4 ) )
   gSBPanels(i).rc.bottom = rc.bottom
   gSBPanels(i).nID       = IDM_GOTO
   nLeft = gSBPanels(i).rc.right + nSpacer

   '''' RIGHT ALIGNED STATUSBAR PANELS ''''
   grcGripper = rc
   grcGripper.top = (rc.bottom-rc.top) / 2
   grcGripper.left = rc.right - (nSpacer * 4)
   
   i = 6  ' CRLF
   dim as long nRight = grcGripper.left
   gSBPanels(i).rc.Left   = nRight - _
      pWindow->ScaleX( getTextWidth( HWND_FRMMAIN_STATUSBAR, gSBPanels(i).wszText, ghStatusBar.hFontStatusBar, 4 ) )
   gSBPanels(i).rc.top    = rc.top
   gSBPanels(i).rc.right  = nRight
   gSBPanels(i).rc.bottom = rc.bottom
   gSBPanels(i).nID       = IDM_LINEENDINGS
   nRight = gSBPanels(i).rc.left - nSpacer

   i -= 1  ' UTF-8
   gSBPanels(i).rc.Left   = nRight - _
      pWindow->ScaleX( getTextWidth( HWND_FRMMAIN_STATUSBAR, gSBPanels(i).wszText, ghStatusBar.hFontStatusBar, 4 ) )
   gSBPanels(i).rc.top    = rc.top
   gSBPanels(i).rc.right  = nRight 
   gSBPanels(i).rc.bottom = rc.bottom
   gSBPanels(i).nID       = IDM_FILEENCODING
   nRight = gSBPanels(i).rc.left - nSpacer

   i -= 1  ' Spaces: 3
   gSBPanels(i).rc.Left   = nRight - _
      pWindow->ScaleX( getTextWidth( HWND_FRMMAIN_STATUSBAR, gSBPanels(i).wszText, ghStatusBar.hFontStatusBar, 4 ) )
   gSBPanels(i).rc.top    = rc.top
   gSBPanels(i).rc.right  = nRight
   gSBPanels(i).rc.bottom = rc.bottom
   gSBPanels(i).nID       = IDM_SPACES
   nRight = gSBPanels(i).rc.left - nSpacer

   i -= 1  ' Main/Resource/Module/Normal
   if gApp.IsProjectActive then
      gSBPanels(i).rc.Left   = nRight - _
         pWindow->ScaleX( getTextWidth( HWND_FRMMAIN_STATUSBAR, gSBPanels(i).wszText, ghStatusBar.hFontStatusBar, 4 ) )
      gSBPanels(i).rc.top    = rc.top
      gSBPanels(i).rc.right  = nRight
      gSBPanels(i).rc.bottom = rc.bottom
      gSBPanels(i).nID       = IDM_PROJECTFILETYPE
      nRight = gSBPanels(i).rc.left - nSpacer
   else
      SetRectEmpty( @gSBPanels(i).rc )
   end if
   
   i -= 1  ' Build Configuration
   gSBPanels(i).rc.Left   = nRight - _
      pWindow->ScaleX( getTextWidth( HWND_FRMMAIN_STATUSBAR, gSBPanels(i).wszText, ghStatusBar.hFontStatusBar, 4 ) )
   gSBPanels(i).rc.top    = rc.top
   gSBPanels(i).rc.right  = nRight
   gSBPanels(i).rc.bottom = rc.bottom
   gSBPanels(i).nID       = IDM_BUILDCONFIG
   nRight = gSBPanels(i).rc.left - nSpacer

   '''' THE MESSAGES PANEL FILLS THE REMAINING AVAILABLE SPACE ''''
   i = 1  ' Compile status and error messages
   gSBPanels(i).rc.Left   = gSBPanels(0).rc.right
   gSBPanels(i).rc.top    = rc.top
   gSBPanels(i).rc.right  = gSBPanels(2).rc.Left
   gSBPanels(i).rc.bottom = rc.bottom
   gSBPanels(i).nID       = IDM_VIEWOUTPUT

   AfxRedrawWindow( HWND_FRMMAIN_STATUSBAR )
   Function = 0
End Function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmStatusBar
' ========================================================================================
Function frmStatusBar_OnSize( _
            ByVal HWnd As HWnd, _
            ByVal state As UINT, _
            ByVal cx As Long, _
            ByVal cy As Long _
            ) As LRESULT
   If state <> SIZE_MINIMIZED Then
      frmStatusBar_PositionWindows()
   End If
   Function = 0
End Function


' ========================================================================================
' frmStatusBar Window procedure
' ========================================================================================
Function frmStatusBar_WndProc( _
            ByVal HWnd   As HWnd, _
            ByVal uMsg   As UINT, _
            ByVal wParam As WPARAM, _
            ByVal lParam As LPARAM _
            ) As LRESULT

   static as long prev_idxHot = -1
   static as long curr_idxHot = -1
   static as boolean resizing
   static as POINT prev_pt
   
   Select Case uMsg
      HANDLE_MSG (HWnd, WM_SIZE, frmStatusBar_OnSize)
   
   case WM_DESTROY
      DIM pWindow AS CWindow PTR = AfxCWindowPtr(HWnd)
      If pWindow = 0 Then Delete pWindow
      
   case WM_ERASEBKGND
      return true

   case WM_LBUTTONDOWN
      ' Are we over the resize gripper
'      dim as POINT pt: GetCursorPos( @pt )
'      MapWindowPoints( HWND_DESKTOP, HWnd, cast( POINT ptr, @pt ), 1 )
'      if PtInRect( @grcGripper, pt ) then
'         SetCursor LoadImage( Null, MAKEINTRESOURCEW(OCR_SIZENWSE), IMAGE_CURSOR, 0, 0, LR_DEFAULTSIZE or LR_SHARED )
'         resizing = true
'         prev_pt = pt
         SetCapture( HWnd )
'      end if
      
   case WM_MOUSEMOVE
      Dim tme As TrackMouseEvent
      tme.cbSize = Sizeof(TrackMouseEvent)
      tme.dwFlags = TME_HOVER Or TME_LEAVE
      tme.hwndTrack = HWnd
      TrackMouseEvent(@tme) 

      dim as POINT pt: GetCursorPos( @pt )
      MapWindowPoints( HWND_DESKTOP, HWnd, cast( POINT ptr, @pt ), 1 )
      
      ' Are we over the resize gripper
'      if PtInRect( @grcGripper, pt ) then
'         SetCursor LoadImage( Null, MAKEINTRESOURCEW(OCR_SIZENWSE), IMAGE_CURSOR, 0, 0, LR_DEFAULTSIZE or LR_SHARED )
'      else
'         SetCursor( LoadCursor( null, IDC_ARROW ))
'      end if
      if resizing then
'         dim as long xDelta = prev_pt.x - pt.x 
'         dim as long yDelta = prev_pt.y - pt.y
'         ' Resize the main form
'         dim as RECT rcMain = AfxGetWindowRect( HWND_FRMMAIN )
'         dim as long nWidth = rcMain.right - rcMain.left - xDelta
'         dim as long nHeight = rcMain.bottom - rcMain.top - yDelta
'         SetWindowPos( HWND_FRMMAIN, 0, 0, 0, nWidth, nHeight, SWP_NOMOVE or SWP_NOZORDER )
      else
         ' Are we over one of the status bar panels
         curr_idxHot = -1
         for i as long = lbound(gSBPanels) to ubound(gSBPanels)
            if PtInRect( @gSBPanels(i).rc, pt ) then
               ' Only make hot panels that have actual text
               if len(gSBPanels(i).wszText) then
                  curr_idxHot = i
                  gSBPanels(i).isHot = true
               end if   
            else   
               gSBPanels(i).isHot = false
            end if
         next
         if curr_idxHot <> prev_idxHot then
            AfxRedrawWindow( HWnd )
            prev_idxHot = curr_idxHot 
         end if
      end if
            
   case WM_MOUSELEAVE
      for i as long = lbound(gSBPanels) to ubound(gSBPanels)
         gSBPanels(i).isHot = false
      next
      prev_idxHot = -1
      curr_idxHot = -1
      AfxRedrawWindow( HWnd )

   case WM_LBUTTONUP
      dim as POINT pt: GetCursorPos( @pt )
      MapWindowPoints( HWND_DESKTOP, HWnd, cast( POINT ptr, @pt ), 1 )
      SetCursor( LoadCursor( null, IDC_ARROW ))
      ReleaseCapture
      resizing = false
      ' Are we over one of the status bar panels
      for i as long = lbound(gSBPanels) to ubound(gSBPanels)
         if PtInRect( @gSBPanels(i).rc, pt ) then
            PostMessage( HWND_FRMMAIN, WM_COMMAND, MAKEWPARAM(gSBPanels(i).nID, 0), 0 )
            exit for
         end if
      next   

   case WM_PAINT
      Dim As PAINTSTRUCT ps
      Dim As HDC hDc

      Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWND_FRMMAIN_STATUSBAR)
      If pWindow = 0 Then Exit Function

      hDC = BeginPaint(hWnd, @ps)

      SaveDC(hDC)
      dim as long nWidth = ps.rcPaint.right - ps.rcPaint.left
      dim as long nHeight = ps.rcPaint.bottom - ps.rcPaint.top

      Dim memDC as HDC      ' Double buffering
      Dim hbit As HBITMAP   ' Double buffering

      memDC = CreateCompatibleDC( hDC )
      hbit  = CreateCompatibleBitmap( hDC, nWidth, nHeight )
      
      SaveDC(memDC)
      SelectObject( memDC, hbit )

      if ghStatusBar.hFontStatusBar then SelectObject(memDC, ghStatusBar.hFontStatusBar)

      FillRect( memDC, @ps.rcPaint, ghStatusBar.hPanelBrush )
      
      ' paint the panels
      for i as long = lbound(gSBPanels) to ubound(gSBPanels)
         SetBkColor( memDC, iif( gSBPanels(i).isHot, ghStatusBar.BackColorHot, ghStatusBar.BackColor) )   
         SetTextColor( memDC, iif( gSBPanels(i).isHot, ghStatusBar.ForeColorHot, ghStatusBar.ForeColor) )
         FillRect( memDC, @gSBPanels(i).rc, iif( gSBPanels(i).isHot, ghStatusBar.hBackBrushHot, ghStatusBar.hBackBrush) )
         dim as long wsStyle = DT_NOPREFIX or DT_CENTER Or DT_VCENTER or DT_SINGLELINE
         DrawText( memDC, gSBPanels(i).wszText.sptr, -1, Cast(lpRect, @gSBPanels(i).rc), wsStyle )
         
         ' if this is the compile result panel then display the green/success or red/failure "icon"
         if i = 1 then
            if (gApp.IsProjectLoading = false) andalso (gApp.IsFileLoading = false) then
               dim as COLORREF clr = ghStatusBar.BackColor
               if gApp.hIconPanel = ghIconGood then clr = ghGeneral.iconsuccess
               if gApp.hIconPanel = ghIconBad then clr = ghGeneral.iconfail
               dim as long txtWidth = getTextWidth( HWND_FRMMAIN_STATUSBAR, gSBPanels(i).wszText, ghStatusBar.hFontStatusBar, 0 ) 
               SetTextColor( memDC, clr )
               dim as RECT rcBitmap = gSBPanels(i).rc
               dim as long halfWidth = (rcBitmap.right-rcBitmap.left) / 2
               rcBitmap.right = rcBitmap.left + halfWidth - pWindow->ScaleX((txtWidth / 2))
               rcBitmap.left = rcBitmap.right - pWindow->ScaleX(16)

               dim as long wsStyle = DT_NOPREFIX or DT_CENTER Or DT_VCENTER or DT_SINGLELINE
               DrawText( memDC, wszCompileResultIcon, -1, Cast(lpRect, @rcBitmap), wsStyle )
            end if
         end if   
      next

      BitBlt hDC, 0, 0, nWidth, nHeight, memDC, 0, 0, SRCCOPY 

      ' Cleanup
      RestoreDC(memDC, -1)
      If hbit  Then DeleteObject(hbit)
      If memDC Then DeleteDC(memDC)
      RestoreDC(hDC, -1)

      EndPaint hWnd, @ps
   
   End Select
   
   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmStatusBar_Show
' ========================================================================================
Function frmStatusBar_Show( ByVal hWndParent As HWnd ) As LRESULT

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   HWND_FRMMAIN_STATUSBAR = pWindow->Create( hWndParent, "", @frmStatusBar_WndProc, _
        0, 0, 0, STATUSBAR_HEIGHT, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   ' Disable background erasing by only assigning the one style
   pWindow->ClassStyle = CS_DBLCLKS
   
   frmStatusBar_PositionWindows()
   Function = 0
   
End Function

