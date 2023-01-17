'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2023 Paul Squires, PlanetSquires Software
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
''  frmPanelVScroll.inc
''   

#include once "frmPanel.bi"


' ========================================================================================
' Calculate the RECT that holds the client coordinates of the scrollbar's vertical thumb
' Returns True if RECT is not empty
' ========================================================================================
function frmPanelVScroll_calcVThumbRect() as boolean
   dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMPANEL)
   if pWindow = 0 then exit function

   ' calculate the vertical scrollbar in client coordinates
   SetRectEmpty( @gPanelVScroll.rc )
  
   dim as long nTopIndex = SendMessage( gPanelVScroll.hListBox, LB_GETTOPINDEX, 0, 0 ) 

   dim as RECT rc 
   with gPanelVScroll
      GetClientRect( gPanelVScroll.hListBox, @rc )
      .listBoxHeight = rc.bottom - rc.top
      .itemHeight = SendMessage( gPanelVScroll.hListBox, LB_GETITEMHEIGHT, 0, 0 )
      .numItems = ListBox_GetCount( gPanelVScroll.hListBox ) 
      .itemsPerPage = .listBoxHeight / .itemHeight
      .thumbHeight = (.itemsPerPage / .numItems) * .listBoxHeight
      GetClientRect( HWND_FRMPANEL_VSCROLLBAR, @rc )
      .rc.Left = rc.Left
      .rc.top = rc.top + ((nTopIndex / .numItems) * .listBoxHeight)
      .rc.Right = rc.right
      .rc.bottom = .rc.top + .thumbHeight
      if .numItems < .itemsPerPage then return true
   end with

   function = 0
end function

    
' ========================================================================================
' Update the VScrollBar UI via UpdateLayeredWindow
' ========================================================================================
function frmPanelVScroll_UpdateUI() as LRESULT
   dim as long nWidth = AfxGetWindowWidth( HWND_FRMPANEL_VSCROLLBAR )
   dim as long nHeight = AfxGetWindowHeight( HWND_FRMPANEL_VSCROLLBAR )
    
   '// Make mem DC + mem  bitmap
   dim as HDC hdcScreen = GetDC(null)
   dim as HDC hDC = CreateCompatibleDC(hdcScreen)
   dim as HBITMAP hBmp = CreateCompatibleBitmap(hdcScreen, nWidth, nHeight)
   dim as HBITMAP hBmpOld = SelectObject(hDC, hBmp)

   '// Draw to StatusBar into the memory DC
   FillRect( hDC, @gPanelVScroll.rc, ghPanel.hBackBrushThumb )

   '// Call UpdateLayeredWindow
   dim as BLENDfunction blend
   blend.BlendOp = AC_SRC_OVER
   blend.SourceConstantAlpha = 200  '255
   blend.AlphaFormat = AC_SRC_ALPHA

   dim as POINT ptPos
   dim as RECT rc = AfxGetWindowRect( HWND_FRMPANEL_VSCROLLBAR )
   ptPos.x = rc.left
   ptPos.y = rc.top
   dim as SIZE sizeWnd = (nWidth, nHeight)
   dim as POINT ptSrc = (0, 0)

   UpdateLayeredWindow ( _
            HWND_FRMPANEL_VSCROLLBAR, _  ' A handle to a layered window
            hdcScreen, _    ' A handle to a DC for the screen
            @ptPos, _       ' New screen position. pptDst can be NULL if not changing
            @sizeWnd, _     ' New size. psize can be NULL if not changing
            hDC, _          ' Handle to DC of layered window from CreateCompatibleDC
            @ptSrc, _       ' The location of the layer in the device context
            0, _            ' COLORREF color key to be used when composing the layered window
            @blend, _       ' Pointer to BLENDfunction structure
            ULW_ALPHA )     ' Use pblend as the blend function

   SelectObject( hDC, hBmpOld )
   DeleteObject( hBmp )
   DeleteDC( hDC )
   ReleaseDC( null, hdcScreen )

   function = 0
end function


' ========================================================================================
' Position the VScrollBar over the Explorer listbox
' ========================================================================================
function frmPanelVScroll_PositionWindows( byval nShowState as long ) as LRESULT

   dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMPANEL_VSCROLLBAR)
   if pWindow = 0 then exit function
   
   ' Position the VScrollBar relative to the Explorer Listbox using screen coordinates
   dim as Rect rc = AfxGetWindowRect( gPanelVScroll.hListBox )
   dim as long nScrollWidth = pWindow->ScaleX(SCROLLBAR_WIDTH_PANEL)
 
   SetWindowPos( HWND_FRMPANEL_VSCROLLBAR, HWND_TOP, _
                  rc.right - nScrollWidth, _
                  rc.top, _
                  nScrollWidth, _
                  rc.bottom - rc.top, _
                  SWP_NOACTIVATE )

   ' returns true if RECT is empty
   if frmPanelVScroll_calcVThumbRect() then nShowState = SW_HIDE 
   ShowWindow( HWND_FRMPANEL_VSCROLLBAR, nShowState )
   if nShowState = SW_HIDE then exit function

   if gApp.isWineActive then
      AfxRedrawWindow( HWND_FRMPANEL_VSCROLLBAR )
   else   
      frmPanelVScroll_UpdateUI()
   end if
      
   function = 0
end function


' ========================================================================================
' frmPanelVScroll Window procedure
' ========================================================================================
function frmPanelVScroll_WndProc( _
            byval HWnd   as HWnd, _
            byval uMsg   as UINT, _
            byval wParam as WPARAM, _
            byval lParam as LPARAM _
            ) as LRESULT

   static as POINT prev_pt   ' screen pt.y cursor position

   select case uMsg
      case WM_LBUTTONDOWN
         dim as POINT pt: GetCursorPos( @pt )
         frmPanelVScroll_calcVThumbRect()   ' in client coordinates
         dim as RECT rc = gPanelVScroll.rc ' covert copy to screen coordinates
         MapWindowPoints( HWND_FRMPANEL_VSCROLLBAR, HWND_DESKTOP, cast(POINT ptr, @rc), 2)
         if PtInRect( @rc, pt ) then
            prev_pt = pt
            gApp.bDragActive = true
            SetCapture( HWnd )
         else
            ' we have clicked on a PageUp or PageDn
            dim as long nTopIndex = SendMessage( gPanelVScroll.hListBox, LB_GETTOPINDEX, 0, 0 ) 
            if pt.y < rc.top then
               nTopIndex = max( nTopIndex - gPanelVScroll.itemsPerPage, 0 )
               SendMessage( gPanelVScroll.hListBox, LB_SETTOPINDEX, nTopIndex, 0 ) 
               frmPanelVScroll_calcVThumbRect()   ' in client coordinates
               if gApp.isWineActive then
                  AfxRedrawWindow( HWND_FRMPANEL_VSCROLLBAR )
               else
                  frmPanelVScroll_UpdateUI()
               end if   
               AfxRedrawWindow( gPanelVScroll.hListBox )
            elseif pt.y > rc.bottom then
               dim as long nMaxTopIndex = gPanelVScroll.numItems - gPanelVScroll.itemsPerPage
               nTopIndex = min( nTopIndex + gPanelVScroll.itemsPerPage, nMaxTopIndex )
               SendMessage( gPanelVScroll.hListBox, LB_SETTOPINDEX, nTopIndex, 0 ) 
               frmPanelVScroll_calcVThumbRect()   ' in client coordinates
               if gApp.isWineActive then
                  AfxRedrawWindow( HWND_FRMPANEL_VSCROLLBAR )
               else
                  frmPanelVScroll_UpdateUI()
               end if   
               AfxRedrawWindow( gPanelVScroll.hListBox )
            end if
         end if
         
      case WM_MOUSEMOVE
         if gApp.bDragActive then
            dim as POINT pt: GetCursorPos( @pt )
            if pt.y <> prev_pt.y then 
               dim as long delta = pt.y - prev_pt.y 
               ' convert to client coordinates for ease of use
               dim as RECT rc: GetClientRect( HWND_FRMPANEL_VSCROLLBAR, @rc )

               gPanelVScroll.rc.top = max(0, gPanelVScroll.rc.top + delta)
               gPanelVScroll.rc.top = min(gPanelVScroll.rc.top, rc.bottom-gPanelVScroll.thumbHeight)
               gPanelVScroll.rc.bottom = gPanelVScroll.rc.top + gPanelVScroll.thumbHeight

               prev_pt = pt
               
               dim as long nPrevTopIndex = SendMessage( gPanelVScroll.hListBox, LB_GETTOPINDEX, 0, 0 ) 
               dim as long nLastIndex = (gPanelVScroll.rc.bottom / rc.bottom) * gPanelVScroll.numItems
               dim as long nTopIndex = nLastIndex - gPanelVScroll.itemsPerPage
               if nTopIndex <> nPrevTopIndex then
                  SendMessage( gPanelVScroll.hListBox, LB_SETTOPINDEX, nTopIndex, 0 ) 
                  AfxRedrawWindow( gPanelVScroll.hListBox )
               end if   

               if gApp.isWineActive then
                  AfxRedrawWindow( HWND_FRMPANEL_VSCROLLBAR )
               else
                  frmPanelVScroll_UpdateUI()
               end if   
            end if
        end if
         
      case WM_LBUTTONUP   
         gApp.bDragActive = false
         prev_pt.x = 0
         prev_pt.y = 0
         ReleaseCapture
         
      
      case WM_ERASEBKGND
         if gApp.isWineActive then
            return true
         end if
         
      case WM_PAINT
         if gApp.isWineActive then
            Dim As PAINTSTRUCT ps
            Dim As HDC hDC
            hDC = BeginPaint( hWnd, @ps )
            FillRect( hDC, @ps.rcPaint, ghPanel.hBackBrushScrollBar )
            FillRect( hDC, @gPanelVScroll.rc, ghPanel.hBackBrushThumb )
            EndPaint hWnd, @ps
            exit function
         end if      

   end select

   ' for messages that we don't deal with
   function = DefWindowProc( HWnd, uMsg, wParam, lParam )

end function


' ========================================================================================
' frmPanelVScroll_Show
' ========================================================================================
function frmPanelVScroll_Show( byval hWndParent as HWND ) as LRESULT

   '  Create the main window and child controls
   dim pWindow as CWindow ptr = new CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   dim as long ExStyles = WS_EX_NOACTIVATE
   if gApp.isWineActive = false then
      ExStyles = ExStyles or WS_EX_LAYERED
   end if
   
   HWND_FRMPANEL_VSCROLLBAR = pWindow->Create( hWndParent, _
        "", @frmPanelVScroll_WndProc, 0, 0, SCROLLBAR_WIDTH_PANEL, 0, _
        WS_POPUP or WS_CLIPSIBLINGS or WS_CLIPCHILDREN, _
        ExStyles )
   pWindow->Brush = ghPanel.hBackBrushScrollBar

   function = 0
   
end function
