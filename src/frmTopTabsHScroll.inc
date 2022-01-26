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


''
''  frmTopTabsHScroll.inc
''   

#include once "frmTopTabs.bi"


' ========================================================================================
' Calculate the RECT that holds the client coordinates of the scrollbar's horizontal thumb
' Returns True if RECT is not empty
' ========================================================================================
function frmTopTabsHScroll_calcHThumbRect() as boolean
   dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMMAIN_TOPTABS)
   if pWindow = 0 Then exit function

   ' calculate the horizontal scrollbar in client coordinates
   SetRectEmpty( @gTopTabsHScroll.rc )
  
   dim as Rect rc: GetClientRect( HWND_FRMMAIN_TOPTABS, @rc )
   'dim as Rect rc: GetClientRect( HWND_FRMMAIN, @rc )
   
   gTopTabsHScroll.rc = rc
   gTopTabsHScroll.rc.right = grcTopTabActionPanel.Left
   gTopTabsHScroll.rc.bottom = gTopTabsHScroll.rc.top + pWindow->ScaleY( TOPTABS_SCROLLBAR_HEIGHT )
'   gTopTabsHScroll.rc = rc
'   gTopTabsHScroll.rc.right = grcTopTabActionPanel.Left
'   gTopTabsHScroll.rc.top = gTopTabsHScroll.rc.bottom - pWindow->ScaleY( TOPTABS_SCROLLBAR_HEIGHT )
  
'   dim as long nTopIndex = SendMessage( HWND_FRMEXPLORER_LISTBOX, LB_GETTOPINDEX, 0, 0 ) 

'   dim as Rect rc 
'   with gExplorerVScroll
'      GetClientRect( HWND_FRMEXPLORER_LISTBOX, @rc )
'      .listBoxHeight = rc.bottom - rc.top
'      .itemHeight = SendMessage( HWND_FRMEXPLORER_LISTBOX, LB_GETITEMHEIGHT, 0, 0 )
'      .numItems = ListBox_GetCount( HWND_FRMEXPLORER_LISTBOX ) 
'      .itemsPerPage = .listBoxHeight / .itemHeight
'      .thumbHeight = (.itemsPerPage / .numItems) * .listBoxHeight
'      GetClientRect( HWND_FRMEXPLORER_VSCROLLBAR, @rc )
'      .rc.Left = rc.Left
'      .rc.top = rc.top + ((nTopIndex / .numItems) * .listBoxHeight)
'      .rc.Right = rc.right
'      .rc.bottom = .rc.top + .thumbHeight
'      if .numItems < .itemsPerPage then return true
'   end with

   function = 0
end function

    
' ========================================================================================
' Update the HScrollBar UI via UpdateLayeredWindow
' ========================================================================================
function frmTopTabsHScroll_UpdateUI() as LRESULT
   dim as long nWidth = AfxGetWindowWidth( HWND_FRMTOPTABS_HSCROLLBAR )
   dim as long nHeight = AfxGetWindowHeight( HWND_FRMTOPTABS_HSCROLLBAR  )

   '// Make mem DC + mem  bitmap
   dim as HDC hdcScreen = GetDC(null)
   dim as HDC hDC = CreateCompatibleDC(hdcScreen)
   dim as HBITMAP hBmp = CreateCompatibleBitmap(hdcScreen, nWidth, nHeight)
   dim as HBITMAP hBmpOld = SelectObject(hDC, hBmp)

   '// Draw to StatusBar into the memory DC
   FillRect( hDC, @gTopTabsHScroll.rc, ghTopTabs.hBackBrushThumb )

   '// Call UpdateLayeredWindow
   dim as BLENDfunction blend
   blend.BlendOp = AC_SRC_OVER
   blend.SourceConstantAlpha = 200  '255
   blend.AlphaFormat = AC_SRC_ALPHA

   dim as POINT ptPos
   dim as RECT rc = AfxGetWindowRect( HWND_FRMTOPTABS_HSCROLLBAR )
   ptPos.x = rc.left
   ptPos.y = rc.top
   dim as SIZE sizeWnd = (nWidth, nHeight)
   dim as POINT ptSrc = (0, 0)
   
   UpdateLayeredWindow ( _
            HWND_FRMTOPTABS_HSCROLLBAR, _  ' A handle to a layered window
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
' Position the HScrollBar over the Top Tabs
' ========================================================================================
function frmTopTabsHScroll_PositionWindows( byval nShowState as long ) as LRESULT

   dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMTOPTABS_HSCROLLBAR)
   if pWindow = 0 Then exit function
   
   ' Position the HScrollBar relative to the Top Tabs using screen coordinates
   dim as Rect rc = AfxGetWindowRect( HWND_FRMMAIN_TOPTABS )
   dim as long nScrollHeight = pWindow->ScaleY( TOPTABS_SCROLLBAR_HEIGHT )
 
   SetWindowPos( HWND_FRMTOPTABS_HSCROLLBAR, HWND_TOP, _
                  rc.left, _
                  rc.bottom - nScrollHeight, _
                  rc.right - rc.left, _
                  nScrollHeight, _
                  SWP_NOACTIVATE )

   ' returns true if RECT is empty
   if frmTopTabsHScroll_calcHThumbRect() then nShowState = SW_HIDE 
   ShowWindow( HWND_FRMTOPTABS_HSCROLLBAR, nShowState )
   if nShowState = SW_HIDE then exit function
   frmTopTabsHScroll_UpdateUI()
   
   function = 0
end function


' ========================================================================================
' frmTopTabsHScroll Window procedure
' ========================================================================================
function frmTopTabsHScroll_WndProc( _
            ByVal HWnd   as HWnd, _
            ByVal uMsg   as UINT, _
            ByVal wParam as WPARAM, _
            ByVal lParam as LPARAM _
            ) as LRESULT

   static as POINT prev_pt   ' screen pt.y cursor position

   Select Case uMsg
      case WM_LBUTTONDOWN
'         dim as POINT pt: GetCursorPos( @pt )
'         frmTopTabsHScroll_calcHThumbRect()   ' in client coordinates
'         dim as RECT rc = gExplorerVScroll.rc ' covert copy to screen coordinates
'         MapWindowPoints( HWND_FRMEXPLORER_VSCROLLBAR, HWND_DESKTOP, cast(POINT ptr, @rc), 2)
'         if PtInRect( @rc, pt ) then
'            prev_pt = pt
'            gApp.bDragActive = true
'            SetCapture( HWnd )
'         else
'            ' we have clicked on a PageUp or PageDn
'            dim as long nTopIndex = SendMessage( HWND_FRMEXPLORER_LISTBOX, LB_GETTOPINDEX, 0, 0 ) 
'            if pt.y < rc.top then
'               nTopIndex = max( nTopIndex - gExplorerVScroll.itemsPerPage, 0 )
'               SendMessage( HWND_FRMEXPLORER_LISTBOX, LB_SETTOPINDEX, nTopIndex, 0 ) 
'               frmTopTabsHScroll_calcHThumbRect()   ' in client coordinates
'               frmTopTabsHScroll_UpdateUI()
'               AfxRedrawWindow( HWND_FRMEXPLORER_LISTBOX )
'            elseif pt.y > rc.bottom then
'               dim as long nMaxTopIndex = gExplorerVScroll.numItems - gExplorerVScroll.itemsPerPage
'               nTopIndex = min( nTopIndex + gExplorerVScroll.itemsPerPage, nMaxTopIndex )
'               SendMessage( HWND_FRMEXPLORER_LISTBOX, LB_SETTOPINDEX, nTopIndex, 0 ) 
'               frmTopTabsHScroll_calcHThumbRect()   ' in client coordinates
'               frmTopTabsHScroll_UpdateUI()
'               AfxRedrawWindow( HWND_FRMEXPLORER_LISTBOX )
'            end if
'         end if
         
      case WM_MOUSEMOVE
'         if gApp.bDragActive then
'            dim as POINT pt: GetCursorPos( @pt )
'            if pt.y <> prev_pt.y then 
'               dim as long delta = pt.y - prev_pt.y 
'               ' convert to client coordinates for ease of use
'               dim as RECT rc: GetClientRect( HWND_FRMEXPLORER_VSCROLLBAR, @rc )

'               gExplorerVScroll.rc.top = max(0, gExplorerVScroll.rc.top + delta)
'               gExplorerVScroll.rc.top = min(gExplorerVScroll.rc.top, rc.bottom-gExplorerVScroll.thumbHeight)
'               gExplorerVScroll.rc.bottom = gExplorerVScroll.rc.top + gExplorerVScroll.thumbHeight

'               prev_pt = pt
               
'               dim as long nPrevTopIndex = SendMessage( HWND_FRMEXPLORER_LISTBOX, LB_GETTOPINDEX, 0, 0 ) 
'               dim as long nLastIndex = (gExplorerVScroll.rc.bottom / rc.bottom) * gExplorerVScroll.numItems
'               dim as long nTopIndex = nLastIndex - gExplorerVScroll.itemsPerPage
'               if nTopIndex <> nPrevTopIndex then
'                  SendMessage( HWND_FRMEXPLORER_LISTBOX, LB_SETTOPINDEX, nTopIndex, 0 ) 
'                  AfxRedrawWindow( HWND_FRMEXPLORER_LISTBOX )
'               end if   

               frmTopTabsHScroll_UpdateUI()
'            end if
'        end if
         
      case WM_LBUTTONUP   
'         gApp.bDragActive = false
'         prev_pt.x = 0
'         prev_pt.y = 0
'         ReleaseCapture
         
   end Select

   ' for messages that we don't deal with
   function = DefWindowProc( HWnd, uMsg, wParam, lParam )

end function


' ========================================================================================
' frmTopTabsHScroll_Show
' ========================================================================================
function frmTopTabsHScroll_Show( ByVal hWndParent as HWnd ) as LRESULT

   '  Create the main window and child controls
   dim pWindow as CWindow ptr = new CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   HWND_FRMTOPTABS_HSCROLLBAR = pWindow->Create( hWndParent, _
        "", @frmTopTabsHScroll_WndProc, 0, 0, 0, TOPTABS_SCROLLBAR_HEIGHT, _
        WS_POPUP or WS_CLIPSIBLINGS or WS_CLIPCHILDREN, _
        WS_EX_LAYERED or WS_EX_NOACTIVATE )
   pWindow->Brush = ghTopTabs.hBackBrushScrollBar

   function = 0
   
end function
