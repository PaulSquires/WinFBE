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
''  frmEditorHScroll.inc
''   

#include once "frmEditorHScroll.bi"

' ========================================================================================
' Calculate the RECT that holds the client coordinates of the scrollbar's horizontal thumb
' Returns True if RECT is not empty
' ========================================================================================
function frmEditorHScroll_calcHThumbRect() as boolean
'   dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMMAIN)
'   if pWindow = 0 Then exit function

'   numLines as long 
'   linesPerPage as long
'   thumbHeight as long
'   rc as RECT
'end type
'dim shared gEditorVScroll(1) as EDITOR_VSCROLL_TYPE

'   ' calculate the vertical scrollbars in client coordinates
'   for i as long = lbound(
'   SetRectEmpty( @gEditorVScroll(0).rc )
'   SetRectEmpty( @gEditorVScroll(1).rc )
  

'   dim as Rect rc: GetClientRect( HWND_FRMMAIN_TOPTABS, @rc )
   
'   gTopTabsHScroll.rc = rc
'   gTopTabsHScroll.rc.right = grcTopTabActionPanel.Left
'   gTopTabsHScroll.rc.bottom = gTopTabsHScroll.rc.top + pWindow->ScaleY( TOPTABS_SCROLLBAR_HEIGHT )
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
' frmEditorHScroll Window procedure
' ========================================================================================
function frmEditorHScroll_WndProc( _
            ByVal HWnd   as HWnd, _
            ByVal uMsg   as UINT, _
            ByVal wParam as WPARAM, _
            ByVal lParam as LPARAM _
            ) as LRESULT

   static as POINT prev_pt   ' screen pt.y cursor position

   Select Case uMsg
      case WM_DESTROY
         if HWnd = HWND_FRMEDITOR_HSCROLLBAR(0) then HWND_FRMEDITOR_HSCROLLBAR(0) = 0
         if HWnd = HWND_FRMEDITOR_HSCROLLBAR(1) then HWND_FRMEDITOR_HSCROLLBAR(1) = 0
              
      case WM_LBUTTONDOWN
'         dim as POINT pt: GetCursorPos( @pt )
'         frmEditorHScroll_calcHThumbRect()   ' in client coordinates
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
'               frmEditorHScroll_calcHThumbRect()   ' in client coordinates
'               frmEditorHScroll_UpdateUI()
'               AfxRedrawWindow( HWND_FRMEXPLORER_LISTBOX )
'            elseif pt.y > rc.bottom then
'               dim as long nMaxTopIndex = gExplorerVScroll.numItems - gExplorerVScroll.itemsPerPage
'               nTopIndex = min( nTopIndex + gExplorerVScroll.itemsPerPage, nMaxTopIndex )
'               SendMessage( HWND_FRMEXPLORER_LISTBOX, LB_SETTOPINDEX, nTopIndex, 0 ) 
'               frmEditorHScroll_calcHThumbRect()   ' in client coordinates
'               frmEditorHScroll_UpdateUI()
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

'               frmEditorHScroll_UpdateUI()
'            end if
'        end if
         
      case WM_LBUTTONUP   
'         gApp.bDragActive = false
'         prev_pt.x = 0
'         prev_pt.y = 0
'         ReleaseCapture
         
      case WM_ERASEBKGND
         return true

      case WM_PAINT
         Dim As PAINTSTRUCT ps
         Dim As HDC hDc
         hDC = BeginPaint( hWnd, @ps )
         SaveDC( hDC )
         FillRect( hDC, @ps.rcPaint, ghEditor.hBackBrushScrollBar )
'         FillRect( hDC, @gEditorVScroll(idxWindow).rc, ghEditor.hBackBrushThumb )
         EndPaint( hWnd, @ps )
         RestoreDC( hDC, -1 )

   end Select

   ' for messages that we don't deal with
   function = DefWindowProc( HWnd, uMsg, wParam, lParam )

end function


' ========================================================================================
' frmEditorHScroll_Show
' ========================================================================================
function frmEditorHScroll_Show( ByVal hWndParent as HWnd ) as LRESULT

   '  Create the main window and child controls
   dim pWindow as CWindow ptr = new CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   for i as long = 0 to 1
      HWND_FRMEDITOR_HSCROLLBAR(i) = pWindow->Create( hWndParent, _
           "", @frmEditorHScroll_WndProc, 0, 0, 0, SCROLLBAR_HEIGHT, _
           WS_CHILD or WS_CLIPSIBLINGS or WS_CLIPCHILDREN, WS_EX_NOACTIVATE )
   next
   
   function = 0
   
end function


' ========================================================================================
' frmEditorHScroll_RepositionEditor
' ========================================================================================
function frmEditorHScroll_RepositionEditor( _
            byval hScrollBar as HWND, _
            byval newState as long, _
            byval pDoc as clsDocument ptr, _
            byval idxWindow as long _
            ) as long
   
   dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMMAIN)
   if pWindow = 0 Then exit function

   if (pDoc->IsDesigner) andalso (IsDesignerView(pDoc)) then
      exit function
   end if

   dim as long nHeightScrollbar = pWindow->ScaleY(SCROLLBAR_HEIGHT)
   
   ' Position the HScrollbar immediately underneath the edit window
   dim as RECT rc = AfxGetWindowRect( pDoc->hWindow(idxWindow) )
   MapWindowPoints( HWND_DESKTOP, HWND_FRMMAIN, cast( POINT ptr, @rc ), 2 )

   ' If we are wanting to show the HScrollbar then test the visible lines to see if any
   ' are actually wider than screen width. If not, then no need to show the HScrollbar.
   ' We have to use PositionFromPoint calculations because the LinesOnScreen message does
   ' not take into account extra ascent and descent values.
   if newState = SW_SHOWNA then
      dim as long GetFirstVisiblePos, GetLastVisiblePos, GetFirstVisibleLine, GetLastVisibleLine
      dim as long ClientWidth = (rc.right - rc.left)
      dim as long ClientHeight = (rc.bottom - rc.top)
      GetFirstVisiblePos = SciExec( pDoc->hWindow(idxWindow), SCI_POSITIONFROMPOINT, 0, 0 )
      GetLastVisiblePos = SciExec( pDoc->hWindow(idxWindow), SCI_POSITIONFROMPOINT, ClientWidth, ClientHeight )
      GetFirstVisibleLine = SciExec( pDoc->hWindow(idxWindow), SCI_LINEFROMPOSITION, GetFirstVisiblePos, 0 )
      GetLastVisibleLine = SciExec( pDoc->hWindow(idxWindow), SCI_LINEFROMPOSITION, GetLastVisiblePos, 0 )

      dim as long GetXOffset, GetLineEndPosition, endPos
      dim as boolean needHSB = false
      'dim as long x = ClientWidth
      for i as long = GetFirstVisibleLine to GetLastVisibleLine
         ' GetXOffset adds left scroll spacing if we are already scrolled left some.
         'GetXOffset = SciExec( pDoc->hWindow(idxWindow), SCI_GETXOFFSET, 0, 0 )
         GetLineEndPosition = SciExec( pDoc->hWindow(idxWindow), SCI_GETLINEENDPOSITION, i, 0 )
         endPos = SciExec( pDoc->hWindow(idxWindow), SCI_POINTXFROMPOSITION, 0, GetLineEndPosition )
         needHSB = endPos > ClientWidth
         if needHSB then exit for
      next
      if needHSB = false then newState = SW_HIDE
   end if
   
   if newState = SW_SHOWNA then
      dim as long nHeight = rc.bottom-rc.top-nHeightScrollbar
      SetWindowPos( pDoc->hWindow(idxWindow), 0, _
                      0, 0, rc.right-rc.left, nHeight, _
                      SWP_NOMOVE or SWP_NOZORDER or SWP_SHOWWINDOW )
      SetWindowPos( hScrollBar, 0, _
                      rc.left, rc.top+nHeight, rc.right-rc.left, nHeightScrollbar, _
                      SWP_NOZORDER )
      ShowWindow( hScrollBar, SW_SHOWNA )                
   else
      ShowWindow( hScrollBar, SW_HIDE ) 
      dim as RECT rc = pDoc->origPosition(idxWindow) 
      MapWindowPoints( HWND_DESKTOP, HWND_FRMMAIN, cast(POINT ptr, @rc), 2 )              
      SetWindowPos( pDoc->hWindow(idxWindow), 0, _
                      0, 0, rc.right-rc.left, rc.bottom-rc.top, _
                      SWP_NOMOVE or SWP_NOZORDER or SWP_SHOWWINDOW )
   end if
   
   function = 0
   
end function
