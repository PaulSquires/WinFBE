'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2025 Paul Squires, PlanetSquires Software
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
' frmEditorHScroll Window procedure
' ========================================================================================
function frmEditorHScroll_WndProc( _
            ByVal HWnd   as HWnd, _
            ByVal uMsg   as UINT, _
            ByVal wParam as WPARAM, _
            ByVal lParam as LPARAM _
            ) as LRESULT

   static as POINT prev_pt   ' screen pt.y cursor position
   static as long prev_xOffset
   dim as long idxWindow = iif( HWnd = HWND_FRMEDITOR_HSCROLLBAR(0), 0, 1 )

   Select Case uMsg
      case WM_DESTROY
         if HWnd = HWND_FRMEDITOR_HSCROLLBAR(0) then HWND_FRMEDITOR_HSCROLLBAR(0) = 0
         if HWnd = HWND_FRMEDITOR_HSCROLLBAR(1) then HWND_FRMEDITOR_HSCROLLBAR(1) = 0
              
      case WM_LBUTTONDOWN
         dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
         if pDoc then
            gApp.SuppressNotify = true   ' prevent SCN_UPDATEUI
            dim as long delta = 0
            dim as POINT pt: GetCursorPos( @pt )
            dim as RECT rc = gEditorHScroll(idxWindow).rc ' covert copy to screen coordinates
            MapWindowPoints( HWND_FRMEDITOR_HSCROLLBAR(idxWindow), HWND_DESKTOP, cast(POINT ptr, @rc), 2)
            if PtInRect( @rc, pt ) then
               prev_pt = pt
               gApp.bDragActive = true
               SetCapture( HWnd )
            elseif pt.x < rc.left then
               delta = -(rc.left - pt.x)
            elseif pt.x > rc.right then
               delta = pt.x - rc.right
            end if
           
            ' The client area was clicked
            if delta <> 0 then
               dim as RECT rc: GetClientRect( HWND_FRMEDITOR_HSCROLLBAR(idxWindow), @rc )
               dim as long xOffset = SciExec( pDoc->hWindow(idxWindow), SCI_GETXOFFSET, 0, 0 )
               gEditorHScroll(idxWindow).rc = rc
               gEditorHScroll(idxWindow).rc.left = xOffset
               gEditorHScroll(idxWindow).rc.right = xOffset + ((rc.right / gEditorHScroll(idxWindow).lineLength) * rc.right)

               if gEditorHScroll(idxWindow).rc.right + delta < rc.right then
                  xOffset = xOffset + delta
                  xOffset = max( 0, xOffset )
                  if xOffset <> prev_xOffset then
                     SendMessage( pDoc->hWindow(idxWindow), SCI_SETXOFFSET, xOffset, 0) 
                     prev_xOffset = xOffset
                     AfxRedrawWindow( HWND_FRMEDITOR_HSCROLLBAR(idxWindow) )
                  end if   
               end if
            end if
         end if
         
      case WM_MOUSEMOVE
         if gApp.bDragActive then
            dim as POINT pt: GetCursorPos( @pt )
            if pt.x <> prev_pt.x then 
               dim as long delta = pt.x - prev_pt.x 

               prev_pt = pt
               
               dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
               if pDoc then
                  dim as RECT rc: GetClientRect( HWND_FRMEDITOR_HSCROLLBAR(idxWindow), @rc )
                  dim as long xOffset = SciExec( pDoc->hWindow(idxWindow), SCI_GETXOFFSET, 0, 0 )
                  gEditorHScroll(idxWindow).rc = rc
                  gEditorHScroll(idxWindow).rc.left = xOffset
                  gEditorHScroll(idxWindow).rc.right = xOffset + ((rc.right / gEditorHScroll(idxWindow).lineLength) * rc.right)

                  if gEditorHScroll(idxWindow).rc.right + delta < rc.right then
                     xOffset = xOffset + delta
                     xOffset = max( 0, xOffset )
                     if xOffset <> prev_xOffset then
                        SendMessage( pDoc->hWindow(idxWindow), SCI_SETXOFFSET, xOffset, 0) 
                        prev_xOffset = xOffset
                        AfxRedrawWindow( HWND_FRMEDITOR_HSCROLLBAR(idxWindow) )
                     end if   
                  end if
               end if
            end if
        end if
         
      case WM_LBUTTONUP   
         gApp.SuppressNotify = false   ' allow SCN_UPDATEUI
         gApp.bDragActive = false
         prev_pt.x = 0
         prev_pt.y = 0
         ReleaseCapture
         
      case WM_ERASEBKGND
         return true

      case WM_PAINT
         Dim As PAINTSTRUCT ps
         Dim As HDC hDc
         hDC = BeginPaint( hWnd, @ps )
         SaveDC( hDC )
         FillRect( hDC, @ps.rcPaint, ghEditor.hBackBrushScrollBar )
         
         dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
         if pDoc then
            dim as RECT rc: GetClientRect( HWND_FRMEDITOR_HSCROLLBAR(idxWindow), @rc )
            dim as long GetXOffset = SciExec( pDoc->hWindow(idxWindow), SCI_GETXOFFSET, 0, 0 )
            gEditorHScroll(idxWindow).rc = rc
            gEditorHScroll(idxWindow).rc.left = GetXOffset
            gEditorHScroll(idxWindow).rc.right = GetXOffset + ((rc.right / gEditorHScroll(idxWindow).lineLength) * rc.right)
            FillRect( hDC, @gEditorHScroll(idxWindow).rc, ghEditor.hBackBrushThumb )
         end if
         
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
   dim pWindow as CWindow ptr

   for i as long = 0 to 1
      pWindow = new CWindow
      pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI
      HWND_FRMEDITOR_HSCROLLBAR(i) = pWindow->Create( hWndParent, _
           "", @frmEditorHScroll_WndProc, 0, 0, 0, SCROLLBAR_HEIGHT, _
           WS_CHILD or WS_CLIPSIBLINGS or WS_CLIPCHILDREN, WS_EX_NOACTIVATE )
   next
   
   function = 0
   
end function


' ========================================================================================
' frmEditorHScroll_NeedScrollBar
' Determine based on visible line lengths if a scrollbar should be displayed
' ========================================================================================
function frmEditorHScroll_NeedScrollBar( _
            byval pDoc as clsDocument ptr, _
            byval idxWindow as long _
            ) as long

   dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMMAIN)
   if pWindow = 0 Then exit function

   dim as RECT rc = AfxGetWindowRect( pDoc->hWindow(idxWindow) )
   MapWindowPoints( HWND_DESKTOP, HWND_FRMMAIN, cast(POINT ptr, @rc), 2 )
   
   ' We have to use PositionFromPoint calculations because the LinesOnScreen message does
   ' not take into account extra ascent and descent values.
   dim as long GetFirstVisiblePos, GetLastVisiblePos, GetFirstVisibleLine, GetLastVisibleLine
   dim as long ClientWidth = (rc.right - rc.left)
   dim as long ClientHeight = (rc.bottom - rc.top)
   GetFirstVisiblePos = SciExec( pDoc->hWindow(idxWindow), SCI_POSITIONFROMPOINT, 0, 0 )
   GetLastVisiblePos = SciExec( pDoc->hWindow(idxWindow), SCI_POSITIONFROMPOINT, ClientWidth, ClientHeight )
   GetFirstVisibleLine = SciExec( pDoc->hWindow(idxWindow), SCI_LINEFROMPOSITION, GetFirstVisiblePos, 0 )
   GetLastVisibleLine = SciExec( pDoc->hWindow(idxWindow), SCI_LINEFROMPOSITION, GetLastVisiblePos, 0 )

   dim as long GetXOffset, GetLineEndPosition, endPos
   dim as boolean needHSB = false
   for i as long = GetFirstVisibleLine to GetLastVisibleLine
      dim as long GetXOffset = SciExec( pDoc->hWindow(idxWindow), SCI_GETXOFFSET, 0, 0 )
      GetLineEndPosition = SciExec( pDoc->hWindow(idxWindow), SCI_GETLINEENDPOSITION, i, 0 )
      endPos = SciExec( pDoc->hWindow(idxWindow), SCI_POINTXFROMPOSITION, 0, GetLineEndPosition )
      if (endPos > ClientWidth) orelse (GetXOffset <> 0) then 
         needHSB = true
      end if
      dim as long lineLength = GetXOffset + endPos + pWindow->ScaleX(200)
      if lineLength > gEditorHScroll(idxWindow).lineLength then
         gEditorHScroll(idxWindow).lineLength = lineLength
      end if
      if needHSB then exit for
   next
   function = iif( needHSB, SW_SHOWNA, SW_HIDE )
end function

