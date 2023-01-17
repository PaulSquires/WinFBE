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
''  frmEditorVScroll.inc
''   

#include once "frmEditorVScroll.bi"

' ========================================================================================
' Calculate the RECT that holds the client coordinates of the scrollbar's vertical thumb
' Returns True if RECT is not empty
' ========================================================================================
function frmEditorVScroll_calcVThumbRect( byval pDoc as clsDocument ptr ) as boolean
   dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMMAIN)
   if pWindow = 0 Then exit function
   if pDoc = 0 then exit function
   dim as RECT rc

   ' calculate the vertical scrollbars in client coordinates
   for i as long = lbound(pDoc->hWindow) to ubound(pDoc->hWindow)
      rc = AfxGetWindowRect(pDoc->hWindow(i))
      MapWindowPoints( HWND_DESKTOP, HWND_FRMMAIN, cast(POINT ptr, @rc), 2 ) 
      dim as long clientHeight = rc.bottom - rc.top
      if clientHeight = 0 then continue for
      dim as long firstVisibleLine = SciExec( pDoc->hWindow(i), SCI_GETFIRSTVISIBLELINE, 0, 0 )
      dim as long iScrollbarWidth = AfxGetWindowWidth(HWND_FRMEDITOR_VSCROLLBAR(i)) 
      dim as long minThumbHeight = pWindow->ScaleY(SCROLLBAR_MINTHUMBSIZE)
      SetRectEmpty( @gEditorVScroll(i).rc )
      with gEditorVScroll(i)
         .linesPerPage = pDoc->linesPerPage(i)
         .numLines = pDoc->GetLineCount + 1
         .thumbHeight = (.linesPerPage / .numLines) * clientHeight
         ' If the thumb height is less than the minimum default then adjust parameters so
         ' that the viewport can compensate for the adjustments.
         .thumbRatio = 1
         'if .thumbHeight < minThumbHeight then
         '   .thumbRatio = 1 + (minThumbHeight-.thumbHeight) /.thumbHeight
         '   .thumbHeight = minThumbHeight
         'end if
         .rc.left = 0
         .rc.right = .rc.left + iScrollbarWidth
         .rc.top = ((firstVisibleLine / .numLines) * clientHeight)
         .rc.bottom = .rc.top + .thumbHeight
      end with
   next

   function = 0
end function

    
' ========================================================================================
' frmEditorVScroll Window procedure
' ========================================================================================
function frmEditorVScroll_WndProc( _
            ByVal HWnd   as HWnd, _
            ByVal uMsg   as UINT, _
            ByVal wParam as WPARAM, _
            ByVal lParam as LPARAM _
            ) as LRESULT

   static as POINT prev_pt   ' screen pt.y cursor position

   dim as long idxWindow = iif( HWnd = HWND_FRMEDITOR_VSCROLLBAR(0), 0, 1 )

   Select Case uMsg
      case WM_DESTROY
         if HWnd = HWND_FRMEDITOR_VSCROLLBAR(0) then HWND_FRMEDITOR_VSCROLLBAR(0) = 0
         if HWnd = HWND_FRMEDITOR_VSCROLLBAR(1) then HWND_FRMEDITOR_VSCROLLBAR(1) = 0
              
      case WM_LBUTTONDOWN
         dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
         if pDoc then
            gApp.SuppressNotify = true   ' prevent SCN_UPDATEUI
            dim as POINT pt: GetCursorPos( @pt )
            frmEditorVScroll_calcVThumbRect(pDoc)   ' in client coordinates
            dim as RECT rc = gEditorVScroll(idxWindow).rc ' covert copy to screen coordinates
            MapWindowPoints( HWND_FRMEDITOR_VSCROLLBAR(idxWindow), HWND_DESKTOP, cast(POINT ptr, @rc), 2)
            if PtInRect( @rc, pt ) then
               prev_pt = pt
               gApp.bDragActive = true
               SetCapture( HWnd )
            else
               ' we have clicked on a PageUp or PageDn
               dim as long nTopIndex = SendMessage( pDoc->hWindow(idxWindow), SCI_GETFIRSTVISIBLELINE, 0, 0 ) 
               if pt.y < rc.top then
                  nTopIndex = max( nTopIndex - gEditorVScroll(idxWindow).linesPerPage, 0 )
                  SendMessage( pDoc->hWindow(idxWindow), SCI_SETFIRSTVISIBLELINE, nTopIndex, 0 ) 
                  frmEditorVScroll_calcVThumbRect(pDoc)   ' in client coordinates
                  AfxRedrawWindow( HWND_FRMEDITOR_VSCROLLBAR(idxWindow) )
               elseif pt.y > rc.bottom then
                  dim as long nMaxTopIndex = gEditorVScroll(idxWindow).numLines - gEditorVScroll(idxWindow).linesPerPage
                  nTopIndex = min( nTopIndex + gEditorVScroll(idxWindow).linesPerPage, nMaxTopIndex )
                  SendMessage( pDoc->hWindow(idxWindow), SCI_SETFIRSTVISIBLELINE, nTopIndex, 0 ) 
                  frmEditorVScroll_calcVThumbRect(pDoc)   ' in client coordinates
                  AfxRedrawWindow( HWND_FRMEDITOR_VSCROLLBAR(idxWindow) )
               end if
            end if
         end if
            
      case WM_MOUSEMOVE
         if gApp.bDragActive then
            dim as POINT pt: GetCursorPos( @pt )
            if pt.y <> prev_pt.y then 
               dim as long delta = pt.y - prev_pt.y 
               ' convert to client coordinates for ease of use
               dim as RECT rc: GetClientRect( HWND_FRMEDITOR_VSCROLLBAR(idxWindow), @rc )
               rc.bottom = rc.bottom * gEditorVScroll(idxWindow).thumbRatio
               gEditorVScroll(idxWindow).rc.top = max(0, gEditorVScroll(idxWindow).rc.top + delta)
               gEditorVScroll(idxWindow).rc.top = min(gEditorVScroll(idxWindow).rc.top, rc.bottom - gEditorVScroll(idxWindow).thumbHeight)
               gEditorVScroll(idxWindow).rc.bottom = gEditorVScroll(idxWindow).rc.top + gEditorVScroll(idxWindow).thumbHeight

               prev_pt = pt
               
               dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
               if pDoc then
                  dim as long nPrevTopLine = SendMessage( pDoc->hWindow(idxWindow), SCI_GETFIRSTVISIBLELINE, 0, 0 ) 
                  dim as long nTopLine = (gEditorVScroll(idxWindow).rc.top / rc.bottom) * gEditorVScroll(idxWindow).numLines
                  if nTopLine <> nPrevTopLine then
                     SendMessage( pDoc->hWindow(idxWindow), SCI_SETFIRSTVISIBLELINE, nTopLine, 0 ) 
                  end if   
               end if
                AfxRedrawWindow( HWND_FRMEDITOR_VSCROLLBAR(idxWindow) )
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
         dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMMAIN)
         Dim As PAINTSTRUCT ps
         Dim As HDC hDc

         hDC = BeginPaint( hWnd, @ps )
         SaveDC( hDC )
         FillRect( hDC, @ps.rcPaint, ghEditor.hBackBrushScrollBar )
         dim as long penWidth = pWindow->ScaleX(1)
         dim as HPEN hPenSolid = CreatePen( PS_SOLID, penWidth, ghEditor.Divider )
         SelectPen( hDC, hPenSolid )
         MoveToEx( hDC, ps.rcPaint.left, ps.rcPaint.top, null )
         LineTo( hDC, ps.rcPaint.left, ps.rcPaint.bottom )   
         FillRect( hDC, @gEditorVScroll(idxWindow).rc, ghEditor.hBackBrushThumb )
         EndPaint( hWnd, @ps )
         RestoreDC( hDC, -1 )
         DeletePen( hPenSolid )

   end Select

   ' for messages that we don't deal with
   function = DefWindowProc( HWnd, uMsg, wParam, lParam )

end function


' ========================================================================================
' frmEditorVScroll_Show
' ========================================================================================
function frmEditorVScroll_Show( ByVal hWndParent as HWnd ) as LRESULT

   '  Create the main window and child controls
   dim pWindow as CWindow ptr 

   for i as long = 0 to 1
      pWindow = new CWindow
      pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI
      HWND_FRMEDITOR_VSCROLLBAR(i) = pWindow->Create( hWndParent, _
           "", @frmEditorVScroll_WndProc, 0, 0, SCROLLBAR_WIDTH_EDITOR + 1, 0, _
           WS_CHILD or WS_CLIPSIBLINGS or WS_CLIPCHILDREN, WS_EX_NOACTIVATE )
   next

   function = 0
   
end function

