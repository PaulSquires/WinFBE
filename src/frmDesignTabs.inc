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

#include once "frmDesignTabs.bi"

' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
function frmDesignTabs_PositionWindows() As LRESULT

   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWND_FRMMAIN_DESIGNTABS)
   If pWindow = 0 Then Exit Function
   
   ' Get the entire client area
   Dim As Rect rc: GetClientRect( HWND_FRMMAIN_DESIGNTABS, @rc )
   
   ' Calculate the RECT positions for all of the top tabs
   dim as long nTextWidth = 0
   dim as long nLeft = rc.left 
   dim as long TabHeight = DESIGNTABS_HEIGHT
   dim as long rightBorderWidth = 2
   dim as long hmargin = 3
   
   for i as long = lbound(gDesignTabs) to ubound(gDesignTabs)
      ' Determine the length of the text
      nTextWidth = getTextWidth( HWND_FRMMAIN_DESIGNTABS, gDesignTabs(0).wszText, ghMenuBar.hFontMenuBar, 0 )
      ' calculate the various tab dimensions
      gDesignTabs(i).rcTab = rc
      gDesignTabs(i).rcTab.Left = nLeft
      gDesignTabs(i).rcTab.Right = nLeft + pWindow->ScaleX(hmargin + nTextWidth + hmargin + rightBorderWidth)
         
      gDesignTabs(i).rcText = gDesignTabs(i).rcTab
      gDesignTabs(i).rcText.Left = gDesignTabs(i).rcText.Left + hmargin
      gDesignTabs(i).rcText.Right = gDesignTabs(i).rcText.Left + pWindow->ScaleX(nTextWidth)
         
      gDesignTabs(i).isHot = false
      nLeft = gDesignTabs(i).rcTab.Right
   next   

   AfxRedrawWindow( HWND_FRMMAIN_DESIGNTABS )

   Function = 0
End Function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmTopTabs
' ========================================================================================
Function frmDesignTabs_OnSize( _
            ByVal HWnd As HWnd, _
            ByVal state As UINT, _
            ByVal cx As Long, _
            ByVal cy As Long _
            ) As LRESULT

   If state <> SIZE_MINIMIZED Then
      frmDesignTabs_PositionWindows()
   End If

   Function = 0
End Function


' ========================================================================================
' Do hit test to determine what tab is currently under the mouse cursor
' ========================================================================================
function getHotDesignTabsHitTest( byval hWin as HWnd ) as long
   dim as POINT pt: GetCursorPos( @pt )
   MapWindowPoints( HWND_DESKTOP, hWin, cast( POINT ptr, @pt ), 1 )
   dim as long hotTab = -1
   for i as long = lbound(gDesignTabs) to ubound(gDesignTabs)
      if PtInRect( @gDesignTabs(i).rcTab, pt ) then
         hotTab = i
         gDesignTabs(i).isHot = true
      else   
         gDesignTabs(i).isHot = false
      end if
   next
   function = hotTab
end function

' ========================================================================================
' frmDesignTabs Window procedure
' ========================================================================================
Function frmDesignTabs_WndProc( _
            ByVal HWnd   As HWnd, _
            ByVal uMsg   As UINT, _
            ByVal wParam As WPARAM, _
            ByVal lParam As LPARAM _
            ) As LRESULT

   static as long curr_idxHot = -1
   
   Select Case uMsg
   
   case WM_ERASEBKGND
      return true

   case WM_MOUSEMOVE
      Dim tme As TrackMouseEvent
      tme.cbSize = Sizeof(TrackMouseEvent)
      tme.dwFlags = TME_HOVER Or TME_LEAVE
      tme.hwndTrack = HWnd
      TrackMouseEvent(@tme) 

      curr_idxHot = getHotDesignTabsHitTest( HWnd )
      AfxRedrawWindow( HWnd )
      
   case WM_MOUSELEAVE
      ' this will reset all tabs isHot to -1 and curr_idxHot to -1
      curr_idxHot = getHotDesignTabsHitTest( HWnd )
      AfxRedrawWindow( HWnd )

   case WM_LBUTTONDOWN
      SetCapture( hWnd )
      
   case WM_LBUTTONUP
      ' Reset the mouse pointer
      ReleaseCapture
      curr_idxHot = getHotDesignTabsHitTest( HWnd )
      dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
      if pDoc then
         if (curr_idxHot = 0) orelse (curr_idxHot = 1) then
            pDoc->DesignTabsCurSel = curr_idxHot
            frmMain_PositionWindows()
            PostMessage( HWND_FRMMAIN, MSG_USER_GENERATECODE, 0, 0 )
         end if   
      end if
      

   case WM_PAINT
      Dim As PAINTSTRUCT ps
      Dim As HDC hDc

      Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWND_FRMMAIN_DESIGNTABS)
      If pWindow = 0 Then Exit Function

      hDC = BeginPaint( hWnd, @ps )

      SaveDC(hDC)
      dim as long nWidth = ps.rcPaint.right - ps.rcPaint.left
      dim as long nHeight = ps.rcPaint.bottom - ps.rcPaint.top

      Dim memDC as HDC      ' Double buffering
      Dim hbit As HBITMAP   ' Double buffering

      memDC = CreateCompatibleDC( hDC )
      hbit  = CreateCompatibleBitmap( hDC, nWidth, nHeight )
      If hbit Then hbit = SelectObject( memDC, hbit )

      ' Fill in the entire back panel width across the top of the screen
      FillRect( memDC, @ps.rcPaint, ghTopTabs.hPanelBrush )

      ' Create a black pen that acts as the divider for each tab
      dim as long penWidth = pWindow->ScaleX(1)
      dim as HPEN hPenSolid = CreatePen( PS_SOLID, penWidth, ghTopTabs.Divider )
      SelectObject( memDC, hPenSolid)
      
      dim as HPEN hPenNull = CreatePen( PS_NULL, 1, 0 )  ' null/invisible pen

      dim as POINT pt: GetCursorPos( @pt )
      dim as RECT rc
      
      dim as long nCurSel = -1
      dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
      if pDoc then nCurSel = pDoc->DesignTabsCurSel
      
      ' All of the rc calculations have already been done in frmTopTabs_PostionWindows
      for i as long = lbound(gDesignTabs) to ubound(gDesignTabs)
         ' paint this tab based on active/inactive status
         if i = nCurSel then
            SetBkColor( memDC, ghTopTabs.BackColorHot )
            SetTextColor( memDC, ghTopTabs.ForeColorHot )
            FillRect( memDC, @gDesignTabs(i).rcTab, ghTopTabs.hBackBrushHot )
         else
            SetBkColor( memDC, ghTopTabs.BackColor )
            SetTextColor( memDC, ghTopTabs.ForeColor )
            FillRect( memDC, @gDesignTabs(i).rcTab, ghTopTabs.hBackBrush )
         end if
         
         ' display the tab text
         dim as long wsStyle = DT_NOPREFIX or DT_CENTER Or DT_VCENTER or DT_SINGLELINE
         SelectObject(memDC, ghMenuBar.hFontMenuBar)
         DrawText( memDC, gDesignTabs(i).wszText.sptr, -1, Cast(lpRect, @gDesignTabs(i).rcText), wsStyle )
         
         ' Draw the righthand side black divider
         SelectPen( memDC, hPenSolid )
         MoveToEx( memDC, gDesignTabs(i).rcTab.Right - penWidth, gDesignTabs(i).rcTab.top, null )
         LineTo( memDC, gDesignTabs(i).rcTab.Right - penWidth, gDesignTabs(i).rcTab.bottom )   
         
      next
   
      BitBlt( hDC, 0, 0, nWidth, nHeight, memDC, 0, 0, SRCCOPY )

      ' Cleanup
      if hPenSolid then DeleteObject( hPenSolid )
      if hPenNull  then DeleteObject( hPenNull )
      if hbit  then DeleteObject( SelectObject(memDC, hbit) )
      if memDC then DeleteDC( memDC )
      RestoreDC( hDC, -1 )

      EndPaint( hWnd, @ps )
   
   End Select
   
   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function



' ========================================================================================
' frmDesignTabs_Show
' ========================================================================================
Function frmDesignTabs_Show( ByVal hWndParent As HWnd ) As LRESULT

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   HWND_FRMMAIN_DESIGNTABS = pWindow->Create( hWndParent, "", @frmDesignTabs_WndProc, _
        0, 0, 0, DESIGNTABS_HEIGHT, _
        WS_CHILD Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   ' Disable background erasing by only assigning the one style
   pWindow->ClassStyle = CS_DBLCLKS

   ' Add the two tabs. Design and Code views.
   gDesignTabs(0).wszText = L(327, "Design")
   gDesignTabs(1).wszText = L(328, "Code")
   
   frmDesignTabs_PositionWindows()
   
   Function = 0
End Function

