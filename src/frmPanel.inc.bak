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
''  frmPanel.inc
''   

#include once "frmPanel.bi"
#include once "frmExplorer.bi"
#include once "frmBookmarks.bi"
#include once "frmFunctions.bi"

' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
function frmPanel_PositionWindows() as LRESULT

   dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMPANEL)
   if pWindow = 0 then exit function
   
   ' Get the entire client area
   dim as Rect rc
   GetClientRect( HWND_FRMPANEL, @rc )

   redim gPanelButton(1) as PANEL_BUTTON_TYPE
   
   gPanelButton(1).rc.left = rc.left
   gPanelButton(1).rc.right = rc.right
   gPanelButton(1).rc.bottom = rc.bottom
   gPanelButton(1).rc.top = rc.bottom - pWindow->ScaleY(TOPTABS_HEIGHT)
   
   gPanelButton(0).rc.left = rc.left
   gPanelButton(0).rc.right = rc.right
   gPanelButton(0).rc.bottom = gPanelButton(1).rc.top 
   gPanelButton(0).rc.top = gPanelButton(1).rc.top - pWindow->ScaleY(TOPTABS_HEIGHT)

   if gApp.GetDocumentCount = 0 then
      SetRectEmpty( @gPanelButton(0).rc )
      SetRectEmpty( @gPanelButton(1).rc )
   end if

   ' show/hide of the child window is done separately in order to prevent blank client area from 
   ' painting before the lists are loaded. Allow a 2 pixel right border between listbox and
   ' form edge in order to allow for panel resizing. 
   dim as long nLeft = 0
   dim as long nTop = rc.top  + pWindow->ScaleY(46)
   dim as long nWidth = rc.right  - rc.Left - pWindow->ScaleX(2)
   dim as long nHeight = gPanelButton(0).rc.top - rc.top - nTop

   SetWindowPos( HWND_FRMEXPLORER, 0, nLeft, nTop, nWidth, nHeight, SWP_NOZORDER or SWP_HIDEWINDOW )
   SetWindowPos( HWND_FRMBOOKMARKS, 0, nLeft, nTop, nWidth, nHeight, SWP_NOZORDER or SWP_HIDEWINDOW )
   SetWindowPos( HWND_FRMFUNCTIONS, 0, nLeft, nTop, nWidth, nHeight, SWP_NOZORDER or SWP_HIDEWINDOW )

   select case gPanel.hActiveChild
      case HWND_FRMEXPLORER
         gPanel.wszLabel = L(247,"EXPLORER")
         ShowWindow( HWND_FRMEXPLORER, SW_SHOW )
         gPanelVScroll.hListBox = HWND_FRMEXPLORER_LISTBOX

         gPanelButton(0).wszCaption = ucase( L(223,"Function List") )
         gPanelButton(0).hActionChild = HWND_FRMFUNCTIONS
         gPanelButton(1).wszCaption = ucase( L(188,"Bookmarks") )
         gPanelButton(1).hActionChild = HWND_FRMBOOKMARKS
      
      case HWND_FRMBOOKMARKS
         gPanel.wszLabel = ucase( L(188,"Bookmarks") )
         ShowWindow( HWND_FRMBOOKMARKS, SW_SHOW )
         gPanelVScroll.hListBox = HWND_FRMBOOKMARKS_LISTBOX
      
         gPanelButton(0).wszCaption = L(247,"EXPLORER")
         gPanelButton(0).hActionChild = HWND_FRMEXPLORER
         gPanelButton(1).wszCaption = ucase( L(223,"Function List") )
         gPanelButton(1).hActionChild = HWND_FRMFUNCTIONS

      case HWND_FRMFUNCTIONS
         gPanel.wszLabel = ucase( L(223,"Function List") )
         ShowWindow( HWND_FRMFUNCTIONS, SW_SHOW )
         gPanelVScroll.hListBox = HWND_FRMFUNCTIONS_LISTBOX

         gPanelButton(0).wszCaption = L(247,"EXPLORER")
         gPanelButton(0).hActionChild = HWND_FRMEXPLORER
         gPanelButton(1).wszCaption = ucase( L(188,"Bookmarks") )
         gPanelButton(1).hActionChild = HWND_FRMBOOKMARKS
   end select

   ' calculate the rect for the panel label
   with gPanel.rcLabel 
      .left = pWindow->ScaleX( 10 )
      .top = pWindow->ScaleY( 10 )
      .bottom = .top + pWindow->ScaleY( 20 )
      .right = .left + _
                pWindow->ScaleX( _
                   getTextWidth(HWND_FRMPANEL, gPanel.wszLabel, ghStatusBar.hFontStatusBar, 10) _
                )
   end with

   ' calculate the actual more actions "..." button itself
   dim as long ActionButtonWidth = 24
   dim as long ActionButtonHeight = 22
   with gPanel.rcActionMenu
      .left = rc.right - pWindow->ScaleX(ActionButtonWidth+10)
      .top = pWindow->ScaleY(7)
      .right = .left + pWindow->ScaleX(ActionButtonWidth)
      .bottom = .top + pWindow->ScaleY(ActionButtonHeight)
   end with
   
   AfxRedrawWindow( gPanelVScroll.hListBox )
   AfxRedrawWindow( HWND_FRMPANEL )
   
   function = 0
end function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmPanel
' ========================================================================================
function frmPanel_OnSize( _
            byval HWnd as HWnd, _
            byval state as UINT, _
            byval cx as long, _
            byval cy as long _
            ) as LRESULT
   if state <> SIZE_MINIMIZED then
      ' Position all of the child windows
      frmPanel_PositionWindows
   end if
   function = 0
end function
       
' ========================================================================================
' Do hit test to determine if "..." button in action Area was clicked
' ========================================================================================
function isPanelActionButtonHitTest( byval hWin as HWnd ) as boolean
   dim as POINT pt: GetCursorPos( @pt )
   MapWindowPoints( HWND_DESKTOP, hWin, cast( POINT ptr, @pt ), 1 )
   if PtInRect( @gPanel.rcActionMenu, pt ) then return true
   function = false
end function

' ========================================================================================
' Do hit test on Button0 
' ========================================================================================
function isPanelButton0HitTest( byval hWin as HWnd ) as boolean
   dim as POINT pt: GetCursorPos( @pt )
   MapWindowPoints( HWND_DESKTOP, hWin, cast( POINT ptr, @pt ), 1 )
   if PtInRect( @gPanelButton(0).rc, pt ) then return true
   function = false
end function

' ========================================================================================
' Do hit test on Button1
' ========================================================================================
function isPanelButton1HitTest( byval hWin as HWnd ) as boolean
   dim as POINT pt: GetCursorPos( @pt )
   MapWindowPoints( HWND_DESKTOP, hWin, cast( POINT ptr, @pt ), 1 )
   if PtInRect( @gPanelButton(1).rc, pt ) then return true
   function = false
end function

' ========================================================================================
' Process WM_PAINT message for window/dialog: frmPanel
' ========================================================================================
function frmPanel_OnPaint( byval HWnd as HWnd ) as LRESULT
            
   dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMPANEL)
   if pWindow = 0 then exit function

   dim as PAINTSTRUCT ps
   dim as HDC hDc
   dim as HPEN oldPen
   dim as HFONT oldFont
   dim as HBRUSH oldBrush
   
   hDC = BeginPaint(hWnd, @ps)

   SaveDC( hDC )

   FillRect( hDC, @ps.rcPaint, ghPanel.hPanelBrush )
   
   ' Draw the panel Label
   SetTextColor( hDC, ghPanel.forecolor )
   SetBkColor( hDC, ghPanel.backcolor )
   oldFont = SelectObject( hDC, ghStatusBar.hFontStatusBar )

   dim as long wsStyle = DT_NOPREFIX or DT_LEFT or DT_VCENTER or DT_SINGLELINE  
   DrawText( hDC, gPanel.wszLabel.sptr, -1, cast(lpRect, @gPanel.rcLabel), wsStyle )
   SelectObject( hDC, oldFont )
   
   ' Draw the "..." menu item
   SetTextColor( hDC, ghPanel.ForeColorHot )
   if isPanelActionButtonHitTest( HWnd ) then
      oldBrush = SelectObject( hDC, ghPanel.hBackBrushHot )
      SetBkColor( hDC, ghPanel.BackColorHot )
   else
      oldBrush = SelectObject( hDC, ghPanel.hPanelBrush )
      SetBkColor( hDC, ghPanel.BackColor )
   end if
   dim as RECT rc = gPanel.rcActionMenu

   dim as HPEN hPenNull = CreatePen( PS_NULL, 1, 0 )  ' null/invisible pen
   oldPen = SelectObject( hDC, hPenNull )
   RoundRect( hDC, rc.left, rc.top, rc.right, rc.bottom, 20, 20 )
   DeleteObject SelectObject( hDC, oldPen )
   
   wsStyle = DT_NOPREFIX or DT_CENTER Or DT_VCENTER
   oldFont = SelectObject( hDC, ghMenuBar.hFontSymbolLargeBold )
   DrawText( hDC, wszMoreActions, -1, Cast(lpRect, @rc), wsStyle )
   SelectObject( hDC, oldFont )
   SelectObject( hDC, oldBrush )

   ' Draw the bottom two Panel buttons
   dim as long penWidth = pWindow->ScaleX(1)
   dim as HPEN hPenSolid = CreatePen( PS_SOLID, penWidth, ghTopTabs.Divider )
   oldPen = SelectObject( hDC, hPenSolid)

   dim as HBRUSH hBrush
   dim as COLORREF foreclr, backclr
   dim as boolean isHot 
   
   for i as long = lbound(gPanelButton) to ubound(gPanelButton)
      dim as RECT rc = gPanelButton(i).rc

      isHot = isMouseOverRECT( HWnd, rc )
      hBrush = iif( isHot, ghPanel.hBackBrushButtonHot, ghPanel.hBackBrushButton)
      backclr = iif( isHot, ghPanel.BackColorButtonHot, ghPanel.BackColorButton)
      foreclr = iif( isHot, ghPanel.ForeColorButtonHot, ghPanel.ForeColorButton)
      
      oldBrush = SelectObject( hDC, hBrush )
      SetBkColor( hDC, backclr )
      SetTextColor( hDC, foreclr )

      SelectPen( hDC, hPenSolid )
      Rectangle( hDC, rc.left, rc.top, rc.right, rc.bottom )
      wsStyle = DT_NOPREFIX or DT_CENTER Or DT_VCENTER or DT_SINGLELINE
      oldFont = SelectObject( hDC, ghStatusBar.hFontStatusBar )
      DrawText( hDC, gPanelButton(i).wszCaption.sptr, -1, Cast(lpRect, @rc), wsStyle )

      SelectObject( hDC, oldFont )
      SelectObject( hDC, oldBrush )
   next   

   DeleteObject SelectObject( hDC, oldPen )
   RestoreDC( hDC, -1 )

   EndPaint( hWnd, @ps )

   function = 0
end function


' ========================================================================================
' frmPanel Window procedure
' ========================================================================================
function frmPanel_WndProc( _
            byval HWnd   as HWnd, _
            byval uMsg   as UINT, _
            byval wParam as WPARAM, _
            byval lParam as LPARAM _
            ) as LRESULT

   static as boolean isPrevHotAction, isPrevHotButton0, isPrevHotButton1
   static hTooltip as HWND

   select case uMsg
      HANDLE_MSG (HWnd, WM_SIZE,  frmPanel_OnSize)
      HANDLE_MSG (HWnd, WM_PAINT, frmPanel_OnPaint)
   
   case WM_ERASEBKGND
      return true

   case WM_MOUSEMOVE
      dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMEXPLORER)
      if pWindow = 0 then exit function

      dim tme as TrackMouseEvent
      tme.cbSize = sizeof(TrackMouseEvent)
      tme.dwFlags = TME_HOVER or TME_LEAVE
      tme.hwndTrack = HWnd
      TrackMouseEvent(@tme) 

      ' Repaint the gPanelMenuRect to ensure Hot highlighting
      dim as boolean isHotAction, isHotButton0, isHotButton1 
      isHotAction = isPanelActionButtonHitTest( HWnd )
      isHotButton0 = isPanelButton0HitTest( HWnd )
      isHotButton1 = isPanelButton1HitTest( HWnd )

      if isHotAction <> isPrevHotAction then   
         AfxRedrawWindow( HWnd )
         isPrevHotAction = isHotAction
      end if   
            
      if isHotButton0 <> isPrevHotButton0 then   
         AfxRedrawWindow( HWnd )
         isPrevHotButton0 = isHotButton0
      end if   

      if isHotButton1 <> isPrevHotButton1 then   
         AfxRedrawWindow( HWnd )
         isPrevHotButton1 = isHotButton1
      end if

      ' PANELS LEFT/RIGHT SPLITTER
      dim as POINT pt
      dim as Rect rc
      GetWindowRect HWND_FRMPANEL, @rc
      rc.Left = rc.Right - pWindow->ScaleX(3)
      GetCursorPos(@pt)
      if PtInRect( @rc, pt ) then
         if WindowFromPoint(pt) = HWND_FRMPANEL then
            SetCursor( ghCursorSizeWE )
         end if
      end if

      if gApp.bDragActive then
         if gApp.hWndPanel = HWND_FRMPANEL then
            ' Get the current rect of the frmExplorer and compare right side to the current cursor position
            ' and then move the right side to equal the cursor position.
            dim as long nDiff = pt.x - rc.Right
            GetClientRect HWND_FRMPANEL, @rc
            dim as long nWidth = (rc.Right - rc.Left + nDiff)
            ' make sure width does not go below a specific size otherwise the user will not
            ' be able to grab the panel edge to resize
            nWidth = max( nWidth, pWindow->ScaleX(80) )
            SetWindowPos( HWND_FRMPANEL, 0, 0, 0, nWidth, rc.Bottom - rc.Top, SWP_NOMOVE or SWP_NOZORDER )
            frmMain_PositionWindows
            AfxRedrawWindow(HWnd)   ' ensure the More Actions menu repaints
            AfxDoEvents   ' allow screen to repaint
         end if
      end if
      
   case WM_LBUTTONDOWN
      dim pWindow as CWindow ptr = AfxCWindowPtr(HWND_FRMPANEL)
      if pWindow = 0 then exit function

      ' PANELS LEFT/RIGHT SPLITTER
      dim as Rect rc
      dim as Point pt 

      gApp.bDragActive = false 
      GetWindowRect HWND_FRMPANEL, @rc
      rc.Left = rc.Right - pWindow->ScaleX(3)
      GetCursorPos(@pt)
      if PtInRect( @rc, pt ) then
         if WindowFromPoint(pt) = HWND_FRMPANEL then
            SetCursor( ghCursorSizeWE )
            gApp.bDragActive = true 
            gApp.hWndPanel   = HWND_FRMPANEL
            SetCapture( HWND_FRMPANEL )
         end if
      end if

   case WM_LBUTTONUP
      if gApp.bDragActive then
         gApp.bDragActive = false 
         gApp.hWndPanel = 0
         ReleaseCapture()
      end if
      if isPanelActionButtonHitTest( HWnd ) then
         dim as HMENU hPopupMenu 
         select case gPanel.hActiveChild
            case HWND_FRMEXPLORER
               hPopupMenu = CreateExplorerActionButtonContextMenu()
            case HWND_FRMFUNCTIONS
               hPopupMenu = CreateFunctionsActionButtonContextMenu()
            case HWND_FRMBOOKMARKS
               hPopupMenu = CreateBookmarksActionButtonContextMenu()
         end select
         ' Popup the menu to the bottom of the Action Button (right aligned)
         dim as RECT rc = gPanel.rcActionMenu   ' work with a copy
         MapWindowPoints( HWND_FRMPANEL, HWND_DESKTOP, cast(POINT ptr, @rc), 2 )
         TrackPopupMenu( hPopUpMenu, TPM_RIGHTALIGN, _
                             rc.right, rc.bottom, 0, HWND_FRMMAIN, byval null)
         DestroyMenu( hPopUpMenu )
         Return true   ' prevent further processing that leads to WM_CONTEXTMENU
      end if
      
      if isPanelButton0HitTest( HWnd ) then
         gPanel.hActiveChild = gPanelButton(0).hActionChild
         frmPanel_PositionWindows()
      end if
      
      if isPanelButton1HitTest( HWnd ) then
         gPanel.hActiveChild = gPanelButton(1).hActionChild
         frmPanel_PositionWindows()
      end if


   case WM_MOUSELEAVE
      isPrevHotAction = false
      isPrevHotButton0 = false
      isPrevHotButton1 = false
      AfxDeleteTooltip( hTooltip, HWnd )
      hTooltip = 0
      AfxRedrawWindow(HWnd)   

   case WM_MOUSEHOVER
      dim as CWSTR wszTooltip 
      if IsWindow(hTooltip) = 0 then hTooltip = AfxAddTooltip( HWnd, "", false, false )
      if isPanelActionButtonHitTest( HWnd ) then
         wszTooltip  = L(440, "More Actions") & "..."
      end if
      ' Display the tooltip
      AfxSetTooltipText( hTooltip, HWnd, wszTooltip )
      AfxRedrawWindow( HWnd )
      
   end select

   ' for messages that we don't deal with
   function = DefWindowProc( HWnd, uMsg, wParam, lParam )

end function


' ========================================================================================
' frmPanel_Show
' ========================================================================================
function frmPanel_Show( byval hWndParent as HWnd ) as LRESULT

   '  Create the main window and child controls
   dim pWindow as CWindow ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   ' Only make the Panel initially visible if it was already visible
   ' when the most previous instance of the program closed. Also, set the width of
   ' the window to the last used visible width.
   dim as long nWidth = iif(gConfig.ShowPanel, gConfig.ShowPanelWidth, 250)
   
   HWND_FRMPANEL = pWindow->Create( hWndParent, "Panel Window", @frmPanel_WndProc, _
        0, 0, nWidth, 0, _
        WS_CHILD or iif(gConfig.ShowPanel, WS_VISIBLE, 0) or WS_CLIPSIBLINGS or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT or WS_EX_LEFT or WS_EX_LTRREADING or WS_EX_RIGHTSCROLLBAR)
   ' Disable background erasing by only assigning the one style
   pWindow->ClassStyle = CS_DBLCLKS

   ' Create the Explorer child window
   frmExplorer_Show( HWND_FRMPANEL )
   frmBookmarks_Show( HWND_FRMPANEL )
   frmFunctions_Show( HWND_FRMPANEL )
   gPanel.hActiveChild = HWND_FRMEXPLORER
   
   function = 0
   
end function
