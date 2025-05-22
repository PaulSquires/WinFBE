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


' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
function frmMenuBar_PositionWindows() As LRESULT
   
   DIM pWindow AS CWindow PTR = AfxCWindowPtr(HWND_FRMMAIN)
   If pWindow = 0 Then Exit Function

   dim as HWND hCtrl
   dim as long nLeft, nTop, nWidth, nHeight
   nLeft = pWindow->ScaleX(6)
   nTop = 0
   nHeight = AfxGetWindowHeight(HWND_FRMMAIN_MENUBAR)

   ' position all of the child label buttons
   for i as long = IDC_MENUBAR_FILE to IDC_MENUBAR_HELP
      hCtrl = GetDlgItem( HWND_FRMMAIN_MENUBAR, i )
      nWidth = AfxGetWindowWidth(hCtrl)
      SetWindowPos( hCtrl, 0, nLeft, nTop, nWidth, nHeight, SWP_SHOWWINDOW Or SWP_NOZORDER )
      AfxRedrawWindow( hCtrl )
      nLeft = nLeft + nWidth
   next
   
   Function = 0
End Function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmMenuBar
' ========================================================================================
function frmMenuBar_OnSize( _
            ByVal HWnd As HWnd, _
            ByVal state As UINT, _
            ByVal cx As Long, _
            ByVal cy As Long _
            ) As LRESULT
   If state <> SIZE_MINIMIZED Then
      ' Position all of the child windows
      frmMenuBar_PositionWindows
   End If
   Function = 0
End Function


' ========================================================================================
' frmMenuBar Window procedure
' ========================================================================================
function frmMenuBar_WndProc( _
            ByVal HWnd   As HWnd, _
            ByVal uMsg   As UINT, _
            ByVal wParam As WPARAM, _
            ByVal lParam As LPARAM _
            ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_SIZE, frmMenuBar_OnSize)
   
   case WM_ERASEBKGND
      return true

   case WM_PAINT
      Dim As PAINTSTRUCT ps
      Dim As HDC hDc
      Dim As Rect rc

      hDC = BeginPaint(hWnd, @ps)
      GetClientRect(HWnd, @rc)
      FillRect( hDC, @rc, ghMenuBar.hPanelBrush )
      EndPaint hWnd, @ps
   
   case WM_DRAWITEM
      Dim memDC as HDC      ' Double buffering
      Dim hbit As HBITMAP   ' Double buffering
      Dim As RECT rc

      dim lpdis As DRAWITEMSTRUCT Ptr = cast( DRAWITEMSTRUCT Ptr, lParam )
      if lpdis = 0 then exit function

      DIM pWindow AS CWindow PTR = AfxCWindowPtr(lpdis->hwndItem)
      If pWindow = 0 Then Exit Function
      
      if ( lpdis->itemAction = ODA_DRAWENTIRE ) orelse _
         ( lpdis->itemAction = ODA_SELECT ) orelse _
         ( lpdis->itemAction = ODA_FOCUS ) then

         GetClientRect( lpdis->hwndItem, @rc )
         
         dim as HFONT oldFont
         dim as HBITMAP oldBmp

         SaveDC(lpdis->hDC)

         memDC = CreateCompatibleDC( lpdis->hDC )
         hbit  = CreateCompatibleBitmap( lpdis->hDC, rc.right, rc.bottom )
         
         SaveDC(memDC)
         oldBmp = SelectObject( memDC, hbit )
         oldFont = SelectObject( memDC, ghMenuBar.hFontMenuBar )
         
         dim as boolean IsHot = false
         if lpdis->hwndItem = ghWndActiveMenuBarButton then IsHot = true

         ' Paint the entire background
         FillRect( memDC, @rc, iif( IsHot, ghMenuBar.hBackBrushHot, ghMenuBar.hBackBrush) )

         ' Prepare and paint the text coloring
         SetBkColor( memDC, iif( IsHot, ghMenuBar.BackColorHot, ghMenuBar.BackColor) )   
         SetTextColor( memDC, iif( IsHot, ghMenuBar.ForeColorHot, ghMenuBar.ForeColor) )

         dim as long wsStyle = DT_NOPREFIX or DT_CENTER Or DT_VCENTER or DT_SINGLELINE
         
         ' pad the drawing rectangle to allow left and right margins
         dim as RECT rcText = lpdis->rcItem
         rcText.left = rcText.left + pWindow->ScaleX(4) 
         rcText.right = rcText.right - pWindow->ScaleX(4) 
         
         dim as CWSTR wszCaption = AfxGetWindowText( lpdis->hwndItem )
         DrawText( memDC, wszCaption.sptr, -1, Cast(lpRect, @rcText), wsStyle )
         SelectObject( memDC, oldFont )
         
         BitBlt lpdis->hDC, 0, 0, rc.right, rc.bottom, memDC, 0, 0, SRCCOPY 

         SelectObject( memDC, oldBmp )

         ' Cleanup
         RestoreDC(memDC, -1)
         if hbit then DeleteObject(hbit)
         If memDC Then DeleteDC(memDC)
         RestoreDC(lpdis->hDC, -1)
         return true
      end if   

   End Select
   
   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmMenuBar_SubclassProc 
' ========================================================================================
function frmMenuBar_SubclassProc ( _
            ByVal hWin   As HWnd, _                 ' // Control window handle
            ByVal uMsg   As UINT, _                 ' // Type of message
            ByVal _wParam As WPARAM, _               ' // First message parameter
            ByVal _lParam As LPARAM, _               ' // Second message parameter
            ByVal uIdSubclass As UINT_PTR, _        ' // The subclass ID
            ByVal dwRefData As DWORD_PTR _          ' // Pointer to reference data
            ) As LRESULT

   DIM pWindow AS CWindow PTR = AfxCWindowPtr(hWin)

   Select Case uMsg
      Case WM_GETDLGCODE
         ' Do not filter any types of messages. We want the dialog box manager to be
         ' able to correctly handle TAB, arrows, and focus rectangles, etc.
      
      Case WM_MOUSEMOVE
         ' If we have moused over a new button then redraw the previous one so the hot
         ' state can be properly rendered. We do it here because WM_MOUSELEAVE can not 
         ' be depended on to fire if we move quickly between adjacent menubar buttons.
         dim as HWND prevCtrl = 0
         if ghWndActiveMenuBarButton then prevCtrl = ghWndActiveMenuBarButton
         ghWndActiveMenuBarButton = hWin
         if prevCtrl then AfxRedrawWindow(prevCtrl)   

         IF ghWndActiveMenuBarButton <> prevCtrl THEN
            ' A new menubar button has been moused over. If a previous popup
            ' menu exists then we will automatically continue to show a popup menu
            ' for the newly selected menubar button.
            if IsWindowVisible(HWND_MENU(0)) then
               frmPopupMenu_Show( ID_POPUP, 0, hWin )
            end if   
            ' Track that we are over the control in order to catch the 
            ' eventual WM_MOUSELEAVE event
            Dim tme As TrackMouseEvent
            tme.cbSize = Sizeof(TrackMouseEvent)
            tme.dwFlags = TME_HOVER Or TME_LEAVE
            tme.hwndTrack = hWin
            TrackMouseEvent(@tme) 
         End If
         ' refresh the label so that the Hot state can draw correctly
         AfxRedrawWindow(hWin)
         
      case WM_MOUSELEAVE
         if IsWindow(HWND_MENU(0)) = 0 then
            ghWndActiveMenuBarButton = 0
            AfxRedrawWindow(hWin)
         end if   
         
      case WM_LBUTTONDOWN
         ' activate the popup menu or toggle an existing one off 
         if IsWindow(HWND_MENU(0)) then
            killPopupMenus()
            killPopupSubMenus()
         else   
            frmPopupMenu_Show( ID_POPUP, 0, hWin )
         end if
         AfxRedrawWindow(hWin)
      
      Case WM_DESTROY
         ' REQUIRED: Remove control subclassing
         RemoveWindowSubclass( hWin, @frmMenuBar_SubclassProc, uIdSubclass )
   End Select
    
   ' For messages that we don't deal with
   Function = DefSubclassProc(hWin, uMsg, _wParam, _lParam)

End Function

      
' ========================================================================================
' frmMenuBar_CreateButton
' ========================================================================================
function frmMenuBar_CreateButton( _
            byval isFirstControl as boolean, _
            byval pWindow as CWindow ptr, _
            byval CtrlID as long _
            ) As HWND

   dim as long dwStyle = WS_CHILD Or SS_NOTIFY or SS_OWNERDRAW or WS_VISIBLE 
   dim wszText as wstring * 200 = getMenuText(CtrlID)
   dim as long nWidth = getTextWidth( HWND_FRMMAIN_MENUBAR, wszText, ghMenuBar.hFontMenuBar, 6 )

   if isFirstControl then dwStyle = dwStyle or WS_GROUP 
   function = pWindow->AddControl ( "LABEL", , CtrlID, wszText, 0, 0, nWidth, 0, dwStyle, 0, 0, _
               Cast(SUBCLASSPROC, @frmMenuBar_SubclassProc), CtrlID, Cast(DWORD_PTR, 0) )

end function

' ========================================================================================
' frmMenuBar_Show
' ========================================================================================
function frmMenuBar_Show( ByVal hWndParent As HWnd ) As LRESULT

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   HWND_FRMMAIN_MENUBAR = pWindow->Create( hWndParent, "", @frmMenuBar_WndProc, _
        0, 0, 0, MENUBAR_HEIGHT, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   ' Disable background erasing by only assigning the one style
   pWindow->ClassStyle = CS_DBLCLKS
   
   frmMenuBar_CreateButton( true, pWindow, IDC_MENUBAR_FILE )
   frmMenuBar_CreateButton( false, pWindow, IDC_MENUBAR_EDIT )
   frmMenuBar_CreateButton( false, pWindow, IDC_MENUBAR_SEARCH )
   frmMenuBar_CreateButton( false, pWindow, IDC_MENUBAR_VIEW )
   frmMenuBar_CreateButton( false, pWindow, IDC_MENUBAR_PROJECT )
   frmMenuBar_CreateButton( false, pWindow, IDC_MENUBAR_COMPILE )
   frmMenuBar_CreateButton( false, pWindow, IDC_MENUBAR_DESIGNER )
   frmMenuBar_CreateButton( false, pWindow, IDC_MENUBAR_HELP )

   buildTopMenuDefinitions()
   
   Function = 0
   
End Function

