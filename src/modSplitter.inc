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


' ========================================================================================
' Process WM_LBUTTONDOWN message for splitter
' ========================================================================================
Function Splitter_OnLButtonDown() As BOOLEAN

   DIM pWindow AS CWindow PTR = AfxCWindowPtr(HWND_FRMMAIN)
   If pWindow = 0 Then Exit Function
   
   Dim As Rect rc
   Dim As Point pt 
   
   
   ' TEST #1: (LEFT/RIGHT SPLITTER) Get the windowrect of the frmExplorer.
   GetWindowRect HWND_FRMEXPLORER, @rc
   
   ' Adjust the left edge of this rect
   rc.Left = rc.Right - pWindow->ScaleX(3)
   
   ' Now test to see if the cursor is within this rect (screen coordinates)
   GetCursorPos(@pt)
   If PtInRect( @rc, pt ) Then
      SetCursor LoadImageW( Null, MAKEINTRESOURCEW(OCR_SIZEWE), IMAGE_CURSOR, 0, 0, LR_DEFAULTSIZE Or LR_SHARED )
      gApp.bDragActive = True 
      gApp.hWndPanel   = HWND_FRMEXPLORER
      SetCapture( HWND_FRMMAIN )
      Exit Function
   End If
   

   ' TEST #2: (TOP/BOTTOM SPLITTER) Get the windowrect of the frmOutput.
   GetWindowRect HWND_FRMOUTPUT, @rc
   
   ' Adjust the top edge of this rect
   rc.Bottom = rc.Top + pWindow->ScaleY(3)
   
   ' Now test to see if the cursor is within this rect (screen coordinates)
   GetCursorPos(@pt)
   If PtInRect( @rc, pt ) Then
      SetCursor LoadImageW( Null, MAKEINTRESOURCEW(OCR_SIZENS), IMAGE_CURSOR, 0, 0, LR_DEFAULTSIZE Or LR_SHARED )
      gApp.bDragActive = True 
      gApp.hWndPanel   = HWND_FRMOUTPUT
      SetCapture( HWND_FRMMAIN )
      Exit Function
   End If
   
   gApp.bDragActive = False 
   
   SetCursor( LoadCursor( null, IDC_ARROW ))
   Function = 0
End Function


' ========================================================================================
' Process WM_LBUTTONUP message for splitter
' ========================================================================================
Function Splitter_OnLButtonUp() As BOOLEAN
   if gApp.bDragActive then
      ReleaseCapture()
      gApp.bDragActive = False 
      gApp.hWndPanel = 0
   end if
   Function = 0
End Function


' ========================================================================================
' Process WM_MOUSEMOVE message for splitter
' ========================================================================================
Function Splitter_OnMouseMove() As BOOLEAN

   ' Determine the right edge of the window and see if the mouse is currently
   ' hovering over it. Can't use WM_NCHITTEST because window not created with resizable borders.
   
   DIM pWindow AS CWindow PTR = AfxCWindowPtr(HWND_FRMMAIN)
   If pWindow = 0 Then Exit Function
 
   Dim As Rect rc
   Dim As Point pt 
   
   
   ' TEST #1 (LEFT/RIGHT SPLITTER)
   ' Get the windowrect of the frmExplorer.
   GetWindowRect HWND_FRMEXPLORER, @rc
   
   ' Adjust the left edge of this rect
   rc.Left = rc.Right - pWindow->ScaleX(3)
   
   ' Now test to see if the cursor is within this rect (screen coordinates)
   GetCursorPos(@pt)
   
   If PtInRect( @rc, pt ) Then
      SetCursor LoadImageW( Null, MAKEINTRESOURCEW(OCR_SIZEWE), IMAGE_CURSOR, 0, 0, LR_DEFAULTSIZE Or LR_SHARED )
   End If
   If gApp.bDragActive Then
      If gApp.hWndPanel = HWND_FRMEXPLORER Then
         ' Get the current rect of the frmExplorer and compare right side to the current cursor position
         ' and then move the right side to equal the cursor position.
         Dim As Long nDiff = pt.x - rc.Right
         GetClientRect HWND_FRMEXPLORER, @rc
         Dim As Long nWidth = (rc.Right - rc.Left + nDiff)
         SetWindowPos HWND_FRMEXPLORER, 0, 0, 0, nWidth, rc.Bottom - rc.Top, SWP_NOMOVE Or SWP_NOZORDER
         frmMain_PositionWindows(HWND_FRMMAIN)
         AfxDoEvents   ' allow screen to repaint
         if ( GetAsyncKeyState(VK_LBUTTON) and &H8000 )=0 then 
            Splitter_OnLButtonUp()
         endif
         Exit Function
      End If
   End If


   ' TEST #2 (TOP/BOTTOM SPLITTER)
   ' Get the windowrect of the frmOutput.
   GetWindowRect HWND_FRMOUTPUT, @rc
   
   ' Adjust the top edge of this rect
   rc.Bottom = rc.Top + pWindow->ScaleY(3)
   
   ' Now test to see if the cursor is within this rect (screen coordinates)
   GetCursorPos(@pt)
   
   If PtInRect( @rc, pt ) Then
      SetCursor LoadImageW( Null, MAKEINTRESOURCEW(OCR_SIZENS), IMAGE_CURSOR, 0, 0, LR_DEFAULTSIZE Or LR_SHARED )
   End If
   If gApp.bDragActive Then
      If gApp.hWndPanel = HWND_FRMOUTPUT Then
         Dim As Long nDiff = pt.y - rc.Top
         ' Adjust the height. The positioning will be taken care of in PositionMainWindows().
         GetClientRect HWND_FRMOUTPUT, @rc
         rc.Top = rc.Top + nDiff
         SetWindowPos HWND_FRMOUTPUT, 0, 0, 0, rc.Right - rc.Left, rc.Bottom-rc.Top, SWP_NOMOVE Or SWP_NOZORDER
         frmMain_PositionWindows(HWND_FRMMAIN)
         if ( GetAsyncKeyState(VK_LBUTTON) and &H8000 )=0 then 
            Splitter_OnLButtonUp()
         endif
         Exit Function
      End If
   End If

   Function = 0
End Function














