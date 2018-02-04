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

'    VISUAL DESIGNER ROUTINES
'

''
''
function DrawGrabHandles( byval hDC as HDC, _
                          byval pDoc as clsDocument ptr, _ 
                          byval bFormOnly as Boolean _
                          ) as long
   if pDoc = 0 THEN exit function
      
   dim pCtrl as clsControl ptr 

   SaveDC hDC

   dim as LOGBRUSH LogBrush
   LogBrush.lbColor = BGR(0,0,0)
   LogBrush.lbStyle = PS_SOLID
   dim as HPEN hDottedPen = ExtCreatePen( PS_COSMETIC or PS_ALTERNATE, 1, @LogBrush, 0, NULL )
   dim as HPEN hSolidPen = CreatePen(PS_SOLID, 1, BGR(0,0,0))
   dim as HBRUSH hWhiteBrush = CreateSolidBrush(BGR(255,255,255))
   dim as HBRUSH hBlackBrush = CreateSolidBrush(BGR(0,0,0))
   
   dim as long cx = AfxScaleX(8)
   dim as long cy = AfxScaleY(8)
   dim as long HMid, VMid    ' horiz and vert middles
   
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)

      if bFormOnly THEN
         if pCtrl->ControlType <> CTRL_FORM THEN continue for
      end if

      dim rc as RECT
      GetWindowRect(pCtrl->hWindow, @rc)
      MapWindowPoints(0, GetParent(pCtrl->hWindow), cast(point ptr, @rc), 2)
      
      ' Set all of the grab handles to empty
      for i as long = GRAB_TOPLEFT to GRAB_LEFT
         SetRectEmpty(@pCtrl->rcHandles(i))
      NEXT

      ' Calculate the grab handle rectangles
      HMid = (rc.right - rc.left) / 2
      VMid = (rc.bottom - rc.top) / 2
      '
      '      1     2     3
      '
      '      8           4
      '
      '      7     6     5
      '
      SetRect(@pCtrl->rcHandles(GRAB_TOPLEFT), rc.left-cx, rc.top-cy, rc.left, rc.top)
      SetRect(@pCtrl->rcHandles(GRAB_TOP), HMid-(cx/2), rc.top-cy, HMid+(cx/2), rc.top)
      SetRect(@pCtrl->rcHandles(GRAB_TOPRIGHT), rc.right, rc.top-cy, rc.right+cx, rc.top)
      SetRect(@pCtrl->rcHandles(GRAB_RIGHT), rc.right, VMid-(cy/2), rc.right+cx, VMid+(cy/2))
      SetRect(@pCtrl->rcHandles(GRAB_BOTTOMRIGHT), rc.right, rc.bottom, rc.right+cx, rc.bottom+cy)
      SetRect(@pCtrl->rcHandles(GRAB_BOTTOM), HMid-(cx/2), rc.bottom, HMid+(cx/2), rc.bottom+cy)
      SetRect(@pCtrl->rcHandles(GRAB_BOTTOMLEFT), rc.left-cx, rc.bottom, rc.left, rc.bottom+cy)
      SetRect(@pCtrl->rcHandles(GRAB_LEFT), rc.left-cx, VMid-(cy/2), rc.left, VMid+(cy/2))

      if pCtrl->ControlType = CTRL_FORM THEN
         SetRectEmpty(@pCtrl->rcHandles(GRAB_TOPLEFT))
         SetRectEmpty(@pCtrl->rcHandles(GRAB_TOP))
         SetRectEmpty(@pCtrl->rcHandles(GRAB_TOPRIGHT))
         SetRectEmpty(@pCtrl->rcHandles(GRAB_BOTTOMLEFT))
         SetRectEmpty(@pCtrl->rcHandles(GRAB_LEFT))
      end if
         
      ' Draw the actual grab handles
      if pCtrl->IsSelected THEN
         ' Draw the dotted rectangle around the control
         dim as long nOffset = AfxScaleX(2)
         SelectObject( hDC, hDottedPen )
         SelectObject( hDC, GetStockObject( NULL_BRUSH ) )
         Rectangle(hDC, rc.left-nOffset, rc.top-nOffset, rc.right+nOffset, rc.bottom+nOffset)
      
         SelectObject( hDC, hSolidPen )
         SelectObject( hDC, iif(pCtrl->IsActive, hWhiteBrush, hBlackBrush) )
         for i = GRAB_TOPLEFT to GRAB_LEFT
            if pCtrl->IsActive THEN
               RoundRect(hDC, pCtrl->rcHandles(i).left, pCtrl->rcHandles(i).top, _
                              pCtrl->rcHandles(i).right, pCtrl->rcHandles(i).bottom, 2, 2 )
            else
               Rectangle(hDC, pCtrl->rcHandles(i).left, pCtrl->rcHandles(i).top, _
                              pCtrl->rcHandles(i).right, pCtrl->rcHandles(i).bottom )
            END IF
         NEXT
      END IF
                              
   NEXT

   RestoreDC hDC, -1 

   DeleteObject(hDottedPen)
   DeleteObject(hSolidPen)
   DeleteObject(hWhiteBrush)
   DeleteObject(hBlackBrush)
   
   function = 0   
end function


' ========================================================================================
' Process WM_PAINT message for Visual Designer Form
' ========================================================================================
Function DesignerForm_OnPaint( ByVal HWnd As HWnd) As LRESULT

   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)
   If pWindow = 0 Then Exit Function
   
   Dim As PAINTSTRUCT ps
   Dim As HDC hDC

   hDC = BeginPaint(hWnd, @ps)

   SaveDC hDC

   RestoreDC hDC, -1 
   
   EndPaint hWnd, @ps
   
   Function = 0
End Function

' ========================================================================================
' Process WM_DESTROY message for Visual Designer Form
' ========================================================================================
Function DesignerForm_OnDestroy(HWnd As HWnd) As LRESULT
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)
   if pWindow then Delete pWindow
   Function = 0
End Function


' ========================================================================================
' Visual Designer Form Window procedure 
' ========================================================================================
Function DesignerForm_WndProc( ByVal HWnd   As HWnd, _
                               ByVal uMsg   As UINT, _
                               ByVal wParam As WPARAM, _
                               ByVal lParam As LPARAM _
                               ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_PAINT,   DesignerForm_OnPaint)
      HANDLE_MSG (HWnd, WM_DESTROY, DesignerForm_OnDestroy)

      Case WM_NCLBUTTONDOWN, WM_NCRBUTTONDOWN  
         ' Will only fire when the Caption bar, Min/Max/Close buttons are clicked.
         SetActiveWindow HWND_FRMMAIN

         'GetPointers hWnd, pForm, pCtrl
         'DeselectAllControls pForm
         '@pForm.NumSelectedControls = 0
         'HighlightSelectedControls pForm
         'RepaintThunderClient pForm
         Function = TRUE: Exit Function
         

      Case WM_NCLBUTTONDBLCLK
         SetActiveWindow HWND_FRMMAIN
         ' Prevent the user from double clicking on the caption bar thus
         ' causing the Form to maximize.
         Function = TRUE: Exit Function


      Case WM_NCHITTEST
         ' Catch certain critical mouseover points on the form so we can stop processing them.
         dim as LRESULT nHitTest = DefWindowProc(hWnd, uMsg, wParam, lParam)
         
         Select Case nHitTest
            ' Border edges of the window and captionbar
            Case HTLEFT, HTTOP, HTTOPLEFT, HTTOPRIGHT, _
                 HTBOTTOMLEFT, HTRIGHT, HTBOTTOM, HTBOTTOMRIGHT, _
                 HTCLOSE, HTMENU, HTMINBUTTON, HTMAXBUTTON 
                 Function = 0  ' Return zero so the mousepointer will not change
                 Exit Function                                    
         End Select      
                                      
         Function = nHitTest    ' Return the default code from the default window handler.
         Exit Function
                              
      Case WM_WINDOWPOSCHANGING 
         'prevent the Form from being dragged (moved) via the mouse. But allow the Form
         'to move when the ViewPort scrollbars are used.
         'If g.ScrollViewPort = %FALSE Then
         '   pWP = lParam
         '   @pWP.flags = @pWP.flags Or %SWP_NOMOVE 
         'End If

   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' Change the mouse cursor if over a valid grab handle
' ========================================================================================
function SetGrabHandleMouseCursor( byval pDoc as clsDocument ptr, _
                                   byval x as long, _
                                   byval y as long _
                                   ) as LRESULT
   
   dim pCtrl as clsControl ptr
   dim pt as point: pt.x = x: pt.y = y
   
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)
      for ii as long = GRAB_TOPLEFT to GRAB_LEFT
         if PtInRect(@pCtrl->rcHandles(ii), pt) then
            select case ii
               Case GRAB_TOP, GRAB_BOTTOM:          SetCursor( LoadCursor(Null, ByVal IDC_SIZENS) )
               case GRAB_LEFT, GRAB_RIGHT:          SetCursor( LoadCursor(Null, ByVal IDC_SIZEWE) )
               case GRAB_TOPLEFT, GRAB_BOTTOMRIGHT: SetCursor( LoadCursor(Null, ByVal IDC_SIZENWSE) )
               case GRAB_TOPRIGHT, GRAB_BOTTOMLEFT: SetCursor( LoadCursor(Null, ByVal IDC_SIZENESW) )
            end select
            return ii   ' return hit code GRAB_TOPLEFT to GRAB_LEFT         
         end if   
      NEXT
      exit for
   next

   function = GRAB_NOHIT
      
end function


' ========================================================================================
' Process WM_LBUTTONDOWN message for Visual Designer Frame
' ========================================================================================
function DesignerFrame_OnLButtonDown( ByVal HWnd As HWnd, _
                                      byval fDoubleClick as Boolean, _
                                      byval x as long, _
                                      byval y as long, _
                                      byval keyFlags as ulong _
                                      ) as LRESULT

   ' Change the mouse cursor if over one of the Form's grab handles
   dim pDoc as clsDocument ptr = gApp.GetDocumentPtrByWindow(hwnd)
   if pDoc = 0 THEN exit function
   
   pDoc->GrabHit = SetGrabHandleMouseCursor(pDoc, x, y)
   if pDoc->GrabHit <> GRAB_NOHIT THEN
      pDoc->bSizing = true
      GetWindowRect(hwnd, @pDoc->rcSize)
      SetCapture(hwnd)
   END IF

   ' Save the current mouse position
   pDoc->ptPrev.x = x
   pDoc->ptPrev.y = y
   
   function = 0
end function

' ========================================================================================
' Process WM_LBUTTONUP message for Visual Designer Frame
' ========================================================================================
function DesignerFrame_OnLButtonUp( ByVal HWnd As HWnd, _
                                    byval x as long, _
                                    byval y as long, _
                                    byval keyFlags as ulong _
                                    ) as LRESULT

   ' Change the mouse cursor if over one of the Form's grab handles
   dim pDoc as clsDocument ptr = gApp.GetDocumentPtrByWindow(hwnd)
   if pDoc = 0 THEN exit function
   
   pDoc->GrabHit = GRAB_NOHIT
   pDoc->bSizing = false
   ReleaseCapture
   
   function = 0
end function

' ========================================================================================
' Process WM_MOUSEMOVE message for Visual Designer Frame
' ========================================================================================
function DesignerFrame_OnMouseMove( ByVal HWnd As HWnd, _
                                    byval x as long, _
                                    byval y as long, _
                                    byval keyFlags as ulong _
                                    ) as LRESULT

   ' Change the mouse cursor if over one of the Form's grab handles
   dim pDoc as clsDocument ptr = gApp.GetDocumentPtrByWindow(hwnd)
   if pDoc = 0 THEN exit function
   
   if pDoc->bSizing THEN
      Select Case pDoc->GrabHit
         Case GRAB_BOTTOMRIGHT 
            pDoc->rcSize.right  = pDoc->rcSize.right + (x - pDoc->ptPrev.x)
            pDoc->rcSize.bottom = pDoc->rcSize.bottom + (y - pDoc->ptPrev.y)
         Case GRAB_RIGHT 
            pDoc->rcSize.right  = pDoc->rcSize.right + (x - pDoc->ptPrev.x)
         Case GRAB_BOTTOM
            pDoc->rcSize.bottom = pDoc->rcSize.bottom + (y - pDoc->ptPrev.y)
      End Select
      
      ' Save the current mouse position
      pDoc->ptPrev.x = x
      pDoc->ptPrev.y = y

      ' NOTE:
      ' Set our control to a default minimum value If Zero.
      ' We want to do this so we don't loose visibility of our control and we can still 
      ' see the handles when selected.

      ' Check for a minimum width and height
      If pDoc->rcSize.right - pDoc->rcSize.left <= AfxScaleX(8) Then pDoc->rcSize.right = pDoc->rcSize.left + AfxScaleX(8)
      If pDoc->rcSize.bottom - pDoc->rcSize.top <= AfxScaleY(8) Then pDoc->rcSize.bottom = pDoc->rcSize.top + AfxScaleY(8)

      ' Map the window based rcSize to client coordinates
      MapWindowPoints(0, hWnd, cast(point ptr, @pDoc->rcSize), 2)
      
      ' Move the form/control to its new position
      SetWindowPos( pDoc->hWndForm, 0, pDoc->rcSize.left, pDoc->rcSize.top, _
                    pDoc->rcSize.right - pDoc->rcSize.left, pDoc->rcSize.bottom - pDoc->rcSize.top, _
                    SWP_NOZORDER)
      
   else
      pDoc->GrabHit = SetGrabHandleMouseCursor(pDoc, x, y)
   end if      

   function = 0
end function


' ========================================================================================
' Process WM_PAINT message for Visual Designer Frame
' ========================================================================================
Function DesignerFrame_OnPaint( ByVal HWnd As HWnd) As LRESULT

   Dim As PAINTSTRUCT ps
   Dim As HDC hDC

   hDC = BeginPaint(hWnd, @ps)

   dim pDoc as clsDocument ptr = gApp.GetDocumentPtrByWindow(hwnd)
   if pDoc THEN
      ' Draw grab handles for the form only
      DrawGrabHandles(hDC, pDoc, true)
   END IF

   EndPaint hWnd, @ps
   
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for Visual Designer frame
' ========================================================================================
Function DesignerFrame_OnDestroy(HWnd As HWnd) As LRESULT
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)
   if pWindow then Delete pWindow
   Function = 0
End Function


' ========================================================================================
' Visual Designer Frame Window procedure (the container frame window)
' ========================================================================================
Function DesignerFrame_WndProc( ByVal HWnd   As HWnd, _
                                ByVal uMsg   As UINT, _
                                ByVal wParam As WPARAM, _
                                ByVal lParam As LPARAM _
                                ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_LBUTTONDOWN, DesignerFrame_OnLButtonDown)
      HANDLE_MSG (HWnd, WM_LBUTTONUP,   DesignerFrame_OnLButtonUp)
      HANDLE_MSG (HWnd, WM_MOUSEMOVE,   DesignerFrame_OnMouseMove)
      HANDLE_MSG (HWnd, WM_PAINT,       DesignerFrame_OnPaint)
      HANDLE_MSG (HWnd, WM_DESTROY,     DesignerFrame_OnDestroy)
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' Process WM_DESTROY message for Visual Designer main window
' ========================================================================================
Function DesignerMain_OnDestroy(HWnd As HWnd) As LRESULT
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)
   if pWindow then Delete pWindow
   Function = 0
End Function


' ========================================================================================
' Visual Designer Main (hWindow) Window procedure 
' ========================================================================================
Function DesignerMain_WndProc( ByVal HWnd   As HWnd, _
                               ByVal uMsg   As UINT, _
                               ByVal wParam As WPARAM, _
                               ByVal lParam As LPARAM _
                               ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_DESTROY, DesignerMain_OnDestroy)

      CASE WM_SIZE
         ' In the WM_SIZE message we need to check the validity of the pointers
         ' because this message is sent before we have had time to store them.
         DIM pWindow AS CWindow PTR = cast(CWindow PTR, GetWindowLongPtr(hwnd, 0))
         DIM pScrollWindow AS CScrollWindow PTR
         IF pWindow THEN 
            ' Resize the DesignerFrame to be same size as DesignerMain
            Dim pFrame As CWindow Ptr = AfxCWindowPtr(GetDlgItem(hWnd, IDC_DESIGNFRAME))
            if pFrame THEN pFrame->SetClientSize(pWindow->ClientWidth, pWindow->ClientHeight)
            pScrollWindow = pWindow->ScrollWindowPtr
         end if
         IF pScrollWindow THEN pScrollWindow->OnSize(wParam, lParam)
         EXIT FUNCTION

      CASE WM_VSCROLL, WM_HSCROLL
         dim as hwnd hWndFrame = GetDlgItem(hWnd, IDC_DESIGNFRAME)
         dim as hwnd hWndForm = GetDlgItem(hWndFrame, IDC_DESIGNFORM)
         DIM pScrollWindow AS CScrollWindow PTR = AfxScrollWindowPtr(hwnd)
         IF pScrollWindow THEN 
            if uMsg = WM_VSCROLL then pScrollWindow->OnVScroll(wParam, lParam)
            if uMsg = WM_HSCROLL then pScrollWindow->OnHScroll(wParam, lParam)
' TODO: Fix the scrolling because the design form does not repaint correctly.
            AfxRedrawWindow(hWndFrame)
            AfxRedrawWindow(hWndForm)
         end if   
         EXIT FUNCTION

   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


