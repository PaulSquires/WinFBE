'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2017 Paul Squires, PlanetSquires Software
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
''  Ensure that the rect's coordinates have the lowest values to the left.
''
Function NormalizeRect( ByRef rc As Rect ) As Long
   If rc.Left > rc.Right   Then Swap rc.Right, rc.Left
   If rc.Top  > rc.Bottom  Then Swap rc.Top, rc.Bottom
   function = 0
End Function


''
''  Draw the rectangle lasso on the Window DC.
''
function DrawLasso( byval pDoc as clsDocument ptr, _
                    ByVal ptStart As point, _
                    byval ptEnd as point _
                    ) as long 

   dim as rect rc
   SetRect(@rc, ptStart.x, ptStart.y, ptEnd.x, ptEnd.y)
   
   ' Normalize the Rect
   NormalizeRect(rc)
   If IsRectEmpty( @rc ) Then Exit function
    
   ' Get DC of entire screen to draw on top of all controls
   dim as HDC hDC = GetWindowDC(0)
   MapWindowPoints(pDoc->hWndForm, 0, cast(point ptr, @rc), 2)
   DrawFocusRect( hDC, @rc)
   ReleaseDC( 0, hDC )

   function = 0
End function


' ========================================================================================
' Change the mouse cursor if over a valid grab handle
' ========================================================================================
function SetGrabHandleMouseCursor( byval pDoc as clsDocument ptr, _
                                   byval x as long, _
                                   byval y as long, _
                                   byref pCtrlAction as clsControl Ptr _
                                   ) as LRESULT
   
   dim pCtrl as clsControl ptr
   dim rcCtrl as RECT
   dim pt as point: pt.x = x: pt.y = y   ' The point in is client coordinates
   
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
            pCtrlAction = pCtrl
            return ii   ' return hit code GRAB_TOPLEFT to GRAB_LEFT         
         end if 
      NEXT
   next

   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)
      if pCtrl->ControlType = CTRL_FORM THEN continue for
      GetWindowRect(pCtrl->hWindow, @rcCtrl)
      MapWindowPoints(0, pDoc->hWndForm, cast(point ptr, @rcCtrl), 2)
      if PtInRect(@rcCtrl, pt) THEN
         SetCursor( LoadCursor(Null, ByVal IDC_SIZEALL) )
         return GRAB_NOHIT
      END IF  
   next
   
   function = GRAB_NOHIT
      
end function


' ========================================================================================
' Calculate the size of form/controls grab handle rectangles 
' ========================================================================================
function CalculateGrabHandles( byval pDoc as clsDocument ptr) as long
   if pDoc = 0 THEN exit function
      
   dim pCtrl as clsControl ptr 

   dim as long cx = AfxScaleX(6)
   dim as long cy = AfxScaleY(6)
   dim as long HMid, VMid    ' horiz and vert middles
   
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)

      dim rc as RECT
      GetWindowRect(pCtrl->hWindow, @rc)
      MapWindowPoints(0, GetParent(pCtrl->hWindow), cast(point ptr, @rc), 2)
      
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
      ' Only calculate the grab handle rectangles if the control is selected
      if pCtrl->IsSelected THEN
         SetRect(@pCtrl->rcHandles(GRAB_TOPLEFT), rc.left-cx, rc.top-cy, rc.left, rc.top)
         SetRect(@pCtrl->rcHandles(GRAB_TOP), rc.left+HMid-(cx/2), rc.top-cy, rc.left+HMid+(cx/2), rc.top)
         SetRect(@pCtrl->rcHandles(GRAB_TOPRIGHT), rc.right, rc.top-cy, rc.right+cx, rc.top)
         SetRect(@pCtrl->rcHandles(GRAB_RIGHT), rc.right, rc.top+VMid-(cy/2), rc.right+cx, rc.top+VMid+(cy/2))
         SetRect(@pCtrl->rcHandles(GRAB_BOTTOMRIGHT), rc.right, rc.bottom, rc.right+cx, rc.bottom+cy)
         SetRect(@pCtrl->rcHandles(GRAB_BOTTOM), rc.left+HMid-(cx/2), rc.bottom, rc.left+HMid+(cx/2), rc.bottom+cy)
         SetRect(@pCtrl->rcHandles(GRAB_BOTTOMLEFT), rc.left-cx, rc.bottom, rc.left, rc.bottom+cy)
         SetRect(@pCtrl->rcHandles(GRAB_LEFT), rc.left-cx, rc.top+VMid-(cy/2), rc.left, rc.top+VMid+(cy/2))
      else
         for i as long = GRAB_TOPLEFT to GRAB_LEFT
            SetRectEmpty(@pCtrl->rcHandles(i))
         NEXT
      end if
      
      ' If this is a form then we don't want to display some of the grab handles
      ' so simply set them to be empty rectangles.
      if pCtrl->ControlType = CTRL_FORM THEN
         SetRectEmpty(@pCtrl->rcHandles(GRAB_TOPLEFT))
         SetRectEmpty(@pCtrl->rcHandles(GRAB_TOP))
         SetRectEmpty(@pCtrl->rcHandles(GRAB_TOPRIGHT))
         SetRectEmpty(@pCtrl->rcHandles(GRAB_BOTTOMLEFT))
         SetRectEmpty(@pCtrl->rcHandles(GRAB_LEFT))
      end if
         
   NEXT

   function = 0
end function

   
' ========================================================================================
' Draw the actual grab handles (this is called from WM_PAINT)
' ========================================================================================
function DrawGrabHandles( byval hDC as HDC, _
                          byval pDoc as clsDocument ptr, _ 
                          byval bFormOnly as Boolean _
                          ) as long
   if pDoc = 0 THEN exit function
      
   dim pCtrl as clsControl ptr 

   CalculateGrabHandles(pDoc)
   
   SaveDC hDC

   dim as LOGBRUSH LogBrush
   LogBrush.lbColor = BGR(0,0,0)
   LogBrush.lbStyle = PS_SOLID
   dim as HPEN hDottedPen = ExtCreatePen( PS_COSMETIC or PS_ALTERNATE, 1, @LogBrush, 0, NULL )
   dim as HPEN hSolidPen = CreatePen(PS_SOLID, 1, BGR(0,0,0))
   dim as HBRUSH hWhiteBrush = CreateSolidBrush(BGR(255,255,255))
   dim as HBRUSH hBlackBrush = CreateSolidBrush(BGR(0,0,0))
   
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)

      if bFormOnly THEN
         if pCtrl->ControlType <> CTRL_FORM THEN continue for
      else
         if pCtrl->ControlType = CTRL_FORM THEN continue for
      end if

      dim rc as RECT
      GetWindowRect(pCtrl->hWindow, @rc)
      MapWindowPoints(0, GetParent(pCtrl->hWindow), cast(point ptr, @rc), 2)
      
      ' Draw the actual grab handles
      if pCtrl->IsSelected THEN
         ' Draw the dotted rectangle around the control
         dim as long nOffset = AfxScaleX(2)
         SelectObject( hDC, hDottedPen )
         SelectObject( hDC, GetStockObject( NULL_BRUSH ) )
         Rectangle(hDC, rc.left-nOffset, rc.top-nOffset, rc.right+nOffset, rc.bottom+nOffset)

         SelectObject( hDC, hSolidPen )
         SelectObject( hDC, iif(pCtrl->IsActive, hWhiteBrush, hBlackBrush) )
         for ii as long = GRAB_TOPLEFT to GRAB_LEFT
            if pCtrl->IsActive THEN
               RoundRect(hDC, pCtrl->rcHandles(ii).left, pCtrl->rcHandles(ii).top, _
                              pCtrl->rcHandles(ii).right, pCtrl->rcHandles(ii).bottom, 2, 2 )
            else
               Rectangle(hDC, pCtrl->rcHandles(ii).left, pCtrl->rcHandles(ii).top, _
                              pCtrl->rcHandles(ii).right, pCtrl->rcHandles(ii).bottom )
            END IF
         NEXT
      END IF
   next                           

   RestoreDC hDC, -1 

   DeleteObject(hDottedPen)
   DeleteObject(hSolidPen)
   DeleteObject(hWhiteBrush)
   DeleteObject(hBlackBrush)
   
   function = 0   
end function


' ========================================================================================
' Handle WM_LBUTTONDOWN messages for the Form and Frame windows
' ========================================================================================
function HandleDesignerLButtonDown( ByVal HWnd As HWnd ) as LRESULT

   dim as POINT pt
   dim as long x, y
   
   dim pDoc as clsDocument ptr = gApp.GetDocumentPtrByWindow(hwnd)
   if pDoc = 0 THEN exit function

   dim pCtrl as clsControl ptr
   SetCapture(hwnd)

   GetCursorPos(@pt)
   MapWindowPoints(0, HWND, @pt, 1)
   x = pt.x
   y = pt.y

   ' #1: Determine if a grab handle has been clicked on
   pDoc->GrabHit = SetGrabHandleMouseCursor(pDoc, x, y, pCtrl)
   if pDoc->GrabHit <> GRAB_NOHIT THEN
      ' One of the sizing handles was clicked on   
      pDoc->bSizing = true
      pDoc->pCtrlAction = pCtrl
      GetWindowRect(pDoc->pCtrlAction->hWindow, @pDoc->rcSize)
   else
      ' #2: Determine what control/form was clicked on
      dim as hwnd hWndCtrl = RealChildWindowFromPoint(pDoc->hWndForm, pt)
      pCtrl = pDoc->Controls.GetCtrlPtr(hWndCtrl)
      if pCtrl then
         ' If Ctrl is held down then toggle adding/removing the control 
         if GetAsyncKeyState(VK_CONTROL) THEN
            pCtrl->IsSelected = not(pCtrl->IsSelected)
         else
            ' If the control being clicked on is already selected then it will become
            ' the actice control. If not already selected then deselect all other controls
            ' the the selection group.
            if pCtrl->IsSelected = false THEN pDoc->Controls.DeselectAllControls
            pCtrl->IsSelected = true
            ' If the Form is clicked on the start the lasso process.
            if pCtrl->ControlType = CTRL_FORM THEN
               pDoc->bLasso = true 
               pDoc->ptStart = pt
               pDoc->ptEnd = pt
            else
               pDoc->pCtrlAction = pCtrl
               pDoc->bMoving = true
            END IF
         END IF
         pDoc->Controls.SetActiveControl( iif(pCtrl->IsSelected, hWndCtrl, 0) )
      END IF
   END IF

   ' Save the current mouse position
   pDoc->ptPrev.x = x
   pDoc->ptPrev.y = y
   
   ' Ensure the grab handles of form and controls are redrawn or hidden
   AfxRedrawWindow(pDoc->hWndFrame)
   AfxRedrawWindow(pDoc->hWndForm)

   function = 0
end function


' ========================================================================================
' Handle WM_LBUTTONUP messages for the Form and Frame windows
' ========================================================================================
function HandleDesignerLButtonUp( ByVal HWnd As HWnd ) as LRESULT

   dim pDoc as clsDocument ptr = gApp.GetDocumentPtrByWindow(hwnd)
   if pDoc = 0 THEN exit function
   
   ' Hide any previous lasso (and select controls)
   if pDoc->bLasso THEN 
      dim pCtrl as clsControl ptr
      dim as Rect rcIntersect, rcLasso, rcCtrl
      DrawLasso(pDoc, pDoc->ptStart, pDoc->ptEnd)
      SetRect(@rcLasso, pDoc->ptStart.x, pDoc->ptStart.y, pDoc->ptEnd.x, pDoc->ptEnd.y)
      NormalizeRect(rcLasso)

      ' If the Toolbox Pointer/Arrow is selected then attempt to select the controls
      ' that intersect with the lasso, otherwise draw and create the new Toolbox control.
      if GetActiveToolboxControlType = CTRL_POINTER THEN
         dim as hwnd hCtrlSel
         pDoc->Controls.DeselectAllControls
         MapWindowPoints(pDoc->hWndForm, 0, cast(point ptr, @rcLasso), 2)
         for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
            pCtrl = pDoc->Controls.ItemAt(i)
            if pCtrl->ControlType <> CTRL_FORM THEN
               GetWindowRect( pCtrl->hWindow, @rcCtrl)
               If IntersectRect( @rcIntersect, @rcCtrl, @rcLasso ) Then
                  hCtrlSel = pCtrl->hWindow
                  pCtrl->IsSelected = true
               end if
            end if
         next
         if hCtrlSel = 0 THEN hCtrlSel = pDoc->hWndForm
         pDoc->Controls.SelectControl(hCtrlSel)
         pDoc->Controls.SetActiveControl(hCtrlSel)
      else
         ' Create the selected Toolbox control.
         CreateToolboxControl( pDoc, GetActiveToolboxControlType, rcLasso )
      end if   
   end if      
    
   pDoc->ptStart  = pDoc->ptEnd
   pDoc->GrabHit  = GRAB_NOHIT
   pDoc->bSizing  = false
   pDoc->bMoving  = false
   pDoc->bLasso   = false
   pDoc->pCtrlAction = 0
   
   SetActiveToolboxControl(CTRL_POINTER)

   ReleaseCapture

   ' Ensure the grab handles of form and controls are redrawn or hidden
   AfxRedrawWindow(pDoc->hWndFrame)
   AfxRedrawWindow(pDoc->hWndForm)
   
   function = 0
end function

' ========================================================================================
' Handle MOUSEMOVE messages for the Form and Frame windows
' ========================================================================================
function HandleDesignerMouseMove( ByVal HWnd As HWnd ) as LRESULT

   dim as POINT pt
   dim as long xDelta, yDelta
   
   dim pDoc as clsDocument ptr = gApp.GetDocumentPtrByWindow(hwnd)
   if pDoc = 0 THEN exit function
   
   dim pCtrl as clsControl ptr

   ' Hide any previous lasso
   if pDoc->bLasso THEN DrawLasso(pDoc, pDoc->ptStart, pDoc->ptEnd)

   GetCursorPos(@pt)
   MapWindowPoints(0, HWND, @pt, 1)
   xDelta = pt.x - pDoc->ptPrev.x
   yDelta = pt.y - pDoc->ptPrev.y
   
   pDoc->ptEnd = pt
   
   if pDoc->bSizing THEN
      Select Case pDoc->GrabHit
         Case GRAB_BOTTOMRIGHT 
            pDoc->rcSize.right  = pDoc->rcSize.right + xDelta
            pDoc->rcSize.bottom = pDoc->rcSize.bottom + yDelta
         Case GRAB_RIGHT 
            pDoc->rcSize.right  = pDoc->rcSize.right + xDelta
         Case GRAB_BOTTOM
            pDoc->rcSize.bottom = pDoc->rcSize.bottom + yDelta
      End Select

      ' NOTE:
      ' Set our control to a default minimum value If Zero.
      ' We want to do this so we don't loose visibility of our control and we can still 
      ' see the handles when selected.

      ' Check for a minimum width and height
      If pDoc->rcSize.right - pDoc->rcSize.left <= AfxScaleX(8) Then pDoc->rcSize.right = pDoc->rcSize.left + AfxScaleX(8)
      If pDoc->rcSize.bottom - pDoc->rcSize.top <= AfxScaleY(8) Then pDoc->rcSize.bottom = pDoc->rcSize.top + AfxScaleY(8)

      ' Resize all selected the form/control 
      SetWindowPos( pDoc->pCtrlAction->hWindow, 0, 0, 0, _
                    pDoc->rcSize.right - pDoc->rcSize.left, pDoc->rcSize.bottom - pDoc->rcSize.top, _
                    SWP_NOMOVE or SWP_NOZORDER)
      
      ' Ensure the grab handles are redrawn
      AfxRedrawWindow(GetParent(pDoc->pCtrlAction->hWindow))
      
   elseif pDoc->bMoving then
      ' Move the control to its new position
      for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
         pCtrl = pDoc->Controls.ItemAt(i)
         if pCtrl->IsSelected THEN
            GetWindowRect(pCtrl->hWindow, @pDoc->rcSize)
            MapWindowPoints(0, pDoc->hWndForm, cast(point ptr, @pDoc->rcSize), 2)
            pDoc->rcSize.left  = pDoc->rcSize.left + xDelta
            pDoc->rcSize.top   = pDoc->rcSize.top + yDelta
            SetWindowPos( pCtrl->hWindow, 0, _
                          pDoc->rcSize.left, pDoc->rcSize.top, _
                          0, 0, SWP_NOSIZE or SWP_NOZORDER)
         END IF
      next
      ' Ensure the grab handles are redrawn
      AfxRedrawWindow(GetParent(pDoc->pCtrlAction->hWindow))
      
   elseif pDoc->bLasso then
      DrawLasso(pDoc, pDoc->ptStart, pDoc->ptEnd)
   else   
      pDoc->GrabHit = SetGrabHandleMouseCursor(pDoc, pt.x, pt.y, pCtrl)
   end if      

   ' Save the current mouse position
   pDoc->ptPrev.x = pt.x
   pDoc->ptPrev.y = pt.y

   function = 0
   
END FUNCTION



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

   dim pDoc as clsDocument ptr = gApp.GetDocumentPtrByWindow(hwnd)
   if pDoc THEN DrawGrabHandles(hDC, pDoc, false)

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
' Lasso popup window
' ========================================================================================
Function Lasso_WndProc( ByVal HWnd   As HWnd, _
                        ByVal uMsg   As UINT, _
                        ByVal wParam As WPARAM, _
                        ByVal lParam As LPARAM _
                        ) As LRESULT

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

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
      HANDLE_MSG (HWnd, WM_PAINT,       DesignerForm_OnPaint)
      HANDLE_MSG (HWnd, WM_DESTROY,     DesignerForm_OnDestroy)

      case WM_LBUTTONDOWN: HandleDesignerLButtonDown(HWND)
      case WM_LBUTTONUP:   HandleDesignerLButtonUp(HWND)
      case WM_MOUSEMOVE:   HandleDesignerMouseMove(HWND)

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
' Processes messages for the subclassed controls.
' ========================================================================================
FUNCTION Control_SubclassProc( BYVAL hwnd   AS HWND, _                 ' Control window handle
                               BYVAL uMsg   AS UINT, _                 ' Type of message
                               BYVAL wParam AS WPARAM, _               ' First message parameter
                               BYVAL lParam AS LPARAM, _               ' Second message parameter
                               BYVAL uIdSubclass AS UINT_PTR, _        ' The subclass ID
                               BYVAL dwRefData AS DWORD_PTR _          ' Pointer to reference data
                               ) AS LRESULT

   dim pDoc as clsDocument ptr = cast(clsDocument ptr, dwRefData)
   
   SELECT CASE uMsg

      CASE WM_GETDLGCODE
         ' All keyboard input
         FUNCTION = DLGC_WANTALLKEYS
         EXIT FUNCTION

      Case WM_RBUTTONDOWN  
         ' Let the parent Form handle the right click mouse message
         if pDoc THEN PostMessage( pDoc->hWndForm, uMsg, wParam, cast(lParam, dwRefData) )
         Exit Function                                                                       

      case WM_LBUTTONDOWN: HandleDesignerLButtonDown(pDoc->hWndForm)
      case WM_LBUTTONUP:   HandleDesignerLButtonUp(pDoc->hWndForm)
      case WM_MOUSEMOVE:   HandleDesignerMouseMove(pDoc->hWndForm)

'      CASE WM_KEYDOWN
'         SELECT CASE GET_WM_COMMAND_ID(wParam, lParam)
'            CASE VK_ESCAPE
'               SendMessage(GetParent(hwnd), WM_CLOSE, 0, 0)
'               EXIT FUNCTION
'         END SELECT

      Case WM_MOUSEACTIVATE
         ' Defeat this message so that mouse clicks do not activate the control
         Function = MA_NOACTIVATE: uMsg = WM_NULL
         Exit Function

      Case WM_SETCURSOR
         Function = TRUE: uMsg = WM_NULL
         Exit Function
          
      Case WM_SETFOCUS
         ' Defeat the caret activation for the RichEdit Control, for some
         ' reason MA_NOACTIVATE does not work for right clicks.
         Function = 0: uMsg = WM_NULL
         Exit Function

      CASE WM_DESTROY
         ' REQUIRED: Remove control subclassing
         RemoveWindowSubclass( hwnd, @Control_SubclassProc, uIdSubclass )

   END SELECT

   ' Default processing of Windows messages
   FUNCTION = DefSubclassProc(hwnd, uMsg, wParam, lParam)

END FUNCTION


