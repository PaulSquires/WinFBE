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
' Draw the actual grab handles (this is called from WM_PAINT)
' ========================================================================================
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
      else
         if pCtrl->ControlType = CTRL_FORM THEN continue for
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
      SetRect(@pCtrl->rcHandles(GRAB_TOP), rc.left+HMid-(cx/2), rc.top-cy, rc.left+HMid+(cx/2), rc.top)
      SetRect(@pCtrl->rcHandles(GRAB_TOPRIGHT), rc.right, rc.top-cy, rc.right+cx, rc.top)
      SetRect(@pCtrl->rcHandles(GRAB_RIGHT), rc.right, rc.top+VMid-(cy/2), rc.right+cx, rc.top+VMid+(cy/2))
      SetRect(@pCtrl->rcHandles(GRAB_BOTTOMRIGHT), rc.right, rc.bottom, rc.right+cx, rc.bottom+cy)
      SetRect(@pCtrl->rcHandles(GRAB_BOTTOM), rc.left+HMid-(cx/2), rc.bottom, rc.left+HMid+(cx/2), rc.bottom+cy)
      SetRect(@pCtrl->rcHandles(GRAB_BOTTOMLEFT), rc.left-cx, rc.bottom, rc.left, rc.bottom+cy)
      SetRect(@pCtrl->rcHandles(GRAB_LEFT), rc.left-cx, rc.top+VMid-(cy/2), rc.left, rc.top+VMid+(cy/2))

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

   dim pDoc as clsDocument ptr = gApp.GetDocumentPtrByWindow(hwnd)
   if pDoc THEN
      ' Draw grab handles for the controls on the form 
      DrawGrabHandles(hDC, pDoc, false)
   END IF

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

'      CASE WM_LBUTTONDOWN
'         MessageBox(GetParent(hwnd), "Click", "FreeBasic", MB_OK)
'         EXIT FUNCTION

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


