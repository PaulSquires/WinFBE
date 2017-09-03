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
      GetWindowRect(pDoc->hWndForm, @pDoc->rcSize)
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

      ' Move the form/control to its new position
      SetWindowPos( pDoc->hWndForm, 0, 0, 0, _
                    pDoc->rcSize.right - pDoc->rcSize.left, pDoc->rcSize.bottom - pDoc->rcSize.top, _
                    SWP_NOMOVE or SWP_NOZORDER)
      
      ' Ensure the grab handles are redrawn
      AfxRedrawWindow(pDoc->hWndFrame)
      
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

