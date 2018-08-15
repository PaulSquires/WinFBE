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

' ========================================================================================
' Process WM_PAINT message for Visual Designer Frame
' ========================================================================================
Function DesignerFrame_OnPaint( ByVal HWnd As HWnd) As LRESULT

   Dim As PAINTSTRUCT ps
   Dim As HDC hDC

   hDC = BeginPaint(hWnd, @ps)

   dim pDoc as clsDocument ptr = gApp.GetDocumentPtrByWindow(hwnd)
   if pDoc THEN DrawGrabHandles(hDC, pDoc, true)

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
      HANDLE_MSG (HWnd, WM_PAINT,   DesignerFrame_OnPaint)
      HANDLE_MSG (HWnd, WM_DESTROY, DesignerFrame_OnDestroy)
      
      case WM_LBUTTONDOWN: HandleDesignerLButtonDown(HWND)
      case WM_LBUTTONUP:   HandleDesignerLButtonUp(HWND)
      case WM_MOUSEMOVE:   HandleDesignerMouseMove(HWND)
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function

