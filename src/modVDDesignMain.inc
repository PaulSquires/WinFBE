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
            AfxRedrawWindow(hWndFrame)
            AfxRedrawWindow(hWndForm)
         end if   
         EXIT FUNCTION

   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


