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


''
''  Owner draw combo box for color selection
''  Based on code by José Roca 2011 and Börje Hagsten, January 2001

const MODCBCOLOR_USERSELECTED = 41

' ========================================================================================
' GetColorInfo
' ========================================================================================
Sub GetColorInfo( ByVal nIndex      As Long, _
                  ByVal wColorName  As WString Ptr, _
                  ByRef nColorValue As COLORREF _
                  ) 
                       
   Select Case nIndex
      Case 0:  *wColorName = "Black":          nColorValue = BGR(000,000,000)
      Case 1:  *wColorName = "Light Black":    nColorValue = BGR(030,030,030)
      Case 2:  *wColorName = "Charcoal Black": nColorValue = BGR(050,050,050)
      Case 3:  *wColorName = "Evening Black":  nColorValue = BGR(090,090,090)
      Case 4:  *wColorName = "Gray":           nColorValue = BGR(128,128,128)
      Case 5:  *wColorName = "Light Gray":     nColorValue = BGR(204,204,204)
      Case 6:  *wColorName = "Pale Gray":      nColorValue = BGR(237,236,235)
      Case 7:  *wColorName = "White":          nColorValue = BGR(255,255,255)
      Case 8:  *wColorName = "Dark Green":     nColorValue = BGR(051,102,000)
      Case 9:  *wColorName = "Green":          nColorValue = BGR(000,128,000)
      Case 10: *wColorName = "Moss Green":     nColorValue = BGR(095,125,049)
      Case 11: *wColorName = "Teal":           nColorValue = BGR(000,128,128)
      Case 12: *wColorName = "Signal Green":   nColorValue = BGR(000,255,000)  
      Case 13: *wColorName = "Military Green": nColorValue = BGR(237,255,102)
      Case 14: *wColorName = "Lime":           nColorValue = BGR(237,255,127)
      Case 15: *wColorName = "Mint":           nColorValue = BGR(238,255,185)
      Case 16: *wColorName = "Pale Green":     nColorValue = BGR(244,255,223)
      Case 17: *wColorName = "Brown":          nColorValue = BGR(128,064,000)
      Case 18: *wColorName = "Dark Red":       nColorValue = BGR(153,000,000)
      Case 19: *wColorName = "Red":            nColorValue = BGR(196,000,000)
      Case 20: *wColorName = "Signal Red":     nColorValue = BGR(255,000,000)
      Case 21: *wColorName = "Orange":         nColorValue = BGR(255,102,000)
      Case 22: *wColorName = "Orange Brown":   nColorValue = BGR(195,137,080)
      Case 23: *wColorName = "Light Orange":   nColorValue = BGR(255,204,000)
      Case 24: *wColorName = "Yellow":         nColorValue = BGR(255,255,000)
      Case 25: *wColorName = "Light Yellow":   nColorValue = BGR(255,255,136)
      Case 26: *wColorName = "Pale Yellow":    nColorValue = BGR(255,255,223)
      Case 27: *wColorName = "Deep Purple":    nColorValue = BGR(128,000,128)
      Case 28: *wColorName = "Dark Purple":    nColorValue = BGR(173,000,173)
      Case 29: *wColorName = "Magenta":        nColorValue = BGR(128,000,128)
      Case 30: *wColorName = "Purple":         nColorValue = BGR(210,000,210)
      Case 31: *wColorName = "Pink":           nColorValue = BGR(255,000,178)
      Case 32: *wColorName = "Light Purple":   nColorValue = BGR(165,134,181)
      Case 33: *wColorName = "Pale Purple":    nColorValue = BGR(250,235,255)
      Case 34: *wColorName = "Dark Blue":      nColorValue = BGR(000,000,128)
      Case 35: *wColorName = "Blue":           nColorValue = BGR(000,000,255)
      Case 36: *wColorName = "Lavender":       nColorValue = BGR(101,158,254)
      Case 37: *wColorName = "Light Blue":     nColorValue = BGR(038,079,120)
      Case 38: *wColorName = "Powder Blue":    nColorValue = BGR(148,202,240)
      Case 39: *wColorName = "Bright Cyan":    nColorValue = BGR(000,255,255)
      Case 40: *wColorName = "Pale Blue":      nColorValue = BGR(230,245,255)
      Case 41: *wColorName = L(341, "User selected..."): nColorValue = BGR(255,255,255)
    
   End Select
   
End Sub


' ========================================================================================
' Subclassed Combobox procedure
' ========================================================================================
Function CBProc( ByVal HWnd   As HWnd, _
                 ByVal wMsg   As UInt, _
                 ByVal wParam As WPARAM, _
                 ByVal lParam As LPARAM _
                 ) As LRESULT

   Dim As LONG_PTR OldWndProc = GetWindowLongPtr( HWnd, GWLP_USERDATA )

   Select Case wMsg
   
      case MSG_USER_SETCOLORCUSTOM
         SendMessage(hwnd, CB_SETITEMDATA, MODCBCOLOR_USERSELECTED, Cast(LPARAM, wParam))
         
         
      Case WM_DESTROY
         ' Unsubclass the combo box
         If OldWndProc Then SetWindowLongPtr( HWnd, GWLP_WNDPROC, Cast(LONG_PTR, OldWndProc) )
         Exit Function

      Case WM_DRAWITEM
         Dim hBrush      As HBRUSH
         Dim lpdis       As DRAWITEMSTRUCT Ptr
         Dim rc          As Rect
         Dim wColorName  As WString * 80
         Dim nColorValue As COLORREF 
                   
         lpdis = Cast(DRAWITEMSTRUCT Ptr, lParam)
         If lpdis->itemID = &HFFFFFFFF Then Exit Function

         GetColorInfo lpdis->itemID, @wColorName, nColorValue
         ' Actually, get the colorvalue from the item of the row
         ' being drawn because the "user selected" value can change.
         nColorValue = lpdis->itemData

         Select Case lpdis->itemAction
            Case ODA_DRAWENTIRE, ODA_SELECT
               Dim pWindow As CWindow Ptr = AfxCWindowOwnerPtr(HWnd)
               ' Clear background
               FillRect( lpdis->hDC, @lpdis->rcItem, GetSysColorBrush(COLOR_WINDOW))
               ' Set the font
               SelectFont( lpdis->hDC, pWindow->Font)
               ' Set text background
               SetBkColor( lpdis->hDC, GetSysColor(COLOR_WINDOW))
               ' Set text color
               SetTextColor( lpdis->hDC, GetSysColor(COLOR_WINDOWTEXT))
               rc = lpdis->rcItem
               rc.Left = pWindow->ScaleX(28)
               DrawText( lpdis->hDC, @wColorName, Len(wColorName), @rc, DT_SINGLELINE Or DT_LEFT Or DT_VCENTER)
               ' Selected item
               If (lpdis->itemState And ODS_SELECTED) Then    ' // If selected
                  If (lpdis->itemState And &H1000) = 0 Then   ' // If not ODS_COMBOBOXEDIT (= &H1000)
                     rc.Left   = pWindow->ScaleX(26)
                     rc.Right  = lpdis->rcItem.Right
                     rc.Top    = lpdis->rcItem.Top
                     rc.Bottom = lpdis->rcItem.Bottom
                     ' Invert area around text only
                     InvertRect( lpdis->hDC, @rc)
                  End If
                  'and draw a focus rectangle around all
                  DrawFocusRect( lpdis->hDC, @lpdis->rcItem)
               End If

               ' color rectangle (using RoundRect for nicer looks..
               If (lpdis->itemState And &H1000) Then          ' // If ODS_COMBOBOXEDIT (= &H1000)
                  ' Set coordinates
                  rc.Left  = pWindow->ScaleX(4)
                  rc.Right = pWindow->ScaleX(24)
               Else
                  ' A tiny bit to the left in list
                  rc.Left  = pWindow->ScaleX(3)
                  rc.Right = pWindow->ScaleX(23)
               End If
               rc.Top    = lpdis->rcItem.Top + pWindow->ScaleY(2)
               rc.Bottom = lpdis->rcItem.Bottom - pWindow->ScaleY(2)
               ' Create brush with proper color
               hBrush = CreateSolidBrush(nColorValue)
               ' Select brush into device context
               hBrush = SelectObject( Cast(HDC, lpdis->hDC), hBrush)
               ' Draw   
               RoundRect( lpdis->hDC, rc.Left, rc.Top, rc.Right, rc.Bottom, pWindow->ScaleX(3), pWindow->ScaleY(3))
               ' Select old brush and delete the new one
               DeleteObject SelectObject(lpdis->hDC, hBrush)
               
         End Select
         Exit Function

   End Select

   ' Pass on for processing to OldWndProc
   If OldWndProc Then
      Function = CallWindowProc( Cast(WNDPROC, OldWndProc), HWnd, wMsg, wParam, lParam) 
   End If
   
End Function


' ========================================================================================
' CreateCBColorList
' ========================================================================================
Function CreateCBColorList( ByVal HWnd As HWnd, ByVal CtrlId As Long, _
                            ByVal nLeft As Long, ByVal nTop As Long, _
                            ByVal nWidth As Long, ByVal nHeight As Long _
                            ) As HWnd

   Dim i           As Long
   Dim hCombo      As HWnd
   Dim idx         As Long
   Dim wColorName  As WString * 80
   Dim nColorValue As COLORREF 
   Dim OldWndProc  As LONG_PTR
   Dim hInst       As HINSTANCE = Cast(HINSTANCE, GetWindowLongPtr(HWnd, GWLP_HINSTANCE))
   
   hCombo = CreateWindowEx( WS_EX_CLIENTEDGE, "COMBOBOX", ByVal 0, WS_VISIBLE Or _
                            WS_CHILD Or WS_VISIBLE Or CBS_OWNERDRAWFIXED Or CBS_HASSTRINGS Or _
                            CBS_DROPDOWNLIST Or WS_TABSTOP Or CBS_DISABLENOSCROLL Or WS_VSCROLL, _
                            nLeft, nTop, nWidth, nHeight, _
                            HWnd, Cast(HMENU, Cast(LONG_PTR, CtrlId)), hInst, Cast(LPVOID, 0))
   If hCombo = 0 Then Exit Function

   ' Subclass the combo and initialize some control specific data
   OldWndProc = SetWindowLongPtr( hCombo, GWLP_WNDPROC, Cast(LONG_PTR, @CBProc) )
   SetWindowLongPtr( hCombo, GWLP_USERDATA, OldWndProc )

   ' Add items to the combo box
   For i = 0 To MODCBCOLOR_USERSELECTED
      GetColorInfo i, @wColorName, nColorValue
      idx = SendMessage(hCombo, CB_ADDSTRING, 0, Cast(LPARAM, @wColorName))
      SendMessage(hCombo, CB_SETITEMDATA, idx, Cast(LPARAM, nColorValue))
   Next

   Function = hCombo

End Function


