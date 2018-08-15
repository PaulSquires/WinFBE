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
''  FRMHOTIMGBTN
''
''  Image button with hot mouse over 
''
''

Type HOTIMGBTN_TYPE
   hWndParent As HWND
   cID        As Long
   hTooltip   as HWND
   hImageList AS HIMAGELIST
   idxNormal  As long         ' position in imagelist of normal icon
   idxHot     As long         ' position in imagelist of hot icon
   clrBg      As COLORREF
   clrBgHot   As COLORREF
   nImgWidth  As Long
   nImgHeight As Long
   fSelected  As BOOLEAN      ' flag that button should remain displaying as selected
   fSelectedFrame as boolean  ' draw frame around selected button
   fEnabled   as boolean      ' should the button allow clicks
End Type


' ========================================================================================
' Return the Enabled/Disabled flag
' ========================================================================================
Function frmHotImgBtn_IsEnabled( ByVal HWnd As HWnd ) As boolean
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)
   If pWindow = 0 Then Exit Function
   
   Dim pData As HOTIMGBTN_TYPE Ptr = Cast(HOTIMGBTN_TYPE Ptr, pWindow->UserData(0))
   Function = pData->fEnabled
End Function

' ========================================================================================
' Set the Enabled/Disabled flag
' ========================================================================================
Function frmHotImgBtn_Enabled( ByVal HWnd As HWnd, ByVal fEnabled as BOOLEAN ) As Long
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)
   If pWindow = 0 Then Exit Function
   
   Dim pData As HOTIMGBTN_TYPE Ptr = Cast(HOTIMGBTN_TYPE Ptr, pWindow->UserData(0))
   If pData Then pData->fEnabled = fEnabled
   AfxRedrawWindow( hWnd )
   
   Function = 0
End Function

' ========================================================================================
' Set the background normal and hot colors. 
' ========================================================================================
Function frmHotImgBtn_SetBackColors( ByVal HWnd As HWnd, _
                                     ByVal clrNormal As COLORREF, _
                                     byval clrHot as COLORREF _
                                     ) As Long
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)
   If pWindow = 0 Then Exit Function
   
   Dim pData As HOTIMGBTN_TYPE Ptr = Cast(HOTIMGBTN_TYPE Ptr, pWindow->UserData(0))
   If pData Then 
      pData->clrBg = clrNormal
      pData->clrBgHot = clrHot
   end if   
   AfxRedrawWindow( hWnd )
   
   Function = 0
End Function


' ========================================================================================
' Set the normal and hot images. 
' ========================================================================================
Function frmHotImgBtn_SetImages( ByVal HWnd As HWnd, _
                                 ByRef wszImgNormal As WString = "", _
                                 ByRef wszImgHot    As WString = "",  _
                                 byval bIsGDI       as Boolean = false _
                                 ) As Long
 
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)
   If pWindow = 0 Then Exit Function
   
   Dim pData As HOTIMGBTN_TYPE Ptr = Cast(HOTIMGBTN_TYPE Ptr, pWindow->UserData(0))
   If pData = 0 Then exit function 

   dim as HICON hIcon
   
   ' Destroy any existing ImageList
   ImageList_Destroy(pData->hImageList)
   
   ' Calculate the size of the icon according the DPI
   DIM cx AS LONG = pData->nImgWidth * pWindow->DPI \ 96

   ' Create an image list to hold the normal and hot images
   pData->hImageList = ImageList_Create(cx, cx, ILC_COLOR32 OR ILC_MASK, 2, 0)
   IF pData->hImageList THEN
      ' Load the image depending on the file type. Use LoadImage for ICO and GDI for PNG/BMP, etc
      if bIsGDI then
         pData->idxNormal = AfxGdipAddIconFromRes(pData->hImageList, pWindow->InstanceHandle, wszImgNormal)
         pData->idxHot    = AfxGdipAddIconFromRes(pData->hImageList, pWindow->InstanceHandle, wszImgHot)
      else   
         hIcon = LoadImage( pWindow->InstanceHandle, wszImgNormal, IMAGE_ICON, pData->nImgWidth, pData->nImgHeight, LR_DEFAULTCOLOR)
         IF hIcon THEN 
            pData->idxNormal = ImageList_ReplaceIcon(pData->hImageList, -1, hIcon)
            DestroyIcon(hIcon)
         end if
         hIcon = LoadImage( pWindow->InstanceHandle, wszImgHot, IMAGE_ICON, pData->nImgWidth, pData->nImgHeight, LR_DEFAULTCOLOR)
         IF hIcon THEN 
            pData->idxHot = ImageList_ReplaceIcon(pData->hImageList, -1, hIcon)
            DestroyIcon(hIcon)
         end if
      end if
   END IF

   AfxRedrawWindow( HWnd )
   
   Function = 0
End Function

' ========================================================================================
' Set the Selected flag for the button. If Selected then the button will always 
' display using the Hot image regardless of whether the button is hot or not.
' ========================================================================================
Function frmHotImgBtn_SetSelected( ByVal HWnd As HWnd, ByVal fSelected As BOOLEAN) As Long
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)
   If pWindow = 0 Then Exit Function
   
   Dim pData As HOTIMGBTN_TYPE Ptr = Cast(HOTIMGBTN_TYPE Ptr, pWindow->UserData(0))
   If pData Then pData->fSelected = fSelected
   AfxRedrawWindow( hWnd )
   
   Function = 0
End Function

' ========================================================================================
' Gets the Selected flag from the button. If Selected then the button will always 
' display using the Hot image regardless of whether the button is hot or not.
' ========================================================================================
Function frmHotImgBtn_GetSelected( ByVal HWnd As HWnd ) As BOOLEAN
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)
   If pWindow = 0 Then Exit Function
   
   Dim pData As HOTIMGBTN_TYPE Ptr = Cast(HOTIMGBTN_TYPE Ptr, pWindow->UserData(0))
   If pData Then Return pData->fSelected

   Function = False
End Function


' ========================================================================================
' Process WM_LBUTTONUP message for frmHotImgBtn
' ========================================================================================
Function frmHotImgBtn_OnLButtonUp(ByVal HWnd As HWnd, ByVal x As Long, ByVal y As Long, ByVal keyFlags As Long ) As LRESULT
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)
   If pWindow = 0 Then Exit Function
   
   Dim pData As HOTIMGBTN_TYPE Ptr = Cast(HOTIMGBTN_TYPE Ptr, pWindow->UserData(0))

   ' Post a clicked message to the parent's WM_COMMAND handler.
   If pData Then
      PostMessageW( pData->hWndParent, WM_COMMAND, MAKEWPARAM(pData->cID, BN_CLICKED), Cast(LPARAM, hWnd) )
   End If
   
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmHotImgBtn
' ========================================================================================
Function frmHotImgBtn_OnDestroy(HWnd As HWnd) As LRESULT
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)
   If pWindow = 0 Then Exit Function
   
   Dim pData As HOTIMGBTN_TYPE Ptr = Cast(HOTIMGBTN_TYPE Ptr, pWindow->UserData(0))
   If pData Then
      ImageList_Destroy(pData->hImageList)
      AfxDeleteTooltip(pData->hTooltip, HWND)
      DestroyWindow pData->hTooltip
      Delete pData
   End If
      
   if pWindow then Delete pWindow

   Function = 0
End Function


' ========================================================================================
' Process WM_PAINT message for window/dialog: frmHotImgBtn
' ========================================================================================
Function frmHotImgBtn_OnPaint( ByVal HWnd As HWnd) As LRESULT

   Dim pWindow As CWindow Ptr = AfxCWindowPtr(hWnd)
   If pWindow = 0 Then Exit Function
   
   Dim As PAINTSTRUCT ps
   Dim As HBITMAP hbit 
   
   Dim As HDC hDC, memDC 
   Dim As HBRUSH hBrush
   Dim As POINT pt
   Dim As RECT rc
   Dim As Long nHMargin, nVMargin
   Dim As long idxImage
   dim as long cx, cy
   Dim As BOOLEAN bIsHot

   hDC = BeginPaint(hWnd, @ps)

   SaveDC hDC

   Dim pData As HOTIMGBTN_TYPE Ptr = Cast(HOTIMGBTN_TYPE Ptr, pWindow->UserData(0))
   
   If pData Then
      GetClientRect hWnd, @rc
      memDC = CreateCompatibleDC( ps.hDC )
      hbit  = CreateCompatibleBitmap( ps.hDC, rc.Right, rc.Bottom )
      If hbit Then hbit = SelectObject( memDC, hbit )

      bIsHot = IIf( GetProp(hWnd, "HOT") = 0, False, True ) 

      dim as HPEN hPen = CreatePen(PS_SOLID, 1, BGR(0,0,255))  ' blue
      SelectObject( memDC, hPen )

      If ( bIsHot = True ) Or (pData->fSelected = True) Then
         idxImage = pData->idxHot
         hBrush = CreateSolidBrush( pData->clrBgHot )
      Else
         idxImage = pData->idxNormal
         hBrush = CreateSolidBrush( pData->clrBg )
      End If
      ' Paint the entire client background
      FillRect(memDC, @rc, hBrush) 
      DeleteBrush(hBrush)

      ' Draw the icon (center it in the rectangle)
      ImageList_GetIconSize(pData->hImageList, @cx, @cy)
      nHMargin = (rc.right - cx) / 2
      nVMargin = (rc.bottom - cy) / 2
      ImageList_Draw(pData->hImageList, idxImage, memDC, _
                     rc.Left + nHMargin, rc.Top + nVMargin, ILD_TRANSPARENT)
       
      if pData->fSelected THEN
         ' If the button is in a selected state then draw a thin
         ' border around the icon.
         if pData->fSelectedFrame THEN
            SelectObject( memDC, GetStockObject( NULL_BRUSH ) )
            Rectangle(memDC, rc.left + 1, rc.top + 1, rc.right - 1, rc.bottom - 1)
         end if
      END IF
   
      ' Copy the entire memory bitmap over to the visual display
      BitBlt ps.hDC, 0, 0, rc.Right, rc.Bottom, memDC, 0, 0, SRCCOPY 

      ' Cleanup
      If hbit  Then DeleteObject SelectObject(memDC, hbit)
      If memDC Then DeleteDC memDC

      DeleteObject(hPen)
   End If

   RestoreDC hDC, -1 
   
   EndPaint hWnd, @ps
   
   Function = 0
End Function


' ========================================================================================
' frmHotImgBtn Window procedure
' ========================================================================================
Function frmHotImgBtn_WndProc( ByVal HWnd   As HWnd, _
                               ByVal uMsg   As UINT, _
                               ByVal wParam As WPARAM, _
                               ByVal lParam As LPARAM _
                               ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_LBUTTONUP,   frmHotImgBtn_OnLButtonUp)
      HANDLE_MSG (HWnd, WM_PAINT,       frmHotImgBtn_OnPaint)
      HANDLE_MSG (HWnd, WM_DESTROY,     frmHotImgBtn_OnDestroy)

      CASE WM_MOUSEMOVE
         ' Need to record trackmouse event in order to get the WM_MOUSELEAVE message 
         ' when the mouse leaves the client area of the tabctl.
         DIM trackMouse AS TRACKMOUSEEVENT
         IF GetPropW(hwnd, "HOT") = 0 THEN
            trackMouse.cbSize = SIZEOF(trackMouse)
            trackMouse.dwFlags = TME_LEAVE
            trackMouse.hwndTrack = hwnd
            trackMouse.dwHoverTime = 1
            TrackMouseEvent(@trackMouse)
            SetPropW hwnd, "HOT", CAST(HANDLE, CTRUE)
            ' Ensure repainting as hot
            AfxRedrawWindow hWnd
         END IF
         Exit Function
            
      CASE WM_MOUSELEAVE
         ' Repaint to remove any Hot state tab
         RemovePropW hwnd, "HOT"
         AfxRedrawWindow hWnd
         Exit Function

      Case WM_ERASEBKGND
         Return True  ' prevent background erase using class brush

   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function



' ========================================================================================
' frmHotImgBtn
' ========================================================================================
Function frmHotImgBtn( ByVal hWndParent   As HWnd, _
                       ByVal cID          As Long = 0, _
                       ByRef wszImgNormal As WString = "", _
                       ByRef wszImgHot    As WString = "",  _
                       ByRef wszTooltip   As WString = "",  _
                       ByVal nImgWidth    As Long = 24, _
                       ByVal nImgHeight   As Long = 24, _
                       ByVal clrBg        As COLORREF = 0, _
                       ByVal clrBgHot     As COLORREF = 0, _
                       BYVAL x            AS LONG = 0, _
                       BYVAL y            AS LONG = 0, _
                       BYVAL nWidth       AS LONG = 0, _
                       BYVAL nHeight      AS LONG = 0, _
                       byval fSelectedFrame as BOOLEAN = true, _
                       byval fEnabled       as BOOLEAN = true, _
                       byval fIsGDI         as Boolean = false _
                       ) As HWND

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI
   Dim As HWND hWindow = _
   pWindow->Create( hWndParent, "", @frmHotImgBtn_WndProc, x, y, nWidth, nHeight, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->ClassStyle = 0  ' to ensure that there are no double clicks
   pWindow->Brush = Null
   
   SetWindowLongPtr( hWindow, GWLP_ID, cID )

   Dim pData As HOTIMGBTN_TYPE Ptr = New HOTIMGBTN_TYPE
   ' Save the control data elements
   pWindow->UserData(0) = Cast(LONG_PTR, pData)
   With *pData
      .hWndParent = hWndParent
      .cID        = cID
      .nImgWidth  = nImgWidth 
      .nImgHeight = nImgHeight
      .clrBg      = clrBg
      .clrBgHot   = clrBgHot
      .fSelectedFrame = fSelectedFrame
      .fEnabled   = fEnabled
   End With
   
   frmHotImgBtn_SetImages(hWindow, wszImgNormal, wszImgHot, fIsGDI)
   
   ' Register a tooltip for this window. Change the text of the tooltip to the full filename/path as
   ' the mouse changes from tab to tab.
   Dim As Hwnd hTooltip
   If Len(wszTooltip) Then
      hTooltip = AfxAddTooltip( hWindow, AfxStrRemove(wszTooltip, "&") )
      Tooltip_SetDelayTime( hTooltip, TTDT_INITIAL, 250 )
      Tooltip_SetDelayTime( hTooltip, TTDT_RESHOW, 1000 )
      Tooltip_SetDelayTime( hTooltip, TTDT_AUTOPOP, 2000 )
      pData->hTooltip = hTooltip
   End If
   
   ' Refresh the control
   AfxRedrawWindow(hWindow)
   
   Function = hWindow
   
End Function





