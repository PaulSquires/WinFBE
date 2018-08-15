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
' Lasso window callback procedure
' ========================================================================================
FUNCTION WndProcLasso( BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT
   
   ' Default processing of Windows messages
   FUNCTION = DefWindowProc(hWnd, uMsg, wParam, lParam)

END FUNCTION


Destructor clsLasso
   if pWindow THEN this.Destroy
end destructor

function clsLasso.IsActive() as Boolean
   function = this.bLasso
END FUNCTION

function clsLasso.GetLassoRect() as RECT
   dim as RECT rc = (this.ptStart.x, this.ptStart.y, this.ptEnd.x, this.ptEnd.y)
   ' Normalize the rect
   If rc.Left > rc.Right   Then Swap rc.Right, rc.Left
   If rc.Top  > rc.Bottom  Then Swap rc.Top, rc.Bottom
   return rc
end function

function clsLasso.SetStartPoint( byval x as long, byval y as Long) as Long
   this.ptStart.x = x 
   this.ptStart.y = y
   function = 0
end function
   
function clsLasso.SetEndPoint( byval x as long, byval y as Long) as Long
   this.ptEnd.x = x 
   this.ptEnd.y = y
   function = 0
end function

function clsLasso.GetStartPoint() as POINT
   return this.ptStart
end function

function clsLasso.GetEndPoint() as POINT
   return this.ptEnd
end function

function clsLasso.Create( byval hWndParent as hwnd ) as boolean
   dim as Boolean bResult = false
   dim as RECT rc
   
   this.pWindow = new CWindow
   if this.pWindow THEN
      GetClientRect(hWndParent, @rc)
      MapWindowPoints(hWndParent, 0, cast( POINT Ptr, @rc), 2)
      this.hWindow = this.pWindow->Create( hWndParent, "", @WndProcLasso, _
                                           rc.left, rc.top, rc.right-rc.left, rc.bottom-rc.top, _
                                           WS_POPUP or WS_VISIBLE, WS_EX_LAYERED)
      this.pWindow->ClassStyle = CS_DBLCLKS
      this.hWndParent = hWndParent
      this.SetStartPoint(-1,-1)
      this.SetEndPoint(-1,-1)
      this.bLasso = true
      SetFocus(hWndParent)
      bResult = iif(this.pWindow->hWindow, true, false)
   end if
  
   function = bResult
END function


function clsLasso.FillAlpha(byval hBmp as HBITMAP) as Boolean

   dim as Boolean bResult = false

   if (hBmp) then 
      dim as BITMAP bmp
      GetObject(hBmp, sizeof(BITMAP), @bmp)
      dim as DWORD dwCount = bmp.bmWidthBytes * bmp.bmHeight
      if (dwCount >= sizeof(DWORD)) then
         dim as DWORD ptr pcBitsWords = cast(DWORD ptr, bmp.bmBits)
         if (pcBitsWords) then
            dim as DWORD dwIndex = (dwCount / sizeof(DWORD)) - 1
            dim as DWORD dwUp = bmp.bmWidth
            dim as DWORD dwDn = dwIndex -dwUp
            dim as DWORD dwR  = bmp.bmWidth -1
            while dwIndex 
               dim as DWORD dwSides = dwIndex mod bmp.bmWidth
               if (dwIndex < dwUp) or (dwIndex > dwDn) or (dwSides = 0) or(dwSides = dwR) then
                  pcBitsWords[dwIndex] = &HFF0080FF  'sm_clrPenA;   // 0xFF0080FF
               else 
                  pcBitsWords[dwIndex] = &H400020FF  'sm_clrBrushA; // 0x400020FF
               end if
               dwIndex = dwIndex - 1
            wend
            bResult = true
         end if
      end if
   end if
   return bResult
end function


function clsLasso.Show() as Long
   if this.bLasso THEN 

'    CRect cMoveRect(crPos);
'    cMoveRect.NormalizeRect();

'    CRect cIntersectRect(cMoveRect);
'    cIntersectRect.IntersectRect(m_cParentCliRect, cMoveRect);

      dim as RECT rcPos = this.GetLassoRect()
      
      dim as HDC hdcScreen = GetDC(0)
      dim as HDC hDC = CreateCompatibleDC(hdcScreen)
      dim as long iWidth  = rcPos.right - rcPos.left
      dim as long iHeight = rcPos.bottom - rcPos.top

      dim as BITMAPINFO sBI      
      sBI.bmiHeader.biSize        = sizeof(BITMAPINFOHEADER)
      sBI.bmiHeader.biWidth       = iWidth
      sBI.bmiHeader.biHeight      = iHeight 
      sBI.bmiHeader.biPlanes      = 1 
      sBI.bmiHeader.biBitCount    = 32 
      sBI.bmiHeader.biCompression = BI_RGB

      dim as HBITMAP hBmp = CreateDIBSection(hDC, @sBI, DIB_RGB_COLORS, NULL, NULL, 0)
      dim as HBITMAP hBmpOld = SelectObject(hDC, hBmp)

      dim as Boolean bFillAlphaOK = FillAlpha(hBmp)
      if (not bFillAlphaOK) then 
      '  dim as Boolean bFillAlphaOK = false
      '  dim as HPEN hPen1 = CreatePen(PS_INSIDEFRAME, 2, &HFF0080FF) 'sm_clrPen)
      '  dim as HBRUSH hBrush1 = CreateSolidBrush(&H400020FF) 'sm_clrBrush);
      '  dim as HPEN hOldPen = SelectObject(hDC, hPen1)
      '  dim as HBRUSH hOldBrush = SelectObject(hDC, hBrush1)
      '  Rectangle(hDC, 0, 0, iWidth, iHeight)
      '  SelectObject(hDC, hOldPen)
      '  SelectObject(hDC, hOldBrush)
      '  DeleteObject(hPen1)
      '  DeleteObject(hBrush1)
      end if

      dim as BLENDFUNCTION blend
      blend.BlendOp             = AC_SRC_OVER
      blend.SourceConstantAlpha = iif(bFillAlphaOK, 160, 64)
      blend.AlphaFormat         = iif(bFillAlphaOK, AC_SRC_ALPHA, 0)
       
      ' Destination position at the screen
      '    POINT ptPos   = {cIntersectRect.left,
      '                     cIntersectRect.top};
      dim as RECT rc
      GetClientRect(this.hWndParent, @rc)
      MapWindowPoints(this.hWndParent, 0, cast(POINT ptr, @rc), 2)
      dim as POINT ptPos  = (rc.left + rcPos.left, rc.top + rcPos.top)

      ' Source position in source (memory DC)
      '    POINT ptSrc   = {cIntersectRect.left - cMoveRect.left,
      '                     cIntersectRect.top  - cMoveRect.top};
      dim as point ptSrc = (0,0) '(rcPos.left, rcPos.top)

      ' Dimensions of the bits transfer
      '    SIZE sizeWnd  = {cIntersectRect.Width(),
      '                     cIntersectRect.Height()};
      dim as SIZE sizeWnd  = (iWidth, iHeight)

      'UpdateLayeredWindow(
      '    __in HWND hWnd,                  // handle of the layered window
      '    __in_opt HDC hdcDst,             // destination DC (in our case screen DC)
      '    __in_opt POINT *pptDst,          // destination position at the screen
      '    __in_opt SIZE *psize,            // dimensions of window at the screen
      '    __in_opt HDC hdcSrc,             // source (memory) DC of prepainting
      '    __in_opt POINT *pptSrc,          // source position for the surface bits transfer
      '    __in COLORREF crKey,             // color to be fully transparent (not our case)
      '    __in_opt BLENDFUNCTION *pblend,  // blending parameters
      '    __in DWORD dwFlags);             // kind of the transfer 
      '                                     // (in our case ULW_ALPHA - for alpha blending)
      UpdateLayeredWindow(this.hWindow, hdcScreen, @ptPos, @sizeWnd, hDC, @ptSrc, 0, @blend, ULW_ALPHA) 

      SelectObject(hDC, hBmpOld)
      DeleteObject(hBmp)
      DeleteDC(hDC)
      ReleaseDC(0, hdcScreen)

   end if   
   
   function = 0
end function


function clsLasso.Destroy() as boolean
   if this.pWindow THEN 
      this.bLasso = false
      DestroyWindow( pWindow->hWindow)
      delete this.pWindow
      this.pWindow = 0
   end if   
   function = true
END FUNCTION


dim shared gLasso as clsLasso


