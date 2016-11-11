' ########################################################################################
' Microsoft Windows
' File: CGraphCtx.inc
' Contents: Graphic control
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "Afx/CWindow.inc"
#include once "win/ddraw.bi"   ' // needed to work together with CGdiPlus.inc
#include once "Afx/AfxGdiplus.inc"

NAMESPACE Afx

' // Notification messages
'NM_CLICK     = NM_FIRST - 2
'NM_DBLCLK    = NM_FIRST - 3
'NM_RCLICK    = NM_FIRST - 5
'NM_RDBLCLK   = NM_FIRST - 6
'NM_SETFOCUS  = NM_FIRST - 7
'NM_KILLFOCUS = NM_FIRST - 8

' // Process them in the main window callback as follows:
'
' CASE WM_NOTIFY
'    DIM phdr AS NMHDR PTR
'    phdr = CAST(NMHDR PTR, lParam)
'    IF wParam = IDC_GRCTX THEN
'       SELECT CASE phdr->code
'          CASE NM_CLICK
'             ' Left button clicked
'          CASE NM_RCLICK
'             ' Right button clicked
'          CASE NM_SETFOCUS
'             ' The control has gained focus
'          CASE NM_KILLFOCUS
'             ' The control has lost focus
'       END SELECT
'    END IF

' ========================================================================================
' CGraphCtx class
' ========================================================================================
TYPE CGraphCtx

   Private:
      m_hCtl        AS HWND          ' // Handle of the control
      m_hFont       AS HFONT         ' // Handle to font used by the control
      m_hMemDc      AS HDC           ' // Memory compatible device context handle
      m_hBmp        AS HBITMAP       ' // Bitmap handle
      m_hOldBmp     AS HBITMAP       ' // Old Bitmap handle
      m_vWidth      AS DWORD         ' // Width of the virtual buffer
      m_vHeight     AS DWORD         ' // Height of the virtual buffer
      m_bkcolor     AS COLORREF      ' // Background color
      m_MaxX        AS LONG          ' // Maximum horizontal range
      m_MaxY        AS LONG          ' // Maximum vertical range
      m_OrgX        AS LONG          ' // x-origin for current display
      m_OrgY        AS LONG          ' // y-origin for current display
      m_coord       AS RECT          ' // Client area
      m_ppvBits     AS ANY PTR PTR   ' // Location of the DIB bit values
      m_BitsPerPel  AS LONG          ' // Bits per pixel
      m_Stretchable AS BOOLEAN       ' // Contents are stretchable
      m_StretchMode AS LONG          ' // Stretch mode
      m_Resizable   AS BOOLEAN       ' // Contents are resizable

   Public:

      DECLARE CONSTRUCTOR (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF wszTitle AS WSTRING = "", _
         BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
         BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL lpParam AS LONG_PTR = 0)
      DECLARE DESTRUCTOR
      DECLARE FUNCTION hWindow () AS HWND
      DECLARE FUNCTION SetVirtualBufferSize (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG) AS LONG
      DECLARE FUNCTION GetVirtualBufferWidth () AS LONG
      DECLARE FUNCTION GetVirtualBufferHeight () AS LONG
      DECLARE FUNCTION GetMemDC () AS HDC
      DECLARE FUNCTION GethBmp () AS HBITMAP
      DECLARE FUNCTION GetBits () AS ANY PTR
      DECLARE FUNCTION Clear (BYVAL RGBColor AS COLORREF) AS BOOLEAN
      DECLARE PROPERTY Stretchable () AS BOOLEAN
      DECLARE PROPERTY Stretchable (BYVAL bStretchable AS BOOLEAN)
      DECLARE PROPERTY StretchMode () AS LONG
      DECLARE PROPERTY StretchMode (BYVAL nStretchMode AS LONG)
      DECLARE PROPERTY Resizable () AS BOOLEAN
      DECLARE PROPERTY Resizable (BYVAL bResizable AS BOOLEAN)
      DECLARE SUB LoadImageFromFile (BYREF wszFileName AS WSTRING)
      DECLARE SUB CreateBitmapFromFile (BYREF wszFileName AS WSTRING, BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE)
      DECLARE FUNCTION SaveImage (BYREF wszFileName AS WSTRING, BYREF wszMimeType AS WSTRING = "image/bmp") AS LONG
      DECLARE FUNCTION SaveImageToBmp (BYREF wszFileName AS WSTRING) AS LONG
      DECLARE FUNCTION SaveImageToJpeg (BYREF wszFileName AS WSTRING) AS LONG
      DECLARE FUNCTION SaveImageToPng (BYREF wszFileName AS WSTRING) AS LONG
      DECLARE FUNCTION SaveImageToGif (BYREF wszFileName AS WSTRING) AS LONG
      DECLARE FUNCTION SaveImageToTiff (BYREF wszFileName AS WSTRING) AS LONG
      DECLARE FUNCTION PrintImage (BYVAL bStretch AS BOOLEAN = FALSE, BYVAL nStretchMode AS LONG = InterpolationModeHighQualityBicubic) AS LONG
      DECLARE STATIC FUNCTION CGraphCtxProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' CGraphCtx class constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CGraphCtx (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF wszTitle AS WSTRING = "", _
   BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
   BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL lpParam AS LONG_PTR = 0)

   ' // Register the class
   DIM wAtom AS ATOM
   DIM wcexw AS WNDCLASSEXW
   DIM wszClassName AS WSTRING * 260 = "AFX_GRAPHCTX"
   IF .GetClassInfoExW(.GetModuleHandleW(NULL), @wszClassName, @wcexw) = 0 THEN
      ' // Fill the WNDCLASSEXW structure
      WITH wcexw
         .cbSize        = SIZEOF(wcexw)
         .style         = CS_DBLCLKS OR CS_HREDRAW OR CS_VREDRAW
         .lpfnWndProc   = @CGraphCtxProc
         .cbClsExtra    = 0
         .cbWndExtra    = SIZEOF(HANDLE)   ' // make room to store a pointer to the class
         .hInstance     = ..GetModuleHandleW(NULL)
         .hCursor       = ..LoadCursorW(NULL, CAST(LPCWSTR, IDC_ARROW))
         .hbrBackground = CAST(HBRUSH, WHITE_BRUSH)
         .lpszMenuName  = NULL
         .lpszClassName = @wszClassName
         .hIcon         = NULL
         .hIconSm       = NULL
      END WITH
      wAtom = .RegisterClassExW(@wcexw)
   END IF

   ' // Create the control
   IF dwStyle = 0 THEN dwStyle = WS_VISIBLE OR WS_CHILD OR WS_TABSTOP
   IF pWindow THEN m_hCtl = pWindow->AddControl(wszClassName, pWindow->hWindow, cID, wszTitle, x, y, nWidth, nHeight, dwStyle, dwExStyle, lpParam)
   IF m_hCtl THEN
      .SetWindowLongPtrW m_hCtl, 0, CAST(LONG_PTR, @this)
      ' // Set the same font used by the parent
      DIM lfw AS LOGFONTW
      IF pWindow->Font THEN
         IF .GetObjectW(pWindow->Font, SIZEOF(lfw), @lfw) THEN m_hFont = CreateFontIndirectW(@lfw)
      END IF
      ' // Creates a compatible memory device context and bitmap
      .GetClientRect m_hCtl, @m_coord
      DIM hDC AS HDC = .GetDC(m_hCtl)
      m_BitsPerPel = .GetDeviceCaps(hDC, BITSPIXEL_)
      m_hMemDc = .CreateCompatibleDc(hDC)
      DIM bi AS BITMAPINFO
      bi.bmiHeader.biSize = SIZEOF(bi.bmiHeader)
      bi.bmiHeader.biWidth = m_coord.Right - m_coord.Left
      bi.bmiHeader.biHeight = m_coord.Bottom - m_coord.Top
      bi.bmiHeader.biPlanes = 1
      bi.bmiHeader.biBitCount = m_BitsPerPel
      bi.bmiHeader.biCompression = BI_RGB
      m_hBmp = .CreateDIBSection(hDC, @bi, DIB_RGB_COLORS, m_ppvBits, NULL, 0)
      m_hOldBmp = .SelectObject(m_hMemDc, m_hBmp)
      FillRect m_hMemDc, @m_coord, CAST(HBRUSH, .GetClassLongPtr(.GetParent(m_hCtl), GCLP_HBRBACKGROUND))
      ReleaseDc m_hCtl, hDc
      ' // Set the virtual size equal to the size of the bitmap
      m_vWidth = m_coord.Right - m_coord.Left
      m_vHeight = m_coord.Bottom - m_coord.Top
      ' // Default stretch mode
      m_StretchMode = HALFTONE
      ' // Initializes scroll bar ranges
      m_MaxX = m_coord.Right - m_coord.Left
      m_MaxY = m_coord.Bottom - m_coord.Top
      DIM si AS SCROLLINFO
      si.cbSize = SIZEOF(si)
      si.fMask = SIF_RANGE
      si.nMin = 0
      si.nMax = m_MaxX - m_coord.Right - m_coord.Left
      SetScrollInfo m_hCtl, SB_HORZ, @si, 1
      si.nMax = m_MaxY - m_coord.Bottom - m_coord.Top
      .SetScrollInfo m_hCtl, SB_VERT, @si, 1
      ' // Invalidates the window to force redrawing
      .InvalidateRect m_hCtl, NULL, 0
   END IF

END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CPgBar3D class destructor
' ========================================================================================
PRIVATE DESTRUCTOR CGraphCtx
   ' // Restore the original bitmap
   .SelectObject(m_hMemDc, m_hOldBmp)
   ' // Destroys the bitmap and the memory device context
   IF m_hBmp THEN .DeleteObject m_hBmp
   IF m_hMemDc THEN .DeleteObject m_hMemDc
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Window procedure
' ========================================================================================
PRIVATE FUNCTION CGraphCtx.CGraphCtxProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   DIM pGraphCtx AS CGraphCtx PTR, rc AS RECT, rcFill AS RECT, hDC AS HDC, hBrush AS HBRUSH
   DIM si AS SCROLLINFO, ps AS PAINTSTRUCT, hdr AS NMHDR, nPage AS LONG

   SELECT CASE uMsg

      CASE WM_CREATE
         EXIT FUNCTION

      CASE WM_COMMAND, WM_NOTIFY
         ' // Forwards the message to the parent window
         .SendMessageW .GetParent(hwnd), uMsg, wParam, lParam

      CASE WM_GETDLGCODE
        ' // Ensures that the control will process all the keystrokes by itself
        FUNCTION = DLGC_WANTALLKEYS
        EXIT FUNCTION

      CASE WM_LBUTTONDOWN
         ' // Forwards the message to the parent window
         hdr.hwndFrom = hwnd
         hdr.idFrom = .GetDlgCtrlId(hwnd)
         hdr.code = NM_CLICK
         .SendMessageW .GetParent(hwnd), WM_NOTIFY, .GetDlgCtrlId(hwnd), CAST(.LPARAM, @hdr)
         EXIT FUNCTION

      CASE WM_RBUTTONDOWN
         ' // Forwards the message to the parent window
         hdr.hwndFrom = hwnd
         hdr.idFrom = .GetDlgCtrlId(hwnd)
         hdr.code = NM_RCLICK
         .SendMessageW .GetParent(hwnd), WM_NOTIFY, .GetDlgCtrlId(hwnd), CAST(.LPARAM, @hdr)
         EXIT FUNCTION

      CASE WM_LBUTTONDBLCLK
         ' // Forwards the message to the parent window
         hdr.hwndFrom = hwnd
         hdr.idFrom = .GetDlgCtrlId(hwnd)
         hdr.code = NM_DBLCLK
         .SendMessageW .GetParent(hwnd), WM_NOTIFY, .GetDlgCtrlId(hwnd), CAST(.LPARAM, @hdr)
         EXIT FUNCTION

      CASE WM_RBUTTONDBLCLK
         ' // Forwards the message to the parent window
         hdr.hwndFrom = hwnd
         hdr.idFrom = .GetDlgCtrlId(hwnd)
         hdr.code = NM_RDBLCLK
         .SendMessageW .GetParent(hwnd), WM_NOTIFY, .GetDlgCtrlId(hwnd), CAST(.LPARAM, @hdr)
         EXIT FUNCTION

      CASE WM_SETFOCUS
         ' // Forwards the message to the parent window
         hdr.hwndFrom = hwnd
         hdr.idFrom = .GetDlgCtrlId(hwnd)
         hdr.code = NM_SETFOCUS
         .SendMessageW .GetParent(hwnd), WM_NOTIFY, .GetDlgCtrlId(hwnd), CAST(.LPARAM, @hdr)
         EXIT FUNCTION

      CASE WM_KILLFOCUS
         ' // Forwards the message to the parent window
         hdr.hwndFrom = hwnd
         hdr.idFrom = .GetDlgCtrlId(hwnd)
         hdr.code = NM_KILLFOCUS
         .SendMessageW .GetParent(hwnd), WM_NOTIFY, .GetDlgCtrlId(hwnd), CAST(.LPARAM, @hdr)
         EXIT FUNCTION

      CASE WM_KEYDOWN
         ' // Keyboard navigation
         SELECT CASE LOWORD(wParam)
            CASE VK_TAB
               IF HIWORD(.GetKeyState(VK_SHIFT)) = 0 THEN
                  .SetFocus .GetNextDlgTabItem(GetParent(hwnd), hwnd, 0)
               ELSE
                  .SetFocus .GetNextDlgTabItem(GetParent(hwnd), hwnd, -1)
               END IF
            CASE VK_RETURN
               .SendMessageW .GetParent(hwnd), WM_COMMAND, MAKELONG(IDOK, BN_CLICKED), CAST(.LPARAM, hwnd)
            CASE VK_ESCAPE
               .SendMessageW .GetParent(hwnd), WM_COMMAND, MAKELONG(IDCANCEL, BN_CLICKED), CAST(.LPARAM, hwnd)
            CASE VK_RIGHT
               IF .GetAsyncKeyState(VK_CONTROL) THEN
                  .SendMessageW hwnd, WM_HSCROLL, SB_PAGERIGHT, 0
               ELSE
                  .SendMessageW hwnd, WM_HSCROLL, SB_LINERIGHT, 0
               END IF
            CASE VK_LEFT
               IF .GetAsyncKeyState(VK_CONTROL) THEN
                  .SendMessageW hwnd, WM_HSCROLL, SB_PAGELEFT, 0
               ELSE
                  .SendMessageW hwnd, WM_HSCROLL, SB_LINELEFT, 0
               END IF
            CASE VK_DOWN
               .SendMessageW hwnd, WM_VSCROLL, SB_LINEDOWN, 0
            CASE VK_UP
               .SendMessageW hwnd, WM_VSCROLL, SB_LINEUP, 0
            CASE VK_NEXT ' = VK_PGDN
               .SendMessageW hwnd, WM_VSCROLL, SB_PAGEDOWN, 0
            CASE VK_PRIOR ' = VK_PGUP
               .SendMessageW hwnd, WM_VSCROLL, SB_PAGEUP, 0
            CASE VK_HOME
               .SendMessageW hwnd, WM_VSCROLL, SB_TOP, 0
            CASE VK_END
               .SendMessageW hwnd, WM_VSCROLL, SB_BOTTOM, 0
         END SELECT
         EXIT FUNCTION

      CASE WM_MOUSEWHEEL
         DIM zDelta AS SHORT = HIWORD(wParam)
         IF (LOWORD(wParam) AND MK_CONTROL) = MK_CONTROL THEN        ' Horizontal scroll
            IF zDelta > 0 THEN                                       ' Scroll to the left
               SendMessageW hwnd, WM_HSCROLL, SB_LINELEFT, 0
            ELSE                                                     ' Scroll to the right
               SendMessageW hwnd, WM_HSCROLL, SB_LINERIGHT, 0
            END IF
         ELSE                                                        ' Vertical scroll
            IF zDelta > 0 THEN                                       ' Scroll up
               SendMessageW hwnd, WM_VSCROLL, SB_LINEUP, 0
            ELSE                                                     ' Scroll down
               SendMessageW hwnd, WM_VSCROLL, SB_LINEDOWN, 0
            END IF
         END IF
         EXIT FUNCTION

      CASE WM_HSCROLL
         ' // Get a pointer to the class
         pGraphCtx = CAST(CGraphCtx PTR, .GetWindowLongPtrW(hwnd, 0))
         IF pGraphCtx = NULL THEN EXIT FUNCTION
         ' // Horizontal scrolling
         SELECT CASE LOWORD(wParam)
            CASE SB_THUMBTRACK
               pGraphCtx->m_OrgX = HIWORD(wParam)
            CASE SB_LINERIGHT
               IF pGraphCtx->m_OrgX < pGraphCtx->m_MaxX - pGraphCtx->m_coord.Right - pGraphCtx->m_coord.Left THEN
                  pGraphCtx->m_OrgX =pGraphCtx->m_OrgX + 10
                  IF pGraphCtx->m_OrgX > pGraphCtx->m_MaxX - pGraphCtx->m_coord.Right - pGraphCtx->m_coord.Left THEN
                     pGraphCtx->m_OrgX = pGraphCtx->m_MaxX - pGraphCtx->m_coord.Right - pGraphCtx->m_coord.Left
                  END IF
               END IF
            CASE SB_LINELEFT
               IF pGraphCtx->m_OrgX > 0 THEN
                  pGraphCtx->m_OrgX = pGraphCtx->m_OrgX - 10
                  IF pGraphCtx->m_OrgX < 0 THEN pGraphCtx->m_OrgX = 0
               END IF
            CASE SB_PAGERIGHT
               nPage = pGraphCtx->m_coord.Right - pGraphCtx->m_coord.Left
               IF pGraphCtx->m_OrgX + nPage < pGraphCtx->m_MaxX - pGraphCtx->m_coord.Right - pGraphCtx->m_coord.Left THEN
                  pGraphCtx->m_OrgX = pGraphCtx->m_OrgX + nPage
               ELSE
                  pGraphCtx->m_OrgX = pGraphCtx->m_MaxX - pGraphCtx->m_coord.Right - pGraphCtx->m_coord.Left
               END IF
            CASE SB_PAGELEFT
               nPage = pGraphCtx->m_coord.Right - pGraphCtx->m_coord.Left
               IF pGraphCtx->m_OrgX - nPage > 0 THEN
                  pGraphCtx->m_OrgX = pGraphCtx->m_OrgX - nPage
               ELSE
                  pGraphCtx->m_OrgX = 0
               END IF
         END SELECT
         si.fMask = SIF_POS
         si.nPos = pGraphCtx->m_OrgX
         .SetScrollInfo hwnd, SB_HORZ, @si, 1
         ' // Invalidates the window to force redrawing
         .InvalidateRect hwnd, NULL, 0
         EXIT FUNCTION

      CASE WM_VSCROLL
         ' // Get a pointer to the class
         pGraphCtx = CAST(CGraphCtx PTR, .GetWindowLongPtrW(hwnd, 0))
         IF pGraphCtx = NULL THEN EXIT FUNCTION
         ' // Vertical scrolling
         SELECT CASE LOWORD(wParam)
            CASE SB_THUMBTRACK
               pGraphCtx->m_OrgY = HIWORD(wParam)
            CASE SB_LINEDOWN
               IF pGraphCtx->m_OrgY < pGraphCtx->m_MaxY - pGraphCtx->m_coord.Bottom - pGraphCtx->m_coord.Top THEN
                  pGraphCtx->m_OrgY = pGraphCtx->m_OrgY + 10
                  IF pGraphCtx->m_OrgY > pGraphCtx->m_MaxY - pGraphCtx->m_coord.Bottom - pGraphCtx->m_coord.Top THEN
                     pGraphCtx->m_OrgY = pGraphCtx->m_MaxY - pGraphCtx->m_coord.Bottom - pGraphCtx->m_coord.Top
                  END IF
               END IF
            CASE SB_LINEUP
               IF pGraphCtx->m_OrgY > 0 THEN
                  pGraphCtx->m_OrgY = pGraphCtx->m_OrgY - 10
                  IF pGraphCtx->m_OrgY < 0 THEN pGraphCtx->m_OrgY = 0
               END IF
            CASE SB_PAGEDOWN
               nPage = pGraphCtx->m_coord.Bottom - pGraphCtx->m_coord.Top
               IF pGraphCtx->m_OrgY + nPage < pGraphCtx->m_MaxY - pGraphCtx->m_coord.Bottom - pGraphCtx->m_coord.Top THEN
                  pGraphCtx->m_OrgY = pGraphCtx->m_OrgY + nPage
               ELSE
                  pGraphCtx->m_OrgY = pGraphCtx->m_MaxY - pGraphCtx->m_coord.Bottom - pGraphCtx->m_coord.Top
               END IF
            CASE SB_PAGEUP
               nPage = pGraphCtx->m_coord.Bottom - pGraphCtx->m_coord.Top
               IF pGraphCtx->m_OrgY - nPage > 0 THEN
                  pGraphCtx->m_OrgY = pGraphCtx->m_OrgY - nPage
               ELSE
                  pGraphCtx->m_OrgY = 0
               END IF
            CASE SB_TOP
               pGraphCtx->m_OrgX = 0
               pGraphCtx->m_OrgY = 0
               si.fMask = SIF_POS
               si.nPos = pGraphCtx->m_OrgX
               .SetScrollInfo hwnd, SB_HORZ, @si, 1
            CASE SB_BOTTOM
               .GetClientRect hwnd, @rc
               pGraphCtx->m_OrgX = pGraphCtx->m_MaxX - rc.Right
               pGraphCtx->m_OrgY = pGraphCtx->m_MaxY - rc.Bottom
               si.fMask = SIF_POS
               si.nPos = pGraphCtx->m_OrgX
               .SetScrollInfo hwnd, SB_HORZ, @si, 1
         END SELECT
         si.fMask = SIF_POS
         si.nPos = pGraphCtx->m_OrgY
         .SetScrollInfo hwnd, SB_VERT, @si, 1
         ' // Invalidates the window to force redrawing
         .InvalidateRect hwnd, NULL, 0
         EXIT FUNCTION

      CASE WM_ENABLE
         ' // Redraws the control
         .InvalidateRect hwnd, NULL, 0
         .UpdateWindow hwnd
         EXIT FUNCTION

      CASE WM_ERASEBKGND
         ' // Get a pointer to the class
         pGraphCtx = CAST(CGraphCtx PTR, .GetWindowLongPtrW(hwnd, 0))
         IF pGraphCtx = NULL THEN EXIT FUNCTION
         ' // If not stretchable...
         IF pGraphCtx->m_Stretchable = FALSE THEN
            ' // Erases only the portion of the window not covered
            ' // by the graphic or image to avoid flicker
            .GetClientRect hwnd, @rc
            IF rc.Right - rc.Left > pGraphCtx->m_vWidth OR _
               rc.Bottom - rc.Top > pGraphCtx->m_vHeight THEN
               hDC = CAST(.HDC, wParam)
               hBrush = .CreateSolidBrush(pGraphCtx->m_bkColor)
               IF hBrush THEN
                  rcFill = rc
'                  rcFill.Left = pGraphCtx->m_vWidth
                  rcFill.Left = pGraphCtx->m_vWidth - pGraphCtx->m_OrgX
                  .FillRect hDC, @rcFill, hBrush
                  rcFill = rc
'                  rcFill.Top = pGraphCtx->m_vHeight
                  rcFill.Top = pGraphCtx->m_vHeight - pGraphCtx->m_OrgY
                  .FillRect hDC, @rcFill, hBrush
                  .DeleteObject hBrush
               END IF
            END IF
         END IF
         FUNCTION = 1
         EXIT FUNCTION

      CASE WM_SIZE
         ' // Get a pointer to the class
         pGraphCtx = CAST(CGraphCtx PTR, .GetWindowLongPtrW(hwnd, 0))
         IF pGraphCtx = NULL THEN EXIT FUNCTION
         ' // Updates virtual window origins
         DIM Increment AS LONG = LOWORD(lParam) - pGraphCtx->m_coord.Right - pGraphCtx->m_coord.Left
         IF Increment < 0 THEN Increment = 0
         IF (Increment > 0) AND (pGraphCtx->m_OrgX >= (LOWORD(lParam) - pGraphCtx->m_coord.Right - pGraphCtx->m_coord.Left)) THEN pGraphCtx->m_OrgX = pGraphCtx->m_OrgX - Increment
         IF pGraphCtx->m_OrgX < 0 THEN  pGraphCtx->m_OrgX = 0
         Increment = HIWORD(lParam) - pGraphCtx->m_coord.Bottom - pGraphCtx->m_coord.Top
         IF Increment < 0 THEN Increment = 0
         IF(Increment > 0) AND (pGraphCtx->m_OrgY >= (HIWORD(lParam) - pGraphCtx->m_coord.Bottom - pGraphCtx->m_coord.Top)) THEN pGraphCtx->m_OrgY = pGraphCtx->m_OrgY - Increment
         IF pGraphCtx->m_OrgY < 0 THEN pGraphCtx->m_OrgY = 0
         ' // Stores new window extents
         pGraphCtx->m_coord.Right  = LOWORD(lParam)
         pGraphCtx->m_coord.Bottom = HIWORD(lParam)
         ' // If resizable or stretchable, disable scroll bars
         IF pGraphCtx->m_Resizable = TRUE OR pGraphCtx->m_Stretchable = TRUE THEN
            ' // Initializes scroll bar ranges
            .GetClientRect hwnd, @rc
            pGraphCtx->m_MaxX = rc.Right - rc.Left
            pGraphCtx->m_MaxY = rc.Bottom - rc.Top
            si.cbSize = SIZEOF(si)
            si.fMask = SIF_RANGE
            si.nMin = 0
            si.nMax = pGraphCtx->m_MaxX - rc.Right - rc.Left
            .SetScrollInfo hwnd, SB_HORZ, @si, 1
            si.nMax = pGraphCtx->m_MaxY - rc.Bottom - rc.Top
            .SetScrollInfo hwnd, SB_VERT, @si, 1
         ELSE
            ' // Reinitializes scroll bar ranges
            si.cbSize = SIZEOF(si)
            si.fMask = SIF_RANGE OR SIF_POS
            si.nMin = 0
            si.nMax = pGraphCtx->m_MaxX - pGraphCtx->m_coord.Right - pGraphCtx->m_coord.Left
            si.nPos = pGraphCtx->m_OrgX
            .SetScrollInfo hwnd, SB_HORZ, @si, 1
            si.nMax = pGraphCtx->m_MaxY - pGraphCtx->m_coord.Bottom - pGraphCtx->m_coord.Top
            si.nPos = pGraphCtx->m_OrgY
            .SetScrollInfo hwnd, SB_VERT, @si, 1
         END IF
         ' // Invalidates the window to force redrawing
         .InvalidateRect hwnd, NULL, 0
         EXIT FUNCTION

      CASE WM_PRINTCLIENT
         ' // Get a pointer to the class
         pGraphCtx = CAST(CGraphCtx PTR, .GetWindowLongPtrW(hwnd, 0))
         IF pGraphCtx = NULL THEN EXIT FUNCTION
         ' // Copies the bitmap to the provided device context
         hDC = CAST(.HDC, wParam)
         .GetClientRect(hwnd, @rc)
         IF pGraphCtx->m_Stretchable = FALSE THEN
            .BitBlt(hDC, 0, 0, .GetDeviceCaps(hDC, HORZRES), .GetDeviceCaps(hDC, VERTRES), _
               pGraphCtx->m_hMemDc, rc.Left + pGraphCtx->m_OrgX, rc.Top + pGraphCtx->m_OrgY, SRCCOPY)
         ELSE
            .SetStretchBltMode hDC, pGraphCtx->m_StretchMode
            IF pGraphCtx->m_StretchMode = HALFTONE THEN .SetBrushOrgEx hDC, 0, 0, NULL
            .StretchBlt(hDC, 0, 0, .GetDeviceCaps(hDC, HORZRES), .GetDeviceCaps(hDC, VERTRES), _
               pGraphCtx->m_hMemDc, 0, 0, pGraphCtx->m_vWidth, pGraphCtx->m_vHeight, SRCCOPY)
         END IF
         EXIT FUNCTION

      CASE WM_PAINT
         ' // Get a pointer to the class
         pGraphCtx = CAST(CGraphCtx PTR, .GetWindowLongPtrW(hwnd, 0))
         IF pGraphCtx = NULL THEN EXIT FUNCTION
         ' // Copies the bitmap to the control's window
         hDC = .BeginPaint(hwnd, @ps)
         IF pGraphCtx->m_Stretchable = FALSE THEN
            .BitBlt(hDC, ps.rcPaint.Left, ps.rcPaint.Top, _
                 ps.rcPaint.Right - ps.rcPaint.Left, _
                 ps.rcPaint.Bottom - ps.rcPaint.Top, _
                 pGraphCtx->m_hMemDc, ps.rcPaint.Left + pGraphCtx->m_OrgX, ps.rcPaint.Top + pGraphCtx->m_OrgY, SRCCOPY)
         ELSE
            .GetClientRect hwnd, @rc
            .SetStretchBltMode hDC, pGraphCtx->m_StretchMode
            IF pGraphCtx->m_StretchMode = HALFTONE THEN .SetBrushOrgEx hDC, 0, 0, NULL
            .StretchBlt(hDC, ps.rcPaint.Left, ps.rcPaint.Top, _
                 ps.rcPaint.Right - ps.rcPaint.Left, _
                 ps.rcPaint.Bottom - ps.rcPaint.Top, _
                 pGraphCtx->m_hMemDc, 0, 0,pGraphCtx->m_vWidth, pGraphCtx->m_vHeight, SRCCOPY)
         END IF
         .EndPaint hwnd, @ps
         EXIT FUNCTION

   END SELECT

   ' // Default processing for other messages.
   FUNCTION = DefWindowProcW(hWnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the handle of the control
' ========================================================================================
PRIVATE FUNCTION CGraphCtx.hWindow () AS HWND
   FUNCTION = m_hCtl
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the size of the virtual buffer.
' Parameters:
' * nWidth = Width, in pixels, of the virtual buffer.
' * nHeight = Height, in pixels, of the virtual buffer.
' Return value:
' * If the function succeeds, the return value is S_OK.
' ========================================================================================
PRIVATE FUNCTION CGraphCtx.SetVirtualBufferSize (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG) AS LONG

   ' // Clone the memory bitmap to restore it later
   DIM bi AS BITMAPINFO
   DIM rcClone AS RECT
   .SetRect (@rcClone, 0, 0, nWidth, nHeight)
   DIM hDC AS .HDC = .GetDC(m_hCtl)
   DIM hCloneDC AS .HDC = .CreateCompatibleDC(hDC)
'   DIM hCloneBmp AS HBITMAP = .CreateBitmap(m_vWidth, m_vHeight, 1, m_BitsPerPel, NULL)
   ' // --- Use CreateDIBSection instead of CreateBitmap --- //
   bi.bmiHeader.biSize = SIZEOF(bi.bmiHeader)
   bi.bmiHeader.biWidth = m_vHeight
   bi.bmiHeader.biHeight = nHeight
   bi.bmiHeader.biPlanes = 1
   bi.bmiHeader.biBitCount = m_BitsPerPel
   bi.bmiHeader.biCompression = BI_RGB
   DIM hCloneBmp AS HBITMAP = .CreateDIBSection(hDC, @bi, DIB_RGB_COLORS, NULL, NULL, 0)
   ' // -------------- //
   IF hCloneBmp THEN
      .SelectObject(hCloneDc, hCloneBmp)
      .BitBlt(hCloneDC, 0, 0, m_vWidth, m_vHeight, m_hMemDc, 0, 0, SRCCOPY)
   END IF
   .ReleaseDc m_hCtl, hDc
   ' // Restore the original bitmap
   SelectObject(m_hMemDc, m_hOldBmp)
   ' // Destroys the bitmap and the memory device context
   IF m_hBmp THEN .DeleteObject m_hBmp
   IF m_hMemDc THEN .DeleteObject m_hMemDc
   m_vWidth = nWidth
   m_vHeight = nHeight

   ' // Creates a virtual compatible memory device context and bitmap.
   hDC = .GetDC(m_hCtl)
   m_hMemDc = .CreateCompatibleDc(hDC)
   bi.bmiHeader.biSize = SIZEOF(bi.bmiHeader)
   bi.bmiHeader.biWidth = nWidth
   bi.bmiHeader.biHeight = nHeight
   bi.bmiHeader.biPlanes = 1
   bi.bmiHeader.biBitCount = m_BitsPerPel
   bi.bmiHeader.biCompression = BI_RGB
   m_hBmp = .CreateDIBSection(hDC, @bi, DIB_RGB_COLORS, m_ppvBits, NULL, 0)
   m_hOldBmp = .SelectObject(m_hMemDc, m_hBmp)

   ' // Adjusts the coordinates
   m_coord.Right = m_coord.Left + nWidth
   m_coord.Bottom = m_coord.Top + nHeight
   ' // Erases the background with the specified or default color
   IF m_bkColor THEN
      DIM hBrush AS .HBRUSH = .CreateSolidBrush(m_bkColor)
      IF hBrush THEN
         .FillRect m_hMemDc, @m_coord, hBrush
         .DeleteObject hBrush
      END IF
   ELSE
      .FillRect m_hMemDc, @m_coord, CAST(.HBRUSH, .GetClassLongPtr(.GetParent(m_hCtl), GCLP_HBRBACKGROUND))
   END IF
   .ReleaseDc m_hCtl, hDC

   IF m_Resizable = FALSE AND m_Stretchable = FALSE THEN
      ' // Adds scrollbars if needed
      DIM rcParent AS RECT, dwStyle AS DWORD
      .GetClientRect .GetParent(m_hCtl), @rcParent
      IF nWidth > rcParent.Right OR nHeight > rcParent.Bottom THEN
         dwStyle = .GetWindowLongPtrW(m_hCtl, GWL_STYLE)
         dwStyle = dwStyle OR WS_HSCROLL OR WS_VSCROLL
         .SetWindowLongPtrW(m_hCtl, GWL_STYLE, dwStyle)
      END IF
      ' // Initializes scroll bar ranges
      m_MaxX = m_coord.Right - m_coord.Left
      m_MaxY = m_coord.Bottom - m_coord.Top
      DIM si AS SCROLLINFO
      si.cbSize = SIZEOF(si)
      si.fMask = SIF_RANGE
      si.nMin = 0
      si.nMax = m_MaxX - m_coord.Right - m_coord.Left
      .SetScrollInfo m_hCtl, SB_HORZ, @si, 1
      si.nMax = m_MaxY - m_coord.Bottom - m_coord.Top
      .SetScrollInfo m_hCtl, SB_VERT, @si, 1
      ' // If a scrollbar isn't needed, hide it
      DIM rc AS RECT
      .GetClientRect m_hCtl, @rc
      IF nWidth <= rc.Right THEN ShowScrollBar m_hCtl, SB_HORZ, FALSE
      IF nHeight <= rc.Bottom THEN ShowScrollBar m_hCtl, SB_VERT, FALSE
      ' // Invalidates the window to force redrawing
      .InvalidateRect m_hCtl, NULL, 0
   END IF

   ' // Copy the cloned bitmap
   IF hCloneBmp THEN
      .BitBlt(m_hMemDc, 0, 0, m_vWidth, m_vHeight, hCloneDc, 0, 0, SRCCOPY)
      .SelectObject(hCloneDc, 0)
      .DeleteDC hCloneDC
      .DeleteObject hCloneBmp
      .InvalidateRect m_hCtl, NULL, 0
   END IF

   FUNCTION = S_OK

END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the handle of the memory device context of the control.
' ========================================================================================
PRIVATE FUNCTION CGraphCtx.GetMemDC () AS HDC
   FUNCTION = m_hMemDc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the handle of the compatible bitmap.
' ========================================================================================
PRIVATE FUNCTION CGraphCtx.GethBmp () AS HBITMAP
   FUNCTION = m_hBmp
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the location of the DIB bit values.
' ========================================================================================
PRIVATE FUNCTION CGraphCtx.GetBits () AS ANY PTR
   FUNCTION = m_ppvBits
END FUNCTION
' ========================================================================================

' ========================================================================================
' Clears the graphic control with the specified RGB color.
' Parameter:
' * RGBColor = RGB color used to fill the control.
' Return value:
' * If the function succeeds, the return value is TRUE.
'   If the function fails, the return value is FALSE.
' ========================================================================================
PRIVATE FUNCTION CGraphCtx.Clear (BYVAL RGBColor AS COLORREF) AS BOOLEAN
   IF m_hMemDc = NULL THEN EXIT FUNCTION
   DIM hBrush AS .HBRUSH = .CreateSolidBrush(RGBColor)
   DIM rc AS RECT
   .GetClientRect m_hCtl, @rc
   IF m_vWidth THEN rc.Right = rc.Left + m_vWidth
   IF m_vHeight THEN rc.Bottom = rc.Top + m_vHeight
   FUNCTION = .FillRect(m_hMemDc, @rc, hBrush)
   .DeleteObject hBrush
   m_bkColor = RGBColor
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the width of the virtual buffer.
' ========================================================================================
PRIVATE FUNCTION CGraphCtx.GetVirtualBufferWidth () AS LONG
   FUNCTION = m_vWidth
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the height of the virtual buffer.
' ========================================================================================
PRIVATE FUNCTION CGraphCtx.GetVirtualBufferHeight () AS LONG
   FUNCTION = m_vHeight
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the value of the stretchable property
' Return value:
' * TRUE or FALSE
' ========================================================================================
PRIVATE PROPERTY CGraphCtx.Stretchable () AS BOOLEAN
   PROPERTY = m_Stretchable
END PROPERTY
' ========================================================================================

' ========================================================================================
' Sets the value of the stretchable property
' Parameters:
' * bResizable = TRUE or FALSE
' Return value:
' * The previous value of the property
' ========================================================================================
PRIVATE PROPERTY CGraphCtx.Stretchable (BYVAL bStretchable AS BOOLEAN)
   m_Stretchable = bStretchable
END PROPERTY
' ========================================================================================

' ========================================================================================
' Returns the value of the stretch mode property
' ========================================================================================
PRIVATE PROPERTY CGraphCtx.StretchMode () AS LONG
   PROPERTY = m_StretchMode
END PROPERTY
' ========================================================================================

' ========================================================================================
' Sets the value of the stretch mode property
' Parameter:
' * nStretchMode = The stretching mode.
' BLACKONWHITE
'    Performs a Boolean AND operation using the color values for the eliminated and existing
'    pixels. If the bitmap is a monochrome bitmap, this mode preserves black pixels at the
'    expense of white pixels.
' COLORONCOLOR
'   Deletes the pixels. This mode deletes all eliminated lines of pixels without trying to
'   preserve their information.
' HALFTONE
'   Maps pixels from the source rectangle into blocks of pixels in the destination rectangle.
'   The average color over the destination block of pixels approximates the color of the
'   source pixels.
'   After setting the HALFTONE stretching mode, an application must call the SetBrushOrgEx
'   function to set the brush origin. If it fails to do so, brush misalignment occurs.
' STRETCH_ANDSCANS
'   Same as BLACKONWHITE.
' STRETCH_DELETESCANS
'   Same as COLORONCOLOR.
' STRETCH_HALFTONE
'   Same as HALFTONE.
' STRETCH_ORSCANS
'   Same as WHITEONBLACK.
' WHITEONBLACK
'   Performs a Boolean OR operation using the color values for the eliminated and existing
'   pixels. If the bitmap is a monochrome bitmap, this mode preserves white pixels at the
'   expense of black pixels.
' Return value:
' * The previous value of the property
' ========================================================================================
PRIVATE PROPERTY CGraphCtx.StretchMode (BYVAL nStretchMode AS LONG)
   m_StretchMode = nStretchMode
END PROPERTY
' ========================================================================================

' ========================================================================================
' Returns the value of the resizable property
' Return value:
' * TRUE or FALSE
' ========================================================================================
PRIVATE PROPERTY CGraphCtx.Resizable () AS BOOLEAN
   PROPERTY = m_Resizable
END PROPERTY
' ========================================================================================

' ========================================================================================
' Sets the value of the resizable property
' Parameter:
' * bResizable = TRUE or FALSE
' Return value:
' * The previous value of the property
' ========================================================================================
PRIVATE PROPERTY CGraphCtx.Resizable (BYVAL bResizable AS BOOLEAN)
   m_Resizable = bResizable
END PROPERTY
' ========================================================================================

' ========================================================================================
' Shows the specified image in the specified device context
' Parameters:
' - wszFileName = Path of the file.
' Remarks:
'   A quirk in the GDI+ GdipCreateBitmapFromFile function causes that black and white
'   images are loaded with increased contrast. Therefore, it's better to use the
'   CreateBitmapFromFile method.
' ========================================================================================
PRIVATE SUB CGraphCtx.LoadImageFromFile (BYREF wszFileName AS WSTRING)
   ' // Initialize Gdiplus
   DIM token AS ULONG_PTR = AfxGdipInit
   IF token = NULL THEN EXIT SUB
   ' // Load the image from file and display it in the control
   DIM pBitmap AS GpBitmap PTR, pGraphics AS GpGraphics PTR, nWidth AS DWORD, nHeight AS DWORD
   GdipCreateBitmapFromFile(wszFileName, @pBitmap)
   IF pBitmap THEN
      ' // Get the width and height of the image
      GdipGetImageWidth(CAST(GpImage PTR, pBitmap), @nWidth)
      GdipGetImageHeight(CAST(GpImage PTR, pBitmap), @nHeight)
      ' // Set the virtual buffer size of the control
      this.SetVirtualBufferSize(nWidth, nHeight)
      this.Clear m_bkcolor
      ' // Create a graphics object from the memory DC of the control
      GdipCreateFromHDC(m_hMemDC, @pGraphics)
      ' // Draw the image
      IF pGraphics THEN GdipDrawImageRectI(pGraphics, CAST(GpImage PTR, pBitmap), 0, 0, nWidth, nHeight)
      ' // Dispose the image and delete the graphics object
      GdipDisposeImage(CAST(GpImage PTR, pBitmap))
      IF pGraphics THEN GdipDeleteGraphics(pGraphics)
   END IF
   ' // Shutdown Gdiplus
   GdiplusShutdown token
END SUB
' ========================================================================================

' ========================================================================================
' Shows the specified image in the specified device context
' Parameters:
' - wszFileName = Path of the file.
' - dimPercent  = Percent of dimming (1-99)
' - bGrayScale  = TRUE or FALSE. Convert to gray scale.
' ========================================================================================
PRIVATE SUB CGraphCtx.CreateBitmapFromFile (BYREF wszFileName AS WSTRING, BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE)
   IF LEN(wszFileName) = 0 THEN EXIT SUB
'   ' // Initialize Gdiplus
'   DIM token AS ULONG_PTR = AfxGdipInit
'   IF token = NULL THEN EXIT SUB
   DIM hbmp AS HBITMAP = AfxGdipBitmapFromFile(wszFileName, dimPercent, bGrayScale)
   IF hbmp THEN
      DIM bm AS BITMAP, hMemDC AS HDC
      hMemDC = .CreateCompatibleDC(m_hMemDC)
      IF hMemDC THEN
         IF .GetObject(hbmp, SIZEOF(BITMAP), @bm) THEN
            this.SetVirtualBufferSize(bm.bmWidth, bm.bmHeight)
            this.Clear m_bkcolor
            .SelectObject hMemDC, hbmp
            .BitBlt m_hMemDC, 0, 0, bm.bmWidth, bm.bmHeight, hMemDC, 0, 0, SRCCOPY
         END IF
         DeleteDC hMemDC
      END IF
   END IF
'   ' // Shutdown Gdiplus
'   GdiplusShutdown token
END SUB
' ========================================================================================

' ========================================================================================
' Saves the image to file.
' Parameters:
' - wszFileName = Path name for the image to be saved.
' - wszMimeType = Mime type (default: "image/bmp").
'   "image/bmp" = Bitmap (.bmp)
'   "image/gif" = GIF (.gif)
'   "image/jpeg" = JPEG (.jpg)
'   "image/png" = PNG (.png)
'   "image/tiff" = TIFF (.tiff)
' Return value:
' If the method succeeds, it returns Ok, which is an element of the Status enumeration.
' If the method fails, it returns one of the other elements of the Status enumeration.
' ========================================================================================
PRIVATE FUNCTION CGraphCtx.SaveImage (BYREF wszFileName AS WSTRING, BYREF wszMimeType AS WSTRING = "image/bmp") AS LONG
   FUNCTION = AfxGdipSaveHBITMAPToFile(m_hBmp, wszFileName, wszMimeType)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGraphCtx.SaveImageToBmp (BYREF wszFileName AS WSTRING) AS LONG
   FUNCTION = AfxGdipSaveHBITMAPToFile(m_hBmp, wszFileName, "image/bmp")
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGraphCtx.SaveImageToJpeg (BYREF wszFileName AS WSTRING) AS LONG
   FUNCTION = AfxGdipSaveHBITMAPToFile(m_hBmp, wszFileName, "image/jpeg")
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGraphCtx.SaveImageToPng (BYREF wszFileName AS WSTRING) AS LONG
   FUNCTION = AfxGdipSaveHBITMAPToFile(m_hBmp, wszFileName, "image/png")
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGraphCtx.SaveImageToGif (BYREF wszFileName AS WSTRING) AS LONG
   FUNCTION = AfxGdipSaveHBITMAPToFile(m_hBmp, wszFileName, "image/gif")
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGraphCtx.SaveImageToTiff (BYREF wszFileName AS WSTRING) AS LONG
   FUNCTION = AfxGdipSaveHBITMAPToFile(m_hBmp, wszFileName, "image/tiff")
END FUNCTION
' ========================================================================================

' ========================================================================================
' Prints the image in the default printer.
' Parameters:
' - bStretch     = Stretch the image.
' - nStretchMode = Stretching mode. Default value = InterpolationModeHighQualityBicubic.
'   InterpolationModeLowQuality = 1
'   InterpolationModeHighQuality = 2
'   InterpolationModeBilinear = 3
'   InterpolationModeBicubic = 4
'   InterpolationModeNearestNeighbor = 5
'   InterpolationModeHighQualityBilinear = 6
'   InterpolationModeHighQualityBicubic = 7
' Return value: Returns TRUE if the bitmap has been printed successfully, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CGraphCtx.PrintImage (BYVAL bStretch AS BOOLEAN = FALSE, BYVAL nStretchMode AS LONG = InterpolationModeHighQualityBicubic) AS LONG
   FUNCTION = AfxGdipPrintHBITMAP(m_hBmp, bStretch, nStretchMode)
END FUNCTION
' ========================================================================================

END NAMESPACE
