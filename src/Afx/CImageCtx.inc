' ########################################################################################
' Microsoft Windows
' File: CImageCtx.inc
' Contents: GDI+ image control
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "Afx/CWindow.inc"
#include once "Afx/AfxGdiPlus.inc"

NAMESPACE Afx

CONST GDIP_IMAGECTX_AUTOSIZE    = 0   ' // Autoadjusts the image to the width or height of the control
CONST GDIP_IMAGECTX_ACTUALSIZE  = 1   ' // Shows the image with its actual size
CONST GDIP_IMAGECTX_FITTOWIDTH  = 2   ' // Adjusts the image to the width of the control
CONST GDIP_IMAGECTX_FITTOHEIGHT = 3   ' // Adjusts the image to the height of the control
CONST GDIP_IMAGECTX_STRETCH     = 4   ' // Adjusts the image to the height and width of the control

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
' CImageCtx class
' ========================================================================================
TYPE CImageCtx

   Private:
      m_hCtl    AS HWND             ' // Handle to the control
      m_hFont   AS HFONT            ' // Handle to the font used by the control
      m_pImage  AS GpImage PTR      ' // Image object pointer
      m_token   AS ULONG_PTR        ' // Token returned by GdiplusStartup
      m_BkColor AS COLORREF         ' // RGB background color
      m_Fit     AS LONG             ' // Fit to width/height
      m_InterpolationMode AS LONG   ' // Interpolation mode

   Public:

      DECLARE CONSTRUCTOR (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF wszTitle AS WSTRING = "", _
         BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
         BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL lpParam AS LONG_PTR = 0)
      DECLARE DESTRUCTOR
      DECLARE FUNCTION hWindow () AS HWND
      DECLARE FUNCTION GetImagePtr () AS GpImage PTR
      DECLARE FUNCTION DrawImage (BYVAL hwnd AS HWND, BYVAL hDC AS HDC, BYVAL pImage AS GpImage PTR) AS LONG
      DECLARE FUNCTION LoadImageFromFile (BYREF wszFileName AS WSTRING) AS LONG
      DECLARE FUNCTION LoadBitmapFromResource (BYVAL hInstance AS HINSTANCE, BYREF wszResourceName AS WSTRING) AS LONG
      DECLARE FUNCTION LoadImageFromResource (BYVAL hInstance AS HINSTANCE, BYREF wszResourceName AS WSTRING) AS LONG
      DECLARE FUNCTION SetImageAdjustment (BYVAL ImageAdjustment AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE) AS LONG
      DECLARE FUNCTION GetImageAdjustment () AS LONG
      DECLARE FUNCTION GetImageWidth () AS LONG
      DECLARE FUNCTION GetImageHeight () AS LONG
      DECLARE FUNCTION SetBkColor (BYVAL clr AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE) AS LONG
      DECLARE FUNCTION GetBkColor () AS LONG
      DECLARE SUB SetInterpolationMode (BYVAL InterpolationMode AS LONG)
      DECLARE FUNCTION GetInterpolationMode () AS LONG
      DECLARE SUB Redraw
      DECLARE SUB Clear
      DECLARE STATIC FUNCTION CImageCtxProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' CImageCtx class constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CImageCtx (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF wszTitle AS WSTRING = "", _
   BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
   BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL lpParam AS LONG_PTR = 0)

   ' // Register the class
   DIM wAtom AS ATOM
   DIM wcexw AS WNDCLASSEXW
   DIM wszClassName AS WSTRING * 260 = "AFX_GDIPIMAGECTX"
   IF .GetClassInfoExW(.GetModuleHandleW(NULL), @wszClassName, @wcexw) = 0 THEN
      ' // Fill the WNDCLASSEXW structure
      WITH wcexw
         .cbSize        = SIZEOF(wcexw)
         .style         = CS_DBLCLKS OR CS_HREDRAW OR CS_VREDRAW
         .lpfnWndProc   = @CImageCtxProc
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
   END IF

   ' // Initialize GDI+
   m_token = AfxGdipInit
   ' // Default interpolation mode
   m_InterpolationMode = InterpolationModeHighQualityBicubic

END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CImageCtx class destructor
' ========================================================================================
PRIVATE DESTRUCTOR CImageCtx
   ' // Delete the font used by the control
   IF m_hFont THEN DeleteObject m_hFont
   ' // Destroy the Image object
   IF m_pImage THEN GdipDisposeImage(m_pImage)
   ' // Shutdown GDI+
   IF m_token THEN GdiplusShutdown m_token
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Window procedure
' ========================================================================================
PRIVATE FUNCTION CImageCtx.CImageCtxProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   DIM hDC AS HDC, ps AS PAINTSTRUCT, hdr AS NMHDR, pImageCtx AS CImageCtx PTR

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

      CASE WM_ERASEBKGND
         ' // Don't erase the background to avoid flicker
         FUNCTION = 1
         EXIT FUNCTION

      CASE WM_PRINTCLIENT
         ' // Get a pointer to the class
         pImageCtx = CAST(CImageCtx PTR, .GetWindowLongPtrW(hwnd, 0))
         IF pImageCtx = NULL THEN EXIT FUNCTION
         ' // Draws the image in the provided device context
         hDC = CAST(.HDC, wParam)
         pImageCtx->DrawImage(hwnd, hDC, pImageCtx->m_pImage)
         EXIT FUNCTION

      CASE WM_PAINT
         ' // Get a pointer to the class
         pImageCtx = CAST(CImageCtx PTR, .GetWindowLongPtrW(hwnd, 0))
         IF pImageCtx = NULL THEN EXIT FUNCTION
         ' // Prepares the window for painting
         hDC = .BeginPaint(hwnd, @ps)
         ' // Draws the image
         pImageCtx->DrawImage(hwnd, hDC, pImageCtx->m_pImage)
         ' // Ends the painting of the window
         .EndPaint hwnd, @ps
         EXIT FUNCTION

      CASE WM_KEYDOWN
         SELECT CASE (wParam)
            ' // Must handle tab key by ourselves, since we have taken full control
            CASE VK_TAB
               IF HIWORD(.GetKeyState(VK_SHIFT)) = 0 THEN
                  .SetFocus .GetNextDlgTabItem(GetParent(hwnd), hwnd, 0)
               ELSE
                  .SetFocus .GetNextDlgTabItem(GetParent(hwnd), hwnd, -1)
               END IF
         END SELECT

      CASE WM_LBUTTONDOWN
         ' // Forwards the message to the parent window
         hdr.hwndFrom = hwnd
         hdr.idFrom = GetDlgCtrlId(hwnd)
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

      CASE WM_ENABLE
         ' // Redraws the control
         .InvalidateRect hwnd, NULL, 0
         .UpdateWindow hwnd
         EXIT FUNCTION

   END SELECT

   ' // Default processing for other messages.
   FUNCTION = .DefWindowProcW(hWnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the handle of the control
' ========================================================================================
PRIVATE FUNCTION CImageCtx.hWindow () AS HWND
   FUNCTION = m_hCtl
END FUNCTION
' ========================================================================================

' ========================================================================================
' Draws the image
' Returns StatusOk or an error code.
' ========================================================================================
PRIVATE FUNCTION CImageCtx.DrawImage (BYVAL hwnd AS HWND, BYVAL hDC AS HDC, BYVAL pImage AS GpImage PTR) AS LONG

   DIM pImageCtx AS CImageCtx PTR    ' // Pointer to the CImageCtx class
   DIM hStatus AS GpStatus           ' // Result code
   DIM pGraphics AS GpGraphics PTR   ' // Graphics object pointer
   DIM hBgBrush AS HBRUSH            ' // Background brush
   DIM ImageWidth AS LONG            ' // Image width
   DIM ImageHeight AS LONG           ' // Image heigth
   DIM p AS LONG                     ' // Percent
   DIM w AS LONG                     ' // Width
   DIM h AS LONG                     ' // Height
   DIM cx AS LONG                    ' // x-coordinate for center fill
   DIM cy AS LONG                    ' // y-coordinate for center fill
   DIM nFit AS LONG                  ' // Fit
   DIM rc AS RECT                    ' // client area
   DIM rcFill AS RECT                ' // Fill area

   ' // Check parameters validity
   IF hWnd = NULL OR hDC = NULL OR pImage = NULL THEN RETURN E_POINTER
   ' // Get a pointer to the class
   pImageCtx = CAST(CImageCtx PTR, .GetWindowLongPtrW(hwnd, 0))
   IF pImageCtx = NULL THEN EXIT FUNCTION
   ' // Retrieves the coordinates of the window's client area
   .GetClientRect hwnd, @rc
   ' // Retrieves the width and height of the image
   hStatus = .GdipGetImageWidth(pImage, @ImageWidth)
   hStatus = .GdipGetImageHeight(pImage, @ImageHeight)
   ' // Creates a Graphics object
   hStatus = .GdipCreateFromHDC(hdc, @pGraphics)
   ' // Draws the image
   IF pGraphics THEN
      nFit = pImageCtx->m_Fit
      IF nFit = GDIP_IMAGECTX_AUTOSIZE THEN
         IF (ImageWidth \ (rc.Right - rc.Left)) > (ImageHeight \ (rc.Bottom - rc.Top)) THEN
            nFit = GDIP_IMAGECTX_FITTOWIDTH
         ELSE
            nFit = GDIP_IMAGECTX_FITTOHEIGHT
         END IF
      END IF
      IF nFit = GDIP_IMAGECTX_ACTUALSIZE THEN
         ' // If the width of the image is smaller than the width of the control
         ' // calculate the position for a centered image
         w = ImageWidth
         h = ImageHeight
         IF w < (rc.Right - rc.Left) THEN
            cx = ((rc.Right - rc.Left) - w) / 2
         END IF
         IF h < (rc.Bottom - rc.Top) THEN
            cy = ((rc.Bottom - rc.Top) - h) / 2
         END IF
         ' // Draw the image
         hStatus = .GdipDrawImageRectI(pGraphics, pImageCtx->m_pImage, cx, cy, ImageWidth, ImageHeight)
         ' // Creates a brush and fills the client area of the window not covered by the image
         hBgBrush = .CreateSolidBrush(pImageCtx->m_BkColor)
         .GetClientRect hwnd, @rcFill
         rcFill.Left = 0
         rcFill.Right = cx
         .FillRect hDC, @rcFill, hBgBrush
         .GetClientRect hwnd, @rcFill
         rcFill.Left = cx + w
         .FillRect hDC, @rcFill, hBgBrush
         .GetClientRect hwnd, @rcFill
         rcFill.Left = cx
         rcFill.Right = cx + w
         rcFill.Bottom = cy
         .FillRect hDC, @rcFill, hBgBrush
         .GetClientRect hwnd, @rcFill
         rcFill.Left = cx
         rcFill.Right = cx + w
         rcFill.Top = cy + h
         .FillRect hDC, @rcFill, hBgBrush
         ' // Deletes the brush
         .DeleteObject hBgBrush
      ELSEIF nFit = GDIP_IMAGECTX_STRETCH THEN
        ' Stretch or shrunk to fill the controls client area
         w = rc.Right - rc.Left
         h = rc.Bottom - rc.Top
         IF m_InterpolationMode THEN GdipSetInterpolationMode(pGraphics, m_InterpolationMode)
         hStatus = .GdipDrawImageRectI(pGraphics, pImageCtx->m_pImage, 0, 0, w, h)
      ELSEIF nFit = GDIP_IMAGECTX_FITTOWIDTH THEN
         ' // Calculate the percent of reduction or ampliation
         ' // and adjust the image height
         p = ((rc.Right - rc.Left) * 100) / ImageWidth
         w = rc.Right - rc.Left
         h = ImageHeight * p / 100
         ' // If the height of the image is smaller than the height of the control
         ' // calculate the position for a centered image
         IF h < (rc.Bottom - rc.Top) THEN
            cy = ((rc.Bottom - rc.Top) - h) / 2
         END IF
         IF m_InterpolationMode THEN GdipSetInterpolationMode(pGraphics, m_InterpolationMode)
         hStatus = .GdipDrawImageRectI(pGraphics, pImageCtx->m_pImage, 0, cy, w, h)
         ' // Creates a brush and fills the client area of the window not covered by the image
         hBgBrush = .CreateSolidBrush(pImageCtx->m_BkColor)
         .GetClientRect hwnd, @rcFill
         rcFill.Bottom = cy
         .FillRect hDC, @rcFill, hBgBrush
         .GetClientRect hwnd, @rcFill
         rcFill.Top = cy + h
         .FillRect hDC, @rcFill, hBgBrush
         ' // Deletes the brush
         .DeleteObject hBgBrush
      ELSE
         ' // Calculate the percent of reduction or ampliation
         ' // and adjust the image width
         IF ImageHeight <> rc.Bottom - rc.Top THEN
            p = ((rc.Bottom - rc.Top) * 100) /  ImageHeight
            h = rc.Bottom - rc.Top
            w = ImageWidth * p / 100
         ELSE
            h = ImageHeight
            w = ImageWidth
         END IF
         ' // If the width of the image is smaller than the width of the control
         ' // calculate the position for a centered image
         IF w < (rc.Right - rc.Left) THEN
            cx = ((rc.Right - rc.Left) - w) / 2
         END IF
         ' // Draws the image
         IF m_InterpolationMode THEN GdipSetInterpolationMode(pGraphics, m_InterpolationMode)
         hStatus = .GdipDrawImageRectI(pGraphics, pImageCtx->m_pImage, cx, 0, w, h)
         ' // Creates a brush and fills the client area of the window not covered by the image
         hBgBrush = .CreateSolidBrush(pImageCtx->m_BkColor)
         .GetClientRect hwnd, @rcFill
         rcFill.Left = 0
         rcFill.Right = cx
         .FillRect hDC, @rcFill, hBgBrush
         .GetClientRect hwnd, @rcFill
         rcFill.Left = cx + w
         .FillRect hDC, @rcFill, hBgBrush
         ' // Deletes the brush
         .DeleteObject hBgBrush
      END IF
      ' // Deletes the Graphics object
      .GdipDeleteGraphics(pGraphics)
   ELSE
      ' // Creates a brush and fills the client area of the window
      hBgBrush = .CreateSolidBrush(pImageCtx->m_BkColor)
      .GetClientRect hwnd, @rcFill
      .FillRect hDC, @rcFill, hBgBrush
      ' // Deletes the brush
      .DeleteObject hBgBrush
   END IF

END FUNCTION
' ========================================================================================

' ========================================================================================
' Loads an image from file using GDI+
' Parameters:
' - wszFileName = Fully qualified path and filename of the image file to load.
' Return value:
' - Ok (0) or an error code.
' ========================================================================================
PRIVATE FUNCTION CImageCtx.LoadImageFromFile (BYREF wszFileName AS WSTRING) AS LONG
   IF LEN(wszFileName) = 0 THEN RETURN E_INVALIDARG
   ' // Dispose a previous instance of the Image object, if any
   IF m_pImage THEN
      .GdipDisposeImage(m_pImage)
      m_pImage = NULL
   END IF
   ' // Loads the image
   .GdipLoadImageFromFile(wszFileName, @m_pImage)
   ' // Erase the window's client area
   DIM hBgBrush AS HBRUSH = .CreateSolidBrush(m_BkColor)
   IF hBgBrush THEN
      DIM hDC AS HDC = .GetDc(m_hCtl)
      IF hDC THEN
         DIM rcFill AS RECT
         .GetClientRect m_hCtl, @rcFill
         .FillRect hDC, @rcFill, hBgBrush
         .DeleteDc hDC
      END IF
      .DeleteObject hBgBrush
   END IF
   ' // Redraws the control
   .InvalidateRect m_hCtl, NULL, 0
   .UpdateWindow m_hCtl
END FUNCTION
' ========================================================================================

' ========================================================================================
' UNTESTED
' Loads a bitmap from a resource file using GDI+
' Parameters:
' - hInstance = Instance handle
' - wszResourceName = Name of the resource.
' Return value:
'   Ok or an error code.
' Note: In Windows 7, it fails to load .jpg and .tif files. Works with .bmp and .png files.
' ========================================================================================
PRIVATE FUNCTION CImageCtx.LoadBitmapFromResource (BYVAL hInstance AS HINSTANCE, BYREF wszResourceName AS WSTRING) AS LONG
   IF LEN(wszResourceName) = 0 THEN RETURN E_INVALIDARG
   ' // Dispose a previous instance of the Image object, if any
   IF m_pImage THEN
      .GdipDisposeImage(m_pImage)
      m_pImage = NULL
   END IF
   DIM hStatus AS LONG = .GdipCreateBitmapFromResource(hInstance, wszResourceName, cast(GpBitmap PTR PTR, @m_pImage))
   ' // Erases the window's client area
   DIM hBgBrush AS HBRUSH = .CreateSolidBrush(m_BkColor)
   IF hBgBrush THEN
      DIM hDC AS HDC = .GetDc(m_hCtl)
      IF hDC THEN
         DIM rcFill AS RECT
         .GetClientRect m_hCtl, @rcFill
         .FillRect hDC, @rcFill, hBgBrush
         .DeleteDc hDC
      END IF
      .DeleteObject hBgBrush
   END IF
   ' // Redraws the control
   .InvalidateRect m_hCtl, NULL, 0
   .UpdateWindow m_hCtl
   FUNCTION = hStatus
END FUNCTION
' ========================================================================================

' ========================================================================================
' UNTESTED
' Loads an image from a resource file using GDI+
' Parameters:
' - hInstance = Instance handle
' - wszResourceName = Name of the resource image.
' Return value:
'   Ok or an error code.
' ========================================================================================
PRIVATE FUNCTION CImageCtx.LoadImageFromResource (BYVAL hInstance AS HINSTANCE, BYREF wszResourceName AS WSTRING) AS LONG
   IF LEN(wszResourceName) = 0 THEN RETURN E_INVALIDARG
   ' // Dispose a previous instance of the Image object, if any
   IF m_pImage THEN
      .GdipDisposeImage(m_pImage)
      m_pImage = NULL
   END IF
   DIM hStatus AS LONG, hDC AS HDC
   ' // Find the resource and lock it
   DIM hResource AS HRSRC = .FindResourceW(hInstance, wszResourceName, CAST(LPCWSTR, RT_RCDATA))
   IF hResource = NULL THEN RETURN E_INVALIDARG
   DIM imageSize AS DWORD = .SizeofResource(hInstance, hResource)
   IF imageSize = 0 THEN RETURN E_INVALIDARG
   DIM pResourceData AS LPVOID = .LockResource(LoadResource(hInstance, hResource))
   IF pResourceData = NULL THEN RETURN E_INVALIDARG
   ' // Allocate memory to hold the image
   DIM hGlobal AS HGLOBAL = .GlobalAlloc(GMEM_MOVEABLE, imageSize)
   IF hGlobal THEN
      ' // Lock the memory
      DIM pGlobalBuffer AS LPVOID = .GlobalLock(hGlobal)
      IF pGlobalBuffer THEN
         ' // Copy the image from the resource file to global memory
         .memcpy pGlobalBuffer, pResourceData, imageSize
         ' // Create an stream in global memory
         DIM pImageStream AS LPSTREAM, pGraphics AS GpGraphics PTR
         IF .CreateStreamOnHGlobal(hGlobal, FALSE, @pImageStream) = S_OK THEN
            ' // Create a bitmap from the data contained in the stream
            hStatus = .GdipCreateBitmapFromStream(pImageStream, cast(GpBitmap PTR PTR, @m_pImage))
            IF hStatus = 0 THEN
               ' // Gets the device context handle
               hDC = .GetDc(m_hCtl)
               ' // Creates a graphics object from it
               IF hDC THEN hStatus = .GdipCreateFromHDC(hDC, @pGraphics)
               ' // Draws the image (required to keep it in memory, since we are
               ' // going to unlock and free the resource)
               IF pGraphics THEN hStatus = .GdipDrawImageI(pGraphics, m_pImage, 0, 0)
               ' // Deletes the graphics object
               IF pGraphics THEN .GdipDeleteGraphics(pGraphics)
               ' // Releases the device context handle
               IF hDC THEN .DeleteDc hDC
            END IF
            pImageStream->lpVtbl->Release(pImageStream)
         END IF
         ' // Unlock the memory
         .GlobalUnlock pGlobalBuffer
      END IF
      ' // Free the memory
      .GlobalFree hGlobal
   END IF
   ' // Erases the window's client area
   DIM hBgBrush AS HBRUSH = .CreateSolidBrush(m_BkColor)
   IF hBgBrush THEN
      hDC = .GetDc(m_hCtl)
      IF hDC THEN
         DIM rcFill AS RECT
         .GetClientRect m_hCtl, @rcFill
         .FillRect hDC, @rcFill, hBgBrush
         .DeleteDc hDC
      END IF
      .DeleteObject hBgBrush
   END IF
   ' // Redraws the control
   .InvalidateRect m_hCtl, NULL, 0
   .UpdateWindow m_hCtl
   FUNCTION = hStatus
END FUNCTION
' ========================================================================================

' ========================================================================================
' Redraws the control.
' ========================================================================================
PRIVATE SUB CImageCtx.Redraw
   .InvalidateRect m_hCtl, NULL, 0
   .UpdateWindow m_hCtl
END SUB
' ========================================================================================

' ========================================================================================
' Sets the image adjustment.
' Parameters:
' - ImageAdjustment =
'     GDIP_IMAGECTX_AUTOSIZE = 0
'        Autoadjusts the image to the width or height of the control
'     GDIP_IMAGECTX_ACTUALSIZE = 1
'        Shows the image with its actual size
'     GDIP_IMAGECTX_FITTOWIDTH = 2
'        Adjusts the image to the width of the control
'     GDIP_IMAGECTX_FITTOHEIGHT = 3
'        Adjusts the image to the height of the control
' - fRedraw = [optional] TRUE or FALSE (redraws the control to reflect the changes)
' Return value:
' * Returns the previous setting.
' ========================================================================================
PRIVATE FUNCTION CImageCtx.SetImageAdjustment (BYVAL ImageAdjustment AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE) AS LONG
   ' // Returns the old setting
   FUNCTION = m_Fit
   ' // Store the new setting
   m_Fit = ImageAdjustment
   ' // Redraws the control
   IF fRedraw THEN
      .InvalidateRect m_hCtl, NULL, 0
      .UpdateWindow m_hCtl
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the image adjustment setting.
' Parameter:
' * hwnd = Control's window handle
' Return value:
' * The current adjustment setting.
' ========================================================================================
PRIVATE FUNCTION CImageCtx.GetImageAdjustment () AS LONG
   FUNCTION = m_Fit
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the width of the image.
' ========================================================================================
PRIVATE FUNCTION CImageCtx.GetImageWidth () AS LONG
   DIM ImageWidth AS DWORD
   IF m_pImage THEN
      .GdipGetImageWidth(m_pImage, @ImageWidth)
      FUNCTION = ImageWidth
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the height of the image.
' Parameter:
' * hwnd = Control's window handle
' Return value:
' * The height of the image.
' ========================================================================================
PRIVATE FUNCTION CImageCtx.GetImageHeight () AS LONG
   DIM ImageHeight AS DWORD
   IF m_pImage THEN
      .GdipGetImageHeight(m_pImage, @ImageHeight)
      FUNCTION = ImageHeight
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the background color of the control.
' Parameters:
' * clr = RGB color.
' * fRedraw = [optional] TRUE or FALSE (redraws the control to reflect the changes)
' Return value:
' * Returns the previous background color.
' ========================================================================================
PRIVATE FUNCTION CImageCtx.SetBkColor (BYVAL clr AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE) AS LONG
   ' // Returns the old setting
   FUNCTION = m_BkColor
   ' // Stores the new color
   m_BkColor = clr
   ' // Redraws the control
   IF fRedraw THEN
      .InvalidateRect m_hCtl, NULL, 0
      .UpdateWindow m_hCtl
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the background color of the control's window.
' ========================================================================================
PRIVATE FUNCTION CImageCtx.GetBkColor () AS LONG
   FUNCTION = m_BkColor
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the interpolation mode used by GDI+. The interpolation mode determines the algorithm
' that is used when images are scaled or rotated.
' ========================================================================================
PRIVATE SUB CImageCtx.SetInterpolationMode (BYVAL InterpolationMode AS LONG)
   m_InterpolationMode = InterpolationMode
END SUB
' ========================================================================================

' ========================================================================================
' Gets the interpolation mode used by GDI+.
' ========================================================================================
PRIVATE FUNCTION CImageCtx.GetInterpolationMode () AS LONG
   FUNCTION = m_InterpolationMode
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the Image object pointer.
' Return value: The Image object pointer.
' NOTE: Don't dispose the returned pointer. The control keeps a copy of it and calls
' GdipDisposeImage when it is destroyed.
' ========================================================================================
PRIVATE FUNCTION CImageCtx.GetImagePtr () AS GpImage PTR
   FUNCTION = m_pImage
END FUNCTION
' ========================================================================================

' ========================================================================================
' Clears the contents of the control
' Return value:
' * If the function succeeds, the return value is TRUE.
'   If the function fails, the return value is FALSE.
' ========================================================================================
PRIVATE SUB CImageCtx.Clear
   IF m_pImage THEN
      .GdipDisposeImage(m_pImage)
      m_pImage = 0
   END IF
   ' // Erase the contents of the window
   DIM rcFill AS RECT
   DIM hBgBrush AS HBRUSH = .CreateSolidBrush(m_BkColor)
   .GetClientRect m_hCtl, @rcFill
   DIM hDC AS HDC = GetDc(m_hCtl)
   .FillRect hDC, @rcFill, hBgBrush
   .ReleaseDC m_hCtl, hDC
   .DeleteObject hBgBrush
END SUB
' ========================================================================================

' ========================================================================================
' Returns a pointer to the class given the handle of its associated window handle.
' hCtl = Handle of the image control.
' ========================================================================================
PRIVATE FUNCTION AfxCImageCtxPtr OVERLOAD (BYVAL hCtl AS HWND) AS CImageCtx PTR
   FUNCTION = CAST(CImageCtx PTR, .GetWindowLongPtrW(hCtl, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
' hParent = Handle of the parent window of the image control.
' cID = Identifier of the image control.
' ========================================================================================
PRIVATE FUNCTION AfxCImageCtxPtr OVERLOAD (BYVAL hParent AS HWND, BYVAL cID AS LONG) AS CImageCtx PTR
   FUNCTION = CAST(CImageCtx PTR, .GetWindowLongPtrW(GetDlgItem(hParent, cID), 0))
END FUNCTION
' ========================================================================================

END NAMESPACE
