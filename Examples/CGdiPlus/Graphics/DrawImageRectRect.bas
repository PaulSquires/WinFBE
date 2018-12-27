' ########################################################################################
' Microsoft Windows
' File: DrawImageRectRect.bas
' Contents: GDI+ - DrawImageRectRect example
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "Afx/CGdiPlus/CGdiPlus.inc"
#INCLUDE ONCE "Afx/CGraphCtx.inc"
USING Afx

CONST IDC_GRCTX = 1001

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG

   END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

' ========================================================================================
' The following example draws the original source image and then draws a portion of the
' image in a specified parallelogram.
' ========================================================================================
SUB Example_DrawImageRectRect (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratios
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set scaling
   graphics.SetPageUnit(UnitPixel)
   graphics.SetPageScale(rxRatio)

   ' // Create an Image object.
   DIM pImage AS CGpImage = "pattern.png"

   ' // Draw the original source image.
   graphics.DrawImage(@pImage, 10, 10)

   ' // Define the portion of the image to draw.
   DIM srcX AS SINGLE = 70.0
   DIM srcY AS SINGLE = 20.0
   DIM srcWidth AS SINGLE = 100.0
   DIM srcHeight AS SINGLE = 100.0

   ' // Create an array of Point objects that specify the destination of the cropped image.
   DIM destPoints(0 TO 2) AS GpPointF
   destPoints(0).x = 230 : destPoints(0).y = 30
   destPoints(1).x = 350 : destPoints(1).y = 50
   destPoints(2).x = 275 : destPoints(2).y = 120

   ' Yet another mess of the FB GdiPlus declares.
'#ifdef __FB_64BIT__
'   DIM redToBlue AS ColorMap_
'   redToBlue.oldColor.value = GDIP_ARGB(255, 255, 0, 0)
'   redToBlue.newColor.value = GDIP_ARGB(255, 0, 0, 255)
'#else
'   DIM redToBlue AS ColorMap
'   redToBlue.from = GDIP_ARGB(255, 255, 0, 0)
'   redToBlue.to = GDIP_ARGB(255, 0, 0, 255)
'#endif

   ' // GDIP_COLORMAP is an union that solves the 32/64-bit incompatibility
   DIM redToBlue AS GDIP_COLORMAP = (GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255))

   ' // Create an ImageAttributes object that specifies a recoloring from red to blue.
   DIM remapAttributes AS CGpImageAttributes
   RemapAttributes.SetRemapTable(1, @redToBlue)

   ' // Draw the cropped image
   graphics.DrawImage(@pImage, @destPoints(0), 3, srcX, srcY, srcWidth, srcHeight, _
                     UnitPixel, @remapAttributes, NULL, NULL)


END SUB
' ========================================================================================

' ========================================================================================
' Main
' ========================================================================================
FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                  BYVAL hPrevInstance AS HINSTANCE, _
                  BYVAL szCmdLine AS ZSTRING PTR, _
                  BYVAL nCmdShow AS LONG) AS LONG

   ' // Set process DPI aware
   ' // The recommended way is to use a manifest file
   AfxSetProcessDPIAware

   ' // Create the main window
   DIM pWindow AS CWindow
   ' -or- DIM pWindow AS CWindow = "MyClassName" (use the name that you wish)
   pWindow.Create(NULL, "GDI+ DrawImageRectRect", @WndProc)
   ' // Change the window style
   pWindow.WindowStyle = WS_OVERLAPPED OR WS_CAPTION OR WS_SYSMENU
   ' // Size it by setting the wanted width and height of its client area
   pWindow.SetClientSize(400, 250)
   ' // Center the window
   pWindow.Center

   ' // Add a graphic control
   DIM pGraphCtx AS CGraphCtx = CGraphCtx(@pWindow, IDC_GRCTX, "", 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   pGraphCtx.Clear BGR(255, 255, 255)
   ' // Get the memory device context of the graphic control
   DIM hdc AS HDC = pGraphCtx.GetMemDc
   ' // Draw the graphics
   Example_DrawImageRectRect(hdc)

   ' // Displays the window and dispatches the Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Main window procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   SELECT CASE uMsg

      CASE WM_COMMAND
         SELECT CASE GET_WM_COMMAND_ID(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF
         END SELECT

    	CASE WM_DESTROY
         ' // Ends the application by sending a WM_QUIT message
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
