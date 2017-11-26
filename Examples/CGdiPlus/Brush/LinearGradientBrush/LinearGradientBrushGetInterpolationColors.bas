' ########################################################################################
' Microsoft Windows
' File: LinearGradientBrushGetInterpolationColors.bas
' Contents: GDI+ - LinearGradientBrushGetInterpolationColors example
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "Afx/CWindow.inc"
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
' The following example sets the colors that are interpolated for this linear gradient
' brush to red, blue, and green and sets the blend positions to 0, 0.3, and 1. The code
' calls the LinearGradientBrush::GetInterpolationColorCount method of a LinearGradientBrush
' object to obtain the number of colors currently set to be interpolated for the brush.
' Next, the code gets the colors and their positions. Then, the code fills a small
' rectangle with each color.
' ========================================================================================
SUB Example_GetInterpolationColors (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratios
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a linear gradient brush, and set the colors to be interpolated.
   DIM colors(0 TO 2) AS ARGB = {GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255), GDIP_ARGB(255, 0, 255, 0)}
   DIM positions(0 TO 2) AS SINGLE = {0.0, 0.3, 1.0}

   DIM pt1 AS GpPoint = GDIP_POINT(0, 0)
   DIM pt2 AS GpPoint = GDIP_POINT(100, 0)

   DIM linGrBrush AS CGpLinearGradientBrush = CGpLinearGradientBrush(@pt1, @pt2, GDIP_ARGB(255, 0, 0, 0), GDIP_ARGB(255, 255, 255, 255))
   linGrBrush.SetInterpolationColors(@colors(0), @positions(0), 3)

   ' // Obtain information about the linear gradient brush.
   ' // How many colors have been specified to be interpolated for this brush?
   DIM colorCount AS LONG = linGrBrush.GetInterpolationColorCount

   ' // Allocate a buffer large enough to hold the set of colors.
   DIM rgcolors(0 TO colorCount - 1) AS ARGB

   ' // Allocate a buffer to hold the relative positions of the colors.
   DIM rgPositions(0 TO colorCount - 1) AS SINGLE

   ' // Get the colors and their relative positions.
   linGrBrush.GetInterpolationColors(@rgcolors(0), @rgPositions(0), colorCount)

   ' // Fill a small rectangle with each of the colors.
   DIM pSolidBrush AS CGpSolidBrush PTR
   FOR j AS LONG = 0 TO colorCount - 1
      pSolidBrush = NEW CGpSolidBrush(rgcolors(j))
      graphics.FillRectangle(pSolidBrush, 15 * j, 0, 10, 10)
      Delete pSolidBrush
   NEXT

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
   pWindow.Create(NULL, "GDI+ LinearGradientBrushGetInterpolationColors", @WndProc)
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
   Example_GetInterpolationColors(hdc)

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
