' ########################################################################################
' Microsoft Windows
' File: PathGradientBrushGetFocusScales.bas
' Contents: GDI+ - PathGradientBrushGetFocusScales example
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
' The following example creates a PathGradientBrush object based on a triangular path.
' The code sets the focus scales of the path gradient brush to (0.2, 0.2) and then uses
' the path gradient brush to fill an area that contains the triangular path. Finally, the
' code calls the PathGradientBrush.GetFocusScales method of the PathGradientBrush object
' to obtain the values of the x focus scale and the y focus scale.
' ========================================================================================
SUB Example_GetFocusScales (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratios
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM points(0 TO 2) AS GpPoint = {GDIP_POINT(100, 0), GDIP_POINT(200, 200), GDIP_POINT(0, 200)}

   ' // No GraphicsPath object is created. The PathGradientBrush
   ' // object is constructed directly from the array of points.
   DIM pthGrBrush AS CGpPathGradientBrush = CGpPathGradientBrush(@points(0), 3)

   DIM colors(0 TO 1) AS ARGB = {GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255)}

   ' // red at the boundary of the outer triangle
   ' // blue at the boundary of the inner triangle
   DIM relativePositions(0 TO 1) AS SINGLE = {0.0, 1.0}
   pthGrBrush.SetInterpolationColors(@colors(0), @relativePositions(0), 2)

   ' // The inner triangle is formed by scaling the outer triangle
   ' // about its centroid. The scaling factor is 0.2 in both the x and y directions.
   pthGrBrush.SetFocusScales(0.2, 0.2)

   ' // Fill a rectangle that is larger than the triangle
   ' // specified in the Point array. The portion of the
   ' // rectangle outside the triangle will not be painted.
   graphics.FillRectangle(@pthGrBrush, 0, 0, 200, 200)

   ' // Obtain information about the path gradient brush.
   DIM xScale AS SINGLE = 0.0
   DIM yScale AS SINGLE = 0.0
   pthGrBrush.GetFocusScales(@xScale, @yScale)

   ' // The value of xScale is now 0.2.
   ' // The value of yScale is now 0.2.

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
   pWindow.Create(NULL, "GDI+ PathGradientBrushGetFocusScales", @WndProc)
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
   Example_GetFocusScales(hdc)

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
