' ########################################################################################
' Microsoft Windows
' File: DrawBeziers.bas
' Contents: GDI+ - DrawBeziers example
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
' The following example draws a pair of Bézier curves.
' ========================================================================================
SUB Example_DrawBeziers (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratios
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Define a Pen object and an array of PointF objects.
   DIM greenPen AS CGpPen = GDIP_ARGB(255, 0, 255, 0)
   DIM startPoint AS GpPointF : startPoint.x = 100.0 : startPoint.y = 100.0
   DIM ctrlPoint1 AS GpPointF : ctrlPoint1.x = 200.0 : ctrlPoint1.y = 50.0
   DIM ctrlPoint2 AS GpPointF : ctrlPoint2.x = 400.0 : ctrlPoint2.y = 10.0
   DIM endPoint1  AS GpPointF : endPoint1.x  = 500.0 : endPoint1.y  = 100.0
   DIM ctrlPoint3 AS GpPointF : ctrlPoint3.x = 600.0 : ctrlPoint3.y = 200.0
   DIM ctrlPoint4 AS GpPointF : ctrlPoint4.x = 700.0 : ctrlPoint4.y = 400.0
   DIM endPoint2  AS GpPointF : endPoint2.x  = 500.0 : endPoint2.y  = 500.0

   DIM curvePoints(6) AS GpPointF
   curvePoints(0) = startPoint
   curvePoints(1) = ctrlPoint1
   curvePoints(2) = ctrlPoint2
   curvePoints(3) = endPoint1
   curvePoints(4) = ctrlPoint3
   curvePoints(5) = ctrlPoint4
   curvePoints(6) = endPoint2

   ' // Draw the Bezier curves.
   graphics.DrawBeziers(@greenPen, @curvePoints(0), 7)

   ' // Draw the control and end points.
   DIM redBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   graphics.FillEllipse(@redBrush, 100 - 5, 100 - 5, 10, 10)
   graphics.FillEllipse(@redBrush, 500 - 5, 100 - 5, 10, 10)
   graphics.FillEllipse(@redBrush, 500 - 5, 500 - 5, 10, 10)
   DIM blueBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 255)
   graphics.FillEllipse(@blueBrush, 200 - 5, 50 - 5, 10, 10)
   graphics.FillEllipse(@blueBrush, 400 - 5, 10 - 5, 10, 10)
   graphics.FillEllipse(@blueBrush, 600 - 5, 200 - 5, 10, 10)
   graphics.FillEllipse(@blueBrush, 700 - 5, 400 - 5, 10, 10)

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
   pWindow.Create(NULL, "GDI+ DrawBeziers", @WndProc)
   ' // Change the window style
   pWindow.WindowStyle = WS_OVERLAPPED OR WS_CAPTION OR WS_SYSMENU
   ' // Size it by setting the wanted width and height of its client area
   pWindow.SetClientSize(800, 530)
   ' // Center the window
   pWindow.Center

   ' // Add a graphic control
   DIM pGraphCtx AS CGraphCtx = CGraphCtx(@pWindow, IDC_GRCTX, "", 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   pGraphCtx.Clear BGR(255, 255, 255)
   ' // Get the memory device context of the graphic control
   DIM hdc AS HDC = pGraphCtx.GetMemDc
   ' // Draw the graphics
   Example_DrawBeziers(hdc)

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
