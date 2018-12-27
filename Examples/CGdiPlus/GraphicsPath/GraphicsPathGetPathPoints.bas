' ########################################################################################
' Microsoft Windows
' File: GraphicsPathGetPathPoints.bas
' Contents: GDI+ - GraphicsPathGetPathPoints example
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

' // Forward declaratiom
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

' ========================================================================================
' The following example creates and draws a path that has a line, a rectangle, an ellipse,
' and a curve. The code calls the path's GraphicsPath::GetPointCount method to determine
' the number of data points that are stored in the path. The code allocates a buffer large
' enough to receive the array of data points and passes the address of that buffer to the
' GraphicsPath.GetPathPoints method. Finally, the code draws each of the path's data points.
' ========================================================================================
SUB Example_GetPathPoints (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratios
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a path that has a line, a rectangle, an ellipse, and a curve.
   DIM path AS CGpGraphicsPath
   DIM points(0 TO 4) AS GpPoint = {GDIP_POINT(200, 200), GDIP_POINT(250, 240), _
      GDIP_POINT(200, 300), GDIP_POINT(300, 310), GDIP_POINT(250, 350)}

   path.AddLine(20, 100, 150, 200)
   path.AddRectangle(40, 30, 80, 60)
   path.AddEllipse(200, 30, 200, 100)
   path.AddCurve(@points(0), 5)

   ' // Draw the path
   DIM pen AS CGpPen = GDIP_ARGB(255, 0, 0, 255)
   graphics.DrawPath(@pen, @path)

   ' // Get the path points.
   DIM nCount AS LONG = path.GetPointCount
   DIM dataPoints(nCount -1) AS GpPointF
   path.GetPathPoints(@dataPoints(0), nCount)

   ' // Draw the path's data points
   DIM brush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   FOR j AS LONG = 0 TO nCount - 1
      graphics.FillEllipse(@brush, dataPoints(j).x - 3.0, _
         dataPoints(j).y - 3.0, 6.0, 6.0)
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

   ' // Creates the main window
   DIM pWindow AS CWindow
   ' -or- DIM pWindow AS CWindow = "MyClassName" (use the name that you wish)
   pWindow.Create(NULL, "GDI+ GraphicsPathGetPathPoints", @WndProc)
   ' // Change the window style
   pWindow.WindowStyle = WS_OVERLAPPED OR WS_CAPTION OR WS_SYSMENU
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(420, 380)
   ' // Centers the window
   pWindow.Center

   ' // Add a graphic control
   DIM pGraphCtx AS CGraphCtx = CGraphCtx(@pWindow, IDC_GRCTX, "", 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   pGraphCtx.Clear BGR(255, 255, 255)
   ' // Get the memory device context of the graphic control
   DIM hdc AS HDC = pGraphCtx.GetMemDc
   ' // Draw the graphics
   Example_GetPathPoints(hdc)

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
