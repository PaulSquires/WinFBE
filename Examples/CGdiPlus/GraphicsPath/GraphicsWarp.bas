' ########################################################################################
' Microsoft Windows
' File: GraphicsPathWarp.bas
' Contents: GDI+ - GraphicsPathWarp example
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
' The following example creates a GraphicsPath object and adds a closed figure to the path.
' The code defines a warp transformation by specifying a source rectangle and an array of
' four destination points. The source rectangle and destination points are passed to the
' Warp method. The code draws the path twice: once before it has been warped and once after
' it has been warped.
' ========================================================================================
SUB Example_Warp (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratios
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a path.
   DIM points(0 TO 7) AS GpPointF = {GDIP_POINTF(20.0, 60.0), GDIP_POINTF(30.0, 90.0), _
      GDIP_POINTF(15.0, 110.0), GDIP_POINTF(15.0, 145.0), GDIP_POINTF(55.0, 145.0), _
      GDIP_POINTF(55.0, 110.0), GDIP_POINTF(40.0, 90.0), GDIP_POINTF(50.0, 60.0)}
   DIM path AS CGpGraphicsPath
   path.AddLines(@points(0), 8)
   path.CloseFigure

   ' // Draw the path before applying a warp transformation.
   DIM bluePen AS CGpPen = GDIP_ARGB(255, 0, 0, 255)
   graphics.DrawPath(@bluePen, @path)

   ' // Define a warp transformation, and warp the path.
   DIM srcRect AS GpRectF = GDIP_RECTF(10.0, 50.0, 50.0, 100.0)
   DIM destPts(0 TO 3) AS GpPointF = {GDIP_POINTF(220.0, 10.0), GDIP_POINTF(280.0, 10.0), _
       GDIP_POINTF(100.0, 150.0), GDIP_POINTF(400.0, 150.0)}
   path.Warp(@destPts(0), 4, @srcRect)

   ' // Draw the warped path.
   graphics.DrawPath(@bluePen, @path)

   ' // Draw the source rectangle and the destination polygon.
   DIM blackPen AS CGpPen = GDIP_ARGB(255, 0, 0, 0)
   graphics.DrawRectangle(@blackPen, @srcRect)
   graphics.DrawLine(@blackPen, @destPts(0), @destPts(1))
   graphics.DrawLine(@blackPen, @destPts(0), @destPts(2))
   graphics.DrawLine(@blackPen, @destPts(1), @destPts(3))
   graphics.DrawLine(@blackPen, @destPts(2), @destPts(3))

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
   pWindow.Create(NULL, "GDI+ GraphicsPathWarp", @WndProc)
   ' // Change the window style
   pWindow.WindowStyle = WS_OVERLAPPED OR WS_CAPTION OR WS_SYSMENU
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(410, 250)
   ' // Centers the window
   pWindow.Center

   ' // Add a graphic control
   DIM pGraphCtx AS CGraphCtx = CGraphCtx(@pWindow, IDC_GRCTX, "", 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   pGraphCtx.Clear BGR(255, 255, 255)
   ' // Get the memory device context of the graphic control
   DIM hdc AS HDC = pGraphCtx.GetMemDc
   ' // Draw the graphics
   Example_Warp(hdc)

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
