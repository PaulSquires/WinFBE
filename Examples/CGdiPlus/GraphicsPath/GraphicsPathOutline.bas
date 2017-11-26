' ########################################################################################
' Microsoft Windows
' File: GraphicsPathOutline.bas
' Contents: GDI+ - GraphicsPathOutline example
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
' The following example creates a GraphicsPath object and calls the GraphicsPath.AddClosedCurve
' method to add a closed cardinal spline to the path. The code calls the GraphicsPath.Widen
' method to widen the path and then draws the path. Next, the code calls the path's Outline
' method. The code calls the TranslateTransform method of a Graphics object so that the
' outlined path drawn by the subsequent call to DrawPath sits to the right of the first path.
' ========================================================================================
SUB Example_Outline (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratios
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM bluePen AS CGpPen = GDIP_ARGB(255, 0, 0, 255)
   DIM greenPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 255, 0), 10)

   DIM points(0 TO 3) AS GpPoint = {GDIP_POINT(20, 20), GDIP_POINT(160, 100), _
       GDIP_POINT(140, 60), GDIP_POINT(60, 100)}

   DIM path AS CGpGraphicsPath
   path.AddClosedCurve(@points(0), 4)

   path.Widen(@greenPen)
   graphics.DrawPath(@bluePen, @path)

   path.Outline

   graphics.TranslateTransform(180, 0)
   graphics.DrawPath(@bluePen, @path)

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
   pWindow.Create(NULL, "GDI+ GraphicsPathOutline", @WndProc)
   ' // Change the window style
   pWindow.WindowStyle = WS_OVERLAPPED OR WS_CAPTION OR WS_SYSMENU
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(400, 250)
   ' // Centers the window
   pWindow.Center

   ' // Add a graphic control
   DIM pGraphCtx AS CGraphCtx = CGraphCtx(@pWindow, IDC_GRCTX, "", 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   pGraphCtx.Clear BGR(255, 255, 255)
   ' // Get the memory device context of the graphic control
   DIM hdc AS HDC = pGraphCtx.GetMemDc
   ' // Draw the graphics
   Example_Outline(hdc)

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
