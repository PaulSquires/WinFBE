' ########################################################################################
' Microsoft Windows
' File: MatrixTransformPoints.bas
' Contents: GDI+ - MatrixTransformPoints example
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
' The following example creates an array of five points and then transforms those points
' by calling the Matrix.TransformPoints method of a Matrix object. The code passes the
' array of points to the DrawCurve Methods method before the transformation and again
' after the transformation.
' ========================================================================================
SUB Example_TransformPoints (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratios
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96

   ' // Create a pen
   DIM myPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), rxRatio)

   DIM rgPoints(0 TO 4) AS GpPointF = {GDIP_POINTF(50 * rxRatio, 100 * ryRatio), GDIP_POINTF(100 * rxRatio, 50 * ryRatio), _
       GDIP_POINTF(160 * rxRatio, 125 * rxRatio), GDIP_POINTF(200 * rxRatio, 100 * ryRatio), GDIP_POINTF(250 * rxRatio, 150 * ryRatio)}

'#ifdef __FB_64BIT__
'   DIM rgPoints(0 TO 4) AS GpPointF = {(50 * rxRatio, 100 * ryRatio), (100 * rxRatio, 50 * ryRatio), _
'       (160 * rxRatio, 125 * rxRatio), (200 * rxRatio, 100 * ryRatio), (250 * rxRatio, 150 * ryRatio)}
'#else
'      ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM rgPoints(0 TO 4) AS GpPointF
'   rgPoints(0).x = 50 * rxRatio : rgPoints(0).y = 100 * ryRatio : rgPoints(1).x = 100 * rxRatio
'   rgPoints(1). y = 50 * ryRatio : rgPoints(2).x = 160 * rxRatio : rgPoints(2).y = 125 * ryRatio
'   rgPoints(3).x = 200 * rxRatio : rgPoints(3).y = 100 * ryRatio rgPoints(4).x = 250 * rxRatio
'   rgPoints(4).y = 150 * ryRatio
'#endif

   DIM matrix AS CGpMatrix = CGpMatrix(1.0, 0.0, 0.0, 2.0, 0.0, 0.0)

   graphics.DrawCurve(@myPen, @rgPoints(0), 5)
   matrix.TransformPoints(@rgPoints(0), 5)
   graphics.DrawCurve(@myPen, @rgPoints(0), 5)

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
   pWindow.Create(NULL, "GDI+ MatrixTransformPoints", @WndProc)
   ' // Change the window style
   pWindow.WindowStyle = WS_OVERLAPPED OR WS_CAPTION OR WS_SYSMENU
   ' // Size it by setting the wanted width and height of its client area
   pWindow.SetClientSize(400, 350)
   ' // Center the window
   pWindow.Center

   ' // Add a graphic control
   DIM pGraphCtx AS CGraphCtx = CGraphCtx(@pWindow, IDC_GRCTX, "", 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   pGraphCtx.Clear BGR(255, 255, 255)
   ' // Get the memory device context of the graphic control
   DIM hdc AS HDC = pGraphCtx.GetMemDc
   ' // Draw the graphics
   Example_TransformPoints(hdc)

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
