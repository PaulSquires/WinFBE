' ########################################################################################
' Microsoft Windows
' File: Restore.bas
' Contents: GDI+ - Restore example
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
' Restoring Nested Saved States
' The following example sets the world transformation of a Graphics object to a rotation
' and then saves the state of the Graphics object. Next, the code calls TranslateTransform,
' and saves the state again. Then the code calls ScaleTransform. At that point, the world
' transformation of the Graphics object is a composite transformation: first rotate, then
' translate, then scale. The code uses a red pen to draw an ellipse that is transformed by
' that composite transformation.
' The code passes state2, which was returned by the second call to Save, to the Graphics.Restore
' method, and draws the ellipse again using a green pen. The green ellipse is rotated and
' translated but not scaled. Finally the code passes state1, which was returned by the
' first call to Save, to the Graphics.Restore method, and draws the ellipse again using a
' blue pen. The blue ellipse is rotated but not translated or scaled.
' ========================================================================================
SUB Example_Restore (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratios
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Three transformations apply: rotate, then translate, then scale.
   DIM AS GraphicsState state1, state2
   graphics.RotateTransform(30)
   state1 = graphics.Save
   graphics.TranslateTransform(100 * rxRatio, 0, MatrixOrderAppend)
   state2 = graphics.Save
   graphics.ScaleTransform(1, 3, MatrixOrderAppend)

   ' // Draw an ellipse.
   DIM redPen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawEllipse(@redPen, 0, 0, 100, 20)

   ' // Restore to state2 and draw the ellipse again.
   ' // Two transformations apply: rotate then translate.
   graphics.Restore(state2)
   DIM greenPen AS CGpPen = GDIP_ARGB(255, 0, 255, 0)
   graphics.DrawEllipse(@greenPen, 0, 0, 100, 20)

   ' // Restore to state1 and draw the ellipse again.
   ' // Only the rotation transformation applies.
   graphics.Restore(state1)
   DIM bluePen AS CGpPen = GDIP_ARGB(255, 0, 0, 255)
   graphics.DrawEllipse(@bluePen, 0, 0, 100, 20)

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
   pWindow.Create(NULL, "GDI+ Restore", @WndProc)
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
   Example_Restore(hdc)

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
