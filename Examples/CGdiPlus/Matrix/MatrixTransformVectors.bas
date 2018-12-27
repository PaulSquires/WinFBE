' ########################################################################################
' Microsoft Windows
' File: MatrixTransformVectors.bas
' Contents: GDI+ - MatrixTransformVectors example
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
' The following example creates a vector and a point. The tip of the vector and the point
' are at the same location: (100, 50). The code creates a Matrix object and initializes its
' elements so that it represents a clockwise rotation followed by a translation 100 units
' to the right. The code calls the TransformPoints Methods method of the matrix to transform
' the point and calls the Matrix.TransformVectors method of the matrix to transform the
' vector. The entire transformation (rotation followed by translation) is performed on the
' point, but only the rotation part of the transformation is performed on the vector. The
' elements of the matrix that represent translation are ignored by the Matrix.TransformVectors method.
' ========================================================================================
SUB Example_TransformVectors (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratios
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a pen
   DIM myPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), rxRatio)
   myPen.SetEndCap(LineCapArrowAnchor)
   DIM myBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 255)

   ' // A point and a vector, same representation but different behavior
   DIM pt AS GpPointF = GDIP_POINTF(100.0, 50.0)
   DIM vector AS GpPointF = GDIP_POINTF(100.0, 50.0)

   ' // Draw the original point and vector in blue
   graphics.FillEllipse(@myBrush, pt.X - 5.0, pt.Y - 5.0, 10.0, 10.0)
   graphics.DrawLine(@MyPen, 0, 0, vector.x, vector.y)

   ' // Transform
   DIM matrix AS CGpMatrix = CGpMatrix(0.8, 0.6, -0.6, 0.8, 100.0, 0.0)
   matrix.TransformPoints(@pt)
   matrix.TransformVectors(@vector)

   ' // Draw the transformed point and vector in red.
   myPen.SetColor(GDIP_ARGB(255, 255, 0, 0))
   myBrush.SetColor(GDIP_ARGB(255, 255, 0, 0))
   graphics.FillEllipse(@myBrush, pt.X - 5.0, pt.Y - 5.0, 10.0, 10.0)
   graphics.DrawLine(@MyPen, 0, 0, vector.x, vector.y)

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
   pWindow.Create(NULL, "GDI+ MatrixTransformVectors", @WndProc)
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
   Example_TransformVectors(hdc)

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
