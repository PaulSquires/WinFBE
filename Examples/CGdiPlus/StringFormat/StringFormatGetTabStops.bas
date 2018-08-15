' ########################################################################################
' Microsoft Windows
' File: StringFormatGetTabStops.bas
' Contents: GDI+ - StringFormatGetTabStops example
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
' The following example creates a StringFormat object, sets tab stops, and uses the
' StringFormat object to draw a string that contains tab characters. The code also draws
' the string's layout rectangle. Then, the code gets the tab stops and proceeds to use or
' inspect the tab stops in some way.
' ========================================================================================
SUB Example_GetTabStops (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratios
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   ' // Create a red solid brush
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 255)
   ' // Create a font family from name
   DIM fontFamily AS CGpFontFamily = "Times New Roman"
   ' // Create a font from the font family
   DIM pFont AS CGpFont = CGpFont(@fontFamily, 16, FontStyleRegular, UnitPixel)

   ' // Create a string format object and set the tab stops
   DIM tabs(0 TO 2) AS SINGLE = {150, 100, 100}
   DIM stringFormat AS CGpStringFormat
   stringFormat.SetTabStops(0, 3, @tabs(0))

   ' // Use the StringFormat object in a call to DrawString
   DIM wszText AS WSTRING * 260 = "Name" & CHR(9) &"Test 1" & CHR(9) & "Test 2" & CHR(9) & "Test 3"
   graphics.DrawString(@wszText, LEN(wszText), @pFont, 20, 20, 500, 100, @stringFormat, @solidBrush)

   ' // Draw the rectangle that encloses the text.
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 255, 0, 0))
   graphics.DrawRectangle(@pen, 20, 20, 500, 100)

   ' // Get the tab stops
   DIM tabStopCount AS LONG, firstTabOffset AS SINGLE
   DIM tabStops(ANY) AS SINGLE
   tabStopCount = stringFormat.GetTabStopCount
   REDIM tabStops(tabStopCount - 1)
   stringFormat.GetTabStops(tabStopCount, @firstTabOffset, @tabStops(0))

   FOR j AS LONG = 0 TO tabStopCount - 1
      ' // Inspect or use the value in tabStops(j)
      PRINT tabStops(j)
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
   pWindow.Create(NULL, "GDI+ StringFormatGetTabStops", @WndProc)
   ' // Change the window style
   pWindow.WindowStyle = WS_OVERLAPPED OR WS_CAPTION OR WS_SYSMENU
   ' // Size it by setting the wanted width and height of its client area
   pWindow.SetClientSize(540, 200)
   ' // Center the window
   pWindow.Center

   ' // Add a graphic control
   DIM pGraphCtx AS CGraphCtx = CGraphCtx(@pWindow, IDC_GRCTX, "", 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   pGraphCtx.Clear BGR(255, 255, 255)
   ' // Get the memory device context of the graphic control
   DIM hdc AS HDC = pGraphCtx.GetMemDc
   ' // Draw the graphics
   Example_GetTabStops(hdc)

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
