' ########################################################################################
' Microsoft Windows
' File: TaskDialog.bas
' Contents: Task dialog example.
' Remarks: Requires the use of a manifest.
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define _WIN32_WINNT &h0602
#INCLUDE ONCE "Afx/CWindow.inc"
USING Afx

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG

   END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hWnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

' ========================================================================================
' Main
' ========================================================================================
FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                  BYVAL hPrevInstance AS HINSTANCE, _
                  BYVAL szCmdLine AS ZSTRING PTR, _
                  BYVAL nCmdShow AS LONG) AS LONG

   ' // Set process DPI aware
   ' // The recommended way is to use a manifest
'   AfxSetProcessDPIAware

   ' // Create the main window
   DIM pWindow AS CWindow
   pWindow.Create(NULL, "Task Dialog", @WndProc)
   pWindow.SetClientSize(500, 320)
   pWindow.Center

   ' // Add the buttons
   pWindow.AddControl("Button", , IDOK, "&Click me", 280, 270, 75, 23)
   pWindow.AddControl("Button", , IDCANCEL, "&Exit", 380, 270, 75, 23)

   ' // Process Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Window procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hWnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   SELECT CASE uMsg

      CASE WM_COMMAND
         SELECT CASE GET_WM_COMMAND_ID(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application sending an WM_CLOSE message
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF
            CASE IDOK
               ' // Display the message
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  DIM nClickedButton AS LONG
                  DIM hr AS HRESULT = TaskDialog(hwnd, NULL, "CWindow", "CWindow", _
                        "An update for the CWindow framework has just been released. Do you want to download it?", _
                        TDCBF_YES_BUTTON OR TDCBF_NO_BUTTON, _
                        TD_INFORMATION_ICON, @nClickedButton)
                  IF hr = S_OK THEN
                     SELECT CASE nClickedButton
                        CASE IDYES : MessageBox 0, "You clicked the Yes button", "", MB_OK
                        CASE IDNO  : MessageBox 0, "You clicked the No button", "", MB_OK
                     END SELECT
                  END IF
               END IF
         END SELECT

    	CASE WM_DESTROY
         ' // End the application sending a WM_QUIT message
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hWnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
