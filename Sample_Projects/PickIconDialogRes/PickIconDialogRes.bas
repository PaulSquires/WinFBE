' ########################################################################################
' Microsoft Windows
' File: PickIconDialogRes.bas
' Contents: Demonstrates the use of the pick icon dialog.
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#INCLUDE ONCE "Afx/CWindow.inc"
#INCLUDE ONCE "win/shlobj.bi"
USING Afx

CONST IDC_PICKDLG = 1001   ' // Pick icon dialog identifier

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
   pWindow.Create(NULL, "Pick Icon dialog", @WndProc)
   pWindow.SetClientSize(500, 320)
   pWindow.Center

   ' // Add buttons without position or size (it will be resized in the WM_SIZE message).
   pWindow.AddControl("Button", , IDC_PICKDLG, "&Pick")
   pWindow.AddControl("Button", , IDCANCEL, "&Close")


   ' // Process Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Window procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hWnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   STATIC wszIconPath AS WSTRING * MAX_PATH   ' // Path of the resource file containing the icons
   STATIC nIconIndex AS LONG                  ' // Icon index
   STATIC hIcon AS HICON                      ' // Icon handle

   SELECT CASE uMsg

      CASE WM_COMMAND
         ' // If ESC key pressed, close the application sending an WM_CLOSE message
         SELECT CASE GET_WM_COMMAND_ID(wParam, lParam)
            CASE IDCANCEL
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF
            ' // Launch the Pick icon dialog
            CASE IDC_PICKDLG
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  IF LEN(wszIconPath) = 0 THEN
                     ' // Get the full path of our executable
                     GetModuleFileNameW(NULL, wszIconPath, MAX_PATH)
                  END IF
                  ' // Activate the Pick Icon Common Dialog Box
                  DIM hr AS LONG = PickIconDlg(0, wszIconPath, SIZEOF(wszIconPath), @nIconIndex)
                  ' // If an icon has been selected...
                  IF hr = 1 THEN
                     ' // Destroy previously loaded icon, if any
                     IF hIcon THEN DestroyIcon(hIcon)
                     ' // Get the handle of the new selected icon
                     hIcon = ExtractIconW(GetModuleHandle(NULL), wszIconPath, nIconIndex)
                     ' // Replace the application icons
                     IF hIcon THEN
                        SendMessageW(hwnd, WM_SETICON, ICON_SMALL, cast(LPARAM, hIcon))
                        SendMessageW(hwnd, WM_SETICON, ICON_BIG, cast(LPARAM, hIcon))
                     END IF
                  END IF
                  EXIT FUNCTION
               END IF
         END SELECT

      CASE WM_SIZE
         IF wParam <> SIZE_MINIMIZED THEN
            ' // Resize the buttons
            DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
            IF pWindow THEN pWindow->MoveWindow GetDlgItem(hwnd, IDC_PICKDLG), pWindow->ClientWidth - 230, pWindow->ClientHeight - 50, 75, 23, CTRUE
            IF pWindow THEN pWindow->MoveWindow GetDlgItem(hwnd, IDCANCEL), pWindow->ClientWidth - 120, pWindow->ClientHeight - 50, 75, 23, CTRUE
         END IF

    	CASE WM_DESTROY
         ' // Destroy the icon
         IF hIcon THEN DestroyIcon(hIcon)
         ' // End the application sending a WM_QUIT message
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Process Windows messages
   FUNCTION = DefWindowProcW(hWnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
