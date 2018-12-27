' ########################################################################################
' Microsoft Windows
' File: Toolbar_HDPI.bas
' Contents: CWindow with a toolbar
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define unicode
#INCLUDE ONCE "windows.bi"
#INCLUDE ONCE "Afx/CWindow.inc"
#INCLUDE ONCE "Afx/AfxGdiplus.inc"
USING Afx

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG

   END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

CONST IDC_TOOLBAR = 1001
enum
   IDM_LEFTARROW = 28000
   IDM_RIGHTARROW
   IDM_HOME
   IDM_SAVEFILE
end enum

' // Forward reference
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

' ========================================================================================
' Main
' ========================================================================================
FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                  BYVAL hPrevInstance AS HINSTANCE, _
                  BYVAL szCmdLine AS ZSTRING PTR, _
                  BYVAL nCmdShow AS LONG) AS LONG

   ' // Set process DPI aware
   ' // The recommended way is to use a manifest file
'   AfxSetProcessDPIAware

   DIM pWindow AS CWindow
   pWindow.Create(NULL, "CWindow with a toolbar", @WndProc)
   ' // Disable background erasing
   pWindow.ClassStyle = CS_DBLCLKS
   ' // Set the client size
   pWindow.SetClientSize(500, 300)
   ' // Center the window
   pWindow.Center

   ' // Add a button
   pWindow.AddControl("Button", , IDCANCEL, "&Close")

   ' // Add a tooolbar
   DIM hToolBar AS HWND = pWindow.AddControl("Toolbar", , IDC_TOOLBAR)

   ' // Calculate the size of the icon according the DPI
   DIM cx AS LONG = 20 * pWindow.DPI \ 96

   ' // Create an image list for the toolbar
   DIM hImageList AS HIMAGELIST
   hImageList = ImageList_Create(cx, cx, ILC_COLOR32 OR ILC_MASK, 4, 0)
   IF hImageList THEN
      AfxGdipAddIconFromFile(hImageList, "./Resources/arrow_left_48.png")
      AfxGdipAddIconFromFile(hImageList, "./Resources/arrow_right_48.png")
      AfxGdipAddIconFromFile(hImageList, "./Resources/home_48.png")
      AfxGdipAddIconFromFile(hImageList, "./Resources/save_48.png")
      ' // Set the normal image list
      Toolbar_SetImageList hToolBar, hImageList
      ' // Set the hot image list with the same images than the normal one
      Toolbar_SetHotImageList hToolBar, hImageList
   END IF

   ' // Create a disabled image list for the toolbar
   DIM hDisabledImageList AS HIMAGELIST
   hDisabledImageList = ImageList_Create(cx, cx, ILC_COLOR32 OR ILC_MASK, 4, 0)
   IF hDisabledImageList THEN
      AfxGdipAddIconFromFile(hDisabledImageList, "./Resources/arrow_left_48.png", 60, TRUE)
      AfxGdipAddIconFromFile(hDisabledImageList, "./Resources/arrow_right_48.png", 60, TRUE)
      AfxGdipAddIconFromFile(hDisabledImageList, "./Resources/home_48.png", 60, TRUE)
      AfxGdipAddIconFromFile(hDisabledImageList, "./Resources/save_48.png", 60, TRUE)
      ' // Set the disabled image list
      Toolbar_SetDisabledImageList hToolBar, hDisabledImageList
   END IF

   ' // Add buttons to the toolbar
   Toolbar_AddButton hToolBar, 0, IDM_LEFTARROW
   Toolbar_AddButton hToolBar, 1, IDM_RIGHTARROW
   Toolbar_AddButton hToolBar, 2, IDM_HOME
   Toolbar_AddButton hToolBar, 3, IDM_SAVEFILE

   ' // Disable the save file button
   Toolbar_DisableButton(hToolBar, IDM_SAVEFILE)

   ' // Size the toolbar
   Toolbar_AutoSize hToolBar

   ' // Process event messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Window procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   SELECT CASE uMsg

      CASE WM_CREATE
         EXIT FUNCTION

      CASE WM_COMMAND
         SELECT CASE GET_WM_COMMAND_ID(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application sending an WM_CLOSE message
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF
'            CASE IDM_LEFTARROW   ' etc.
'               MessageBoxW hwnd, "You have clicked the left arrow button", "Toolbar", MB_OK
'               EXIT FUNCTION
         END SELECT

      CASE WM_SIZE
         IF wParam <> SIZE_MINIMIZED THEN
            ' // Update the size and position of the Toolbar control
            Toolbar_Autosize GetDlgItem(hwnd, IDC_TOOLBAR)
            ' // Resize the button
            DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
            pWindow->MoveWindow GetDlgItem(hwnd, IDCANCEL), pWindow->ClientWidth - 95, pWindow->ClientHeight - 35, 75, 23, CTRUE
         END IF

    	CASE WM_DESTROY
         ' // Destroy the image lists
         ImageList_Destroy CAST(HIMAGELIST, SendMessageW(GetDlgItem(hwnd, IDC_TOOLBAR), TB_SETIMAGELIST, 0, 0))
         ImageList_Destroy CAST(HIMAGELIST, SendMessageW(GetDlgItem(hwnd, IDC_TOOLBAR), TB_SETHOTIMAGELIST, 0, 0))
         ImageList_Destroy CAST(HIMAGELIST, SendMessageW(GetDlgItem(hwnd, IDC_TOOLBAR), TB_SETDISABLEDIMAGELIST, 0, 0))
         ' // Quit the application
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hWnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
