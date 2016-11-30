' ########################################################################################
' Microsoft Windows
' File: Split_Button.bas
' Contents: CWindow with a split button
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#define _WIN32_WINNT &h0602
#INCLUDE ONCE "Afx/CWindow.inc"
#INCLUDE ONCE "Afx/AfxGdiplus.inc"
USING Afx

CONST IDC_MENUCOMMAND1 = 28000
CONST IDC_MENUCOMMAND2 = 28001

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

   ' // Create the main window
   DIM pWindow AS CWindow
   pWindow.Create(NULL, "CWindow with a split button", @WndProc)
   pWindow.SetClientSize(300, 150)
   pWindow.Center

   ' // Add a button without position or size (it will be resized in the WM_SIZE message).
   DIM hSplitButton AS HWND = pWindow.AddControl("Button", pWindow.hWindow, IDCANCEL, "&Shutdown", , , , , BS_SPLITBUTTON)

   ' // Calculate an appropriate icon size
   DIM cx AS LONG = 16 * pWindow.DPI \ 96
   ' // Create an image list for the button
   DIM hImageList AS HIMAGELIST = ImageList_Create(cx, cx, ILC_COLOR32 OR ILC_MASK, 1, 0)
   ' // --> Remember to change the path and name of the icon
   IF hImageList THEN ImageList_ReplaceIcon(hImageList, -1, AfxGdipImageFromFile("./Resources/Shutdown_48.png"))
   ' // Fill a BUTTON_IMAGELIST structure and set the image list
   DIM bi AS BUTTON_IMAGELIST = (hImageList, (3, 3, 3, 3), BUTTON_IMAGELIST_ALIGN_LEFT)
   Button_SetImageList(hSplitButton, @bi)

   ' // Dispatch Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Main window callback procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hWnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

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

      CASE WM_NOTIFY
         ' // Processs notify messages sent by the split button
         DIM pNmh AS NMHDR PTR = CAST(NMHDR PTR, lParam)
         IF pNmh->idFrom = IDCANCEL AND pNmh->code = BCN_DROPDOWN THEN
            DIM pDropDown AS NMBCDROPDOWN PTR = CAST(NMBCDROPDOWN PTR, lParam)
            ' // Get screen coordinates of the button
            DIM pt AS POINT = (pDropdown->rcButton.left, pDropDown->rcButton.bottom)
            ClientToScreen(pNmh->hwndFrom, @pt)
            ' // Create a menu and add items
            DIM hSplitMenu AS HMENU = CreatePopupMenu
            AppendMenuW(hSplitMenu, MF_BYPOSITION, IDC_MENUCOMMAND1, "Menu item 1")
            AppendMenuW(hSplitMenu, MF_BYPOSITION, IDC_MENUCOMMAND2, "Menu item 2")
            ' // Display the menu
            TrackPopupMenu(hSplitMenu, TPM_LEFTALIGN OR TPM_TOPALIGN, pt.x, pt.y, 0, hwnd, NULL)
            FUNCTION = CTRUE
            EXIT FUNCTION
         END IF

      CASE WM_SIZE
         IF wParam <> SIZE_MINIMIZED THEN
            ' // Resize the button
            DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
            IF pWindow THEN pWindow->MoveWindow GetDlgItem(hwnd, IDCANCEL), pWindow->ClientWidth - 200, pWindow->ClientHeight - 90, 110, 23, CTRUE
         END IF

    	CASE WM_DESTROY
         ' // Destroy the image list
         ImageList_Destroy CAST(HIMAGELIST, Toolbar_SetImageList(GetDlgItem(hwnd, IDCANCEL), 0))
         ' // End the application by sending an WM_QUIT message
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hWnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
