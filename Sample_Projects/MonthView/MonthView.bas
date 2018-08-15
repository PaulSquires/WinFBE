' ########################################################################################
' Microsoft Windows
' Contents: Embedded MonthView Calendar OCX
' Compiler: FreeBasic 32 bit only
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "Afx/CAxHost/CAxHost.inc"
#INCLUDE ONCE "Afx/CDispInvoke.inc"
#INCLUDE ONCE "Afx/AfxCOM.inc"
USING Afx

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG

   END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

CONST IDC_MONTHVIEW = 1001

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

   ' // Creates the main window
   DIM pWindow AS CWindow
   ' -or- DIM pWindow AS CWindow = "MyClassName" (use the name that you wish)
   DIM hwndMain AS HWND = pWindow.Create(NULL, "CAxHost - Embedded MonthView Calendar OCX", @WndProc)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(580, 360)
   ' // Centers the window
   pWindow.Center

   DIM wszLibName AS WSTRING * 260 = ExePath & "\MSCOMCT2.OCX"
   DIM CLSID_MSComCtl2_MonthView AS CLSID = (&h232E456A, &h87C3, &h11D1, {&h8B, &hE3,&h00, &h00, &hF8, &h75, &h4D, &hA1})
   DIM IID_MSComCtl2_MonthView AS IID = (&h232E4565, &h87C3, &h11D1, {&h8B, &hE3,&h00, &h00, &hF8, &h75, &h4D, &hA1})
   DIM RTLKEY_MSCOMCT2 AS WSTRING * 260 = "651A8940-87C5-11d1-8BE3-0000F8754DA1"

   DIM pHost AS CAxHost = CAxHost(@pWindow, IDC_MONTHVIEW, wszLibName, CLSID_MSComCtl2_MonthView, _
       IID_MSComCtl2_MonthView, RTLKEY_MSCOMCT2, 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   SetFocus pHost.hWindow

   ' // Use CDispInvoke to set the date
   DIM pdisp AS CDispInvoke = pHost.OcxDispObj
   pdisp.Put("Year") = 2017
   pdisp.Put("Month") = 12
   pdisp.Put("Day") = 1

   ' // Display the window
   ShowWindow(hwndMain, nCmdShow)
   UpdateWindow(hwndMain)

   ' // Dispatch Windows messages
   DIM uMsg AS MSG
   WHILE (GetMessageW(@uMsg, NULL, 0, 0) <> FALSE)
      IF AfxCAxHostForwardMessage(GetFocus, @uMsg) = FALSE THEN
         IF IsDialogMessageW(hwndMain, @uMsg) = 0 THEN
            TranslateMessage(@uMsg)
            DispatchMessageW(@uMsg)
         END IF
      END IF
   WEND
   FUNCTION = uMsg.wParam

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

      CASE WM_SIZE
         ' // Optional resizing code
         IF wParam <> SIZE_MINIMIZED THEN
            ' // Retrieve a pointer to the CWindow class
            DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
            ' // Move the position of the button
            IF pWindow THEN pWindow->MoveWindow GetDlgItem(hwnd, 1001), _
               0, 0, pWindow->ClientWidth, pWindow->ClientHeight, CTRUE
         END IF

    	CASE WM_DESTROY
         ' // Ends the application by sending a WM_QUIT message
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
