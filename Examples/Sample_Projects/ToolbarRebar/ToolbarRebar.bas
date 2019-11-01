' ########################################################################################
' Microsoft Windows
' File: ToolbarRebar.bas
' Contents: CWindow with a toolbar inside a rebar
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
CONST IDC_REBAR = 1002
enum
   IDM_LEFTARROW = 28000
   IDM_RIGHTARROW
   IDM_HOME
   IDM_SAVEFILE
   IDM_HISTORY
   IDM_MAIL
   IDM_PRINT
   IDM_SEARCH
   IDM_VIEW
   IDM_STOP
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
   pWindow.Create(NULL, "CWindow with a toolbar-rebar", @WndProc)
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
   DIM cx AS LONG = 24 * pWindow.DPI \ 96

   ' // Create an image list for the toolbar
   DIM hImageList AS HIMAGELIST
   hImageList = ImageList_Create(cx, cx, ILC_COLOR32 OR ILC_MASK, 7, 0)
   IF hImageList THEN
      AfxGdipAddIconFromRes(hImageList, hInstance, "IDI_ARROW_LEFT")
      AfxGdipAddIconFromRes(hImageList, hInstance, "IDI_ARROW_RIGHT")
      AfxGdipAddIconFromRes(hImageList, hInstance, "IDI_HOME")
      AfxGdipAddIconFromRes(hImageList, hInstance, "IDI_SAVE")
      AfxGdipAddIconFromRes(hImageList, hInstance, "IDI_HISTORY")
      AfxGdipAddIconFromRes(hImageList, hInstance, "IDI_MAIL")
      AfxGdipAddIconFromRes(hImageList, hInstance, "IDI_PRINT")
      AfxGdipAddIconFromRes(hImageList, hInstance, "IDI_SEARCH")
      AfxGdipAddIconFromRes(hImageList, hInstance, "IDI_VIEW")
      AfxGdipAddIconFromRes(hImageList, hInstance, "IDI_STOP")
      ' // Set the normal image list
      Toolbar_SetImageList hToolBar, hImageList
      ' // Set the hot image list with the same images than the normal one
      Toolbar_SetHotImageList hToolBar, hImageList
   END IF

   ' // Create a disabled image list for the toolbar
   DIM hDisabledImageList AS HIMAGELIST
   hDisabledImageList = ImageList_Create(cx, cx, ILC_COLOR32 OR ILC_MASK, 4, 0)
   IF hDisabledImageList THEN
      AfxGdipAddIconFromRes(hDisabledImageList, hInstance, "IDI_ARROW_LEFT", 60, TRUE)
      AfxGdipAddIconFromRes(hDisabledImageList, hInstance, "IDI_ARROW_RIGHT", 60, TRUE)
      AfxGdipAddIconFromRes(hDisabledImageList, hInstance, "IDI_HOME", 60, TRUE)
      AfxGdipAddIconFromRes(hDisabledImageList, hInstance, "IDI_SAVE", 60, TRUE)
      AfxGdipAddIconFromRes(hDisabledImageList, hInstance, "IDI_HISTORY", 60, TRUE)
      AfxGdipAddIconFromRes(hDisabledImageList, hInstance, "IDI_MAIL", 60, TRUE)
      AfxGdipAddIconFromRes(hDisabledImageList, hInstance, "IDI_PRINT", 60, TRUE)
      AfxGdipAddIconFromRes(hDisabledImageList, hInstance, "IDI_SEARCH", 60, TRUE)
      AfxGdipAddIconFromRes(hDisabledImageList, hInstance, "IDI_VIEW", 60, TRUE)
      AfxGdipAddIconFromRes(hDisabledImageList, hInstance, "IDI_STOP", 60, TRUE)
      ' // Set the disabled image list
      Toolbar_SetDisabledImageList hToolBar, hDisabledImageList
   END IF

   ' // Add buttons to the toolbar
   Toolbar_AddButton hToolBar, 0, IDM_LEFTARROW
   Toolbar_AddButton hToolBar, 1, IDM_RIGHTARROW
   Toolbar_AddButton hToolBar, 2, IDM_HOME
   Toolbar_AddButton hToolBar, 3, IDM_SAVEFILE
   Toolbar_AddButton hToolBar, 4, IDM_HISTORY
   Toolbar_AddButton hToolBar, 5, IDM_MAIL
   Toolbar_AddButton hToolBar, 6, IDM_PRINT
   Toolbar_AddButton hToolBar, 7, IDM_SEARCH
   Toolbar_AddButton hToolBar, 8, IDM_VIEW
   Toolbar_AddButton hToolBar, 9, IDM_STOP

   ' // Disable the save file button
   Toolbar_DisableButton(hToolBar, IDM_SAVEFILE)

   ' // Size the toolbar
   Toolbar_AutoSize hToolBar

   ' // Create a rebar control
   DIM hRebar AS HWND = pWindow.AddControl("Rebar", , IDC_REBAR)

   ' // Add the band containing the toolbar control to the rebar
   ' // The size of the REBARBANDINFOW is different in Vista/Windows 7
   DIM rc AS RECT, wszText AS WSTRING * 260
   DIM rbbi AS REBARBANDINFOW
   IF AfxWindowsVersion >= 600 AND AfxComCtlVersion >= 600 THEN
      rbbi.cbSize  = REBARBANDINFO_V6_SIZE
   ELSE
      rbbi.cbSize  = REBARBANDINFO_V3_SIZE
   END IF

   ' // Insert the toolbar in the rebar control
   rbbi.fMask      = RBBIM_STYLE OR RBBIM_CHILD OR RBBIM_CHILDSIZE OR _
                     RBBIM_SIZE OR RBBIM_ID OR RBBIM_IDEALSIZE OR RBBIM_TEXT
   rbbi.fStyle     = RBBS_CHILDEDGE
   rbbi.hwndChild  = hToolbar
   rbbi.cxMinChild = 270 * pWindow.rxRatio
   rbbi.cyMinChild = Toolbar_GetButtonHeight(hToolbar)
   rbbi.cx         = 270 * pWindow.rxRatio
   rbbi.cxIdeal    = 270 * pWindow.rxRatio
   wszText = "Toolbar"
   rbbi.lpText = @wszText
   '// Insert band into rebar
   Rebar_InsertBand(hRebar, -1, @rbbi)

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

      CASE WM_NOTIFY
         ' -------------------------------------------------------
         ' Notification messages are handled here.
         ' The TTN_GETDISPINFO message is sent by a ToolTip control
         ' to retrieve information needed to display a ToolTip window.
         ' ------------------------------------------------------
         DIM ptnmhdr AS NMHDR PTR              ' // Information about a notification message
         DIM ptttdi AS NMTTDISPINFOW PTR       ' // Tooltip notification message information
         DIM wszTooltipText AS WSTRING * 260   ' // Tooltip text

         ptnmhdr = CAST(NMHDR PTR, lParam)
         SELECT CASE ptnmhdr->code
            ' // The height of the rebar has changed
            CASE RBN_HEIGHTCHANGE
               ' // Get the coordinates of the client area
               DIM rc AS RECT
               GetClientRect hwnd, @rc
               ' // Send a WM_SIZE message to resize the controls
               SendMessageW hwnd, WM_SIZE, SIZE_RESTORED, MAKELONG(rc.Right - rc.Left, rc.Bottom - rc.Top)
            ' // Toolbar tooltips
            CASE TTN_GETDISPINFOW
               ptttdi = CAST(NMTTDISPINFOW PTR, lParam)
               ptttdi->hinst = NULL
               wszTooltipText = ""
               SELECT CASE ptttdi->hdr.hwndFrom
                  CASE Toolbar_GetTooltips(GetDlgItem(GetDlgItem(hwnd, IDC_REBAR), IDC_TOOLBAR))
                     SELECT CASE ptttdi->hdr.idFrom
                        CASE IDM_LEFTARROW  : wszTooltipText = "Left arrow"
                        CASE IDM_RIGHTARROW : wszTooltipText = "Right arrow"
                        CASE IDM_HOME       : wszTooltipText = "Home"
                        CASE IDM_SAVEFILE   : wszTooltipText = "Save"
                        CASE IDM_HISTORY    : wszTooltipText = "History"
                        CASE IDM_MAIL       : wszTooltipText = "Mail"
                        CASE IDM_PRINT      : wszTooltipText = "Print"
                        CASE IDM_SEARCH     : wszTooltipText = "Search"
                        CASE IDM_VIEW       : wszTooltipText = "View"
                        CASE IDM_STOP       : wszTooltipText = "Stop"
                     END SELECT
                     IF LEN(wszTooltipText) THEN ptttdi->lpszText = @wszTooltipText
               END SELECT
         END SELECT

      CASE WM_SIZE
         IF wParam <> SIZE_MINIMIZED THEN
            ' // Update the size and position of the Rebar control
            SendMessageW GetDlgItem(hwnd, IDC_REBAR), WM_SIZE, wParam, lParam
            EXIT FUNCTION
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
