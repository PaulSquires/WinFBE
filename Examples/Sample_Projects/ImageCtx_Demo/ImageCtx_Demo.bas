' ########################################################################################
' Microsoft Windows
' File: ImageCtx_Demo
' Contents: Image control demo
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define _WIN32_WINNT &h0602
#INCLUDE ONCE "Afx/CWindow.inc"
#INCLUDE ONCE "Afx/CImageCtx.inc"

USING Afx

CONST IDC_LOAD              = 100
CONST IDC_IMAGECTX          = 101
CONST IDC_RADIO_AUTOSIZE    = 102
CONST IDC_RADIO_ACTUALSIZE  = 103
CONST IDC_RADIO_FITTOHEIGHT = 104
CONST IDC_RADIO_FITTOWIDTH  = 105
CONST IDC_RADIO_STRETCH     = 106

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG

   END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

' ========================================================================================
' Displays the File Open Dialog and loads the selected file in the image control.
' ========================================================================================
SUB AfxLoadFileDialog (BYVAL hwndOwner AS HWND, BYVAL pImageCtx AS CImageCtx PTR, BYVAL sigdnName AS SIGDN = SIGDN_FILESYSPATH)

   ' // Create an instance of the FileOpenDialog interface
   DIM hr AS LONG
   DIM pofd AS IFileOpenDialog PTR
   hr = CoCreateInstance(@CLSID_FileOpenDialog, NULL, CLSCTX_INPROC_SERVER, @IID_IFileOpenDialog, @pofd)
   IF pofd = NULL THEN EXIT SUB

   ' // Set the file types
   DIM rgFileTypes(1 TO 5) AS COMDLG_FILTERSPEC
   rgFileTypes(1).pszName = @WSTR("JPG Files (*.jpg)")
   rgFileTypes(2).pszName = @WSTR("BMP Files (*.bmp)")
   rgFileTypes(3).pszName = @WSTR("PNG Files (*.png)")
   rgFileTypes(4).pszName = @WSTR("TIF Files (*.tif)")
   rgFileTypes(5).pszName = @WSTR("All files")
   rgFileTypes(1).pszSpec = @WSTR("*.jpg")
   rgFileTypes(2).pszSpec = @WSTR("*.bmp")
   rgFileTypes(3).pszSpec = @WSTR("*.png")
   rgFileTypes(4).pszSpec = @WSTR("*.tif")
   rgFileTypes(5).pszSpec = @WSTR("*.*")
   pofd->lpVtbl->SetFileTypes(pofd, 5, @rgFileTypes(1))

   ' // Set the title of the dialog
   hr = pofd->lpVtbl->SetTitle(pofd, "A Single-Selection Dialog")
   ' // Display the dialog
   hr = pofd->lpVtbl->Show(pofd, hwndOwner)

   ' // Set the default folder
   DIM pFolder AS IShellItem PTR
   SHCreateItemFromParsingName (CURDIR, NULL, @IID_IShellItem, @pFolder)
   IF pFolder THEN
      pofd->lpVtbl->SetDefaultFolder(pofd, pFolder)
      pFolder->lpVtbl->Release(pFolder)
   END IF

   ' // Get the result
   DIM pItem AS IShellItem PTR
   DIM pwszName AS WSTRING PTR
   IF SUCCEEDED(hr) THEN
      hr = pofd->lpVtbl->GetResult(pofd, @pItem)
      IF SUCCEEDED(hr) THEN
         hr = pItem->lpVtbl->GetDisplayName(pItem, sigdnName, @pwszName)
         pImageCtx->LoadImageFromFile(*pwszName)
         CoTaskMemFree(pwszName)
      END IF
   END IF

   ' // Cleanup
   IF pItem THEN pItem->lpVtbl->Release(pItem)
   IF pofd THEN pofd->lpVtbl->Release(pofd)

END SUB
' ========================================================================================

' ========================================================================================
' Main
' ========================================================================================
FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                  BYVAL hPrevInstance AS HINSTANCE, _
                  BYVAL szCmdLine AS ZSTRING PTR, _
                  BYVAL nCmdShow AS LONG) AS LONG

   ' // Initialize the COM library
   CoInitialize(NULL)

   ' // Set process DPI aware
   ' // The recommended way is to use a manifest
'   AfxSetProcessDPIAware

   ' // Create the main window
   DIM pWindow AS CWindow
   pWindow.Create(NULL, "CWindow with an image control", @WndProc)
   pWindow.SetClientSize(600, 350)
   pWindow.Center

   ' // Add a button without coordinates (it will be reiszed in WM_SIZE, below)
   DIM hCtl AS HWND = pWindow.AddControl("RadioButton", , IDC_RADIO_AUTOSIZE, "Autosize")
   SendMessageW hCtl, BM_SETCHECK, BST_CHECKED, 0
   SetFocus hCtl
   pWindow.AddControl("RadioButton", , IDC_RADIO_ACTUALSIZE, "Actual size")
   pWindow.AddControl("RadioButton", , IDC_RADIO_FITTOWIDTH, "Fit to width")
   pWindow.AddControl("RadioButton", , IDC_RADIO_FITTOHEIGHT, "Fit to height")
   pWindow.AddControl("RadioButton", , IDC_RADIO_STRETCH, "Stretch")
   pWindow.AddControl("Button", , IDC_LOAD, "Load")
   pWindow.AddControl("Button", , IDCANCEL, "Close")

   ' // Add an image control
   DIM pImageCtx AS CImageCtx = CImageCtx(@pWindow, IDC_IMAGECTX, , 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   ' // Load an image from disk
   pImageCtx.LoadImageFromFile ".\Images\Ciutat_de_les_Arts_i_de_les_Ciencies_01.jpg"


   ' // Process Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

   ' // Uninitialize the COM library
   CoUninitialize

END FUNCTION
' ========================================================================================

' ========================================================================================
' Window procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   DIM pWindow AS CWindow PTR, pImageCtx AS CImageCtx PTR

   SELECT CASE uMsg

      CASE WM_COMMAND
         SELECT CASE GET_WM_COMMAND_ID(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application sending an WM_CLOSE message
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF
            CASE IDC_LOAD
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  ' // Load the picture
                  AfxLoadFileDialog hwnd, AfxCImageCtxPtr(hwnd, IDC_IMAGECTX)
               END IF
            CASE IDC_RADIO_AUTOSIZE
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  AfxCImageCtxPtr(hwnd, IDC_IMAGECTX)->SetImageAdjustment GDIP_IMAGECTX_AUTOSIZE, CTRUE
                  ' // Alternate way: Get a pointer to the CImageCtx class
'                  pImageCtx = CAST(CImageCtx PTR, GetWindowLongPtr(GetDlgItem(hwnd, IDC_IMAGECTX), 0))
'                  pImageCtx->SetImageAdjustment GDIP_IMAGECTX_AUTOSIZE, CTRUE
                  EXIT FUNCTION
               END IF
            CASE IDC_RADIO_ACTUALSIZE
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  AfxCImageCtxPtr(hwnd, IDC_IMAGECTX)->SetImageAdjustment GDIP_IMAGECTX_ACTUALSIZE, CTRUE
                  ' // Alternate way: Get a pointer to the CImageCtx class
'                  pImageCtx = CAST(CImageCtx PTR, GetWindowLongPtr(GetDlgItem(hwnd, IDC_IMAGECTX), 0))
'                  pImageCtx->SetImageAdjustment GDIP_IMAGECTX_ACTUALSIZE, CTRUE
                  EXIT FUNCTION
               END IF
            CASE IDC_RADIO_FITTOWIDTH
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  AfxCImageCtxPtr(hwnd, IDC_IMAGECTX)->SetImageAdjustment GDIP_IMAGECTX_FITTOWIDTH, CTRUE
                  ' // Alternate way: Get a pointer to the CImageCtx class
'                  pImageCtx = CAST(CImageCtx PTR, GetWindowLongPtr(GetDlgItem(hwnd, IDC_IMAGECTX), 0))
'                  pImageCtx->SetImageAdjustment GDIP_IMAGECTX_FITTOWIDTH, CTRUE
                  EXIT FUNCTION
               END IF
            CASE IDC_RADIO_FITTOHEIGHT
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  AfxCImageCtxPtr(hwnd, IDC_IMAGECTX)->SetImageAdjustment GDIP_IMAGECTX_FITTOHEIGHT, CTRUE
                  ' // Alternate way: Get a pointer to the CImageCtx class
'                  pImageCtx = CAST(CImageCtx PTR, GetWindowLongPtr(GetDlgItem(hwnd, IDC_IMAGECTX), 0))
'                  pImageCtx->SetImageAdjustment GDIP_IMAGECTX_FITTOHEIGHT, CTRUE
                  EXIT FUNCTION
               END IF
            CASE IDC_RADIO_STRETCH
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  AfxCImageCtxPtr(hwnd, IDC_IMAGECTX)->SetImageAdjustment GDIP_IMAGECTX_STRETCH, CTRUE
                  ' // Alternate way: Get a pointer to the CImageCtx class
'                  pImageCtx = CAST(CImageCtx PTR, GetWindowLongPtr(GetDlgItem(hwnd, IDC_IMAGECTX), 0))
'                  pImageCtx->SetImageAdjustment GDIP_IMAGECTX_STRETCH, CTRUE
                  EXIT FUNCTION
               END IF
         END SELECT

      CASE WM_NOTIFY
         ' // Process notification messages
         DIM phdr AS NMHDR PTR
         phdr = cast(NMHDR PTR, lParam)
         IF wParam = IDC_IMAGECTX THEN
            SELECT CASE phdr->code
               CASE NM_CLICK : MessageBox hwnd, "NM_CLICK", "", MB_OK
               CASE NM_DBLCLK
               CASE NM_RCLICK
               CASE NM_RDBLCLK
               CASE NM_SETFOCUS
               CASE NM_KILLFOCUS
            END SELECT
         END IF

      CASE WM_SIZE
         pWindow = AfxCWindowPtr(hwnd)
         ' // Alternate way:
'         pWindow = CAST(CWindow PTR, GetWindowLongPtr(hwnd, 0))
         ' // If the window isn't minimized, resize it
         IF wParam <> SIZE_MINIMIZED THEN
            pWindow->MoveWindow GetDlgItem(hwnd, IDC_IMAGECTX), 10, 10, pWindow->ClientWidth - 20, pWindow->ClientHeight - 65, CTRUE
            pWindow->MoveWindow GetDlgItem(hwnd, IDC_LOAD), pWindow->ClientWidth - 185, pWindow->ClientHeight - 35, 75, 23, CTRUE
            pWindow->MoveWindow GetDlgItem(hwnd, IDCANCEL), pWindow->ClientWidth - 95, pWindow->ClientHeight - 35, 75, 23, CTRUE
            pWindow->MoveWindow GetDlgItem(hwnd, IDC_RADIO_AUTOSIZE), 15, pWindow->ClientHeight - 49, 100, 21, CTRUE
            pWindow->MoveWindow GetDlgItem(hwnd, IDC_RADIO_ACTUALSIZE), 15, pWindow->ClientHeight - 28, 100, 21, CTRUE
            pWindow->MoveWindow GetDlgItem(hwnd, IDC_RADIO_FITTOWIDTH), 150, pWindow->ClientHeight - 49, 100, 21, CTRUE
            pWindow->MoveWindow GetDlgItem(hwnd, IDC_RADIO_FITTOHEIGHT), 150, pWindow->ClientHeight - 28, 100, 21, CTRUE
            pWindow->MoveWindow GetDlgItem(hwnd, IDC_RADIO_STRETCH), 275, pWindow->ClientHeight - 28, 60, 21, CTRUE
         END IF

    	CASE WM_DESTROY
         ' // Ends the application by sending a WM_QUIT message
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hWnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
