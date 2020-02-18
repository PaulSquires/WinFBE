' ########################################################################################
' Microsoft Windows
' DirectShow example.
' Allows to select a video clip and plays it.
' Based on an example by Vladimir Shulakov posted in the PowerBASIC forums:
' https://forum.powerbasic.com/forum/user-to-user-discussions/source-code/24615-directshow-small-example?t=23966
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "Afx/CWindow.inc"
#INCLUDE ONCE "Afx/AfxCOM.inc"
#INCLUDE ONCE "win/uuids.bi"
#INCLUDE ONCE "win/dshow.bi"
#INCLUDE ONCE "win/strmif.bi"
#INCLUDE ONCE "win/control.bi"
USING Afx

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG

   END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

' // Menu identifiers
CONST ID_FILE_OPENCLIP = 40001
CONST ID_FILE_EXIT     = 40002

' // Custom message
CONST WM_GRAPHNOTIFY   = WM_USER + 13

' // Globals
DIM SHARED bIsPlaying AS BOOLEAN
DIM SHARED pIMediaControl AS IMediaControl PTR
DIM SHARED pIVideoWindow AS IVideoWindow PTR
DIM SHARED pIMediaEventEx AS IMediaEventEx PTR
DIM SHARED pIGraphBuilder AS IGraphBuilder PTR

' // Forward declaration
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

   ' // Initialize the COM library
   CoInitialize(NULL)

   ' // Creates the main window
   DIM pWindow AS CWindow
   pWindow.Create(NULL, "DirectShow demo: Play clip", @WndProc)
   ' // Use a black brush for the background
   pWindow.Brush = CreateSolidBrush(BGR(0, 0, 0))
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(400, 320)
   ' // Centers the window
   pWindow.Center

   ' // Create the menu
   DIM hMenu AS HMENU = CreateMenu
   DIM hMenuFile AS HMENU = CreatePopUpMenu
   AppendMenuW hMenu, MF_POPUP OR MF_ENABLED, CAST(UINT_PTR, hMenuFile), "&File"
   AppendMenuW hMenuFile, MF_ENABLED, ID_FILE_OPENCLIP, "&Open clip..."
   AppendMenuW hMenuFile, MF_ENABLED, ID_FILE_EXIT, "E&xit"
   SetMenu pWindow.hWindow, hMenu

   ' // Displays the window and dispatches the Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

   ' // Uninitalize COM
   CoUninitialize

END FUNCTION
' ========================================================================================

' ========================================================================================
' Play the movie inside the window.
' ========================================================================================
SUB PlayMovieInWindow (BYVAL hwnd AS HWND, BYREF wszFileName AS WSTRING)

   DIM hr AS HRESULT

   ' // If there is a clip loaded, stop it
   IF pIMediaControl THEN
      hr = pIMediaControl->lpvtbl->Stop(pIMediaControl)
      AfxSafeRelease(pIMediaControl) : pIMediaControl = NULL
      AfxSafeRelease(pIVideoWindow) : pIVideoWindow = NULL
      AfxSafeRelease(pIMediaEventEx) : pIMediaEventEx = NULL
      AfxSafeRelease(pIGraphBuilder) : pIGraphBuilder = NULL
   END IF

   ' // Create an instance of the IGraphBuilder object
   pIGraphBuilder = AfxNewCom(CLSID_FilterGraph)
   IF pIGraphBuilder = NULL THEN EXIT SUB

   ' // Further error checking omitted for brevity
   ' // We should add IF hr <> S_OK THEN ...

   ' // Retrieve interafce pointers
   hr = pIGraphBuilder->lpvtbl->QueryInterface(pIGraphBuilder, @IID_IMediaControl, @pIMediaControl)
   hr = pIGraphBuilder->lpvtbl->QueryInterface(pIGraphBuilder, @IID_IMediaEventEx, @pIMediaEventEx)
   hr = pIGraphBuilder->lpvtbl->QueryInterface(pIGraphBuilder, @IID_IVideoWindow, @pIVideoWindow)

   ' // Render the file
   hr = pIMediaControl->lpvtbl->RenderFile(pIMediaControl, @wszFileName)

   '// Set the window owner and style
   hr = pIVideoWindow->lpvtbl->put_Visible(pIVideoWindow, OAFALSE)
   hr = pIVideoWindow->lpvtbl->put_Owner(pIVideoWindow, cast(OAHWND, hwnd))
   hr = pIVideoWindow->lpvtbl->put_WindowStyle(pIVideoWindow, WS_CHILD OR WS_CLIPSIBLINGS OR WS_CLIPCHILDREN)

   ' // Have the graph signal event via window callbacks for performance
   hr = pIMediaEventEx->lpvtbl->SetNotifyWindow(pIMediaEventEx, cast(OAHWND, hwnd), WM_GRAPHNOTIFY, 0)

   ' // Set the window position
   DIM rc AS RECT
   GetClientRect(hwnd, @rc)
   hr = pIVideoWindow->lpvtbl->SetWindowPosition(pIVideoWindow, rc.Left, rc.Top, rc.Right, rc.Bottom)
   ' // Make the window visible
   hr = pIVideoWindow->lpvtbl->put_Visible(pIVideoWindow, OATRUE)

   ' // Run the graph
   hr = pIMediaControl->lpvtbl->Run(pIMediaControl)
   bIsPlaying = TRUE

END SUB
' ========================================================================================

' ========================================================================================
' Main window procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   SELECT CASE uMsg

      CASE WM_SYSCOMMAND
         ' // Capture this message and send a WM_CLOSE message
         IF (wParam AND &HFFF0) = SC_CLOSE THEN
            SendMessageW hwnd, WM_CLOSE, 0, 0
            EXIT FUNCTION
         END IF

      CASE WM_COMMAND
         SELECT CASE GET_WM_COMMAND_ID(wParam, lParam)
            CASE IDCANCEL, ID_FILE_EXIT
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF
            CASE ID_FILE_OPENCLIP
               ' // Open file dialog
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  DIM wszFile AS WSTRING * MAX_PATH = "*.wmv"
                  DIM wszInitialDir AS STRING * MAX_PATH = CURDIR
                  DIM wszFilter AS WSTRING * 260 = "Video Files (*.MPG;*.MPEG;*.AVI;*.MOV;*.QTM;*.WMV)|*.MPG;*.MPEG;*.AVI;*.MOV;*.QT;*.WMV|" & "All Files (*.*)|*.*|"
                  DIM dwFlags AS DWORD = OFN_EXPLORER OR OFN_FILEMUSTEXIST OR OFN_HIDEREADONLY
                  DIM wszFileName AS WSTRING * MAX_PATH
                  wszFileName = AfxOpenFileDialog(hwnd, "", wszFile, wszInitialDir, wszFilter, "wmv", @dwFlags, NULL)
                  IF LEN(wszFileName) THEN PlayMovieInWindow(hwnd, wszFileName)
                  EXIT FUNCTION
               END IF
         END SELECT

      CASE WM_GRAPHNOTIFY

         ' // WM_GRAPHNOTIFY is an ordinary Windows message. Whenever the Filter Graph Manager
         ' // puts a new event on the event queue, it posts a WM_GRAPHNOTIFY message to the
         ' // designated application window. The message's lParam parameter is equal to the third
         ' // parameter in SetNotifyWindow. This parameter enables you to send instance data with
         ' // the message. The window message's wParam parameter is always zero.

         DIM lEventCode AS LONG
         DIM lParam1 AS LONG_PTR
         DIM lParam2 AS LONG_PTR

         IF pIMediaEventEx THEN
            DO
               DIM hr AS HRESULT
               hr = pIMediaEventEx->lpvtbl->GetEvent(pIMediaEventEx, @lEventCode, @lParam1, @lParam2, 0)
               IF hr <> S_OK THEN EXIT DO
               pIMediaEventEx->lpvtbl->FreeEventParams(pIMediaEventEx, lEventCode, lParam1, lParam2)
               IF lEventCode = EC_COMPLETE THEN
                  IF pIVideoWindow THEN
                     pIVideoWindow->lpvtbl->put_Visible(pIVideoWindow, OAFALSE)
                     pIVideoWindow->lpvtbl->put_Owner(pIVideoWindow, NULL)
                     AfxSafeRelease(pIVideoWindow)
                     pIVideoWindow = NULL
                  END IF
                  AfxSafeRelease(pIMediaControl): pIMediaControl = NULL
                  AfxSafeRelease(pIMediaEventEx) : pIMediaEventEx = NULL
                  AfxSafeRelease(pIGraphBuilder) : pIGraphBuilder = NULL
                  bIsPlaying = FALSE
                  AfxRedrawWindow hwnd
                  EXIT DO
               END IF
            LOOP
         END IF

      CASE WM_SIZE
         ' // Optional resizing code
         IF wParam <> SIZE_MINIMIZED THEN
            ' // Resize the window and the video
            DIM rc AS RECT
            GetClientRect(hwnd, @rc)
            IF pIVideoWindow THEN
               pIVideoWindow->lpvtbl->SetWindowPosition(pIVideoWindow, rc.Left, rc.Top, rc.Right, rc.Bottom)
               RedrawWindow hwnd, @rc, 0, RDW_INVALIDATE OR RDW_UPDATENOW
            END IF
         END IF

      CASE WM_ERASEBKGND
         ' // Erase the window's background
         IF bIsPlaying = FALSE THEN
            DIM hDC AS HDC = cast(HDC, wParam)
            DIM rc AS RECT
            GetClientRect(hwnd, @rc)
            FillRect hDC, @rc, GetStockObject(BLACK_BRUSH)
            FUNCTION = CTRUE
            EXIT FUNCTION
         END IF

    	CASE WM_DESTROY
         ' // Stop de video if it is playing
         IF pIMediaControl THEN
            pIMediaControl->lpvtbl->Stop(pIMediaControl)
            AfxSafeRelease(pIMediaControl)
         END IF
         ' // Clean resources
         IF pIVideoWindow THEN
            pIVideoWindow->lpvtbl->put_Visible(pIVideoWindow, OAFALSE)
            pIVideoWindow->lpvtbl->put_Owner(pIVideoWindow, NULL)
            AfxSafeRelease(pIVideoWindow)
         END IF
         AfxSafeRelease(pIMediaEventEx)
         AfxSafeRelease(pIGraphBuilder)
         ' // Ends the application by sending a WM_QUIT message
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
