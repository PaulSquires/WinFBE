' ########################################################################################
' Microsoft Windows
' Contents: SAPI example
' Creates an instance of the ISpVoice interface, sets the the object of interest to word
' boundaries and sets the handle of the window that will receive the user-defined
' MSG_SAPI_EVENT message. When processing the MSG_SAPI_EVENT message, it calls the
' GetEvents method of the ISpVoice interface to retrieve the word boundaries and highlights
' the word in the edit control.
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "Afx/CWindow.inc"
#INCLUDE ONCE "Afx/AfxSapi.bi"
USING Afx

CONST IDC_START = 1001
CONST IDC_STOP = 1002
CONST IDC_TEXTBOX = 1003
CONST MSG_SAPI_EVENT = WM_USER + 1   ' --> change me

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG

   END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

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
   AfxSetProcessDPIAware

   ' // Initialize the COM library
   CoInitialize NULL

   ' // Creates the main window
   DIM pWindow AS CWindow
   ' -or- DIM pWindow AS CWindow = "MyClassName" (use the name that you wish)
   pWindow.Create(NULL, "SAPI example", @WndProc)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(340, 200)
   ' // Centers the window
   pWindow.Center

   ' // Adds a button
   pWindow.AddControl("Button", , IDC_START, "Start", 130, 160, 75, 23)
   pWindow.AddControl("Button", , IDC_STOP, "Stop", 230, 160, 75, 23)

   ' // Displays the window and dispatches the Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

   ' // Uninitialize the COM library
   CoUninitialize

END FUNCTION
' ========================================================================================

' ========================================================================================
' Main window procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   STATIC pSpVoice AS Afx_ISpVoice PTR, cwsText AS CWSTR

   SELECT CASE uMsg

      CASE WM_CREATE
         ' // Create an instance of the SpVoice object
         pSpVoice = AfxNewCom("SAPI.SpVoice")
         IF pSpVoice THEN
            ' // Set the object of interest to word boundaries
            pSpVoice->SetInterest(SPFEI(SPEI_WORD_BOUNDARY), SPFEI(SPEI_WORD_BOUNDARY))
            ' // Set the handle of the window that will receive the MSG_SAPI_EVENT message
            pSpVoice->SetNotifyWindowMessage(hwnd, MSG_SAPI_EVENT, 0, 0)
         END IF
         cwsText = "Now is the time for all good men to read" & _
            " this sentence and head to the enlistment" & _
            " center to help their country fight for justice!"
         DIM pWindow AS CWindow PTR = AfxCWindowPtr(lParam)
         IF pWindow THEN
            pWindow->AddControl("EditMultiline", hwnd, IDC_TEXTBOX, cwsText, 20, 20, 300, 120, _
               WS_VISIBLE OR WS_TABSTOP OR WS_VSCROLL OR ES_LEFT OR ES_MULTILINE OR ES_NOHIDESEL OR ES_WANTRETURN)
         END IF
         EXIT FUNCTION

      CASE WM_COMMAND
         SELECT CASE GET_WM_COMMAND_ID(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF
            CASE IDC_START
               ' // Start speaking
               IF pSpVoice THEN pSpVoice->Speak(cwsText, SPF_ASYNC, NULL)
               EXIT FUNCTION
            CASE IDC_STOP
               ' // Stop speaking
               DIM cws AS CWSTR = ""
               IF pSpVoice THEN pSpVoice->Speak(cws, SVSFPurgeBeforeSpeak, NULL)
               EXIT FUNCTION
         END SELECT

      CASE MSG_SAPI_EVENT
         DIM eventItem AS SPEVENT, eventStatus AS SPVOICESTATUS
         DO
            IF pSpVoice->GetEvents(1, @eventItem, NULL) <> S_OK THEN EXIT DO
            IF eventItem.eEventId = SPEI_WORD_BOUNDARY THEN
               pSpVoice->GetStatus(@eventStatus, NULL)
               DIM nStart AS LONG = eventStatus.ulInputWordPos
               DIM nLen AS LONG = eventStatus.ulInputWordLen
               Edit_SetSel(GetDlgItem(hwnd, IDC_TEXTBOX), nStart, nStart + nLen)
            END IF
         LOOP
    	CASE WM_DESTROY
          ' // Release the ISpVoice interface
         AfxSafeRelease(pSpVoice)
         ' // Ends the application by sending a WM_QUIT message
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
