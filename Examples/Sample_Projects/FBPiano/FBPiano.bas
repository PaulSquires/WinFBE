' ########################################################################################
' Microsoft Windows
' File: FBPiano.bas
' Contents: FBPiano.
' Based on PBPiano, public domain control originally written by Börje Hagsten in
' June 2016 for PowerBASIC.
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "Afx/CWindow.inc"
#INCLUDE ONCE "CPiano.inc"
USING Afx

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG

   END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT
DECLARE SUB Midi_ListInstruments (BYVAL hCtl AS HWND, BYVAL iPatches AS LONG)

ENUM FBPIANOIDS
   IDC_PIANO = 1001
   IDC_LABEL1      ' octave info
   IDC_LABEL2      ' < "button"
   IDC_LABEL3      ' > "button"
   IDC_LABEL4      ' F1 - help text
   IDC_FRAME1      ' Options
   IDC_FRAME2      ' Patch/Instrument
   IDC_FRAME3      ' Effects
   IDC_FRAME4      ' Pitch bend
   IDC_CHECKBOX1   ' Always on top
   IDC_CHECKBOX2   ' Show Key text
   IDC_CHECKBOX3   ' Show Note text
   IDC_CHECKBOX4   ' Sustain
   IDC_LISTBOX1    ' Instrument listbox
   IDC_COMBOBOX1   ' Channels ComboBox
   IDC_COMBOBOX2   ' Midi devices ComboBox
   IDC_COMBOBOX3   ' Instrument categories ComboBox
   IDC_KEYBOARD    ' Graphic control - piano keys
   IDC_TRACKBAR1   ' Volume
   IDC_TRACKBAR2   ' Balance
   IDC_TRACKBAR3   ' Vibrato
   IDC_TRACKBAR4   ' Pitch bend
END ENUM

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
   DIM hwndMain AS HWND = pWindow.Create(NULL, "FBPiano", @WndProc)
   ' // Change the window style
'   pWindow.WindowStyle = WS_OVERLAPPED OR WS_CAPTION OR WS_SYSMENU
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(646, 410)
   ' // Centers the window
   pWindow.Center

   ' // Create an instance of the piano control
   DIM pPiano AS CPiano = CPiano(@pWindow, IDC_PIANO, 8, 209, 634, 194)
   ' // Options frame
   DIM hCtl AS HWND = pWindow.AddControl("GROUPBOX", hwndMain, IDC_FRAME1, "Options", 8, 3, 195, 198)
   '// List of devices
   hCtl = pWindow.AddControl("COMBOBOX", hwndMain, IDC_COMBOBOX1, "", 17, 21, 177, 160, WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR CBS_DROPDOWNLIST OR CBS_HASSTRINGS)
   DIM moc AS MIDIOUTCAPSW
   FOR i AS LONG = 1 TO midiOutGetNumDevs  ' there's usually only one...
      midiOutGetDevCapsW i - 1, @moc, SIZEOF(moc)
      Combobox_AddString(hCtl, @moc.szPname)
   NEXT
   Combobox_SetCursel(hCtl, 0)
   ' // List of channels
   hCtl = pWindow.AddControl("COMBOBOX", hwndMain, IDC_COMBOBOX2, "", 17, 52, 177, 280, WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR CBS_DROPDOWNLIST OR CBS_HASSTRINGS)
   DIM wszText AS WSTRING * 260
   FOR i AS LONG = 1 TO 16
      wszText = " Channel " & WSTR(i)
      IF i = 10 THEN wszText += "  (Percussion set)"
      Combobox_AddString(hCtl, @wszText)
   NEXT
   Combobox_SetCursel(hCtl, 0)
   ' // Checkboxes
   hCtl = pWindow.AddControl("CHECKBOX", hwndMain, IDC_CHECKBOX1, "Always on top ", 17, 85, 120, 16)
   hCtl = pWindow.AddControl("CHECKBOX", hwndMain, IDC_CHECKBOX2, "Show Key text ", 17, 111, 120, 16)
   Button_SetCheck(hCtl, CTRUE)
   hCtl = pWindow.AddControl("CHECKBOX", hwndMain, IDC_CHECKBOX3, "Show Note text ", 17, 136, 120, 16)
   Button_SetCheck(hCtl, CTRUE)
   ' // Help label
   hCtl = pWindow.AddControl("LABEL", hwndMain, IDC_LABEL4, "Press F1 for help, Esc to Exit", 17, 169, 177, 17, WS_CHILD OR WS_VISIBLE OR SS_CENTER)

   ' // Patches/Instrument frame
   hCtl = pWindow.AddControl("GROUPBOX", hwndMain, IDC_FRAME2, "Patches/Instrument", 213, 3, 160, 198)
   ' // List of patches
   hCtl = pWindow.AddControl("COMBOBOX", hwndMain, IDC_COMBOBOX3, "", 222, 20, 142, 280, WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR CBS_DROPDOWNLIST OR CBS_HASSTRINGS)
   wszText = "Piano" : Combobox_AddString(hCtl, @wszText)
   wszText = "Chromatic percussion" : Combobox_AddString(hCtl, @wszText)
   wszText = "Organ" : Combobox_AddString(hCtl, @wszText)
   wszText = "Guitar" : Combobox_AddString(hCtl, @wszText)
   wszText = "Bass" : Combobox_AddString(hCtl, @wszText)
   wszText = "Strings" : Combobox_AddString(hCtl, @wszText)
   wszText = "Ensemble" : Combobox_AddString(hCtl, @wszText)
   wszText = "Brass" : Combobox_AddString(hCtl, @wszText)
   wszText = "Reed" : Combobox_AddString(hCtl, @wszText)
   wszText = "Pipe" : Combobox_AddString(hCtl, @wszText)
   wszText = "Synth lead" : Combobox_AddString(hCtl, @wszText)
   wszText = "Synth pad" : Combobox_AddString(hCtl, @wszText)
   wszText = "Synth effects" : Combobox_AddString(hCtl, @wszText)
   wszText = "Ethnic" : Combobox_AddString(hCtl, @wszText)
   wszText = "Percussive" : Combobox_AddString(hCtl, @wszText)
   wszText = "Sound effects" : Combobox_AddString(hCtl, @wszText)
   Combobox_SetCursel(hCtl, 0)

   ' // List instruments
   hCtl = pWindow.AddControl("LISTBOX", hwndMain, IDC_LISTBOX1, "", 222, 47, 142, 148, WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR LBS_NOTIFY OR LBS_NOINTEGRALHEIGHT)
   Midi_ListInstruments(hCtl, 0)
   ListBox_SetCursel(hCtl, 0)

   ' // Other controls
   hCtl = pWindow.AddControl("GROUPBOX", hwndMain, IDC_FRAME3, "Effects", 383, 3, 174, 198)
   hCtl = pWindow.AddControl("LABEL", hwndMain, -1, "", 396, 20, 150, 36, WS_CHILD OR WS_VISIBLE OR SS_SUNKEN, WS_EX_STATICEDGE)
   hCtl = pWindow.AddControl("LABEL", hwndMain, IDC_LABEL1, "Octave 4 to 5" & CHR(13, 10) & "(< or > to change)", 416, 21, 110, 34, WS_CHILD OR WS_VISIBLE OR SS_CENTER)
   hCtl = pWindow.AddControl("LABEL", hwndMain, IDC_LABEL2, "<", 397, 21, 20, 34, WS_CHILD OR WS_VISIBLE OR SS_CENTERIMAGE OR SS_NOTIFY OR SS_CENTER, WS_EX_DLGMODALFRAME OR WS_EX_WINDOWEDGE)
   hCtl = pWindow.AddControl("LABEL", hwndMain, IDC_LABEL3, ">", 524, 21, 20, 34, WS_CHILD OR WS_VISIBLE OR SS_CENTERIMAGE OR SS_NOTIFY OR SS_CENTER, WS_EX_DLGMODALFRAME OR WS_EX_WINDOWEDGE)
   hCtl = pWindow.AddControl("LABEL", hwndMain, -1, "Volume:", 393, 72, 45, 16, WS_CHILD OR WS_VISIBLE)
   hCtl = pWindow.AddControl("TRACKBAR", hwndMain, IDC_TRACKBAR1, "", 441, 71, 112, 32, WS_CHILD OR WS_VISIBLE OR TBS_AUTOTICKS OR TBS_BOTTOM OR TBS_TOOLTIPS)
   Trackbar_SetRange(hCtl, 0, 10, CTRUE)
   Trackbar_SetTicFreq(hCtl, 1)
   Trackbar_SetPos(hCtl, CTRUE, 10) ' position
   hCtl = pWindow.AddControl("LABEL", hwndMain, -1, "Balance:", 393, 106, 45, 16)
   hCtl = pWindow.AddControl("TRACKBAR", hwndMain, IDC_TRACKBAR2, "", 441, 104, 112, 32, WS_CHILD OR WS_VISIBLE OR TBS_AUTOTICKS OR TBS_BOTTOM OR TBS_TOOLTIPS)
   Trackbar_SetRange(hCtl, -10, 10, CTRUE)
   Trackbar_SetTicFreq(hCtl, 5)
   Trackbar_SetPos(hCtl, CTRUE, 0) ' balance center
   hCtl = pWindow.AddControl("LABEL", hwndMain, -1, "Vibrato:", 393, 138, 45, 16)
   hCtl = pWindow.AddControl("TRACKBAR", hwndMain, IDC_TRACKBAR3, "", 441, 136, 112, 32, WS_CHILD OR WS_VISIBLE OR TBS_AUTOTICKS OR TBS_BOTTOM OR TBS_TOOLTIPS)
   Trackbar_SetRange(hCtl, 0, 10, CTRUE)
   Trackbar_SetTicFreq(hCtl, 1)
   Trackbar_SetPos(hCtl, CTRUE, 0) ' position
   hCtl = pWindow.AddControl("CHECKBOX", hwndMain, IDC_CHECKBOX4, "Sustain ", 394, 170, 70, 16)
   hCtl = pWindow.AddControl("GROUPBOX", hwndMain, IDC_FRAME1, "Pitch bend", 566, 3, 71, 198)
   hCtl = pWindow.AddControl("TRACKBAR", hwndMain, IDC_TRACKBAR4, "", 584, 25, 45, 170, WS_CHILD OR WS_VISIBLE OR TBS_AUTOTICKS OR TBS_VERT OR TBS_RIGHT OR TBS_BOTH  OR TBS_TOOLTIPS)
   Trackbar_SetRange(hCtl, -64, 64, CTRUE)
   Trackbar_SetTicFreq(hCtl, 10)
   ' // Set the focus in the control
   SetFocus GetDlgItem(pWindow.hWindow, IDC_PIANO)

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

            CASE IDC_LABEL2  ' Arrow left < - octave change down
               IF HIWORD(wParam) = STN_CLICKED OR HIWORD(wParam) = STN_DBLCLK THEN
                  DIM hCtl AS HWND = GetDlgItem(hwnd, IDC_PIANO)
                  DIM pPiano AS CPiano PTR = CAST(CPiano PTR, GetWindowLongPtrW(hCtl, 0))
                  IF pPiano THEN
                     DIM nOctave AS LONG = pPiano->GetOctave
                     nOctave = MAX(0, nOctave - 1)
                     pPiano->SetOctave(nOctave, CTRUE)
                     DIM keyCount AS LONG = pPiano->GetKeyCount
                     DIM wszText AS WSTRING * 260 = "Octave " & WSTR(nOctave) & " to " & _
                         WSTR(nOctave + keyCount \ 12 - 1) & CHR(13, 10) & "(< or > to change)"
                     DIM hLabel AS HWND = GetDlgItem(hwnd, IDC_LABEL1)
                     SetWindowTextW(hLabel, @wszText)
                  END IF
               END IF

            CASE IDC_LABEL3
               IF HIWORD(wParam) = STN_CLICKED OR HIWORD(wParam) = STN_DBLCLK THEN
                  DIM hCtl AS HWND = GetDlgItem(hwnd, IDC_PIANO)
                  DIM pPiano AS CPiano PTR = CAST(CPiano PTR, GetWindowLongPtrW(hCtl, 0))
                  IF pPiano THEN
                     DIM keyCount AS LONG = pPiano->GetKeyCount
                     DIM nOctave AS LONG = pPiano->GetOctave
                     nOctave = MIN(10 - keyCount \ 12, nOctave + 1)
                     pPiano->SetOctave(nOctave, CTRUE)
                     DIM wszText AS WSTRING * 260 = "Octave " & WSTR(nOctave) & " to " & _
                         WSTR(nOctave + keyCount \ 12 - 1) & CHR(13, 10) & "(< or > to change)"
                     DIM hLabel AS HWND = GetDlgItem(hwnd, IDC_LABEL1)
                     SetWindowTextW(hLabel, @wszText)
                  END IF
               END IF

            CASE IDC_CHECKBOX1  ' Topmost or not on screen
               IF HIWORD(wParam) = BN_CLICKED THEN
                  DIM nCheck AS LONG = SendMessageW(CAST(HWND, lParam), BM_GETCHECK, 0, 0)
                  DIM hwndAfter AS HWND
                  IF nCheck THEN hwndAfter = HWND_TOPMOST ELSE hwndAfter = HWND_NOTOPMOST
                  SetWindowPos hwnd, hwndAfter, 0, 0, 0, 0, SWP_NOMOVE OR SWP_NOSIZE
                  SetFocus GetDlgItem(hwnd, IDC_PIANO)
               END IF

            CASE IDC_CHECKBOX2 ' Show key text?
               IF HIWORD(wParam) = BN_CLICKED THEN
                  DIM nCheck AS LONG = SendMessageW(CAST(HWND, lParam), BM_GETCHECK, 0, 0)
                  DIM hCtl AS HWND = GetDlgItem(hwnd, IDC_PIANO)
                  DIM pPiano AS CPiano PTR = CAST(CPiano PTR, GetWindowLongPtrW(hCtl, 0))
                  IF pPiano THEN pPiano->SetKeyText(nCheck, CTRUE)
               END IF

            CASE IDC_CHECKBOX3  ' Show note text?
               IF HIWORD(wParam) = BN_CLICKED THEN
                  DIM nCheck AS LONG = SendMessageW(CAST(HWND, lParam), BM_GETCHECK, 0, 0)
                  DIM hCtl AS HWND = GetDlgItem(hwnd, IDC_PIANO)
                  DIM pPiano AS CPiano PTR = CAST(CPiano PTR, GetWindowLongPtrW(hCtl, 0))
                  IF pPiano THEN pPiano->SetNoteText(nCheck, CTRUE)
               END IF

            CASE IDC_CHECKBOX4  ' Sustain on/off
               IF HIWORD(wParam) = BN_CLICKED THEN
                  DIM nCheck AS LONG = SendMessageW(CAST(HWND, lParam), BM_GETCHECK, 0, 0)
                  DIM hCtl AS HWND = GetDlgItem(hwnd, IDC_PIANO)
                  DIM pPiano AS CPiano PTR = CAST(CPiano PTR, GetWindowLongPtrW(hCtl, 0))
                  IF pPiano THEN pPiano->SetSustain(nCheck, CTRUE)
               END IF

            CASE IDC_COMBOBOX2  ' Channel
               IF HIWORD(wParam) = CBN_SELCHANGE THEN
                  DIM idx AS LONG = ComboBox_GetCursel(CAST(HWND, lParam))
                  DIM hCtl AS HWND = GetDlgItem(hwnd, IDC_PIANO)
                  DIM pPiano AS CPiano PTR = CAST(CPiano PTR, GetWindowLongPtrW(hCtl, 0))
                  IF pPiano THEN pPiano->SetChannel(idx)
               ELSEIF HIWORD(wParam) = CBN_CLOSEUP THEN
                  SetFocus GetDlgItem(hwnd, IDC_PIANO)
               END IF

            CASE IDC_COMBOBOX3  ' Patches
               IF HIWORD(wParam) = CBN_SELCHANGE THEN
                  DIM idx AS LONG = ComboBox_GetCursel(CAST(HWND, lParam))
                  DIM hListBox AS HWND = GetDlgItem(hwnd, IDC_LISTBOX1)
                  Midi_ListInstruments(hListBox, idx)
                  ListBox_SetCursel(hListBox, 0)
                  DIM iPatch AS LONG = ListBox_GetItemData(hListBox, 0)
                  DIM hCtl AS HWND = GetDlgItem(hwnd, IDC_PIANO)
                  DIM pPiano AS CPiano PTR = CAST(CPiano PTR, GetWindowLongPtrW(hCtl, 0))
                  IF pPiano THEN pPiano->SetInstrument(iPatch, CTRUE)
               ELSEIF HIWORD(wParam) = CBN_CLOSEUP THEN
                  SetFocus GetDlgItem(hwnd, IDC_PIANO)
               END IF

            CASE IDC_LISTBOX1  ' Instruments
               IF HIWORD(wParam) = LBN_SELCHANGE THEN
                  DIM idx AS LONG = ListBox_GetCursel(CAST(HWND, lParam))
                  DIM iPatch AS LONG = ListBox_GetItemData(CAST(HWND, lParam), idx)
                  DIM hCtl AS HWND = GetDlgItem(hwnd, IDC_PIANO)
                  DIM pPiano AS CPiano PTR = CAST(CPiano PTR, GetWindowLongPtrW(hCtl, 0))
                  IF pPiano THEN pPiano->SetInstrument(iPatch, CTRUE)
                  ' // Debugging code - Sometimes the octave changes to 0 after setting the instrument
                  DIM octv AS LONG = pPiano->GetOctave
                  DIM keyCount AS LONG = pPiano->GetKeyCount + 1
                  DIM wszText AS WSTRING * 260 = "Octave " & WSTR(octv) & " to " & _
                      WSTR(octv + keyCount \ 12 - 1) & CHR(13, 10) & "(< or > to change)"
                  DIM hLabel AS HWND = GetDlgItem(hwnd, IDC_LABEL1)
                  SetWindowTextW(hLabel, @wszText)
               END IF

            CASE IDC_PIANO
               SELECT CASE HIWORD(wParam)
                  CASE PN_PLAYNOTE   ' when a note is played, notification is sent here
                     ' // Debugging code - Sometimes the octave changes to 0 after playing a note?
                     DIM hCtl AS HWND = GetDlgItem(hwnd, IDC_PIANO)
                     DIM pPiano AS CPiano PTR = CAST(CPiano PTR, GetWindowLongPtrW(hCtl, 0))
                     IF pPiano THEN
                        DIM octv AS LONG = pPiano->GetOctave
                        DIM keyCount AS LONG = pPiano->GetKeyCount + 1
                        DIM wszText AS WSTRING * 260 = "Octave " & WSTR(octv) & " to " & _
                            WSTR(octv + keyCount \ 12 - 1) & CHR(13, 10) & "(< or > to change)"
                        DIM hLabel AS HWND = GetDlgItem(hwnd, IDC_LABEL1)
                        SetWindowTextW(hLabel, @wszText)
                     END IF
                  CASE PN_STOPNOTE   ' when a note is stopped, notification is sent here
                     ' // Debugging code - Sometimes the octave changes to 0 after playing a note?
                     DIM hCtl AS HWND = GetDlgItem(hwnd, IDC_PIANO)
                     DIM pPiano AS CPiano PTR = CAST(CPiano PTR, GetWindowLongPtrW(hCtl, 0))
                     IF pPiano THEN
                        DIM octv AS LONG = pPiano->GetOctave
                        DIM keyCount AS LONG = pPiano->GetKeyCount + 1
                        DIM wszText AS WSTRING * 260 = "Octave " & WSTR(octv) & " to " & _
                            WSTR(octv + keyCount \ 12 - 1) & CHR(13, 10) & "(< or > to change)"
                        DIM hLabel AS HWND = GetDlgItem(hwnd, IDC_LABEL1)
                        SetWindowTextW(hLabel, @wszText)
                     END IF

                  CASE PN_PITCHBEND  ' on up/down arrow key, pitch bend notification is sent here
                     Trackbar_SetPos(GetDlgItem(hwnd, IDC_TRACKBAR4), CTRUE, lParam)
                  CASE PN_OCTAVE     ' on left/right arrow key, octave change notification is sent here
                     DIM hCtl AS HWND = GetDlgItem(hwnd, IDC_PIANO)
                     DIM pPiano AS CPiano PTR = CAST(CPiano PTR, GetWindowLongPtrW(hCtl, 0))
                     IF pPiano THEN
                        DIM keyCount AS LONG = pPiano->GetKeyCount
                        DIM wszText AS WSTRING * 260 = "Octave " & WSTR(lParam) & " to " & _
                            WSTR(lParam + keyCount \ 12 - 1) & CHR(13, 10) & "(< or > to change)"
                        DIM hLabel AS HWND = GetDlgItem(hwnd, IDC_LABEL1)
                        SetWindowTextW(hLabel, @wszText)
                     END IF
               END SELECT

         END SELECT

    	CASE WM_DESTROY
         ' // Ends the application by sending a WM_QUIT message
         PostQuitMessage(0)
         EXIT FUNCTION

      CASE WM_CTLCOLOREDIT, WM_CTLCOLORLISTBOX
         SetBkColor CAST(HDC, wParam), GetSysColor(COLOR_INFOBK)
         SetTextColor CAST(HDC, wParam), GetSysColor(COLOR_INFOTEXT)
         FUNCTION = CAST(LRESULT, GetSysColorBrush(COLOR_INFOBK))
         EXIT FUNCTION

      CASE WM_CTLCOLORSTATIC  ' Simple stunt to avoid focus rect flashing in TrackBars
         SELECT CASE GetDlgCtrlID(CAST(HWND, lParam))
            CASE IDC_TRACKBAR1 TO IDC_TRACKBAR4
               SetFocus GetDlgItem(hwnd, IDC_PIANO)
            CASE IDC_LABEL1
               SetBkColor CAST(HDC, wParam), GetSysColor(COLOR_INFOBK)
               SetTextColor CAST(HDC, wParam), GetSysColor(COLOR_INFOTEXT)
               FUNCTION = CAST(LRESULT, GetSysColorBrush(COLOR_INFOBK))
               EXIT FUNCTION
         END SELECT

      CASE WM_KEYDOWN, WM_KEYUP  ' keyboard keys
         SendMessageW GetDlgItem(hwnd, IDC_PIANO), uMsg, wParam, lParam
         EXIT FUNCTION

      CASE WM_HSCROLL ' Horizontal trackbar scrollpos change
         IF lParam = GetDlgItem(hwnd, IDC_TRACKBAR1) THEN  ' Volume
            IF LOWORD(wParam) = TB_THUMBTRACK THEN      ' get trackbar pos
               DIM nPos AS LONG = Trackbar_GetPos(CAST(HWND, lParam))
               DIM hCtl AS HWND = GetDlgItem(hwnd, IDC_PIANO)
               DIM pPiano AS CPiano PTR = CAST(CPiano PTR, GetWindowLongPtrW(hCtl, 0))
               IF pPiano THEN pPiano->SetVolume(MIN(127, nPos * 13), CTRUE)
            END IF
         ELSEIF lParam = GetDlgItem(hwnd, IDC_TRACKBAR2) THEN ' Balance
            IF LOWORD(wParam) = TB_THUMBTRACK THEN      ' get trackbar pos
               DIM nPos AS LONG = Trackbar_GetPos(CAST(HWND, lParam))
               DIM hCtl AS HWND = GetDlgItem(hwnd, IDC_PIANO)
               DIM pPiano AS CPiano PTR = CAST(CPiano PTR, GetWindowLongPtrW(hCtl, 0))
               IF pPiano THEN pPiano->SetBalance(MIN(127, nPos * 6.4 + 64), CTRUE)
            END IF
         ELSEIF lParam = GetDlgItem(hwnd, IDC_TRACKBAR3) THEN ' Modulation (Vibrato)
            IF LOWORD(wParam) = TB_THUMBTRACK THEN      ' get trackbar pos
               DIM nPos AS LONG = Trackbar_GetPos(CAST(HWND, lParam))
               DIM hCtl AS HWND = GetDlgItem(hwnd, IDC_PIANO)
               DIM pPiano AS CPiano PTR = CAST(CPiano PTR, GetWindowLongPtrW(hCtl, 0))
               IF pPiano THEN pPiano->SetVibrato(MIN(127, nPos * 13), CTRUE)
            END IF
         END IF

      CASE WM_VSCROLL 'Pitch bend - Vertical trackbar scrollpos change
         IF lParam = GetDlgItem(hWnd, IDC_TRACKBAR4) THEN ' Pitch Bend
            SELECT CASE LOWORD(wParam)
               CASE TB_THUMBTRACK ' get trackbar pos and set pitch
                  DIM nPos AS LONG = Trackbar_GetPos(CAST(HWND, lParam))
                  DIM hCtl AS HWND = GetDlgItem(hwnd, IDC_PIANO)
                  DIM pPiano AS CPiano PTR = CAST(CPiano PTR, GetWindowLongPtrW(hCtl, 0))
                  IF pPiano THEN pPiano->SetPitch(64 - nPos, CTRUE)
               CASE TB_ENDTRACK   ' reset trackbar (middle) pos and pitch
                  Trackbar_SetPos(CAST(HWND, lParam), CTRUE, 0)
                  DIM hCtl AS HWND = GetDlgItem(hwnd, IDC_PIANO)
                  DIM pPiano AS CPiano PTR = CAST(CPiano PTR, GetWindowLongPtrW(hCtl, 0))
                  IF pPiano THEN pPiano->SetPitch(64, CTRUE)
            END SELECT
         END IF

      CASE WM_HELP
         STATIC F1Help AS LONG = 0
         IF F1Help = 0 THEN  ' need a static prevention flag there, else
            F1Help = 1      ' pressing F1 again produces a new messagebox
            DIM cwsText AS CWSTR = "PBPiano accepts both computer keyboard and/or mouse play." + CHR(13, 10) + CHR(13, 10) + _
               " Options:" + CHR(13, 10) + _
               "   ComboBox 1 shows available midi devices (usually 1)."  + CHR(13, 10) + _
               "   ComboBox 2 shows a collection of 16 midi channels, but we usually only use channel 1 for normal play."    + CHR(13, 10) + _
               "   Note: Channel 10 is reserved for percussion sets, turning each key from D2# to D7# into an instrument."  + CHR(13, 10) + CHR(13, 10) + _
               "   Always on top - program stays topmost on screen."  + CHR(13, 10) + _
               "   Show Key text - show/hide computer keyboard key text."  + CHR(13, 10) + _
               "   Show Note text - show/hide piano key notes."  + CHR(13, 10) + CHR(13, 10) + _
               " Effects:" + CHR(13, 10) + _
               "   Arrow left - decrease octave"  + CHR(13, 10) + _
               "   Arrow right - increase octave" + CHR(13, 10) + _
               "   Arrow up - pitch bend up"      + CHR(13, 10) + _
               "   Arrow down - pitch bend down"  + CHR(13, 10) + CHR(13, 10) + _
               "   Sustain on/off, for sustain (reverb) on played keys."      + CHR(13, 10) + _
               "   Sustain can cause eternal notes to for example string and brass instruments and should be used with care." + CHR(13, 10) + _
               "   If eternal notes occur, just uncheck this checkbox."       + CHR(13, 10) + _
               "   On octave change, pressed notes are repeated for a nice replay effect." + CHR(13, 10) + _
               "   Mouse controls Volume, Balance, Vibrato and Pitch bend."
            MessageBoxW(hwnd, cwsText, "Information", MB_OK OR MB_ICONINFORMATION)
            F1Help = 0
         END IF

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Fill a ListBox or ComboBox with a list of instruments
' ========================================================================================
SUB Midi_ListInstruments (BYVAL hCtl AS HWND, BYVAL iPatches AS LONG)

   DIM AS LONG iStart, iStop
   SELECT CASE iPatches  ' divide into iPatch categories
      CASE  0 : iStart =   1 : iStop =   8  ' Piano
         ListBox_ResetContent(hCtl)
         DIM wszText AS WSTRING * 260
         FOR i AS LONG = iStart TO iStop
            IF i = 1 THEN wszText = " " & WSTR(i) & "  " & "Acoustic Grand Piano"
            IF i = 2 THEN wszText = " " & WSTR(i) & "  " & "Bright Acoustic Piano"
            IF i = 3 THEN wszText = " " & WSTR(i) & "  " & "Electric Grand Piano"
            IF i = 4 THEN wszText = " " & WSTR(i) & "  " & "Honky-tonk Piano"
            IF i = 5 THEN wszText = " " & WSTR(i) & "  " & "Rhodes Piano"
            IF i = 6 THEN wszText = " " & WSTR(i) & "  " & "Chorused Piano"
            IF i = 7 THEN wszText = " " & WSTR(i) & "  " & "Harpsichord"
            IF i = 8 THEN wszText = " " & WSTR(i) & "  " & "Clavinet"
            ListBox_Addstring(hCtl, @wszText)
            ListBox_SetItemData(hCtl, i - iStart, i - 1)
         NEXT
      CASE  1 : iStart =   9 : iStop =  16  ' Chromatic Percussion
         ListBox_ResetContent(hCtl)
         DIM wszText AS WSTRING * 260
         FOR i AS LONG = iStart TO iStop
            IF i = 9 THEN wszText = "  " & WSTR(i) & "  " & "Celesta"
            IF i = 10 THEN wszText = " " & WSTR(i) & "  " & "Glockenspiel"
            IF i = 11 THEN wszText = " " & WSTR(i) & "  " & "Music Box"
            IF i = 12 THEN wszText = " " & WSTR(i) & "  " & "Vibraphone"
            IF i = 13 THEN wszText = " " & WSTR(i) & "  " & "Marimba"
            IF i = 14 THEN wszText = " " & WSTR(i) & "  " & "Xylophone"
            IF i = 15 THEN wszText = " " & WSTR(i) & "  " & "Tubular Bells"
            IF i = 16 THEN wszText = " " & WSTR(i) & "  " & "Dulcimer"
            ListBox_Addstring(hCtl, @wszText)
            ListBox_SetItemData(hCtl, i - iStart, i - 1)
         NEXT
      CASE  2 : iStart =  17 : iStop =  24  ' Organ
         ListBox_ResetContent(hCtl)
         DIM wszText AS WSTRING * 260
         FOR i AS LONG = iStart TO iStop
            IF i = 17 THEN wszText = " " & WSTR(i) & "  " & "Hammond Organ"
            IF i = 18 THEN wszText = " " & WSTR(i) & "  " & "Percussive Organ"
            IF i = 19 THEN wszText = " " & WSTR(i) & "  " & "Rock Organ"
            IF i = 20 THEN wszText = " " & WSTR(i) & "  " & "Church Organ"
            IF i = 21 THEN wszText = " " & WSTR(i) & "  " & "Reed Organ"
            IF i = 22 THEN wszText = " " & WSTR(i) & "  " & "Accordion"
            IF i = 23 THEN wszText = " " & WSTR(i) & "  " & "Harmonica"
            IF i = 24 THEN wszText = " " & WSTR(i) & "  " & "Tango Accordion"
            ListBox_Addstring(hCtl, @wszText)
            ListBox_SetItemData(hCtl, i - iStart, i - 1)
         NEXT
      CASE  3 : iStart =  25 : iStop =  32  ' Guitar
         ListBox_ResetContent(hCtl)
         DIM wszText AS WSTRING * 260
         FOR i AS LONG = iStart TO iStop
            IF i = 25 THEN wszText = " " & WSTR(i) & "  " & "Acoustic Guitar (nylon)"
            IF i = 26 THEN wszText = " " & WSTR(i) & "  " & "Acoustic Guitar (steel)"
            IF i = 27 THEN wszText = " " & WSTR(i) & "  " & "Electric Guitar (jazz)"
            IF i = 28 THEN wszText = " " & WSTR(i) & "  " & "Electric Guitar (clean)"
            IF i = 29 THEN wszText = " " & WSTR(i) & "  " & "Electric Guitar (muted)"
            IF i = 30 THEN wszText = " " & WSTR(i) & "  " & "Overdriven Guitar"
            IF i = 31 THEN wszText = " " & WSTR(i) & "  " & "Distortion Guitar"
            IF i = 32 THEN wszText = " " & WSTR(i) & "  " & "Guitar Harmonics"
            ListBox_Addstring(hCtl, @wszText)
            ListBox_SetItemData(hCtl, i - iStart, i - 1)
         NEXT
      CASE  4 : iStart =  33 : iStop =  40  ' Bass
         ListBox_ResetContent(hCtl)
         DIM wszText AS WSTRING * 260
         FOR i AS LONG = iStart TO iStop
            IF i = 33 THEN wszText = " " & WSTR(i) & "  " & "Acoustic Bass"
            IF i = 34 THEN wszText = " " & WSTR(i) & "  " & "Electric Bass (finger)"
            IF i = 35 THEN wszText = " " & WSTR(i) & "  " & "Electric Bass (pick)"
            IF i = 36 THEN wszText = " " & WSTR(i) & "  " & "Fretless Bass"
            IF i = 37 THEN wszText = " " & WSTR(i) & "  " & "Slap Bass 1"
            IF i = 38 THEN wszText = " " & WSTR(i) & "  " & "Slap Bass 2"
            IF i = 39 THEN wszText = " " & WSTR(i) & "  " & "Synth Bass 1"
            IF i = 40 THEN wszText = " " & WSTR(i) & "  " & "Synth Bass 2"
            ListBox_Addstring(hCtl, @wszText)
            ListBox_SetItemData(hCtl, i - iStart, i - 1)
         NEXT
      CASE  5 : iStart =  41 : iStop =  48  ' Strings
         ListBox_ResetContent(hCtl)
         DIM wszText AS WSTRING * 260
         FOR i AS LONG = iStart TO iStop
            IF i = 41 THEN wszText = " " & WSTR(i) & "  " & "Violin"
            IF i = 42 THEN wszText = " " & WSTR(i) & "  " & "Viola"
            IF i = 43 THEN wszText = " " & WSTR(i) & "  " & "Cello"
            IF i = 44 THEN wszText = " " & WSTR(i) & "  " & "Contrabass"
            IF i = 45 THEN wszText = " " & WSTR(i) & "  " & "Tremolo Strings"
            IF i = 46 THEN wszText = " " & WSTR(i) & "  " & "Pizzicato Strings"
            IF i = 47 THEN wszText = " " & WSTR(i) & "  " & "Orchestral Harp"
            IF i = 48 THEN wszText = " " & WSTR(i) & "  " & "Timpani"
            ListBox_Addstring(hCtl, @wszText)
            ListBox_SetItemData(hCtl, i - iStart, i - 1)
         NEXT
      CASE  6 : iStart =  49 : iStop =  56  ' Ensamble
         ListBox_ResetContent(hCtl)
         DIM wszText AS WSTRING * 260
         FOR i AS LONG = iStart TO iStop
            IF i = 49 THEN wszText = " " & WSTR(i) & "  " & "String Ensemble 1"
            IF i = 50 THEN wszText = " " & WSTR(i) & "  " & "String Ensemble 2"
            IF i = 51 THEN wszText = " " & WSTR(i) & "  " & "SynthStrings 1"
            IF i = 52 THEN wszText = " " & WSTR(i) & "  " & "SynthStrings 2"
            IF i = 53 THEN wszText = " " & WSTR(i) & "  " & "Choir Aahs"
            IF i = 54 THEN wszText = " " & WSTR(i) & "  " & "Voice Oohs"
            IF i = 55 THEN wszText = " " & WSTR(i) & "  " & "Synth Voice"
            IF i = 56 THEN wszText = " " & WSTR(i) & "  " & "Orchestra Hit"
            ListBox_Addstring(hCtl, @wszText)
            ListBox_SetItemData(hCtl, i - iStart, i - 1)
         NEXT
      CASE  7 : iStart =  57 : iStop =  64  ' Brass
         ListBox_ResetContent(hCtl)
         DIM wszText AS WSTRING * 260
         FOR i AS LONG = iStart TO iStop
            IF i = 57 THEN wszText = " " & WSTR(i) & "  " & "Trumpet"
            IF i = 58 THEN wszText = " " & WSTR(i) & "  " & "Trombone"
            IF i = 59 THEN wszText = " " & WSTR(i) & "  " & "Tuba"
            IF i = 60 THEN wszText = " " & WSTR(i) & "  " & "Muted Trumpet"
            IF i = 61 THEN wszText = " " & WSTR(i) & "  " & "French Horn"
            IF i = 62 THEN wszText = " " & WSTR(i) & "  " & "Brass Section"
            IF i = 63 THEN wszText = " " & WSTR(i) & "  " & "Synth Brass 1"
            IF i = 64 THEN wszText = " " & WSTR(i) & "  " & "Synth Brass 2"
            ListBox_Addstring(hCtl, @wszText)
            ListBox_SetItemData(hCtl, i - iStart, i - 1)
         NEXT
      CASE  8 : iStart =  65 : iStop =  72  ' Reed
         ListBox_ResetContent(hCtl)
         DIM wszText AS WSTRING * 260
         FOR i AS LONG = iStart TO iStop
            IF i = 65 THEN wszText = " " & WSTR(i) & "  " & "Soprano Sax"
            IF i = 66 THEN wszText = " " & WSTR(i) & "  " & "Alto Sax"
            IF i = 67 THEN wszText = " " & WSTR(i) & "  " & "Tenor Sax"
            IF i = 68 THEN wszText = " " & WSTR(i) & "  " & "Baritone Sax"
            IF i = 69 THEN wszText = " " & WSTR(i) & "  " & "Oboe"
            IF i = 70 THEN wszText = " " & WSTR(i) & "  " & "English Horn"
            IF i = 71 THEN wszText = " " & WSTR(i) & "  " & "Bassoon"
            IF i = 72 THEN wszText = " " & WSTR(i) & "  " & "Clarinet"
            ListBox_Addstring(hCtl, @wszText)
            ListBox_SetItemData(hCtl, i - iStart, i - 1)
         NEXT
      CASE  9 : iStart =  73 : iStop =  80  ' Pipe
         ListBox_ResetContent(hCtl)
         DIM wszText AS WSTRING * 260
         FOR i AS LONG = iStart TO iStop
            IF i = 73 THEN wszText = " " & WSTR(i) & "  " & "Piccolo"
            IF i = 74 THEN wszText = " " & WSTR(i) & "  " & "Flute"
            IF i = 75 THEN wszText = " " & WSTR(i) & "  " & "Recorder"
            IF i = 76 THEN wszText = " " & WSTR(i) & "  " & "Pan Flute"
            IF i = 77 THEN wszText = " " & WSTR(i) & "  " & "Bottle Blow"
            IF i = 78 THEN wszText = " " & WSTR(i) & "  " & "Shakuhachi"
            IF i = 79 THEN wszText = " " & WSTR(i) & "  " & "Whistle"
            IF i = 80 THEN wszText = " " & WSTR(i) & "  " & "Ocarina"
            ListBox_Addstring(hCtl, @wszText)
            ListBox_SetItemData(hCtl, i - iStart, i - 1)
         NEXT
      CASE 10 : iStart =  81 : iStop =  88  ' Synth Lead
         ListBox_ResetContent(hCtl)
         DIM wszText AS WSTRING * 260
         FOR i AS LONG = iStart TO iStop
            IF i = 81 THEN wszText = " " & WSTR(i) & "  " & "Lead 1 (square)"
            IF i = 82 THEN wszText = " " & WSTR(i) & "  " & "Lead 2 (sawtooth)"
            IF i = 83 THEN wszText = " " & WSTR(i) & "  " & "Lead 3 (calliope lead)"
            IF i = 84 THEN wszText = " " & WSTR(i) & "  " & "Lead 4 (chiff lead)"
            IF i = 85 THEN wszText = " " & WSTR(i) & "  " & "Lead 5 (charang)"
            IF i = 86 THEN wszText = " " & WSTR(i) & "  " & "Lead 6 (voice)"
            IF i = 87 THEN wszText = " " & WSTR(i) & "  " & "Lead 7 (fifths)"
            IF i = 88 THEN wszText = " " & WSTR(i) & "  " & "Lead 8 (bass + lead)"
            ListBox_Addstring(hCtl, @wszText)
            ListBox_SetItemData(hCtl, i - iStart, i - 1)
         NEXT
      CASE 11 : iStart =  89 : iStop =  96  ' Synth Pad
         ListBox_ResetContent(hCtl)
         DIM wszText AS WSTRING * 260
         FOR i AS LONG = iStart TO iStop
            IF i = 89 THEN wszText = " " & WSTR(i) & "  " & "Pad 1 (new age)"
            IF i = 90 THEN wszText = " " & WSTR(i) & "  " & "Pad 2 (warm)"
            IF i = 91 THEN wszText = " " & WSTR(i) & "  " & "Pad 3 (polysynth)"
            IF i = 92 THEN wszText = " " & WSTR(i) & "  " & "Pad 4 (choir)"
            IF i = 93 THEN wszText = " " & WSTR(i) & "  " & "Pad 5 (bowed)"
            IF i = 94 THEN wszText = " " & WSTR(i) & "  " & "Pad 6 (metallic)"
            IF i = 95 THEN wszText = " " & WSTR(i) & "  " & "Pad 7 (halo)"
            IF i = 96 THEN wszText = " " & WSTR(i) & "  " & "Pad 8 (sweep)"
            ListBox_Addstring(hCtl, @wszText)
            ListBox_SetItemData(hCtl, i - iStart, i - 1)
         NEXT
      CASE 12 : iStart =  97 : iStop = 104  ' Synth Effects
         ListBox_ResetContent(hCtl)
         DIM wszText AS WSTRING * 260
         FOR i AS LONG = iStart TO iStop
            IF i = 97 THEN wszText = "   " & WSTR(i) & "  " & "FX 1 (rain)"
            IF i = 98 THEN wszText = "   " & WSTR(i) & "  " & "FX 2 (soundtrack)"
            IF i = 99 THEN wszText = "   " & WSTR(i) & "  " & "FX 3 (crystal)"
            IF i = 100 THEN wszText = " " & WSTR(i) & "  " & "FX 4 (atmosphere)"
            IF i = 101 THEN wszText = " " & WSTR(i) & "  " & "FX 5 (brightness)"
            IF i = 102 THEN wszText = " " & WSTR(i) & "  " & "FX 6 (goblins)"
            IF i = 103 THEN wszText = " " & WSTR(i) & "  " & "FX 7 (echoes)"
            IF i = 104 THEN wszText = " " & WSTR(i) & "  " & "FX 8 (sci-fi)"
            ListBox_Addstring(hCtl, @wszText)
            ListBox_SetItemData(hCtl, i - iStart, i - 1)
         NEXT
      CASE 13 : iStart = 105 : iStop = 112  ' Ethnic
         ListBox_ResetContent(hCtl)
         DIM wszText AS WSTRING * 260
         FOR i AS LONG = iStart TO iStop
            IF i = 105 THEN wszText = " " & WSTR(i) & "  " & "Sitar"
            IF i = 106 THEN wszText = " " & WSTR(i) & "  " & "Banjo"
            IF i = 107 THEN wszText = " " & WSTR(i) & "  " & "Shamisen"
            IF i = 108 THEN wszText = " " & WSTR(i) & "  " & "Koto"
            IF i = 109 THEN wszText = " " & WSTR(i) & "  " & "Kalimba"
            IF i = 110 THEN wszText = " " & WSTR(i) & "  " & "Bagpipe"
            IF i = 111 THEN wszText = " " & WSTR(i) & "  " & "Fiddle"
            IF i = 112 THEN wszText = " " & WSTR(i) & "  " & "Shanai"
            ListBox_Addstring(hCtl, @wszText)
            ListBox_SetItemData(hCtl, i - iStart, i - 1)
         NEXT
      CASE 14 : iStart = 113 : iStop = 119  ' Percussive
         ListBox_ResetContent(hCtl)
         DIM wszText AS WSTRING * 260
         FOR i AS LONG = iStart TO iStop
            IF i = 113 THEN wszText = " " & WSTR(i) & "  " & "Tinkle Bell"
            IF i = 114 THEN wszText = " " & WSTR(i) & "  " & "Agogo"
            IF i = 115 THEN wszText = " " & WSTR(i) & "  " & "Steel Drums"
            IF i = 116 THEN wszText = " " & WSTR(i) & "  " & "Woodblock"
            IF i = 117 THEN wszText = " " & WSTR(i) & "  " & "Taiko Drum"
            IF i = 118 THEN wszText = " " & WSTR(i) & "  " & "Melodic Tom"
            IF i = 119 THEN wszText = " " & WSTR(i) & "  " & "Synth Drum"
            ListBox_Addstring(hCtl, @wszText)
            ListBox_SetItemData(hCtl, i - iStart, i - 1)
         NEXT
      CASE 15 : iStart = 120 : iStop = 128  ' Sound effects
         ListBox_ResetContent(hCtl)
         DIM wszText AS WSTRING * 260
         FOR i AS LONG = iStart TO iStop
            IF i = 120 THEN wszText = " " & WSTR(i) & "  " & "Reverse Cymbal"
            IF i = 121 THEN wszText = " " & WSTR(i) & "  " & "Guitar Fret Noise"
            IF i = 122 THEN wszText = " " & WSTR(i) & "  " & "Breath Noise"
            IF i = 123 THEN wszText = " " & WSTR(i) & "  " & "Seashore"
            IF i = 124 THEN wszText = " " & WSTR(i) & "  " & "Bird Tweet"
            IF i = 125 THEN wszText = " " & WSTR(i) & "  " & "Telephone Ring"
            IF i = 126 THEN wszText = " " & WSTR(i) & "  " & "Helicopter"
            IF i = 127 THEN wszText = " " & WSTR(i) & "  " & "Applause"
            IF i = 128 THEN wszText = " " & WSTR(i) & "  " & "Gunshot"
            ListBox_Addstring(hCtl, @wszText)
            ListBox_SetItemData(hCtl, i - iStart, i - 1)
         NEXT

  END SELECT
  
END SUB
