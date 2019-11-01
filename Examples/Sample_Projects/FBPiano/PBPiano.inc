' PBPiano.inc - Public Domain by Borje Hagsten, June, 2016.
' Compiles with both PBWIN 9 and 10. See PBPiano.bas and
' PBPianoDemo.bas for examples of how to use this custom control
' include file.
'
' To begin with, computer keyboards are lousy midi keyboards. Some
' key combinations when playing chords simply don't work, probably
' because of bit mask conflicts. But it's fun to learn new stuff
' and investigate, which is why I did this one. Pitch bending and
' vibrato were the hardest part to do, because I could not find any
' examples or good info how to do them by code, so had to experiment
' and finally found a way. Should be easy enough to understand how
' to use all features and learn limitations. Both mouse and keyboard
' can be used to play at same time.
' 
' As always, code is free to use and abuse any way you want. Could
' serve as base for a note learning game or a sequenser, build own
' custom signals or just use it as it is, for fun.


'====================================================================
' PBPiano.inc, PB Midi Piano custom control.
'--------------------------------------------------------------------
#IF NOT %DEF(%WINAPI)
  #INCLUDE "WIN32API.INC"
#ENDIF

'--------------------------------------------------------------------
'  How to include and call in parent:
' in declares:
'     #INCLUDE "PBPiano.inc"

' in code:
'  hPiano = CreatePiano(hDlg, %IDC_PIANO, "", x, y, w, h, _
'                       %WS_CHILD OR %WS_TABSTOP OR %WS_VISIBLE, 0, 1)
'
'--------------------------------------------------------------------
' Messages for use with SendMessage or CONTROL SEND..
'--------------------------------------------------------------------
%PM_SETDEVICE       = %WM_APP + 100  ' wParam sets midi device
%PM_SETCHANNEL      = %WM_APP + 101  ' wParam sets channel
%PM_PLAYNOTE        = %WM_APP + 105  ' wParam sets note to play (0 to 127), lParam sets sound volume (0 to 127)
%PM_STOPNOTE        = %WM_APP + 106  ' wParam determines which note to stop playing (0 to 127)

%PM_LISTDEVICES     = %WM_APP + 110  ' List midi devices. wParam is list id
%PM_LISTCHANNELS    = %WM_APP + 111  ' List 16 midi channels. wParam list id
%PM_LISTPATCHES     = %WM_APP + 112  ' List patches. wParam is list id
%PM_LISTINSTRUMENTS = %WM_APP + 113  ' List instrument. wParam is list id

' For the following %PM_SET.. messages, lParam (1/0) sets control focus
%PM_SETINSTRUMENT   = %WM_APP + 120  ' wParam sets instrument (0-127)
%PM_SETOCTAVE       = %WM_APP + 121  ' wParam sets octave (0-10)
%PM_SETKEYCOUNT     = %WM_APP + 122  ' wParam sets total number of keys (1 to 128)
%PM_SETPITCH        = %WM_APP + 123  ' wParam sets pitch bend (-64 to 64, where 0 is normal pitch)
%PM_SETVOLUME       = %WM_APP + 124  ' wParam sets sound volume (0 to 127)
%PM_SETBALANCE      = %WM_APP + 125  ' wParam sets sound balance (-64 to 64, where 0 is centered)
%PM_SETVIBRATO      = %WM_APP + 126  ' wParam sets vibrato/modulation (0 to 127)
%PM_SETSUSTAIN      = %WM_APP + 127  ' wParam sets sustain off/on (0/127)
%PM_SETKEYTEXT      = %WM_APP + 130  ' wParam 0/1, if to print computer key letters
%PM_SETNOTETEXT     = %WM_APP + 131  ' wParam 0/1, if to print note desciptions

%PM_GETOCTAVE       = %WM_APP + 221  ' get base octave
%PM_GETKEYCOUNT     = %WM_APP + 222  ' get number of keys (1 to 128)

%PN_PLAYNOTE        = %WM_APP + 300  ' note played notification
%PN_STOPNOTE        = %WM_APP + 301  ' note stopped notification
%PN_OCTAVE          = %WM_APP + 310  ' octave change notification
%PN_PITCHBEND       = %WM_APP + 320  ' pitch bend notification

'--------------------------------------------------------------------
TYPE kbData            ' Program data
  hParent    AS DWORD  ' Parent dialog handle
  CtlId      AS DWORD  ' Piano control id
  hFont      AS DWORD  ' font handle
  hMidi      AS DWORD  ' midi device
  memDC      AS DWORD  ' memory device context
  hBitmap    AS DWORD  ' 
  hBitOld    AS DWORD  ' 
  KeyCount   AS DWORD  ' Number of keys to draw
  instrument AS LONG   ' selected instrument (0 - 127)
  chPlay     AS LONG   ' channel &H90 to &H9F, where ch &H99 = percussion
  chStop     AS LONG   ' channel &H80 to &H8F, where ch &H89 = percussion
  chControl  AS LONG   ' channel &HB0 to &HBF, Control mode change
  chInstr    AS LONG   ' channel &HC0 to &HCF, Program (instrument) change
  chBend     AS LONG   ' channel &HE0 to &HEF, Pitch bend
  KeyText    AS LONG   ' Key description on keys on/off
  NoteText   AS LONG   ' Note description on keys on/off
  note       AS LONG   ' note to play (0 - 127)
  octave     AS LONG   ' base octave (1 octave = 12 notes)
  sustain    AS LONG   ' sustain on/off
  volume     AS LONG   ' sound volume
  balance    AS LONG   ' sound balance Left/Right
  vibrato    AS LONG   ' Modulation/Vibrato
END TYPE

TYPE pKey              ' Key array data
  nt  AS LONG          ' note (0 - 127)
  dn  AS LONG          ' key up/down flag, 0/1
  rc  AS RECT          ' key coordinates
END TYPE

$KEY1 = "ZSXDCVGBHNJMQ2W3ER5T6Y7UI9O0P"  ' 29 computer key note layout

'====================================================================
' Initialize and create the control in one and same call
'--------------------------------------------------------------------
FUNCTION CreatePiano(BYVAL hParent AS DWORD, BYVAL CtrlId AS LONG, _
                     BYVAL sTxt AS STRING, _
                     BYVAL cLeft AS LONG,  BYVAL cTop AS LONG, _
                     BYVAL cWidth AS LONG, BYVAL cHeight AS LONG, _
                     BYVAL wStyle AS LONG, BYVAL wStyleEx AS LONG, _
                     BYVAL isDDT AS LONG) AS DWORD
'--------------------------------------------------------------------
  LOCAL hCtrl AS DWORD, wc AS WNDCLASSEX, szName AS ASCIIZ * 20

  szName = "PBPianoCtrl"

  wc.cbSize        = LEN(wc)
  wc.Style         = %CS_DBLCLKS
  wc.lpfnWndProc   = CODEPTR(PBPianoProc)
  wc.cbClsExtra    = 0
  wc.cbWndExtra    = 4  ' 4 extra bytes 
  wc.hInstance     = GetWindowLong(hParent, %GWL_HINSTANCE)
  wc.hIcon         = 0
  wc.hCursor       = LoadCursor(%NULL, BYVAL %IDC_HAND)
  wc.hbrBackground = %NULL
  wc.lpszMenuName  = %NULL
  wc.lpszClassName = VARPTR(szName)
  wc.hIconSm       = 0

  RegisterClassEx wc

'--------------------------------------------------------------------
  IF isDDT THEN ' create control using DDT dialog way
     CONTROL ADD szName, hParent, CtrlId, sTxt, _
                 cLeft, cTop, cWidth, cHeight, wStyle, wStyleEx
     CONTROL HANDLE hParent, CtrlId TO hCtrl

  ELSE          ' create control using SDK/API way
     hCtrl = CreateWindowEx(wStyleEx, szName, BYVAL STRPTR(sTxt), wStyle, _
                            cLeft, cTop, cWidth, cHeight, _
                            hParent, CtrlId, wc.hInstance, BYVAL 0)
  END IF
'--------------------------------------------------------------------

  FUNCTION = hCtrl
END FUNCTION


'====================================================================
' Piano Control Procedure
'--------------------------------------------------------------------
FUNCTION PBPianoProc (BYVAL hWnd AS DWORD,   BYVAL wMsg AS DWORD, _
                      BYVAL wParam AS DWORD, BYVAL lParam AS LONG) AS LONG

  LOCAL ps AS PAINTSTRUCT, pt AS POINTAPI, rc AS RECT
  LOCAL c, d AS LONG, hDC AS DWORD

  STATIC chPitch, CurKey, OldKey AS LONG
  DIM hRgn(127) AS STATIC DWORD ' static key region array
  DIM pk(127) AS STATIC pKey    ' static key data array
  STATIC kbd AS kbData          ' static keyboard data

  SELECT CASE wMsg
  CASE %WM_CREATE
      IF midiOutOpen(kbd.hMidi, -1, 0, 0, 0) <> %MMSYSERR_NOERROR THEN
          MessageBox (GetParent(hWnd), "midiOutOpen failed!", "Error message!", _
                      %MB_OK OR %MB_ICONERROR)
          EXIT FUNCTION
      END IF

      initKbd(hWnd, kbd)
      DrawKeyBoard(hWnd, kbd, pk(), hRgn())
      '--------------------------------------------------------------
      chPitch = 64        ' pitch bend start value (64 is no bend)
      CurKey  = -1

  CASE %WM_DESTROY       ' exit, time to clean up memory, etc
      IF kbd.hMidi  THEN midiOutClose(kbd.hMidi) ' Close the midi device
      IF kbd.memDC THEN  ' restore and delete what we have created
          IF kbd.hBitOld THEN SelectObject kbd.memDC, kbd.hBitOld
          IF kbd.hBitmap THEN DeleteObject kbd.hBitmap
          DeleteDC kbd.memDC
      END IF
      IF kbd.hFont THEN DeleteObject kbd.hFont
      FOR c = LBOUND(hRgn) TO UBOUND(hRgn)
          IF hRgn(c) THEN DeleteObject(hRgn(c))
      NEXT
      EXIT FUNCTION

  CASE %WM_GETDLGCODE ' Ensure the control processes all keys by itself
      FUNCTION = %DLGC_WANTALLKEYS
      EXIT FUNCTION

  CASE %WM_PAINT  ' on repaint, simply copy kbd.memDC to control DC
      BeginPaint hWnd, ps
        GetClientRect hWnd, rc
        BitBlt ps.hDC, 0, 0, rc.nRight, rc.nBottom, kbd.memDC, 0, 0, %SRCCOPY
      EndPaint hWnd, ps
      EXIT FUNCTION

  CASE %WM_KEYDOWN
      SELECT CASE AS LONG wParam
      CASE %VK_ESCAPE ' enable Exit on Esc key
           SendMessage(kbd.hParent, %WM_COMMAND, MAK(DWORD, %IDCANCEL, %BN_CLICKED), 0)

      CASE %VK_UP  ' Arrow up ^  - Pitch bend up
          IF chPitch = 64 THEN  ' send notification on each change to parent's/control id's %WM_COMMAND
              FOR chPitch = 64 TO 128
                  Midi_PitchBend(kbd.hMidi, kbd.chBend, chPitch)
                  SendMessage(kbd.hParent, %WM_COMMAND, MAK(DWORD, kbd.CtlId, %PN_PITCHBEND), -(chPitch - 64))
                  IF (GetAsyncKeyState(%VK_UP) AND &H8000) = 0 THEN EXIT FOR ' if key is released
                  SLEEP 8
              NEXT
          END IF

      CASE %VK_DOWN  ' Arrow down v  - Pitch bend down
          IF chPitch = 64 THEN
              FOR chPitch = 64 TO 0 STEP -1
                  Midi_PitchBend(kbd.hMidi, kbd.chBend, chPitch)
                  SendMessage(kbd.hParent, %WM_COMMAND, MAK(DWORD, kbd.CtlId, %PN_PITCHBEND), -(chPitch - 64))
                  IF (GetAsyncKeyState(%VK_DOWN) AND &H8000) = 0 THEN EXIT FOR
                  SLEEP 8
              NEXT
          END IF

      CASE %VK_LEFT  ' Arrow left < - octave change down
          kbd.octave = MAX&(0, kbd.octave - 1)
          SendMessage(hWnd, %PM_SETOCTAVE, kbd.octave, 1)
          SendMessage(kbd.hParent, %WM_COMMAND, MAK(DWORD, kbd.CtlId, %PN_OCTAVE), kbd.octave)

      CASE %VK_RIGHT  ' Arrow right > - octave change up
          kbd.octave = MIN&(10 - kbd.KeyCount \ 12, kbd.octave + 1)
          SendMessage(hWnd, %PM_SETOCTAVE, kbd.octave, 1)
          SendMessage(kbd.hParent, %WM_COMMAND, MAK(DWORD, kbd.CtlId, %PN_OCTAVE), kbd.octave)

      CASE %VK_0 TO %VK_Z                 ' See if a note key has been pressed
          c = INSTR($KEY1, CHR$(wParam))  ' $KEY1 is used for pk array compare
          IF c THEN
              DECR c
              IF pk(c).dn = 0 THEN        ' if not pressed before
                  kbd.note = kbd.octave * 12 + pk(c).nt   ' actual note
                  pk(c).dn = 2                            ' 2 = keyboard key pressed it
                  SendMessage(kbd.hParent, %WM_COMMAND, MAK(DWORD, kbd.CtlId, %PN_PLAYNOTE), kbd.note)
                  Midi_PlayNote(kbd.hMidi, kbd.chPlay, kbd.note, kbd.volume)

                  SELECT CASE c MOD 12
                  CASE 1, 3, 6, 8, 10  ' black keys
                      PrintKeyText(kbd, pk(c), IIF(kbd.KeyText, %GRAY, %BLACK), %BLACK, IIF(kbd.NoteText, %YELLOW, %BLACK), %BLACK)
                  CASE ELSE    ' white keys
                      PrintKeyText(kbd, pk(c), IIF(kbd.KeyText, %LTGRAY, %WHITE), %WHITE, IIF(kbd.NoteText, %BLUE, %WHITE), %WHITE)
                  END SELECT
                  InvalidateRect hWnd, BYVAL 0, 0 : UpdateWindow hWnd
              END IF
          END IF
      END SELECT
      EXIT FUNCTION

  CASE %WM_KEYUP
      SELECT CASE AS LONG wParam
      CASE %VK_UP    ' Arrow up - Pitch has been bended up
          IF chPitch > 64 THEN
              chPitch = 64
              Midi_PitchBend(kbd.hMidi, kbd.chBend, chPitch)
              SendMessage(kbd.hParent, %WM_COMMAND, MAK(DWORD, kbd.CtlId, %PN_PITCHBEND), 0)
          END IF

      CASE %VK_DOWN  ' Arrow down - Pitch has been bended down
          IF chPitch < 64 THEN
              chPitch = 64
              Midi_PitchBend(kbd.hMidi, kbd.chBend, chPitch)
              SendMessage(kbd.hParent, %WM_COMMAND, MAK(DWORD, kbd.CtlId, %PN_PITCHBEND), 0)
          END IF

      CASE %VK_0 TO %VK_Z  ' same as on key down, but here we stop playing note and reset key text instead.
          c = INSTR($KEY1, CHR$(wParam))  ' see if it's a defined key
          IF c THEN
              DECR c
              IF pk(c).dn = 2 THEN            ' if it was pressed by a keyboard key
                  SendMessage(kbd.hParent, %WM_COMMAND, MAK(DWORD, kbd.CtlId, %PN_STOPNOTE), kbd.note)
                  Midi_StopNote(kbd.hMidi, kbd.chStop, kbd.octave * 12 + pk(c).nt)
                  pk(c).dn = 0

                  SELECT CASE c MOD 12 ' reset key text
                  CASE 1, 3, 6, 8, 10  ' black keys
                      PrintKeyText(kbd, pk(c), IIF(kbd.KeyText, %GRAY, %BLACK), %BLACK, IIF(kbd.NoteText, %LTGRAY, %BLACK), %BLACK)
                  CASE ELSE : INCR d   ' white keys
                      PrintKeyText(kbd, pk(c), IIF(kbd.KeyText, %LTGRAY, %WHITE), %WHITE, IIF(kbd.NoteText, %GRAY, %WHITE), %WHITE)
                  END SELECT
                  InvalidateRect hWnd, BYVAL 0, 0 : UpdateWindow hWnd
                  EXIT FUNCTION
              END IF
          END IF
      END SELECT


  CASE %WM_LBUTTONDOWN, %WM_LBUTTONDBLCLK
      SetCapture(hWnd)
      pt.x = LO(WORD, lParam)
      pt.y = HI(WORD, lParam)
      FOR c = LBOUND(hRgn) TO UBOUND(hRgn)
          IF  c <> CurKey AND PtInRegion(hRgn(c), pt.x, pt.y) THEN
              pk(c).dn = 1 ' 1 = mouse key pressed it
              CurKey   = c
              kbd.note = kbd.octave * 12 + CurKey
              SendMessage(kbd.hParent, %WM_COMMAND, MAK(DWORD, kbd.CtlId, %PN_PLAYNOTE), kbd.note)
              Midi_PlayNote(kbd.hMidi, kbd.chPlay, kbd.note, kbd.volume)
              OldKey = CurKey

              SELECT CASE CurKey MOD 12
              CASE 1, 3, 6, 8, 10  ' black keys
                  PrintKeyText(kbd, pk(c), IIF(kbd.KeyText, %GRAY, %BLACK), %BLACK, IIF(kbd.NoteText, %YELLOW, %BLACK), %BLACK)
              CASE ELSE : INCR d   ' white keys
                  PrintKeyText(kbd, pk(c), IIF(kbd.KeyText, %LTGRAY, %WHITE), %WHITE, IIF(kbd.NoteText, %RED, %WHITE), %WHITE)
              END SELECT
              InvalidateRect hWnd, BYVAL 0, 0 : UpdateWindow hWnd
              EXIT FOR
          END IF
      NEXT

  CASE %WM_MOUSEMOVE
      IF (GetAsyncKeyState(%VK_LBUTTON) AND &H8000) THEN
          IF GetCapture() <> hWnd THEN SetCapture(hWnd)  ' ensure mouse capture

          pt.x = LO(WORD, lParam)
          pt.y = HI(WORD, lParam)
          FOR c = LBOUND(hRgn) TO UBOUND(hRgn)
              IF PtInRegion(hRgn(c), pt.x, pt.y) AND c <> CurKey THEN
                  CurKey = c
                  EXIT FOR
              END IF
          NEXT
          IF CurKey <> OldKey THEN
              pk(OldKey).dn = 0 ' 1 = mouse key pressed it
              kbd.note = kbd.octave * 12 + OldKey
              Midi_StopNote(kbd.hMidi, kbd.chStop, kbd.note) ' stop playing this note
              SELECT CASE OldKey MOD 12 ' reset key text
              CASE 1, 3, 6, 8, 10  ' black keys
                  PrintKeyText(kbd, pk(OldKey), IIF(kbd.KeyText, %GRAY, %BLACK), %BLACK, IIF(kbd.NoteText, %LTGRAY, %BLACK), %BLACK)
              CASE ELSE : INCR d   ' white keys
                  PrintKeyText(kbd, pk(OldKey), IIF(kbd.KeyText, %LTGRAY, %WHITE), %WHITE, IIF(kbd.NoteText, %GRAY, %WHITE), %WHITE)
              END SELECT

              pk(CurKey).dn = 1 ' 1 = mouse key pressed it
              kbd.note = kbd.octave * 12 + CurKey
              Midi_PlayNote(kbd.hMidi, kbd.chPlay, kbd.note, kbd.volume)
              SELECT CASE CurKey MOD 12
              CASE 1, 3, 6, 8, 10  ' black keys
                  PrintKeyText(kbd, pk(CurKey), IIF(kbd.KeyText, %GRAY, %BLACK), %BLACK, IIF(kbd.NoteText, %YELLOW, %BLACK), %BLACK)
              CASE ELSE : INCR d   ' white keys
                  PrintKeyText(kbd, pk(CurKey), IIF(kbd.KeyText, %LTGRAY, %WHITE), %WHITE, IIF(kbd.NoteText, %RED, %WHITE), %WHITE)
              END SELECT
              InvalidateRect hWnd, BYVAL 0, 0 : UpdateWindow hWnd

              OldKey = CurKey
          END IF
      END IF

  CASE %WM_LBUTTONUP
      pk(CurKey).dn = 0
      kbd.note = kbd.octave * 12 + CurKey
      SendMessage(kbd.hParent, %WM_COMMAND, MAK(DWORD, kbd.CtlId, %PN_STOPNOTE), kbd.note)
      Midi_StopNote(kbd.hMidi, kbd.chStop, kbd.note) ' and stop playing this note

      SELECT CASE CurKey MOD 12 ' reset key text
      CASE 1, 3, 6, 8, 10  ' black keys
          PrintKeyText(kbd, pk(CurKey), IIF(kbd.KeyText, %GRAY, %BLACK), %BLACK, IIF(kbd.NoteText, %LTGRAY, %BLACK), %BLACK)
      CASE ELSE : INCR d   ' white keys
          PrintKeyText(kbd, pk(CurKey), IIF(kbd.KeyText, %LTGRAY, %WHITE), %WHITE, IIF(kbd.NoteText, %GRAY, %WHITE), %WHITE)
      END SELECT

      InvalidateRect hWnd, BYVAL 0, 0 : UpdateWindow hWnd
      CurKey = -1
      OldKey = -1
      ReleaseCapture()

  '------------------------------------------------------------------
  ' a bunch of useful %PM_ Piano Messages
  '------------------------------------------------------------------
  CASE %PM_SETDEVICE
      ' can't test this one, because like most, I only have one midi device. Hope it works..
      IF kbd.hMidi THEN midiOutClose(kbd.hMidi) ' Close current midi device
      IF midiOutOpen(kbd.hMidi, wParam, 0, 0, 0) <> %MMSYSERR_NOERROR THEN
          MSGBOX "midiOutOpen failed - trying default device instead!", _
                 %MB_ICONERROR, "Error message!"
          midiOutOpen(kbd.hMidi, -1, 0, 0, 0)
      END IF
      Midi_SetInstrument(kbd.hMidi, kbd.chInstr, kbd.instrument)  ' reset selected instrument
      Midi_SetVolume(kbd.hMidi, kbd.chControl, kbd.volume)        ' reset volume 
      Midi_SetBalance(kbd.hMidi, kbd.chControl, kbd.balance)      ' reset balance
      Midi_SetSustain(kbd.hMidi, kbd.chControl, kbd.sustain)      ' reset sustain
      SetFocus hWnd
      EXIT FUNCTION

  CASE %PM_SETCHANNEL
      Midi_StopAllNotes(kbd.hMidi, kbd.chControl)  ' stop all notes
      Midi_SetSustain(kbd.hMidi, kbd.chControl, 0) ' reset sustain

      kbd.chControl = &HB0 + wParam
      kbd.chInstr   = &HC0 + wParam
      kbd.chStop    = &H80 + wParam
      kbd.chPlay    = &H90 + wParam
      kbd.chBend    = &HE0 + wParam

      Midi_SetInstrument(kbd.hMidi, kbd.chInstr, kbd.instrument)  ' reset selected instrument
      Midi_SetVolume(kbd.hMidi, kbd.chControl, kbd.volume)        ' reset volume 
      Midi_SetBalance(kbd.hMidi, kbd.chControl, kbd.balance)      ' reset balance
      Midi_SetSustain(kbd.hMidi, kbd.chControl, kbd.sustain)      ' reset sustain
      SetFocus hWnd
      EXIT FUNCTION

  CASE %PM_LISTDEVICES     ' wParam is list id
      Midi_ListDevices(kbd.hParent, wParam)
      EXIT FUNCTION

  CASE %PM_LISTCHANNELS    ' wParam is list id
      Midi_ListChannels(kbd.hParent, wParam)

  CASE %PM_LISTPATCHES     ' wParam is list id
      Midi_ListPatches(kbd.hParent, wParam)
      EXIT FUNCTION

  CASE %PM_LISTINSTRUMENTS ' wParam is list id, lParam is patches number
      Midi_ListInstruments(kbd.hParent, wParam, lParam)
      EXIT FUNCTION

  CASE %PM_GETOCTAVE  'returns currently selected base octave
      FUNCTION = kbd.octave
      EXIT FUNCTION

  CASE %PM_SETOCTAVE  ' sets 0 to 9 octaves
      Midi_StopAllNotes(kbd.hMidi, kbd.chControl)   ' stop all notes
      kbd.octave = wParam
      FOR c = LBOUND(pk) TO UBOUND(pk)
          IF pk(c).dn THEN
              kbd.note = kbd.octave * 12 + pk(c).nt ' refresh note
              Midi_PlayNote(kbd.hMidi, kbd.chPlay, kbd.note, kbd.volume)
          END IF
      NEXT
      IF lParam THEN SetFocus hWnd
      EXIT FUNCTION

  CASE %PM_SETINSTRUMENT  ' 0 to 127 instruments
      FOR c = LBOUND(pk) TO UBOUND(pk) ' On change, stop all notes
          Midi_StopNote(kbd.hMidi, kbd.chStop, kbd.octave * 12 + pk(c).nt)
          pk(c).dn = 0
      NEXT
      kbd.instrument = wParam
      Midi_SetInstrument(kbd.hMidi, kbd.chInstr, kbd.instrument)  'set selected instrument
      IF lParam THEN SetFocus hWnd
      EXIT FUNCTION

  CASE %PM_SETPITCH   ' -64 to 64, where 0 is normal pitch
      Midi_PitchBend(kbd.hMidi, kbd.chBend, wParam)
      IF lParam THEN SetFocus hWnd
      EXIT FUNCTION

  CASE %PM_SETVOLUME  ' 0 to 127, where 127 is max sound volume
      kbd.volume = wParam
      Midi_SetVolume(kbd.hMidi, kbd.chControl, kbd.volume)
      IF lParam THEN SetFocus hWnd
      EXIT FUNCTION

  CASE %PM_SETBALANCE  ' -64 to 64, where 0 is centered
      kbd.balance = wParam
      Midi_SetBalance(kbd.hMidi, kbd.chControl, kbd.balance)
      IF lParam THEN SetFocus hWnd
      EXIT FUNCTION

  CASE %PM_SETVIBRATO  ' 0 to 127, where 0 is no vibrato
      kbd.vibrato = wParam
      Midi_SetModulation(kbd.hMidi, kbd.chControl, kbd.vibrato)
      IF lParam THEN SetFocus hWnd
      EXIT FUNCTION

  CASE %PM_SETSUSTAIN  ' on/off, where anything below 64 is off (0/127 = off/on)
      kbd.sustain = wParam
      Midi_SetSustain(kbd.hMidi, kbd.chControl, kbd.sustain)
      IF lParam THEN SetFocus hWnd
      EXIT FUNCTION

  CASE %PM_GETKEYCOUNT
      FUNCTION = kbd.KeyCount + 1
      EXIT FUNCTION

  CASE %PM_SETKEYCOUNT  ' set total number of keys to draw/play
      kbd.KeyCount = MIN&(127, wParam - 1)
      DrawKeyBoard(hWnd, kbd, pk(), hRgn())
      InvalidateRect hWnd, BYVAL 0, 0 : UpdateWindow hWnd
      IF lParam THEN SetFocus hWnd
      EXIT FUNCTION

  CASE %PM_SETKEYTEXT  ' if to print keyboard key letter
      kbd.KeyText = wParam
      DrawKeyBoard(hWnd, kbd, pk(), hRgn())
      InvalidateRect hWnd, BYVAL 0, 0 : UpdateWindow hWnd
      IF lParam THEN SetFocus hWnd
      EXIT FUNCTION

  CASE %PM_SETNOTETEXT  ' if to print key note
      kbd.NoteText = wParam
      DrawKeyBoard(hWnd, kbd, pk(), hRgn())
      InvalidateRect hWnd, BYVAL 0, 0 : UpdateWindow hWnd
      IF lParam THEN SetFocus hWnd
      EXIT FUNCTION

  CASE %PM_PLAYNOTE  ' to enable playing notes via this message
      kbd.note   = MIN&(127, wParam) ' play this note
      kbd.volume = MIN&(127, lParam) ' 127 is max volume
      Midi_PlayNote(kbd.hMidi, kbd.chPlay, kbd.note, kbd.volume)

  CASE %PM_STOPNOTE  ' to enable stop playing notes via this message
      kbd.note = MIN&(127, wParam)
      Midi_StopNote(kbd.hMidi, kbd.chStop, kbd.note) ' stop playing this note

  END SELECT

  FUNCTION = DefWindowProc(hWnd, wMsg, wParam, lParam)
END FUNCTION


'==============================================================================
' Initialize keyboard variables
'------------------------------------------------------------------------------
SUB initKbd (hWnd AS DWORD, kbd AS kbData)
  kbd.hParent    = GetParent(hWnd)
  kbd.CtlId      = GetDlgCtrlID(hWnd)
  kbd.hFont      = MakeFontEx(0, "Arial", 9, 0, %FW_BOLD, 0, 0) 'GetStockObject(%ANSI_VAR_FONT) 
  kbd.KeyText    = 1
  kbd.NoteText   = 1
  kbd.octave     = 4         ' start with octave 3 and 5
  kbd.KeyCount   = 28        ' zerobased number of keys to draw (28 = 29 keys)
  kbd.volume     = 127       ' max sound volume
  kbd.instrument = 0         ' start off with piano
  kbd.sustain    = 0         ' start off with no sustain
  kbd.vibrato    = 0         ' start off with no vibrato
  kbd.balance    = 64        ' set balance to centered
  kbd.chStop     = &H80      ' Note off for channel 1 to 16 / &H80 to &H8F
  kbd.chPlay     = &H90      ' Note oon for channel 1 to 16 / &H90 to &H9F
  kbd.chControl  = &HB0      ' Control mode change for channel 1 to 16 / &HB0 to &HBF
  kbd.chInstr    = &HC0      ' Program change (instrument) for channel 1 to 16 / &HC0 to &HCF
  kbd.chBend     = &HE0      ' Pitch wheel range (0 TO 127, +/- two notes) for channel 1 to 16 / &HE0 to &HEF
  Midi_SetInstrument(kbd.hMidi, kbd.chInstr, kbd.instrument)  ' init instrument
  Midi_SetVolume(kbd.hMidi, kbd.chControl, kbd.volume)        ' init volume
  Midi_SetBalance(kbd.hMidi, kbd.chControl, kbd.balance)      ' init balance
END SUB


'==============================================================================
' Initialize keyboard memDC. Look worse than it is...  or maybe not. :) 
'------------------------------------------------------------------------------
SUB DrawKeyBoard (hWnd AS DWORD, kbd AS kbData, pk() AS pKey, hRgn() AS DWORD)
  LOCAL b, c, d, h, hb, w, wb AS LONG
  LOCAL dwRes, hDC, hBrush, hOldPen, hPen, tmpFont AS DWORD
  LOCAL rc AS RECT
  DIM pts(1 TO 8) AS POINTAPI

  IF kbd.memDC THEN  ' restore and delete what we have created
      IF kbd.hBitOld THEN SelectObject kbd.memDC, kbd.hBitOld
      IF kbd.hBitmap THEN DeleteObject kbd.hBitmap
      DeleteDC kbd.memDC
  END IF

  hDC = GetDc(hWnd)
    GetClientRect hWnd, rc
    kbd.memDC   = CreateCompatibleDC(hDC)
    kbd.hBitmap = CreateCompatibleBitmap(hDC, rc.nRight, rc.nBottom)
    kbd.hBitOld = SelectObject(kbd.memDC, kbd.hBitmap)
  ReleaseDC hWnd, hDC

  FillRect(kbd.memDC, rc, GetSysColorBrush(%COLOR_3DFACE))
  tmpFont = SelectObject(kbd.memDC, kbd.hFont)

  IF hRgn(LBOUND(hRgn)) THEN
      FOR c = LBOUND(hRgn) TO UBOUND(hRgn)
          IF hRgn(c) THEN DeleteObject(hRgn(c))
      NEXT
  END IF

  d = 0
  FOR c = 0 TO kbd.KeyCount
      SELECT CASE c MOD 12
      CASE 1, 3, 6, 8, 10  ' black keys
      CASE ELSE : INCR d   ' count white keys
      END SELECT
      pk(c).nt = c
  NEXT

  w  = rc.nRight / d                ' white key width
  IF d * w > rc.nRight THEN DECR w  ' ensure last key is fully inside control
  IF w MOD 2 THEN DECR w            ' need to ensure clean divide by 2

  IF d * w < rc.nRight THEN         ' shrink piano control to visible keys width
      SetWindowPos(hWnd, 0, 0, 0, d * w, rc.nBottom, %SWP_NOMOVE OR %SWP_NOZORDER)
  END IF

  wb = w * 0.7              ' black key width
  IF wb MOD 2 THEN DECR wb  ' need to ensure clean divide by 2
  h  = rc.nBottom           ' white key height
  hb = h * 0.7              ' black key height

  '--------------------------------------------------------------
  ' draw left white key. We use regions for mouse detection.
  '--------------------------------------------------------------
  hBrush  = SelectObject(kbd.memDC, GetStockObject(%WHITE_BRUSH))
  hPen    = CreatePen(%PS_SOLID, 1, RGB(96,96,96))
  hOldPen = SelectObject(kbd.memDC, hPen)

  pts(1).x =  0          : pts(1).y = 0   ' Region coordinates for Left white key
  pts(2).x =  w - wb / 2 : pts(2).y = 0
  pts(3).x =  pts(2).x   : pts(3).y = hb
  pts(4).x =  w          : pts(4).y = pts(3).y
  pts(5).x =  w          : pts(5).y = h
  pts(6).x =  0          : pts(6).y = h

  b = 0 : d = 0
  FOR c = 0 TO kbd.KeyCount
      IF (c > 0) AND (c MOD 12 = 0) THEN
          b = b + 7 * w
      END IF
      SELECT CASE c MOD 12
      CASE 0, 5
          SELECT CASE c MOD 12
          CASE 0 : d = 0
          CASE 5 : d = 3
          END SELECT
          hRgn(c) = CreatePolygonRgn(pts(1), 6, %ALTERNATE)  ' create region
          OffsetRgn hRgn(c), b + d * w, 0   ' move region into place
          GetRgnBox hRgn(c), rc             ' get region rect
          pk(c).rc = rc                     ' and use rect for RoundRect draw
          RoundRect(kbd.memDC, rc.nLeft, rc.nTop, rc.nRight, rc.nBottom, 5, 5)
          PrintKeyText(kbd, pk(c), IIF(kbd.KeyText, %LTGRAY, %WHITE), %WHITE, IIF(kbd.NoteText, %GRAY, %WHITE), %WHITE)
      END SELECT
  NEXT

  '--------------------------------------------------------------
  ' Draw middle white key
  '--------------------------------------------------------------
  pts(1).x =  wb / 2         : pts(1).y = 0
  pts(2).x =  w - wb / 2     : pts(2).y = 0
  pts(3).x =  pts(2).x       : pts(3).y = hb
  pts(4).x =  w              : pts(4).y = pts(3).y
  pts(5).x =  w              : pts(5).y = h
  pts(6).x =  0              : pts(6).y = h
  pts(7).x =  0              : pts(7).y = hb
  pts(8).x = pts(1).x        : pts(8).y = hb

  b = 0 : d = 0
  FOR c = 1 TO kbd.KeyCount
      SELECT CASE c MOD 12
      CASE 2, 7, 9
          SELECT CASE c MOD 12
          CASE 2 : d = 1
          CASE 7 : d = 4
          CASE 9 : d = 5
          END SELECT
          hRgn(c) = CreatePolygonRgn(pts(1), 8, %ALTERNATE)
          OffsetRgn hRgn(c), b + d * w, 0
          GetRgnBox hRgn(c), rc
          pk(c).rc = rc
          RoundRect(kbd.memDC, rc.nLeft, rc.nTop, rc.nRight, rc.nBottom, 5, 5)
          PrintKeyText(kbd, pk(c), IIF(kbd.KeyText, %LTGRAY, %WHITE), %WHITE, IIF(kbd.NoteText, %GRAY, %WHITE), %WHITE)
      END SELECT
      IF c MOD 12 = 0 THEN b = b + 7 * w
  NEXT

  '------------------------------------------------------------------
  ' draw Right white key
  '------------------------------------------------------------------
  pts(1).x =  wb / 2   : pts(1).y = 0
  pts(2).x =  w        : pts(2).y = 0
  pts(3).x =  w        : pts(3).y = h
  pts(4).x =  0        : pts(4).y = h
  pts(5).x =  0        : pts(5).y = hb
  pts(6).x =  wb / 2   : pts(6).y = hb

  b = 0 : d = 0
  FOR c = 1 TO kbd.KeyCount
      SELECT CASE c MOD 12
      CASE 4, 11
          SELECT CASE c MOD 12
          CASE  4 : d = 2
          CASE 11 : d = 6
          END SELECT
          hRgn(c) = CreatePolygonRgn(pts(1), 6, %ALTERNATE)
          OffsetRgn hRgn(c), b + d * w, 0
          GetRgnBox hRgn(c), rc
          pk(c).rc = rc
          RoundRect(kbd.memDC, rc.nLeft, rc.nTop, rc.nRight, rc.nBottom, 5, 5)
          PrintKeyText(kbd, pk(c), IIF(kbd.KeyText, %LTGRAY, %WHITE), %WHITE, IIF(kbd.NoteText, %GRAY, %WHITE), %WHITE)
      END SELECT
      IF c MOD 12 = 0 THEN b = b + 7 * w
  NEXT

  hOldPen = SelectObject(kbd.memDC, hOldPen)
  dwRes   = DeleteObject(hPen)
  SelectObject kbd.memDC, hBrush

  '--------------------------------------------------------------
  ' draw black keys
  '--------------------------------------------------------------
  pts(1).x = 0        : pts(1).y = 0
  pts(2).x = wb       : pts(2).y = 0
  pts(3).x = pts(2).x : pts(3).y = hb
  pts(4).x = 0        : pts(4).y = hb

  hBrush  = SelectObject(kbd.memDC, GetStockObject(%BLACK_BRUSH))
  hPen    = CreatePen(%PS_SOLID, 1, RGB(0,0,0))
  hOldPen = SelectObject(kbd.memDC, hPen)

  b = 0 : d = 0
  FOR c = 1 TO kbd.KeyCount  ' 0 is white key, C
      IF c MOD 12 = 0 THEN b = b + 7 * w
      SELECT CASE c MOD 12
      CASE 1, 3, 6, 8, 10
          SELECT CASE c MOD 12
          CASE  1 : d = 1
          CASE  3 : d = 2
          CASE  6 : d = 4
          CASE  8 : d = 5
          CASE 10 : d = 6
          END SELECT
          hRgn(c) = CreatePolygonRgn(pts(1), 4, %ALTERNATE)
          OffsetRgn hRgn(c), b + d * w - wb / 2, 0
          GetRgnBox hRgn(c), rc
          pk(c).rc = rc
          RoundRect(kbd.memDC, rc.nLeft, rc.nTop, rc.nRight, rc.nBottom, 3, 3)
      END SELECT
  NEXT

  hOldPen = SelectObject(kbd.memDC, hOldPen)
  dwRes   = DeleteObject(hPen)
  SelectObject kbd.memDC, hBrush

  '--------------------------------------------------------------
  ' draw thin white rectangle inside black keys
  '--------------------------------------------------------------
  hBrush  = SelectObject(kbd.memDC, GetStockObject(%NULL_BRUSH))
  hPen    = CreatePen(%PS_SOLID, 1, RGB(255,255,255))
  hOldPen = SelectObject(kbd.memDC, hPen)

  FOR c = 1 TO kbd.KeyCount  ' 0 is white key, C
      SELECT CASE c MOD 12
      CASE 1, 3, 6, 8, 10
          GetRgnBox hRgn(c), rc
          InflateRect(rc, -2, -2)
          Rectangle(kbd.memDC, rc.nLeft, rc.nTop, rc.nRight, rc.nBottom)
          PrintKeyText(kbd, pk(c), IIF(kbd.KeyText, %GRAY, %BLACK), %BLACK, IIF(kbd.NoteText, %LTGRAY, %BLACK), %BLACK)
      END SELECT
  NEXT

  hOldPen = SelectObject(kbd.memDC, hOldPen)
  dwRes   = DeleteObject(hPen)
  SelectObject kbd.memDC, hBrush
  IF tmpFont THEN SelectObject kbd.memDC, tmpFont

 END SUB
 

'==============================================================================
' Create a desirable font and return its handle.
'------------------------------------------------------------------------------
FUNCTION MakeFontEx (BYVAL hDC AS DWORD,  BYVAL FontName AS STRING, _
                     BYVAL PointSize AS LONG, BYVAL Angle AS LONG, BYVAL fBold AS LONG, _
                     BYVAL fItalic AS LONG, BYVAL fUnderline AS LONG) AS DWORD

  LOCAL CharSet AS LONG, CyPixels AS LONG

  IF hDC = 0 THEN
      hDC = GetDC(%HWND_DESKTOP)
      CyPixels  = GetDeviceCaps(hDC, %LOGPIXELSY)
      EnumFontFamilies hDC, BYVAL STRPTR(FontName), CODEPTR(EnumCharSet), BYVAL VARPTR(CharSet)
      ReleaseDC %HWND_DESKTOP, hDC
  ELSE
      CyPixels  = GetDeviceCaps(hDC, %LOGPIXELSY)
      EnumFontFamilies hDC, BYVAL STRPTR(FontName), CODEPTR(EnumCharSet), BYVAL VARPTR(CharSet)
  END IF
  PointSize = 0 - (PointSize * CyPixels) \ 72

  FUNCTION = CreateFont(PointSize, 0, _  'height, width (default=0)
             Angle, Angle, _             'escapement(angle), orientation
             fBold, _                    'weight (%FW_DONTCARE = 0, %FW_NORMAL = 400, %FW_BOLD = 700)
             fItalic, _                  'Italic
             fUnderline, _               'Underline
             %FALSE, _                   'StrikeThru
             CharSet, %OUT_TT_PRECIS, _
             %CLIP_DEFAULT_PRECIS, %DEFAULT_QUALITY, _
             %FF_DONTCARE , BYCOPY FontName)

END FUNCTION
'==============================================================================
' Get type of character set - ansi, symbol.. a must for some fonts.
'------------------------------------------------------------------------------
FUNCTION EnumCharSet (elf AS ENUMLOGFONT, ntm AS NEWTEXTMETRIC, _
                      BYVAL FontType AS LONG, CharSet AS LONG) AS LONG
  CharSet = elf.elfLogFont.lfCharSet
END FUNCTION


'==============================================================================
' Custom ptInRect function to enable compile with older compiler/Win32API.inc
'------------------------------------------------------------------------------
FUNCTION fnPtInRect(rc AS RECT, pt AS POINTAPI) AS LONG
  #IF %PB_REVISION >= &H1000         ' if compiler PBWIN10 or later
      FUNCTION = ptInRect(rc, pt)
  #ELSE                              ' else
      FUNCTION = ptInRect(rc, pt.x, pt.y)
  #ENDIF
END FUNCTION


'==============================================================================
' Get note description for 7+5=12 keys in any octave
'------------------------------------------------------------------------------
FUNCTION GetNoteText(BYVAL iNote AS BYTE) AS STRING
  SELECT CASE iNote MOD 12
  CASE  0 : FUNCTION = "C"
  CASE  1 : FUNCTION = "C#"
  CASE  2 : FUNCTION = "D"
  CASE  3 : FUNCTION = "D#"
  CASE  4 : FUNCTION = "E"
  CASE  5 : FUNCTION = "F"
  CASE  6 : FUNCTION = "F#"
  CASE  7 : FUNCTION = "G"
  CASE  8 : FUNCTION = "G#"
  CASE  9 : FUNCTION = "A"
  CASE 10 : FUNCTION = "A#"
  CASE 11 : FUNCTION = "B"
  END SELECT
END FUNCTION


'==============================================================================
' Print note + key text on key, using desired colors.
'------------------------------------------------------------------------------
SUB PrintKeyText(kbd AS kbData, pk AS pKey, _
                 iForeColor1 AS LONG, iBackColor1 AS LONG, _
                 iForeColor2 AS LONG, iBackColor2 AS LONG)

  LOCAL tmpFont AS DWORD, sText AS STRING, rc AS RECT, sz AS SIZEL

  tmpFont = SelectObject(kbd.memDC, kbd.hFont)
  GetTextExtentPoint32 kbd.memDC, "W", 1, sz  ' need font height
  rc = pk.rc

  sText = MID$($KEY1, pk.nt + 1, 1)  ' draw keyboard key text
  SetBkColor kbd.memDC, iBackColor1
  SetTextColor kbd.memDC, iForeColor1
  rc.nBottom = pk.rc.nBottom - 1.7 * sz.cy
  DrawText kbd.memDC, BYVAL STRPTR(sText), LEN(sText), rc, %DT_BOTTOM OR %DT_CENTER OR %DT_SINGLELINE

  sText = GetNoteText(pk.nt)         ' draw note text
  SetBkColor kbd.memDC, iBackColor2
  SetTextColor kbd.memDC, iForeColor2
  rc.nBottom = pk.rc.nBottom - 0.5 * sz.cy
  DrawText kbd.memDC, BYVAL STRPTR(sText), LEN(sText), rc, %DT_BOTTOM OR %DT_CENTER OR %DT_SINGLELINE

  IF tmpFont THEN SelectObject kbd.memDC, tmpFont  ' restore font

END SUB


'==============================================================================
' Enumerate midi devices - cannot test if it works 100%, but think it's ok..
'----------------------------------------------------------------------------------
SUB Midi_ListDevices (hDlg AS LONG, ctlId AS LONG)
  LOCAL c AS LONG, hCtrl, dwMsg1, dwMsg2 AS DWORD
  LOCAL MidiCaps AS MIDIOUTCAPS, zTxt AS ASCIIZ * 32

  hCtrl = GetDlgItem(hDlg, CtlId)
  GetClassName hCtrl, zTxt, SIZEOF(zTxt) ' see what kind of control to fill

  IF UCASE$(zTxt) = "COMBOBOX" THEN      ' if to add results to a ComboBox
      SendMessage(hCtrl, %CB_RESETCONTENT, 0, 0) 
      dwMsg1 = %CB_ADDSTRING
      dwMsg2 = %CB_SETCURSEL
  ELSE                                   ' elseif to add results to a ListBox
      SendMessage(hCtrl, %LB_RESETCONTENT, 0, 0) 
      dwMsg1 = %LB_ADDSTRING
      dwMsg2 = %LB_SETCURSEL
  END IF
  '------------------------------------------------------------------
  FOR c = 1 TO midiOutGetNumDevs  ' there's usually only one...
      midiOutGetDevCaps c - 1, BYVAL VARPTR(MidiCaps), LEN(MidiCaps)
      SendMessage(hCtrl, dwMsg1, 0, BYVAL VARPTR(MidiCaps.szPname)) 
  NEXT
  SendMessage(hCtrl, dwMsg2, 0, 0) 
END SUB


'==============================================================================
' Fill a ComboBox or ListBox with midi channels
'------------------------------------------------------------------------------
SUB Midi_ListChannels(BYVAL hDlg AS DWORD, BYVAL CtlId AS DWORD)

  LOCAL c AS LONG, hCtrl, dwMsg1, dwMsg2 AS DWORD, zTxt AS ASCIIZ * 64

  hCtrl = GetDlgItem(hDlg, CtlId)
  GetClassName hCtrl, zTxt, SIZEOF(zTxt) ' see what kind of control to fill

  IF UCASE$(zTxt) = "COMBOBOX" THEN      ' if to add channels to a ComboBox
      SendMessage(hCtrl, %CB_RESETCONTENT, 0, 0) 
      dwMsg1 = %CB_ADDSTRING
      dwMsg2 = %CB_SETCURSEL
  ELSEIF UCASE$(zTxt) = "LISTBOX" THEN   ' elseif to add channels to a ListBox
      SendMessage(hCtrl, %LB_RESETCONTENT, 0, 0) 
      dwMsg1 = %LB_ADDSTRING
      dwMsg2 = %LB_SETCURSEL
  END IF

  FOR c = 1 TO 16
      IF c = 10 THEN
          zTxt = " Channel" + STR$(c) + "  (Percussion set)"
      ELSE
          zTxt = " Channel" + STR$(c)
      END IF
      SendMessage(hCtrl, dwMsg1, 0, BYVAL VARPTR(zTxt)) 
  NEXT
  SendMessage(hCtrl, dwMsg2, 0, 0) 

END SUB


'==============================================================================
' Fill a ComboBox or ListBox with midi patches (instrument categories)
'------------------------------------------------------------------------------
SUB Midi_ListPatches(BYVAL hDlg AS DWORD, BYVAL CtlId AS DWORD)

  LOCAL c AS LONG, hCtrl, dwMsg AS DWORD, zTxt AS ASCIIZ * 64

  hCtrl = GetDlgItem(hDlg, CtlId)
  GetClassName hCtrl, zTxt, SIZEOF(zTxt) ' see what kind of control to fill

  IF UCASE$(zTxt) = "COMBOBOX" THEN      ' if to add data to a ComboBox
      SendMessage(hCtrl, %CB_RESETCONTENT, 0, 0) 
      dwMsg = %CB_ADDSTRING
  ELSEIF UCASE$(zTxt) = "LISTBOX" THEN   ' elseif to add data to a ListBox
      SendMessage(hCtrl, %LB_RESETCONTENT, 0, 0) 
      dwMsg = %LB_ADDSTRING
  END IF

  FOR c = 1 TO DATACOUNT
     zTxt = READ$(c)
     SendMessage(hCtrl, dwMsg, 0, BYVAL VARPTR(zTxt)) 
  NEXT

DATA "Piano", "Chromatic Percussion", "Organ", "Guitar"
DATA "Bass", "Strings", "Ensamble", "Brass"
DATA "Reed", "Pipe", "Synth Lead", "Synth Pad"
DATA "Synth Effects", "Ethnic", "Percussive", "Sound effects"

END SUB


'==============================================================================
' Fill a ListBox or ComboBox with midi instruments
'------------------------------------------------------------------------------
SUB Midi_ListInstruments (BYVAL hDlg AS DWORD, BYVAL CtlId AS DWORD, BYVAL iPatches AS LONG)
  LOCAL c, iStart, iStop AS LONG
  LOCAL hCtrl, dwMsg1, dwMsg2 AS DWORD, zTxt AS ASCIIZ * 64

  SELECT CASE iPatches  ' divide into iPatch categories
  CASE  0 : iStart =   1 : iStop =   8  ' Piano
  CASE  1 : iStart =   9 : iStop =  16  ' Chromatic Percussion
  CASE  2 : iStart =  17 : iStop =  24  ' Organ
  CASE  3 : iStart =  25 : iStop =  32  ' Guitar
  CASE  4 : iStart =  33 : iStop =  40  ' Bass
  CASE  5 : iStart =  41 : iStop =  48  ' Strings
  CASE  6 : iStart =  49 : iStop =  56  ' Ensamble
  CASE  7 : iStart =  57 : iStop =  64  ' Brass
  CASE  8 : iStart =  65 : iStop =  72  ' Reed
  CASE  9 : iStart =  73 : iStop =  80  ' Pipe
  CASE 10 : iStart =  81 : iStop =  88  ' Synth Lead
  CASE 11 : iStart =  89 : iStop =  96  ' Synth Pad
  CASE 12 : iStart =  97 : iStop = 104  ' Synth Effects
  CASE 13 : iStart = 105 : iStop = 112  ' Ethnic
  CASE 14 : iStart = 113 : iStop = 119  ' Percussive
  CASE 15 : iStart = 120 : iStop = 128  ' Sound effects
  END SELECT
  
  hCtrl = GetDlgItem(hDlg, CtlId)
  GetClassName hCtrl, zTxt, SIZEOF(zTxt) ' see what kind of control to fill

  IF UCASE$(zTxt) = "LISTBOX" THEN       ' if to add data to a ListBox
      SendMessage(hCtrl, %LB_RESETCONTENT, 0, 0)
      dwMsg1 = %LB_ADDSTRING
      dwMsg2 = %LB_SETITEMDATA
  ELSEIF UCASE$(zTxt) = "COMBOBOX" THEN  ' elseif to add data to a ComboBox
      SendMessage(hCtrl, %CB_RESETCONTENT, 0, 0)
      dwMsg1 = %CB_ADDSTRING
      dwMsg2 = %CB_SETITEMDATA
  END IF

  FOR c = iStart TO iStop
     zTxt = FORMAT$(c, "* #") + "  " + READ$(c)
     SendMessage(hCtrl, dwMsg1, 0, BYVAL VARPTR(zTxt)) 
     SendMessage(hCtrl, dwMsg2, c - iStart, c - 1)
  NEXT

'------------------------------------------------------------------------------
DATA "Acoustic Grand Piano", "Bright Acoustic Piano", "Electric Grand Piano"
DATA "Honky-tonk Piano", "Rhodes Piano", "Chorused Piano", "Harpsichord"
DATA "Clavinet", "Celesta", "Glockenspiel", "Music Box", "Vibraphone"
DATA "Marimba", "Xylophone", "Tubular Bells", "Dulcimer", "Hammond Organ"
DATA "Percussive Organ", "Rock Organ", "Church Organ", "Reed Organ"
DATA "Accordion", "Harmonica", "Tango Accordion", "Acoustic Guitar (nylon)"
DATA "Acoustic Guitar (steel)", "Electric Guitar (jazz)", "Electric Guitar (clean)"
DATA "Electric Guitar (muted)", "Overdriven Guitar", "Distortion Guitar"
DATA "Guitar Harmonics", "Acoustic Bass", "Electric Bass (finger)", "Electric Bass (pick)"
DATA "Fretless Bass", "Slap Bass 1", "Slap Bass 2", "Synth Bass 1", "Synth Bass 2"
DATA "Violin", "Viola", "Cello", "Contrabass", "Tremolo Strings", "Pizzicato Strings"
DATA "Orchestral Harp", "Timpani", "String Ensemble 1", "String Ensemble 2"
DATA "SynthStrings 1", "SynthStrings 2", "Choir Aahs", "Voice Oohs", "Synth Voice"
DATA "Orchestra Hit", "Trumpet", "Trombone", "Tuba", "Muted Trumpet", "French Horn"
DATA "Brass Section", "Synth Brass 1", "Synth Brass 2", "Soprano Sax", "Alto Sax"
DATA "Tenor Sax", "Baritone Sax", "Oboe", "English Horn", "Bassoon", "Clarinet"
DATA "Piccolo", "Flute", "Recorder", "Pan Flute", "Bottle Blow", "Shakuhachi"
DATA "Whistle", "Ocarina", "Lead 1 (square)", "Lead 2 (sawtooth)", "Lead 3 (calliope lead)"
DATA "Lead 4 (chiff lead)", "Lead 5 (charang)", "Lead 6 (voice)", "Lead 7 (fifths)"
DATA "Lead 8 (bass + lead)", "Pad 1 (new age)", "Pad 2 (warm)", "Pad 3 (polysynth)"
DATA "Pad 4 (choir)", "Pad 5 (bowed)", "Pad 6 (metallic)", "Pad 7 (halo)", "Pad 8 (sweep)"
DATA "FX 1 (rain)", "FX 2 (soundtrack)", "FX 3 (crystal)", "FX 4 (atmosphere)"
DATA "FX 5 (brightness)", "FX 6 (goblins)", "FX 7 (echoes)", "FX 8 (sci-fi)"
DATA "Sitar", "Banjo", "Shamisen", "Koto", "Kalimba", "Bagpipe", "Fiddle"
DATA "Shanai", "Tinkle Bell", "Agogo", "Steel Drums", "Woodblock", "Taiko Drum"
DATA "Melodic Tom", "Synth Drum", "Reverse Cymbal", "Guitar Fret Noise"
DATA "Breath Noise", "Seashore", "Bird Tweet", "Telephone Ring", "Helicopter"
DATA "Applause", "Gunshot"

END SUB


'==============================================================================
' Set desired instrument, event/channel &HB0 to &HBF.
'------------------------------------------------------------------------------
FUNCTION Midi_SetInstrument(BYVAL hMidi AS DWORD, _
                            BYVAL bEvent AS BYTE, _            ' event/channel
                            BYVAL bInstrument AS BYTE) AS LONG ' instrument
  LOCAL dwData AS DWORD
  dwData = bEvent + bInstrument * &H100     ' setup dwData
  FUNCTION = midiOutShortMsg(hMidi, dwData) ' set instrument
END FUNCTION


'==============================================================================
' Set sound volume for selected channel, event/channel &HC0 to &HCF.
'------------------------------------------------------------------------------
FUNCTION Midi_SetVolume(BYVAL hMidi    AS DWORD, _
                        BYVAL iEvent   AS LONG, _       'channel &HB0 to &HBF
                        BYVAL iVolume AS LONG) AS LONG  'volume, 0 to 127
  LOCAL dwData AS DWORD
  dwData   = iEvent + &H7 * &H100 + iVolume * &H10000   ' &H7 = 7
  FUNCTION = midiOutShortMsg(hMidi, dwData)
END FUNCTION


'==============================================================================
' Set balance for selected channel.  0 = Left, 64 = Center, 127 = Right
'------------------------------------------------------------------------------
FUNCTION Midi_SetBalance(BYVAL hMidi    AS DWORD, _
                         BYVAL iEvent   AS LONG, _       'channel &HB0 to &HBF
                         BYVAL iBalance AS LONG) AS LONG 'balance
  LOCAL dwData AS DWORD
  dwData   = iEvent + &HA * &H100 + iBalance * &H10000   ' &HA = 10
  FUNCTION = midiOutShortMsg(hMidi, dwData)
END FUNCTION


'==============================================================================
' Set Modulation/Vibrato for selected channel. iModulation = 1 to 127
'------------------------------------------------------------------------------
FUNCTION Midi_SetModulation(BYVAL hMidi       AS DWORD, _
                            BYVAL iEvent      AS LONG, _       'channel &HB0 to &HBF
                            BYVAL iModulation AS LONG) AS LONG 'balance, 127 = on, 0 = off
  LOCAL dwData AS DWORD
  dwData   = iEvent + &H1 * &H100 + iModulation * &H10000  ' &H1 = 1
  FUNCTION = midiOutShortMsg(hMidi, dwData)
END FUNCTION


'==============================================================================
' Pitch Bend +/- 2 notes, 0 = Max down, 64=Normal note, 127 = Max up
'------------------------------------------------------------------------------
FUNCTION Midi_PitchBend (BYVAL hMidi AS DWORD, _
                         BYVAL bEvent AS BYTE, _       ' event/channel &HE0 to &HEF
                         BYVAL bPitch AS BYTE) AS LONG ' Pitch bend 0 to 127
  LOCAL dwData AS DWORD
  dwData = bEvent + bPitch * &H10000        ' setup dwData
  FUNCTION = midiOutShortMsg(hMidi, dwData) ' set instrument
END FUNCTION


'==============================================================================
' Set Sustain for selected channel
'------------------------------------------------------------------------------
FUNCTION Midi_SetSustain(BYVAL hMidi AS DWORD, _
                         BYVAL bEvent AS LONG, _        'channel &HB0 to &HBF
                         BYVAL sustain AS LONG) AS LONG 'sustain, 127 = on, 0 = off
  LOCAL dwData AS DWORD
  dwData = bEvent + &H40 * &H100 + sustain * &H10000    ' &H40 = 64
  FUNCTION = midiOutShortMsg(hMidi, dwData)
END FUNCTION


'==============================================================================
' Play a note using midiOutShortMsg, channel &H90 to &H9F (ch &H99 = percussion)
'------------------------------------------------------------------------------
FUNCTION Midi_PlayNote (BYVAL hMidi AS DWORD, _           ' 
                        BYVAL bEvent AS BYTE, _           ' event/channel
                        BYVAL bNote AS BYTE, _            ' play this note
                        BYVAL bVolume AS BYTE) AS DWORD   ' sound volume, 0 to 127
  LOCAL dwData AS DWORD
  dwData = bEvent + bNote * &H100 + bVolume * &H10000     ' setup for playing
  FUNCTION = midiOutShortMsg(hMidi, dwData)               ' play note
END FUNCTION


'==============================================================================
' Turn off all notes in selected channel, &HB0 to &HBF.
'------------------------------------------------------------------------------
FUNCTION Midi_StopAllNotes(BYVAL hMidi AS DWORD, _
                           BYVAL bEvent AS LONG) AS LONG  'event/channel
  LOCAL dwData AS DWORD
  dwData = bEvent + &H7B * &H100  ' &H7B = 123
  FUNCTION = midiOutShortMsg(hMidi, dwData)
END FUNCTION


'==============================================================================
' Stop playing a note using midiOutShortMsg, channel &H80 to &H8F.
'------------------------------------------------------------------------------
FUNCTION Midi_StopNote (BYVAL hMidi AS DWORD, _
                        BYVAL bEvent AS BYTE, _       ' event/channel
                        BYVAL bNote AS BYTE) AS LONG  ' note
  LOCAL dwData AS DWORD
  dwData = bEvent + (bNote * &H100)          ' setup dwData
  FUNCTION = midiOutShortMsg(hMidi, dwData)  ' stop playing
END FUNCTION

