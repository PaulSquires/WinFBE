' ########################################################################################
' Microsoft Windows
' File: CPiano.inc
' Contents: Piano control.
' Based on the Piano.inc, public domain control originally written by Börje Hagsten in
' June 2016 for PowerBASIC.
' Compiler: FreeBasic 32 & 64-bit
' Written in 2017 by José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "win/mmsystem.bi"
#include once "Afx/CWindow.inc"
USING Afx

NAMESPACE Afx

CONST PN_PLAYNOTE        = WM_APP + 300  ' note played notification
CONST PN_STOPNOTE        = WM_APP + 301  ' note stopped notification
CONST PN_OCTAVE          = WM_APP + 310  ' octave change notification
CONST PN_PITCHBEND       = WM_APP + 320  ' pitch bend notification

TYPE CPIANO_KEY        ' Key array data
  nt  AS LONG          ' Note (0 - 127)
  dn  AS LONG          ' key up/down flag, 0/1
  rc  AS RECT          ' key coordinates
END TYPE

CONST CPIANO_KEY1 = WSTR("ZSXDCVGBHNJMQ2W3ER5T6Y7UI9O0P")  ' 29 computer key note layout

' // Note: Declared as global because declaring them as instance variables of the class
' // causes strange memory problems, like the m_Octave variable being sometimes reset to 0
' // for no apparent reason.
DIM SHARED m_pk(127) AS CPIANO_KEY    ' // Static key data array
DIM SHARED m_hRgn(127) AS HRGN        ' // Region array

' ========================================================================================
' CPiano class
' ========================================================================================
TYPE CPiano

Private:
   m_hCtl       AS HWND          ' // Button handle
   m_hParent    AS HWND          ' // Parent dialog handle
   m_CtlId      AS LONG_PTR      ' // Piano control id
   m_hFont      AS HFONT         ' // Font handle
   m_hMidi      AS HMIDIOUT      ' // MIDI device handle
   m_MemDC      AS HDC           ' // Memory device context
   m_hBitmap    AS HBITMAP       ' // Bitmap handle
   m_hBmpOld    AS HBITMAP       ' // Old bitmap handle
   m_KeyCount   AS DWORD = 28    ' // Number of keys to draw
   m_Instrument AS LONG = 0      ' // Selected m_Instrument (0 - 127)
   m_chPlay     AS LONG = &h90   ' // Channel &H90 to &H9F, where ch &H99 = percussion
   m_chStop     AS LONG = &h80   ' // Channel &H80 to &H8F, where ch &H89 = percussion
   m_chControl  AS LONG = &hB0   ' // Channel &HB0 to &HBF, Control mode change
   m_chInstr    AS LONG = &hC0   ' // Channel &HC0 to &HCF, Program instrument change
   m_chBend     AS LONG = &hE0   ' // Channel &HE0 to &HEF, Pitch bend
   m_KeyText    AS LONG = 1      ' // Key description on keys on/off
   m_Note       AS LONG = 0      ' // Note to play (0 - 127)
   m_NoteText   AS LONG = 1      ' // Note description on keys on/off
   m_Octave     AS LONG = 4      ' // Base octave (1 octave = 12 notes)
   m_Sustain    AS LONG = 0      ' // Sustain on/off
   m_Volume     AS LONG = 127    ' // Sound volume
   m_Balance    AS LONG = 64     ' // Sound balance Left/Right
   m_Vibrato    AS LONG = 0      ' // Modulation/vibrato
'   m_pk(127)    AS CPIANO_KEY    ' // Static key data array
'   m_hRgn(127)  AS HRGN          ' // Region array

Private:
   DECLARE FUNCTION SetInstrument_ (BYVAL nEvent AS LONG, BYVAL nInstrument AS LONG) AS MMRESULT
   DECLARE FUNCTION SetVolume_ (BYVAL nEvent AS LONG, BYVAL nVolume AS LONG) AS MMRESULT
   DECLARE FUNCTION SetBalance_ (BYVAL nEvent AS LONG, BYVAL nBalance AS LONG) AS MMRESULT
   DECLARE FUNCTION SetModulation_ (BYVAL nEvent AS LONG, BYVAL nModulation AS LONG) AS MMRESULT
   DECLARE FUNCTION PitchBend_ (BYVAL nEvent AS LONG, BYVAL nPitch AS LONG) AS MMRESULT
   DECLARE FUNCTION SetSustain_ (BYVAL nEvent AS LONG, BYVAL nSustain AS LONG) AS MMRESULT
   DECLARE FUNCTION GetNoteText_ (BYVAL iNote AS LONG) AS STRING
   DECLARE SUB PrintKeyText_ (BYREF pk AS CPIANO_KEY, BYVAL iForeColor1 AS LONG, BYVAL iBackColor1 AS LONG, BYVAL iForeColor2 AS LONG, BYVAL iBackColor2 AS LONG)
   DECLARE SUB DrawKeyboard_

Public:
   DECLARE CONSTRUCTOR (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, _
      BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
      BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL lpParam AS LONG_PTR = 0)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION hWindow () AS HWND
   DECLARE STATIC FUNCTION CPianoProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   DECLARE FUNCTION PlayNote (BYVAL nEvent AS LONG, BYVAL nNote AS LONG, BYVAL nVolume AS LONG) AS MMRESULT
   DECLARE FUNCTION StopAllNotes (BYVAL nEvent AS LONG) AS MMRESULT
   DECLARE FUNCTION StopNote (BYVAL nEvent AS LONG, BYVAL nNote AS LONG) AS MMRESULT
   DECLARE FUNCTION SetDevice (BYVAL uDeviceID AS UINT) AS MMRESULT
   DECLARE SUB SetChannel (BYVAL nChannel AS LONG)
   DECLARE FUNCTION GetOctave () AS LONG
   DECLARE SUB SetOctave (BYVAL nOctave AS LONG, BYVAL bSetFocus AS BOOLEAN = TRUE)
   DECLARE SUB SetInstrument (BYVAL nInstrument AS LONG, BYVAL bSetFocus AS BOOLEAN = TRUE)
   DECLaRE SUB SetPitch (BYVAL nValue AS LONG, BYVAL bSetFocus AS BOOLEAN = TRUE)
   DECLARE SUB SetVolume (BYVAL nValue AS LONG, BYVAL bSetFocus AS BOOLEAN = TRUE)
   DECLARE SUB SetBalance (BYVAL nValue AS LONG, BYVAL bSetFocus AS BOOLEAN = TRUE)
   DECLARE SUB SetVibrato (BYVAL nValue AS LONG, BYVAL bSetFocus AS BOOLEAN = TRUE)
   DECLARE SUB SetSustain (BYVAL nValue AS LONG, BYVAL bSetFocus AS BOOLEAN = TRUE)
   DECLARE FUNCTION GetKeyCount () AS LONG
   DECLARE SUB SetKeyCount (BYVAL nKeys AS LONG, BYVAL bSetFocus AS BOOLEAN = TRUE)
   DECLARE SUB SetKeyText (BYVAL bOnOff AS WINBOOL, BYVAL bSetFocus AS BOOLEAN = TRUE)
   DECLARE SUB SetNoteText (BYVAL bOnOff AS WINBOOL, BYVAL bSetFocus AS BOOLEAN = TRUE)

END TYPE
' ========================================================================================

' ========================================================================================
' CPiano class constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CPiano (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, _
   BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
   BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL lpParam AS LONG_PTR = 0)

   IF pWindow = NULL THEN EXIT CONSTRUCTOR
   ' // Opens a MIDI output device for playback
   IF midiOutOpen(@m_hMidi, -1, NULL, NULL, 0) <> MMSYSERR_NOERROR THEN
      MessageBoxW(pWindow->hWindow, "midiOutOpen failed!", "Error", MB_OK OR MB_ICONERROR)
      EXIT CONSTRUCTOR
   END IF

   ' // Register the class
   DIM wAtom AS ATOM
   DIM wcexw AS WNDCLASSEXW
   DIM wszClassName AS WSTRING * 260 = "FBPIANOCTX"
   IF .GetClassInfoExW(.GetModuleHandleW(NULL), @wszClassName, @wcexw) = 0 THEN
      ' // Fill the WNDCLASSEXW structure
      WITH wcexw
         .cbSize        = SIZEOF(wcexw)
         .style         = CS_DBLCLKS
         .lpfnWndProc   = @CPiano.CpianoProc
         .cbClsExtra    = 0
         .cbWndExtra    = SIZEOF(HANDLE)
         .hInstance     = ..GetModuleHandleW(NULL)
         .hCursor       = ..LoadCursorW(NULL, CAST(LPCWSTR, IDC_HAND))
         .hbrBackground = NULL
         .lpszMenuName  = NULL
         .lpszClassName = @wszClassName
         .hIcon         = NULL
         .hIconSm       = NULL
      END WITH
      wAtom = .RegisterClassExW(@wcexw)
   END IF

   ' // Create the control
   IF dwStyle = 0 THEN dwStyle = WS_CHILD OR WS_TABSTOP OR WS_VISIBLE OR WS_BORDER
   IF pWindow THEN m_hCtl = pWindow->AddControl(wszClassName, pWindow->hWindow, cID, "", x, y, nWidth, nHeight, dwStyle, dwExStyle, lpParam)
   IF m_hCtl = NULL THEN EXIT CONSTRUCTOR
   SetWindowLongPtrW m_hCtl, 0, CAST(LONG_PTR, @this)
   ' // Set the same font used by the parent, but bolded
   DIM lfw AS LOGFONTW
   IF pWindow->Font THEN
      IF GetObjectW(pWindow->Font, SIZEOF(lfw), @lfw) THEN
         lfw.lfWeight = FW_BOLD
         m_hFont = CreateFontIndirectW(@lfw)
      END IF
   END IF
   
   ' // Store values
   m_hParent = pWindow->hWindow
   m_CtlId = cID
   m_Octave = 4
   ' // Set the initial instrument
   this.SetInstrument_(m_chInstr, m_Instrument)
   ' // Set the initial volume
   this.SetVolume_(m_chControl, m_Volume)
   ' // Set the initial balance
   this.SetBalance_(m_chControl, m_Balance)
   ' // Draw the kwyboard
   this.DrawKeyBoard_

END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CPiano class destructor
' ========================================================================================
PRIVATE DESTRUCTOR CPiano
   ' // Close the MIDI device
   IF m_hMidi THEN midiOutClose(m_hMidi)
   ' // Free resources
   IF m_MemDC THEN
      IF m_hBmpOld THEN SelectObject(m_MemDC, m_hBmpOld)
      IF m_hBitmap THEN DeleteObject(m_hBitmap)
      DeleteDC m_MemDC
   END IF
   IF m_hFont THEN DeleteObject m_hFont
   FOR i AS LONG = LBOUND(m_hRgn) TO UBOUND(m_hRgn)
      IF m_hRgn(i) THEN DeleteObject(m_hRgn(i))
   NEXT
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the handle of the button
' ========================================================================================
PRIVATE FUNCTION CPiano.hWindow () AS HWND
   FUNCTION = m_hCtl
END FUNCTION
' ========================================================================================

' ========================================================================================
' Set desired instrument, event/channel &HB0 to &HBF.
' ========================================================================================
FUNCTION CPiano.SetInstrument_ (BYVAL nEvent AS LONG, BYVAL nInstrument AS LONG) AS MMRESULT
   DIM dwData AS DWORD = nEvent + nInstrument * &h100
   RETURN midiOutShortMsg(m_hMidi, dwData)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Set sound volume for selected channel, event/channel &HC0 to &HCF.
' ========================================================================================
FUNCTION CPiano.SetVolume_ (BYVAL nEvent AS LONG, BYVAL nVolume AS LONG) AS MMRESULT
   DIM dwData AS DWORD = nEvent + &h7 * &h100 + nVolume * &h10000
   RETURN midiOutShortMsg(m_hMidi, dwData)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Set balance for selected channel.  0 = Left, 64 = Center, 127 = Right
' ========================================================================================
FUNCTION CPiano.SetBalance_ (BYVAL nEvent AS LONG, BYVAL nBalance AS LONG) AS MMRESULT
   DIM dwData AS DWORD = nEvent + &hA * &h100 + nBalance * &h10000
   RETURN midiOutShortMsg(m_hMidi, dwData)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Set modulation/vibrato for selected channel. nModulation = 1 to 127
' ========================================================================================
FUNCTION CPiano.SetModulation_ (BYVAL nEvent AS LONG, BYVAL nModulation AS LONG) AS MMRESULT
   DIM dwData AS DWORD = nEvent + &h1 * &h100 + nModulation * &h10000
   RETURN midiOutShortMsg(m_hMidi, dwData)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Set desired instrument, event/channel &HB0 to &HBF.
' ========================================================================================
FUNCTION CPiano.PitchBend_ (BYVAL nEvent AS LONG, BYVAL nPitch AS LONG) AS MMRESULT
   DIM dwData AS DWORD = nEvent + nPitch * &h10000
   RETURN midiOutShortMsg(m_hMidi, dwData)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Set sustain for selected channel
' ========================================================================================
FUNCTION CPiano.SetSustain_ (BYVAL nEvent AS LONG, BYVAL nSustain AS LONG) AS MMRESULT
   DIM dwData AS DWORD = nEvent + &h40 * &h100 + nSustain * &h10000
   RETURN midiOutShortMsg(m_hMidi, dwData)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Play a note, channel &H90 to &H9F (ch &H99 = percussion)
' ========================================================================================
FUNCTION CPiano.PlayNote (BYVAL nEvent AS LONG, BYVAL nNote AS LONG, BYVAL nVolume AS LONG) AS MMRESULT
   DIM dwData AS DWORD = nEvent + nNote * &h100 + nVolume * &h10000
   RETURN midiOutShortMsg(m_hMidi, dwData)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Turn off all notes in selected channel, &HB0 to &HBF.
' ========================================================================================
FUNCTION CPiano.StopAllNotes (BYVAL nEvent AS LONG) AS MMRESULT
   DIM dwData AS DWORD = nEvent + &h7B * &h100
   RETURN midiOutShortMsg(m_hMidi, dwData)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Stop playing a note, channel &H80 to &H8F.
' ========================================================================================
FUNCTION CPiano.StopNote (BYVAL nEvent AS LONG, BYVAL nNote AS LONG) AS MMRESULT
   DIM dwData AS DWORD = nEvent + nNote * &h100
   RETURN midiOutShortMsg(m_hMidi, dwData)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Set the MIDI device
' Can't test this one, because like most, I only have one midi device. Hope it works..
' ========================================================================================
FUNCTION CPiano.SetDevice (BYVAL uDeviceID AS UINT) AS MMRESULT
   ' // Close current MIDI device
   IF m_hMidi THEN midiOutClose(m_hMidi)
   IF midiOutOpen(@m_hMidi, uDeviceID, NULL, NULL, 0) <> MMSYSERR_NOERROR THEN
      MessageBoxW(m_hParent, "midiOutOpen failed - trying default device instead!", _
         "Error", MB_OK OR MB_ICONERROR)
      DIM res AS MMRESULT = midiOutOpen(@m_hMidi, -1, NULL, NULL, 0)
      IF res <> MMSYSERR_NOERROR THEN RETURN res
   END IF
   this.SetInstrument_(m_chInstr, m_Instrument)   ' Reset selected instrument
   this.SetVolume_(m_chControl, m_Volume)         ' Reset volume 
   this.SetBalance_(m_chControl, m_Balance)       ' Reset balance
   this.SetSustain_(m_chControl, m_Sustain)       ' Reset sustain
   SetFocus m_hCtl
   RETURN MMSYSERR_NOERROR
END FUNCTION
' ========================================================================================

' ========================================================================================
' Set the channel
' ========================================================================================
SUB CPiano.SetChannel (BYVAL nChannel AS LONG)
   this.StopAllNotes(m_chControl)
   this.SetSustain_(m_chControl, 0)
   m_chControl = &hB0 + nChannel
   m_chInstr   = &hC0 + nChannel
   m_chStop    = &h80 + nChannel
   m_chPlay    = &h90 + nChannel
   m_chBend    = &hE0 + nChannel
   this.SetInstrument_(m_chInstr, m_Instrument)
   this.SetVolume_(m_chControl, m_Volume)
   this.SetBalance_(m_chControl, m_Balance)
   this.SetSustain_(m_chControl, m_Sustain)
   SetFocus m_hCtl
END SUB
' ========================================================================================

' ========================================================================================
' Returns the currently selected octave.
' ========================================================================================
FUNCTION CPiano.GetOctave () AS LONG
   RETURN m_Octave
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the octave (0-9)
' ========================================================================================
SUB CPiano.SetOctave (BYVAL nOctave AS LONG, BYVAL bSetFocus AS BOOLEAN = TRUE)
   this.StopAllNotes(m_chControl)
   m_Octave = nOctave
   FOR i AS LONG = LBOUND(m_pk) TO UBOUND(m_pk)
      IF m_pk(i).dn THEN
         m_Note = m_Octave * 12 + m_pk(i).nt
         this.PlayNote(m_chPlay, m_Note, m_Volume)
      END IF
   NEXT
   IF bSetFocus THEN SetFocus m_hCtl
END SUB
' ========================================================================================

' ========================================================================================
' Sets the instrument (0-127)
' ========================================================================================
SUB CPiano.SetInstrument (BYVAL nInstrument AS LONG, BYVAL bSetFocus AS BOOLEAN = TRUE)
   FOR i AS LONG = LBOUND(m_pk) TO UBOUND(m_pk)
      this.StopNote(m_chStop, m_Octave * 12 + m_pk(i).nt)
      m_pk(i).dn = 0
   NEXT
   m_Instrument = nInstrument
   this.SetInstrument_(m_chInstr, m_Instrument)
   IF bSetFocus THEN SetFocus m_hCtl
END SUB
' ========================================================================================

' ========================================================================================
' Sets the pitch (-64 to 64, where 0 is normal pitch)
' ========================================================================================
SUB CPiano.SetPitch (BYVAL nValue AS LONG, BYVAL bSetFocus AS BOOLEAN = TRUE)
   this.PitchBend_(m_chBend, nValue)
   IF bSetFocus THEN SetFocus m_hCtl
END SUB
' ========================================================================================

' ========================================================================================
' Sets the volume (0 to 127, where 127 is max sound volume)
' ========================================================================================
SUB CPiano.SetVolume (BYVAL nValue AS LONG, BYVAL bSetFocus AS BOOLEAN = TRUE)
   m_Volume = nValue
   this.SetVolume_(m_chControl, nValue)
   IF bSetFocus THEN SetFocus m_hCtl
END SUB
' ========================================================================================

' ========================================================================================
' Set the balance (-64 to 64, where 0 is centered)
' ========================================================================================
SUB CPiano.SetBalance (BYVAL nValue AS LONG, BYVAL bSetFocus AS BOOLEAN = TRUE)
   m_Balance = nValue
   this.SetBalance_(m_chControl, nValue)
   IF bSetFocus THEN SetFocus m_hCtl
END SUB
' ========================================================================================

' ========================================================================================
' Set the vibrato (0 to 127, where 0 is no vibrato)
' ========================================================================================
SUB CPiano.SetVibrato (BYVAL nValue AS LONG, BYVAL bSetFocus AS BOOLEAN = TRUE)
   m_Balance = nValue
   this.SetModulation_(m_chControl, nValue)
   IF bSetFocus THEN SetFocus m_hCtl
END SUB
' ========================================================================================

' ========================================================================================
' Set the vibrato (on/off, where anything below 64 is off (0/127 = off/on)
' ========================================================================================
SUB CPiano.SetSustain (BYVAL nValue AS LONG, BYVAL bSetFocus AS BOOLEAN = TRUE)
   m_Sustain = nValue
   this.SetSustain_(m_chControl, nValue)
   IF bSetFocus THEN SetFocus m_hCtl
END SUB
' ========================================================================================

' ========================================================================================
' Get the key count
' ========================================================================================
FUNCTION CPiano.GetKeyCount () AS LONG
   RETURN m_KeyCount + 1
END FUNCTION
' ========================================================================================

' ========================================================================================
' Get note description for 7+5=12 keys in any octave
' ========================================================================================
FUNCTION CPiano.GetNoteText_ (BYVAL iNote AS LONG) AS STRING
  SELECT CASE iNote MOD 12
     CASE  0 : RETURN "C"
     CASE  1 : RETURN "C#"
     CASE  2 : RETURN "D"
     CASE  3 : RETURN "D#"
     CASE  4 : RETURN "E"
     CASE  5 : RETURN "F"
     CASE  6 : RETURN "F#"
     CASE  7 : RETURN "G"
     CASE  8 : RETURN "G#"
     CASE  9 : RETURN "A"
     CASE 10 : RETURN "A#"
     CASE 11 : RETURN "B"
  END SELECT
END FUNCTION
' ========================================================================================

' ========================================================================================
' Print note + key text on key, using desired colors.
' ========================================================================================
SUB CPiano.PrintKeyText_ (BYREF pk AS CPIANO_KEY, _
                          BYVAL iForeColor1 AS LONG, BYVAL iBackColor1 AS LONG, _
                          BYVAL iForeColor2 AS LONG, BYVAL iBackColor2 AS LONG)

  DIM tmpFont AS HFONT = SelectObject(m_MemDC, m_hFont)
  DIM sz AS SIZEL
  GetTextExtentPoint32 m_MemDC, "W", 1, @sz  ' need font height
  DIM rc AS RECT = pk.rc

  DIM wszText AS WSTRING * 260 = MID(CPIANO_KEY1, pk.nt + 1, 1)  ' draw keyboard key text
  SetBkColor m_MemDC, iBackColor1
  SetTextColor m_MemDC, iForeColor1
  rc.Bottom = pk.rc.Bottom - 1.7 * sz.cy
  DrawTextW m_MemDC, @wszText, LEN(wszText), @rc, DT_BOTTOM OR DT_CENTER OR DT_SINGLELINE

  wszText = this.GetNoteText_(pk.nt)         ' draw note text
  SetBkColor m_MemDC, iBackColor2
  SetTextColor m_MemDC, iForeColor2
  rc.Bottom = pk.rc.Bottom - 0.5 * sz.cy
  DrawTextW m_MemDC, @wszText, LEN(wszText), @rc, DT_BOTTOM OR DT_CENTER OR DT_SINGLELINE

  IF tmpFont THEN SelectObject m_MemDC, tmpFont  ' restore font

END SUB
' ========================================================================================

' ========================================================================================
' Initialize keyboard memory device context
' ========================================================================================
SUB CPiano.DrawKeyBoard_

   IF m_MemDC THEN
      IF m_hBmpOld THEN SelectObject(m_MemDC, m_hBmpOld)
      IF m_hBitmap THEN DeleteObject(m_hBitmap)
      DeleteDC m_MemDC
   END IF

   DIM hDC AS HDC = GetDC(m_hCtl)
   DIM rc AS RECT
   GetClientRect m_hCtl, @rc
   m_MemDC   = CreateCompatibleDC(hDC)
   m_hBitmap = CreateCompatibleBitmap(hDC, rc.Right, rc.Bottom)
   m_hBmpOld = SelectObject(m_MemDC, m_hBitmap)
   ReleaseDC m_hCtl, hDC

   FillRect(m_MemDC, @rc, GetSysColorBrush(COLOR_3DFACE))
   DIM tmpFont AS HFONT = SelectObject(m_MemDC, m_hFont)

   IF m_hRgn(LBOUND(m_hRgn)) THEN
      FOR i AS LONG = LBOUND(m_hRgn) TO UBOUND(m_hRgn)
         IF m_hRgn(i) THEN DeleteObject(m_hRgn(i))
      NEXT
   END IF

   DIM d AS LONG
   FOR i AS LONG = 0 TO m_KeyCount
      SELECT CASE i MOD 12
         CASE 1, 3, 6, 8, 10  ' black keys
         CASE ELSE : d += 1   ' count white keys
      END SELECT
      m_pk(i).nt = i
   NEXT

   DIM w AS LONG = rc.Right / d       ' white key width
'   IF d * w > rc.Right THEN w -= 1   ' ensure last key is fully inside control
'   IF w MOD 2 THEN w -= 1            ' need to ensure clean divide by 2
   IF d * w < rc.Right THEN           ' shrink piano control to visible keys width
      SetWindowPos(m_hCtl, 0, 0, 0, d * w, rc.Bottom, SWP_NOMOVE OR SWP_NOZORDER)
   END IF

   DIM wb AS LONG = w * 0.7           ' black key width
   IF wb MOD 2 THEN wb -= 1           ' need to ensure clean divide by 2
   DIM h AS LONG  = rc.Bottom         ' white key height
   DIM hb AS LONG = h * 0.7           ' black key height

  '--------------------------------------------------------------
  ' Draw left white key. We use regions for mouse detection.
  '--------------------------------------------------------------
  DIM hBrush AS HBRUSH  = SelectObject(m_MemDC, GetStockObject(WHITE_BRUSH))
  DIM hPen AS HPEN = CreatePen(PS_SOLID, 1, BGR(96,96,96))
  DIM hOldPen AS .HPEN = SelectObject(m_MemDC, hPen)

  DIM pts(1 TO 8) AS POINT
  pts(1).x =  0          : pts(1).y = 0   ' Region coordinates for Left white key
  pts(2).x =  w - wb / 2 : pts(2).y = 0
  pts(3).x =  pts(2).x   : pts(3).y = hb
  pts(4).x =  w          : pts(4).y = pts(3).y
  pts(5).x =  w          : pts(5).y = h
  pts(6).x =  0          : pts(6).y = h

  DIM b AS LONG
  d = 0
  FOR i AS LONG = 0 TO m_KeyCount
      IF (i > 0) AND (i MOD 12 = 0) THEN
         b = b + 7 * w
      END IF
      SELECT CASE i MOD 12
      CASE 0, 5
         SELECT CASE i MOD 12
            CASE 0 : d = 0
            CASE 5 : d = 3
         END SELECT
         m_hRgn(i) = CreatePolygonRgn(@pts(1), 6, ALTERNATE)  ' create region
         OffsetRgn m_hRgn(i), b + d * w, 0                    ' move region into place
         GetRgnBox m_hRgn(i), @rc                             ' get region rect
         m_pk(i).rc = rc                                      ' and use rect for RoundRect draw
         RoundRect(m_MemDC, rc.Left, rc.Top, rc.Right, rc.Bottom, 5, 5)
         ' &hD3D3D3 = light gray; &hFFFFFF = white; &h808080 = gray
         this.PrintKeyText_(m_pk(i), IIF(m_KeyText, &hD3D3D3, &hFFFFFF), &hFFFFFF, IIF(m_NoteText, &h808080, &hFFFFFF), &hFFFFFF)
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
  FOR i AS LONG = 1 TO m_KeyCount
      SELECT CASE i MOD 12
      CASE 2, 7, 9
         SELECT CASE i MOD 12
            CASE 2 : d = 1
            CASE 7 : d = 4
            CASE 9 : d = 5
         END SELECT
         m_hRgn(i) = CreatePolygonRgn(@pts(1), 8, ALTERNATE)
         OffsetRgn m_hRgn(i), b + d * w, 0
         GetRgnBox m_hRgn(i), @rc
         m_pk(i).rc = rc
         RoundRect(m_MemDC, rc.Left, rc.Top, rc.Right, rc.Bottom, 5, 5)
         ' &hD3D3D3 = light gray; &hFFFFFF = white; &h808080 = gray
         this.PrintKeyText_(m_pk(i), IIF(m_KeyText, &hD3D3D3, &hFFFFFF), &hFFFFFF, IIF(m_NoteText, &h808080, &hFFFFFF), &hFFFFFF)
      END SELECT
      IF i MOD 12 = 0 THEN b = b + 7 * w
   NEXT

   '------------------------------------------------------------------
   ' Draw Right white key
   '------------------------------------------------------------------
   pts(1).x =  wb / 2   : pts(1).y = 0
   pts(2).x =  w        : pts(2).y = 0
   pts(3).x =  w        : pts(3).y = h
   pts(4).x =  0        : pts(4).y = h
   pts(5).x =  0        : pts(5).y = hb
   pts(6).x =  wb / 2   : pts(6).y = hb

   b = 0 : d = 0
   FOR i AS LONG = 1 TO m_KeyCount
      SELECT CASE i MOD 12
      CASE 4, 11
         SELECT CASE i MOD 12
            CASE  4 : d = 2
            CASE 11 : d = 6
         END SELECT
         m_hRgn(i) = CreatePolygonRgn(@pts(1), 6, ALTERNATE)
         OffsetRgn m_hRgn(i), b + d * w, 0
         GetRgnBox m_hRgn(i), @rc
         m_pk(i).rc = rc
         RoundRect(m_MemDC, rc.Left, rc.Top, rc.Right, rc.Bottom, 5, 5)
         ' &hD3D3D3 = light gray; &hFFFFFF = white; &h808080 = gray
         this.PrintKeyText_(m_pk(i), IIF(m_KeyText, &hD3D3D3, &hFFFFFF), &hFFFFFF, IIF(m_NoteText, &h808080, &hFFFFFF), &hFFFFFF)
      END SELECT
      IF i MOD 12 = 0 THEN b = b + 7 * w
   NEXT

   hOldPen = SelectObject(m_MemDC, hOldPen)
   DeleteObject(hPen)
   SelectObject m_MemDC, hBrush

   '--------------------------------------------------------------
   ' draw black keys
   '--------------------------------------------------------------
   pts(1).x = 0        : pts(1).y = 0
   pts(2).x = wb       : pts(2).y = 0
   pts(3).x = pts(2).x : pts(3).y = hb
   pts(4).x = 0        : pts(4).y = hb

   hBrush  = SelectObject(m_MemDC, GetStockObject(BLACK_BRUSH))
   hPen    = CreatePen(PS_SOLID, 1, BGR(0,0,0))
   hOldPen = SelectObject(m_MemDC, hPen)

   b = 0 : d = 0
   FOR i AS LONG = 1 TO m_KeyCount  ' 0 is white key, C
      IF i MOD 12 = 0 THEN b = b + 7 * w
      SELECT CASE i MOD 12
      CASE 1, 3, 6, 8, 10
         SELECT CASE i MOD 12
            CASE  1 : d = 1
            CASE  3 : d = 2
            CASE  6 : d = 4
            CASE  8 : d = 5
            CASE 10 : d = 6
         END SELECT
         m_hRgn(i) = CreatePolygonRgn(@pts(1), 4, ALTERNATE)
         OffsetRgn m_hRgn(i), b + d * w - wb / 2, 0
         GetRgnBox m_hRgn(i), @rc
         m_pk(i).rc = rc
         RoundRect(m_MemDC, rc.Left, rc.Top, rc.Right, rc.Bottom, 3, 3)
      END SELECT
   NEXT

   hOldPen = SelectObject(m_MemDC, hOldPen)
   DeleteObject(hPen)
   SelectObject m_MemDC, hBrush

   '--------------------------------------------------------------
   ' draw thin white rectangle inside black keys
   '--------------------------------------------------------------
   hBrush  = SelectObject(m_MemDC, GetStockObject(NULL_BRUSH))
   hPen    = CreatePen(PS_SOLID, 1, BGR(255,255,255))
   hOldPen = SelectObject(m_MemDC, hPen)

   FOR i AS LONG = 1 TO m_KeyCount  ' 0 is white key, C
      SELECT CASE i MOD 12
      CASE 1, 3, 6, 8, 10
         GetRgnBox m_hRgn(i), @rc
         InflateRect(@rc, -2, -2)
         Rectangle(m_MemDC, rc.Left, rc.Top, rc.Right, rc.Bottom)
         ' &hD3D3D3 = light gray; &hFFFFFF = white; &h808080 = gray; &h000000 = black
         this.PrintKeyText_(m_pk(i), IIF(m_KeyText, &h808080, &h000000), &h000000, IIF(m_NoteText, &hD3D3D3, &h000000), &h000000)
      END SELECT
   NEXT

   hOldPen = SelectObject(m_MemDC, hOldPen)
   DeleteObject(hPen)
   SelectObject m_MemDC, hBrush
   IF tmpFont THEN SelectObject m_MemDC, tmpFont

END SUB
' ========================================================================================

' ========================================================================================
' Sets the key count
' ========================================================================================
SUB CPiano.SetKeyCount (BYVAL nKeys AS LONG, BYVAL bSetFocus AS BOOLEAN = TRUE)
   m_KeyCount = MIN(127, nKeys - 1)
   this.DrawKeyBoard_
   InvalidateRect m_hCtl, NULL, FALSE
   UpdateWIndow m_hCtl
   IF bSetFocus THEN SetFocus m_hCtl
END SUB
' ========================================================================================

' ========================================================================================
' Prints the keyboard key letter
' ========================================================================================
SUB CPiano.SetKeyText (BYVAL bOnOff AS WINBOOL, BYVAL bSetFocus AS BOOLEAN = TRUE)
   m_KeyText = bOnOff
   this.DrawKeyboard_
   InvalidateRect m_hCtl, NULL, FALSE
   UpdateWIndow m_hCtl
   IF bSetFocus THEN SetFocus m_hCtl
END SUB
' ========================================================================================

' ========================================================================================
' Prints the key note
' ========================================================================================
SUB CPiano.SetNoteText (BYVAL bOnOff AS WINBOOL, BYVAL bSetFocus AS BOOLEAN = TRUE)
   m_NoteText = bOnOff
   this.DrawKeyboard_
   InvalidateRect m_hCtl, NULL, FALSE
   UpdateWIndow m_hCtl
   IF bSetFocus THEN SetFocus m_hCtl
END SUB
' ========================================================================================

' ========================================================================================
' Window procedure
' ========================================================================================
PRIVATE FUNCTION CPiano.CpianoProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   STATIC curKey AS LONG = -1, oldKey AS LONG = 0, chPitch AS LONG = 64

   SELECT CASE uMsg

      ' // Ensure the control processes all keys by itself
      CASE WM_GETDLGCODE
         FUNCTION = DLGC_WANTALLKEYS
         EXIT FUNCTION

      ' // On repaint, simply copy m_MemDC to control DC
      CASE WM_PAINT
         DIM ps AS PAINTSTRUCT
         BeginPaint hwnd, @ps
         DIM rc AS RECT
         GetClientRect hwnd, @rc
         DIM pPiano AS CPiano PTR = CAST(CPiano PTR, GetWindowLongPtrW(hwnd, 0))
         IF pPiano THEN BitBlt ps.hDC, 0, 0, rc.Right, rc.Bottom, pPiano->m_MemDC, 0, 0, SRCCOPY
         EndPaint hwnd, @ps
         EXIT FUNCTION

      CASE WM_KEYDOWN
         ' // Enable exit on escape key
         SELECT CASE wParam
            CASE VK_ESCAPE
               SendMessageW(GetParent(hwnd), WM_COMMAND, MAKELONG(IDCANCEL, BN_CLICKED), 0)

            CASE VK_UP  ' Arrow up ^  - Pitch bend up
               DIM pPiano AS CPiano PTR = CAST(CPiano PTR, GetWindowLongPtrW(hwnd, 0))
               IF pPiano THEN
                  IF chPitch = 64 THEN  ' send notification on each change to parent's/control id's %WM_COMMAND
                     FOR chPitch = 64 TO 128
                        pPiano->PitchBend_(pPiano->m_chBend, chPitch)
                        SendMessageW(GetParent(hwnd), WM_COMMAND, MAKELONG(pPiano->m_CtlId, PN_PITCHBEND), -(chPitch - 64))
                        IF (GetAsyncKeyState(VK_UP) AND &h8000) = 0 THEN EXIT FOR ' if key is released
                        SLEEP 8
                     NEXT
                  END IF
               END IF

            CASE VK_DOWN  ' Arrow down v  - Pitch bend down
               DIM pPiano AS CPiano PTR = CAST(CPiano PTR, GetWindowLongPtrW(hwnd, 0))
               IF pPiano THEN
                  IF chPitch = 64 THEN
                     FOR chPitch = 64 TO 0 STEP -1
                        pPiano->PitchBend_(pPiano->m_chBend, chPitch)
                        SendMessageW(GetParent(hwnd), WM_COMMAND, MAKELONG(pPiano->m_CtlId, PN_PITCHBEND), -(chPitch - 64))
                        IF (GetAsyncKeyState(VK_DOWN) AND &h8000) = 0 THEN EXIT FOR
                        SLEEP 8
                     NEXT
                  END IF
               END IF

            CASE VK_LEFT  ' Arrow left < - Octave change down
               DIM pPiano AS CPiano PTR = CAST(CPiano PTR, GetWindowLongPtrW(hwnd, 0))
               IF pPiano THEN
                  DIM nOctave AS LONG = MAX(0, pPiano->GetOctave - 1)
                  pPiano->SetOctave(nOctave, 1)
                  SendMessageW(GetParent(hwnd), WM_COMMAND, MAKELONG(pPiano->m_CtlId, PN_OCTAVE), pPiano->GetOctave)
               END IF

            CASE VK_RIGHT  ' Arrow right > - Octave change up
               DIM pPiano AS CPiano PTR = CAST(CPiano PTR, GetWindowLongPtrW(hwnd, 0))
               IF pPiano THEN
                  DIM nOctave AS LONG = MIN(10 - pPiano->GetKeyCount \ 12, pPiano->GetOctave + 1)
                  pPiano->SetOctave(nOctave, 1)
                  SendMessageW(GetParent(hwnd), WM_COMMAND, MAKELONG(pPiano->m_CtlId, PN_OCTAVE), pPiano->GetOctave)
               END IF

            CASE VK_0 TO VK_Z   ' See if a note key has been pressed
               DIM pPiano AS CPiano PTR = CAST(CPiano PTR, GetWindowLongPtrW(hwnd, 0))
               IF pPiano THEN
                  DIM i AS LONG = INSTR(CPIANO_KEY1, CHR(wParam))  ' CPIANO_KEY1 is used for pk array compare
                  IF i THEN
                     i -= 1
                     IF m_pk(i).dn = 0 THEN   ' if not pressed before
                        pPiano->m_Note = pPiano->GetOctave * 12 + m_pk(i).nt   ' actual note
                        m_pk(i).dn = 2   ' 2 = keyboard key pressed it
                        SendMessageW(GetParent(hwnd), WM_COMMAND, MAKELONG(pPiano->m_CtlId, PN_PLAYNOTE), pPiano->m_Note)
                        pPiano->PlayNote(pPiano->m_chPlay, pPiano->m_Note, pPiano->m_Volume)
                        SELECT CASE i MOD 12
                           CASE 1, 3, 6, 8, 10  ' black keys
                              ' &hD3D3D3 = light gray; &hFFFFFF = white; &h808080 = gray; &h000000 = black; &h00FFFF = yellow; &hFF0000 = blue
                              pPiano->PrintKeyText_(m_pk(i), IIF(pPiano->m_KeyText, &h808080, &h000000), &h000000, IIF(pPiano->m_NoteText, &h00FFFF, &h000000), &h000000)
                        CASE ELSE   ' white keys
                           pPiano->PrintKeyText_(m_pk(i), IIF(pPiano->m_KeyText, &hD3D3D3, &hFFFFFF), &hFFFFFF, IIF(pPiano->m_NoteText, &hFF0000, &hFFFFFF), &hFFFFFF)
                        END SELECT
                        InvalidateRect hwnd, NULL, FALSE
                        UpdateWindow hwnd
                     END IF
                  END IF
               END IF
         END SELECT
         EXIT FUNCTION

      CASE WM_KEYUP
         SELECT CASE wParam
            CASE VK_UP   ' Arrow up - Pitch has been bended up
               DIM pPiano AS CPiano PTR = CAST(CPiano PTR, GetWindowLongPtrW(hwnd, 0))
               IF pPiano THEN
                  IF chPitch > 64 THEN
                     chPitch = 64
                     pPiano->PitchBend_(pPiano->m_chBend, chPitch)
                     SendMessageW(GetParent(hwnd), WM_COMMAND, MAKELONG(pPiano->m_CtlId, PN_PITCHBEND), 0)
                  END IF
               END IF

            CASE VK_DOWN  ' Arrow down - Pitch has been bended down
               DIM pPiano AS CPiano PTR = CAST(CPiano PTR, GetWindowLongPtrW(hwnd, 0))
               IF pPiano THEN
                  chPitch = 64
                  pPiano->PitchBend_(pPiano->m_chBend, chPitch)
                  SendMessageW(GetParent(hwnd), WM_COMMAND, MAKELONG(pPiano->m_CtlId, PN_PITCHBEND), 0)
               END IF

            CASE VK_0 TO VK_Z  ' same as on key down, but here we stop playing the note and reset key text instead.
               DIM pPiano AS CPiano PTR = CAST(CPiano PTR, GetWindowLongPtrW(hwnd, 0))
               IF pPiano THEN
                  DIM i AS LONG = INSTR(CPIANO_KEY1, CHR(wParam))  ' see if it's a defined key
                  IF i THEN
                     i -= 1
                     IF m_pk(i).dn = 2 THEN   ' if it was pressed by a keyboard key
                        SendMessageW(GetParent(hwnd), WM_COMMAND, MAKELONG(pPiano->m_CtlId, PN_STOPNOTE), pPiano->m_Note)
                        pPiano->StopNote(pPiano->m_chStop, pPiano->GetOctave * 12 + m_pk(i).nt)
                        m_pk(i).dn = 0
                        SELECT CASE i MOD 12 ' reset key text
                           CASE 1, 3, 6, 8, 10  ' black keys
                              ' &hD3D3D3 = light gray; &hFFFFFF = white; &h808080 = gray; &h000000 = black; &h00FFFF = yellow; &hFF0000 = blue
                              pPiano->PrintKeyText_(m_pk(i), IIF(pPiano->m_KeyText, &h808080, &h000000), &h000000, IIF(pPiano->m_NoteText, &hD3D3D3, &h000000), &h000000)
                           CASE ELSE
                              ' white keys
                              pPiano->PrintKeyText_(m_pk(i), IIF(pPiano->m_KeyText, &hD3D3D3, &hFFFFFF), &hFFFFFF, IIF(pPiano->m_NoteText, &h808080, &hFFFFFF), &hFFFFFF)
                        END SELECT
                        InvalidateRect hwnd, NULL, FALSE
                        UpdateWindow hwnd
                        EXIT FUNCTION
                     END IF
                  END IF
               END IF
            END SELECT

      CASE WM_LBUTTONDOWN, WM_LBUTTONDBLCLK
         DIM pPiano AS CPiano PTR = CAST(CPiano PTR, GetWindowLongPtrW(hwnd, 0))
         IF pPiano THEN
            SetCapture(hwnd)
            DIM pt AS POINT
            pt.x = LOWORD(lParam)
            pt.y = HIWORD(lParam)
            FOR i AS LONG = LBOUND(m_hRgn) TO UBOUND(m_hRgn)
               IF i <> curKey AND PtInRegion(m_hRgn(i), pt.x, pt.y) THEN
                  m_pk(i).dn = 1 ' 1 = mouse key pressed it
                  curKey = i
                  pPiano->m_Note = pPiano->GetOctave * 12 + curKey
                  SendMessageW(GetParent(hwnd), WM_COMMAND, MAKELONG(pPiano->m_CtlId, PN_PLAYNOTE), pPiano->m_Note)
                  pPiano->PlayNote(pPiano->m_chPlay, pPiano->m_Note, pPiano->m_Volume)
                  oldKey = curKey
                  SELECT CASE curKey MOD 12
                     CASE 1, 3, 6, 8, 10  ' black keys
                        ' &hD3D3D3 = light gray; &hFFFFFF = white; &h808080 = gray; &h000000 = black; &h00FFFF = yellow; &hFF0000 = blue: &h0000FF = red
                        pPiano->PrintKeyText_(m_pk(i), IIF(pPiano->m_KeyText, &h808080, &h000000), &h000000, IIF(pPiano->m_NoteText, &h00FFFF, &h000000), &h000000)
                     CASE ELSE  ' white keys
                        pPiano->PrintKeyText_(m_pk(i), IIF(pPiano->m_KeyText, &hD3D3D3, &hFFFFFF), &hFFFFFF, IIF(pPiano->m_NoteText, &h0000FF, &hFFFFFF), &hFFFFFF)
                  END SELECT
                  InvalidateRect hwnd, NULL, FALSE
                  UpdateWindow hwnd
                  EXIT FOR
               END IF
            NEXT
         END IF

      CASE WM_LBUTTONUP
         DIM pPiano AS CPiano PTR = CAST(CPiano PTR, GetWindowLongPtrW(hwnd, 0))
         IF pPiano THEN
            m_pk(curKey).dn = 0
            pPiano->m_Note = pPiano->GetOctave * 12 + curKey
            SendMessageW(GetParent(hwnd), WM_COMMAND, MAKELONG(pPiano->m_CtlId, PN_STOPNOTE), pPiano->m_Note)
            pPiano->StopNote(pPiano->m_chStop, pPiano->m_Note) ' and stop playing this note
            SELECT CASE curKey MOD 12 ' reset key text
               CASE 1, 3, 6, 8, 10  ' black keys
                  ' &hD3D3D3 = light gray; &hFFFFFF = white; &h808080 = gray; &h000000 = black; &h00FFFF = yellow; &hFF0000 = blue: &h0000FF = red
                  pPiano->PrintKeyText_(m_pk(curKey), IIF(pPiano->m_KeyText, &h808080, &h000000), &h000000, IIF(pPiano->m_NoteText, &hD3D3D3, &h000000), &h000000)
               CASE ELSE   ' white keys
                  pPiano->PrintKeyText_(m_pk(curKey), IIF(pPiano->m_KeyText, &hD3D3D3, &hFFFFFF), &hFFFFFF, IIF(pPiano->m_NoteText, &h808080, &hFFFFFF), &hFFFFFF)
            END SELECT
            InvalidateRect hwnd, NULL, FALSE
            UpdateWindow hwnd
            curKey = -1
            oldKey = -1
            ReleaseCapture
         END IF

      CASE WM_MOUSEMOVE
         DIM pPiano AS CPiano PTR = CAST(CPiano PTR, GetWindowLongPtrW(hwnd, 0))
         IF pPiano THEN
            IF (GetAsyncKeyState(VK_LBUTTON) AND &h8000) THEN
               IF GetCapture <> hwnd THEN SetCapture(hwnd)  ' ensure mouse capture
               DIM pt AS POINT
               pt.x = LOWORD(lParam)
               pt.y = HIWORD(lParam)
               FOR i AS LONG = LBOUND(m_hRgn) TO UBOUND(m_hRgn)
                  IF PtInRegion(m_hRgn(i), pt.x, pt.y) AND i <> curKey THEN
                     curKey = i
                     EXIT FOR
                  END IF
               NEXT
               IF curKey <> oldKey THEN
                  m_pk(oldKey).dn = 0 ' 1 = mouse key pressed it
                  pPiano->m_Note = pPiano->GetOctave * 12 + oldKey
                  pPiano->StopNote(pPiano->m_chStop, pPiano->m_Note) ' stop playing this note
                  SELECT CASE oldKey MOD 12 ' reset key text
                     CASE 1, 3, 6, 8, 10  ' black keys
                        ' &hD3D3D3 = light gray; &hFFFFFF = white; &h808080 = gray; &h000000 = black; &h00FFFF = yellow; &hFF0000 = blue: &h0000FF = red
                        pPiano->PrintKeyText_(m_pk(oldKey), IIF(pPiano->m_KeyText, &h808080, &h000000), &h000000, IIF(pPiano->m_NoteText, &hD3D3D3, &h000000), &h000000)
                     CASE ELSE   ' white keys
                        pPiano->PrintKeyText_(m_pk(oldKey), IIF(pPiano->m_KeyText, &hD3D3D3, &hFFFFFF), &hFFFFFF, IIF(pPiano->m_NoteText, &h808080, &hFFFFFF), &hFFFFFF)
                  END SELECT
                  m_pk(curKey).dn = 1 ' 1 = mouse key pressed it
                  pPiano->m_Note = pPiano->GetOctave * 12 + curKey
                  pPiano->PlayNote(pPiano->m_chPlay, pPiano->m_Note, pPiano->m_Volume)
                  SELECT CASE CurKey MOD 12
                     CASE 1, 3, 6, 8, 10  ' black keys
                        pPiano->PrintKeyText_(m_pk(curKey), IIF(pPiano->m_KeyText, &h808080, &h000000), &h000000, IIF(pPiano->m_NoteText, &h00FFFF, &h000000), &h000000)
                     CASE ELSE   ' white keys
                        pPiano->PrintKeyText_(m_pk(CurKey), IIF(pPiano->m_KeyText, &hD3D3D3, &hFFFFFF), &hFFFFFF, IIF(pPiano->m_NoteText, &h0000FF, &hFFFFFF), &hFFFFFF)
                  END SELECT
                  InvalidateRect hwnd, NULL, FALSE
                  UpdateWindow hwnd
                  oldKey = curKey
               END IF
            END IF
         END IF

   END SELECT

   ' // Default processing for other messages.
   FUNCTION = DefWindowProcW(hWnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================

END NAMESPACE
