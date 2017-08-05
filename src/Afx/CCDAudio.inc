' ########################################################################################
' Microsoft Windows
' File: CCDAudio.inc
' Contents: MCI CD Audio wrapper class
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "win/mmsystem.bi"
#include once "Afx/CWSTR.inc"

NAMESPACE Afx

' ########################################################################################
' CCDAudio - MCI CD Audio class
' ########################################################################################
TYPE CCDAudio

Private:
   m_dwError AS MCIERROR

Public:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastError () AS MCIERROR
   DECLARE FUNCTION SetError (BYVAL dwError AS MCIERROR) AS MCIERROR
   DECLARE FUNCTION GetErrorString (BYVAL dwError AS MCIERROR = 0) AS CWSTR
   DECLARE FUNCTION Open () AS MCIERROR
   DECLARE FUNCTION Close () AS MCIERROR
   DECLARE FUNCTION IsReady () AS BOOLEAN
   DECLARE FUNCTION IsMediaInserted () AS BOOLEAN
   DECLARE FUNCTION IsPaused () AS BOOLEAN
   DECLARE FUNCTION IsStopped () AS BOOLEAN
   DECLARE FUNCTION IsPlaying () AS BOOLEAN
   DECLARE FUNCTION IsSeeking () AS BOOLEAN
   DECLARE FUNCTION OpenDoor () AS MCIERROR
   DECLARE FUNCTION CloseDoor () AS MCIERROR
   DECLARE FUNCTION Play () AS MCIERROR
   DECLARE FUNCTION PlayFrom (BYVAL nTrack AS LONG) AS MCIERROR
   DECLARE FUNCTION PlayFromTo (BYVAL nStartTrack AS LONG, BYVAL nEndTrack AS LONG) AS MCIERROR
   DECLARE FUNCTION Stop () AS MCIERROR
   DECLARE FUNCTION Pause () AS MCIERROR
   DECLARE FUNCTION GetTracksCount () AS LONG
   DECLARE FUNCTION GetCurrentTrack () AS LONG
   DECLARE FUNCTION GetTrackLengthString (BYVAL nTrack AS LONG) AS CWSTR
   DECLARE FUNCTION GetTrackLength (BYVAL nTrack AS LONG) AS LONG
   DECLARE FUNCTION GetAllTracksLengthString () AS CWSTR
   DECLARE FUNCTION GetAllTracksLength () AS LONG
   DECLARE FUNCTION GetCurrentPosString () AS CWSTR
   DECLARE FUNCTION GetCurrentPos () AS LONG
   DECLARE FUNCTION ToStart () AS MCIERROR
   DECLARE FUNCTION ToEnd () AS MCIERROR
   DECLARE FUNCTION Forward () AS BOOLEAN
   DECLARE FUNCTION Backward () AS BOOLEAN
   DECLARE FUNCTION GetTrackStartTimeString (BYVAL nTrack AS LONG) AS CWSTR
   DECLARE FUNCTION GetTrackStartTime (BYVAL nTrack AS LONG) AS LONG

END TYPE
' ########################################################################################

' ========================================================================================
' Constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CCDAudio
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR CCDAudio
   ' // Stop and close the audio CD
   DIM wszBuff AS WSTRING * 64
   m_dwError = mciSendStringW("status cdaudio mode wait", @wszBuff, 64, NULL)
   IF m_dwError <> 0 THEN EXIT DESTRUCTOR
   IF UCASE(wszBuff) = UCASE("not ready") THEN EXIT DESTRUCTOR
   IF UCASE(wszBuff) = UCASE("paused") OR UCASE(wszBuff) = UCASE("playing") THEN
      m_dwError = mciSendStringW("stop cdaudio", NULL, 0, NULL)
   END IF
   m_dwError = mciSendStringW("close cdaudio", NULL, 0, NULL)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last result code.
' ========================================================================================
PRIVATE FUNCTION CCDAudio.GetLastError () AS MCIERROR
   RETURN m_dwError
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last result code.
' ========================================================================================
PRIVATE FUNCTION CCDAudio.SetError (BYVAL dwError AS MCIERROR) AS MCIERROR
   m_dwError = dwError
   RETURN m_dwError
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves a string that describes the specified MCI error code.
' ========================================================================================
PRIVATE FUNCTION CCDAudio.GetErrorString (BYVAL dwError AS MCIERROR = 0) AS CWSTR
   DIM wszBuff AS WSTRING * 128
   IF dwError = 0 THEN dwError = m_dwError
   IF dwError = 0 THEN RETURN ""
   mciGetErrorStringW(dwError, @wszBuff, 128)
   RETURN wszBuff
END FUNCTION
' ========================================================================================

' ========================================================================================
' Opens the audio CD
' ========================================================================================
PRIVATE FUNCTION CCDAudio.Open () AS MCIERROR
   RETURN this.SetError(mciSendStringW("open cdaudio", NULL, 0, NULL))
END FUNCTION
' ========================================================================================

' =====================================================================================
' Closes the audio CD
' =====================================================================================
PRIVATE FUNCTION CCDAudio.Close () AS MCIERROR
   this.SetError(0)
   IF this.IsReady THEN
      IF this.IsPaused OR this.IsPlaying THEN mciSendStringW("stop cdaudio", NULL, 0, NULL)
      RETURN this.SetError(mciSendStringW("close cdaudio", NULL, 0, NULL))
   END IF
END FUNCTION
' =====================================================================================

' =====================================================================================
' Checks if the CD is ready.
' =====================================================================================
PRIVATE FUNCTION CCDAudio.IsReady () AS BOOLEAN
   DIM wszBuff AS WSTRING * 64
   this.SetError(mciSendStringW("status cdaudio mode wait", @wszBuff, 64, NULL))
   IF UCASE(wszBuff) = UCASE("not ready") THEN RETURN FALSE
   RETURN TRUE
END FUNCTION
' =====================================================================================

' =====================================================================================
' Checks whether CD media is inserted.
' =====================================================================================
PRIVATE FUNCTION CCDAudio.IsMediaInserted () AS BOOLEAN
   DIM wszBuff AS WSTRING * 16
   this.SetError(mciSendStringW("status cdaudio media present", @wszBuff, 16, NULL))
   IF UCASE(wszBuff) <> UCASE("true") THEN RETURN FALSE
   RETURN TRUE
END FUNCTION
' =====================================================================================

' =====================================================================================
' Checks whether is in paused mode.
' =====================================================================================
PRIVATE FUNCTION CCDAudio.IsPaused () AS BOOLEAN
   DIM wszBuff AS WSTRING * 64
   this.SetError(mciSendStringW("status cdaudio mode wait", @wszBuff, 64, NULL))
   IF UCASE(wszBuff) <> UCASE("paused") THEN RETURN FALSE
   RETURN TRUE
END FUNCTION
' =====================================================================================

' =====================================================================================
' Checks whether is in stopped mode.
' =====================================================================================
PRIVATE FUNCTION CCDAudio.IsStopped () AS BOOLEAN
   DIM wszBuff AS WSTRING * 64
   this.SetError(mciSendStringW("status cdaudio mode wait", @wszBuff, 64, NULL))
   IF UCASE(wszBuff) <> UCASE("stopped") THEN RETURN FALSE
   RETURN TRUE
END FUNCTION
' =====================================================================================

' =====================================================================================
' Checks whether is in play mode.
' =====================================================================================
PRIVATE FUNCTION CCDAudio.IsPlaying () AS BOOLEAN
   DIM wszBuff AS WSTRING * 64
   this.SetError(mciSendStringW("status cdaudio mode wait", @wszBuff, 64, NULL))
   IF UCASE(wszBuff) <> UCASE("playing") THEN RETURN FALSE
   RETURN TRUE
END FUNCTION
' =====================================================================================

' =====================================================================================
' Checks whether is in seeking mode.
' =====================================================================================
PRIVATE FUNCTION CCDAudio.IsSeeking () AS BOOLEAN
   DIM wszBuff AS WSTRING * 64
   this.SetError(mciSendStringW("status cdaudio mode wait", @wszBuff, 64, NULL))
   IF UCASE(wszBuff) <> UCASE("seeking") THEN RETURN FALSE
   RETURN TRUE
END FUNCTION
' =====================================================================================

' =====================================================================================
' Opens the CDRom door.
' =====================================================================================
PRIVATE FUNCTION CCDAudio.OpenDoor () AS MCIERROR
   RETURN this.SetError(mciSendStringW("set cdaudio door open", NULL, 0, NULL))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Closes the CDRom door.
' =====================================================================================
PRIVATE FUNCTION CCDAudio.CloseDoor () AS MCIERROR
   RETURN this.SetError(mciSendStringW("set cdaudio door closed", NULL, 0, NULL))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Starts playing CD Audio.
' =====================================================================================
PRIVATE FUNCTION CCDAudio.Play () AS MCIERROR
   RETURN this.SetError(mciSendStringW("play cdaudio", NULL, 0, NULL))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Starts playing CD Audio on the given track.
' =====================================================================================
PRIVATE FUNCTION CCDAudio.PlayFrom (BYVAL nTrack AS LONG) AS MCIERROR
   RETURN this.SetError(mciSendStringW("play cdaudio from " & WSTR(nTrack), NULL, 0, NULL))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Starts playing CD Audio from a given track to a given track.
' =====================================================================================
PRIVATE FUNCTION CCDAudio.PlayFromTo (BYVAL nStartTrack AS LONG, BYVAL nEndTrack AS LONG) AS MCIERROR
   RETURN this.SetError(mciSendStringW("play cdaudio from " & WSTR(nStartTrack) & " to " & WSTR(nEndTrack), NULL, 0, NULL))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Stops playing CD Audio.
' =====================================================================================
PRIVATE FUNCTION CCDAudio.Stop () AS MCIERROR
   mciSendStringW("stop cdaudio", NULL, 0, NULL)
   RETURN this.SetError(mciSendStringW("seek cdaudio to start", NULL, 0, NULL))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Pauses playing CD Audio.
' =====================================================================================
PRIVATE FUNCTION CCDAudio.Pause () AS MCIERROR
   RETURN this.SetError(mciSendStringW("pause cdaudio", NULL, 0, NULL))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the count of tracks.
' =====================================================================================
PRIVATE FUNCTION CCDAudio.GetTracksCount () AS LONG
   DIM wszBuff AS WSTRING * 16
   this.SetError(mciSendStringW("status cdaudio number of tracks", @wszBuff, 16, NULL))
   RETURN VAL(wszBuff)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the current track number.
' =====================================================================================
PRIVATE FUNCTION CCDAudio.GetCurrentTrack () AS LONG
   DIM wszBuff AS WSTRING * 16
   this.SetError(mciSendStringW("status cdaudio current track", @wszBuff, 16, NULL))
   RETURN VAL(wszBuff)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the length of the given track.
' =====================================================================================
PRIVATE FUNCTION CCDAudio.GetTrackLengthString (BYVAL nTrack AS LONG) AS CWSTR
   DIM wszBuff AS WSTRING * 64
   IF this.IsMediaInserted  = FALSE THEN RETURN ""
   this.SetError(mciSendStringW("status cdaudio length track " & WSTR(nTrack), @wszBuff, 64, NULL))
   RETURN wszBuff
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the length in seconds of the given track.
' =====================================================================================
PRIVATE FUNCTION CCDAudio.GetTrackLength (BYVAL nTrack AS LONG) AS LONG
   DIM wszBuff AS WSTRING * 64
   DIM wszMin AS WSTRING * 3
   DIM wszSec AS WSTRING * 3
   IF this.IsMediaInserted = FALSE THEN RETURN 0
   this.SetError(mciSendStringW("status cdaudio length track " & WSTR(nTrack), @wszBuff, 64, NULL))
   wszMin = LEFT(wszBuff, 2)
   wszSec = MID(wszBuff, 4, 2)
   RETURN VAL(wszMin) * 60 + VAL(wszSec)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the total length of all the tracks.
' =====================================================================================
PRIVATE FUNCTION CCDAudio.GetAllTracksLengthString () AS CWSTR
   DIM wszBuff AS WSTRING * 16
   IF this.IsMediaInserted = FALSE THEN RETURN ""
   this.SetError(mciSendStringW("status cdaudio length", @wszBuff, 16, NULL))
   RETURN wszBuff
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the total length in seconds of all the tracks.
' =====================================================================================
PRIVATE FUNCTION CCDAudio.GetAllTracksLength () AS LONG
   DIM wszBuff AS WSTRING * 16
   DIM wszMin AS WSTRING * 3
   DIM wszSec AS WSTRING * 3
   IF this.IsMediaInserted = FALSE THEN RETURN 0
   this.SetError(mciSendStringW("status cdaudio length", @wszBuff, 16, NULL))
   wszMin = LEFT(wszBuff, 2)
   wszSec = MID(wszBuff, 4, 2)
   RETURN VAL(wszMin) * 60 + VAL(wszSec)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the current track position.
' =====================================================================================
PRIVATE FUNCTION CCDAudio.GetCurrentPosString () AS CWSTR
   DIM wszBuff AS WSTRING * 16
   IF this.IsMediaInserted = FALSE THEN RETURN ""
   this.SetError(mciSendStringW("status cdaudio position", @wszBuff, 16, NULL))
   RETURN wszBuff
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the current track position in seconds.
' =====================================================================================
PRIVATE FUNCTION CCDAudio.GetCurrentPos () AS LONG
   DIM wszBuff AS WSTRING * 16
   DIM wszMin AS WSTRING * 3
   DIM wszSec AS WSTRING * 3
   IF this.IsMediaInserted = FALSE THEN RETURN 0
   this.SetError(mciSendStringW("status cdaudio position", @wszBuff, 16, NULL))
   wszMin = LEFT(wszBuff, 2)
   wszSec = MID(wszBuff, 4, 2)
   RETURN VAL(wszMin) * 60 + VAL(wszSec)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the position to the start of the audio CD.
' =====================================================================================
PRIVATE FUNCTION CCDAudio.ToStart () AS MCIERROR
   RETURN this.SetError(mciSendStringW("seek cdaudio to start", NULL, 0, NULL))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the position to the end of the audio CD.
' =====================================================================================
PRIVATE FUNCTION CCDAudio.ToEnd () AS MCIERROR
   RETURN this.SetError(mciSendStringW("seek cdaudio to end", NULL, 0, NULL))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Moves to the next track.
' =====================================================================================
PRIVATE FUNCTION CCDAudio.Forward () AS BOOLEAN
   DIM wszBuff AS WSTRING * 32
   DIM nCurrTrack AS LONG = this.GetCurrentTrack
   DIM nLastTrack AS LONG = this.GetTracksCount
   IF nCurrTrack = nLastTrack OR nLastTrack = 0 THEN RETURN FALSE
   nCurrTrack += 1
   this.SetError(mciSendStringW("status cdaudio position track " & WSTR(nCurrTrack), @wszBuff, 32, NULL))
   IF m_dwError THEN RETURN FALSE
   wszBuff = "play cdaudio from " & wszBuff
   this.SetError(mciSendStringW(wszBuff, NULL, 0, NULL))
   IF m_dwError = 0 THEN RETURN TRUE
END FUNCTION
' =====================================================================================

' =====================================================================================
' Moves to the previous track.
' =====================================================================================
PRIVATE FUNCTION CCDAudio.Backward () AS BOOLEAN
   DIM wszBuff AS WSTRING * 32
   DIM nCurrTrack AS LONG = this.GetCurrentTrack
   DIM nLastTrack AS LONG = this.GetTracksCount
   IF nCurrTrack < 2 OR nLastTrack = 0 THEN RETURN 0
   nCurrTrack -=1
   this.SetError(mciSendStringW("status cdaudio position track " & WSTR(nCurrTrack), @wszBuff, 32, NULL))
   IF m_dwError THEN RETURN FALSE
   wszBuff = "play cdaudio from " & wszBuff
   this.SetError(mciSendStringW(wszBuff, NULL, 0, NULL))
   IF m_dwError = 0 THEN RETURN TRUE
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the start time of the given track.
' =====================================================================================
PRIVATE FUNCTION CCDAudio.GetTrackStartTimeString (BYVAL nTrack AS LONG) AS CWSTR
   DIM wszBuff AS WSTRING * 64
   IF this.IsMediaInserted = FALSE THEN RETURN ""
   this.SetError(mciSendStringW("status cdaudio position track " & WSTR(nTrack), @wszBuff, 64, NULL))
   RETURN wszBuff
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the start time of the given track.
' =====================================================================================
PRIVATE FUNCTION CCDAudio.GetTrackStartTime (BYVAL nTrack AS LONG) AS LONG
   DIM wszBuff AS WSTRING * 64
   DIM wszMin AS WSTRING * 3
   DIM wszSec AS WSTRING * 3
   IF this.IsMediaInserted = FALSE THEN RETURN 0
   this.SetError(mciSendStringW("status cdaudio position track " & WSTR(nTrack), @wszBuff, 64, NULL))
   wszMin = LEFT(wszBuff, 2)
   wszSec = MID(wszBuff, 4, 2)
   RETURN VAL(wszMin) * 60 + VAL(wszSec)
END FUNCTION
' =====================================================================================

END NAMESPACE
