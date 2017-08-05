' ########################################################################################
' Microsoft Windows
' File: CDSAudio.inc
' Contents: Direct Show audio class.
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' Based on Simple C++ DirectShow MP3 Player Class by Wong Shao Voon, 26 Apr 2012.
' http://www.codeproject.com/Articles/373613/Simple-Cplusplus-DirectShow-MP3-Player-Class
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "win/uuids.bi"
#include once "win/strmif.bi"
#include once "win/evcode.bi"
#include once "win/control.bi"

NAMESPACE Afx

' ========================================================================================
' CDSAudio class
' ========================================================================================
TYPE CDSAudio

   Public:
      m_pIGraphBuilder AS IGraphBuilder PTR
      m_pIMediaControl AS IMediaControl PTR
      m_pIMediaEventEx AS IMediaEventEx PTR
      m_pIBasicAudio AS IBasicAudio PTR
      m_pIMediaSeeking AS IMediaSeeking PTR
      m_duration AS LONGLONG

   Private:
      m_bUninitCOM AS BOOLEAN

   Public:
      DECLARE CONSTRUCTOR
      DECLARE CONSTRUCTOR (BYREF wszFileName AS WSTRING)
      DECLARE DESTRUCTOR
      DECLARE SUB Cleanup
      DECLARE FUNCTION Load (BYREF wszFileName AS WSTRING) AS HRESULT
      DECLARE FUNCTION Run () AS HRESULT
      DECLARE FUNCTION Pause () AS HRESULT
      DECLARE FUNCTION Stop () AS HRESULT
      DECLARE FUNCTION WaitForCompletion (BYVAL msTimeout AS LONG, BYREF EvCode AS LONG) AS HRESULT
      DECLARE FUNCTION GetVolume () AS LONG
      DECLARE FUNCTION SetVolume (BYVAL nVol AS LONG) AS HRESULT
      DECLARE FUNCTION GetBalance () AS LONG
      DECLARE FUNCTION SetBalance (BYVAL nVol AS LONG) AS HRESULT
      DECLARE FUNCTION GetDuration () AS LONGLONG
      DECLARE FUNCTION GetCurrentPosition () AS LONGLONG
      DECLARE FUNCTION SetPositions (BYREF pCurrent AS LONGLONG, BYREF pStop AS LONGLONG, BYVAL bAbsolutePositioning AS BOOLEAN) AS HRESULT
      DECLARE FUNCTION SetNotifyWindow (BYVAL hwnd AS HWND, BYVAL lMsg AS LONG, BYVAL lParam AS LONG_PTR) AS HRESULT
      DECLARE FUNCTION GetEvent(BYREF lEventCode AS LONG, BYREF lParam1 AS LONG_PTR, BYREF lParam2 AS LONG_PTR, BYVAL msTimeout AS LONG = 0) AS HRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' CDSAudio constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CDSAudio
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CDSAudio (BYREF wszFileName AS WSTRING)
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Load the file
   this.Load(wszFileName)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CDSAudio destructor
' ========================================================================================
PRIVATE DESTRUCTOR CDSAudio
   IF m_pIMediaEventEx THEN SetNotifyWindow(NULL, 0, NULL)   ' // Cancel event notification
   this.Cleanup
   ' // Free the COM library
   IF m_bUninitCOM THEN CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Cleanup. Used internally.
' ========================================================================================
PRIVATE SUB CDSAudio.Cleanup
   IF m_pIGraphBuilder THEN IGraphBuilder_Release(m_pIGraphBuilder) : m_pIGraphBuilder = NULL
   IF m_pIMediaControl THEN IMediaControl_Release(m_pIMediaControl) : m_pIMediaControl = NULL
   IF m_pIMediaEventEx THEN IMediaEventEx_Release(m_pIMediaEventEx) : m_pIMediaEventEx = NULL
   IF m_pIBasicAudio   THEN IBasicAudio_Release(m_pIBasicAudio)     : m_pIBasicAudio   = NULL
   IF m_pIMediaSeeking THEN IMediaSeeking_Release(m_pIMediaSeeking) : m_pIMediaSeeking = NULL
END SUB
' ========================================================================================

' ========================================================================================
' Builds a filter graph that renders the specified file.
' Return code:
'   S_OK = Success.
'   VFW_S_AUDIO_NOT_RENDERED = Partial success; the audio was not rendered.
'   VFW_S_DUPLICATE_NAME = Success; the Filter Graph Manager modified the filter name to avoid duplication.
'   VFW_S_PARTIAL_RENDER = Some of the streams in this movie are in an unsupported format.
'   VFW_S_VIDEO_NOT_RENDERED = Partial success; some of the streams in this movie are in an unsupported format.
'   E_ABORT = Operation aborted.
'   E_FAIL = Failure.
'   E_INVALIDARG = Argument is invalid.
'   E_OUTOFMEMORY = Insufficient memory.
'   E_POINTER = NULL pointer argument.
'   VFW_E_CANNOT_CONNECT = No combination of intermediate filters could be found to make the connection.
'   VFW_E_CANNOT_LOAD_SOURCE_FILTER = The source filter for this file could not be loaded.
'   VFW_E_CANNOT_RENDER = No combination of filters could be found to render the stream.
'   VFW_E_INVALID_FILE_FORMAT = The file format is invalid.
'   VFW_E_NOT_FOUND = An object or name was not found.
'   VFW_E_UNKNOWN_FILE_TYPE = The media type of this file is not recognized.
'   VFW_E_UNSUPPORTED_STREAM = Cannot play back the file: the format is not supported.
' ========================================================================================
PRIVATE FUNCTION CDSAudio.Load (BYREF wszFileName AS WSTRING) AS HRESULT
   DIM hr AS HRESULT
   this.Cleanup
   hr = CoCreateInstance(@CLSID_FilterGraph, NULL, CLSCTX_INPROC_SERVER, @IID_IGraphBuilder, @m_pIGraphBuilder)
   IF hr <> S_OK THEN RETURN hr
   hr = IGraphBuilder_QueryInterface(m_pIGraphBuilder, @IID_IMediaControl, @m_pIMediaControl)
   IF hr <> S_OK THEN this.Cleanup : RETURN hr
   hr = IGraphBuilder_QueryInterface(m_pIGraphBuilder, @IID_IMediaEventEx, @m_pIMediaEventEx)
   IF hr <> S_OK THEN this.Cleanup : RETURN hr
   hr = IGraphBuilder_QueryInterface(m_pIGraphBuilder, @IID_IBasicAudio, @m_pIBasicAudio)
   IF hr <> S_OK THEN this.Cleanup : RETURN hr
   hr = IGraphBuilder_QueryInterface(m_pIGraphBuilder, @IID_IMediaSeeking, @m_pIMediaSeeking)
   IF hr <> S_OK THEN this.Cleanup : RETURN hr
   hr = IGraphBuilder_RenderFile(m_pIGraphBuilder, @wszFileName, NULL)
   IF FAILED(hr) THEN this.Cleanup : RETURN hr
   hr = IMediaSeeking_SetTimeFormat(m_pIMediaSeeking, @TIME_FORMAT_MEDIA_TIME)   ' // Reference time (100-nanosecond units).
   IF hr <> S_OK THEN this.Cleanup : RETURN hr
   hr = IMediaSeeking_GetDuration(m_pIMediaSeeking, @m_duration)                 ' // Returns 10,000,000 for a second.
   IF hr <> S_OK THEN this.Cleanup : RETURN hr
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Runs all the filters in the filter graph. While the graph is running, data moves through
' the graph and is rendered.
' Return code:
'   E_POINTER Invalid interface pointer.
'   S_FALSE The graph is preparing to run, but some filters have not completed the transition to a running state.
'   S_OK = All filters in the graph completed the transition to a running state.
' ========================================================================================
PRIVATE FUNCTION CDSAudio.Run () AS HRESULT
   IF m_pIMediaControl = NULL THEN RETURN E_POINTER
   RETURN IMediaControl_Run(m_pIMediaControl)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Pauses all the filters in the filter graph.
' Result code:
'   E_POINTER Invalid interface pointer.
'   S_FALSE = The graph paused successfully, but some filters have not completed the state transition.
'   S_OK = All filters in the graph completed the transition to a paused state.
' ========================================================================================
PRIVATE FUNCTION CDSAudio.Pause () AS HRESULT
   IF m_pIMediaControl = NULL THEN RETURN E_POINTER
   RETURN IMediaControl_Pause(m_pIMediaControl)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Stops all the filters in the filter graph.
' Result code:
'    Returns S_OK if successful, or an HRESULT value that indicates the cause of the error.
' ========================================================================================
PRIVATE FUNCTION CDSAudio.Stop () AS LONG
   IF m_pIMediaControl = NULL THEN RETURN E_POINTER
   RETURN IMediaControl_Stop(m_pIMediaControl)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Waits for the filter graph to render all available data.
' Parameters:
' msTimeout
'   [in] Duration of the time-out, in milliseconds. Pass zero to return immediately.
'   To block indefinitely, pass INFINITE (&hffffffff).
' EvCode
'   [out] Event that terminated the wait. This value can be one of the following:
'   EC_COMPLETE  = Operation completed.
'   EC_ERRORABORT = Error. Playback cannot continue.
'   EC_USERABORT  = User terminated the operation.
'   Zero (0) = Operation has not completed.
' Return Value = %TRUE or %FALSE.
' Return code:
'   S_OK = Success.
'   E_POINTER Invalid interface pointer.
'   E_ABORT = Time-out expired.
'   VFW_E_WRONG_STATE = The filter graph is not running.
' Remarks
' This method blocks until the time-out expires, or one of the following events occurs:
'   EC_COMPLETE
'   EC_ERRORABORT
'   EC_USERABORT
' During the wait, the method discards all other event notifications.
' If the return value is S_OK, the EvCode parameter receives the event code that ended
' the wait. When the method returns, the filter graph is still running. The application
' can pause or stop the graph, as appropriate.
' ========================================================================================
PRIVATE FUNCTION CDSAudio.WaitForCompletion (BYVAL msTimeout AS LONG, BYREF EvCode AS LONG) AS HRESULT
   IF m_pIMediaEventEx = NULL THEN RETURN E_POINTER
   RETURN IMediaEventEx_WaitForcompletion(m_pIMediaEventEx, msTimeout, @EvCode)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the volume (amplitude) of the audio signal.
' Return Value = The volume. Divide by 100 to get equivalent decibel value.
'                For example, –10,000 is –100 dB.
' If it fails, it returns 0.
' ========================================================================================
PRIVATE FUNCTION CDSAudio.GetVolume () AS LONG
   IF m_pIBasicAudio = NULL THEN RETURN 0
   DIM nVol AS LONG
   IBasicAudio_get_Volume(m_pIBasicAudio, @nVol)
   RETURN nVol
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the volume (amplitude) of the audio signal.
' nVol = Specifies the volume, as a number from –10,000 to 0, inclusive. Full volume is 0, and
'        –10,000 is silence. Multiply the desired decibel level by 100. For example, –10,000 = –100 dB.
' Return code:
'   S_OK = Success.
'   E_POINTER Invalid interface pointer.
'   E_FAIL = The underlying audio device returned an error.
'   E_INVALIDARG = The value of nVol is invalid.
'   E_NOTIMPL = The filter graph does not contain an audio renderer filter.
'               (Possibly the source does not contain an audio stream.)
' ========================================================================================
PRIVATE FUNCTION CDSAudio.SetVolume (BYVAL nVol AS LONG) AS HRESULT
   IF m_pIBasicAudio = NULL THEN RETURN E_POINTER
   RETURN IBasicAudio_put_Volume(m_pIBasicAudio, nVol)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the balance of the audio signal.
' Return Value = The balance ranges from -10,000 to 10,000. The value -10,000 means the
'                right channel is attenuated by 100 dB and is effectively silent. The
'                value 10,000 means the left channel is silent. The neutral value is 0,
'                which means that both channels are at full volume. When one channel is
'                attenuated, the other remains at full volume.
' If it fails, it returns 0.
' ========================================================================================
PRIVATE FUNCTION CDSAudio.GetBalance () AS LONG
   IF m_pIBasicAudio = NULL THEN RETURN 0
   DIM nBalance AS LONG
   IBasicAudio_get_Balance(m_pIBasicAudio, @nBalance)
   RETURN nBalance
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the balance of the audio signal.
' nBalance = Specifies the balance. The value can range from -10,000 to 10,000. The value
'            -10,000 means the right channel is attenuated by 100 dB and is effectively
'            silent. The value 10,000 means the left channel is silent. The neutral value
'            is 0, which means that both channels are at full volume. When one channel is
'            attenuated, the other remains at full volume.
' Return code:
'   S_OK = Success.
'   E_POINTER Invalid interface pointer.
'   E_FAIL = The underlying audio device returned an error.
'   E_INVALIDARG = The value of nBalance is invalid.
'   E_NOTIMPL = The filter graph does not contain an audio renderer filter.
'               (Possibly the source does not contain an audio stream.)
' ========================================================================================
PRIVATE FUNCTION CDSAudio.SetBalance (BYVAL nBalance AS LONG) AS HRESULT
   IF m_pIBasicAudio = NULL THEN RETURN E_POINTER
   RETURN IBasicAudio_put_Balance(m_pIBasicAudio, nBalance)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the duration of the stream, in 100-nanosecond units.
' Returns 10,000,000 for a second.
' ========================================================================================
PRIVATE FUNCTION CDSAudio.GetDuration () AS LONGLONG
   RETURN m_duration
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current position, relative to the total duration of the stream.
' If it fails, it returns 0.
' ========================================================================================
PRIVATE FUNCTION CDSAudio.GetCurrentPosition () AS LONGLONG
   IF m_pIMediaSeeking = NULL THEN RETURN 0
   DIM curPos AS LONGLONG
   IMediaSeeking_GetCurrentPosition(m_pIMediaSeeking, @curPos)
   RETURN curPos
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the current position and the stop position.
' Parameters:
'   pCurrent = [in, out] Pointer to a variable that specifies the current position, in units of the current time format.
'   pStop = [in, out] Pointer to a variable that specifies the stop time, in units of the current time format.
'   bAbsolutePositioning = [in] TRUE or FALSE. IF TRUE, the specified position is absolute.
'                          IF FALSE, the specified position is relative.
' Result code:
'   S_FALSE = No position change. (Both flags specify no seeking.)
'   E_POINTER Invalid interface pointer.
'   S_OK = Success.
'   E_INVALIDARG = Invalid argument.
'   E_NOTIMPL = Method is not supported.
'   E_POINTER = NULL pointer argument.
' ========================================================================================
PRIVATE FUNCTION CDSAudio.SetPositions (BYREF pCurrent AS LONGLONG, BYREF pStop AS LONGLONG, BYVAL bAbsolutePositioning AS BOOLEAN) AS HRESULT
   IF m_pIMediaSeeking = NULL THEN RETURN E_POINTER
   DIM dwFlags AS DWORD
   IF bAbsolutePositioning THEN
      dwFlags = AM_SEEKING_AbsolutePositioning OR AM_SEEKING_SeekToKeyFrame
   ELSE
      dwFlags = AM_SEEKING_RelativePositioning OR AM_SEEKING_SeekToKeyFrame
   END IF
   RETURN IMediaSeeking_SetPositions(m_pIMediaSeeking, @pCurrent, dwFlags, @pStop, dwFlags)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Registers a window to process event notifications.
' Parameters:
'   hwnd = Window handle.
'   lMsg = Window message to be passed as the notification.
'   lInstanceData = Value to be passed as the lParam parameter for the lMsg message.
' Result code:
'   E_POINTER Invalid interface pointer.
'   Returns S_OK if successful or E_INVALIDARG if the hwnd parameter is not a valid handle
'   to a window.
' Using event notification:
'   An application can instruct the Filter Graph Manager to send a Windows message to a
'   designated window whenever a new event occurs. This enables the application to respond
'   inside the window's message loop. First, define the message that will be sent to the
'   application window. Applications can use message numbers in the range from WM_APP
'   through 0xBFFF as private message:
'     #define WM_GRAPHNOTIFY  WM_APP + 1
'   Next call the SetNotifyWindow method:
'     pCDSAudio.SetNotifyWindow(hwnd, WM_GRAPHNOTIFY, 0)
'   This method designates the specified window (hwnd) as the recipient of the message.
'   WM_GRAPHNOTIFY is an ordinary Windows message. Whenever the Filter Graph Manager puts
'   a new event on the event queue, it posts a WM_GRAPHNOTIFY message to the designated
'   application window. The message's lParam parameter is equal to the third parameter in
'   SetNotifyWindow. This parameter enables you to send instance data with the message.
'   The window message's wParam parameter is always zero.
'   In your application's WindowProc function, add a case statement for the WM_GRAPHNOTIFY message:
'     CASE WM_GRAPHNOTIFY
'        WHILE (SUCCEEDED(pCDSAudio.GetEvent(evCode, param1, param2)))
'           SELECT CASE evCode
'              CASE EC_COMPLETE:     ' Handle event
'              CASE EC_USERABORT:    ' Handle event
'              CASE EC_ERRORABORT:   ' Handle event
'           END SELECT
'        END WHILE
'   The GetEvent method retrieves the event code and the two event parameters.
'   Event notification and the message loop are both asynchronous, so the queue might hold
'   more than one event by the time your application responds to the message. Also, the
'   Filter Graph Manager can remove certain events from the queue, if they become invalid.
'   Therefore, you should call GetEvent until it returns a failure code, indicating the
'   queue is empty.
' ========================================================================================
PRIVATE FUNCTION CDSAudio.SetNotifyWindow (BYVAL hwnd AS HWND, BYVAL lMsg AS LONG, BYVAL lInstanceData AS LONG_PTR) AS HRESULT
   IF m_pIMediaEventEx = NULL THEN RETURN E_POINTER
   RETURN IMediaEventEx_SetNotifyWindow(m_pIMediaEventEx, cast(OAHWND, hwnd), lMsg, lInstanceData)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the next event notification from the event queue.
' Parameters:
'   lEventCode [out]: Variable that receives the event code.
'   lParam1 [out]   : Variable that receives the first event parameter.
'   lParam2 [out]   : Variable that receives the second event parameter.
'   msTimeout [in]  : Time-out interval, in milliseconds. Use INFINITE to block until there
'                     is an event.
' Return code:
'   S_OK Success.
'   E_ABORT Timeout expired.
'   E_POINTER Invalid interface pointer.
' ========================================================================================
PRIVATE FUNCTION CDSAudio.GetEvent(BYREF lEventCode AS LONG, BYREF lParam1 AS LONG_PTR, BYREF lParam2 AS LONG_PTR, BYVAL msTimeout AS LONG = 0) AS HRESULT
   IF m_pIMediaEventEx = NULL THEN RETURN E_POINTER
   DIM hr AS HRESULT = IMediaEventEx_GetEvent(m_pIMediaEventEx, @lEventCode, @lParam1, @lParam2, msTimeout)
   IF SUCCEEDED(hr) THEN IMediaEventEx_FreeEventParams(m_pIMediaEventEx, lEventCode, lParam1, lParam2)
   RETURN hr
END FUNCTION
' ========================================================================================


END NAMESPACE
