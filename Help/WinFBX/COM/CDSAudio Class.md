# CDSAudio Class

The **CDSAudio** class allows to play audio files of a variety of formats using **Direct Show**.

**Include file**: CDSAudio.inc

### Constructor

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#Constructor) | Creates an instance of the **CDSAudio** class and loads the specified audio file. |

### Methods

| Name       | Description |
| ---------- | ----------- |
| [GetBalance](#GetBalance) | Gets the balance for the audio signal. |
| [CurrentPosition](#CurrentPosition) | Gets the current position, relative to the total duration of the stream. |
| [GetDuration](#GetDuration) | Gets the duration of the stream, in 100-nanosecond units. |
| [GetEvent](#GetEvent) | Retrieves the next event notification from the event queue. |
| [GetVolume](#GetVolume) | Gets the volume (amplitude) of the audio signal. |
| [Load](#Load) | Builds a filter graph that renders the specified file. |
| [Pause](#Pause) | Pauses all the filters in the filter graph. |
| [Run](#Run) | Runs all the filters in the filter graph. |
| [SetBalance](#SetBalance) | Sets the balance for the audio signal. |
| [SetNotifyWindow](#SetNotifyWindow) | Registers a window to process event notifications. |
| [SetPositions](#SetPositions) | Sets the current position and the stop position. |
| [SetVolume](#SetVolume) | Sets the volume (amplitude) of the audio signal. |
| [Stop](#Stop) | Stops all the filters in the filter graph. |
| [WaitForCompletion](#WaitForCompletion) | Waits for the filter graph to render all available data. |

# <a name="Constructor"></a>Constructor

Creates an instance of the **CDSAudio** class and loads the specified audio file.

```
CONSTRUCTOR CDSAudio (BYREF wszFileName AS WSTRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFilename* | The file to load. |

You can create an instance of the class and load the file at the same time with:

```
DIM pCDSAudio AS CDSAudio = "MyAudioFile.mp3"
```

or you can use the default constructor and then load the file:

```
DIM pCDSAudio AS CDSAudio
pCDSAudio.Load("MyAudioFile.mp3")
```

With the **Load** method you can change the audio file on the fly.

To receive event messages, you can define a custom message

```
#define WM_GRAPHNOTIFY  WM_APP + 1
```

and pass it to the **SetNotifyWindow** method, together with the handle of the window that will process the message:

```
pCDSAudio.SetNotifyWindow(pWindow.hWindow, WM_GRAPHNOTIFY, 0)
```

The messages will be processed in the window callback procedure:

```
      CASE WM_GRAPHNOTIFY
         DIM AS LONG lEventCode
         DIM AS LONG_PTR lParam1, lParam2
         WHILE (SUCCEEDED(pCDSAudio.GetEvent(lEventCode, lParam1, lParam2)))
            SELECT CASE lEventCode
               CASE EC_COMPLETE:    ' Handle event
               CASE EC_USERABORT:    ' Handle event
               CASE EC_ERRORABORT:   ' Handle event
            END SELECT
         WEND
```

Once loaded, you can start playing the audio file calling the **Run** method:

```
pCDSAudio.Run
```

There are other methods to get/set the volume and balance, to get the duration and current position, to set the positions, and to pause or stop.

# <a name="GetBalance"></a>GetBalance

Gets the balance for the audio signal.

```
FUNCTION GetBalance () AS LONG
```

#### Return value

The balance for the audio signal.

#### Remarks

The balance ranges from -10,000 to 10,000. The value -10,000 means the right channel is attenuated by 100 dB and is effectively silent. The value 10,000 means the left channel is silent. The neutral value is 0, which means that both channels are at full volume. When one channel is attenuated, the other remains at full volume.

# <a name="CurrentPosition"></a>CurrentPosition

Gets the current position, relative to the total duration of the stream.

```
FUNCTION CurrentPosition () AS LONG
```

# <a name="GetDuration"></a>GetDuration

Gets the duration of the stream, in 100-nanosecond units.

```
FUNCTION GetDuration () AS LONG
```

# <a name="GetEvent"></a>GetEvent

Retrieves the next event notification from the event queue.

```
FUNCTION GetEvent(BYREF lEventCode AS LONG, BYREF lParam1 AS LONG_PTR, _
   BYREF lParam2 AS LONG_PTR, BYVAL msTimeout AS LONG = 0) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *lEventCode* | Out. Variable that receives the event code. |
| *lParam1* | Out. Variable that receives the first event parameter. |
| *lParam2* | Out. Variable that receives the second event parameter. |
| *msTimeout* | In. Time-out interval, in milliseconds. Use INFINITE to block until there is an event. |

| Result code | Description |
| ----------- | ----------- |
| S_OK | Success. |
| E_ABORT | Timeout expired. |
| E_POINTER | The **IMediaEventEx** interface pointer is null. |

# <a name="GetVolume"></a>GetVolume

Gets the volume (amplitude) of the audio signal.

```
FUNCTION GetVolume () AS LONG
```

#### Return value

The volume (amplitude) of the audio signal. Specifies the volume, as a number from –10,000 to 0, inclusive. Full volume is 0, and –10,000 is silence. Multiply the desired decibel level by 100. For example, –10,000 = –100 dB.

# <a name="Load"></a>Load

Builds a filter graph that renders the specified file.

```
FUNCTION Load (BYREF wszFileName AS WSTRING) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFilename* | The file to load. |

#### Return value

| Result code | Description |
| ----------- | ----------- |
| S_OK | Success. |
| VFW_S_AUDIO_NOT_RENDERED | Partial success; the audio was not rendered. |
| VFW_S_DUPLICATE_NAME | Success; the Filter Graph Manager modified the filter name to avoid duplication. |
| VFW_S_PARTIAL_RENDER | Some of the streams in this movie are in an unsupported format. |
| VFW_S_VIDEO_NOT_RENDERED | Partial success; some of the streams in this movie are in an unsupported format. |
| E_ABORT | Operation aborted. |
| E_FAIL | Failure. |
| E_INVALIDARG | Argument is invalid. |
| E_OUTOFMEMORY | Insufficient memory. |
| E_POINTER | NULL pointer argument.. |
| VFW_E_CANNOT_CONNECT | No combination of intermediate filters could be found to make the connection.. |
| VFW_E_CANNOT_LOAD_SOURCE_FILTER | The source filter for this file could not be loaded.. |
| VFW_E_CANNOT_RENDER | No combination of filters could be found to render the stream.. |
| VFW_E_INVALID_FILE_FORMAT | The file format is invalid.. |
| VFW_E_NOT_FOUND | An object or name was not found.. |
| VFW_E_UNKNOWN_FILE_TYPE | The media type of this file is not recognized.. |
| VFW_E_UNSUPPORTED_STREAM | Cannot play back the file: the format is not supported.. |

# <a name="Pause"></a>Pause

Pauses all the filters in the filter graph.

```
FUNCTION Pause () AS HRESULT
```

#### Return value

| Result code | Description |
| ----------- | ----------- |
| S_OK | All filters in the graph completed the transition to a paused state. |
| S_FALSE | The graph paused successfully, but some filters have not completed the state transition. |
| E_POINTER | The **IMediaControl** interface pointer is null. |

# <a name="Run"></a>Run

Runs all the filters in the filter graph.

```
FUNCTION Run () AS HRESULT
```

#### Return value

| Result code | Description |
| ----------- | ----------- |
| S_OK | All filters in the graph completed the transition to a running state. |
| S_FALSE | The graph is preparing to run, but some filters have not completed the transition to a running state. |
| E_POINTER | The **IMediaControl** interface pointer is null. |

# <a name="SetBalance"></a>SetBalance

Sets the balance for the audio signal.

```
FUNCTION SetBalance (BYVAL nBalance AS LONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *nBalance* | The balance ranges from -10,000 to 10,000. The value -10,000 means the right channel is attenuated by 100 dB and is effectively silent. The value 10,000 means the left channel is silent. The neutral value is 0, which means that both channels are at full volume. When one channel is attenuated, the other remains at full volume. |

#### Return value

| Result code | Description |
| ----------- | ----------- |
| S_OK | Success. |
| E_FAIL | The underlying audio device returned an error. |
| E_INVALIDARG | The value of *nBalance* is invalid. |
| E_NOTIMPL | The filter graph does not contain an audio renderer filter. (Possibly the source does not contain an audio stream.) |
| E_POINTER | The **IBasicAudio** interface pointer is null. |

# <a name="SetNotifyWindow"></a>SetNotifyWindow

Sets the balance for the audio signal.

```
FUNCTION SetNotifyWindow (BYVAL hwnd AS HWND, BYVAL lMsg AS LONG, BYVAL lParam AS LONG_PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |
| *lMsg* | Window message to be passed as the notification. |
| *lInstanceData* | Value to be passed as the lParam parameter for the lMsg message. |

#### Return value

Returns S_OK if successful or E_INVALIDARG if the hwnd parameter is not a valid handle to a window.

# <a name="SetPositions"></a>SetPositions

Sets the current position and the stop position.

```
FUNCTION SetPositions (BYREF pCurrent AS LONGLONG, BYREF pStop AS LONGLONG, _
   BYVAL bAbsolutePositioning AS BOOLEAN) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pCurrent* | In/Out. Pointer to a variable that specifies the current position, in units of the current time format. |
| *pStop* | In/Out. Pointer to a variable that specifies the stop time, in units of the current time format. |
| *lInstanceData* | In. TRUE or FALSE. If TRUE, the specified position is absolute. If FALSE, the specified position is relative. |

#### Return value

| Result code | Description |
| ----------- | ----------- |
| S_OK | Success. |
| S_FALSE | No position change. (Both flags specify no seeking.) |
| E_INVALIDARG | Invalid argument. |
| E_NOTIMPL | Method is not supported. |
| E_POINTER | NULL pointer argument. |

# <a name="SetVolume"></a>SetVolume

Sets the volume (amplitude) of the audio signal.

```
FUNCTION SetVolume (BYVAL nVol AS LONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *nVol* | The volume (amplitude) of the audio signal. Specifies the volume, as a number from –10,000 to 0, inclusive. Full volume is 0, and –10,000 is silence. Multiply the desired decibel level by 100. For example, –10,000 = –100 dB. |


# <a name="Stop"></a>Stop

Stops all the filters in the filter graph.

```
FUNCTION Stop () AS HRESULT
```

#### Return value

Returns S_OK if successful, or an HRESULT value that indicates the cause of the error.

# <a name="WaitForCompletion"></a>WaitForCompletion

Waits for the filter graph to render all available data.

```
FUNCTION WaitForCompletion (BYVAL msTimeout AS LONG, BYREF EvCode AS LONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *msTimeout* | Duration of the time-out, in milliseconds. Pass zero to return immediately. To block indefinitely, pass INFINITE. |
| *EvCode* | Out. Event that terminated the wait. This value can be one of the following:<br>**EC_COMPLETE** : Operation completed.<br>**EC_ERRORABORT** : Error. Playback cannot continue.<br>**EC_USERABORT** : User terminated the operation.<br>**Zero** (0) = Operation has not completed. |

#### Return value

| Result code | Description |
| ----------- | ----------- |
| S_OK | Success. |
| E_ABORT | Time-out expired. |
| E_POINTER | The **IMediaEventEx** interface is null. |
| VFW_E_WRONG_STATE | The filter graph is not running. |

#### Remarks

This method blocks until the time-out expires, or one of the following events occurs:

**EC_COMPLETE**<br>
**EC_ERRORABORT**<br>
**EC_USERABORT**

During the wait, the method discards all other event notifications.

If the return value is S_OK, the *EvCode* parameter receives the event code that ended the wait. When the method returns, the filter graph is still running. The application can pause or stop the graph, as appropriate.
