# CCDAudio Class

The **CCDAudio** class allows to play a CD Rom using MCI.

### Constructor

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#Constructor) | Creates an instance of the **CCDAudio** class. |

### Methods

| Name       | Description |
| ---------- | ----------- |
| [Backward](#Backward) | Moves to the previous track. |
| [Close](#Close) | Closes the device or file and any associated resources. MCI unloads a device when all instances of the device or all files are closed.  |
| [CloseDoor](#CloseDoor) | Closes the CDRom door. |
| [Forward](#Forward) | Moves to the next track. |
| [GetAllTracksLength](#GetAllTracksLength) | Returns the total length in seconds of all the tracks. |
| [GetAllTracksLengthString](#GetAllTracksLengthString) | Returns the total length of all the tracks. |
| [GetCurrentPos](#GetCurrentPos) | Returns the current track position in seconds. |
| [GetCurrentPosString](#GetCurrentPosString) | Returns the current track position. |
| [GetCurrentTrack](#GetCurrentTrack) | Returns the current track number. |
| [GetErrorString](#GetErrorString) | Retrieves a string that describes the specified MCI error code. |
| [GetLastError](#GetLastError) | Retrieves a The last MCI error code. |
| [GetTrackLength](#GetTrackLength) | Returns the length in seconds of the given track. |
| [GetTrackLengthString](#GetTrackLengthString) | Returns the length of the given track. |
| [GetTracksCount](#GetTracksCount) | Returns the count of tracks. |
| [GetTrackStartTime](#GetTrackStartTime) | Returns the start time of the given track. |
| [GetTrackStartTimeString](#GetTrackStartTimeString) | Returns the start time of the given track. |
| [IsMediaInserted](#IsMediaInserted) | Checks whether CD media is inserted. |
| [IsPaused](#IsPaused) | Checks whether is in paused mode. |
| [IsPlaying](#IsPlaying) | Checks whether is in play mode. |
| [IsReady](#IsReady) | Checks if the device is ready. |
| [IsSeeking](#IsSeeking) | Checks whether is in seeking mode. |
| [IsStopped](#IsStopped) | Checks whether is in stopped mode. |
| [Open](#Open) | Initializes the device. |
| [OpenDoor](#OpenDoor) | Opens the CDRom door. |
| [Pause](#Pause) | Pauses playing CD Audio. |
| [Play](#Play) | Starts playing CD Audio. |
| [PlayFrom](#PlayFrom) | Starts playing CD Audio on the given track. |
| [PlayFromTo](#PlayFromTo) | Starts playing CD Audio from a given track to a given track. |
| [Stop](#Stop) | Stops playing CD Audio. |
| [ToEnd](#ToEnd) | Sets the position to the end of the audio CD. |
| [ToStart](#ToStart) | Sets the position to the start of the audio CD. |


# <a name="Constructor"></a>Constructor

Creates an instance of the **CCDAudio** class.

```
CONSTRUCTOR CCDAudio
```

#### Usage example:

```
DIM pAudio AS CCDAudio
pAudio.Open
pAudio.Play
```

# <a name="Backward"></a>Backward

Moves to the previous track.

```
FUNCTION Backward () AS BOOLEAN
```

#### Return value

TRUE or FALSE.

# <a name="Close"></a>Close

Closes the device or file and any associated resources. MCI unloads a device when all instances of the device or all files are closed. 

```
FUNCTION Close () AS MCIERROR
```

#### Return value

Returns zero if successful or an error otherwise.

# <a name="CloseDoor"></a>CloseDoor

Closes the CDRom door.

```
FUNCTION CloseDoor () AS MCIERROR
```

#### Return value

Returns zero if successful or an error otherwise.

# <a name="Forward"></a>Forward

Moves to the next track.

```
FUNCTION Forward () AS BOOLEAN
```

#### Return value

TRUE or FALSE.

# <a name="GetAllTracksLength"></a>GetAllTracksLength

Returns the total length in seconds of all the tracks.

```
FUNCTION GetAllTracksLength () AS LONG
```

# <a name="GetAllTracksLengthString"></a>GetAllTracksLengthString

Returns the total length of all the tracks as a string.

```
FUNCTION GetAllTracksLengthString () AS CWSTR
```

# <a name="GetCurrentPos"></a>GetCurrentPos

Returns the current track position in seconds.

```
FUNCTION GetCurrentPos () AS LONG
```

# <a name="GetCurrentPosString"></a>GetCurrentPosString

Returns the current track position as a string.

```
FUNCTION GetCurrentPosString () AS CWSTR
```

# <a name="GetCurrentTrack"></a>GetCurrentTrack

Returns the current track number.

```
FUNCTION GetCurrentTrack () AS LONG
```

# <a name="GetErrorString"></a>GetErrorString

Retrieves a string that describes the specified MCI error code.

```
FUNCTION GetErrorString (BYVAL dwError AS MCIERROR = 0) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwError* | Optional. The MCI error code. If this parameter is omitted, the last error code is used. |

# <a name="GetLastError"></a>GetLastError

Returns the last MCI error code.

```
FUNCTION GetLastError () AS MCIERROR
```

# <a name="GetTrackLength"></a>GetTrackLength

Returns the length in seconds of the given track.

```
FUNCTION GetTrackLength (BYVAL nTrack AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nTrack* | The track number. |

# <a name="GetTrackLengthString"></a>GetTrackLengthString

Returns the length of the given track as a string.

```
FUNCTION GetTrackLengthString (BYVAL nTrack AS LONG) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nTrack* | The track number. |

# <a name="GetTracksCount"></a>GetTracksCount

Returns the count of tracks.

```
FUNCTION GetTracksCount () AS LONG
```

# <a name="GetTrackStartTime"></a>GetTrackStartTime

Returns the start time of the given track.

```
FUNCTION GetTrackStartTime (BYVAL nTrack AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nTrack* | The track number. |

# <a name="GetTrackStartTimeString"></a>GetTrackStartTimeString

Returns the start time of the given track as a string.

```
FUNCTION GetTrackStartTimeString (BYVAL nTrack AS LONG) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nTrack* | The track number. |

# <a name="IsMediaInserted"></a>IsMediaInserted

Checks whether CD media is inserted.

```
FUNCTION IsMediaInserted () AS BOOLEAN
```

#### Return value

TRUE or FALSE.

# <a name="IsPaused"></a>IsPaused

Checks whether is in paused mode.

```
FUNCTION IsPaused () AS BOOLEAN
```

#### Return value

TRUE or FALSE.

# <a name="IsPlaying"></a>IsPlaying

Checks whether is in play mode.

```
FUNCTION IsPlaying () AS BOOLEAN
```

#### Return value

TRUE or FALSE.

# <a name="IsReady"></a>IsReady

Checks if the device is ready.

```
FUNCTION IsReady () AS BOOLEAN
```

#### Return value

TRUE or FALSE.

# <a name="IsSeeking"></a>IsSeeking

Checks whether is in seeking mode.

```
FUNCTION IsSeeking () AS BOOLEAN
```

#### Return value

TRUE or FALSE.

# <a name="IsStopped"></a>IsStopped

Checks whether is in stopped mode.

```
FUNCTION IsStopped () AS BOOLEAN
```

#### Return value

TRUE or FALSE.

# <a name="Open"></a>Open

Initializes the device.

```
FUNCTION Open () AS DWORD
```

#### Return value

Returns zero if successful or an error otherwise.

# <a name="OpenDoor"></a>OpenDoor

Opens the CDRom door.

```
FUNCTION OpenDoor () AS MCIERROR
```

#### Return value

Returns zero if successful or an error otherwise.

# <a name="Pause"></a>Pause

Pauses playing CD Audio.

```
FUNCTION Pause () AS MCIERROR
```

#### Return value

Returns zero if successful or an error otherwise.

# <a name="Play"></a>Play

Starts playing CD Audio.

```
FUNCTION Play () AS MCIERROR
```

#### Return value

Returns zero if successful or an error otherwise.

# <a name="PlayFrom"></a>PlayFrom

Starts playing CD Audio on the given track.

```
FUNCTION PlayFrom (BYVAL nTrack AS LONG) AS MCIERROR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nTrack* | The track number. |

#### Return value

Returns zero if successful or an error otherwise.

# <a name="PlayFromTo"></a>PlayFromTo

Starts playing CD Audio from a given track to a given track.

```
FUNCTION PlayFromTo (BYVAL nStartTrack AS LONG, BYVAL nEndTrack AS LONG) AS MCIERROR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nStartTrack* | The starting track number. |
| *nEndTrack* | The ending track number. |

#### Return value

Returns zero if successful or an error otherwise.

# <a name="Stop"></a>Stop

Stops playing CD Audio.

```
FUNCTION Stop () AS MCIERROR
```

# <a name="ToEnd"></a>ToEnd

Sets the position to the end of the audio CD.

```
FUNCTION ToEnd () AS MCIERROR
```

# <a name="ToStart"></a>ToStart

Sets the position to the start of the audio CD.

```
FUNCTION ToStart () AS MCIERROR
```
