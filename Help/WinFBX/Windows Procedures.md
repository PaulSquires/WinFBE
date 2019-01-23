# Windows Procedures

Assorted Windows procedures.

**Include File**: AfxWin.inc.

# Window

| Name       | Description |
| ---------- | ----------- |
| [AfxCenterWindow](#AfxCenterWindow) | Centers a window on the screen or over another window. |
| [AfxForceSetForegroundWindow](#AfxForceSetForegroundWindow) | Brings the thread that created the specified window into the foreground and activates the window. |
| [AfxGetTopEnabledWindow](#AfxGetTopEnabledWindow) | Retrieves the handle of the enabled and visible window at the top of the z-order in an application. |
| [AfxGetTopLevelParent](#AfxGetTopLevelParent) | Retrieves the window's top-level parent window. |
| [AfxGetTopLevelWindow](#AfxGetTopLevelWindow) | Retrieves the window's top-level parent or owner window. |
| [AfxGetWindowClassName](#AfxGetWindowClassName) | Retrieves the name of the class to which the specified window belongs. |
| [AfxGetWindowClientHeight](#AfxGetWindowClientHeight) | Returns the height of the client area of window, in pixels. |
| [AfxGetWindowClientRect](#AfxGetWindowClientRect) | Retrieves the coordinates of a window's client area. |
| [AfxGetWindowClientWidth](#AfxGetWindowClientWidth) | Returns the width of the client area of a window, in pixels. |
| [AfxGetWindowHeight](#AfxGetWindowHeight) | Returns the height of a window, in pixels. |
| [AfxGetWindowLocation](#AfxGetWindowLocation) | Returns the location of the top left corner of the window, in pixels. |
| [AfxGetWindowRect](#AfxGetWindowRect) | Retrieves the dimensions of the bounding rectangle of the specified window. |
| [AfxGetWindowSize](#AfxGetWindowSize) | Gets the width and height of a window, in pixels. |
| [AfxGetWindowText](#AfxGetWindowText) | Gets the text of a window. |
| [AfxGetWindowTextLength](#AfxGetWindowTextLength) | Gets the length of the text of a window. |
| [AfxGetWindowWidth](#AfxGetWindowWidth) | Returns the width of a window, in pixels. |
| [AfxGetWorkAreaHeight](#AfxGetWorkAreaHeight) | Retrieves the height of the work area on the primary display monitor expressed in virtual screen coordinates. |
| [AfxGetWorkAreaRect](#AfxGetWorkAreaRect) | Retrieves the coordinates of the work area on the primary display monitor expressed in virtual screen coordinates |
| [AfxGetWorkAreaWidth](#AfxGetWorkAreaWidth) | Retrieves the width of the work area on the primary display monitor expressed in virtual screen coordinates. |
| [AfxRedrawNonClientArea](#AfxRedrawNonClientArea) | Redraws the non-client area of the specified window. |
| [AfxRedrawWindow](#AfxRedrawWindow) | Redraws the specified window. |
| [AfxSetWindowClientSize](#AfxSetWindowClientSize) | Adjusts the bounding rectangle of a window based on the desired size of the client area. |
| [AfxSetWindowIcon](#AfxSetWindowIcon) | Associates a new large icon with a window. |
| [AfxSetWindowLocation](#AfxSetWindowLocation) | Sets the location of the top left corner of the window, in pixels. |
| [AfxSetWindowSize](#AfxSetWindowSize) | Sets the size of the specified window, in pixels. |
| [AfxSetWindowText](#AfxSetWindowText) | Sets the text of a window. |
| [AfxShowWindowState](#AfxShowWindowState) | Sets the specified window's show state. |

# Messages

| Name       | Description |
| ---------- | ----------- |
| [AfxDoEvents](#AfxDoEvents) | Processes pending Windows messages. |
| [AfxForwardSizeMessage](#AfxForwardSizeMessage) | Sends a WM_SIZE message to the specified window. |
| [AfxPumpMessages](#AfxPumpMessages) | Processes pending Windows messages. |

# Handles

| Name       | Description |
| ---------- | ----------- |
| [AfxGetControlHandle](#AfxGetControlHandle) | Returns the handle of the control with the specified identifier. |
| [AfxGetFormHandle](#AfxGetFormHandle) | Finds the handle of the top-level window or MDI child window that is the ancestor of the specified window handle. |
| [AfxGetHwndFromPID](#AfxGetHwndFromPID) | Retrieves a window handle given it's process identifier. |
| [AfxGetPathFromWindowHandle](#AfxGetPathFromWindowHandle) | Retrieves the path of the executable file that created the specified window. |

# Window styles

| Name       | Description |
| ---------- | ----------- |
| [AfxAddWindowExStyle](#AfxAddWindowExStyle) | Adds a new extended style to the specified window. |
| [AfxAddWindowStyle](#AfxAddWindowStyle) | Adds a new style to the specified window. |
| [AfxGetWindowExStyle](#AfxGetWindowExStyle) | Retrieves the extended window styles. |
| [AfxGetWindowStyle](#AfxGetWindowStyle) | Retrieves the window styles. |
| [AfxRemoveWindowExStyle](#AfxRemoveWindowExStyle) | Removes an extended style from the specified window. |
| [AfxRemoveWindowStyle](#AfxRemoveWindowStyle) | Removes a style from the specified window. |
| [AfxSetWindowExStyle](#AfxSetWindowExStyle) | Sets the extended style(s) of the specified window. |
| [AfxSetWindowStyle](#AfxSetWindowStyle) | Sets the style(s) of the specified window. |

# Display

| Name       | Description |
| ---------- | ----------- |
| [AfxForceVisibleDisplay](#AfxForceVisibleDisplay) | Force visibility of an off-screen window. |
| [AfxGetDisplayBitsPerPixel](#AfxGetDisplayBitsPerPixel) | Returns the color resolution, in bits per pixel, of the display device. |
| [AfxGetDisplayFrequency](#AfxGetDisplayFrequency) | Returns the frequency, in hertz (cycles per second), of the display device in a particular mode. |
| [AfxGetDisplayPixelsHeight](#AfxGetDisplayPixelsHeight) | Returns the height, in pixels, of the current display device on the computer on which the calling thread is running. |
| [AfxGetDisplayPixelsWidth](#AfxGetDisplayPixelsWidth) | Returns the width, in pixels, of the current display device on the computer on which the calling thread is running. |

# Dialogs

| Name       | Description |
| ---------- | ----------- |
| [AfxChooseColorDialog](#AfxChooseColorDialog) | Displays the Windows choose color dialog. |
| [AfxControlRunDLL](#AfxControlRunDLL) | Launches control panel applications. |
| [AfxShowSysInfo](#AfxShowSysInfo) | Displays the Windows Information System dialog. |

# High DPI

| Name       | Description |
| ---------- | ----------- |
| [AfxGetDpi](#AfxLogPixelsX) | Retrieves the number of pixels per logical inch. |
| [AfxGetDpiX](#AfxLogPixelsX) | Retrieves the number of pixels per logical inch along the screen width. |
| [AfxGetDpiY](#AfxLogPixelsY) | Retrieves the number of pixels per logical inch along the screen height. |
| [AfxGetMonitorHorizontalScaling](#AfxGetMonitorHorizontalScaling) | Returns the horizontal scaling of the monitor that the window is currently displayed on. |
| [AfxGetMonitorVerticalScaling](#AfxGetMonitorVerticalScaling) | Returns the vertical scaling of the monitor that the window is currently displayed on. |
| [AfxGetMonitorLogicalHeight](#AfxGetMonitorLogicalHeight) | Returns the logical height of the monitor that the window is currently displayed on. |
| [AfxGetMonitorLogicalWidth](#AfxGetMonitorLogicalWidth) | Returns the logical width of the monitor that the window is currently displayed on. |
| [AfxIsDPIResolutionAtLeast](#AfxIsDPIResolutionAtLeast) | Determines if screen resolution meets minimum requirements in relative pixels. |
| [AfxIsProcessDPIAware](#AfxIsProcessDPIAware) | Determines whether the current process is dots per inch (dpi) aware. |
| [AfxIsResolutionAtLeast](#AfxIsResolutionAtLeast) | Determines if screen resolution meets minimum requirements. |
| [AfxLoadIconMetric](#AfxLoadIconMetric) | Loads a specified icon resource with a client-specified system metric. |
| [AfxLogPixelsX](#AfxLogPixelsX) | Retrieves the number of pixels per logical inch along the screen width. |
| [AfxLogPixelsY](#AfxLogPixelsY) | Retrieves the number of pixels per logical inch along the screen height. |
| [AfxScaleRatioX](#AfxScaleRatioX) | Retrieves the desktop horizontal scaling ratio. |
| [AfxScaleRatioY](#AfxScaleRatioY) | Retrieves the desktop vertical scaling ratio. |
| [AfxScaleX](#AfxScaleX) | Scales an horizontal coordinate according the DPI (dots per pixel) being used by the operating system. |
| [AfxScaleY](#AfxScaleY) | Scales an vertical coordinate according the DPI (dots per pixel) being used by the operating system. |
| [AfxSetProcessDPIAware](#AfxSetProcessDPIAware) | Sets the current process as dots per inch (dpi) aware. |
| [AfxUnscaleX](#AfxUnscaleX) | Unscales an horizontal coordinate according the DPI (dots per pixel) being used by the operating system. |
| [AfxUnscaleY](#AfxUnscaleY) | Unscales a vertical coordinate according the DPI (dots per pixel) being used by the operating system. |
| [AfxUseDpiScaling](#AfxUseDpiScaling) | Returns TRUE if the Windows Platform is NT; FALSE, otherwise. |

# Fonts

| Name       | Description |
| ---------- | ----------- |
| [AfxCreateFont](#AfxCreateFont) | Creates a logical font. |
| [AfxGetFontHeight](#AfxGetFontHeight) | Returns the logical height of a font given its point size. |
| [AfxGetFontPointSize](#AfxGetFontPointSize) | Returns the point size of a font given its logical height. |
| [AfxGetWindowFont](#AfxGetWindowFont) | Retrieves the font with which the control is currently drawing its text. |
| [AfxGetWindowsFontInfo](#AfxGetWindowsFontInfo) | Retrieves information about the fonts used by Windows. |
| [AfxGetWindowsFontPointSize](#AfxGetWindowsFontPointSize) | Retrieves the point size of the fonts used by Windows. |
| [AfxModifyFontFaceName](#AfxModifyFontFaceName) | Modifies the face name of the font of a window or control. |
| [AfxModifyFontHeight](#AfxModifyFontHeight) | Modifies the height of the font used by a window of control. |
| [AfxModifyFontSettings](#AfxModifyFontSettings) | Modifies settings of the font used by a window of control. |
| [AfxSetWindowFont](#AfxSetWindowFont) | Sets the font that a control is to use when drawing text. |

# Clipboard

| Name       | Description |
| ---------- | ----------- |
| [AfxClearClipboard](#AfxClearClipboard) | Clears the contents of the clipboard. |
| [AfxGetClipboardData](#AfxGetClipboardData) | Retrieves data from the clipboard in the specified format. |
| [AfxGetClipboardText](#AfxGetClipboardText) | Returns a text string from the clipboard. |
| [AfxSetClipboardData](#AfxSetClipboardData) | Places data on the clipboard in a specified clipboard format. |
| [AfxSetClipboardText](#AfxSetClipboardText) | Places a text string into the clipboard. |

# Bitmap

| Name       | Description |
| ---------- | ----------- |
| [AfxCaptureDisplay](#AfxCaptureDisplay) | Captures the display and returns an handle to a bitmap. |
| [AfxGetBitmapHeight](#AfxGetBitmapHeight) | Retrieves the height of the specified bitmap. |
| [AfxGetBitmapWidth](#AfxGetBitmapWidth) | Retrieves the width of the specified bitmap. |

# Device Independent Bitmap (DIB)

| Name       | Description |
| ---------- | ----------- |
| [AfxCreateDIBSection](#AfxCreateDIBSection) | Creates a DIB section. |
| [AfxDibLoadImage](#AfxDibLoadImage) | Loads a DIB in memory and returns a pointer to it. |
| [AfxDibSaveImage](#AfxDibSaveImage) | Saves a DIB to a file. |

# Metric conversions

| Name       | Description |
| ---------- | ----------- |
| [AfxHiMetricToPixelsX](#AfxHiMetricToPixelsX) | Converts from HiMetric to Pixels (horizontal resolution). |
| [AfxHiMetricToPixelsY](#AfxHiMetricToPixelsY) | Converts from HiMetric to Pixels (vertical resolution). |
| [AfxPixelsToHiMetricX](#AfxPixelsToHiMetricX) | Converts from Pixels to HiMetric (horizontal resolution). |
| [AfxPixelsToHiMetricY](#AfxPixelsToHiMetricY) | Converts from Pixels to HiMetric (vertical resolution). |
| [AfxPixelsToPointsX](#AfxPixelsToPointsX) | Converts pixels to points size (1/72 of an inch) (horizontal resolution). |
| [AfxPixelsToPointsY](#AfxPixelsToPointsY) | Converts pixels to points size (1/72 of an inch) (vertical resolution). |
| [AfxPixelsToTwipsX](#AfxPixelsToTwipsX) | Converts pixels to twips (horizontal resolution). |
| [AfxPixelsToTwipsY](#AfxPixelsToTwipsY) | Converts pixels to twips (vertical resolution). |
| [AfxPointSizeToDip](#AfxPointSizeToDip) | Converts point size to DIP (device independent pixel). |
| [AfxPointsToPixelsX](#AfxPointsToPixelsX) | Converts a point size (1/72 of an inch) to pixels (horizontal resolution). |
| [AfxPointsToPixelsY](#AfxPointsToPixelsY) | Converts a point size (1/72 of an inch) to pixels (vertical resolution). |
| [AfxTwipsPerPixelX](#AfxTwipsPerPixelX) | Returns the width of a pixel in twips (horizontal resolution). |
| [AfxTwipsPerPixelY](#AfxTwipsPerPixelY) | Returns the width of a pixel in twips (vertical resolution). |
| [AfxTwipsToPixelsX](#AfxTwipsToPixelsX) | Converts twips to pixels (horizontal resolution). |
| [AfxTwipsToPixelsY](#AfxTwipsToPixelsY) | Converts twips to pixels (vertical resolution). |

# Mail and Internet

| Name       | Description |
| ---------- | ----------- |
| [AfxGetBrowserHandle](#AfxGetBrowserHandle) | Retrieves the handle of the top level window of the web browser. |
| [AfxGetDefaultBrowserName](#AfxGetDefaultBrowserName) | Retrieves the name of the default browser. |
| [AfxGetDefaultBrowserPath](#AfxGetDefaultBrowserPath) | Retrieves the path of the default browser. |
| [AfxGetDefaultMailClientName](#AfxGetDefaultMailClientName) | Retrieves the name of the default client mail application. |
| [AfxGetDefaultMailClientPath](#AfxGetDefaultMailClientPath) | Retrieves the path of the default mail client application. |
| [AfxGetInternetExplorerVersion](#AfxGetInternetExplorerVersion) | Returns the Internet Explorer version installed. |

# Miscellaneous procedures

| Name       | Description |
| ---------- | ----------- |
| [AfxCommand](#AfxCommand) | Returns command line parameters used to call the program. |
| [AfxEnviron](#AfxEnviron) | Retrieves the contents of the specified variable from the environment block of the calling process. |
| [AfxExtractResource](#AfxExtractResource) | Extracts resource data and returns it as a string. |
| [AfxExtractResourceToFile](#AfxExtractResourceToFile) | Extracts resource data and saves it to a file. |
| [AfxGetComputerName](#AfxGetComputerName) | Retrieves the NetBIOS name of the local computer. |
| [AfxGetMACAddress](#AfxGetMACAddress) | Retrieves the MAC address of a machine's Ethernet card. |
| [AfxGetUserName](#AfxGetUserName) | Retrieves the name of the user associated with the current thread. |
| [AfxGetWinErrMsg](#AfxGetWinErrMsg) | Retrieves the localized description of the specified Windows error code. |
| [AfxMsg](#AfxMsg) | Displays an application modal message box. |
| [AfxSetDlgMsgResult](#AfxSetDlgMsgResult) | Sets the return value of a message processed in the dialog box procedure. |
| [AfxSetEnviron](#AfxSetEnviron) | Sets the contents of the specified environment variable for the current process. |

# Versioning

| Name       | Description |
| ---------- | ----------- |
| [AfxComCtlVersion](#AfxComCtlVersion) | Returns the version of CommCtl32.dll. |
| [AfxIsPlatformNT](#AfxIsPlatformNT) | Returns TRUE if the Windows Platform is NT; FALSE, otherwise. |
| [AfxWindowsBitness](#AfxWindowsBitness) | Returns the bitness of the operating system (32 or 64 bit). |
| [AfxWindowsBuild](#AfxWindowsBuild) | Returns the Windows build number. |
| [AfxWindowsPlatform](#AfxWindowsPlatform) | Returns the Windows platform. |
| [AfxWindowsVersion](#AfxWindowsVersion) | Returns the Windows version. |

# <a name="AfxWindowsVersion"></a>AfxWindowsVersion

Returns the Windows version.

```
FUNCTION AfxWindowsVersion () AS LONG
```

#### Return value

Platform 1:
```
  400 Windows 95
  410 Windows 98
  490 Windows ME
```

Platform 2:

```
  400 Windows NT
  500 Windows 2000
  501 Windows XP
  502 Windows Server 2003
  600 Windows Vista and Windows Server 2008
  601 Windows 7
  602 Windows 8
  603 Windows 8.1
 1000 Windows 10
```

**Note**: As Windows 95 and Windows NT return the same version number, we also need to call GetWindowsPlatform to differentiate them.

# <a name="AfxWindowsBuild"></a>AfxWindowsBuild

Returns the Windows build number.

```
FUNCTION AfxWindowsBuild () AS LONG
```

# <a name="AfxWindowsPlatform"></a>AfxWindowsPlatform

Returns the Windows platform.

```
FUNCTION AfxWindowsPlatform () AS LONG
```

#### Return value

| Value      | Description |
| ---------- | ----------- |
| 1 | Windows 95/98/ME |
| 2 | Windows NT/2000/XP/Server/Vista/Windows 7 |

# <a name="AfxWindowsBitness"></a>AfxWindowsBitness

Returns the Windows bitness (32 or 64 bit).

```
FUNCTION AfxWindowsBitness () AS LONG
```

# <a name="AfxIsPlatformNT"></a>AfxIsPlatformNT

Returns TRUE if the Windows Platform is NT; FALSE, otherwise.

```
FUNCTION AfxIsPlatformNT () AS BOOLEAN
```

# <a name="AfxComCtlVersion"></a>AfxComCtlVersion

Returns the version of CommCtl32.dll

```
FUNCTION AfxComCtlVersion () AS LONG
```

#### Return value

The version of CommCtl32.dll multiplied by 100, e.g. 582 for version 5.82.

# <a name="AfxMsg"></a>AfxMsg

Displays an application modal message box. Helper proceudre to display feedback and errors.

```
FUNCTION AfxMsg (BYVAL pwszText AS WSTRING PTR, BYREF wszCaption AS WSTRING = "Message", _
   BYVAL uType AS DWORD = 0) AS LONG
FUNCTION AfxMsg (BYREF cws AS CWSTR, BYREF wszCaption AS WSTRING = "Message", _
   BYVAL uType AS DWORD = 0) AS LONG
FUNCTION AfxMsg (BYREF cbs AS CBSTR, BYREF wszCaption AS WSTRING = "Message", _
   BYVAL uType AS DWORD = 0) AS LONG
FUNCTION AfxMsg (BYREF cv AS CVAR, BYREF wszCaption AS WSTRING = "Message", _
   BYVAL uType AS DWORD = 0) AS LONG
FUNCTION AfxMsg (BYVAL nValue AS DOUBLE, BYREF wszCaption AS WSTRING = "Message", _
   BYVAL uType AS DWORD = 0) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszText* | Pointer to a WSTRING. |
| *cws* | A CWSTR variable. |
| *cws* | A CBSTR variable. |
| *cv* | A CVAR variable. |
| *nValue* | A numeric variable. |
| *wszCaption* | Optional. The message box caption. Default title is "Message". |
| *uType* | Optional. For a list of available types, see the Microsoft documentation for the MessageBoxW function. The MB_APPLMODAL type is always added. |

#### Return value

The version of CommCtl32.dll multiplied by 100, e.g. 582 for version 5.82.

# <a name="AfxGetWinErrMsg"></a>AfxGetWinErrMsg

Retrieves the localized description of the specified Windows error code.

```
FUNCTION AfxGetWinErrMsg (BYVAL dwError AS DWORD) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwError* | The Windows error code. |

#### Return value

The localized description of the error code.

# <a name="AfxCommand"></a>AfxCommand

Returns command line parameters used to call the program.

```
FUNCTION AfxCommand (BYVAL nIndex AS LONG = -1) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nIndex* | Zero-based index for a particular command-line argument. |

#### Return value

Returns the command-line arguments(s).

#### Remarks

**AfxCommand** returns command-line arguments passed to the program upon execution.

If index is less than zero (< 0), a space-separated list of all command-line arguments is returned, otherwise, a single argument is returned. A value of zero (0) returns the name of the executable; and values of one (1) and greater return each command-line argument.

If index is greater than the number of arguments passed to the program, a null string ("") is returned.

# <a name="AfxEnviron"></a>AfxEnviron

Retrieves the contents of the specified variable from the environment block of the calling process.

```
FUNCTION AfxEnviron (BYVAL pwszName AS LPCWSTR) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszName* | The name of the environment variable. |

**Include file**: AfxStr.inc

#### Usage example

```
DIM cws AS CWSTR = AfxEnviron("path")
```

# <a name="AfxSetEnviron"></a>AfxSetEnviron

Sets the contents of the specified environment variable for the current process.

```
FUNCTION AfxSetEnviron (BYVAL pwszName AS LPCWSTR, BYVAL pwszValue AS LPCWSTR) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszName* | The name of the environment variable. The operating system creates the environment variable if it does not exist and *pwszValue* is not NULL. |
| *pwszValue* | The name of the environment variable. The maximum size of a user-defined environment variable is 32,767 characters. |

**Include file**: AfxStr.inc

#### Return value

If the function succeeds, the return value is TRUE.<br>
If the function fails, the return value is FALSE.<br>
To get extended error information, call **GetLastError**.

#### Usage example

```
AfxSetEnviron("path", "c:")
```

# AfxSetEnviron (Overload)

Sets the contents of the specified environment variable for the current process.<br>
Unicode replacement for Free Basic's **SetEnviron** function.

```
FUNCTION AfxSetEnviron (BYVAL pwszName AS LPCWSTR, BYVAL pwszValue AS LPCWSTR) AS BOOLEAN
FUNCTION AfxSetEnviron (BYREF varexp AS WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszName* | The name of the environment variable. The operating system creates the environment variable if it does not exist and *pwszValue* is not NULL. |
| *varexp* | Name and setting of an environment variable in the following (or equivalent) form: varname=varstring (*varname* being the name of the environment variable, and *varstring* being its text value to set).

**Include file**: AfxStr.inc

#### Return value

Returns 0 on success, or -1 on failure.

# <a name="AfxSetDlgMsgResult"></a>AfxSetDlgMsgResult

Sets the return value of a message processed in the dialog box procedure.

```
FUNCTION AfxSetDlgMsgResult (BYVAL hDlg AS HWND, BYVAL msg AS UINT, BYVAL result AS LONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hDlg* | A handle to a dialog box. |
| *msg* | The message type. |
| *result* | The return value. |

#### Return value

If the function succeeds, the return value is TRUE. If the function fails, the return value is FALSE.

# <a name="AfxExtractResource"></a>AfxExtractResource

Extracts resource data and returns it as a string.

```
FUNCTION AfxExtractResource (BYVAL hInstance AS HINSTANCE, _
   BYREF wszResourceName AS WSTRING, BYVAL pResourceType AS LPWSTR = RT_RCDATA) AS STRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *hInstance* | A handle to the module whose portable executable file or an accompanying MUI file contains the resource. If this parameter is NULL, the function searches the module used to create the current process. |
| *wszResourceName* | Name of the resource. If the resource is an image that uses an integral identifier, *wszResourceName* should begin with a number symbol (#) followed by the identifier in an ASCII format, e.g., "#998". Otherwise, use the text identifier name for the image. Only images embedded as raw data (type RCDATA) are valid. These must be in format .png, .jpg, .gif, .tiff. |
| *pResourceType* | Type of the resource, e.g. RT_RCDATA. For a list of predefined resource types see: [Resource Types](https://docs.microsoft.com/en-us/windows/desktop/menurc/resource-types) |

#### Return value

A string containing the resource data.

#### Example

```
DIM strResData AS STRING = AfxExtractResource(NULL, "IDI_ARROW_RIGHT")
where IDI_ARROW_RIGHT is the identifier in the resource file for
IDI_ARROW_RIGHT RCDATA ".\Resources\arrow_right_64.png"
--or--
DIM strResData AS STRING = AfxExtractResource(NULL, "#111")
' where "#111" is the identifier in the resource file for
' 111 RCDATA ".\Resources\VEGA_PAZ_01.jpg"
-----
' // Write the resource data to a file
DIM hFile AS HANDLE = CreateFileW("PazVega.jpg", GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL)
IF hFile THEN
   DIM dwBytesWritten AS DWORD
   DIM bSuccess AS BOOLEAN = WriteFile(hFile, STRPTR(strResData), LEN(strResData), @dwBytesWritten, NULL)
   CloseHandle(hFile)
   print bSuccess
END IF
```

# <a name="AfxExtractResourceToFile"></a>AfxExtractResourceToFile

Extracts resource data and saves it to a file.

```
FUNCTION AfxExtractResourceToFile (BYVAL hInstance AS HINSTANCE, BYREF wszResourceName AS WSTRING, _
   BYREF wszFileName AS WSTRING, BYVAL pResourceType AS LPWSTR = RT_RCDATA) AS STRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *hInstance* | A handle to the module whose portable executable file or an accompanying MUI file contains the resource. If this parameter is NULL, the function searches the module used to create the current process. |
| *wszResourceName* | Name of the resource. If the resource is an image that uses an integral identifier, *wszResourceName* should begin with a number symbol (#) followed by the identifier in an ASCII format, e.g., "#998". Otherwise, use the text identifier name for the image. Only images embedded as raw data (type RCDATA) are valid. These must be in format .png, .jpg, .gif, .tiff. |
| *wszFileName* | Path of the file where to save the extracted resource. |
| *pResourceType* | Type of the resource, e.g. RT_RCDATA. For a list of predefined resource types see: [Resource Types](https://docs.microsoft.com/en-us/windows/desktop/menurc/resource-types) |

#### Return value

TRUE on success of FALSE on failure.

#### Example

```
AfxExtractResourceToFile(NULL, "IDI_ARROW_RIGHT", "IDI_ARROW_RIGHT.png")
where IDI_ARROW_RIGHT is the identifier in the resource file for
IDI_ARROW_RIGHT RCDATA ".\Resources\arrow_right_64.png"
```
#### Example

```
AfxExtractResourceToFile(NULL, "#111", "VEGA_PAZ_01.jpg")
where "#111" is the identifier in the resource file for
111 RCDATA ".\Resources\VEGA_PAZ_01.jpg"
```

# <a name="AfxGetComputerName"></a>AfxGetComputerName

Retrieves the NetBIOS name of the local computer. This name is established at system startup, when the system reads it from the registry.

```
FUNCTION AfxGetComputerName () AS CWSTR
```

#### Return value

The NetBIOS name of the local computer.

#### Remarks

The behavior of this function can be affected if the local computer is a node in a cluster. For more information, see **ResUtilGetEnvironmentWithNetName** and **UseNetworkName**.

# <a name="AfxGetUserName"></a>AfxGetUserName

Retrieves the name of the user associated with the current thread.

```
FUNCTION AfxGetUserName () AS CWSTR
```

#### Return value

The name of the current user associated with the current thread.

#### Remarks

If the current thread is impersonating another client, the **AfxGetUserName** function returns the user name of the client that the thread is impersonating.

# <a name="AfxGetMACAddress"></a>AfxGetMACAddress

Retrieves the MAC address of a machine's Ethernet card.

```
FUNCTION AfxGetMACAddress () AS STRING
```

#### Return value

The MAC address in the following format: MM-MM-MM-SS-SS-SS. The leftmost 6 digits, called a "prefix", is associated with the adapter manufacturer. The rightmost digits of a MAC address represent an identification number for the specific device.

#### Remarks

This function only supports one NIC card on your PC.

# <a name="AfxDoEvents"></a>AfxDoEvents

Processes pending Windows messages. Call this procedure if you are performing a tight FOR/NEXT or DO/LOOP and need to allow your application to be responsive to user input.

```
SUB AfxDoEvents (BYVAL hwnd AS HWND = NULL)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Optional. Handle of the window or dialog. If NULL, the window handle to the active window attached to the calling thread's message queue is used. |

# <a name="AfxPumpMessages"></a>AfxPumpMessages

Processes pending Windows messages. Call this procedure if you are performing a tight FOR/NEXT or DO/LOOP and need to allow your application to be responsive to user input.

```
SUB AfxPumpMessages
```

# <a name="AfxForwardSizeMessage"></a>AfxForwardSizeMessage

Sends a WM_SIZE message to the specified window.

```
FUNCTION AfxForwardSizeMessage (BYVAL hwnd AS HWND, BYVAL nResizeType AS DWORD, _
   BYVAL nWidth AS LONG, BYVAL nHeight AS LONG) AS LRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | A handle to a window. |
| *nResizeType* | Type of resizing requested. |
| *nWidth* | The new width of the client area. |
| *nHeight* | The new height of the client ara. |

| Resizing type  | Description |
| -------------- | ----------- |
| SIZE_MAXHIDE | Message is sent to all pop-up windows when some other window is maximized. |
| SIZE_MAXIMIZED | Maximize the window. |
| SIZE_MAXSHOW | Message is sent to all pop-up windows when some other window has been restored to its former size. |
| SIZE_MINIMIZED | Minimize the window. |
| SIZE_RESTORED | The window has been resized, but neither the SIZE_MINIMIZED nor SIZE_MAXIMIZED value applies. |

#### Remark

If an application processes this message, it should return zero.

# <a name="AfxChooseColorDialog"></a>AfxChooseColorDialog

Displays the Windows choose color dialog.

```
FUNCTION AfxChooseColorDialog (BYVAL hwnd AS HWND, BYVAL rgbDefaultColor AS COLORREF = 0, _
   BYVAL lpCustColors AS COLORREF PTR = NULL) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | A handle to the parent window or NULL. |
| *rgbDefaultColor* | The color initially selected when the dialog box is created. If the specified color value is not among the available colors, the system selects the nearest solid color available. If *rgbDefaultColor* is zero, the initially selected color is black. |
| *lpCustColors* | Out. A pointer to an array of 16 values that contain red, green, blue (RGB) values for the custom color boxes in the dialog box. If the user modifies these colors, the system updates the array with the new RGB values. To preserve new custom colors between calls to the **AfxChooseColorDialog** function, you should allocate static memory for the array. To create a COLORREF color value, use the BGR macro. |

#### Return value

The selected color, or -1 if the user has canceled the dialog.

# <a name="AfxControlRunDLL"></a>AfxControlRunDLL

Control_RunDLL is an undocumented procedure in the Shell32.dll which can be used to launch control panel applications. You’ve to pass the name of the control panel file (.cpl) and the tool represented by it will be launched. For launching some control panel applications, you’ve to provide a valid windows handle (hwnd parameter) and program instance (*hInst*) parameter).

```
FUNCTION AfxControlRunDLL (BYVAL hwnd AS HWND, BYVAL hInst AS HINSTANCE, _
   BYVAL cmd AS WSTRING PTR, BYVAL nCmdShow AS LONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to a window. This parameter can be NULL. |
| *hInst* | Instance handle. This parameter can be NULL. |
| *cmd* | The command and parameters. |
| *nCmdShow* | Controls how the window is to be shown, e.g. SW_SHOWNORMAL. |

| nCmdShow value  | Description |
| --------------- | ----------- |
| SW_FORCEMINIMIZE | Minimizes a window, even if the thread that owns the window is not responding. This flag should only be used when minimizing windows from a different thread. |
| SW_HIDE | Hides the window and activates another window. |
| SW_MAXIMIZE | Maximizes the specified window. |
| SW_MINIMIZE | Minimizes the specified window and activates the next top-level window in the Z order. |
| SW_RESTORE | Activates and displays the window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when restoring a minimized window. |
| SW_SHOW | Activates the window and displays it in its current size and position. |
| SW_SHOWDEFAULT | Sets the show state based on the SW_ value specified in the STARTUPINFO structure passed to the **CreateProcess** function by the program that started the application. |
| SW_SHOWMAXIMIZED | Activates the window and displays it as a maximized window. |
| SW_SHOWMINIMIZED | Activates the window and displays it as a minimized window. |
| SW_SHOWMINNOACTIVE | Displays the window as a minimized window. This value is similar to SW_SHOWMINIMIZED, except the window is not activated. |
| SW_SHOWNA | Displays the window in its current size and position. This value is similar to SW_SHOW, except that the window is not activated. |
| SW_SHOWNOACTIVATE | Displays a window in its most recent size and position. This value is similar to SW_SHOWNORMAL, except that the window is not activated. |
| SW_SHOWNORMAL | Activates and displays a window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when displaying the window for the first time. |

#### Usage examples

```
AfxControlRunDLL(0, 0, "", SW_SHOWNORMAL)   ' Opens the control panel
AfxControlRunDLL(0, 0, "appwiz.cpl", SW_SHOWNORMAL)   ' Opens the applications wizard
```

# <a name="AfxShowSysInfo"></a>AfxShowSysInfo

Displays the Windows Information System dialog.

```
FUNCTION AfxShowSysInfo (BYVAL hwnd AS HWND) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | A handle to the parent window or NULL. |

#### Return value

If the function succeeds, the return value is TRUE. If the function fails, the return value is FALSE.

# <a name="AfxGetBrowserHandle"></a>AfxGetBrowserHandle

Retrieves the handle of the top level window of the web browser.

```
FUNCTION AfxGetBrowserHandle (BYVAL pwszClassName AS WSTRING PTR) AS HWND
FUNCTION AfxGetInternetExplorerHandle () AS HWND
FUNCTION AfxGetFireFoxHandle () AS HWND
FUNCTION AfxGetGoogleChromeHandle () AS HWND
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszClassName* | Class name of the browser. |

#### Return value

The handle of the top level window of the browser. If there is not any instance of the browser running, it will return NULL.

#### Examples

```
DIM hwndBrowser AS HWND = AfxGetBrowserHandle("IEFrame")              ' // Internet Explorer
DIM hwndBrowser AS HWND = AfxGetBrowserHandle("MozillaWindowClass")   ' // Firefox
DIM hwndBrowser AS HWND = AfxGetBrowserHandle("Chrome_WidgetWin_1")   ' // Chrome
```

#### Remarks

**AfxGetInternetExplorerHandle**, **AfxGetFireFoxHandle** and **AfxGetGoogleChromeHandle** are wrappers functions that call AfxGetBrowserHandle passing the appropriate class name ("IEFrame", "MozillaWindowClass" and "Chrome_WidgetWin_1").

# <a name="AfxGetDefaultBrowserName"></a>AfxGetDefaultBrowserName

Retrieves the name of the default browser.

```
FUNCTION AfxGetDefaultBrowserName () AS CWSTR
```

#### Return value

The retrieved name or an empty string.

# <a name="AfxGetDefaultBrowserPath"></a>AfxGetDefaultBrowserPath

Retrieves the path of the default browser.

```
FUNCTION AfxGetDefaultBrowserPath () AS CWSTR
```

#### Return value

The retrieved path or an empty string.

# <a name="AfxGetDefaultMailClientName"></a>AfxGetDefaultMailClientName

Retrieves the name of the default client mail application.

```
FUNCTION AfxGetDefaultMailClientName () AS CWSTR
```

#### Return value

The retrieved name or an empty string.

# <a name="AfxGetDefaultMailClientPath"></a>AfxGetDefaultMailClientPath

Retrieves the path of the default client mail application.

```
FUNCTION AfxGetDefaultMailClientPath () AS CWSTR
```

#### Return value

The retrieved path or an empty string.

# <a name="AfxGetInternetExplorerVersion"></a>AfxGetInternetExplorerVersion

Returns the Internet Explorer version installed.

```
FUNCTION AfxGetInternetExplorerVersion () AS SINGLE
```

#### Return value

The Internet Explorer version (major.minor).

# <a name="AfxHiMetricToPixelsX"></a>AfxHiMetricToPixelsX

Converts from HiMetric to Pixels (horizontal resolution). Himetric is a scaling unit similar to twips used in computing. It is one thousandth of a centimeter and is independent of the screen resolution. HiMetric per inch = 2540; 1 inch = 2.54 mm.

```
SUB AfxHiMetricToPixelsX (BYVAL hm AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *hm* | The size in HiMetric units. |

#### Return value

The size in pixels.

# <a name="AfxHiMetricToPixelsY"></a>AfxHiMetricToPixelsY

Converts from HiMetric to Pixels (vertical resolution). Himetric is a scaling unit similar to twips used in computing. It is one thousandth of a centimeter and is independent of the screen resolution. HiMetric per inch = 2540; 1 inch = 2.54 mm.

```
SUB AfxHiMetricToPixelsY (BYVAL hm AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *hm* | The size in HiMetric units. |

#### Return value

The size in pixels.


# <a name="AfxPixelsToHiMetricX"></a>AfxPixelsToHiMetricX

Converts from Pixels to HiMetric (horizontal resolution). Himetric is a scaling unit similar to twips used in computing. It is one thousandth of a centimeter and is independent of the screen resolution. HiMetric per inch = 2540; 1 inch = 2.54 mm.

```
SUB AfxPixelsToHiMetricX (BYVAL cx AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *cx* | The size in pixels. |

#### Return value

The size in HiMetric units.

# <a name="AfxPixelsToHiMetricY"></a>AfxPixelsToHiMetricY

Converts from Pixels to HiMetric (vertical resolution). Himetric is a scaling unit similar to twips used in computing. It is one thousandth of a centimeter and is independent of the screen resolution. HiMetric per inch = 2540; 1 inch = 2.54 mm.

```
SUB AfxPixelsToHiMetricY (BYVAL cx AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *cy* | The size in pixels. |

#### Return value

The size in HiMetric units.

# <a name="AfxPixelsToPointsX"></a>AfxPixelsToPointsX

Converts pixels to points size (1/72 of an inch). Horizontal resolution.

```
SUB AfxPixelsToPointsX (BYVAL pix AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pix* | The number of pixels. |

#### Return value

The number of points.

# <a name="AfxPixelsToPointsY"></a>AfxPixelsToPointsY

Converts pixels to points size (1/72 of an inch). Vertical resolution.

```
SUB AfxPixelsToPointsY (BYVAL pix AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pix* | The number of pixels. |

#### Return value

The number of points.

# <a name="AfxPixelsToTwipsX"></a>AfxPixelsToTwipsX

Converts pixels to twips. Horizontal resolution.

```
FUNCTION AfxPixelsToTwipsX (BYVAL nPixels AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nPixels* | The number of pixels. |

#### Return value

The number of twips.

# <a name="AfxPixelsToTwipsY"></a>AfxPixelsToTwipsY

Converts pixels to twips. Vertical resolution.

```
FUNCTION AfxPixelsToTwipsY (BYVAL nPixels AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nPixels* | The number of pixels. |

#### Return value

The number of twips.

# <a name="AfxPointSizeToDip"></a>AfxPointSizeToDip

Converts point size to DIP (device independent pixel). DIP is defined as 1/96 of an inch and a point is 1/72 of an inch.

```
FUNCTION AfxPointSizeToDip (BYVAL ptsize AS SINGLE) AS SINGLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *ptsize* | The point size to convert. |

#### Return value

The number of DIP pixels.

# <a name="AfxPointsToPixelsX"></a>AfxPointsToPixelsX

Converts a point size (1/72 of an inch) to pixels. Horizontal resolution.

```
FUNCTION AfxPointsToPixelsX (BYVAL pts AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pts* | The number of points. |

#### Return value

The number of pixels.

# <a name="AfxPointsToPixelsY"></a>AfxPointsToPixelsY

Converts a point size (1/72 of an inch) to pixels. Vertical resolution.

```
FUNCTION AfxPointsToPixelsY (BYVAL pts AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pts* | The number of points. |

#### Return value

The number of pixels.

# <a name="AfxTwipsPerPixelX"></a>AfxTwipsPerPixelX

Returns the width of a pixel in twips (horizontal resolution). Pixel dimensions can vary between systems and may not always be square, so separate functions for pixel width and height are required.

```
FUNCTION AfxTwipsPerPixelX () AS LONG
```

#### Return value

The number of twips per pixel.

# <a name="AfxTwipsPerPixelY"></a>AfxTwipsPerPixelY

Returns the width of a pixel in twips (vertical resolution). Pixel dimensions can vary between systems and may not always be square, so separate functions for pixel width and height are required.

```
FUNCTION AfxTwipsPerPixelY () AS LONG
```

#### Return value

The number of twips per pixel.

# <a name="AfxTwipsToPixelsX"></a>AfxTwipsToPixelsX

Converts twips to pixels. Horizontal resolution.

```
FUNCTION AfxTwipsToPixelsX (BYVAL nTwips AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nTwips* | The number of twips. |

#### Return value

The number of pixels.

# <a name="AfxTwipsToPixelsY"></a>AfxTwipsToPixelsY

Converts twips to pixels. Vertical resolution.

```
FUNCTION AfxTwipsToPixelsY (BYVAL nTwips AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nTwips* | The number of twips. |

#### Return value

The number of pixels.

# <a name="AfxCaptureDisplay"></a>AfxCaptureDisplay

Captures the display and returns an handle to a bitmap.

```
FUNCTION AfxCaptureDisplay () AS HBITMAP
```

#### Return value

The handle of a bitmap.

#### Usage example

```
DIM hBitmap AS HBITMAP = AfxCaptureDisplay
```

# <a name="AfxGetBitmapHeight"></a>AfxGetBitmapHeight

Retrieves the height of the specified bitmap.

```
FUNCTION AfxGetBitmapHeight (BYVAL hBitmap AS HBITMAP) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *hBitmap* | Handle to the bitmap. |

#### Return value

The height of the bitmap on success or 0 on failure.

# <a name="AfxGetBitmapWidth"></a>AfxGetBitmapWidth

Retrieves the width of the specified bitmap.

```
FUNCTION AfxGetBitmapWidth (BYVAL hBitmap AS HBITMAP) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *hBitmap* | Handle to the bitmap. |

#### Return value

The width of the bitmap on success or 0 on failure.

# <a name="AfxCreateDIBSection"></a>AfxCreateDIBSection

Creates a DIB section.

```
FUNCTION AfxCreateDIBSection (BYVAL hdc AS HDC, BYVAL nWidth AS DWORD, BYVAL nHeight AS DWORD, _
   BYVAL bpp AS LONG = 0, BYVAL ppvBits AS ANY PTR PTR = NULL) AS HBITMAP
```

| Parameter  | Description |
| ---------- | ----------- |
| *hdc* | A handle to the device context. |
| *nWidth* | The width of the bitmap, in pixels. |
| *nHeight* | The height of the bitmap, in pixels. |
| *bpp* | The number of bits-per-pixel. If this parameter is 0, the function will use the value returned by GetDeviceCaps(hDC, BITSPIXEL_). |
| *ppvBits* | Out, optional. A pointer to a variable that receives a pointer to the location of the DIB bit values. Can be NULL. |

#### Return value

If the function succeeds, the return value is a handle to the newly created DIB, and *ppvBits* points to the bitmap bit values.

If the function fails, the return value is NULL, and *ppvBits* is NULL. The function can fail if one or more of the input parameters is invalid.

This function can return the following value: ERROR_INVALID_PARAMETER (One or more of the input parameters is invalid).

#### Remarks

You must delete the returned bitmap handle with **DeleteObject** when no longer needed to avoid memory leaks.

You cannot paste a DIB section from one application into another application.

**AfxCreateDIBSection** does not use the **BITMAPINFOHEADER** parameters *biXPelsPerMeter* or *biYPelsPerMeter* and will not provide resolution information in the **BITMAPINFO** structure.

#### Usage example

```
DIM hdcWindow AS HDC, hbmp AS HBITMAP, pvBits AS ANY PTR
hdcWindow = GetWindowDC(hwnd)   ' where hwnd is the handle of the wanted window or control
hbmp = AfxCreateDIBSection(hdcWindow, 10, 10, @pvBits)
ReleaseDC(hwnd, hdcWindow)
```

# <a name="AfxDibLoadImage"></a>AfxDibLoadImage

Loads a DIB in memory and returns a pointer to it.

```
FUNCTION AfxDibLoadImage (BYVAL pwszFileName AS WSTRING PTR) AS BITMAPFILEHEADER PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszFileName* | Path of the bitmap file. |

#### Return value

A pointer to the bitmap file header. You must release it with **CoTaskMemFree** when no longer needed.


# <a name="AfxDibSaveImage"></a>AfxDibSaveImage

Saves a DIB to a file.

```
FUNCTION AfxDibSaveImage (BYVAL pwszFileName AS WSTRING PTR, BYVAL pbmfh AS BITMAPFILEHEADER PTR) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszFileName* | Path of the bitmap file. |
| *pbmfh* | Pointer to the bitmap file header. |

#### Return value

TRUE if the DIB has been saved successfully; FALSE otherwise.

# <a name="AfxClearClipboard"></a>AfxClearClipboard

Clears the contents of the clipboard.

```
FUNCTION AfxClearClipboard () AS LONG
```

#### Return value

If the function succeeds, the return value is nonzero. If the function fails, the return value is zero.

# <a name="AfxGetClipboardData"></a>AfxGetClipboardData

Retrieves data from the clipboard in the specified format.

```
FUNCTION AfxGetClipboardData (BYVAL cfFormat AS DWORD) AS HGLOBAL
```

| Parameter  | Description |
| ---------- | ----------- |
| *cfFormat* | The clipboard format. This parameter can be a registered format or any of the standard clipboard formats. |

Standard clipboard formats:

| Constant/Value  | Description |
| --------------- | ----------- |
| CF_BITMAP | A handle to a bitmap. |
| CF_DIB | A memory object containing a BITMAPINFO structure followed by the bitmap bits. |
| CF_DIBV5 | A memory object containing a BITMAPV5HEADER structure followed by the bitmap color space information and the bitmap bits. |
| CF_DIF | Software Arts' Data Interchange Format. |
| CF_DSPBITMAP | Bitmap display format associated with a private format. The *hMem* parameter must be a handle to data that can be displayed in bitmap format in lieu of the privately formatted data. |
| CF_DSPENHMETAFILE | Enhanced metafile display format associated with a private format. The *hMem* parameter must be a handle to data that can be displayed in enhanced metafile format in lieu of the privately formatted data. |
| CF_DSPMETAFILEPICT | Metafile-picture display format associated with a private format. The *hMem* parameter must be a handle to data that can be displayed in metafile-picture format in lieu of the privately formatted data. |
| CF_DSPTEXT | Text display format associated with a private format. The *hMem* parameter must be a handle to data that can be displayed in text format in lieu of the privately formatted data. |
| CF_ENHMETAFILE | A handle to an enhanced metafile. |
| CF_GDIOBJFIRST | Start of a range of integer values for application-defined GDI object clipboard formats. The end of the range is CF_GDIOBJLAST. Handles associated with clipboard formats in this range are not automatically deleted using the **GlobalFree** function when the clipboard is emptied. Also, when using values in this range, the hMem parameter is not a handle to a GDI object, but is a handle allocated by the **GlobalAlloc** function with the GMEM_MOVEABLE flag. |
| CF_GDIOBJLAST | See CF_GDIOBJFIRST. |
| CF_HDROP | A handle that identifies a list of files. An application can retrieve information about the files by passing the handle to the **DragQueryFile** function. |
| CF_LOCALE | The data is a handle to the locale identifier associated with text in the clipboard. When you close the clipboard, if it contains CF_TEXT data but no CF_LOCALE data, the system automatically sets the CF_LOCALE format to the current input language. You can use the CF_LOCALE format to associate a different locale with the clipboard text.<br>An application that pastes text from the clipboard can retrieve this format to determine which character set was used to generate the text.<br>Note that the clipboard does not support plain text in multiple character sets. To achieve this, use a formatted text data type such as RTF instead.<br>The system uses the code page associated with CF_LOCALE to implicitly convert from CF_TEXT to CF_UNICODETEXT. Therefore, the correct code page table is used for the conversion. |
| CF_METAFILEPICT | Handle to a metafile picture format as defined by the **METAFILEPICT** structure. When passing a CF_METAFILEPICT handle by means of DDE, the application responsible for deleting hMem should also free the metafile referred to by the CF_METAFILEPICT handle. |
| CF_OEMTEXT | Text format containing characters in the OEM character set. Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals the end of the data. |
| CF_OWNERDISPLAY | Owner-display format. The clipboard owner must display and update the clipboard viewer window, and receive the WM_ASKCBFORMATNAME, WM_HSCROLLCLIPBOARD, WM_PAINTCLIPBOARD, WM_SIZECLIPBOARD, and WM_VSCROLLCLIPBOARD messages. The *hMem* parameter must be NULL. |
| CF_PALETTE | Handle to a color palette. Whenever an application places data in the clipboard that depends on or assumes a color palette, it should place the palette on the clipboard as well.<br>If the clipboard contains data in the CF_PALETTE (logical color palette) format, the application should use the** SelectPalette** and **RealizePalette** functions to realize (compare) any other data in the clipboard against that logical palette.<br>When displaying clipboard data, the clipboard always uses as its current palette any object on the clipboard that is in the CF_PALETTE format. |
| CF_PENDATA | Data for the pen extensions to the Microsoft Windows for Pen Computing. |
| CF_PRIVATEFIRST | Start of a range of integer values for private clipboard formats. The range ends with CF_PRIVATELAST. Handles associated with private clipboard formats are not freed automatically; the clipboard owner must free such handles, typically in response to the WM_DESTROYCLIPBOARD message. |
| CF_PRIVATELAST | See CF_PRIVATEFIRST. |
| CF_RIFF | Represents audio data more complex than can be represented in a CF_WAVE standard wave format. |
| CF_SYLK | Microsoft Symbolic Link (SYLK) format. |
| CF_TEXT | Text format. Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals the end of the data. Use this format for ANSI text. |
| CF_TIFF | Tagged-image file format. |
| CF_UNICODETEXT | Unicode text format. Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals the end of the data. |
| CF_WAVE | Represents audio data in one of the standard wave formats, such as 11 kHz or 22 kHz PCM. |

#### Return value

If the function succeeds, the return value is the handle to the data. If the function fails, the return value is NULL.

# <a name="AfxGetClipboardText"></a>AfxGetClipboardText

Returns a text string from the clipboard.

```
FUNCTION AfxGetClipboardText () AS CWSTR
```

#### Return value

The retrieved text string.

# <a name="AfxSetClipboardData"></a>AfxSetClipboardData

Places data on the clipboard in a specified clipboard format.

```
FUNCTION AfxSetClipboardData (BYVAL cfFormat AS DWORD, BYVAL hData AS HANDLE) AS HANDLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *cfFormat* | The clipboard format. This parameter can be a registered format or any of the standard clipboard formats. |
| *hData* | Handle to the data in the specified format. |

Standard clipboard formats:

| Constant/Value  | Description |
| --------------- | ----------- |
| CF_BITMAP | A handle to a bitmap. |
| CF_DIB | A memory object containing a BITMAPINFO structure followed by the bitmap bits. |
| CF_DIBV5 | A memory object containing a BITMAPV5HEADER structure followed by the bitmap color space information and the bitmap bits. |
| CF_DIF | Software Arts' Data Interchange Format. |
| CF_DSPBITMAP | Bitmap display format associated with a private format. The *hMem* parameter must be a handle to data that can be displayed in bitmap format in lieu of the privately formatted data. |
| CF_DSPENHMETAFILE | Enhanced metafile display format associated with a private format. The *hMem* parameter must be a handle to data that can be displayed in enhanced metafile format in lieu of the privately formatted data. |
| CF_DSPMETAFILEPICT | Metafile-picture display format associated with a private format. The *hMem* parameter must be a handle to data that can be displayed in metafile-picture format in lieu of the privately formatted data. |
| CF_DSPTEXT | Text display format associated with a private format. The *hMem* parameter must be a handle to data that can be displayed in text format in lieu of the privately formatted data. |
| CF_ENHMETAFILE | A handle to an enhanced metafile. |
| CF_GDIOBJFIRST | Start of a range of integer values for application-defined GDI object clipboard formats. The end of the range is CF_GDIOBJLAST. Handles associated with clipboard formats in this range are not automatically deleted using the **GlobalFree** function when the clipboard is emptied. Also, when using values in this range, the hMem parameter is not a handle to a GDI object, but is a handle allocated by the **GlobalAlloc** function with the GMEM_MOVEABLE flag. |
| CF_GDIOBJLAST | See CF_GDIOBJFIRST. |
| CF_HDROP | A handle that identifies a list of files. An application can retrieve information about the files by passing the handle to the **DragQueryFile** function. |
| CF_LOCALE | The data is a handle to the locale identifier associated with text in the clipboard. When you close the clipboard, if it contains CF_TEXT data but no CF_LOCALE data, the system automatically sets the CF_LOCALE format to the current input language. You can use the CF_LOCALE format to associate a different locale with the clipboard text.<br>An application that pastes text from the clipboard can retrieve this format to determine which character set was used to generate the text.<br>Note that the clipboard does not support plain text in multiple character sets. To achieve this, use a formatted text data type such as RTF instead.<br>The system uses the code page associated with CF_LOCALE to implicitly convert from CF_TEXT to CF_UNICODETEXT. Therefore, the correct code page table is used for the conversion. |
| CF_METAFILEPICT | Handle to a metafile picture format as defined by the **METAFILEPICT** structure. When passing a CF_METAFILEPICT handle by means of DDE, the application responsible for deleting hMem should also free the metafile referred to by the CF_METAFILEPICT handle. |
| CF_OEMTEXT | Text format containing characters in the OEM character set. Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals the end of the data. |
| CF_OWNERDISPLAY | Owner-display format. The clipboard owner must display and update the clipboard viewer window, and receive the WM_ASKCBFORMATNAME, WM_HSCROLLCLIPBOARD, WM_PAINTCLIPBOARD, WM_SIZECLIPBOARD, and WM_VSCROLLCLIPBOARD messages. The *hMem* parameter must be NULL. |
| CF_PALETTE | Handle to a color palette. Whenever an application places data in the clipboard that depends on or assumes a color palette, it should place the palette on the clipboard as well.<br>If the clipboard contains data in the CF_PALETTE (logical color palette) format, the application should use the** SelectPalette** and **RealizePalette** functions to realize (compare) any other data in the clipboard against that logical palette.<br>When displaying clipboard data, the clipboard always uses as its current palette any object on the clipboard that is in the CF_PALETTE format. |
| CF_PENDATA | Data for the pen extensions to the Microsoft Windows for Pen Computing. |
| CF_PRIVATEFIRST | Start of a range of integer values for private clipboard formats. The range ends with CF_PRIVATELAST. Handles associated with private clipboard formats are not freed automatically; the clipboard owner must free such handles, typically in response to the WM_DESTROYCLIPBOARD message. |
| CF_PRIVATELAST | See CF_PRIVATEFIRST. |
| CF_RIFF | Represents audio data more complex than can be represented in a CF_WAVE standard wave format. |
| CF_SYLK | Microsoft Symbolic Link (SYLK) format. |
| CF_TEXT | Text format. Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals the end of the data. Use this format for ANSI text. |
| CF_TIFF | Tagged-image file format. |
| CF_UNICODETEXT | Unicode text format. Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals the end of the data. |
| CF_WAVE | Represents audio data in one of the standard wave formats, such as 11 kHz or 22 kHz PCM. |

#### Return value

If the function succeeds, the return value is the handle to the data. If the function fails, the return value is NULL.

#### Remarks

If **AfxSetClipboardData** succeeds, the system owns the object identified by the *hData* parameter. The application may not write to or free the data once ownership has been transferred to the system.

# <a name="AfxSetClipboardText"></a>AfxSetClipboardText

Places a text string into the clipboard.

```
FUNCTION AfxSetClipboardText (BYREF wszText AS WSTRING) AS HANDLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszText* | The text to be placed in the clipboard. |

#### Return value

If the function succeeds, the return value is the handle to the data. If the function fails, the return value is NULL.

# <a name="AfxCreateFont"></a>AfxCreateFont

Creates a logical font.

```
FUNCTION AfxCreateFont (BYREF wszFaceName AS WSTRING, BYVAL lPointSize AS LONG, BYVAL DPI AS LONG = 96, _
   BYVAL lWeight AS LONG = 0, BYVAL bItalic AS UBYTE = FALSE, BYVAL bUnderline AS UBYTE = FALSE, _
   BYVAL bStrikeOut AS UBYTE = FALSE, BYVAL bCharSet AS UBYTE = DEFAULT_CHARSET) AS HFONT
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFaceName* | The typeface name. |
| *lPointSize* | The point size. |
| *DPI* | Dots per inch to calculate scaling. Default value = 96 (no scaling). If you pass -1 and the application is DPI aware, the DPI value used by the operating system will be used. |
| *lWeight* | Initial weight of the font. If the weight is below 550 (the average of FW_NORMAL, 400, and FW_BOLD, 700), then the Bold property is also initialized to FALSE. If the weight is above 550, the Bold property is set to TRUE.<br>The following values are defined for convenience: FW_DONTCARE (0), FW_THIN (100), FW_EXTRALIGHT (200), FW_ULTRALIGHT (200), FW_LIGHT (300), FW_NORMAL (400), FW_REGULAR (400), FW_MEDIUM (500), FW_SEMIBOLD (600), FW_DEMIBOLD (600), FW_BOLD (700), FW_EXTRABOLD (800), FW_ULTRABOLD (800), FW_HEAVY (900), FW_BLACK (900) |
| *bItalic* | Italic flag. CTRUE or FALSE. |
| *bUnderline* | Underline flag. CTRUE or FALSE. |
| *bStrikeOut* | StrikeOut flag. CTRUE or FALSE |
| *bCharSet* | Specifies the character set. The following values are predefined:<br>ANSI_CHARSET, BALTIC_CHARSET, CHINESEBIG5_CHARSET, DEFAULT_CHARSET, EASTEUROPE_CHARSET, GB2312_CHARSET, GREEK_CHARSET, HANGUL_CHARSET, MAC_CHARSET, OEM_CHARSET, RUSSIAN_CHARSET, SHIFTJIS_CHARSET, SYMBOL_CHARSET, TURKISH_CHARSET.<br>Korean Windows: JOHAB_CHARSET.<br>Middle-Eastern Windows: HEBREW_CHARSET, ARABIC_CHARSET.<br>Thai Windows: THAI_CHARSET.<br>The OEM_CHARSET value specifies a character set that is operating-system dependent. DEFAULT_CHARSET is set to a value based on the current system locale. For example, when the system locale is English (United States), it is set as ANSI_CHARSET. Fonts with other character sets may exist in the operating system. If an application uses a font with an unknown character set, it should not attempt to translate or interpret strings that are rendered with that font. This parameter is important in the font mapping process. To ensure consistent results, specify a specific character set. If you specify a typeface name in the *wszFaceName* parameter, make sure that the *bCharSet* value matches the character set of the typeface specified in *wszFaceName*. |

#### Return value

The handle of the font or NULL on failure.

#### Remarks

The returned font must be destroyed with **DeleteObject** or the macro **DeleteFont** when no longer needed to prevent memory leaks.

#### Usage examples

```
hFont = AfxCreateFont("MS Sans Serif", 8, , FW_NORMAL, , , , DEFAULT_CHARSET)
hFont = AfxCreateFont("Courier New", 10, 96 , FW_BOLD, , , , DEFAULT_CHARSET)
hFont = AfxCreateFont("Marlett", 8, -1, FW_NORMAL, , , , SYMBOL_CHARSET)
```

# <a name="AfxGetFontHeight"></a>AfxGetFontHeight

Returns the logical height of a font given its point size.

```
FUNCTION AfxGetFontHeight (BYVAL nPointSize AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nPointSize* | The point size of the font. |

# <a name="AfxGetFontPointSize"></a>AfxGetFontPointSize

Returns the point size of a font given its logical height.

```
FUNCTION AfxGetFontPointSize (BYVAL nHeight AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nHeight* | The logical height of the font. |

# <a name="AfxGetWindowFont"></a>AfxGetWindowFont

Retrieves the font with which the control is currently drawing its text.

```
FUNCTION AfxGetWindowFont (BYVAL hwnd AS HWND) AS HFONT
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | A handle to a window. |

# <a name="AfxGetWindowsFontInfo"></a>AfxGetWindowsFontInfo

Retrieves information about the fonts used by Windows.

```
FUNCTION AfxGetWindowsFontInfo (BYVAL nType AS LONG, BYVAL plfw AS LOGFONTW PTR) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *nType* | The type of the font: AFX_FONT_CAPTION, AFX_FONT_SMALLCAPTION, AFX_FONT_MENU, AFX_FONT_STATUS, AFX_FONT_MESSAGE. |
| *plfw* | Pointer to a LOGFONTW structure that receives the font information. |

#### Return value

TRUE on succes or FALSE on failure. To get extended error information, call **GetLastError**.

# <a name="AfxGetWindowsFontPointSize"></a>AfxGetWindowsFontPointSize

Retrieves the point size of the fonts used by Windows.

```
FUNCTION AfxGetWindowsFontPointSize (BYVAL nType AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nType* | The type of the font: AFX_FONT_CAPTION, AFX_FONT_SMALLCAPTION, AFX_FONT_MENU, AFX_FONT_STATUS, AFX_FONT_MESSAGE. |

# <a name="AfxModifyFontFaceName"></a>AfxModifyFontFaceName

Modifies the face name of the font of a window or control.

```
FUNCTION AfxModifyFontFaceName (BYVAL hwnd AS HWND, BYREF wszNewFaceName AS WSTRING) AS HFONT
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window or control. |
| *wszNewFaceName* | The new face name of the font. |

#### Return value

The handle of the new font on success, or NULL on failure.

To get extended error information call **GetLastError**.

#### Remarks

The returned font must be destroyed with **DeleteObject** or the macro **DeleteFont** when no longer needed to prevent memory leaks.

# <a name="AfxModifyFontHeight"></a>AfxModifyFontHeight

Modifies the height of the font used by a window of control.

```
FUNCTION AfxModifyFontHeight (BYVAL hwnd AS HWND, BYVAL nValue AS LONG) AS HFONT
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window or control. |
| *nValue* | The base is 100. To increase the font a 20% pass 120; to reduce it a 20% pass 80%. |

#### Return value

The handle of the new font on success, or NULL on failure.

To get extended error information call **GetLastError**.

#### Remarks

The returned font must be destroyed with **DeleteObject** or the macro **DeleteFont** when no longer needed to prevent memory leaks.

# <a name="AfxModifyFontSettings"></a>AfxModifyFontSettings

Modifies settings of the font used by a window of control.

```
FUNCTION AfxModifyFontSettings (BYVAL hwnd AS HWND, BYVAL nSetting AS LONG, BYVAL nValue AS LONG) AS HFONT
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window or control. |
| *nSetting* | One of the AFX_FONT_xxx constants (see below). |
| *nValue* | Depends of the nSetting value. **AFX_FONT_HEIGHT** : The base is 100. To increase the font a 20% pass 120; to reduce it a 20% pass 80%.<br>**AFX_FONT_WEIGHT** : The weight of the font in the range 0 through 1000. For example, 400 is normal and  700 is bold. If this value is zero, a default weight is used. The following values are defined for convenience. FW_DONTCARE (0), FW_THIN (100), FW_EXTRALIGHT (200), FW_ULTRALIGHT (200), FW_LIGHT (300), FW_NORMAL (400), FW_REGULAR (400), FW_MEDIUM (500), FW_SEMIBOLD (600), FW_DEMIBOLD (600), FW_BOLD (700), FW_EXTRABOLD (800), FW_ULTRABOLD (800), FW_HEAVY (900), FW_BLACK (900)<br>**AFX_FONT_ITALIC** : TRUE or FALSE.<br>**AFX_FONT_UNDERLINE** : TRUE or FALSE.<br>**AFX_FONT_STRIKEOUT** : TRUE or FALSE.<br>**AFX_FONT_CHARSET**: The following values are predefined: ANSI_CHARSET, BALTIC_CHARSET, CHINESEBIG5_CHARSET, DEFAULT_CHARSET, EASTEUROPE_CHARSET, GB2312_CHARSET, GREEK_CHARSET, HANGUL_CHARSET, MAC_CHARSET, OEM_CHARSET, RUSSIAN_CHARSET, SHIFTJIS_CHARSET, SYMBOL_CHARSET, TURKISH_CHARSET, VIETNAMESE_CHARSET, JOHAB_CHARSET (Korean language edition of Windows), ARABIC_CHARSET and HEBREW_CHARSET (Middle East language edition of Windows), THAI_CHARSET (Thai language edition of Windows). The OEM_CHARSET value specifies a character set that is operating-system dependent. DEFAULT_CHARSET is set to a value based on the current system locale. For example, when the system locale is English (United States), it is set as ANSI_CHARSET. |

#### Return value

The handle of the new font on success, or NULL on failure.

To get extended error information call **GetLastError**.

#### Remarks

The returned font must be destroyed with **DeleteObject** or the macro **DeleteFont** when no longer needed to prevent memory leaks.

# <a name="AfxSetWindowFont"></a>AfxSetWindowFont

Sets the font that a control is to use when drawing text.

```
SUB AfxSetWindowFont (BYVAL hwnd AS HWND, BYVAL hFont AS HFONT, BYVAL fRedraw AS LONG = CTRUE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window or control. |
| *hFont* | A handle to the font. If this parameter is NULL, the control uses the default system font to draw text. |
| *fRedraw* | Optional. Specifies whether the control should be redrawn immediately upon setting the font. If this parameter is CTRUE, the control redraws itself. |

#### Return value

The handle of the new font on success, or NULL on failure.

To get extended error information call **GetLastError**.

#### Remarks

The application should call the **DeleteObject** function to delete the font when it is no longer needed; for example, after it destroys the control.

The size of the control does not change as a result of receiving this message. To avoid clipping text that does not fit within the boundaries of the control, the application should correct the size of the control window before it sets the font.

# <a name="AfxForceVisibleDisplay"></a>AfxForceVisibleDisplay

If you use dual (or even triple/quad) displays then you have undoubtedly encountered the following situation: You change the physical order of your displays, or otherwise reconfigure the logical ordering using your display software. This sometimes has the side-effect of changing your desktop coordinates from zero-based to negative starting coordinates (i.e. the top-left coordinate of your desktop changes from 0,0 to -1024,-768).

This effects many Windows programs which restore their last on-screen position whenever they are started. Should the user reorder their display configuration this can sometimes result in a Windows program subsequently starting in an off-screen position (i.e. at a location that used to be visible) - and is now effectively invisible, preventing the user from closing it down or otherwise moving it back on-screen.

The **AfxForceVisibleDisplay** function can be called at program start-time right after the main window has been created and positioned 'on-screen'. Should the window be positioned in an off-screen position, it is forced back onto the nearest display to its last position. The user will be unaware this is happening and won't even realize to thank you for keeping their user-interface visible, even though they changed their display settings.

Source: Catch-22 web site.

```
SUB AfxForceVisibleDisplay (BYVAL hwnd AS HWND)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |

# <a name="AfxGetDisplayBitsPerPixel"></a>AfxGetDisplayBitsPerPixel

Returns the color resolution, in bits per pixel, of the display device.

```
FUNCTION AfxGetDisplayBitsPerPixel () AS DWORD
```

# <a name="AfxGetDisplayFrequency"></a>AfxGetDisplayFrequency

Returns the frequency, in hertz (cycles per second), of the display device in a particular mode. This value is also known as the display device's vertical refresh rate.

```
FUNCTION AfxGetDisplayBitsPerPixel () AS DWORD
```

# <a name="AfxGetDisplayPixelsHeight"></a>AfxGetDisplayPixelsHeight

Returns the height, in pixels, of the current display device on the computer on which the calling thread is running.

```
FUNCTION AfxGetDisplayPixelsHeight () AS DWORD
```

# <a name="AfxGetDisplayPixelsWidth"></a>AfxGetDisplayPixelsWidth

Returns the width, in pixels, of the current display device on the computer on which the calling thread is running.

```
FUNCTION AfxGetDisplayPixelsWidth () AS DWORD
```

#### Remarks

Contrarily to **GetSystemMetrics** or **GetDeviceCaps**, it returns the real width even when it is called from an application that is not DPI aware, e.g. an application running virtualized in a monitor 1920 pixels width and a DPI of 192, will return 960 pixels if it calls **GetSystemMetrics** or **GetDeviceCaps**, but will return 1920 pixels calling **AfxGetDisplayPixelsWidth**.

# <a name="AfxLogPixelsX"></a>AfxLogPixelsX / AfxGetDpi / AfxGetDpiX

Retrieves the number of pixels per logical inch along the screen width. In a system with multiple display monitors, this value is the same for all monitors.

```
FUNCTION AfxLogPixelsX () AS LONG
FUNCTION AfxGetDpi () AS LONG
FUNCTION AfxGetDpiX () AS LONG
```

# <a name="AfxLogPixelsY"></a>AfxLogPixelsX / AfxGetDpiY

Retrieves the number of pixels per logical inch along the screen height. In a system with multiple display monitors, this value is the same for all monitors.

```
FUNCTION AfxLogPixelsY () AS LONG
FUNCTION AfxGetDpiY () AS LONG
```

# <a name="AfxGetMonitorHorizontalScaling"></a>AfxGetMonitorHorizontalScaling

Returns the horizontal scaling of the monitor that the window is currently displayed on.

```
FUNCTION AfxGetMonitorHorizontalScaling (BYVAL hwnd AS HWND = NULL) AS DWORD
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Optional. A handle to the window. If NULL, the desktop window handle will be used. |

#### Remarks

If the application to which the window belongs is not DPI aware, a computer using 192 DPI, will return an scaling ratio of 2.

# <a name="AfxGetMonitorVerticalScaling"></a>AfxGetMonitorVerticalScaling

Returns the vertical scaling of the monitor that the window is currently displayed on.

```
FUNCTION AfxGetMonitorVerticalScaling (BYVAL hwnd AS HWND = NULL) AS DWORD
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Optional. A handle to the window. If NULL, the desktop window handle will be used. |

#### Remarks

If the application to which the window belongs is not DPI aware, a computer using 192 DPI, will return an scaling ratio of 2.

# <a name="AfxGetMonitorLogicalHeight"></a>AfxGetMonitorLogicalHeight

Returns the logical height of the monitor that the window is currently displayed on.

```
FUNCTION AfxGetMonitorLogicalHeight (BYVAL hwnd AS HWND = NULL) AS DWORD
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Optional. A handle to the window. If NULL, the desktop window handle will be used. |

#### Remarks

If the application to which the window belongs is not DPI aware, a monitor with an height resolution of 1080 pixels in a computer using 192 DPI, will return 540 pixels.

# <a name="AfxGetMonitorLogicalWidth"></a>AfxGetMonitorLogicalWidth

Returns the logical width of the monitor that the window is currently displayed on.

```
FUNCTION AfxGetMonitorLogicalWidth (BYVAL hwnd AS HWND = NULL) AS DWORD
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Optional. A handle to the window. If NULL, the desktop window handle will be used. |

#### Remarks

If the application to which the window belongs is not DPI aware, a monitor with a width resolution of 1920 pixels in a computer using 192 DPI, will return 960 pixels.

# <a name="AfxIsDPIResolutionAtLeast"></a>AfxIsDPIResolutionAtLeast

Determines if screen resolution meets minimum requirements in relative pixels, e.g. for a screen resolution of 1920x1080 pixels and a DPI of 192 (scaling ratio = 2), the maximum relative pixels for a DPI aware application is 960x540.

```
FUNCTION AfxIsDPIResolutionAtLeast (BYVAL cxMin AS LONG, BYVAL cyMin AS LONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *cxMin* | Minimum screen resolution width in relative pixels. |
| *cyMin* | Minimum screen resolution height in relative pixels. |

#### Return value

TRUE or FALSE.

# <a name="AfxIsResolutionAtLeast"></a>AfxIsResolutionAtLeast

Determines if screen resolution meets minimum requirements.

```
FUNCTION AfxIsResolutionAtLeast (BYVAL cxMin AS LONG, BYVAL cyMin AS LONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *cxMin* | Minimum screen resolution width in relative pixels. |
| *cyMin* | Minimum screen resolution height in relative pixels. |

#### Return value

TRUE or FALSE.

# <a name="AfxSetProcessDPIAware"></a>AfxSetProcessDPIAware

Sets the current process as dots per inch (dpi) aware.

Note: **AfxSetProcessDPIAware** is subject to a possible race condition if a DLL caches dpi settings during initialization. For this reason, it is recommended that dpi-aware be set through the application (.exe) manifest rather than by calling **AfxSetProcessDPIAware**.

```
FUNCTION AfxSetProcessDPIAware () AS BOOLEAN
```

#### Return value

If the function succeeds, the return value is TRUE. Otherwise, the return value is FALSE.

#### Remarks

DLLs should accept the dpi setting of the host process rather than call **AfxSetProcessDPIAware** themselves. To be set properly, *dpiAware* should be specified as part of the application (.exe) manifest. (*dpiAware* defined in an embedded DLL manifest has no affect.) The following markup shows how to set *dpiAware* as part of an application (.exe) manifest.

```
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0" xmlns:asmv3="urn:schemas-microsoft-com:asm.v3" >
 ...
  <asmv3:application>
    <asmv3:windowsSettings xmlns="http://schemas.microsoft.com/SMI/2005/WindowsSettings">
      <dpiAware>true</dpiAware>
    </asmv3:windowsSettings>
  </asmv3:application>
 ...
</assembly>
```

# <a name="AfxIsProcessDPIAware"></a>AfxIsProcessDPIAware

Determines whether the current process is dots per inch (dpi) aware such that it adjusts the sizes of UI elements to compensate for the dpi setting.

```
FUNCTION AfxIsProcessDPIAware () AS BOOLEAN
```

#### Return value

TRUE if the process is dpi aware; otherwise, FALSE.

# <a name="AfxLoadIconMetric"></a>AfxLoadIconMetric

Loads a specified icon resource with a client-specified system metric.

```
FUNCTION AfxLoadIconMetric (BYVAL hinst AS HINSTANCE, BYVAL pwszName AS WSTRING PTR, _
   BYVAL lims AS LONG, BYVAL phico AS HICON PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *hinst* | A handle to the module of either a DLL or executable (.exe) file that contains the icon to be loaded. For more information, see **GetModuleHandle**. To load a predefined icon or a standalone icon file, set this parameter to NULL. |
| *pwszName* | A pointer to a null-terminated, Unicode buffer that contains location information about the icon to load. It is interpreted as follows: If *hinst* is NULL, *pwszName* can specify one of two things.<br>1) The identifier of a predefined icon to load. These identifiers are recognized: IDI_APPLICATION, IDI_INFORMATION, IDI_ERROR, IDI_WARNING, IDI_SHIELD, IDI_QUESTION.<br>To pass these constants to the **AfxLoadIconMetric** function, use the MAKEINTRESOURCE macro. For example, to load the IDI_ERROR icon, pass MAKEINTRESOURCE(IDI_ERROR) as the *pwszName* parameter and NULL as the *hinst* parameter.<br>2) The name of a standalone icon (.ico) file.<br>If hinst is non-null, *pwszName* can specify one of two things.<br>1) The name of the icon resource, if the icon resource is to be loaded by name from the module.<br>2) The icon ordinal, if the icon resource is to be loaded by ordinal from the module. This ordinal must be packaged by using the MAKEINTRESOURCE macro. |
| *lims* | The desired metric. One of the following values:<br>**LIM_SMALL** : Corresponds to SM_CXSMICON, the recommended pixel width of a small icon.<br>**LIM_LARGE** : Corresponds to SM_CXICON, the default pixel width of an icon. |
| *phico* | When this function returns, contains a pointer to the handle of the loaded icon. |

#### Return value

Returns S_OK if successful, otherwise an error, including the following value: *E_INVALIDARG* : The contents of the buffer pointed to by pszName do not fit any of the expected interpretations.

#### Remarks

**LoadIconMetric** is similar to **LoadIcon**, but with the capability to specify the icon metric. It is used in place of **LoadIcon** when the calling application wants to ensure a high quality icon. This is particularly useful in high dots per inch (dpi) situations.

Icons are extracted or created as follows.

1) If an exact size match is found in the resource, that icon is used.<br>
2) If an exact size match cannot be found and a larger icon is available, a new icon is created by scaling the larger version down to the desired size.<br>
3) If an exact size match cannot be found and no larger icon is available, a new icon is created by scaling a smaller icon up to the desired size.

# <a name="AfxScaleRatioX"></a>AfxScaleRatioX

Retrieves the desktop horizontal scaling ratio.

```
FUNCTION AfxScaleRatioX () AS LONG
```

# <a name="AfxScaleRatioY"></a>AfxScaleRatioY

Retrieves the desktop vertical scaling ratio.

```
FUNCTION AfxScaleRatioY () AS LONG
```

# <a name="AfxScaleX"></a>AfxScaleX

Scales an horizontal coordinate according the DPI (dots per pixel) being used by the operating system.

```
FUNCTION AfxScaleX (BYVAL cx AS SINGLE) AS SINGLE
```

Return value

The scaled coordinate.

# <a name="AfxScaleY"></a>AfxScaleY

Scales a vertical coordinate according the DPI (dots per pixel) being used by the operating system.

```
FUNCTION AfxScaleY (BYVAL cx AS SINGLE) AS SINGLE
```

Return value

The scaled coordinate.

# <a name="AfxUnscaleX"></a>AfxUnscaleX

Unscales an horizontal coordinate according the DPI (dots per pixel) being used by the operating system.

```
FUNCTION AfxUnscaleX (BYVAL cx AS SINGLE) AS SINGLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *cx* | The value of the horizontal coordinate, in pixels. |

Return value

The unscaled coordinate.

# <a name="AfxUnscaleY"></a>AfxUnscaleY

Unscales a vertical coordinate according the DPI (dots per pixel) being used by the operating system.

```
FUNCTION AfxUnscaleY (BYVAL cx AS SINGLE) AS SINGLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *cy* | The value of the vertical coordinate, in pixels. |

Return value

The unscaled coordinate.

# <a name="AfxUseDpiScaling"></a>AfxUseDpiScaling

Returns TRUE if the OS uses DPI scaling; FALSE otherwise.

```
FUNCTION AfxUseDpiScaling () AS BOOLEAN
```

# <a name="AfxAddWindowExStyle"></a>AfxAddWindowExStyle

Adds a new extended style to the specified window.

```
FUNCTION AfxAddWindowExStyle (BYVAL hwnd AS HWND, BYVAL dwExStyle AS DWORD) AS DWORD
```
| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |
| *dwExStyle* | Extended style to add. |

#### Return value

The previous window extended styles.

#### Usage example

```
AfxAddWindowExStyle(hwnd, WS_EX_COMPOSITED)
```

#### Remarks

If the window has a class style of CS_CLASSDC or CS_OWNDC, do not set the extended window styles WS_EX_COMPOSITED or WS_EX_LAYERED.

# <a name="AfxAddWindowStyle"></a>AfxAddWindowStyle

Adds a new style to the specified window.

```
FUNCTION AfxAddWindowStyle (BYVAL hwnd AS HWND, BYVAL dwStyle AS DWORD) AS DWORD
```
| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |
| *dwStyle* | Style to add. |

#### Return value

The previous window styles.

#### Usage example

```
AfxAddWindowStyle(hwnd, WS_HSCROLL)
```

# <a name="AfxGetWindowExStyle"></a>AfxGetWindowExStyle

Retrieves the extended window styles.

```
FUNCTION AfxGetWindowExStyle (BYVAL hwnd AS HWND) AS DWORD

```
| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |

# <a name="AfxGetWindowStyle"></a>AfxGetWindowStyle

Retrieves the window styles.

```
FUNCTION AfxGetWindowStyle (BYVAL hwnd AS HWND) AS DWORD

```
| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |

# <a name="AfxRemoveWindowExStyle"></a>AfxRemoveWindowExStyle

Removes an extended style from the specified window.

```
FUNCTION AfxRemoveWindowExStyle (BYVAL hwnd AS HWND, BYVAL dwExStyle AS DWORD) AS DWORD
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |
| *dwExStyle* | The extended style to remove. |

#### Return value

The previous extended window styles.

# <a name="AfxRemoveWindowStyle"></a>AfxRemoveWindowStyle

Removes a style from the specified window.

```
FUNCTION AfxRemoveWindowStyle (BYVAL hwnd AS HWND, BYVAL dwStyle AS DWORD) AS DWORD
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |
| *dwStyle* | The style to remove. |

#### Return value

The previous window styles.

# <a name="AfxSetWindowExStyle"></a>AfxSetWindowExStyle

Sets the extended style(s) of the specified window.

```
FUNCTION AfxSetWindowExStyle (BYVAL hwnd AS HWND, BYVAL dwExStyle AS DWORD) AS DWORD
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |
| *dwExStyle* | The extended style(s) to set. |

#### Return value

The previous extended window styles.

# <a name="AfxSetWindowStyle"></a>AfxSetWindowStyle

Sets the style(s) of the specified window.

```
FUNCTION AfxSetWindowStyle (BYVAL hwnd AS HWND, BYVAL dwStyle AS DWORD) AS DWORD
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |
| *dwStyle* | The style(s) to set. |

#### Return value

The previous window styles.

# <a name="AfxGetControlHandle"></a>AfxGetControlHandle

Returns the handle of the control with the specified identifier. The reference handle can be the handle of the form or the handle of any other control on the form.

```
FUNCTION AfxGetControlHandle (BYVAL hwnd AS HWND, BYVAL wCtrlID AS WORD) AS HWND
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |
| *wCtrlID* | Control identifier. |

#### Return value

Returns the handle of the control or NULL.

# <a name="AfxGetFormHandle"></a>AfxGetFormHandle

Finds the handle of the top-level window or MDI child window that is the ancestor of the specified window handle. The reference handle is the handle of any control on the form.

```
FUNCTION AfxGetFormHandle (BYVAL hwnd AS HWND) AS HWND
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | A handle to the control. |

#### Return value

Handle of the ancestor window.

# <a name="AfxGetHwndFromPID"></a>AfxGetHwndFromPID

Retrieves a window handle given it's process identifier.

```
FUNCTION AfxGetHwndFromPID (BYVAL PID AS DWORD) AS HWND
```

| Parameter  | Description |
| ---------- | ----------- |
| *PID* | The process identifier. |

#### Return value

The window handle or NULL.

# <a name="AfxGetPathFromWindowHandle"></a>AfxGetPathFromWindowHandle

Retrieves the path of the executable file that created the specified window.

```
FUNCTION AfxGetPathFromWindowHandle (BYVAL hwnd AS HWND) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | The window handle. |

#### Return value

The path of the executable file.

# <a name="AfxCenterWindow"></a>AfxCenterWindow

Centers a window on the screen or over another window. It also ensures that the placement is done within the work area.

```
SUB AfxCenterWindow (BYVAL hwnd AS HWND = NULL, BYVAL hwndParent AS HWND = NULL)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Optional. Handle to the window. |
| *hwndParent* | Optional. Handle to the parent window. |

# <a name="AfxForceSetForegroundWindow"></a>AfxForceSetForegroundWindow

Brings the thread that created the specified window into the foreground and activates the window. Keyboard input is directed to the window, and various visual cues are changed for the user. The system assigns a slightly higher priority to the thread that created the foreground window than it does to other threads.

```
SUB AfxForceSetForegroundWindow (BYVAL hwnd AS HWND)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |

#### Remarks

Replacement for the **SetForegroundWindow** API function, that sometimes fails.

#### Warning

Don't use it. Apparently, **AttachThreadInput** can lead to a deadlock.
See: [The dangers of attaching input queues](https://blogs.msdn.microsoft.com/oldnewthing/20080801-00/?p=21393)

# <a name="AfxGetTopEnabledWindow"></a>AfxGetTopEnabledWindow

Retrieves the handle of the enabled and visible window at the top of the z-order in an application.

```
FUNCTION AfxGetTopEnabledWindow () AS HWND
```

#### Return value

Handle of the window at top of z-order or NULL.

# <a name="AfxGetTopLevelParent"></a>AfxGetTopLevelParent

Retrieves the window's top-level parent window.

```
FUNCTION AfxGetTopEnabledWindow () AS HWND
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |

#### Return value

Handle of the top-level parent window.

# <a name="AfxGetTopLevelWindow"></a>AfxGetTopLevelWindow

Retrieves the window's top-level parent or owner window.

```
FUNCTION AfxGetTopLevelWindow (BYVAL hwnd AS HWND) AS HWND
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |

#### Return value

Handle of the top-level parent or owner window.

# <a name="AfxGetWindowClassName"></a>AfxGetWindowClassName

Retrieves the name of the class to which the specified window belongs. 

```
FUNCTION AfxGetWindowClassName (BYVAL hwnd AS HWND) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |

#### Return value

The name of the class.

# <a name="AfxGetWindowClientHeight"></a>AfxGetWindowClientHeight

Returns the height of the client area of window, in pixels.

```
FUNCTION AfxGetWindowClientHeight (BYVAL hwnd AS HWND) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |

# <a name="AfxGetWindowClientRect"></a>AfxGetWindowClientRect

Retrieves the coordinates of a window's client area. The client coordinates specify the upper-left and lower-right corners of the client area. Because client coordinates are relative to the upper-left corner of a window's client area, the coordinates of the upper-left corner are (0,0).

```
FUNCTION AfxGetWindowClientRect (BYVAL hwnd AS HWND) AS RECT
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |

#### Return value

A RECT structure with the retrieved coordinates of the window's client area.

# <a name="AfxGetWindowClientWidth"></a>AfxGetWindowClientWidth

Returns the width of the client area of a window, in pixels.

```
FUNCTION AfxGetWindowClientWidth (BYVAL hwnd AS HWND) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |

# <a name="AfxGetWindowHeight"></a>AfxGetWindowHeight

Returns the height of a window, in pixels.

```
FUNCTION AfxGetWindowHeight (BYVAL hwnd AS HWND) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |

# <a name="AfxGetWindowWidth"></a>AfxGetWindowWidth

Returns the width of a window, in pixels.

```
FUNCTION AfxGetWindowWidth (BYVAL hwnd AS HWND) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |

# <a name="AfxGetWindowLocation"></a>AfxGetWindowLocation

Returns the location of the top left corner of the window, in pixels. The location is relative to the upper-left corner of the client area in the parent window.

```
SUB AfxGetWindowLocation (BYVAL hwnd AS HWND, BYREF nLeft AS LONG, BYREF nTop AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |
| *nLeft* | Out. The horizontal location. |
| *nTop* | Out. The vertical location. |

# <a name="AfxGetWindowRect"></a>AfxGetWindowRect

Retrieves the dimensions of the bounding rectangle of the specified window. The dimensions are given in screen coordinates that are relative to the upper-left corner of the screen.

```
FUNCTION AfxGetWindowRect (BYVAL hwnd AS HWND) AS RECT
```

#### Return value

A RECT structure with the retrieved dimensions.

# <a name="AfxGetWindowSize"></a>AfxGetWindowSize

Gets the width and height of the specified window, in pixels.

```
FUNCTION AfxGetWindowSize (BYVAL hwnd AS HWND, BYVAL nWidth AS LONG, BYVAL nHeight AS LONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |
| *nWidth* | The width of the window. |
| *nHeight* | The height of the window. |

# <a name="AfxGetWindowText"></a>AfxGetWindowText

Gets the text of a window. This function can also be used to retrieve the text of buttons, edit and static controls.

```
FUNCTION AfxGetWindowText (BYVAL hwnd AS HWND) AS CWSTR
```

#### Return value

The text of the window.

For an edit control, the text returned is the content of the edit control. For a combo box, the text is the content of the edit control (or static-text) portion of the combo box. For a button, the text is the button name. For other windows, the text is the window title. To retrieve the text of an item in a list box, an application can use the **ListBox_GetText** function. 

Rich Edit: If the text to be copied exceeds 64K, use either the EM_STREAMOUT or EM_GETSELTEXT message.

#### Remarks

This function uses the WM_GETTEXT message because **GetWindowText** cannot retrieve the text of a window in another application.

#### Usage examples

DIM cwsText AS CWSTR = AfxGetWindowText(hwnd)
MessageBoxW 0, cwsText, "", MB_OK

# <a name="AfxGetWindowTextLength"></a>AfxGetWindowTextLength

Retrieves the length of the text of a window. This function can also be used to retrieve the length of the text of buttons, edit and static controls.

```
FUNCTION AfxGetWindowTextLength (BYVAL hwnd AS HWND) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |

#### Return value

If the function succeeds, the return value is the length of the text in characters, not including the terminating null character.

If the function fails, the return value is zero.

For an edit control, the text returned is the content of the edit control. For a combo box, the text is the content of the edit control (or static-text) portion of the combo box. For a button, the text is the button name. For other windows, the text is the window title. To retrieve the text of an item in a list box, an application can use the **ListBox_GetTextLength** function. 

**AfxGetWindowTextLength** sends a WM_GETTEXTLENGTH message. When the WM_GETTEXTLENGTH message is sent, the **DefWindowProc** function returns the length, in characters, of the text. Under certain conditions, the **DefWindowProc** function returns a value that is larger than the actual length of the text. This occurs with certain mixtures of ANSI and Unicode, and is due to the system allowing for the possible existence of double-byte character set (DBCS) characters within the text. The return value, however, will always be at least as large as the actual length of the text; you can thus always use it to guide buffer allocation. This behavior can occur when an application uses both ANSI functions and common dialogs, which use Unicode.

To obtain the exact length of the text, use the WM_GETTEXT, LB_GETTEXT, or CB_GETLBTEXT messages, or the **GetWindowText** function.

Sending a WM_GETTEXTLENGTH message to a non-text static control, such as a static bitmap or static icon control, does not return a string value. Instead, it returns zero.

# <a name="AfxGetWorkAreaHeight"></a>AfxGetWorkAreaHeight

Retrieves the height of the work area on the primary display monitor expressed in virtual screen coordinates. The work area is the portion of the screen not obscured by the system taskbar or by application desktop toolbars. To get the work area of a monitor other than the primary display monitor, call the **GetMonitorInfo** function.

```
FUNCTION AfxGetWorkAreaHeight () AS LONG
```

# <a name="AfxGetWorkAreaRect"></a>AfxGetWorkAreaRect

Retrieves the coordinates of the work area on the primary display monitor expressed in virtual screen coordinates. The work area is the portion of the screen not obscured by the system taskbar or by application desktop toolbars. To get the work area of a monitor other than the primary display monitor, call the **GetMonitorInfo** function.

```
FUNCTION AfxGetWorkAreaRect () AS RECT
```

#### Return value

A RECT structure with the retrieved coordinates.

# <a name="AfxGetWorkAreaWidth"></a>AfxGetWorkAreaWidth

Retrieves the width of the work area on the primary display monitor expressed in virtual screen coordinates. The work area is the portion of the screen not obscured by the system taskbar or by application desktop toolbars. To get the work area of a monitor other than the primary display monitor, call the **GetMonitorInfo** function.

```
FUNCTION AfxGetWorkAreaWidth () AS LONG
```

# <a name="AfxRedrawNonClientArea"></a>AfxRedrawNonClientArea

Redraws the non-client area of the specified window.

```
FUNCTION AfxRedrawNonClientArea (BYVAL hwnd AS HWND) AS BOOLEAN
```

#### Return value

If the function succeeds, the return value is TRUE.

If the function fails, the return value is FALSE. To get extended error information, call **GetLastError**.

# <a name="AfxRedrawWindow"></a>AfxRedrawWindow

Redraws the specified window.

```
SUB AfxRedrawWindow (BYVAL hwnd AS HWND)
```

# <a name="AfxSetWindowClientSize"></a>AfxSetWindowClientSize

Adjusts the bounding rectangle of a window based on the desired size of the client area.

```
SUB AfxSetWindowClientSize (BYVAL hwnd AS HWND, BYVAL nWidth AS LONG, BYVAL nHeight AS LONG, _
   BYVAL rxRatio AS SINGLE = 1, BYVAL ryRatio AS SINGLE = 1)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |
| *nWidth* | The new width of the client area of the window. |
| *nHeight* | The new height of the client area of the window. |
| *rxRatio* | Horizontal scaling ratio. |
| *ryRatio* | Vertical scaling ratio. |

# <a name="AfxSetWindowIcon"></a>AfxSetWindowIcon

Associates a new large icon with a window. The system displays the large icon in the ALT+TAB dialog box, and the small icon in the window caption.

```
FUNCTION AfxSetWindowIcon (BYVAL hwnd AS HWND, BYVAL nIconType AS LONG, BYVAL hIcon AS HICON) AS HICON
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |
| *nIconType* | The type of icon to be set. This parameter can be one of the following values.<br>**ICON_BIG** : Set the large icon for the window.<br>**ICON_SMALL** : Set the small icon for the window. |
| *hIcon* | A handle to the new large icon. If this parameter is NULL, the icon is removed. |

#### Return value

The return value is a handle to the previous large or small icon, depending on the value of *nIconType*. It is NULL if the window previously had no icon of the type indicated by *nIconType*.

# <a name="AfxSetWindowLocation"></a>AfxSetWindowLocation

Sets the location of the top left corner of the window, in pixels.The location is relative to the upper-left corner of the client area in the parent window.

```
FUNCTION AfxSetWindowLocation (BYVAL hwnd AS HWND, BYVAL nLeft AS LONG, BYVAL nTop AS LONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |
| *nLeft* | The new position of the left side of the window, in client coordinates. |
| *nTop* | The new position of the top side of the window, in client coordinates. |

#### Return value

If the function succeeds, the return value is TRUE.

If the function fails, the return value is FALSE. To get extended error information, call **GetLastError**.

# <a name="AfxSetWindowSize"></a>AfxSetWindowSize

Sets the size of the specified window, in pixels.

```
FUNCTION AfxSetWindowSize (BYVAL hwnd AS HWND, BYVAL nWidth AS LONG, BYVAL nHeight AS LONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |
| *nWidth* | The new width of the window. |
| *nHeight* | The new height of the window. |

# <a name="AfxSetWindowText"></a>AfxSetWindowText

Sets the text of a window. This function can also be used to set the text of buttons, edit and static controls.

```
FUNCTION AfxSetWindowText (BYVAL hwnd AS HWND, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |
| *pwszText* | The text to set. |

#### Return value

If the function succeeds, the return value is TRUE.

If the function fails, the return value is FALSE.

# <a name="AfxShowWindowState"></a>AfxShowWindowState

Sets the specified window's show state.

```
FUNCTION AfxShowWindowState (BYVAL hwnd AS HWND, BYVAL nShowState AS LONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |
| *nShowState* | Controls how the window is to be shown. This parameter is ignored the first time an application calls **AfxShowWindowState**, if the program that launched the application provides a **STARTUPINFO** structure. Otherwise, the first time **AfxShowWindowState** is called, the value should be the value obtained by the **WinMain** function in its *nCmdShow* parameter. In subsequent calls, this parameter can be one of the values listed below. |

| Value      | Maning |
| ---------- | ----------- |
| SW_FORCEMINIMIZE | Minimizes a window, even if the thread that owns the window is not responding. This flag should only be used when minimizing windows from a different thread. |
| SW_HIDE | Hides the window and activates another window. |
| SW_MAXIMIZE | Maximizes the specified window. |
| SW_MINIMIZE | Minimizes the specified window and activates the next top-level window in the Z order. |
| SW_RESTORE | Activates and displays the window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when restoring a minimized window. |
| SW_SHOW | Activates the window and displays it in its current size and position. |
| SW_SHOWDEFAULT | Sets the show state based on the SW_ value specified in the **STARTUPINFO** structure passed to the **CreateProcess** function by the program that started the application. |
| SW_SHOWMAXIMIZED | Activates the window and displays it as a maximized window. |
| SW_SHOWMINIMIZED | Activates the window and displays it as a minimized window. |
| SW_SHOWMINNOACTIVE | Displays the window as a minimized window. This value is similar to SW_SHOWMINIMIZED, except the window is not activated. |
| SW_SHOWNA | Displays the window in its current size and position. This value is similar to SW_SHOW, except that the window is not activated. |
| SW_SHOWNOACTIVATE | Displays a window in its most recent size and position. This value is similar to SW_SHOWNORMAL, except that the window is not activated. |
| SW_SHOWNORMAL | Activates and displays a window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when displaying the window for the first time. |

#### Return value

If the window was previously visible, the return value is TRUE.

If the window was previously hidden, the return value is FALSE.

#### Remarks

To perform certain special effects when showing or hiding a window, use **AnimateWindow**.

The first time an application calls **AfxShowWindowState**, it should use the **WinMain** function's *nCmdShow* parameter as its *nCmdShow* parameter. Subsequent calls to **AfxShowWindowState** must use one of the values in the given list, instead of the one specified by the **WinMain** function's *nCmdShow* parameter.

As noted in the discussion of the *nCmdShow* parameter, the *nCmdShow* value is ignored in the first call to **AfxShowWindowState** if the program that launched the application specifies startup information in the structure. In this case, **AfxShowWindowState** uses the information specified in the **STARTUPINFO** structure to show the window. On subsequent calls, the application must call **AfxShowWindowState** with *nCmdShow* set to SW_SHOWDEFAULT to use the startup information provided by the program that launched the application. This behavior is designed for the following situations:

* Applications create their main window by calling **CreateWindow** with the WS_VISIBLE flag set.
* Applications create their main window by calling **CreateWindow** with the WS_VISIBLE flag cleared, and later call **AfxShowWindowState** with the SW_SHOW flag set to make it visible.
