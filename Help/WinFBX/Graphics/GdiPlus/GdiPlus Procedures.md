# GdiPlus Procedures

Assorted GDI+ helper procedures.

**Include File**: AfxGdiPlus.inc.

| Name       | Description |
| ---------- | ----------- |
| [AfxGdipAddIconFromFile](#AfxGdipAddIconFromFile) | Loads an image from a file, converts it to an icon and adds it to specified image list. |
| [AfxGdipAddIconFromRes](#AfxGdipAddIconFromRes) | Loads an image from a resource file, converts it to an icon and adds it to specified image list. |
| [AfxGdipBitmapFromBuffer](#AfxGdipBitmapFromBuffer) | Converts an image stored in a buffer into a bitmap and returns the handle. |
| [AfxGdipBitmapFromFile](#AfxGdipBitmapFromFile) | Loads an image from a file, converts it to a bitmap and returns the handle. |
| [AfxGdipBitmapFromRes](#AfxGdipBitmapFromRes) | Loads an image from a resource, converts it to a bitmap and returns the handle. |
| [AfxGdipDllVersion](#AfxGdipDllVersion) | Returns the version of Gdiplus.dll, e.g. 601 for version 6.01. |
| [AfxGdipGetEncoderClsid](#AfxGdipGetEncoderClsid) | Retrieves the encoder's clsid. |
| [AfxGdipGetImageSizeFromFile](#AfxGdipGetImageSizeFromFile) | Returns the size of the image. |
| [AfxGdipIconFromBuffer](#AfxGdipIconFromBuffer) | Converts an image stored in a buffer into an icon and returns the handle. |
| [AfxGdipIconFromFile](#AfxGdipIconFromFile) | Loads an image from a file, converts it to an icon and returns the handle. |
| [AfxGdipIconFromRes](#AfxGdipIconFromRes) | Loads an image from a resource, converts it to an icon and returns the handle. |
| [AfxGdipImageFromBuffer](#AfxGdipImageFromBuffer) | Converts an image stored in a buffer into an icon or bitmap and returns the handle. |
| [AfxGdipImageFromFile](#AfxGdipImageFromFile) | Loads an image from a file, converts it to an icon or bitmap and returns the handle. |
| [AfxGdipImageFromFile2](#AfxGdipImageFromFile2) | Loads an image from a file using GDI+, converts it to an icon or bitmap and returns the handle. |
| [AfxGdipImageFromRes](#AfxGdipImageFromRes) | Loads an image from a resource, converts it to an icon or bitmap and returns the handle. |
| [AfxGdipInit](#AfxGdipInit) | Initializes GDI+. |
| [AfxGdipLoadTexture](#AfxGdipLoadTexture) | Loads an image from disk or a resource an converts it to a texture for use with OpenGL. |
| [AfxGdipPrintHBITMAP](#AfxGdipPrintHBITMAP) | Prints a Windows bitmap in the default printer. |
| [AfxGdipSaveHBITMAPToFile](#AfxGdipSaveHBITMAPToFile) | Saves a Windows bitmap to file. |
| [AfxGdipSaveImageToFile](#AfxGdipSaveImageToFile) | Saves a GDI+ image to file. |
| [AfxGdipShutdown](#AfxGdipShutdown) | Cleans up resources used by Windows GDI+. Each call to **GdiplusStartup** should be paired with a call to **GdiplusShutdown**. |
| [GDIP_ARGB](#GDIP_ARGB) | Returns an ARGB color value initialized with the specified values for the alpha, red, green, and blue components. |
| [GDIP_BGRA](#GDIP_BGRA) | Returns a BGRA color value initialized with the specified values for the blue, green, red and alpha components. |
| [GDIP_COLOR](#GDIP_COLOR) | Returns an ARGB color value initialized with the specified values for the alpha, red, green, and blue components. |
| [GDIP_GetAlpha](#GDIP_GetAlpha) | Returns the alpha component of an ARGB color value. |
| [GDIP_GetBlue](#GDIP_GetBlue) | Returns the blue component of an ARGB color value. |
| [GDIP_GetGreen](#GDIP_GetGreen) | Returns the green component of an ARGB color value. |
| [GDIP_GetRed](#GDIP_GetRed) | Returns the red component of an ARGB color value. |
| [GDIP_POINT](#GDIP_POINT) | Returns a GpPoint color value initialized with the specified values for the x and y coordinates. |
| [GDIP_POINTF](#GDIP_POINTF) | Returns a GpPointF color value initialized with the specified values for the x and y coordinates. |
| [GDIP_RGBA](#GDIP_RGBA) | Returns a RGBA color value initialized with the specified values for the red, green, blue and alpha components. |
| [GDIP_RECT](#GDIP_RECT) | Returns a GpRect structure initialized with the specified values for the x, y, width, and height components. |
| [GDIP_RECTF](#GDIP_RECTF) | Returns a GpRectF structrure initialized with the specified values for the x, y, width, and height components. |
| [GDIP_XRGB](#GDIP_XRGB) | Returns a XRGB color value initialized with the specified values for the red, green, and blue components. |

# <a name="AfxGdipAddIconFromFile"></a>AfxGdipAddIconFromFile

Loads an image from a file, converts it to an icon and adds it to specified image list.

```
FUNCTION AfxGdipAddIconFromFile (BYVAL hIml AS HIMAGELIST, BYREF wszFileName AS WSTRING, _
   BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *hIml* | Handle to the image list. |
| *wszFileName* | Path of the image to load and convert. |
| *dimPercent* | Optional. Percent of dimming (1-99). |
| *bGrayScale* | Optional. TRUE or FALSE. Convert to gray scale. |

#### Return value

Returns the index of the image if successful, or -1 otherwise.

# <a name="AfxGdipAddIconFromRes"></a>AfxGdipAddIconFromRes

Loads an image from a resource file, converts it to an icon and adds it to specified image list.

```
FUNCTION AfxGdipAddIconFromRes (BYVAL hIml AS HIMAGELIST, BYVAL hInstance AS HINSTANCE, _
   BYREF wszFileName AS WSTRING, BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *hIml* | Handle to the image list. |
| *hInstance* | A handle to the module whose portable executable file or an accompanying MUI file contains the resource. If this parameter is NULL, the function searches the module used to create the current process. |
| *wszFileName* | Path of the image to load and convert. |
| *dimPercent* | Optional. Percent of dimming (1-99). |
| *bGrayScale* | Optional. TRUE or FALSE. Convert to gray scale. |

#### Return value

Returns the index of the image if successful, or -1 otherwise.

# <a name="AfxGdipBitmapFromBuffer"></a>AfxGdipBitmapFromBuffer

Converts an image stored in a buffer into a bitmap and returns the handle.

```
FUNCTION AfxGdipBitmapFromBuffer (BYVAL pBuffer AS ANY PTR, BYVAL bufferSize AS SIZE_T_, _
   BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE, BYVAL clrBkg AS ARGB = 0) AS HANDLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *pBuffer* | Pointer to the buffer. |
| *bufferSize* | Size of the buffer |
| *dimPercent* | Optional. Percent of dimming (1-99). |
| *bGrayScale* | Optional. TRUE or FALSE. Convert to gray scale. |
| *clrBkg* | Optional. The background color. This parameter is ignored if the image is totally opaque. |

#### Return value

If the function succeeds, the return value is the handle of the created icon or bitmap.

If the function fails, the return value is NULL.

# <a name="AfxGdipBitmapFromFile"></a>AfxGdipBitmapFromFile

Loads an image from a file, converts it to a bitmap and returns the handle.

```
FUNCTION AfxGdipBitmapFromFile (BYREF wszFileName AS WSTRING, BYVAL dimPercent AS LONG = 0, _
   BYVAL bGrayScale AS LONG = FALSE, BYVAL clrBkg AS ARGB = 0) AS HANDLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileName* | Path of the image to load and convert. |
| *dimPercent* | Optional. Percent of dimming (1-99). |
| *bGrayScale* | Optional. TRUE or FALSE. Convert to gray scale. |
| *clrBkg* | Optional. The background color. This parameter is ignored if the image is totally opaque. |

#### Return value

If the function succeeds, the return value is the handle of the created bitmap.

If the function fails, the return value is NULL.

# <a name="AfxGdipBitmapFromRes"></a>AfxGdipBitmapFromRes

Loads an image from a resource, converts it to a bitmap and returns the handle.

```
FUNCTION AfxGdipBitmapFromRes (BYVAL hInstance AS HINSTANCE, BYREF wszImageName AS WSTRING, _
   BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE, BYVAL clrBkg AS ARGB = 0) AS HANDLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *hInstance* | A handle to the module whose portable executable file or an accompanying MUI file contains the resource. If this parameter is NULL, the function searches the module used to create the current process. |
| *wszFileName* | Name of the image in the resource file (.RES). If the image resource uses an integral identifier, wszImage should begin with a number symbol (#) followed by the identifier in an ASCII format, e.g., "#998". Otherwise, use the text identifier name for the image. Only images embedded as raw data (type RCDATA) are valid. These must be icons in format .png, .jpg, .gif, .tiff. |
| *dimPercent* | Optional. Percent of dimming (1-99). |
| *bGrayScale* | Optional. TRUE or FALSE. Convert to gray scale. |
| *clrBkg* | Optional. The background color. This parameter is ignored if the image is totally opaque. |

#### Return value

If the function succeeds, the return value is the handle of the created bitmap.

If the function fails, the return value is NULL.

# <a name="AfxGdipDllVersion"></a>AfxGdipDllVersion

Returns the version of Gdiplus.dll, e.g. 601 for version 6.01.

```
FUNCTION AfxGdipDllVersion () AS LONG
```

# <a name="AfxGdipGetEncoderClsid"></a>AfxGdipGetEncoderClsid

Retrieves the encoder's clsid.

```
FUNCTION AfxGdipGetEncoderClsid (BYREF wszMimeType AS WSTRING) AS GUID
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMimeType* | The mime type. |

# <a name="AfxGdipGetImageSizeFromFile"></a>AfxGdipGetImageSizeFromFile

Returns the size of the image.

```
FUNCTION AfxGdipGetImageSizeFromFile (BYREF wszFileName AS WSTRING, BYVAL nWidth AS DWORD PTR, _
   BYVAL nHeight AS DWORD PTR) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileName* | The filename path. |
| *nWidth* | Pointer to a DWORD variable that will receive the width, in pixels, of the image. |
| *nHeight* | Pointer to a DWORD variable that will receive the height, in pixels, of the image. |

#### Return value

If the function succeeds, it returns Ok, which is an element of the Status enumeration.

If the function fails, it returns one of the other elements of the Status enumeration.

# <a name="AfxGdipIconFromBuffer"></a>AfxGdipIconFromBuffer

Converts an image stored in a buffer into an icon and returns the handle.

```
FUNCTION AfxGdipIconFromBuffer (BYVAL pBuffer AS ANY PTR, BYVAL bufferSize AS SIZE_T_, _
   BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE) AS HANDLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *pBuffer* | Pointer to the buffer. |
| *bufferSize* | Size of the buffer |
| *dimPercent* | Optional. Percent of dimming (1-99). |
| *bGrayScale* | Optional. TRUE or FALSE. Convert to gray scale. |

#### Return value

If the function succeeds, the return value is the handle of the created icon.

If the function fails, the return value is NULL.

# <a name="AfxGdipIconFromFile"></a>AfxGdipIconFromFile

Loads an image from a file, converts it to an icon and returns the handle.

```
FUNCTION AfxGdipIconFromFile (BYREF wszFileName AS WSTRING, BYVAL dimPercent AS LONG = 0, _
   BYVAL bGrayScale AS LONG = FALSE) AS HANDLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileName* | Path of the image to load and convert. |
| *dimPercent* | Optional. Percent of dimming (1-99). |
| *bGrayScale* | Optional. TRUE or FALSE. Convert to gray scale. |

#### Return value

If the function succeeds, the return value is the handle of the created icon.

If the function fails, the return value is NULL.

# <a name="AfxGdipIconFromRes"></a>AfxGdipIconFromRes

Loads an image from a resource, converts it to an icon and returns the handle.

```
FUNCTION AfxGdipIconFromRes (BYVAL hInstance AS HINSTANCE, BYREF wszImageName AS WSTRING, _
   BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE) AS HANDLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *hInstance* | A handle to the module whose portable executable file or an accompanying MUI file contains the resource. If this parameter is NULL, the function searches the module used to create the current process. |
| *wszImageName* | Name of the image in the resource file (.RES). If the image resource uses an integral identifier, wszImage should begin with a number symbol (#) followed by the identifier in an ASCII format, e.g., "#998". Otherwise, use the text identifier name for the image. Only images embedded as raw data (type RCDATA) are valid. These must be icons in format .png, .jpg, .gif, .tiff. |
| *dimPercent* | Optional. Percent of dimming (1-99). |
| *bGrayScale* | Optional. TRUE or FALSE. Convert to gray scale. |

#### Return value

If the function succeeds, the return value is the handle of the created icon.

If the function fails, the return value is NULL.

# <a name="AfxGdipImageFromBuffer"></a>AfxGdipImageFromBuffer

Converts an image stored in a buffer into an icon or bitmap and returns the handle.

```
FUNCTION AfxGdipImageFromBuffer (BYVAL pBuffer AS ANY PTR, BYVAL bufferSize AS SIZE_T_, _
   BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE, _
   BYVAL imageType AS LONG = IMAGE_ICON, BYVAL clrBkg AS ARGB = 0) AS HANDLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *pBuffer* | Pointer to the buffer. |
| *bufferSize* | Size of the buffer |
| *dimPercent* | Optional. Percent of dimming (1-99). |
| *bGrayScale* | Optional. TRUE or FALSE. Convert to gray scale. |
| *imageType* | Optional. IMAGE_ICON or IMAGE_BITMAP. Default value: IMAGE_ICON. |
| *clrBkg* | Optional. The background color. This parameter is ignored if the image type is IMAGE_ICON or the bitmap is totally opaque. |

#### Return value

If the function succeeds, the return value is the handle of the created icon or bitmap.

If the function fails, the return value is NULL.

#### Usage example

```
DIM wszFileName AS WSTRING * MAX_PATH
wszFileName = ExePath & "\arrow_left_256.png"
DIM bufferSize AS SIZE_T_
DIM nFile AS LONG
nFile = FREEFILE
OPEN wszFileName FOR BINARY AS nFile
IF ERR THEN EXIT FUNCTION
bufferSize = LOF(nFile)
DIM pBuffer AS UBYTE PTR
pBuffer = CAllocate(1, bufferSize)
GET #nFile, , *pBuffer, bufferSize
CLOSE nFile
IF pBuffer THEN
   ImageList_ReplaceIcon(hImageList, -1, AfxGdipIconFromBuffer(pBuffer, ImageSize))
   DeAllocate(pBuffer)
END IF
```

# <a name="AfxGdipImageFromFile"></a>AfxGdipImageFromFile

Loads an image from a file, converts it to an icon or bitmap and returns the handle.

```
FUNCTION AfxGdipImageFromFile (BYREF wszFileName AS WSTRING, BYVAL dimPercent AS LONG = 0, _
   BYVAL bGrayScale AS LONG = FALSE, BYVAL imageType AS LONG = IMAGE_ICON, _
   BYVAL clrBkg AS ARGB = 0) AS HANDLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileName* | Path of the image to load and convert. |
| *dimPercent* | Optional. Percent of dimming (1-99). |
| *bGrayScale* | Optional. TRUE or FALSE. Convert to gray scale. |
| *imageType* | Optional. IMAGE_ICON or IMAGE_BITMAP. Default value: IMAGE_ICON. |
| *clrBkg* | Optional. The background color. This parameter is ignored if the image type is IMAGE_ICON or the bitmap is totally opaque. |

#### Return value

If the function succeeds, the return value is the handle of the created icon or bitmap.

If the function fails, the return value is NULL.

# <a name="AfxGdipImageFromFile2"></a>AfxGdipImageFromFile2

Loads an image from a file using GDI+, converts it to an icon or bitmap and returns the handle.

```
FUNCTION AfxGdipImageFromFile2 (BYREF wszFileName AS WSTRING, BYVAL dimPercent AS LONG = 0, _
   BYVAL bGrayScale AS LONG = FALSE, BYVAL imageType AS LONG = IMAGE_ICON, _
   BYVAL clrBkg AS ARGB = 0) AS HANDLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileName* | Path of the image to load and convert. |
| *dimPercent* | Optional. Percent of dimming (1-99). |
| *bGrayScale* | Optional. TRUE or FALSE. Convert to gray scale. |
| *imageType* | Optional. IMAGE_ICON or IMAGE_BITMAP. Default value: IMAGE_ICON. |
| *clrBkg* | Optional. The background color. This parameter is ignored if the image type is IMAGE_ICON or the bitmap is totally opaque. |

#### Return value

If the function succeeds, the return value is the handle of the created icon or bitmap.

If the function fails, the return value is NULL.

#### Remarks

A quirk in the GDI+ **GdipLoadImageFromFile** function causes that dim gray images (often used for disabled icons) are converted to darker shades of gray. Therefore, is better to use **AfxGdipImageFromFile**.

# <a name="AfxGdipImageFromRes"></a>AfxGdipImageFromRes

Loads an image from a resource, converts it to an icon or bitmap and returns the handle.

```
FUNCTION AfxGdipImageFromRes (BYVAL hInstance AS HINSTANCE, BYREF wszImageName AS WSTRING, _
   BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE, _
   BYVAL imageType AS LONG = IMAGE_ICON, BYVAL clrBkg AS ARGB = 0) AS HANDLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *hInstance* | A handle to the module whose portable executable file or an accompanying MUI file contains the resource. If this parameter is NULL, the function searches the module used to create the current process. |
| *wszImageName* | Name of the image in the resource file (.RES). If the image resource uses an integral identifier, wszImage should begin with a number symbol (#) followed by the identifier in an ASCII format, e.g., "#998". Otherwise, use the text identifier name for the image. Only images embedded as raw data (type RCDATA) are valid. These must be icons in format .png, .jpg, .gif, .tiff. |
| *dimPercent* | Optional. Percent of dimming (1-99). |
| *bGrayScale* | Optional. TRUE or FALSE. Convert to gray scale. |
| *imageType* | Optional. IMAGE_ICON or IMAGE_BITMAP. Default value: IMAGE_ICON. |
| *clrBkg* | Optional. The background color. This parameter is ignored if the image type is IMAGE_ICON or the bitmap is totally opaque. |

#### Return value

If the function succeeds, the return value is the handle of the created icon or bitmap.

If the function fails, the return value is NULL.

# <a name="AfxGdipInit"></a>AfxGdipInit

Initializes GDI+.

```
FUNCTION AfxGdipInit (BYVAL version AS UINT32 = 1) AS ULONG_PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *version* | Optional. The GDI+ version number. Must be 1. |

#### Return value

Returns a token on success or 0 in failure. The returned token will be used in the call to GdiplusShutdown when you have finished using GDI+.

# <a name="AfxGdipLoadTexture"></a>AfxGdipLoadTexture

Loads an image from disk or a resource an converts it to a texture for use with OpenGL.

```
FUNCTION AfxGdipLoadTexture (BYREF wszFileName AS WSTRING, BYREF TextureWidth AS LONG, _
   BYREF TextureHeight AS LONG, BYREF strTextureData AS STRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileName* | The path of the image. |
| *TextureWidth* | Out. Width of the texture. |
| *TextureHeight* | Out. Height of the texture. |
| *strTextureData* | Out. The texture data. |

```
PRIVATE FUNCTION AfxGdipLoadTexture (BYVAL hInstance AS HINSTANCE, BYREF wszResourceName AS WSTRING, _
   BYREF TextureWidth AS LONG, BYREF TextureHeight AS LONG, BYREF strTextureData AS STRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *hInstance* | The instance handle. |
| *wszResourceName* | The name of the resource |
| *TextureWidth* | Out. Width of the texture. |
| *TextureHeight* | Out. Height of the texture. |
| *strTextureData* | Out. The texture data. |

#### Return value

ERROR_FILE_NOT_FOUND = File not found.<br>
ERROR_INVALID_DATA = Bad image size..<br>
A GdiPlus status value.

# <a name="AfxGdipPrintHBITMAP"></a>AfxGdipPrintHBITMAP

Prints a Windows bitmap in the default printer.

```
FUNCTION AfxGdipPrintHBITMAP (BYVAL hbmp AS HBITMAP, BYVAL bStretch AS BOOLEAN = FALSE, _
   BYVAL nStretchMode AS LONG = InterpolationModeHighQualityBicubic) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *hbmp* | Handle to a Windows bitmap. |
| *bStretch* | Optional. Stretch the image. |
| *nStretchMode* | Optional. Stretching mode.<br>InterpolationModeLowQuality = 1<br>InterpolationModeHighQuality = 2<br>InterpolationModeBilinear = 3<br>InterpolationModeBicubic = 4<br>InterpolationModeNearestNeighbor = 5<br>InterpolationModeHighQualityBilinear = 6<br>InterpolationModeHighQualityBicubic = 7<br>Default value = InterpolationModeHighQualityBicubic. |

#### Return value

Returns TRUE if the bitmap has been printed successfully, or FALSE otherwise.

# <a name="AfxGdipSaveHBITMAPToFile"></a>AfxGdipSaveHBITMAPToFile

Saves a Windows bitmap to file.

```
FUNCTION AfxGdipSaveHBITMAPToFile (BYVAL hbmp AS HBITMAP, BYREF wszFileName AS WSTRING, _
   BYREF wszMimeType AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *hbmp* | Handle to a Windows bitmap. |
| *wszFileName* | Path name for the image to be saved. |
| *wszMimeType* | The mime type.<br>"image/bmp" = Bitmap (.bmp)<br>"image/gif" = GIF (.gif)<br>"image/jpeg" = JPEG (.jpg)<br>"image/png" = PNG (.png)<br>"image/tiff" = TIFF (.tiff) |

#### Return value

If the method succeeds, it returns Ok, which is an element of the Status enumeration.

If the method fails, it returns one of the other elements of the Status enumeration.

# <a name="AfxGdipSaveImageToFile"></a>AfxGdipSaveImageToFile

Saves a GDI+ image to file.

```
FUNCTION AfxGdipSaveImageToFile (BYVAL pImage AS GpImage PTR, BYREF wszFileName AS WSTRING, _
   BYREF wszMimeType AS WSTRING) AS LONG
```
```
FUNCTION AfxGdipSaveImageToBmp (BYVAL pImage AS GpImage PTR, BYREF wszFileName AS WSTRING) AS LONG
FUNCTION AfxGdipSaveImageToJpeg (BYVAL pImage AS GpImage PTR, BYREF wszFileName AS WSTRING) AS LONG
FUNCTION AfxGdipSaveImageToPng (BYVAL pImage AS GpImage PTR, BYREF wszFileName AS WSTRING) AS LONG
FUNCTION AfxGdipSaveImageToGif (BYVAL pImage AS GpImage PTR, BYREF wszFileName AS WSTRING) AS LONG
FUNCTION AfxGdipSaveImageToTiff (BYVAL pImage AS GpImage PTR, BYREF wszFileName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pImage* | Pointer to the GDI+ image to save. |
| *wszFileName* | Path name for the image to be saved. |
| *wszMimeType* | The mime type.<br>"image/bmp" = Bitmap (.bmp)<br>"image/gif" = GIF (.gif)<br>"image/jpeg" = JPEG (.jpg)<br>"image/png" = PNG (.png)<br>"image/tiff" = TIFF (.tiff) |

#### Return value

If the method succeeds, it returns Ok, which is an element of the Status enumeration.

If the method fails, it returns one of the other elements of the Status enumeration.

# <a name="AfxGdipShutdown"></a>AfxGdipShutdown

Cleans up resources used by Windows GDI+. Each call to **GdiplusStartup** should be paired with a call to **GdiplusShutdown**.

```
SUB AfxGdipShutdown (BYVAL token AS ULONG_PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *token* | Token returned by a previous call to **GdiplusStartup**. |

# <a name="GDIP_ARGB"></a>GDIP_ARGB

Returns an ARGB color value initialized with the specified values for the alpha, red, green, and blue components.

```
FUNCTION GDIP_ARGB (BYVAL a AS UBYTE, BYVAL r AS UBYTE, BYVAL g AS UBYTE, BYVAL b AS UBYTE) AS COLORREF
```

| Parameter  | Description |
| ---------- | ----------- |
| *a* | The alpha component value. |
| *r* | The red component value. |
| *g* | The green component value. |
| *b*| The bue ponent value. |

#### Return value

The ARGB value.

# <a name="GDIP_BGRA"></a>GDIP_BGRA

Returns a BGRA color value initialized with the specified values for the blue, green, red and alpha components.

```
FUNCTION GDIP_BGRA (BYVAL b AS UBYTE, BYVAL g AS UBYTE, BYVAL r AS UBYTE, BYVAL a AS UBYTE) AS COLORREF
```

| Parameter  | Description |
| ---------- | ----------- |
| *b*| The bue ponent value. |
| *g* | The green component value. |
| *r* | The red component value. |
| *a* | The alpha component value. |

#### Return value

The BGRA value.

# <a name="GDIP_COLOR"></a>GDIP_COLOR

Returns an ARGB color value initialized with the specified values for the alpha, red, green, and blue components.

```
FUNCTION GDIP_COLOR (BYVAL a AS UBYTE, BYVAL r AS UBYTE, BYVAL g AS UBYTE, BYVAL b AS UBYTE) AS COLORREF
```

| Parameter  | Description |
| ---------- | ----------- |
| *a* | The alpha component value. |
| *r* | The red component value. |
| *g* | The green component value. |
| *b*| The bue ponent value. |

#### Return value

The ARGB value.

# <a name="GDIP_GetAlpha"></a>GDIP_GetAlpha

Returns the alpha component of an ARGB color value.

```
FUNCTION GDIP_GetAlpha (BYVAL argbcolor AS COLORREF) AS BYTE
```

| Parameter  | Description |
| ---------- | ----------- |
| *argbcolor* | The ARGB color value. |

# <a name="GDIP_GetBlue"></a>GDIP_GetBlue

Returns the blue component of an ARGB color value.

```
FUNCTION GDIP_GetBlue (BYVAL argbcolor AS COLORREF) AS BYTE
```

| Parameter  | Description |
| ---------- | ----------- |
| *argbcolor* | The ARGB color value. |

# <a name="GDIP_GetGreen"></a>GDIP_GetGreen

Returns the green component of an ARGB color value.

```
FUNCTION GDIP_GetGreen (BYVAL argbcolor AS COLORREF) AS BYTE
```

| Parameter  | Description |
| ---------- | ----------- |
| *argbcolor* | The ARGB color value. |

# <a name="GDIP_GetRed"></a>GDIP_GetRed

Returns the red component of an ARGB color value.

```
FUNCTION GDIP_GetRed (BYVAL argbcolor AS COLORREF) AS BYTE
```

| Parameter  | Description |
| ---------- | ----------- |
| *argbcolor* | The ARGB color value. |

# <a name="GDIP_POINT"></a>GDIP_POINT

Returns a GpPoint color value initialized with the specified values for the *x* and *y* coordinates.

```
FUNCTION GDIP_POINT (BYVAL x AS LONG, BYVAL y AS LONG) AS GpPoint
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | The x coordinate. |
| *y* | The y coordinate. |

# <a name="GDIP_POINTF"></a>GDIP_POINTF

Returns a GpPointF color value initialized with the specified values for the *x* and *y* coordinates.

```
FUNCTION GDIP_POINTF (BYVAL x AS SINGLE, BYVAL y AS SINGLE) AS GpPointF
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | The x coordinate. |
| *y* | The y coordinate. |

# <a name="GDIP_RGBA"></a>GDIP_RGBA

Returns a RGBA color value initialized with the specified values for the red, green, blue and alpha components.

```
FUNCTION GDIP_RGBA (BYVAL r AS UBYTE, BYVAL g AS UBYTE, BYVAL b AS UBYTE, BYVAL a AS UBYTE) AS COLORREF
```

| Parameter  | Description |
| ---------- | ----------- |
| *r* | The red component value. |
| *g* | The green component value. |
| *b*| The bue ponent value. |
| *a* | The alpha component value. |

#### Return value

The RGBA value.

# <a name="GDIP_RECT"></a>GDIP_RECT

Returns a GpRect structure initialized with the specified values for the x, y, width, and height components.

```
FUNCTION GDIP_RECT (BYVAL x AS LONG, BYVAL y AS LONG, BYVAL nWidth AS LONG, BYVAL nHeight AS LONG) AS GpRect
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | The x coordinate. |
| *y* | The y coordinate. |
| *nWidth*| The width. |
| *nHeight* | The height. |

#### Return value

The filled GpRect structure.

# <a name="GDIP_RECTF"></a>GDIP_RECTF

Returns a GpRectF structure initialized with the specified values for the x, y, width, and height components.

```
FUNCTION GDIP_RECT (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, _
   BYVAL nHeight AS SINGLE) AS GpRectF
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | The x coordinate. |
| *y* | The y coordinate. |
| *nWidth*| The width. |
| *nHeight* | The height. |

#### Return value

The filled GpRectF structure.

# <a name="GDIP_XRGB"></a>GDIP_XRGB

Returns a XRGB color value initialized with the specified values for the red, green, and blue components.

```
FUNCTION GDIP_XRGB (BYVAL r AS UBYTE, BYVAL g AS UBYTE, BYVAL b AS UBYTE) AS COLORREF
```

| Parameter  | Description |
| ---------- | ----------- |
| *r* | The red component value. |
| *g* | The green component value. |
| *b*| The bue ponent value. |

#### Return value

The XRGB value.

# INFO: Interoperability Between GDI and GDI+

Microsoft article: http://support.microsoft.com/kb/311221/en-us

Article ID: 311221 - Last Review: February 12, 2007 - Revision: 3.4
INFO: Interoperability Between GDI and GDI+
This article was previously published under Q311221

SUMMARY

It is sometimes desirable to mix GDI and GDI+ drawing operations in the same code path. There are some caveats to keep in mind when you are writing code that allows GDI and GDI+ to interoperate. This article outlines those caveats and provides additional information to help you successfully write such code.

MORE INFORMATION

While it is permissible to mix GDI and GDI+ code, there are certain rules that must be followed. Generally, you should not interleave GDI and GDI+ calls on one target object. For example, it is okay to wrap a Graphics object around an HDC, but you should not access the HDC directly from GDI until the Graphics object is destroyed.

Four primary scenarios for interoperability between GDI and GDI+ are covered in this article:

* Using GDI on a GDI+ Graphics object backed by the screen
* Using GDI on a GDI+ Graphics object backed by a bitmap
* Using GDI+ on a GDI HDC
* Using GDI+ on a GDI memory HBITMAP

### Using GDI on a GDI+ Graphics Object Backed by the Screen

One example of a need to use GDI on a GDI+ Graphics object backed by the screen would be to draw a "rubber band" or "focus" rectangle. GDI+ currently has no direct support for raster operations (ROPs), so GDI must be used directly if R2_XOR pen operations are required. In this case, you would use GdipGetDC to obtain an HDC to which the GDI output would be directed. GDI+ output should not be attempted on the Graphics object for the life of the HDC (that is, until GdipReleaseDC is called).

### Using GDI on a GDI+ Graphics Object Backed by a Bitmap

When GdipGetDC is called for a Graphics object that is backed by a bitmap rather than the screen, a memory HDC is created and a new HBITMAP is created and selected into the memory HDC. This new memory bitmap is not initialized with the original bitmap's image but rather with a sentinel pattern, which allows GDI+ to track changes to the bitmap. Any changes that are made to the memory bitmap through the use of GDI code become apparent in changes to the sentinel pattern. When GdipReleaseDC is called, those changes are copied back to the original bitmap. Because the memory bitmap is not initialized with the bitmap's image, an HDC that is obtained in this way should be considered "write only" and is therefore not suitable for use with ROPs, the use of which requires the ability to read the target, like R2_XOR. Also, there is a performance cost to this approach because GDI+ must copy the changes back to the original bitmap.

### Using GDI+ on a GDI HDC

You can facilitate the use of GDI+ on an HDC by using the Graphics constructor that takes an HDC as a parameter. The drawing members of the Graphics class can be used to draw on the HDC in this way. Once the Graphics object is attached to the HDC, no GDI operations should be performed on the HDC until the Graphics object is destroyed or goes out of scope. If GDI output is required on the HDC, either destroy the Graphics object before using the original HDC or use GdipGetDC to get a new HDC and then follow the rules described earlier in this article for interoperability while using GDI on a GDI+ object.

Using GDI+ on a GDI Memory HBITMAP

The GDI+ Bitmap constructor that takes an HBITMAP as a parameter does not use the actual source HBITMAP as the backing image for the bitmap. Rather, a copy of the image is made in the constructor, and changes are not written back to the original bitmap, even during execution of the destructor. The new bitmap can be thought of as "copy on creation," so to get GDI+ to draw on a memory HBITMAP from GDI and have the changes apply to the HBITMAP, an approach like the following is needed instead:

1. Create a DIBSection.
2. Select the DIBSection into a memory HDC.
3. To use GDI+ to draw to the DIBSection, wrap a Graphics object around the HDC.
4. To use GDI to draw to/from the DIBSection, destroy the Graphics object, and use the HDC.
5. Destroy the Graphics objects, and then clear the DIBSection selection from the HDC.

Later, a bitmap can be constructed from the DIBSection and used as a source image in GdipDrawImage if needed.
