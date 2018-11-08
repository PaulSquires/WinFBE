# CMemBmp Class

The **CMemBmp** class implements a memory bitmap.

You can create an empty bitmap of the specified width an height, e.g.

```
DIM pMemBmp AS CMemBmp = CMemBmp(500, 400)
```

or loading an image

```
DIM pMemBmp AS CMemBmp = CMemBmp($"C:\Users\Pepe\Pictures\MyPicture.jpg")
```

You can manipulate its contents using GDI+ or GDI, e.g.

```
Rectangle pMemBmp.GetMemDC, 10, 10, 150, 150
LineTo pMemBmp.GetMemDC, 30, 180
```

And you can display it in a Graphic Control, e.g.

```
pGraphCtx.DrawBitmap pMemBmp
```

The bitmap can be saved to a file with

```
SaveBitmap
SaveBitmapAsBmp
SaveBitmapAsJpeg
SaveBitmapAsPng
SaveBitmapAsGif
SaveBitmapAsTiff
```

Finally, the **PrintBitmap** method prints the bitmap in the default printer.

The bitmap created by **CMemBmp** is a DDB (Device-Dependent Bitmap), compatible with the device that is associated with the specified device context (hDC parameter of the default constructor). If no device context is specified or the memory bitmap is constructed using the other two additional constructors (the one that use an image from a disk file or a resource file), the created memory bitmap will be compatible with the device context of the screen.

To create a DIB (device-independent bitmap) from the memory bitmap, use the function **GetDIBits**.

To create a GDI+ bitmap from the memory bitmap, use the function **GdipCreateBitmapFromHBITMAP**.

As the **CGraphCtx** graphic control uses DIBs, to draw a **CMemBmp** memory bitmap into it call the method **DrawBitmap** of the **CGraphCtx** class.

# Constructors

| Name       | Description |
| ---------- | ----------- |
| [Constructor(Width,Height)](#Constructor1) | Creates an empty memory bitmap of the specified size. |
| [Constructor(FileName)](#Constructor2) | Creates a memory bitmap of the same size that the specified image to load. |
| [Constructor(ResourceFile)](#Constructor3) | Creates a memory bitmap of the same size that the specified image to load from a resource file. |

# Methods

| Name       | Description |
| ---------- | ----------- |
| [DrawBitmap](#DrawBitmap) | Draws a bitmap in the memory bitmap. |
| [GetBitsPixel](#GetBitsPixel) | Returns the number of bits required to indicate the color of a pixel. |
| [GethBmp](#GethBmp) | Returns the handle of the compatible memory bitmap. |
| [GetHeight](#GetHeight) | Returns the height of the memory bitmap in pixels. |
| [GetMemDC](#GetMemDC) | Returns the handle of the memory device context of the bitmap. |
| [GetPixel](#GetPixel) | Retrieves the red, green, blue (RGB) color value of the pixel at the specified coordinates. |
| [GetPlanes](#GetPlanes) | Returns the count of color planes of the memory bitmap. |
| [GetWidth](#GetWidth) | Returns the width of the memory bitmap in pixels. |
| [GetWidthBytes](#GetWidthBytes) | Returns the number of bytes in each scan line of the memory bitmap. |
| [PrintBtmap](#PrintBtmap) | Prints the bitmap in the default printer. |
| [SaveBitmap](#SaveBitmap) | Saves the bitmap to a file. |
| [SetPixel](#SetPixel) | Sets the pixel at the specified coordinates to the specified color. |

# <a name="Constructor1"></a>Constructor(Width, Height)

Creates an empty memory bitmap of the specified size.

```
CONSTRUCTOR CMemBmp (BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
   BYVAL hdc AS .HDC = NULL, BYVAL clrBkg AS COLORREF = BGR(255, 255, 255))
```

| Parameter  | Description |
| ---------- | ----------- |
| *nWidth* | Optional. The width of the bitmap. |
| *nHeight* | Optional. The height of the bitmap. |
| *hdc* | Optional. A handle to a device context. If no device context is provided, the one of the screen will be used. |
| *clrBkg* | Optional. The background color of the bitmap. Default is white. |

# <a name="Constructor2"></a>Constructor(File Name)

Creates a memory bitmap of the same size that the specified image to load. It also allows to convert the image to gray scale and/or dim the image.

```
CONSTRUCTOR CMemBmp (BYREF wszFileName AS WSTRING, BYVAL dimPercent AS LONG = 0, _
   BYVAL bGrayScale AS LONG = FALSE, BYVAL clrBackground AS ARGB = 0)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileName* | Absolute path to the file. |
| *dimPercent* | Optional. The height of the bitmap. |
| *bGrayScale* | Optional. A handle to a device context. If no device context is provided, the one of the screen will be used. |
| *clrBackground* | Optional. The background color of the bitmap. Default is white. |

# <a name="Constructor3"></a>Constructor(Resource File)

Creates a memory bitmap of the same size that the specified image to load from a resource file. It also allows to convert the image to gray scale and/or dim the image.

```
CONSTRUCTOR CMemBmp (BYVAL hInst AS HINSTANCE, BYREF wszImageName AS WSTRING, BYVAL dimPercent AS LONG = 0, _
   BYVAL bGrayScale AS LONG = FALSE, BYVAL clrBackground AS ARGB = 0) AS HANDLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *hInst* | A handle to the module whose portable executable file or an accompanying MUI file contains the resource. If this parameter is NULL, the function searches the module used to create the current process. |
| *wszImageName* | Name of the image in the resource file (.RES). If the image resource uses an integral identifier, wszImage should begin with a number symbol (#) followed by the identifier in an ASCII format, e.g., "#998". Otherwise, use the text identifier name for the image. Only images embedded as raw data (type RCDATA) are valid. These must be icons in format .png, .jpg, .gif, .tiff. |
| *dimPercent* | Optional. The height of the bitmap. |
| *bGrayScale* | Optional. A handle to a device context. If no device context is provided, the one of the screen will be used. |
| *clrBackground* | Optional. The background color of the bitmap. Default is white. |

# <a name="DrawBitmap"></a>DrawBitmap

Draws a bitmap in the memory bitmap.

```
SUB DrawBitmap (BYVAL hbmp AS HBITMAP, BYVAL x AS SINGLE = 0, BYVAL y AS SINGLE = 0, _
   BYVAL nRight AS SINGLE = 0, BYVAL nBottom AS SINGLE = 0) AS GpStatus
```
```
SUB DrawBitmap (BYVAL pBitmap AS GpBitmap PTR, BYVAL x AS SINGLE = 0, BYVAL y AS SINGLE = 0, _
   BYVAL nRight AS SINGLE = 0, BYVAL nBottom AS SINGLE = 0) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *hbmp* | A handle to the bitmap. |
| *pBitmap* | A pointer to a GDI+ bitmap |
| *x* | Optional. Left position. |
| *y* | Optional. Top position. |
| *nRight* | Optional. Right position. |
| *nBottom* | Optional. Bottom position. |

#### Return value

Returns OK (0) on success or a GdiPlus status code on failure.

# <a name="GetBitsPixel"></a>GetBitsPixel

Returns the number of bits required to indicate the color of a pixel.

```
FUNCTION GetBitsPixel () AS LONG
```

# <a name="GethBmp"></a>GethBmp

Returns the handle of the compatible memory bitmap.

```
FUNCTION GethBmp () AS HBITMAP
```

# <a name="GetHeight"></a>GetHeight

Returns the height of the memory bitmap in pixels.

```
FUNCTION GetHeight () AS LONG
```

# <a name="GetMemDC"></a>GetMemDC

Returns the handle of the memory device context of the bitmap.

```
FUNCTION GetMemDC () AS HDC
```

# <a name="GetPixel"></a>GetPixel

Retrieves the red, green, blue (RGB) color value of the pixel at the specified coordinates.

```
FUNCTION GetPixel (BYVAL x AS LONG, BYVAL y AS LONG) AS COLORREF
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | The x-coordinate, in logical units, of the pixel to be examined. |
| *y* | The y-coordinate, in logical units, of the pixel to be examined. |

#### Return value

The return value is the COLORREF value that specifies the RGB of the pixel. If the pixel is outside of the current clipping region, the return value is CLR_INVALID (hFFFFFFFF defined in Wingdi.bi).

#### Remarks

The function fails if the pixel coordinates lie outside of the current clipping region.

# <a name="GetPlanes"></a>GetPlanes

Returns the count of color planes of the memory bitmap.

```
FUNCTION GetPlanes () AS LONG
```

# <a name="GetWidth"></a>GetWidth

Returns the width of the memory bitmap in pixels.

```
FUNCTION GetWidth () AS LONG
```

# <a name="GetWidthBytes"></a>GetWidthBytes

Returns the number of bytes in each scan line of the memory bitmap.

```
FUNCTION GetWidthBytes () AS LONG
```

# <a name="PrintBtmap"></a>PrintBtmap

Prints the bitmap in the default printer.

```
FUNCTION PrintBtmap (BYVAL bStretch AS BOOLEAN = FALSE, _
   BYVAL nStretchMode AS LONG = InterpolationModeHighQualityBicubic) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *bStretch* | Optional. Stretch the image (TRUE or FALSE). |
| *nStretchMode* | Stretching mode.<br>InterpolationModeLowQuality = 1<br>InterpolationModeHighQuality = 2<br>InterpolationModeBilinear = 3<br>InterpolationModeBicubic = 4<br>InterpolationModeNearestNeighbor = 5<br>InterpolationModeHighQualityBilinear = 6<br>InterpolationModeHighQualityBicubic = 7<br>Default value = InterpolationModeHighQualityBicubic.|

#### Return value

Returns TRUE if the bitmap has been printed successfully, or FALSE otherwise.

# <a name="SaveBitmap"></a>SaveBitmap

Saves the bitmap to a file.

```
FUNCTION SaveBitmap (BYREF wszFileName AS WSTRING, BYREF wszMimeType AS WSTRING) AS LONG
```
```
FUNCTION SaveBitmapAsBmp (BYREF wszFileName AS WSTRING) AS LONG
FUNCTION SaveBitmapAsJpeg (BYREF wszFileName AS WSTRING) AS LONG
FUNCTION SaveBitmapAsPng (BYREF wszFileName AS WSTRING) AS LONG
FUNCTION SaveBitmapAsGif (BYREF wszFileName AS WSTRING) AS LONG
FUNCTION SaveBimapAsTiff (BYREF wszFileName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileName* | Absolute path name of the file. |
| *wszMimeType* | The mime type.<br>"image/bmp" = Bitmap (.bmp)<br>"image/gif" = GIF (.gif)<br>"image/jpeg" = JPEG (.jpg)<br>"image/png" = PNG (.png)<br>"image/tiff" = TIFF (.tiff) |


# <a name="SetPixel"></a>SetPixel

Sets the pixel at the specified coordinates to the specified color.

```
FUNCTION SetPixel (BYVAL x AS LONG, BYVAL y AS LONG, BYVAL crColor AS COLORREF) AS COLORREF
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | The x-coordinate, in logical units, of the point to be set. |
| *y* | The y-coordinate, in logical units, of the point to be set. |
| *crColor* | The color to be used to paint the point. To create a COLORREF color value, use the BGR macro. |

#### Return value

If the function succeeds, the return value is the RGB value that the function sets the pixel to. This value may differ from the color specified by **crColor**; that occurs when an exact match for the specified color cannot be found. If the function fails, the return value is CLR_INVALID (hFFFFFFFF defined in Wingdi.bi).

#### Remarks

The function fails if the pixel coordinates lie outside of the current clipping region.
