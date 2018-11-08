# CGdiPlus

GDI+ is the successor to GDI, the graphics device interface included with earlier versions of Windows. Windows XP or Windows .NET Server supports GDI for compatibility with existing applications, but programmers of new applications should use GDI+ for all their graphics needs because GDI+ optimizes many of the capabilities of GDI and also provides additional features.

A graphics device interface, such as GDI+, allows application programmers to display information on a screen or printer without having to be concerned about the details of a particular display device. The application programmer makes calls to the provided functions and those functions in turn make the appropriate calls to specific device drivers. GDI+ insulates the application from the graphics hardware, and it is this insulation that allows developers to create device-independent applications.

**Include file**: CGdiPlus.inc.

# CGpBase class

Initializes the GDI+ library and keeps a reference count of it. The destructor shutdowns the library. All the other GDI+ classes inherit frrom **CGpBase**.

### GetLastStatus method

Returns a value that indicates the nature of this object's most recent method failure.

```
FUNCTION GetLastStatus () AS GpStatus
```

#### Return value

Returns an element of the **Status** enumeration.

If no methods invoked on this object have failed since the previous call to **GetLastStatus**, then **GetLastStatus** returns **Ok** (0). If at least one method has failed since the previous call to **GetLastStatus**, then **GetLastStatus** returns a value that indicates the nature of the most recent failure.

For a detailed list of status codes see the **Status** enumeration.

# Structures

| Name       | Description |
| ---------- | ----------- |
| [BitmapData](#BitmapData) | Stores attributes of a bitmap. |
| [CharacterRange](#CharacterRange) | Specifies a range of character positions within a string. |
| [ColorMap](#ColorMap) | Contains two COLORREF values. |
| [ColorMatrix](#ColorMatrix) | Contains a 5x5 matrix of real numbers. |
| [ColorPalette](#ColorPalette) | Defines an array of colors that make up a color palette. |
| [EncoderParameter](#EncoderParameter) | Holds a parameter that can be passed to an image encoder. |
| [EncoderParameters](#EncoderParameters) | Holds array of **EncoderParameter** structures. |
| [ENHMETAHEADER3](#ENHMETAHEADER3) | Contains enhanced-metafile data. |
| [GdiplusStartupInput](#GdiplusStartupInput) | Holds a block of arguments that are required by the GdiplusStartup function. |
| [GdiplusStartupOutput](#GdiplusStartupOutput) | Uses the **GdiplusStartupOutput** structure to return (in its output parameter) a pointer to a hook function and a pointer to an unhook function. |
| [GpPathData](#GpPathData) | Used to get or set the data points (and their types) of a path. |
| [GpPoint](#GpPoint) | Defines the x- and y-coordinates of a point. |
| [GpPointF](#GpPointF) | Defines the x- and y-coordinates of a point. |
| [GpRect](#GpRect) | Defines the upper-left corner, width, and height of a rectangle. |
| [GpRectF](#GpRectF) | Defines the upper-left corner, width, and height of a rectangle. |
| [GpSize](#GpSize) | Encapsulates a Width and Height dimension in a 2-D coordinate system. |
| [GpSizeF](#GpSizeF) | Encapsulates a Width and Height dimension in a 2-D coordinate system. |
| [ImageCodecInfo](#ImageCodecInfo) | Stores information about an image codec. |
| [LOGFONTA](#LOGFONTA) | Defines the attributes of a font. |
| [LOGFONTW](#LOGFONTW) | Defines the attributes of a font. |
| [MetafileHeader](#MetafileHeader) | Stores properties of an associated metafile. |
| [PropertyItem](#PropertyItem) | Holds one piece of image metadata. |
| [PWMFRect16](#PWMFRect16) | Defines a rectangle that bounds a placeable metafile. |
| [WmfPlaceableFileHeader](#WmfPlaceableFileHeader) | Defines the fields of a placeable metafile header. |

# Enumerations

| Name       | Description |
| ---------- | ----------- |
| [BrushType](#BrushType) | Indicates the type of brush. |
| [Color](#Color) | Predefined ARGB colors. |
| [ColorAdjustType](#ColorAdjustType) | Specifies which GDI+ objects use color-adjustment information. |
| [ColorChannelFlags](#ColorChannelFlags) | Specifies individual channels in the CMYK (cyan, magenta, yellow, black) color space. |
| [ColorMatrixFlags](#ColorMatrixFlags) | Specifies the types of images and colors that will be affected by the color and grayscale adjustment settings. |
| [ColorMode](#ColorMode) | ColorModeARGB32 = 0, ColorModeARGB64 = 1. |
| [CombineMode](#CombineMode) | Specifies how a new region is combined with an existing region. |
| [CompositingMode](#CompositingMode) | Specifies how rendered colors are combined with background colors. |
| [CompositingQuality](#CompositingQuality) | Specifies whether gamma correction is applied when colors are blended with background colors. |
| [CoordinateSpace](#CoordinateSpace) | Specifies coordinate spaces. |
| [CustomLineCapType](#CustomLineCapType) | CustomLineCapTypeDefault = 0, CustomLineCapTypeAdjustableArrow = 1. |
| [DashCap](#DashCap) | Specifies the type of graphic shape to use on both ends of each dash in a dashed line. |
| [DashStyle](#DashStyle) | Specifies the line style of a line drawn with a Microsoft Windows GDI+ pen. |
| [DriverStringOptions](#DriverStringOptions) | Sècifies the spacing, orientation, and quality of the rendering for driver strings. |
| [EmfPlusRecordType](#EmfPlusRecordType) | Identifies metafile record types used in Microsoft Windows Metafile Format (WMF), Enhanced Metafile (EMF), and EMF+ files. |
| [EmfToWmfBitsFlags](#EmfToWmfBitsFlags) | Specifies options for the **GdipEmfToWmfBits** method, which converts an Enhanced Metafile (EMF) metafile to a Microsoft Windows Metafile Format (WMF) metafile. |
| [EmfType](#EmfType) | Specifies the nature of the records that are placed in an Enhanced Metafile (EMF) file. |
| [EncoderParameterValueType](#EncoderParameterValueType) | Specifies data types for image codec (encoder/decoder) parameters. |
| [EncoderValue](#EncoderValue) | Specifies values that can be passed as arguments to image encoders. |
| [FillMode](#FillMode) | Specifies how to fill areas that are formed when a path or curve intersects itself. |
| [FlushIntention](#FlushIntention) | Specifies when to flush the queue of graphics operations. |
| [FontStyle](#FontStyle) | Specifies the style of the typeface of a font. Styles can be combined. |
| [GenericFontFamily](#GenericFontFamily) | GenericFontFamilySerif = 0, GenericFontFamilySansSerif = 1, GenericFontFamilyMonospace = 2. |
| [HatchStyle](#HatchStyle) | Specifies the hatch pattern used by a brush of type **HatchBrush**. |
| [HistogramFormat](#HistogramFormat) | Specifies the number and type of histograms that represent the color channels of a bitmap. |
| [HotkeyPrefix](#HotkeyPrefix) | Specifies how to display hot keys. |
| [ImageFlags](#ImageFlags) | Specifies the attributes of the pixel data contained in an image. |
| [ImageLockMode](#ImageLockMode) | Specifies flags that are passed to the flags parameter of the **GdipBitmapLockBits** function. |
| [ImageType](#ImageType) | Indicates whether an image is a bitmap or a metafile. |
| [InterpolationMode](#InterpolationMode) | Specifies the algorithm that is used when images are scaled or rotated. |
| [ItemDataPosition](#ItemDataPosition) | Specifies the location of custom metadata in an image file. |
| [LinearGradientMode](#LinearGradientMode) | Specifies the direction in which the change of color occurs for a linear gradient brush. |
| [LineCap](#LineCap) | Specifies the type of graphic shape to use on the end of a line drawn with a Microsoft Windows GDI+ pen. |
| [LineJoin](#LineJoin) | Specifies how to join two lines that are drawn by the same pen and whose ends meet. |
| [MatrixOrder](#MatrixOrder) | Specifies the order of multiplication when a new matrix is multiplied by an existing matrix. |
| [MetafileFrameUnit](#MetafileFrameUnit) | Specifies the unit of measure for a metafile frame rectangle. |
| [MetafileType](#MetafileType) | Specifies types of metafiles. |
| [ObjectType](#ObjectType) | Indicates the object type value of an EMF+ record. |
| [PaletteFlags](#PaletteFlags) | Indicates attributes of the color data in a palette. |
| [PaletteType](#PaletteType) | Identify several standard color palette formats. |
| [PathPointType](#PathPointType) | Indicates point types and flags for the data points in a path. |
| [PenAlignment](#PenAlignment) | Specifies the alignment of a pen relative to the stroke that is being drawn. |
| [PenType](#PenType) | Indicates the type of pattern, texture, or gradient that a pen draws. |
| [PixelOffsetMode](#PixelOffsetMode) | Specifies the pixel offset mode. |
| [QualityMode](#QualityMode) | Specifies the quality mode. |
| [RotateFlipType](#RotateFlipType) | Specifies the direction of an image's rotation and the axis used to flip the image. |
| [SmoothingMode](#SmoothingMode) | Specifies the type of smoothing (antialiasing) that is applied to lines and curves. |
| [Status](#Status) | Indicates the result of a Microsoft Windows GDI+ function call. |
| [StringAlignment](#StringAlignment) | Specifies how a string is aligned in reference to the bounding rectangle. |
| [StringDigitSubstitute](#StringDigitSubstitute) | Specifies how to substitute digits in a string according to a user's locale or language. |
| [StringFormatFlags](#StringFormatFlags) | Specifies text layout information (such as orientation and clipping) and display manipulations (such as ellipsis insertion, digit substitution, and representation of characters that are not supported by a font). |
| [StringTrimming](#StringTrimming) | Specifies how to trim characters from a string so that the string fits into a layout rectangle. |
| [TestControl](#TestControl) | TestControlForceBilinear = 0, TestControlNoICM = 1, TestControlGetBuildNumber = 2. |
| [TextRenderingHint](#TextRenderingHint) | Specifies the process used to render text. |
| [Unit](#Unit) | Specifies the unit of measure for a given data type. |
| [WarpMode](#WarpMode) | Specifies warp modes that can be used to transform images. |
| [WrapMode](#WrapMode) | Specifies how repeated copies of an image are used to tile an area. |

# Constants

| Name       | Description |
| ---------- | ----------- |
| [Image Encoder Constants](#Constants1) | GUIDs that specify the various parameter categories of the **EncodeParameter** objects. |
| [Image File Format Constants](#Constants2) | GUIDs that specify the that identify the file formats of an image. |
| [Image Frame Dimension Constants](#Constants3) | GUIDs that identify the page and time dimensions. |
| [Image Pixel Format Constants](#Constants4) | Pixel formats used in bitmaps. |
| [Image Property Tag Type Constants](#Constants5) | Specifies the data type of the values stored in the value data member of a **PropertyItem** structure. |
| [Image Property Tag Constants](#Constants6) | Property items supported by Microsoft Windows GDI+. |

# <a name="BitmapData"></a>BitmapData Structure

A **BitmapData** structure stores attributes of a bitmap. Used by the **GdipBitmapLockBits** and **GdipBitmapUnlockBits** functions.

```
TYPE BitmapData
   Width AS UINT
   Height AS UINT
   Stride AS INT_
   PixelFormat AS PixelFormat
   Scan0 AS ANY PTR
   Reserved AS UINT_PTR
END TYPE
```

| Member     | Description |
| ---------- | ----------- |
| **Width** | Number of pixels in one scan line of the bitmap. |
| **Height** | Number of scan lines in the bitmap. |
| **Stride** | Offset, in bytes, between consecutive scan lines of the bitmap. If the stride is positive, the bitmap is top-down. If the stride is negative, the bitmap is bottom-up. |
| **PixelFormat** | Integer that specifies the pixel format of the bitmap. |
| **Scan0** | Pointer to the first (index 0) scan line of the bitmap. |
| **Reserved** | Reserved for future use. |

# <a name="CharacterRange"></a>CharacterRange Structure

A **CharacterRange** structure specifies a range of character positions within a string.

```
TYPE CharacterRange
   First AS INT_
   Length AS INT_
END TYPE
```

| Member     | Description |
| ---------- | ----------- |
| **First** | Specifies the first position of this range. |
| **Length** | Specifies the number of positions in this range. |

#### Remarks

A character range is a range of character positions within a string of text. The area of the display that is occupied by a group of characters that are specified by the character range is the bounding region. A character range is set by **GdipSetStringFormatMeasurableCharacterRanges**. The number of ranges that are currently set can be determined by calling **GdipGetStringFormatMeasurableCharacterRangeCount**. This number is also the number of regions expected to be obtained by the **GdipMeasureCharacterRanges** function.

# <a name="ColorMap"></a>ColorMap Structure

The **ColorMap** structure contains two COLORREF values. Several of the **ImageAttributes** functions adjust image colors by using a color remap table, which is an array of **ColorMap** structures.

```
TYPE ColorMap
   oldColor AS Color
   newColor AS Color
END TYPE
```

| Member     | Description |
| ---------- | ----------- |
| **oldColor** | The original color. |
| **newColor** | The new color. |

#### Remarks

Because the FreeBasic 32-bit headers don't define the GDI+ **ColorMap** structure ad use the GDI **ColorMap** one, that has different member names, there is a conflict. To solve it, I have defined an union that can be used to replace **ColorMap** in both 32 and 64-bit.

```
UNION GDIP_COLORMAP
   ' // For compatibility with GDI+
   TYPE
      oldColor AS COLORREF
      newColor AS COLORREF
   END TYPE
   ' // For compatibility with GDI
   TYPE
      from AS COLORREF
      to AS COLORREF
   END TYPE
END UNION
```

# <a name="ColorMatrix"></a>ColorMatrix Structure

The **ColorMap** structure contains two COLORREF values. Several of the **ImageAttributes** functions adjust image colors by using a color remap table, which is an array of **ColorMap** structures.

```
TYPE ColorMatrix
   m(0 to 5-1, 0 to 5-1) AS REAL
END TYPE
```

| Member     | Description |
| ---------- | ----------- |
| **m** | 5x5 array of real numbers. |

#### Remarks

A 5x5 color matrix is a homogeneous matrix for a 4-space transformation. The element in the fifth row and fifth column of a 5x5 homogeneous matrix must be 1, and all of the other elements in the fifth column must be 0. Color matrices are used to transform color vectors. The first four components of a color vector hold the red, green, blue, and alpha components (in that order) of a color. The fifth component of a color vector is always 1.

# <a name="ColorPalette"></a>ColorPalette Structure

The **ColorPalette** structure defines an array of colors that make up a color palette. The colors are 32-bit ARGB colors.

```
TYPE ColorPalette
   Flags AS UINT
   Count AS UINT
   Entries(0) AS ARGB
END TYPE
```

| Member     | Description |
| ---------- | ----------- |
| **Flags** | Combination of flags from the **PaletteFlags** enumeration. |
| **Count** | Number of elements in the **Entries** array. |
| **Entries** | Array of ARGB colors.  |

# <a name="EncoderParameter"></a>EncoderParameter Structure

An **EncoderParameter** structure holds a parameter that can be passed to an image encoder. An **EncoderParameter** structure can also be used to receive a list of possible values supported by a particular parameter of a particular image encoder.

```
TYPE EncoderParameter
   Guid AS GUID
   NumberOfValues AS ULONG
   Type AS ULONG
   Value AS ANY PTR
END TYPE
```

| Member     | Description |
| ---------- | ----------- |
| **Guid** | Identifies the parameter category. GUIDs that represent various parameter categories (**EncoderCompression**, **EncoderColorDepth**, and the like) are defined in Gdiplusimaging.inc. |
| **NumberOfValues** | Number of values in the array pointed to by the **Value** data member. |
| **Type** | Identifies the data type of the parameter. The **EncoderParameterValueType** enumeration in Gdiplusenums.inc defines several possible value types. |
| **Value** | Pointer to an array of values. Each value has the type specified by the **Type** data member. |

# <a name="EncoderParameters"></a>EncoderParameters Structure

An **EncoderParameters** structure is an array of **EncoderParameter** structures along with a data member that specifies the number of **EncoderParameter** structures in the array.

```
TYPE EncoderParameters
   Count AS UINT
   Parameter(0) AS EncoderParameter
END TYPE
```

| Member     | Description |
| ---------- | ----------- |
| **Count** | Number of **EncoderParameter** structures in the array. |
| **Parameter ()** | Array of **EncoderParameter** structures. |

#### Remarks

When you create an **EncoderParameters** structure, you must allocate enough memory to hold all of the **EncoderParameter** structures that will eventually be placed in the array.

# <a name="ENHMETAHEADER3"></a>ENHMETAHEADER3 Structure

The **ENHMETAHEADER3** structure contains enhanced-metafile data including the dimensions of the metafile image, the number of records in the metafile, and the resolution of the device on which the metafile was created.

```
TYPE ENHMETAHEADER3
   itype AS DWORD
   nSize AS DWORD
   rclBounds AS RECTL
   rclFrame AS RECTL
   dSignature AS DWORD
   nVersion AS DWORD
   nBytes AS DWORD
   nRecords AS DWORD
   nHandles AS WORD
   sReserved AS WORD
   nDescription AS DWORD
   offDescription AS DWORD
   nPalEntries AS DWORD
   szlDevice AS SIZEL
   szlMillimeters AS SIZEL
END TYPE
```

| Member     | Description |
| ---------- | ----------- |
| **iType** | Record type. Value is always EMR_HEADER. |
| **nSize** | Structure size, in bytes. This may be greater than the value returned by SIZEOF (ENHMETAHEADER3). |
| **rclBounds** | Bounding rectangle, in device units, for the image stored in the metafile. |
| **rclFrame** | Rectangle, in 0.01 millimeter units, that surrounds the image stored in the metafile. |
| **dSignature** | Must be ENHMETA_SIGNATURE. |
| **nVersion** | Version number of the metafile format. The current version is &H10000. |
| **nBytes** | Size, in bytes, of the metafile. |
| **nRecords** | Number of records in the metafile. |
| **nHandles** | Number of handles in the metafile handle table. Handle index zero is reserved. |
| **sReserved** | Reserved. Must be zero. |
| **nDescription** | Number of characters in the string that contains the description of the metafile's contents. This member is 0 if the metafile does not have a description string. |
| **offDescription** | Offset from the beginning of the ENHMETAHEADER3 structure to the string that contains the description of the metafile's contents. This member is 0 if the metafile does not have a description string. |
| **nPalEntries** | Number of entries in the metafile palette. |
| **szlDevice** | Resolution, in pixels, of the reference device. |
| **szlMillimeters** | Resolution, in millimeters, of the reference device. |

# <a name="GdiplusStartupInput"></a>GdiplusStartupInput Structure

The **GdiplusStartupInput** structure holds a block of arguments that are required by the **GdiplusStartup** function.

```
TYPE GdiplusStartupInput
   GdiplusVersion AS UINT32
   DebugEventCallback AS DebugEventProc
   SuppressBackgroundThread AS BOOL
   SuppressExternalCodecs AS BOOL
END TYPE
```

| Member     | Description |
| ---------- | ----------- |
| **GdiplusVersion** | Specifies the version of Microsoft Windows GDI+. Must be 1. |
| **DebugEventCallback** | Pointer to a callback function that GDI+ can call, on debug builds, for assertions and warnings. The default value is NULL. |
| **SuppressBackgroundThread** | Boolean value that specifies whether to suppress the GDI+ background thread. If you set this member to TRUE, GdiplusStartup returns (in its output parameter) a pointer to a hook function and a pointer to an unhook function. You must call those functions appropriately to replace the background thread. If you do not want to be responsible for calling the hook and unhook functions, set this member to FALSE. The default value is FALSE. |
| **SuppressExternalCodecs** | Boolean value that specifies whether you want GDI+ to suppress external image codecs. GDI+ version 1.0 does not support external image codecs, so this parameter is ignored. |

#### Remarks

The **GdiplusStartupInput** structure provides a constructor that sets the **GdiplusVersion** member to 1 and allows you to specify values for the other three members. All of the constructor parameters are optional, so you can declare a variable of type GdiplusStartupInput without passing any arguments to the constructor, and all of the members will be initialized with appropriate default values.

If you set the **SuppressBackgroundThread** member of the **GdiplusStartupinput** parameter to TRUE, you must call the hook and unhook functions returned in the output parameter. Call those functions before and after the application's main message loop; that is, a message loop that is active for the lifetime of GDI+. Call the hook function before the loop starts, and call the unhook function after the loop ends.

# <a name="GdiplusStartupOutput"></a>GdiplusStartupOutput Structure

The **GdiplusStartup** function uses the **GdiplusStartupOutput** structure to return (in its output parameter) a pointer to a hook function and a pointer to an unhook function. If you set the **SuppressBackgroundThread** member of the input parameter to TRUE, then you are responsible for calling those functions to replace the Microsoft Windows GDI+ background thread.

Call the hook and unhook functions before and after the application's main message loop; that is, a message loop that is active for the lifetime of GDI+. Call the hook function before the loop starts, and call the unhook function after the loop ends. The token parameter of the hook function receives an identifier that you should later pass to the unhook function. If you do not pass the proper identifier (the one returned by the hook function) to the unhook function, there will be resource leaks that won't be cleaned up until the process exits.

If you do not want to be responsible for calling the hook and unhook functions, set the **SuppressBackgroundThread** member of the input parameter (passed to **GdiplusStartup**) to FALSE.

```
TYPE GdiplusStartupOutput
   NotificationHook AS NotificationHookProc
   NotificationUnhook AS NotificationUnhookProc
END TYPE
```

| Member     | Description |
| ---------- | ----------- |
| **NotificationHook** | Receives a pointer to a hook function. |
| **NotificationUnhook** | Receives a pointer to an unhook function. |

# <a name="GpPathData"></a>GpPathData Structure

The **GpPathData** structure is used by the **GdipGetPathData** function. It has an array of points and an array of types. Each element in the array of types is a byte that specifies the point type and a set of flags for the corresponding element in the array of points. You can use a **GpathData** structure to get or set the data points (and their types) of a path.

```
TYPE GpPathData
   Count AS INT_
   Points AS GpPointF PTR
   Types AS BYTE PTR
END TYPE
```
-also-
```
type GpPathData as any
```

| Member     | Description |
| ---------- | ----------- |
| **NotificationHook** | Receives a pointer to a hook function. |
| **NotificationUnhook** | Receives a pointer to an unhook function. |

#### Remarks

Because of another mess in the FreeBASIC Gdi+ headers, this structure can't be used with 32 and 64-bit at the same time. Therefore, I have declared a replacement in AfxGdiPlus.inc.

```
TYPE GDIP_PATHDATA
   Count AS LONG
   Points AS GpPointF PTR
   Types AS UBYTE PTR
END TYPE
```

# <a name="GpPoint"></a>GpPoint Structure

Defines the x- and y-coordinates of a point.

```
TYPE GpPoint
   x AS INT_
   y AS INT_
END TYPE
```

| Member     | Description |
| ---------- | ----------- |
| **x** | Specifies the x-coordinate of a point. |
| **y** | Specifies the y-coordinate of a point. |

# <a name="GpPointF"></a>GpPointF Structure

Defines the x- and y-coordinates of a point.

```
TYPE GpPoint
   x AS FLOAT
   y AS FLOAT
END TYPE
```

| Member     | Description |
| ---------- | ----------- |
| **x** | Specifies the x-coordinate of a point. |
| **y** | Specifies the y-coordinate of a point. |

# <a name="GpRect"></a>GpRect Structure

Defines the upper-left corner, width, and height of a rectangle.

```
TYPE GpRect
   x AS INT_
   y AS INT_
   width AS INT_
   height AS INT_
END TYPE
```

| Member     | Description |
| ---------- | ----------- |
| **x** | Specifies the x-coordinate of the upper-left corner of a rectangle. |
| **y** | Specifies the y-coordinate of the upper-left corner of a rectangle. |
| **width** | Specifies the width of a rectangle. |
| **height** | Specifies the heigth of a rectangle. |

# <a name="GpRectF"></a>GpRectF Structure

Defines the upper-left corner, width, and height of a rectangle.

```
TYPE GpRect
   x AS REAL
   y AS REAL
   width AS REAL
   height AS REAL
END TYPE
```

| Member     | Description |
| ---------- | ----------- |
| **x** | Specifies the x-coordinate of the upper-left corner of a rectangle. |
| **y** | Specifies the y-coordinate of the upper-left corner of a rectangle. |
| **width** | Specifies the width of a rectangle. |
| **height** | Specifies the heigth of a rectangle. |

# <a name="GpSize"></a>GpSize Structure

Encapsulates a Width and Height dimension in a 2-D coordinate system.

```
TYPE GpSize
   Width AS INT_
   Height AS INT_
END TYPE
```

| Member     | Description |
| ---------- | ----------- |
| **Width** | Vertical measurement. |
| **Height** | Horizontal measurement. |

# <a name="GpSizeF"></a>GpSizeF Structure

Encapsulates a Width and Height dimension in a 2-D coordinate system.

```
TYPE GpSizeF
   Width AS REAL
   Height AS REAL
END TYPE
```

| Member     | Description |
| ---------- | ----------- |
| **Width** | Vertical measurement. |
| **Height** | Horizontal measurement. |

# <a name="ImageCodecInfo"></a>ImageCodecInfo Structure

An **ImageCodecInfo** structure stores information about an image codec (encoder/decoder). GDI+ provides several built-in image codecs. You can obtain information about those codecs by calling the **GetImageEncoders** function and the **GetImageDecoders** function. Each of those functions returns an array of **ImageCodecInfo** structures, one for each available encoder or decoder.

```
TYPE ImageCodecInfo
   Clsid AS CLSID
   FormatID AS GUID
   CodecName AS WCHAR PTR
   DllName AS WCHAR PTR
   FormatDescription AS WCHAR PTR
   FilenameExtension AS WCHAR PTR
   MimeType AS WCHAR PTR
   Flags AS DWORD
   Version AS DWORD
   SigCount AS DWORD
   SigSize AS DWORD
   SigPattern AS UBYTE PTR
   SigMask AS UBYTE PTR
END TYPE
```

| Member     | Description |
| ---------- | ----------- |
| **Clsid** | Codec identifier. |
| **FormatID** | File format identifier. GUIDs that identify various file formats (**ImageFormatBMP**, **ImageFormatEMF**, and the like) are defined in Gdiplusimaging.inc. |
| **CodecName** | Pointer to a null-terminated wide string that contains the codec name. |
| **DllName** | Pointer to a null-terminated wide string that contains the path name of the DLL in which the codec resides. If the codec is not in a DLL, this pointer is NULL. |
| **FormatDescription** | Pointer to a null-terminated wide string that contains the name of the file format used by the codec. |
| **FilenameExtension** | Pointer to a null-terminated wide string that contains all file-name extensions associated with the codec. The extensions are separated by semicolons. |
| **MimeType** | Pointer to a null-terminated wide string that contains the mime type of the codec. |
| **Flags** | Combination of flags from the **ImageCodecFlags** enumeration. |
| **Version** | Integer that indicates the version of the codec. |
| **SigCount** | Integer that indicates the number of signatures used by the file format associated with the codec. |
| **SigSize** | Integer that indicates the number of bytes in each signature. |
| **SigPattern** | Pointer to an array of bytes that contains the pattern for each signature. |
| **SigMask** | Pointer to an array of bytes that contains the mask for each signature. |

# <a name="LOGFONTA"></a>LOGFONTA Structure

Defines the attributes of a font.

```
TYPE LOGFONTA
   lfHeight AS LONG
   lfWidth AS LONG
   lfEscapement AS LONG
   lfOrientation AS LONG
   lfWeight AS LONG
   lfItalic AS UBYTE
   lfUnderline AS UBYTE
   lfStrikeOut AS UBYTE
   lfCharSet AS UBYTE
   lfOutPrecision AS UBYTE
   lfClipPrecision AS UBYTE
   lfQuality AS UBYTE
   lfPitchAndFamily AS UBYTE
   lfFaceName AS ZSTRING * 32   ' LF_FACESIZE
END TYPE
```

| Member     | Description |
| ---------- | ----------- |
| **lfHeight** | Value of type LONG that specifies the height, in logical units, of the font's character cell or character. |
| **lfWidth** | Value of type LONG that specifies the width, in logical units, of characters in the font.  |
| **lfEscapement** | Value of type LONG that contains the angle, in tenths of degrees, between the escapement vector and the x-axis of the device. The escapement vector is parallel to the base line of a row of text. |
| **lfOrientation** | Value of type LONG that specifies the angle, in tenths of degrees, between each character's base line and the x-axis of the device. |
| **lfWeight** | Value of type LONG that specifies the weight of the font in the range from 0 through 1000. |
| **lfItalic** | Value of type BYTE that specifies an italic font if set to TRUE. |
| **lfUnderline** | Value of type BYTE that specifies an underlined font if set to TRUE. |
| **lfStrikeOut** | Value of type BYTE that specifies a strikeout font if set to TRUE. |
| **lfCharSet** | Value of type BYTE that specifies the character set. |
| **lfOutPrecision** | Value of type BYTE that specifies the output precision. The output precision defines how closely the output must match the requested font's height, width, character orientation, escapement, pitch, and font type. |
| **lfClipPrecision** | Value of type BYTE that specifies the clipping precision. The clipping precision defines how to clip characters that are partially outside the clipping region. |
| **lfQuality** | Value of type BYTE that specifies the output quality. The output quality defines how carefully the Microsoft Windows Graphics Device Interface (GDI) must attempt to match the logical-font attributes to those of an actual physical font. |
| **lfPitchAndFamily** | Value of type BYTE that specifies the pitch and family of the font. |
| **lfFaceName** | Array of characters that contains a null-terminated string that specifies the typeface name of the font. The length of the string must not exceed 32 characters, including the NULL terminator. |

# <a name="LOGFONTW"></a>LOGFONTW Structure

Defines the attributes of a font.

```
TYPE LOGFONTW
   lfHeight AS LONG
   lfWidth AS LONG
   lfEscapement AS LONG
   lfOrientation AS LONG
   lfWeight AS LONG
   lfItalic AS UBYTE
   lfUnderline AS UBYTE
   lfStrikeOut AS UBYTE
   lfCharSet AS UBYTE
   lfOutPrecision AS UBYTE
   lfClipPrecision AS UBYTE
   lfQuality AS UBYTE
   lfPitchAndFamily AS UBYTE
   lfFaceName AS WSTRING * 32   ' LF_FACESIZE
END TYPE
```

| Member     | Description |
| ---------- | ----------- |
| **lfHeight** | Value of type LONG that specifies the height, in logical units, of the font's character cell or character. |
| **lfWidth** | Value of type LONG that specifies the width, in logical units, of characters in the font.  |
| **lfEscapement** | Value of type LONG that contains the angle, in tenths of degrees, between the escapement vector and the x-axis of the device. The escapement vector is parallel to the base line of a row of text. |
| **lfOrientation** | Value of type LONG that specifies the angle, in tenths of degrees, between each character's base line and the x-axis of the device. |
| **lfWeight** | Value of type LONG that specifies the weight of the font in the range from 0 through 1000. |
| **lfItalic** | Value of type BYTE that specifies an italic font if set to TRUE. |
| **lfUnderline** | Value of type BYTE that specifies an underlined font if set to TRUE. |
| **lfStrikeOut** | Value of type BYTE that specifies a strikeout font if set to TRUE. |
| **lfCharSet** | Value of type BYTE that specifies the character set. |
| **lfOutPrecision** | Value of type BYTE that specifies the output precision. The output precision defines how closely the output must match the requested font's height, width, character orientation, escapement, pitch, and font type. |
| **lfClipPrecision** | Value of type BYTE that specifies the clipping precision. The clipping precision defines how to clip characters that are partially outside the clipping region. |
| **lfQuality** | Value of type BYTE that specifies the output quality. The output quality defines how carefully the Microsoft Windows Graphics Device Interface (GDI) must attempt to match the logical-font attributes to those of an actual physical font. |
| **lfPitchAndFamily** | Value of type BYTE that specifies the pitch and family of the font. |
| **lfFaceName** | Array of wide characters that contains a null-terminated string that specifies the typeface name of the font. The length of the string must not exceed 32 characters, including the NULL terminator. |

# <a name="MetafileHeader"></a>MetafileHeader Structure

A **MetafileHeader** structure stores properties of an associated metafile.

```
UNION MetafileHeader_UNION
   WmfHeader AS METAHEADER
   EmfHeader AS ENHMETAHEADER3
END UNION

TYPE MetafileHeader
   Type_ AS MetafileType     ' // The type of the associated metafile
   Size AS UINT              ' // Size of the metafile (in bytes)
   Version AS UINT           ' // EMF+, EMF, or WMF version
   EmfPlusFlags AS UINT
   DpiX AS REAL
   DpiY AS REAL
   X AS INT_                 ' // Bounds in device units
   Y AS INT_
   Width AS INT_
   Height AS INT_
   union
      WmfHeader as METAHEADER
      EmfHeader as ENHMETAHEADER3
   end union
   EmfPlusHeaderSize AS INT_ ' // size of the EMF+ header in file
   LogicalDpiX AS INT_       ' // Logical Dpi of reference Hdc
   LogicalDpiY AS INT_       ' // usually valid only for EMF+
END TYPE
```

#### Remarks

Metafiles provide a device-independent and application-independent way to share pictures. They contain records that describe a sequence of graphics application programming interfaces (APIs) to invoke in a particular order with their associated graphics data. Metafiles can be recorded by an application and later played back by that application or by another one to reproduce a particular picture. They can also be used to send content to a print spooler. Enhanced metafiles support the ability to provide both Microsoft Windows GDI+ and Windows Graphics Device Interface (GDI) descriptions of the same picture so that both GDI+ and down-level GDI applications can render it.

# <a name="PropertyItem"></a>PropertyItem Structure

A **PropertyItem** structure holds one piece of image metadata.

```
TYPE PropertyItem
   id AS PROPID
   length AS ULONG
   wtype AS WORD
   value AS ANY PTR
END TYPE
```

| Member     | Description |
| ---------- | ----------- |
| **id** | Integer that identifies the kind of metadata stored in this PropertyItem object. Constants that identify various kinds of metadata (**PropertyTagEquipMake**, P**ropertyTagEquipModel**, and the like) are defined in Gdiplusimaging.inc. The **PROPID** data type is defined in Wtypes.inc. |
| **length** | Size, in bytes, of the value array. |
| **wtype** | Integer that identifies the data type of the values in the value array. Constants that identify various data types (**PropertyTagTypeByte**, **PropertyTagTypeASCII**, and the like) are defined in Gdiplusimaging.inc. |
| **value** | Pointer to an array of values. Each value in the array has the data type specified by the type data member. |

# <a name="PWMFRect16"></a>PWMFRect16 Structure

The **PWMFRect16** structure defines a rectangle that bounds a placeable metafile. The rectangle defines the size and position of the metafile.

```
TYPE PWMFRect16
   Left AS INT16
   Top AS INT16
   Right AS INT16
   Bottom AS INT16
END TYPE
```

| Member     | Description |
| ---------- | ----------- |
| **Left** | Left coordinate of the rectangle that bounds the metafile. |
| **Top** | Top coordinate of the rectangle that bounds the metafile. |
| **Right** | Right coordinate of the rectangle that bounds the metafile. |
| **Bottom** | Bottom coordinate of the rectangle that bounds the metafile. |

# <a name="WmfPlaceableFileHeader"></a>WmfPlaceableFileHeader Structure

The **WmfPlaceableFileHeader** structure defines the fields of a placeable metafile header. Placeable metafiles were created as a way of specifying how a metafile is mapped and scaled on a display device.

```
TYPE WmfPlaceableFileHeader
   Key AS UINT32
   Hmf AS INT16
   BoundingBox AS PWMFRect16
   Inch AS INT16
   Reserved AS UINT32
   Checksum AS INT16
END TYPE
```

| Member     | Description |
| ---------- | ----------- |
| **Key** | Identification value that indicates the presence of a placeable metafile header. This value is always &h9AC6CDD7. |
| **Hmf** | Handle to the metafile in memory. When written to disk, this field is not used and will always contains the value 0. |
| **BoundingBox** | Destination rectangle, measured in twips, for displaying the metafile. |
| **Inch** | Number of twips per inch used to represent the image.<br>Normally, there are 1440 twips per inch; however, this number can be changed to scale the image.<br>* A value of 720 specifies that the image is twice its normal size.<br>* A value of 360 specifies that the image is four times its normal size.<br>* A value of 2880 specifies that the image is half its normal size. |
| **Reserved** | Not used and is always set to 0. |
| **Checksum** | Checksum for the previous 10 WORDs in the header. This value can be used to determine whether the metafile has become corrupted. |

#### Remarks

Although placeable metafiles are quite common, they are not directly supported by the Microsoft Windows application programming interface (API). To display a placeable metafile using the Windows API, you must first strip the placeable metafile header from the file. This is typically performed by copying the metafile to a temporary file starting at file offset 22 (0x16). This is because each placeable metafile begins with a 22-byte header that is followed by a standard metafile. 

# <a name="BrushType"></a>BrushType Enumeration

The **BrushType** enumeration indicates the type of brush. There are five types of brushes.

```
BrushTypeSolidColor = 0
BrushTypeHatchFill = 1
BrushTypeTextureFill = 2
BrushTypePathGradient = 3
BrushTypeLinearGradient = 4
```

| Constant   | Meaning     |
| ---------- | ----------- |
| **BrushTypeSolidColor** | Indicates a brush of type SolidBrush. A solid brush paints a single, constant color that can be opaque or transparent. |
| **BrushTypeHatchFill** | Indicates a brush of type **HatchBrush**. A hatch brush paints a background and paints, over that background, a pattern of lines, dots, dashes, squares, crosshatch, or some variation of these. The hatch brush consists of two colors: one for the background and one for the pattern over the background. The color of the background is called the background color, and the color of the pattern is called the foreground color.  |
| **BrushTypeTextureFill** | Indicates a brush of type **TextureBrush**. A texture brush paints an image. The image or texture is either a portion of a specified image or a scaled version of a specified image. The type of image (metafile or nonmetafile) determines whether the texture is a portion of the image or a scaled version of the image. |
| **BrushTypePathGradient** | Indicates a brush of type **PathGradientBrush**. A path gradient brush paints a color gradient in which the color changes from a center point outward to a boundary that is defined by a closed curve or path. The color gradient has one color at the center point and one or multiple colors at the boundary. |
| **BrushTypeLinearGradient** | Indicates a brush of type **LinearGradientBrush**. A linear gradient brush paints a color gradient in which the color changes evenly from the starting boundary line of the linear gradient brush to the ending boundary line of the linear gradient brush. The boundary lines of a linear gradient brush are two parallel straight lines. The color gradient is perpendicular to the boundary lines of the linear gradient brush, changing gradually across the stroke from the starting boundary line to the ending boundary line. The color gradient has one color at the starting boundary line and another color at the ending boundary line.  |

# <a name="Color"></a>Color Enumeration

```
AliceBlue = &hFFF0F8FF
AntiqueWhite = &hFFFAEBD7
Aqua = &hFF00FFFF
Aquamarine = &hFF7FFFD4
Azure = &hFFF0FFFF
Beige = &hFFF5F5DC
Bisque = &hFFFFE4C4
Black = &hFF000000
BlanchedAlmond = &hFFFFEBCD
Blue = &hFF0000FF
BlueViolet = &hFF8A2BE2
Brown = &hFFA52A2A
BurlyWood = &hFFDEB887
CadetBlue = &hFF5F9EA0
Chartreuse = &hFF7FFF00
Chocolate = &hFFD2691E
Coral = &hFFFF7F50
CornflowerBlue = &hFF6495ED
Cornsilk = &hFFFFF8DC
Crimson = &hFFDC143C
Cyan = &hFF00FFFF
DarkBlue = &hFF00008B
DarkCyan = &hFF008B8B
DarkGoldenrod = &hFFB8860B
DarkGray = &hFFA9A9A9
DarkGreen = &hFF006400
DarkKhaki = &hFFBDB76B
DarkMagenta = &hFF8B008B
DarkOliveGreen = &hFF556B2F
DarkOrange = &hFFFF8C00
DarkOrchid = &hFF9932CC
DarkRed = &hFF8B0000
DarkSalmon = &hFFE9967A
DarkSeaGreen = &hFF8FBC8B
DarkSlateBlue = &hFF483D8B
DarkSlateGray = &hFF2F4F4F
DarkTurquoise = &hFF00CED1
DarkViolet = &hFF9400D3
DeepPink = &hFFFF1493
DeepSkyBlue = &hFF00BFFF
DimGray = &hFF696969
DodgerBlue = &hFF1E90FF
Firebrick = &hFFB22222
FloralWhite = &hFFFFFAF0
ForestGreen = &hFF228B22
Fuchsia = &hFFFF00FF
Gainsboro = &hFFDCDCDC
GhostWhite = &hFFF8F8FF
Gold = &hFFFFD700
Goldenrod = &hFFDAA520
Gray = &hFF808080
Green = &hFF008000
GreenYellow = &hFFADFF2F
Honeydew = &hFFF0FFF0
HotPink = &hFFFF69B4
IndianRed = &hFFCD5C5C
Indigo = &hFF4B0082
Ivory = &hFFFFFFF0
Khaki = &hFFF0E68C
Lavender = &hFFE6E6FA
LavenderBlush = &hFFFFF0F5
LawnGreen = &hFF7CFC00
LemonChiffon = &hFFFFFACD
LightBlue = &hFFADD8E6
LightCoral = &hFFF08080
LightCyan = &hFFE0FFFF
LightGoldenrodYellow = &hFFFAFAD2
LightGray = &hFFD3D3D3
LightGreen = &hFF90EE90
LightPink = &hFFFFB6C1
LightSalmon = &hFFFFA07A
LightSeaGreen = &hFF20B2AA
LightSkyBlue = &hFF87CEFA
LightSlateGray = &hFF778899
LightSteelBlue = &hFFB0C4DE
LightYellow = &hFFFFFFE0
Lime = &hFF00FF00
LimeGreen = &hFF32CD32
Linen = &hFFFAF0E6
Magenta = &hFFFF00FF
Maroon = &hFF800000
MediumAquamarine = &hFF66CDAA
MediumBlue = &hFF0000CD
MediumOrchid = &hFFBA55D3
MediumPurple = &hFF9370DB
MediumSeaGreen = &hFF3CB371
MediumSlateBlue = &hFF7B68EE
MediumSpringGreen = &hFF00FA9A
MediumTurquoise = &hFF48D1CC
MediumVioletRed = &hFFC71585
MidnightBlue = &hFF191970
MintCream = &hFFF5FFFA
MistyRose = &hFFFFE4E1
Moccasin = &hFFFFE4B5
NavajoWhite = &hFFFFDEAD
Navy = &hFF000080
OldLace = &hFFFDF5E6
Olive = &hFF808000
OliveDrab = &hFF6B8E23
Orange = &hFFFFA500
OrangeRed = &hFFFF4500
Orchid = &hFFDA70D6
PaleGoldenrod = &hFFEEE8AA
PaleGreen = &hFF98FB98
PaleTurquoise = &hFFAFEEEE
PaleVioletRed = &hFFDB7093
PapayaWhip = &hFFFFEFD5
PeachPuff = &hFFFFDAB9
Peru = &hFFCD853F
Pink = &hFFFFC0CB
Plum = &hFFDDA0DD
PowderBlue = &hFFB0E0E6
Purple = &hFF800080
Red = &hFFFF0000
RosyBrown = &hFFBC8F8F
RoyalBlue = &hFF4169E1
SaddleBrown = &hFF8B4513
Salmon = &hFFFA8072
SandyBrown = &hFFF4A460
SeaGreen = &hFF2E8B57
SeaShell = &hFFFFF5EE
Sienna = &hFFA0522D
Silver = &hFFC0C0C0
SkyBlue = &hFF87CEEB
SlateBlue = &hFF6A5ACD
SlateGray = &hFF708090
Snow = &hFFFFFAFA
SpringGreen = &hFF00FF7F
SteelBlue = &hFF4682B4
Tan = &hFFD2B48C
Teal = &hFF008080
Thistle = &hFFD8BFD8
Tomato = &hFFFF6347
Transparent = &hFFFFFF
Turquoise = &hFF40E0D0
Violet = &hFFEE82EE
Wheat = &hFFF5DEB3
White = &hFFFFFFFF
WhiteSmoke = &hFFF5F5F5
Yellow = &hFFFFFF00
YellowGreen = &hFF9ACD32
```
```
AlphaShift = 24
RedShift = 16
GreenShift = 8
BlueShift = 0
```
```
AlphaMask = &hFF000000
RedMask = &h00FF0000
GreenMask = &h0000FF00
BlueMask = &hh000000FF
```

# <a name="ColorAdjustType"></a>ColorAdjustType Enumeration

The **ColorAdjustType** enumeration specifies which GDI+ objects use color-adjustment information. You can adjust the colors in a rendered image using the **ImageAttributes functions**. There are five adjustment categories: default, bitmap, brush, pen, and text.

```
ColorAdjustTypeDefault = 0
ColorAdjustTypeBitmap  = 1
ColorAdjustTypeBrush   = 2
ColorAdjustTypePen     = 3
ColorAdjustTypeText    = 4
ColorAdjustTypeCount   = 5
ColorAdjustTypeAny     = 6
```

| Constant   | Meaning     |
| ---------- | ----------- |
| ColorAdjustTypeDefault | Specifies that color or grayscale adjustment applies to all categories that do not have adjustment settings of their own. |
| ColorAdjustTypeBitmap | Specifies that color or grayscale adjustment applies to bitmapped images. |
| ColorAdjustTypeBrush | Specifies that color or grayscale adjustment applies to brush operations in metafiles. |
| ColorAdjustTypePen | Specifies that color or grayscale adjustment applies to pen operations in metafiles. |
| ColorAdjustTypeText | Specifies that color or grayscale adjustment applies to text drawn in metafiles. |
| ColorAdjustTypeCount | Used internally to record the number of color adjustment types. |
| ColorAdjustTypeAny | Reserved. |

# <a name="ColorChannelFlags"></a>ColorChannelFlags Enumeration

The **ColorChannelFlags** enumeration specifies individual channels in the CMYK (cyan, magenta, yellow, black) color space. This enumeration is used by the GdipSetImageAttributesOutputChannel function.

```
ColorChannelFlagsC    = 0
ColorChannelFlagsM    = 1
ColorChannelFlagsY    = 2
ColorChannelFlagsK    = 3
ColorChannelFlagsLast = 4
```

| Constant   | Meaning     |
| ---------- | ----------- |
| ColorChannelFlagsC | Specifies the cyan color channel. |
| ColorChannelFlagsM | Specifies the magenta color channel. |
| ColorChannelFlagsY | Specifies the yellow color channel. |
| ColorChannelFlagsK | Specifies the black color channel. |

# <a name="ColorMatrixFlags"></a>ColorMatrixFlags Enumeration

The **ColorMatrixFlags** enumeration specifies the types of images and colors that will be affected by the color and grayscale adjustment settings. It is used by the GdipSetImageAttributesColorMatrix function.

```
ColorMatrixFlagsDefault = 0
ColorMatrixFlagsSkipGrays = 1
ColorMatrixFlagsAltGray = 2
```

| Constant   | Meaning     |
| ---------- | ----------- |
| ColorMatrixFlagsDefault | Specifies that all color values (including grays) are adjusted by the same color-adjustment matrix. |
| ColorMatrixFlagsSkipGrays | Specifies that colors are adjusted but gray shades are not adjusted. A gray shade is any color that has the same value for its red, green, and blue components. |
| ColorMatrixFlagsAltGray | Specifies that colors are adjusted by one matrix and gray shades are adjusted by another matrix. |

# <a name="ColorMode"></a>ColorMode Enumeration

```
ColorModeARGB32 = 0
ColorModeARGB64 = 1
```

# <a name="CombineMode"></a>CombineMode Enumeration

The **CombineMode** enumeration specifies how a new region is combined with an existing region.

```
CombineModeReplace      = 0
CombineModeIntersect    = 1
CombineModeUnion        = 2
CombineModeXor          = 3
CombineModeExclude      = 4
CombineModeComplement   = 5
```

| Constant   | Meaning     |
| ---------- | ----------- |
| CombineModeReplace | Specifies that the existing region is replaced by the new region. |
| CombineModeIntersect | Specifies that the existing region is replaced by the intersection of itself and the new region. |
| CombineModeUnion | Specifies that the existing region is replaced by the union of itself and the new region. |
| CombineModeXor | Specifies that the existing region is replaced by the result of performing an XOR on the two regions. A point is in the XOR of two regions if it is in one region or the other but not in both regions. |
| CombineModeExclude | Specifies that the existing region is replaced by the portion of itself that is outside of the new region. |
| CombineModeComplement | Specifies that the existing region is replaced by the portion of the new region that is outside of the existing region. |

# <a name="CompositingMode"></a>CompositingMode Enumeration

The **CompositingMode** enumeration specifies how rendered colors are combined with background colors. This enumeration is used by the **GdipGetCompositingMode** and **GdipSetCompositingMode** functions.

```
CompositingModeSourceOver = 0
CompositingModeSourceCopy = 1
```

| Constant   | Meaning     |
| ---------- | ----------- |
| CompositingModeSourceOver | Specifies that when a color is rendered, it is blended with the background color. The blend is determined by the alpha component of the color being rendered. |
| CompositingModeSourceCopy | Specifies that when a color is rendered, it overwrites the background color. This mode cannot be used along with **TextRenderingHintClearTypeGridFit**. |

# <a name="CompositingQuality"></a>CompositingQuality Enumeration

The **CompositingQuality** enumeration specifies whether gamma correction is applied when colors are blended with background colors. This enumeration is used by the **GdipGetCompositingQuality** and **GdipSetCompositingQuality** functions.

```
CompositingQualityInvalid        = QualityModeInvalid
CompositingQualityDefault        = QualityModeDefault
CompositingQualityHighSpeed      = QualityModeLow
CompositingQualityHighQuality    = QualityModeHigh
CompositingQualityGammaCorrected = 3
CompositingQualityAssumeLinear   = 4
```

| Constant   | Meaning     |
| ---------- | ----------- |
| CompositingQualityDefault | Specifies that gamma correction is not applied. |
| CompositingQualityHighSpeed | Specifies that gamma correction is not applied. |
| CompositingQualityHighQuality | Specifies that gamma correction is applied. |
| CompositingQualityGammaCorrected | Specifies that gamma correction is applied. |
| CompositingQualityAssumeLinear | Specifies that gamma correction is not applied. |

#### Remarks

When you specify that gamma correction should not be applied, the image data to be rendered (blended with the background) is assumed to be in a linear color space with a gamma value of 1.0. As a result, no gamma adjustment is applied to the image data before or after blending the image with the background.

When you specify that gamma correction should be applied, the image data to be rendered (blended with the background) is assumed to be in the sRGB color space with a gamma value of 2.2. To ensure accurate blending, the input image data is transformed into a linear (gamma = 1.0) space before the colors are blended and transformed back into sRGB (gamma = 2.2) space afterward. This mode results in a more accurate blend at the expense of additional processing time. 

# <a name="CoordinateSpace"></a>CoordinateSpace Enumeration

The **CoordinateSpace** enumeration specifies coordinate spaces. This enumeration is used by the **GdipTransformPoints** and **GdipTransformPointsI** functions, which converts points from one coordinate space to another.

```
CoordinateSpaceWorld  = 0
CoordinateSpacePage   = 1
CoordinateSpaceDevice = 2
```

| Constant   | Meaning     |
| ---------- | ----------- |
| CoordinateSpaceWorld | Specifies the world coordinate space. |
| CoordinateSpacePage | Specifies the page coordinate space. |
| CoordinateSpaceDevice | Specifies the device coordinate space. |

# <a name="CustomLineCapType"></a>CustomLineCapType Enumeration

```
CustomLineCapTypeDefault = 0
CustomLineCapTypeAdjustableArrow = 1
```
# <a name="DashCap"></a>DashCap Enumeration

The **DashCap** enumeration specifies the type of graphic shape to use on both ends of each dash in a dashed line.

```
DashCapFlat = 0
DashCapRound = 2
DashCapTriangle = 3
```

| Constant   | Meaning     |
| ---------- | ----------- |
| DashCapFlat | Specifies a square cap that squares off both ends of each dash. |
| DashCapRound | Specifies a circular cap that rounds off both ends of each dash. |
| DashCapTriangle | Specifies a triangular cap that points both ends of each dash. |

#### Remarks

If you set the alignment of a Pen to **PenAlignmentInset**, you cannot use that pen to draw triangular dash caps.

# <a name="DashStyle"></a>DashStyle Enumeration

The **DashStyle** enumeration specifies the line style of a line drawn with a Microsoft Windows GDI+ pen. The line can be drawn by using one of several predefined styles or a custom style.

```
DashStyleSolid      = 0
DashStyleDash       = 1
DashStyleDot        = 2
DashStyleDashDot    = 3
DashStyleDashDotDot = 4
DashStyleCustom     = 5
```

| Constant   | Meaning     |
| ---------- | ----------- |
| DashStyleSolid | Specifies a solid line. |
| DashStyleDash | Specifies a dashed line. |
| DashStyleDot | Specifies a dotted line. |
| DashStyleDashDot | Specifies an alternating dash-dot line. |
| DashStyleDashDotDot | Specifies an alternating dash-dot-dot line. |
| DashStyleCustom | Specifies a user-defined, custom dashed line. |

# <a name="DriverStringOptions"></a>DriverStringOptions Enumeration

The **DriverStringOptions** enumeration specifies the spacing, orientation, and quality of the rendering for driver strings.

```
DriverStringOptionsCmapLookup      = 1
DriverStringOptionsVertical        = 2
DriverStringOptionsRealizedAdvance = 4
DriverStringOptionsLimitSubpixel   = 8
```

| Constant   | Meaning     |
| ---------- | ----------- |
| DriverStringOptionsCmapLookup | Specifies that the positions of the glyphs are taken from a character map lookup table. If this flag is not set, the positions are taken from an array of coordinates. |
| DriverStringOptionsVertical | Specifies that the string is displayed vertically. |
| DriverStringOptionsRealizedAdvance | Specifies that the glyph positions are calculated from the position of the first glyph. If this flag is not set, the glyph positions are obtained from an array of coordinates. |
| DriverStringOptionsLimitSubpixel | Specifies that less memory should be used for cache of antialiased glyphs. This also produces lower quality. If this flag is set, more memory is used, but the quality is higher. |

# <a name="EmfToWmfBitsFlags"></a>EmfToWmfBitsFlags Enumeration

Specifies options for the **GdipEmfToWmfBits** method, which converts an Enhanced Metafile (EMF) metafile to a Microsoft Windows Metafile Format (WMF) metafile.

```
EmfToWmfBitsFlagsDefault          = &H00000000
EmfToWmfBitsFlagsEmbedEmf         = &H00000001
EmfToWmfBitsFlagsIncludePlaceable = &H00000002
EmfToWmfBitsFlagsNoXORClip        = &H00000004
```

| Constant   | Meaning     |
| ---------- | ----------- |
| EmfToWmfBitsFlagsDefault | Specifies the default conversion. |
| EmfToWmfBitsFlagsEmbedEmf | Specifies that the source EMF metafile is embedded as a comment in the resulting WMF metafile.|
| EmfToWmfBitsFlagsIncludePlaceable | Specifies that the resulting WMF metafile is in the placeable metafile format; that is, it has the additional 22-byte header required by a placeable metafile. |
| EmfToWmfBitsFlagsNoXORClip | Specifies that the clipping region is stored in the metafile in the traditional way. If you do not set this flag, the **GdipEmfToWmfBits** function applies an optimization that stores the clipping region as a path and simulates clipping by using the XOR operator. |

# <a name="EmfType"></a>EmfType Enumeration

The **EmfType** enumeration specifies the nature of the records that are placed in an Enhanced Metafile (EMF) file.

```
EmfTypeEmfOnly = MetafileTypeEmf
EmfTypeEmfPlusOnly = MetafileTypeEmfPlusOnly
EmfTypeEmfPlusDual = MetafileTypeEmfPlusDual
```

| Constant   | Meaning     |
| ---------- | ----------- |
| EmfTypeEmfOnly | Specifies that all of the records in the metafile are EMF records, which can be displayed by GDI or GDI+. |
| EmfTypeEmfPlusOnly | Specifies that all of the records in the metafile are EMF+ records, which can be displayed by GDI+ but not by GDI. |
| EmfTypeEmfPlusDual | Specifies that all EMF+ records in the metafile are associated with an alternate EMF record. Metafiles of type **EmfTypeEmfPlusDual** can be displayed by GDI or by GDI+. |

# <a name="EncoderParameterValueType"></a>EncoderParameterValueType Enumeration

The **EncoderParameterValueType** enumeration specifies data types for image codec (encoder/decoder) parameters. An element of this enumeration is assigned to the Type data member of an EncoderParameter structure. 

```
EncoderParameterValueTypeByte = 1
EncoderParameterValueTypeASCII = 2
EncoderParameterValueTypeShort = 3
EncoderParameterValueTypeLong = 4
EncoderParameterValueTypeRational = 5
EncoderParameterValueTypeLongRange = 6
EncoderParameterValueTypeUndefined = 7
EncoderParameterValueTypeRationalRange = 8
EncoderParameterValueTypePointer = 9
```

| Constant   | Meaning     |
| ---------- | ----------- |
| EncoderParameterValueTypeByte | Specifies that the parameter is an 8-bit unsigned integer. |
| EncoderParameterValueTypeASCII | Specifies that the parameter is a null-terminated character string. |
| EncoderParameterValueTypeShort | Specifies that the parameter is a 16-bit unsigned integer. |
| EncoderParameterValueTypeLong | Specifies that the parameter is a 32-bit unsigned integer. |
| EncoderParameterValueTypeRational | Specifies that the parameter is an array of two, 32-bit unsigned integers. The pair of integers represents a fraction. The first integer in the pair is the numerator, and the second integer in the pair is the denominator. |
| EncoderParameterValueTypeLongRange | Specifies that the parameter is an array of two, 32-bit unsigned integers. The pair of integers represents a range of numbers. The first integer is the smallest number in the range, and the second integer is the largest number in the range. |
| EncoderParameterValueTypeUndefined | Specifies that the parameter is an array of bytes that can hold values of any type. |
| EncoderParameterValueTypeRationalRange | Specifies that the parameter is an array of four, 32-bit unsigned integers. The first two integers represent one fraction, and the second two integers represent a second fraction. The two fractions represent a range of rational numbers. The first fraction is the smallest rational number in the range, and the second fraction is the largest rational number in the range. |
| EncoderParameterValueTypePointer | Specifies that the parameter is a pointer to a block of custom metadata.|

# <a name="EncoderValue"></a>EncoderValue Enumeration

The **EncoderValue** enumeration specifies values that can be passed as arguments to image encoders.

```
EncoderValueColorTypeCMYK             = 0
EncoderValueColorTypeYCCK             = 1
EncoderValueCompressionLZW            = 2
EncoderValueCompressionCCITT3         = 3
EncoderValueCompressionCCITT4         = 4
EncoderValueCompressionRle            = 5
EncoderValueCompressionNone           = 6
EncoderValueScanMethodInterlaced      = 7
EncoderValueScanMethodNonInterlaced   = 8
EncoderValueVersionGif87              = 9
EncoderValueVersionGif89              = 10
EncoderValueRenderProgressive         = 11
EncoderValueRenderNonProgressive      = 12
EncoderValueTransformRotate90         = 13
EncoderValueTransformRotate180        = 14
EncoderValueTransformRotate270        = 15
EncoderValueTransformFlipHorizontal   = 16
EncoderValueTransformFlipVertical     = 17
EncoderValueMultiFrame                = 18
EncoderValueLastFrame                 = 19
EncoderValueFlush                     = 20
EncoderValueFrameDimensionTime        = 21
EncoderValueFrameDimensionResolution  = 22
EncoderValueFrameDimensionPage        = 23
EncoderValueColorTypeGray             = 24
EncoderValueColorTypeRGB              = 25
```

| Constant   | Meaning     |
| ---------- | ----------- |
| EncoderValueColorTypeCMYK | Not used in GDI+ version 1.0. |
| EncoderValueColorTypeYCCK | Not used in GDI+ version 1.0. |
| EncoderValueCompressionLZW | For a TIFF image, specifies the LZW compression method. |
| EncoderValueCompressionCCITT3 | For a TIFF image, specifies the CCITT3 compression method. |
| EncoderValueCompressionCCITT4 | For a TIFF image, specifies the CCITT4 compression method. |
| EncoderValueCompressionRle | For a TIFF image, specifies the RLE compression method. |
| EncoderValueCompressionNone | For a TIFF image, specifies no compression. |
| EncoderValueScanMethodInterlaced | Not used in GDI+ version 1.0. |
| EncoderValueScanMethodNonInterlaced | Not used in GDI+ version 1.0. |
| EncoderValueVersionGif87 | Not used in GDI+ version 1.0. |
| EncoderValueVersionGif89 | Not used in GDI+ version 1.0. |
| EncoderValueRenderProgressive | Not used in GDI+ version 1.0. |
| EncoderValueRenderNonProgressive | Not used in GDI+ version 1.0. |
| EncoderValueTransformRotate90 | For a JPEG image, specifies lossless 90-degree clockwise rotation. |
| EncoderValueTransformRotate180 | For a JPEG image, specifies lossless 180-degree rotation. |
| EncoderValueTransformRotate270 | For a JPEG image, specifies lossless 270-degree clockwise rotation. |
| EncoderValueTransformFlipHorizontal | For a JPEG image, specifies a lossless horizontal flip. |
| EncoderValueTransformFlipVertical | For a JPEG image, specifies a lossless vertical flip. |
| EncoderValueMultiFrame | Specifies multiple-frame encoding. |
| EncoderValueLastFrame | Specifies the last frame of a multiple-frame image. |
| EncoderValueFlush | Specifies that the encoder object is to be closed. |
| EncoderValueFrameDimensionTime | Not used in GDI+ version 1.0. |
| EncoderValueFrameDimensionResolution | Not used in GDI+ version 1.0. |
| EncoderValueFrameDimensionPage | For a TIFF image, specifies the page frame dimension |

# <a name="FillMode"></a>FillMode Enumeration

The **FillMode** enumeration specifies how to fill areas that are formed when a path or curve intersects itself.

```
FillModeAlternate = 0
FillModeWinding   = 1
```

| Constant   | Meaning     |
| ---------- | ----------- |
| FillModeAlternate | Specifies that areas are filled according to the even-odd parity rule. According to this rule, you can determine whether a test point is inside or outside a closed curve as follows: Draw a line from the test point to a point that is distant from the curve. If that line crosses the curve an odd number of times, the test point is inside the curve; otherwise, the test point is outside the curve. |
| FillModeWinding | Specifies that areas are filled according to the nonzero winding rule. According to this rule, you can determine whether a test point is inside or outside a closed curve as follows: Draw a line from a test point to a point that is distant from the curve. Count the number of times the curve crosses the test line from left to right, and count the number of times the curve crosses the test line from right to left. If those two numbers are the same, the test point is outside the curve; otherwise, the test point is inside the curve. |

# <a name="FlushIntention"></a>FlushIntention Enumeration

The **FlushIntention** enumeration specifies when to flush the queue of graphics operations.

```
FlushIntentionFlush = 0
FlushIntentionSync = 1
```

| Constant   | Meaning     |
| ---------- | ----------- |
| FlushIntentionFlush | When passed to the GdipFlush function, specifies that pending rendering operations are executed as soon as possible. The GdipFlush  function is not synchronized with the completion of the rendering operations and might return before the rendering operations are completed. |
| FlushIntentionSync | When passed to the GdipFlush function, specifies that pending rendering operations are executed as soon as possible. The **GdipFlush** function is synchronized with the completion of the rendering operations; that is, it will not return until after the rendering operations are completed. |

# <a name="FontStyle"></a>FontStyle Enumeration

The **FontStyle** enumeration specifies the style of the typeface of a font. Styles can be combined.

```
FontStyleRegular = 0
FontStyleBold = 1
FontStyleItalic = 2
FontStyleBoldItalic = 3
FontStyleUnderline = 4
FontStyleStrikeout = 8
```

| Constant   | Meaning     |
| ---------- | ----------- |
| FontStyleRegular | Specifies normal weight or thickness of the typeface. |
| FontStyleBold | Specifies bold typeface. Bold is a heavier weight or thickness. |
| FontStyleItalic | Specifies italic typeface, which produces a noticeable slant to the vertical stems of the characters. |
| FontStyleBoldItalic | Specifies the typeface as both bold and italic. |
| FontStyleUnderline | Specifies underline, which displays a line underneath the baseline of the characters. |
| FontStyleStrikeout | Specifies strikeout, which displays a horizontal line drawn through the middle of the characters. |

# <a name="GenericFontFamily"></a>GenericFontFamily Enumeration

```
GenericFontFamilySerif = 0
GenericFontFamilySansSerif = 1
GenericFontFamilyMonospace = 2
```

# <a name="HatchStyle"></a>HatchStyle Enumeration

The **HatchStyle** enumeration specifies the hatch pattern used by a brush of type **HatchBrush**. The hatch pattern consists of a solid background color and lines drawn over the background. 

```
HatchStyleHorizontal                   = 0
HatchStyleVertical                     = 1
HatchStyleForwardDiagonal              = 2
HatchStyleBackwardDiagonal             = 3
HatchStyleCross                        = 4
HatchStyleDiagonalCross                = 5
HatchStyle05Percent                    = 6
HatchStyle10Percent                    = 7
HatchStyle20Percent                    = 8
HatchStyle25Percent                    = 9
HatchStyle30Percent                    = 10
HatchStyle40Percent                    = 11
HatchStyle50Percent                    = 12
HatchStyle60Percent                    = 13
HatchStyle70Percent                    = 14
HatchStyle75Percent                    = 15
HatchStyle80Percent                    = 16
HatchStyle90Percent                    = 17
HatchStyleLightDownwardDiagonal        = 18
HatchStyleLightUpwardDiagonal          = 19
HatchStyleDarkDownwardDiagonal         = 20
HatchStyleDarkUpwardDiagonal           = 21
HatchStyleWideDownwardDiagonal         = 22
HatchStyleWideUpwardDiagonal           = 23
HatchStyleLightVertical                = 24
HatchStyleLightHorizontal              = 25
HatchStyleNarrowVertical               = 26
HatchStyleNarrowHorizontal             = 27
HatchStyleDarkVertical                 = 28
HatchStyleDarkHorizontal               = 29
HatchStyleDashedDownwardDiagonal       = 30
HatchStyleDashedUpwardDiagonal         = 31
HatchStyleDashedHorizontal             = 32
HatchStyleDashedVertical               = 33
HatchStyleSmallConfetti                = 34
HatchStyleLargeConfetti                = 35
HatchStyleZigZag                       = 36
HatchStyleWave                         = 37
HatchStyleDiagonalBrick                = 38
HatchStyleHorizontalBrick              = 39
HatchStyleWeave                        = 40
HatchStylePlaid                        = 41
HatchStyleDivot                        = 42
HatchStyleDottedGrid                   = 43
HatchStyleDottedDiamond                = 44
HatchStyleShingle                      = 45
HatchStyleTrellis                      = 46
HatchStyleSphere                       = 47
HatchStyleSmallGrid                    = 48
HatchStyleSmallCheckerBoard            = 49
HatchStyleLargeCheckerBoard            = 50
HatchStyleOutlinedDiamond              = 51
HatchStyleSolidDiamond                 = 52
HatchStyleTotal     = 53
HatchStyleLargeGrid = HatchStyleCross
HatchStyleMin       = HatchStyleHorizontal
HatchStyleMax       = HatchStyleTotal - 1
```

| Constant   | Meaning     |
| ---------- | ----------- |
| HatchStyleHorizontal | Specifies horizontal lines. |
| HatchStyleVertical | Specifies vertical lines. |
| HatchStyleForwardDiagonal | Specifies diagonal lines that slant to the right from top points to bottom points. The lines are antialiased. |
| HatchStyleBackwardDiagonal | Specifies diagonal lines that slant to the left from top points to bottom points. The lines are antialiased. |
| HatchStyleCross | Specifies horizontal and vertical lines that cross at 90-degree angles. |
| HatchStyleDiagonalCross | Specifies forward diagonal and backward diagonal lines that cross at 90-degree angles. The lines are antialiased. |
| HatchStyle05Percent | Specifies a 5-percent hatch. The ratio of foreground color to background color is 5:100. |
| HatchStyle10Percent | Specifies a 10-percent hatch. The ratio of foreground color to background color is 10:100. |
| HatchStyle20Percent | Specifies a 20-percent hatch. The ratio of foreground color to background color is 20:100. |
| HatchStyle25Percent | Specifies a 25-percent hatch. The ratio of foreground color to background color is 25:100. |
| HatchStyle30Percent | Specifies a 30-percent hatch. The ratio of foreground color to background color is 30:100. |
| HatchStyle40Percent | Specifies a 40-percent hatch. The ratio of foreground color to background color is 40:100. |
| HatchStyle50Percent | Specifies a 50-percent hatch. The ratio of foreground color to background color is 50:100. |
| HatchStyle60Percent | Specifies a 60-percent hatch. The ratio of foreground color to background color is 60:100. |
| HatchStyle70Percent | Specifies a 70-percent hatch. The ratio of foreground color to background color is 70:100. |
| HatchStyle75Percent | Specifies a 75-percent hatch. The ratio of foreground color to background color is 75:100. |
| HatchStyle80Percent | Specifies an 80-percent hatch. The ratio of foreground color to background color is 80:100. |
| HatchStyle90Percent | Specifies a 90-percent hatch. The ratio of foreground color to background color is 90:100. |
| HatchStyleLightDownwardDiagonal | Specifies diagonal lines that slant to the right from top points to bottom points and are spaced 50 percent closer together than **HatchStyleForwardDiagonal** but are not antialiased. | 
| HatchStyleLightUpwardDiagonal | Specifies diagonal lines that slant to the left from top points to bottom points and are spaced 50 percent closer together than **HatchStyleBackwardDiagonal** but are not antialiased. |
| HatchStyleDarkDownwardDiagonal | Specifies diagonal lines that slant to the right from top points to bottom points, are spaced 50 percent closer together than **HatchStyleForwardDiagonal**, and are twice the width of **HatchStyleForwardDiagonal** but are not antialiased. |
| HatchStyleDarkUpwardDiagonal | Specifies diagonal lines that slant to the left from top points to bottom points, are spaced 50 percent closer together than **HatchStyleBackwardDiagonal**, and are twice the width of **HatchStyleBackwardDiagonal** but are not antialiased. |
| HatchStyleWideDownwardDiagonal | Specifies diagonal lines that slant to the right from top points to bottom points, have the same spacing as **HatchStyleForwardDiagonal**, and are triple the width of HatchStyleForwardDiagonal but are not antialiased. |
| HatchStyleWideUpwardDiagonal | Specifies diagonal lines that slant to the left from top points to bottom points, have the same spacing as **HatchStyleBackwardDiagonal**, and are triple the width of **HatchStyleBackwardDiagonal** but are not antialiased. |
| HatchStyleLightVertical | Specifies vertical lines that are spaced 50 percent closer together than **HatchStyleVertical**. |
| HatchStyleLightHorizontal | Specifies horizontal lines that are spaced 50 percent closer together than **HatchStyleHorizontal**. |
| HatchStyleNarrowVertical | Specifies vertical lines that are spaced 75 percent closer together than **HatchStyleVertical** (or 25 percent closer together than **HatchStyleLightVertical**). |
| HatchStyleNarrowHorizontal | Specifies horizontal lines that are spaced 75 percent closer together than **HatchStyleHorizontal** ( or 25 percent closer together than **HatchStyleLightHorizontal**). | 
| HatchStyleDarkVertical | Specifies vertical lines that are spaced 50 percent closer together than **HatchStyleVerical** and are twice the width of **HatchStyleVertical**. |
| HatchStyleDarkHorizontal | Specifies horizontal lines that are spaced 50 percent closer together than **HatchStyleHorizontal** and are twice the width of **HatchStyleHorizontal**. |
| HatchStyleDashedDownwardDiagonal | Specifies horizontal lines that are composed of forward diagonals. |
| HatchStyleDashedUpwardDiagonal | Specifies horizontal lines that are composed of backward diagonals. |
| HatchStyleDashedHorizontal | Specifies horizontal dashed lines. |
| HatchStyleDashedVertical | Specifies vertical dashed lines. |
| HatchStyleSmallConfetti | Specifies a hatch that has the appearance of confetti. |
| HatchStyleLargeConfetti | Specifies a hatch that has the appearance of confetti composed of larger pieces than **HatchStyleSmallConfetti**. |
| HatchStyleZigZag | Specifies horizontal lines of zigzags. |
| HatchStyleWave | Specifies horizontal lines of tildes. |
| HatchStyleDiagonalBrick | Specifies a hatch that has the appearance of a wall of bricks laid in a backward diagonal direction. |
| HatchStyleHorizontalBrick | Specifies a hatch that has the appearance of a wall of bricks laid horizontally. |
| HatchStyleWeave | Specifies a hatch that has the appearance of a woven material. |
| HatchStylePlaid | Specifies a hatch that has the appearance of a plaid material. |
| HatchStyleDivot | Specifies a hatch that has the appearance of divots. |
| HatchStyleDottedGrid | Specifies horizontal and vertical dotted lines that cross at 90-degree angles. |
| HatchStyleDottedDiamond | Specifies forward diagonal and backward diagonal dotted lines that cross at 90-degree angles. |
| HatchStyleShingle | Specifies a hatch that has the appearance of shingles laid in a forward diagonal direction. |
| HatchStyleTrellis | Specifies a hatch that has the appearance of a trellis. |
| HatchStyleSphere | Specifies a hatch that has the appearance of a checkerboard of spheres. |
| HatchStyleSmallGrid | Specifies horizontal and vertical lines that cross at 90-degree angles and are spaced 50 percent closer together than **HatchStyleCross**. |
| HatchStyleSmallCheckerBoard | Specifies a hatch that has the appearance of a checkerboard. |
| HatchStyleLargeCheckerBoard | Specifies a hatch that has the appearance of a checkerboard with squares that are twice the size of **HatchStyleSmallCheckerBoard**. |
| HatchStyleOutlinedDiamond | Specifies forward diagonal and backward diagonal lines that cross at 90-degree angles but are not antialiased. |
| HatchStyleSolidDiamond | Specifies a hatch that has the appearance of a checkerboard placed diagonally. |
| HatchStyleTotal | Specifies no hatch thereby allowing the brush to be transparent. |
| HatchStyleLargeGrid | Specifies **HatchStyleCross**. |
| HatchStyleMin | Specifies **HatchStyleHorizonal**. |
| HatchStyleMax | Specifies **HatchStyleSolidDiamond**. |

# <a name="HistogramFormat"></a>HistogramFormat Enumeration

The **HistogramFormat** enumeration specifies the number and type of histograms that represent the color channels of a bitmap. This enumeration is used with the GdipBitmapGetHistogram function.

```
HistogramFormatARGB  = 0
HistogramFormatPARGB = 1
HistogramFormatRGB   = 2
HistogramFormatGray  = 3
HistogramFormatB     = 4
HistogramFormatG     = 5
HistogramFormatR     = 6
HistogramFormatA     = 7
```

| Constant   | Meaning     |
| ---------- | ----------- |
| HistogramFormatARGB | Specifies that the GdipBitmapGetHistogram function returns four histograms: one each for the alpha, red, green, and blue channels. The alpha-channel histogram is written to the buffer pointed to by the channel0 parameter of the **GdipBitmapGetHistogram** function. The red-channel histogram is written to the buffer pointed to by the channel1 parameter. The green-channel histogram is written to the buffer pointed to by the channel2 parameter. The blue-channel histogram is written to the buffer pointed to by the channel3 parameter. |
| HistogramFormatPARGB | Specifies that the GdipBitmapGetHistogram function returns four histograms: one each for the alpha, red, green, and blue channels. The red, green, and blue channels are each multiplied by the alpha channel before the histograms are created. The bitmap is not permanently altered when the color channels are multiplied by the alpha channel; that multiplication is only for the purpose of creating the histograms. The alpha-channel histogram is written to the buffer pointed to by the channel0 parameter of the GdipBitmapGetHistogram function. The red-channel histogram is written to the buffer pointed to by the channel1 parameter. The green-channel histogram is written to the buffer pointed to by the channel2 parameter. The blue-channel histogram is written to the buffer pointed to by the channel3 parameter. |
| HistogramFormatRGB | Specifies that the **GdipBitmapGetHistogram** function returns three histograms: one each for the red, green, and blue channels. The red-channel histogram is written to the buffer pointed to by the channel0 parameter of the GdipBitmapGetHistogram function. The green-channel histogram is written to the buffer pointed to by the channel1 parameter. The blue-channel histogram is written to the buffer pointed to by the channel2 parameter. The channel3 parameter must be set to NULL. |
| HistogramFormatGray | Specifies that each pixel is converted to a grayscale value in the range 0 through 255, and then one histogram, based on those grayscale value, is returned. The bitmap is not permanently altered by the conversion to grayscale values; those values are calculated only for the purpose of creating the histogram. The grayscale histogram is written to the buffer pointed to by the channel0 parameter of the **GdipBitmapGetHistogram** function. The channel1, channel2, and channel3 parameters must be set to NULL. |
| HistogramFormatB | Specifies that the **GdipBitmapGetHistogram** function returns a histogram for the blue channel. The blue-channel histogram is written to the buffer pointed to by the channel0 parameter of the **GdipBitmapGetHistogram** function. The channel1, channel2, and channel3 parameters must be set to NULL. |
| HistogramFormatG | Specifies that the **GdipBitmapGetHistogram** function returns a histogram for the green channel. The green-channel histogram is written to the buffer pointed to by the channel0 parameter of the **GdipBitmapGetHistogram** function. The channel1, channel2, and channel3 parameters must be set to NULL. |
| HistogramFormatR | Specifies that the **GdipBitmapGetHistogram** function returns a histogram for the red channel. The red-channel histogram is written to the buffer pointed to by the channel0 parameter of the **GdipBitmapGetHistogram** function. The channel1, channel2, and channel3 parameters must be set to NULL. |
| HistogramFormatA | Specifies that the **GdipBitmapGetHistogram** function returns a histogram for the alpha channel. The alpha-channel histogram is written to the buffer pointed to by the channel0 parameter of the **GdipBitmapGetHistogram** function. The channel1, channel2, and channel3 parameters must be set to NULL. |

# <a name="HotkeyPrefix"></a>HotkeyPrefix Enumeration

The **HotkeyPrefix** enumeration specifies how to display hot keys. There are three options: do nothing, display hot keys underlined, and hide the hot key underlines.

```
HotkeyPrefixNone = 0
HotkeyPrefixShow = 1
HotkeyPrefixHide = 2
```

| Constant   | Meaning     |
| ---------- | ----------- |
| HotkeyPrefixNone | Specifies that no hot key processing occurs. |
| HotkeyPrefixShow | Specifies that Unicode text is scanned for ampersands (&), which are interpreted as hot key markers in the same way as menu and dialog resources are processed in the Windows user interface. All pairs of ampersands are replaced by a single ampersand. All single ampersands are removed and the first character that follows the first single ampersand is displayed underlined. |
| HotkeyPrefixHide | Specifies that Unicode text is scanned for ampersands (&), which are substituted and removed as in HotkeyPrefixShow but no underline is shown. This option can be used when accessibility hot key underlines are not required. |

#### Remarks

Hot keys are keys that are programmed to provide keyboard shortcuts to functionality and are activated by pressing the ALT key. The keys are application dependent and are identified by an underscored letter, typically in a menu. An example would be the File menu with the letter F underscored. This makes the F key a hot key to launch the File menu.

# <a name="ImageFlags"></a>ImageFlags Enumeration

The **ImageFlags** enumeration specifies the attributes of the pixel data contained in an image. The GdipGetImageFlags function returns an element of this enumeration.

```
ImageFlagsNone                = 0
'// Low-word: shared with SINKFLAG_x
ImageFlagsScalable            = &H00000001
ImageFlagsHasAlpha            = &H00000002
ImageFlagsHasTranslucent      = &H00000004
ImageFlagsPartiallyScalable   = &H00000008
'// Low-word: color space definition
ImageFlagsColorSpaceRGB       = &H00000010
ImageFlagsColorSpaceCMYK      = &H00000020
ImageFlagsColorSpaceGRAY      = &H00000040
ImageFlagsColorSpaceYCBCR     = &H00000080
ImageFlagsColorSpaceYCCK      = &H00000100
'// Low-word: image size info
ImageFlagsHasRealDPI          = &H00001000
ImageFlagsHasRealPixelSize    = &H00002000
'// High-word
ImageFlagsReadOnly            = &H00010000
ImageFlagsCaching             = &H00020000
```

| Constant   | Meaning     |
| ---------- | ----------- |
| ImageFlagsNone | Specifies no format information. |
| ImageFlagsScalable | Specifies that the image can be scaled. |
| ImageFlagsHasAlpha | Specifies that the pixel data contains alpha values. |
| ImageFlagsHasTranslucent | Specifies that the pixel data has alpha values other than 0 (transparent) and 255 (opaque). |
| ImageFlagsPartiallyScalable | Specifies that the pixel data is partially scalable with some limitations. |
| ImageFlagsColorSpaceRGB | Specifies that the image is stored using an RGB color space. |
| ImageFlagsColorSpaceCMYK | Specifies that the image is stored using a CMYK color space. |
| ImageFlagsColorSpaceGRAY | Specifies that the image is a grayscale image. |
| ImageFlagsColorSpaceYCBCR | Specifies that the image is stored using a YCBCR color space. |
| ImageFlagsColorSpaceYCCK | Specifies that the image is stored using a YCCK color space. |
| ImageFlagsHasRealDPI | Specifies that dots per inch information is stored in the image. |
| ImageFlagsHasRealPixelSize | Specifies that the pixel size is stored in the image. |
| ImageFlagsReadOnly | Specifies that the pixel data is read-only. |
| ImageFlagsCaching | Specifies that the pixel data can be cached for faster access. |

# <a name="ImageLockMode"></a>ImageLockMode Enumeration

The **ImageLockMode** enumeration specifies flags that are passed to the flags parameter of the **GdipBitmapLockBits** function. The **GdipBitmapLockBits** function locks a portion of an image so that you can read or write the pixel data.

```
ImageLockModeRead         = &H0001
ImageLockModeWrite        = &H0002
ImageLockModeUserInputBuf = &H0004
```

| Constant   | Meaning     |
| ---------- | ----------- |
| ImageLockModeRead | Specifies that a portion of the image is locked for reading. |
| ImageLockModeWrite | Specifies that a portion of the image is locked for writing. |
| ImageLockModeUserInputBuf | Specifies that the buffer used for reading or writing pixel data is allocated by the user. If this flag is set, then the lockedBitmapData parameter of the **GdipBitmapLockBits** function serves as an input parameter (and possibly as an output parameter). If this flag is cleared, then the lockedBitmapData parameter serves only as an output parameter. |

# <a name="ImageType"></a>ImageType Enumeration

The **ImageType** enumeration indicates whether an image is a bitmap or a metafile. The **GdipGetImageType** function returns an element of this enumeration.

```
ImageTypeUnknown   = 0
ImageTypeBitmap    = 1
ImageTypeMetafile  = 2
```

| Constant   | Meaning     |
| ---------- | ----------- |
| ImageTypeUnknown | Indicates that the image type is not known. |
| ImageTypeBitmap | Indicates a bitmap image. |
| ImageTypeMetafile | Indicates a metafile image. | 

# <a name="InterpolationMode"></a>InterpolationMode Enumeration

The InterpolationMode enumeration specifies the algorithm that is used when images are scaled or rotated. This enumeration is used by the **GdipGetInterpolationMode** and **GdipSetInterpolationMode** functions of the Graphics functions.

```
InterpolationModeInvalid = QualityModeInvalid
InterpolationModeDefault = QualityModeDefault
InterpolationModeLowQuality = QualityModeLow
InterpolationModeHighQuality = QualityModeHigh
InterpolationModeBilinear = 3
InterpolationModeBicubic = 4
InterpolationModeNearestNeighbor = 5
InterpolationModeHighQualityBilinear = 6
InterpolationModeHighQualityBicubic = 7
```

| Constant   | Meaning     |
| ---------- | ----------- |
| InterpolationModeInvalid | Used internally.|
| InterpolationModeDefault | Specifies the default interpolation mode. |
| InterpolationModeLowQuality | Specifies a low-quality mode. |
| InterpolationModeHighQuality | Specifies a high-quality mode. |
| InterpolationModeBilinear | Specifies bilinear interpolation. No prefiltering is done. This mode is not suitable for shrinking an image below 50 percent of its original size. |
| InterpolationModeBicubic | Specifies bicubic interpolation. No prefiltering is done. This mode is not suitable for shrinking an image below 25 percent of its original size. |
| InterpolationModeNearestNeighbor | Specifies nearest-neighbor interpolation. |
| InterpolationModeHighQualityBilinear | Specifies high-quality, bilinear interpolation. Prefiltering is performed to ensure high-quality shrinking. |
| InterpolationModeHighQualityBicubic | Specifies high-quality, bicubic interpolation. Prefiltering is performed to ensure high-quality shrinking. This mode produces the highest quality transformed images. |

# <a name="ItemDataPosition"></a>ItemDataPosition Enumeration

The **ItemDataPosition** enumeration is used to specify the location of custom metadata in an image file. 

```
ItemDataPositionAfterHeader = 0
ItemDataPositionAfterPalette = 1
ItemDataPositionAfterBits = 2
```

| Constant   | Meaning     |
| ---------- | ----------- |
| ItemDataPositionAfterHeader | Specifies that custom metadata is stored after the file header. Valid for JPEG, PNG, and GIF. |
| ItemDataPositionAfterPalette | Specifies that custom metadata is stored after the palette. Valid for PNG. |
| ItemDataPositionAfterBits | Specifies that custom metadata is stored after the pixel data. Valid for GIF and PNG. |

#### Remarks

GDI+ supports custom metadata for JPEG, PNG, and GIF image files.

# <a name="LinearGradientMode"></a>LinearGradientMode Enumeration

The **LinearGradientMode** enumeration specifies the direction in which the change of color occurs for a linear gradient brush.

```
LinearGradientModeHorizontal         = 0
LinearGradientModeVertical           = 1
LinearGradientModeForwardDiagonal    = 2
LinearGradientModeBackwardDiagonal   = 3
```

| Constant   | Meaning     |
| ---------- | ----------- |
| LinearGradientModeHorizontal | Specifies the color to change in a horizontal direction from the left of the display to the right of the display. |
| LinearGradientModeVertical | Specifies the color to change in a vertical direction from the top of the display to the bottom of the display. |
| LinearGradientModeForwardDiagonal | Specifies the color to change in a forward diagonal direction from the upper-left corner to the lower-right corner of the display. |
| LinearGradientModeBackwardDiagonal | Specifies the color to change in a backward diagonal direction from the upper-right corner to the lower-left corner of the display. |

# <a name="LineCap"></a>LineCap Enumeration

The **LineCap** enumeration specifies the type of graphic shape to use on the end of a line drawn with a Microsoft Windows GDI+ pen. The cap can be a square, circle, triangle, arrowhead, custom, or masked (hidden). End caps can also "anchor" the line by centering the cap at the end of the line. 

```
LineCapFlat = 0
LineCapSquare = 1
LineCapRound = 2
LineCapTriangle = 3
LineCapNoAnchor = &H10
LineCapSquareAnchor = &H11
LineCapRoundAnchor = &H12
LineCapDiamondAnchor = &H13
LineCapArrowAnchor = &H14
LineCapCustom = &HFF
LineCapAnchorMask = &HF0
```

| Constant   | Meaning     |
| ---------- | ----------- |
| LineCapFlat | Specifies that the line ends at the last point. The end is squared off. |
| LineCapSquare | Specifies a square cap. The center of the square is the last point in the line. The height and width of the square are the line width. |
| LineCapRound | Specifies a circular cap. The center of the circle is the last point in the line. The diameter of the circle is the line width. |
| LineCapTriangle | Specifies a triangular cap. The base of the triangle is the last point in the line. The base of the triangle is the line width. |
| LineCapNoAnchor | Specifies that the line ends are not anchored. |
| LineCapSquareAnchor | Specifies that the line ends are anchored with a square. The center of the square is the last point in the line. The height and width of the square are the line width. |
| LineCapRoundAnchor | Specifies that the line ends are anchored with a circle. The center of the circle is at the last point in the line. The circle is wider than the line. |
| LineCapDiamondAnchor | Specifies that the line ends are anchored with a diamond (a square turned at 45 degrees). The center of the diamond is at the last point in the line. The diamond is wider than the line. |
| LineCapArrowAnchor | Specifies that the line ends are anchored with arrowheads. The arrowhead point is located at the last point in the line. The arrowhead is wider than the line. |
| LineCapCustom | Specifies that the line ends are made from a **CustomLineCap**. |

# <a name="LineJoin"></a>LineJoin Enumeration

The **LineJoin** enumeration specifies how to join two lines that are drawn by the same pen and whose ends meet. At the intersection of the two line ends, a line join makes the join look more continuous. 

```
LineJoinMiter = 0
LineJoinBevel = 1
LineJoinRound = 2
LineJoinMiterClipped = 3
```

| Constant   | Meaning     |
| ---------- | ----------- |
| LineJoinMiter | Specifies a mitered join. This produces a sharp corner or a clipped corner, depending on whether the length of the miter exceeds the miter limit. |
| LineJoinBevel | Specifies a beveled join. This produces a diagonal corner. |
| LineJoinRound | Specifies a circular join. This produces a smooth, circular arc between the lines. |
| LineJoinMiterClipped | Specifies a mitered join. This produces a sharp corner or a beveled corner, depending on whether the length of the miter exceeds the miter limit. |

#### Remarks

The miter length is the distance from the intersection of the line walls on the inside of the join to the intersection of the line walls outside of the join. The miter length can be large when the angle between two lines is small. The miter limit is the maximum allowed ratio of miter length to stroke width. The default value is 10.0f.

When using **LineJoinMiter** and the actual ratio exceeds the miter limit, the corner is clipped perpendicular to the miter at a distance from the inner corner that is the product of the miter limit and the pen width. 

When using **LineJoinMiterClipped** and the miter limit is exceeded, the join is drawn as if its type were LineJoinBevel; that is, when the line walls on the inside of the join meet, then a joining line is drawn between the line walls on the outside of the join.

# <a name="MatrixOrder"></a>MatrixOrder Enumeration

The **MatrixOrder** enumeration specifies the order of multiplication when a new matrix is multiplied by an existing matrix. 

```
MatrixOrderPrepend = 0
MatrixOrderAppend = 1
```

| Constant   | Meaning     |
| ---------- | ----------- |
| MatrixOrderPrepend | Specifies that the new matrix is on the left and the existing matrix is on the right. |
| MatrixOrderAppend | Specifies that the existing matrix is on the left and the new matrix is on the right. |

# <a name="MetafileFrameUnit"></a>MetafileFrameUnit Enumeration

The **MetafileFrameUnit** enumeration specifies the unit of measure for a metafile frame rectangle. 

```
MetafileFrameUnitPixel      = UnitPixel
MetafileFrameUnitPoint      = UnitPoint
MetafileFrameUnitInch       = UnitInch
MetafileFrameUnitDocument   = UnitDocument
MetafileFrameUnitMillimeter = UnitMillimeter
MetafileFrameUnitGdi        = 7 
```

| Constant   | Meaning     |
| ---------- | ----------- |
| MetafileFrameUnitPixel | Specifies that a unit is 1 pixel. |
| MetafileFrameUnitPoint | Specifies that a unit is 1 point. |
| MetafileFrameUnitInch | Specifies that a unit is 1 inch. |
| MetafileFrameUnitDocument | Specifies that a unit is 1/300 inch. |
| MetafileFrameUnitMillimeter | Specifies that a unit is 1 millimeter. |
| MetafileFrameUnitGdi | Specifies that a unit is 0.01 millimeter. This element is provided for compatibility with Microsoft Windows Graphics Device Interface (GDI). |

# <a name="MetafileType"></a>MetafileType Enumeration

The **MetafileType** enumeration specifies types of metafiles.

```
MetafileTypeInvalid           = 0
MetafileTypeWmf               = 1
MetafileTypeWmfPlaceable      = 2
MetafileTypeEmf               = 3
MetafileTypeEmfPlusOnly       = 4
MetafileTypeEmfPlusDual       = 5
```

| Constant   | Meaning     |
| ---------- | ----------- |
| MetafileTypeInvalid | Specifies a metafile format that is not recognized in Microsoft Windows GDI+. |
| MetafileTypeWmf | Specifies a Windows Metafile Format (WMF) file. Such a file contains only Windows Graphics Device Interface (GDI) records. |
| MetafileTypeWmfPlaceable | Specifies a WMF file that has a placeable metafile header in front of it. |
| MetafileTypeEmf | Specifies an Enhanced Metafile (EMF) file. Such a file contains only GDI records. |
| MetafileTypeEmfPlusOnly | Specifies an EMF+ file. Such a file contains only GDI+ records and must be displayed by using GDI+. Displaying the records using GDI may cause unpredictable results. |
| MetafileTypeEmfPlusDual | Specifies an EMF+ Dual file. Such a file contains GDI+ records along with alternative GDI records and can be displayed by using either GDI or GDI+. Displaying the records using GDI may cause some quality degradation. |

# <a name="ObjectType"></a>ObjectType Enumeration

The **ObjectType** enumeration indicates the object type value of an EMF+ record.

```
ObjectTypeInvalid          = 0
ObjectTypeBrush            = 1
ObjectTypePen              = 2
ObjectTypePath             = 3
ObjectTypeRegion           = 4
ObjectTypeImage            = 5
ObjectTypeFont             = 6
ObjectTypeStringFormat     = 7
ObjectTypeImageAttributes  = 8
ObjectTypeCustomLineCap    = 9
ObjectTypeMax = ObjectTypeCustomLineCap
ObjectTypeMin = ObjectTypeBrush
```

| Constant   | Meaning     |
| ---------- | ----------- |
| ObjectTypeInvalid | Object type is invalid. |
| ObjectTypeBrush | Object type is a brush. |
| ObjectTypePen | Object type is a pen. |
| ObjectTypePath | Object type is a path. |
| ObjectTypeRegion | Object type is a region. |
| ObjectTypeFont | Object type is a font. |
| ObjectTypeStringFormat | Object type is a string format. |
| ObjectTypeImageAttributes | Object type is an image attribute. |
| ObjectTypeCustomLineCap | Object type is a custom line cap. |
| ObjectTypeGraphics | Object type is graphics. |
| ObjectTypeMax | Maximum enumeration value. Currently, it is **ObjectTypeGraphics**. | 
| ObjectTypeMin | Minimum enumeration value. Currently, it is **ObjectTypeBrush**. |

#### Remarks

To determine whether the object type value of an EMF+ record is valid, call **ObjectTypeIsValid**.

# <a name="PaletteFlags"></a>PaletteFlags Enumeration

The **PaletteFlags** enumeration indicates attributes of the color data in a palette.

```
PaletteFlagsHasAlpha    = &H0001
PaletteFlagsGrayScale   = &H0002
PaletteFlagsHalftone    = &H0004
```

| Constant   | Meaning     |
| ---------- | ----------- |
| PaletteFlagsHasAlpha | Indicates that one or more of the palette entries contains alpha (transparency) information. |
| PaletteFlagsGrayScale | Indicates that the palette contains only grayscale entries. |
| PaletteFlagsHalftone | Indicates that the palette is the Microsoft Windows halftone palette. |

# <a name="PaletteType"></a>PaletteType Enumeration

The **PaletteType** enumeration is used by the **GdipInitializePalette** and **GdipBitmapConvertFormat** functions. The members of the enumeration identify several standard color palette formats.

```
PaletteTypeCustom = 0
PaletteTypeOptimal = 1
PaletteTypeFixedBW = 2
PaletteTypeFixedHalftone8 = 3
PaletteTypeFixedHalftone27 = 4
PaletteTypeFixedHalftone64 = 5
PaletteTypeFixedHalftone125 = 6
PaletteTypeFixedHalftone216 = 7
PaletteTypeFixedHalftone252 = 8
PaletteTypeFixedHalftone256 = 9
```

| Constant   | Meaning     |
| ---------- | ----------- |
| PaletteTypeCustom | An arbitrary custom palette provided by the caller. |
| PaletteTypeOptimal | An palette of colors that are optimal for a particular bitmap. To create an optimal palette, pass **PaletteTypeOptimal**, the number of colors you want in the palette, and the address of a Bitmap to the **GdipInitializePalette** function. |
| PaletteTypeFixedBW | A palette that has two colors. This palette type is suitable for bitmaps that store 1 bit per pixel. |
| PaletteTypeFixedHalftone8 | A palette based on two intensities each (off or full) for the red, green, and blue channels. Also contains the 16 colors of the system palette. Because all eight of the on/off combinations of red, green, and blue are already in the system palette, this palette is the same as the system palette. This palette type is suitable for bitmaps that store 4 bits per pixel. |
| PaletteTypeFixedHalftone27 | A palette based on three intensities each for the red, green, and blue channels. Also contains the 16 colors of the system palette. Eight of the 16 system palette colors are among the 27 three-intensity combinations of red, green, and blue, so the total number of colors in the palette is 35. If the palette also includes the transparent color, the total number of colors is 36. |
| PaletteTypeFixedHalftone64 | A palette based on four intensities each for the red, green, and blue channels. Also contains the 16 colors of the system palette. Eight of the 16 system palette colors are among the 64 four-intensity combinations of red, green, and blue, so the total number of colors in the palette is 72. If the palette also includes the transparent color, the total number of colors is 73. |
| PaletteTypeFixedHalftone125 | A palette based on five intensities each for the red, green, and blue channels. Also contains the 16 colors of the system palette. Eight of the 16 system palette colors are among the 125 five-intensity combinations of red, green, and blue, so the total number of colors in the palette is 133. If the palette also includes the transparent color, the total number of colors is 134. |
| PaletteTypeFixedHalftone216 | A palette based on six intensities each for the red, green, and blue channels. Also contains the 16 colors of the system palette. Eight of the 16 system palette colors are among the 216 six-intensity combinations of red, green, and blue, so the total number of colors in the palette is 224. If the palette also includes the transparent color, the total number of colors is 225. This palette is sometimes called the Windows halftone palette or the Web palette. |
| PaletteTypeFixedHalftone252 | A palette based on 6 intensities of red, 7 intensities of green, and 6 intensities of blue. The system palette is not included. The total number of colors is 252. If the palette also includes the transparent color, the total number of colors is 253. |
| PaletteTypeFixedHalftone256 | A palette based on 8 intensities of red, 8 intensities of green, and 4 intensities of blue. The system palette is not included. The total number of colors is 256. If the transparent color is included in this palette, it must replace one of the RGB combinations. |

# <a name="PathPointType"></a>PathPointType Enumeration

The **PathPointType** enumeration indicates point types and flags for the data points in a path. Bits 0 through 2 indicate the type of a point, and bits 3 through 7 hold a set of flags that specify attributes of a point.

```
PathPointTypeStart           = 0
PathPointTypeLine            = 1
PathPointTypeBezier          = 3
PathPointTypePathTypeMask    = &H07
PathPointTypeDashMode        = &H10
PathPointTypePathMarker      = &H20
PathPointTypeCloseSubpath    = &H80
PathPointTypeBezier3         = 3
```

| Constant   | Meaning     |
| ---------- | ----------- |
| PathPointTypeStart | Indicates that the point is the start of a figure. |
| PathPointTypeLine | Indicates that the point is one of the two endpoints of a line. |
| PathPointTypeBezier | Indicates that the point is an endpoint or control point of a cubic Bezier spline. |
| PathPointTypePathTypeMask | Masks all bits except for the three low-order bits, which indicate the point type. |
| PathPointTypePathDashMode | Not used. |
| PathPointTypePathMarker | Specifies that the point is a marker. |
| PathPointTypeCloseSubpath | Specifies that the point is the last point in a closed subpath (figure). |
| PathPointTypeBezier3 | Indicates that the point is an endpoint or control point of a cubic Bezier spline. |

# <a name="PenAlignment"></a>PenAlignment Enumeration

The **PenAlignment** enumeration specifies the alignment of a pen relative to the stroke that is being drawn.

```
PenAlignmentCenter = 0
PenAlignmentInset = 1
```

| Constant   | Meaning     |
| ---------- | ----------- |
| PenAlignmentCenter | Specifies that the pen is aligned on the center of the line that is drawn. |
| PenAlignmentInset | Specifies, when drawing a polygon, that the pen is aligned on the inside of the edge of the polygon. |

#### Remarks

If you set the alignment of a Pen object to **PenAlignmentInset**, you cannot use that pen to draw compound lines or triangular dash caps.

# <a name="PenType"></a>PenType Enumeration

The **PenType** enumeration indicates the type of pattern, texture, or gradient that a pen draws. 

```
PenTypeSolidColor = BrushTypeSolidColor
PenTypeHatchFill = BrushTypeHatchFill
PenTypeTextureFill = BrushTypeTextureFill
PenTypePathGradient = BrushTypePathGradient
PenTypeLinearGradient = BrushTypeLinearGradient
PenTypeUnknown = -1
```

| Constant   | Meaning     |
| ---------- | ----------- |
| PenTypeSolidColor | Indicates that the pen draws with a solid color. |
| PenTypeHatchFill | Indicates that the pen draws with a hatch pattern |
| PenTypeTextureFill | Indicates that the pen draws with a texture. |
| PenTypePathGradient | Indicates that the pen draws with a color gradient. |
| PenTypeLinearGradient | Indicates that the pen draws with a color gradient. |
| PenTypeUnknown | Indicates that the pen type is unknown. |

# <a name="PixelOffsetMode"></a>PixelOffsetMode Enumeration

The **PixelOffsetMode** enumeration specifies the pixel offset mode. This enumeration is used by the **GdipGetPixelOffsetMode** and **GdipSetPixelOffsetMode** methods of the Graphics class.

```
PixelOffsetModeInvalid = QualityModeInvalid
PixelOffsetModeDefault = QualityModeDefault
PixelOffsetModeHighSpeed = QualityModeLow
PixelOffsetModeHighQuality = QualityModeHigh
PixelOffsetModeNone = 3
PixelOffsetModeHalf = 4 
```

| Constant   | Meaning     |
| ---------- | ----------- |
| PixelOffsetModeInvalid | Used internally. |
| PixelOffsetModeDefault | Equivalent to **PixelOffsetModeNone**. |
| PixelOffsetModeHighSpeed | Equivalent to **PixelOffsetModeNone**. |
| PixelOffsetModeHighQuality | Equivalent to **PixelOffsetModeHalf**. |
| PixelOffsetModeNone | Indicates that pixel centers have integer coordinates. |
| PixelOffsetModeHalf | Indicates that pixel centers have coordinates that are half way between integer values. |

#### Remarks

Consider the pixel in the upper-left corner of an image with address (0, 0). With **PixelOffsetModeNone**, the pixel covers the area between 0.5 and 0.5 in both the x and y directions; that is, the pixel center is at (0, 0). With **PixelOffsetModeHalf**, the pixel covers the area between 0 and 1 in both the x and y directions; that is, the pixel center is at (0.5, 0.5).

# <a name="QualityMode"></a>QualityMode Enumeration

The **QualityMode** enumeration specifies the quality mode.

```
QualityModeInvalid = -1
QualityModeDefault = 0
QualityModeLow     = 1
QualityModeHigh    = 2
```

| Constant   | Meaning     |
| ---------- | ----------- |
| QualityModeInvalid | Invalid mode. |
| QualityModeDefault | Default quality mode. |
| QualityModeLow | Low quality mode. Best performance. |
| QualityModeHigh | High quality mode. Best rendering quality. |

# <a name="RotateFlipType"></a>RotateFlipType Enumeration

The **RotateFlipType** enumeration specifies the direction of an image's rotation and the axis used to flip the image.

```
RotateNoneFlipNone = 0
Rotate90FlipNone   = 1
Rotate180FlipNone  = 2
Rotate270FlipNone  = 3
RotateNoneFlipX    = 4
Rotate90FlipX      = 5
Rotate180FlipX     = 6
Rotate270FlipX     = 7
RotateNoneFlipY    = Rotate180FlipX
Rotate90FlipY      = Rotate270FlipX
Rotate180FlipY     = RotateNoneFlipX
Rotate270FlipY     = Rotate90FlipX
RotateNoneFlipXY   = Rotate180FlipNone
Rotate90FlipXY     = Rotate270FlipNone
Rotate180FlipXY    = RotateNoneFlipNone
Rotate270FlipXY    = Rotate90FlipNone
```

| Constant   | Meaning     |
| ---------- | ----------- |
| RotateNoneFlipNone | Specifies no rotation and no flipping. |
| Rotate90FlipNone | Specifies a 90-degree rotation without flipping. |
| Rotate180FlipNone | Specifies a 180-degree rotation without flipping. |
| Rotate270FlipNone | Specifies a 270-degree rotation without flipping. |
| RotateNoneFlipX | Specifies no rotation and a horizontal flip. |
| Rotate90FlipX | Specifies a 90-degree rotation followed by a horizontal flip. |
| Rotate180FlipX | Specifies a 180-degree rotation followed by a horizontal flip. |
| Rotate270FlipX | Specifies a 270-degree rotation followed by a horizontal flip. |
| RotateNoneFlipY | Specifies no rotation and a vertical flip. |
| Rotate90FlipY | Specifies a 90-degree rotation followed by a vertical flip. |
| Rotate180FlipY | Specifies a 180-degree rotation followed by a vertical flip. |
| Rotate270FlipY | Specifies a 270-degree rotation followed by a vertical flip. |
| RotateNoneFlipXY | Specifies no rotation, a horizontal flip, and then a vertical flip. |
| Rotate90FlipXY | Specifies a 90-degree rotation followed by a horizontal flip and then a vertical flip. |
| Rotate180FlipXY | Specifies a 180-degree rotation followed by a horizontal flip and then a vertical flip. |
| Rotate270FlipXY | Specifies a 270-degree rotation followed by a horizontal flip and then a vertical flip. |

# <a name="SmoothingMode"></a>SmoothingMode Enumeration

The **SmoothingMode** enumeration specifies the type of smoothing (antialiasing) that is applied to lines and curves. This enumeration is used by the **GdipGetSmoothingMode** and **GdipSetSmoothingMode** functions.

```
SmoothingModeInvalid = QualityModeInvalid
SmoothingModeDefault = QualityModeDefault
SmoothingModeHighSpeed = QualityModeLow
SmoothingModeHighQuality = QualityModeHigh
SmoothingModeNone = 3
SmoothingModeAntiAlias = 4
```

| Constant   | Meaning     |
| ---------- | ----------- |
| SmoothingModeInvalid | Reserved. |
| SmoothingModeDefault | Specifies that smoothing is not applied. |
| SmoothingModeHighSpeed | Specifies that smoothing is not applied. |
| SmoothingModeHighQuality | Specifies that smoothing is applied using an 8 X 4 box filter. |
| SmoothingModeNone | Specifies that smoothing is not applied. |
| SmoothingModeAntiAlias8x4 | Specifies that smoothing is applied using an 8 X 4 box filter. |
| SmoothingModeAntiAlias | Specifies that smoothing is applied using an 8 X 4 box filter. |
| SmoothingModeAntiAlias8x8 | Specifies that smoothing is applied using an 8 X 8 box filter. |

#### Remarks

Smoothing performed by an 8 X 4 box filter gives better results for nearly vertical lines than it does for nearly horizontal lines. Smoothing performed by an 8 X 8 box filter gives equally good results for nearly vertical and nearly horizontal lines. The 8x8 algorithm produces higher quality smoothing but is slower than the 8 X 4 algorithm.

# <a name="Status"></a>Status Enumeration

The **Status** enumeration indicates the result of a Microsoft Windows GDI+ function call.

```
Ok = 0
GenericError = 1
InvalidParameter = 2
OutOfMemory = 3
ObjectBusy = 4
InsufficientBuffer = 5
NotImplemented = 6
Win32Error = 7
WrongState = 8
Aborted = 9
FileNotFound = 10
ValueOverflow = 11
AccessDenied = 12
UnknownImageFormat = 13
FontFamilyNotFound = 14
FontStyleNotFound = 15
NotTrueTypeFont = 16
UnsupportedGdiplusVersion = 17
GdiplusNotInitialized = 18
PropertyNotFound = 19
PropertyNotSupported = 20
ProfileNotFound = 21
```

| Constant   | Meaning     |
| ---------- | ----------- |
| Ok | Indicates that the function call was successful. |
| GenericError | Indicates that there was an error on the function call, which is identified as something other than those defined by the other elements of this enumeration. |
| InvalidParameter | Indicates that one of the arguments passed to the function was not valid. |
| OutOfMemory | Indicates that the operating system is out of memory and could not allocate memory to process the function call. |
| ObjectBusy | Indicates that one of the arguments specified in the application programming interface (API) call is already in use in another thread. | 
| InsufficientBuffer | Indicates that a buffer specified as an argument in the API call is not large enough to hold the data to be received. | 
| NotImplemented | Indicates that the function is not implemented. |
| Win32Error | Indicates that the function generated a Microsoft Win32 error. |
| WrongState | Indicates that the object is in an invalid state to satisfy the API call. |
| Aborted | Indicates that the function was aborted. |
| FileNotFound | Indicates that the specified image file or metafile cannot be found. |
| ValueOverflow | Indicates that the function performed an arithmetic operation that produced a numeric overflow. |
| AccessDenied | Indicates that a write operation is not allowed on the specified file. |
| UnknownImageFormat | Indicates that the specified image file format is not known. |
| FontFamilyNotFound | Indicates that the specified font family cannot be found. Either the font family name is incorrect or the font family is not installed. |
| FontStyleNotFound | Indicates that the specified style is not available for the specified font family. |
| NotTrueTypeFont | Indicates that the font retrieved from an HDC or LOGFONT is not a TrueType font and cannot be used with GDI+. |
| UnsupportedGdiplusVersion | Indicates that the version of GDI+ that is installed on the system is incompatible with the version with which the application was compiled. |
| GdiplusNotInitialized | Indicates that the GDI+API is not in an initialized state. To function, all GDI+ objects require that GDI+ be in an initialized state. Initialize GDI+ by calling GdiplusStartup. |
| PropertyNotFound | Indicates that the specified property does not exist in the image. |
| PropertyNotSupported | Indicates that the specified property is not supported by the format of the image and, therefore, cannot be set. |
| ProfileNotFound | Indicates that the color profile required to save an image in CMYK format was not found. |

# <a name="StringAlignment"></a>StringAlignment Enumeration

The **StringAlignment** enumeration specifies how a string is aligned in reference to the bounding rectangle. A bounding rectangle is used to define the area in which the text displays.

```
StringAlignmentNear = 0
StringAlignmentCenter = 1
StringAlignmentFar = 2
```

| Constant   | Meaning     |
| ---------- | ----------- |
| StringAlignmentNear | Specifies that alignment is towards the origin of the bounding rectangle. May be used for alignment of characters along the line or for alignment of lines within the rectangle. For a right to left bounding rectangle (**StringFormatFlagsDirectionRightToLeft**), the origin is at the upper right. |
| StringAlignmentCenter | Specifies that alignment is centered between origin and extent (width) of the formatting rectangle. |
| StringAlignmentFar | Specifies that alignment is to the far extent (right side) of the formatting rectangle. |

# <a name="StringDigitSubstitute"></a>StringDigitSubstitute Enumeration

The **StringDigitSubstitute** enumeration specifies how to substitute digits in a string according to a user's locale or language.

```
StringDigitSubstituteUser = 0
StringDigitSubstituteNone = 1
StringDigitSubstituteNational = 2
StringDigitSubstituteTraditional = 3
```

| Constant   | Meaning     |
| ---------- | ----------- |
| StringDigitSubstituteUser | Specifies a user-defined substitution scheme. |
| StringDigitSubstituteNone | Specifies to disable substitutions. |
| StringDigitSubstituteNational | Specifies substitution digits that correspond with the official national language of the user's locale. |
| StringDigitSubstituteTraditional | Specifies substitution digits that correspond with the user's native script or language, which may be different from the official national language of the user's locale. |

# <a name="StringFormatFlags"></a>StringFormatFlags Enumeration

The **StringFormatFlags** enumeration specifies text layout information (such as orientation and clipping) and display manipulations (such as ellipsis insertion, digit substitution, and representation of characters that are not supported by a font).

```
StringFormatFlagsDirectionRightToLeft  = &h00000001
StringFormatFlagsDirectionVertical     = &h00000002
StringFormatFlagsNoFitBlackBox         = &h00000004
StringFormatFlagsDisplayFormatControl  = &h00000020
StringFormatFlagsNoFontFallback        = &h00000400
StringFormatFlagsMeasureTrailingSpaces = &h00000800
StringFormatFlagsNoWrap                = &h00001000
StringFormatFlagsLineLimit             = &h00002000
StringFormatFlagsNoClip                = &h00004000
```

| Constant   | Meaning     |
| ---------- | ----------- |
| StringFormatFlagsDirectionRightToLeft | Specifies that reading order is right to left. For horizontal text, characters are read from right to left. For vertical text, columns are read from right to left. By default, horizontal or vertical text is read from left to right. |
| StringFormatFlagsDirectionVertical | Specifies that individual lines of text are drawn vertically on the display device. By default, lines of text are horizontal, each new line below the previous line. |
| StringFormatFlagsNoFitBlackBox | Specifies that parts of characters are allowed to overhang the string's layout rectangle. By default, characters are first aligned inside the rectangle's boundaries, then any characters which still overhang the boundaries are repositioned to avoid any overhang and thereby avoid affecting pixels outside the layout rectangle. An italic, lowercase letter F ( f) is an example of a character that may have overhanging parts. Setting this flag ensures that the character aligns visually with the lines above and below but may cause parts of characters, which lie outside the layout rectangle, to be clipped or painted. |
| StringFormatFlagsDisplayFormatControl | Specifies that Unicode layout control characters are displayed with a representative character. |
| StringFormatFlagsNoFontFallback | Specifies that an alternate font is used for characters that are not supported in the requested font. By default, any missing characters are displayed with the "fonts missing" character, usually an open square. |
| StringFormatFlagsMeasureTrailingSpaces | Specifies that the space at the end of each line is included in a string measurement. By default, the boundary rectangle returned by the GdipMeasureString function excludes the space at the end of each line. Set this flag to include that space in the measurement. |
| StringFormatFlagsNoWrap | Specifies that the wrapping of text to the next line is disabled. NoWrap is implied when a point of origin is used instead of a layout rectangle. When drawing text within a rectangle, by default, text is broken at the last word boundary that is inside the rectangle's boundary and wrapped to the next line. |
| StringFormatFlagsLineLimit | Specifies that only entire lines are laid out in the layout rectangle. By default, layout continues until the end of the text or until no more lines are visible as a result of clipping, whichever comes first. The default settings allow the last line to be partially obscured by a layout rectangle that is not a whole multiple of the line height. To ensure that only whole lines are seen, set this flag and be careful to provide a layout rectangle at least as tall as the height of one line. |
| StringFormatFlagsNoClip | Specifies that characters overhanging the layout rectangle and text extending outside the layout rectangle are allowed to show. By default, all overhanging characters and text that extends outside the layout rectangle are clipped. Any trailing spaces (spaces that are at the end of a line) that extend outside the layout rectangle are clipped. Therefore, the setting of this flag will have an effect on a string measurement if trailing spaces are being included in the measurement. If clipping is enabled, trailing spaces that extend outside the layout rectangle are not included in the measurement. If clipping is disabled, all trailing spaces are included in the measurement, regardless of whether they are outside the layout rectangle. |

#### Remarks

Multiple flags set can produce combined effects: 

When both **StringFormatFlagsDirectionVertical** and **StringFormatFlagsDirectionRightToLeft** are set, individual lines of text are drawn vertically. The first line starts at the right edge of the layout rectangle; the second line of text is to the left of the first line, and so on. 

When **StringFormatFlagsDirectionVertical** is set and **StringFormatFlagsDirectionRightToLeft** is not set, individual lines of text are drawn vertically. The first line starts at the left edge of the layout rectangle; the second line of text is to the right of the first line. 

When **StringFormatFlagsDirectionRightToLeft** is set and **StringFormatFlagsDirectionVertical** is not set, the individual lines of text are horizontal and the reading order is from right to left. This setting does not change the order in which characters are displayed, it simply specifies the order in which characters can be read. 

The **StringFormatFlagsDirectionVertical** and **StringFormatFlagsDirectionRightToLeft** flags can affect string alignment.

The StringTrimming enumeration specifies how to trim characters from a string so that the string fits into a layout rectangle. The layout rectangle is used to position and size the display string.

# <a name="StringTrimming"></a>StringTrimming Enumeration

```
StringTrimmingNone = 0
StringTrimmingCharacter = 1
StringTrimmingWord = 2
StringTrimmingEllipsisCharacter = 3
StringTrimmingEllipsisWord = 4
StringTrimmingEllipsisPath = 5
```

| Constant   | Meaning     |
| ---------- | ----------- |
| StringTrimmingNone | Specifies that no trimming is done. |
| StringTrimmingCharacter | Specifies that the string is broken at the boundary of the last character that is inside the layout rectangle. This is the default. |
| StringTrimmingWord | Specifies that the string is broken at the boundary of the last word that is inside the layout rectangle. |
| StringTrimmingEllipsisCharacter | Specifies that the string is broken at the boundary of the last character that is inside the layout rectangle and an ellipsis (...) is inserted after the character. |
| StringTrimmingEllipsisWord | Specifies that the string is broken at the boundary of the last word that is inside the layout rectangle and an ellipsis (...) is inserted after the word. |
| StringTrimmingEllipsisPath | Specifies that the center is removed from the string and replaced by an ellipsis. The algorithm keeps as much of the last portion of the string as possible. |

#### Remarks

Trimming affects only the last visible or partly visible (due to clipping) line of text.

# <a name="TestControl"></a>TestControl Enumeration

```
TestControlForceBilinear = 0
TestControlNoICM = 1
TestControlGetBuildNumber = 2
```

# <a name="TextRenderingHint"></a>TextRenderingHint Enumeration

The enumeration specifies the process used to render text. The process affects the quality of the text.

```
TextRenderingHintSystemDefault            = 0
TextRenderingHintSingleBitPerPixelGridFit = 1
TextRenderingHintSingleBitPerPixel        = 2
TextRenderingHintAntiAliasGridFit         = 3
TextRenderingHintAntiAlias                = 4
TextRenderingHintClearTypeGridFit         = 5
```

| Constant   | Meaning     |
| ---------- | ----------- |
| TextRenderingHintSystemDefault | Specifies that a character is drawn using the currently selected system font smoothing mode (also called a rendering hint). |
| TextRenderingHintSingleBitPerPixelGridFit | Specifies that a character is drawn using its glyph bitmap and hinting to improve character appearance on stems and curvature. |
| TextRenderingHintSingleBitPerPixel | Specifies that a character is drawn using its glyph bitmap and no hinting. This results in better performance at the expense of quality. |
| TextRenderingHintAntiAliasGridFit | Specifies that a character is drawn using its antialiased glyph bitmap and hinting. This results in much better quality due to antialiasing at a higher performance cost. |
| TextRenderingHintAntiAlias | Specifies that a character is drawn using its antialiased glyph bitmap and no hinting. Stem width differences may be noticeable because hinting is turned off. |
| TextRenderingHintClearTypeGridFit | Specifies that a character is drawn using its glyph Microsoft ClearType bitmap and hinting. This type of text rendering cannot be used along with **CompositingModeSourceCopy**. Microsoft Windows XP and Windows Server 2003 only: **ClearType** rendering is supported only on Windows XP and Windows Server 2003. Therefore, **TextRenderingHintClearTypeGridFit** is ignored on other operating systems even though Windows GDI+ is supported on those operating systems. |

#### Remarks

The quality associated with each process varies according to the circumstances. **TextRenderingHintClearTypeGridFit** provides the best quality for most LCD monitors and relatively small font sizes. **TextRenderingHintAntiAlias** provides the best quality for rotated text. Generally, a process that produces higher quality text is slower than a process that produces lower quality text.

# <a name="Unit"></a>Unit Enumeration

The **Unit** enumeration specifies the unit of measure for a given data type.

```
UnitWorld      = 0
UnitDisplay    = 1
UnitPixel      = 2
UnitPoint      = 3
UnitInch       = 4
UnitDocument   = 5
UnitMillimeter = 6
```

| Constant   | Meaning     |
| ---------- | ----------- |
| UnitWorld | Specifies world coordinates, a nonphysical unit. |
| UnitDisplay | Specifies display units. For example, if the display device is a monitor, then the unit is 1 pixel. |
| UnitPixel | Specifies that a unit is 1 pixel. |
| UnitPoint | Specifies that a unit is 1 point or 1/72 inch. |
| UnitInch | Specifies that a unit is 1 inch. |
| UnitDocument | Specifies that a unit is 1/300 inch. |
| UnitMillimeter | Specifies that a unit is 1 millimeter. |

# <a name="WarpMode"></a>WarpMode Enumeration

The **WarpMode** enumeration specifies warp modes that can be used to transform images.

```
WarpModePerspective = 0
WarpModeBilinear    = 1
```

| Constant   | Meaning     |
| ---------- | ----------- |
| WarpModePerspective | Specifies the perspective warp mode. |
| WarpModeBilinear | Specifies the bilinear warp mode. |

# <a name="WrapMode"></a>WrapMode Enumeration

The **WrapMode** enumeration specifies how repeated copies of an image are used to tile an area.

```
WrapModeTile         = 0
WrapModeTileFlipX    = 1
WrapModeTileFlipY    = 2
WrapModeTileFlipXY   = 3
WrapModeClamp        = 4
```

| Constant   | Meaning     |
| ---------- | ----------- |
| WrapModeTile | Specifies tiling without flipping. |
| WrapModeTileFlipX | Species that tiles are flipped horizontally as you move from one tile to the next in a row. |
| WrapModeTileFlipY | Specifies that tiles are flipped vertically as you move from one tile to the next in a column. |
| WrapModeTileFlipXY | Specifies that tiles are flipped horizontally as you move along a row and flipped vertically as you move along a column. |
| WrapModeClamp | Specifies that no tiling takes place. |

# <a name="Constants1"></a>Image Encoder Cnstants

The **GdipSaveImageToFile**, **GdipSaveImageToStream**, **GdipSaveAdd** and **GdipSaveAddImage** functions receive an **EncoderParameters** structure that contains an array of **EncoderParameter** objects. Each **EncoderParameter** object has a globally unique identifier (GUID) data member that specifies the parameter category. The following constants represent GUIDs that specify the various parameter categories. 

```
EncoderCompression      = "{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"
EncoderColorDepth       = "{66087055-AD66-4C7C-9A18-38A2310B8337}"
EncoderScanMethod       = "{3A4E2661-3109-4E56-8536-42C156E7DCFA}"
EncoderVersion          = "{24D18C76-814A-41A4-BF53-1C219CCCF797}"
EncoderRenderMethod     = "{6D42C53A-229A-4825-8BB7-5C99E2B9A8B8}"
EncoderQuality          = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
EncoderTransformation   = "{8D0EB2D1-A58E-4EA8-AA14-108074B7B6F9}"
EncoderLuminanceTable   = "{EDB33BCE-0266-4A77-B904-27216099E717}"
EncoderChrominanceTable = "{F2E455DC-09B3-4316-8260-676ADA32481C}"
EncoderSaveFlag         = "{292266FC-AC40-47BF-8CFC-A85B89A655DE}"
```

# <a name="Constants2"></a>Image File Format Constants

The **GdipGetImageRawFormat** function returns a globally unique identifier (GUID) that indicates the file format of an image. The following constants represent the GUIDs that identify those file formats. 

```
ImageFormatUndefined = "{B96B3CA9-0728-11D3-9D7B-0000F81EF32E}"
ImageFormatMemoryBMP = "{B96B3CAA-0728-11D3-9D7B-0000F81EF32E}"
ImageFormatBMP       = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
ImageFormatEMF       = "{B96B3CAC-0728-11D3-9D7B-0000F81EF32E}"
ImageFormatWMF       = "{B96B3CAD-0728-11D3-9D7B-0000F81EF32E}"
ImageFormatJPEG      = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
ImageFormatPNG       = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
ImageFormatGIF       = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}"
ImageFormatTIFF      = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"
ImageFormatEXIF      = "{B96B3CB2-0728-11D3-9D7B-0000F81EF32E}"
ImageFormatIcon      = "{B96B3CB5-0728-11D3-9D7B-0000F81EF32E}"
```

# <a name="Constants3"></a>Image Frame Dimension Constants 

The Graphics Interchange Format (GIF) and Tagged Image File Format (TIFF) image file formats enable you to store multiple frames in a single image file. Multiple frames in a GIF file are used for animation, so the frames are said to be in the time dimension. Multiple frames in a TIFF file are typically used as separate pages, so the frames are said to be in the page dimension. 

The following constants represent globally unique identifiers (GUIDs) that identify the page and time dimensions. 

```
FrameDimensionTime       = "{6AEDBD6D-3FB5-418A-83A6-7F45229DC872}"
FrameDimensionResolution = "{84236F7B-3BD3-428F-8DAB-4EA1439CA315}"
FrameDimensionPage       = "{7462DC86-6180-4C7E-8E3F-EE7333A7A483}"
```

# <a name="Constants4"></a>Image Pixel Format Constants

The following constants specify various pixel formats used in bitmaps. 

```
PixelFormat1bppIndexed    = 196865
PixelFormat4bppIndexed    = 197634
PixelFormat8bppIndexed    = 198659
PixelFormat16bppGrayScale = 1052676
PixelFormat16bppRGB555    = 135173
PixelFormat16bppRGB565    = 135174
PixelFormat16bppARGB1555  = 397319
PixelFormat24bppRGB       = 137224
PixelFormat32bppRGB       = 139273
PixelFormat32bppARGB      = 2498570
PixelFormat32bppPARGB     = 925707
PixelFormat48bppRGB       = 1060876
PixelFormat64bppARGB      = 3424269
PixelFormat64bppPARGB     = 29622286
PixelFormatMax            = 15
```

| Constant   | Meaning     |
| ---------- | ----------- |
| PixelFormat1bppIndexed | Specifies that the format is 1 bit per pixel, indexed. |
| PixelFormat4bppIndexed | Specifies that the format is 4 bits per pixel, indexed. |
| PixelFormat8bppIndexed | Specifies that the format is 8 bits per pixel, indexed. |
| PixelFormat16bppARGB1555  | Specifies that the format is 16 bits per pixel; 1 bit is used for the alpha component, and 5 bits each are used for the red, green, and blue components. |
| PixelFormat16bppGrayScale | Specifies that the format is 16 bits per pixel, grayscale. 
| PixelFormat16bppRGB555 | Specifies that the format is 16 bits per pixel; 5 bits each are used for the red, green, and blue components. The remaining bit is not used. |
| PixelFormat16bppRGB565 | Specifies that the format is 16 bits per pixel; 5 bits are used for the red component, 6 bits are used for the green component, and 5 bits are used for the blue component. |
| PixelFormat24bppRGB | Specifies that the format is 24 bits per pixel; 8 bits each are used for the red, green, and blue components. |
| PixelFormat32bppARGB | Specifies that the format is 32 bits per pixel; 8 bits each are used for the alpha, red, green, and blue components. |
| PixelFormat32bppPARGB | Specifies that the format is 32 bits per pixel; 8 bits each are used for the alpha, red, green, and blue components. The red, green, and blue components are premultiplied according to the alpha component. |
| PixelFormat32bppRGB | Specifies that the format is 32 bits per pixel; 8 bits each are used for the red, green, and blue components. The remaining 8 bits are not used. |
| PixelFormat48bppRGB | Specifies that the format is 48 bits per pixel; 16 bits each are used for the red, green, and blue components. |
| PixelFormat64bppARGB | Specifies that the format is 64 bits per pixel; 16 bits each are used for the alpha, red, green, and blue components. |
| PixelFormat64bppPARGB | Specifies that the format is 64 bits per pixel; 16 bits each are used for the alpha, red, green, and blue components. The red, green, and blue components are premultiplied according to the alpha component. |

#### Remarks

**PixelFormat48bppRGB**, **PixelFormat64bppARGB**, and **PixelFormat64bppPARGB** use 16 bits per color component (channel). Microsoft Windows GDI+ version 1.0 can read 16-bits-per-channel images, but such images are converted to an 8-bits-per-channel format for processing, displaying, and saving.

# <a name="Constants5"></a>Image Property Tag Type Constants

You can store and retrieve image metadata with the help of a **PropertyItem** structure. The type data member of a **PropertyItem**  structure specifies the data type of the values stored in the value data member of that same PropertyItem structure.

The following constants can be assigned to the type data member of a **PropertyItem** structure.

```
PropertyTagTypeByte       = 1
PropertyTagTypeASCII      = 2
PropertyTagTypeShort      = 3
PropertyTagTypeLong       = 4
PropertyTagTypeRational   = 5
PropertyTagTypeUndefined  = 7
PropertyTagTypeSLONG      = 9
PropertyTagTypeSRational  = 10
```

| Constant   | Meaning     |
| ---------- | ----------- |
| PixelFormat4bppIndexed  | Specifies that the format is 4 bits per pixel, indexed. |
| PropertyTagTypeASCII | Specifies that the value data member is a null-terminated ASCII string. If you set the type data member of a **PropertyItem** object to **PropertyTagTypeASCII**, you should set the length data member to the length of the string including the NULL terminator. For example, the string HELLO would have a length of 6. |
| PropertyTagTypeByte | Specifies that the value data member is an array of bytes. |
| PropertyTagTypeLong | Specifies that the value data member is an array of unsigned long (32-bit) integers. |
| PropertyTagTypeRational | Specifies that the value data member is an array of pairs of unsigned long integers. Each pair represents a fraction; the first integer is the numerator and the second integer is the denominator. |
| PropertyTagTypeShort | Specifies that the value data member is an array of unsigned short (16-bit) integers. |
| PropertyTagTypeSLONG | Specifies that the value data member is an array of signed long (32-bit) integers. |
| PropertyTagTypeSRational | Specifies that the value data member is an array of pairs of signed long integers. Each pair represents a fraction; the first integer is the numerator and the second integer is the denominator. |
| PropertyTagTypeUndefined | Specifies that the value data member is an array of bytes that can hold values of any data type. |

# <a name="Constants6"></a>Image Property Tag Constants

Several image file formats enable you to store metadata along with an image. Metadata is information about an image, for example, title, width, camera model, and artist. Microsoft Windows GDI+ provides a uniform way of storing and retrieving metadata from image files in the Tagged Image File Format (TIFF), JPEG, Portable Network Graphics (PNG), and Exchangeable Image File (EXIF) formats.

In GDI+, a piece of metadata is called a property item, and an individual property item is identified by a numerical constant called a property tag. You can store and retrieve metadata by calling the **GdipSetPropertyItem** and **GdipGetPropertyItem** functions, and you don't have to be concerned with the details of how a particular file format stores that metadata.

The following table lists the Microsoft Windows GDI+ image property tags in numerical order.

| Property tag | Value |
| ------------ | ----- |
| PropertyTagGpsVer | &h0000 |
| PropertyTagGpsLatitudeRef | &h0001 |
| PropertyTagGpsLatitude | &h0002 |
| PropertyTagGpsLongitudeRef | &h0003 |
| PropertyTagGpsLongitude | &h0004 |
| PropertyTagGpsAltitudeRef | &h0005 |
| PropertyTagGpsAltitude | &h0006 |
| PropertyTagGpsGpsTime | &h0007 |
| PropertyTagGpsGpsSatellites | &h0008 |
| PropertyTagGpsGpsStatus | &h0009 |
| PropertyTagGpsGpsMeasureMode | &h000A |
| PropertyTagGpsGpsDop | &h000B |
| PropertyTagGpsSpeedRef | &h000C |
| PropertyTagGpsSpeed | &h000D |
| PropertyTagGpsTrackRef | &h000E |
| PropertyTagGpsTrack | &h000F |
| PropertyTagGpsImgDirRef | &h0010 |
| PropertyTagGpsImgDir | &h0011 |
| PropertyTagGpsMapDatum | &h0012 |
| PropertyTagGpsDestLatRef | &h0013 |
| PropertyTagGpsDestLat | &h0014 |
| PropertyTagGpsDestLongRef | &h0015 |
| PropertyTagGpsDestLong | &h0016 |
| PropertyTagGpsDestBearRef | &h0017 |
| PropertyTagGpsDestBear | &h0018 |
| PropertyTagGpsDestDistRef | &h0019 |
| PropertyTagGpsDestDist | &h001A |
| PropertyTagNewSubfileType | &h00FE |
| PropertyTagSubfileType | &h00FF |
| PropertyTagImageWidth | &h0100 |
| PropertyTagImageHeight | &h0101 |
| PropertyTagBitsPerSample | &h0102 |
| PropertyTagCompression | &h0103 |
| PropertyTagPhotometricInterp | &h0106 |
| PropertyTagThreshHolding | &h0107 |
| PropertyTagCellWidth | &h0108 |
| PropertyTagCellHeight | &h0109 |
| PropertyTagFillOrder | &h010A |
| PropertyTagDocumentName | &h010D |
| PropertyTagImageDescription | &h010E |
| PropertyTagEquipMake | &h010F |
| PropertyTagEquipModel | &h0110 |
| PropertyTagStripOffsets | &h0111 |
| PropertyTagOrientation | &h0112 |
| PropertyTagSamplesPerPixel | &h0115 |
| PropertyTagRowsPerStrip | &h0116 |
| PropertyTagStripBytesCount | &h0117 |
| PropertyTagMinSampleValue | &h0118 |
| PropertyTagMaxSampleValue | &h0119 |
| PropertyTagXResolution | &h011A |
| PropertyTagYResolution | &h011B |
| PropertyTagPlanarConfig | &h011C |
| PropertyTagPageName | &h011D |
| PropertyTagXPosition | &h011E |
| PropertyTagYPosition | &h011F |
| PropertyTagFreeOffset | &h0120 |
| PropertyTagFreeByteCounts | &h0121 |
| PropertyTagGrayResponseUnit | &h0122 |
| PropertyTagGrayResponseCurve | &h0123 |
| PropertyTagT4Option | &h0124 |
| PropertyTagT6Option | &h0125 |
| PropertyTagResolutionUnit | &h0128 |
| PropertyTagPageNumber | &h0129 |
| PropertyTagTransferFunction | &h012D |
| PropertyTagSoftwareUsed | &h0131 |
| PropertyTagDateTime | &h0132 |
| PropertyTagArtist | &h013B |
| PropertyTagHostComputer | &h013C |
| PropertyTagPredictor | &h013D |
| PropertyTagWhitePoint | &h013E |
| PropertyTagPrimaryChromaticities | &h013F |
| PropertyTagColorMap | &h0140 |
| PropertyTagHalftoneHints | &h0141 |
| PropertyTagTileWidth | &h0142 |
| PropertyTagTileLength | &h0143 |
| PropertyTagTileOffset | &h0144 |
| PropertyTagTileByteCounts | &h0145 |
| PropertyTagInkSet | &h014C |
| PropertyTagInkNames | &h014D |
| PropertyTagNumberOfInks | &h014E |
| PropertyTagDotRange | &h0150 |
| PropertyTagTargetPrinter | &h0151 |
| PropertyTagExtraSamples | &h0152 |
| PropertyTagSampleFormat | &h0153 |
| PropertyTagSMinSampleValue | &h0154 |
| PropertyTagSMaxSampleValue | &h0155 |
| PropertyTagTransferRange | &h0156 |
| PropertyTagJPEGProc | &h0200 |
| PropertyTagJPEGInterFormat | &h0201 |
| PropertyTagJPEGInterLength | &h0202 |
| PropertyTagJPEGRestartInterval | &h0203 |
| PropertyTagJPEGLosslessPredictors | &h0205 |
| PropertyTagJPEGPointTransforms | &h0206 |
| PropertyTagJPEGQTables | &h0207 |
| PropertyTagJPEGDCTables | &h0208 |
| PropertyTagJPEGACTables | &h0209 |
| PropertyTagYCbCrCoefficients | &h0211 |
| PropertyTagYCbCrSubsampling | &h0212 |
| PropertyTagYCbCrPositioning | &h0213 |
| PropertyTagREFBlackWhite | &h0214 |
| PropertyTagGamma | &h0301 |
| PropertyTagICCProfileDescriptor | &h0302 |
| PropertyTagSRGBRenderingIntent | &h0303 |
| PropertyTagImageTitle | &h0320 |
| PropertyTagResolutionXUnit | &h5001 |
| PropertyTagResolutionYUnit | &h5002 |
| PropertyTagResolutionXLengthUnit | &h5003 |
| PropertyTagResolutionYLengthUnit | &h5004 |
| PropertyTagPrintFlags | &h5005 |
| PropertyTagPrintFlagsVersion | &h5006 |
| PropertyTagPrintFlagsCrop | &h5007 |
| PropertyTagPrintFlagsBleedWidth | &h5008 |
| PropertyTagPrintFlagsBleedWidthScale | &h5009 |
| PropertyTagHalftoneLPI | &h500A |
| PropertyTagHalftoneLPIUnit | &h500B |
| PropertyTagHalftoneDegree | &h500C |
| PropertyTagHalftoneShape | &h500D |
| PropertyTagHalftoneMisc | &h500E |
| PropertyTagHalftoneScreen | &h500F |
| PropertyTagJPEGQuality | &h5010 |
| PropertyTagGridSize | &h5011 |
| PropertyTagThumbnailFormat | &h5012 |
| PropertyTagThumbnailWidth | &h5013 |
| PropertyTagThumbnailHeight | &h5014 |
| PropertyTagThumbnailColorDepth | &h5015 |
| PropertyTagThumbnailPlanes | &h5016 |
| PropertyTagThumbnailRawBytes | &h5017 |
| PropertyTagThumbnailSize | &h5018 |
| PropertyTagThumbnailCompressedSize | &h5019 |
| PropertyTagColorTransferFunction | &h501A |
| PropertyTagThumbnailData | &h501B |
| PropertyTagThumbnailImageWidth | &h5020 |
| PropertyTagThumbnailImageHeight | &h5021 |
| PropertyTagThumbnailBitsPerSample | &h5022 |
| PropertyTagThumbnailCompression | &h5023 |
| PropertyTagThumbnailPhotometricInterp | &h5024 |
| PropertyTagThumbnailImageDescription | &h5025 |
| PropertyTagThumbnailEquipMake | &h5026 |
| PropertyTagThumbnailEquipModel | &h5027 |
| PropertyTagThumbnailStripOffsets | &h5028 |
| PropertyTagThumbnailOrientation | &h5029 |
| PropertyTagThumbnailSamplesPerPixel | &h502A |
| PropertyTagThumbnailRowsPerStrip | &h502B |
| PropertyTagThumbnailStripBytesCount | &h502C |
| PropertyTagThumbnailResolutionX | &h502D |
| PropertyTagThumbnailResolutionY | &h502E |
| PropertyTagThumbnailPlanarConfig | &h502F |
| PropertyTagThumbnailResolutionUnit | &h5030 |
| PropertyTagThumbnailTransferFunction | &h5031 |
| PropertyTagThumbnailSoftwareUsed | &h5032 |
| PropertyTagThumbnailDateTime | &h5033 |
| PropertyTagThumbnailArtist | &h5034 |
| PropertyTagThumbnailWhitePoint | &h5035  |
| PropertyTagThumbnailPrimaryChromaticities | &h5036 |
| PropertyTagThumbnailYCbCrCoefficients | &h5037 |
| PropertyTagThumbnailYCbCrSubsampling | &h5038 |
| PropertyTagThumbnailYCbCrPositioning | &h5039 |
| PropertyTagThumbnailRefBlackWhite | &h503A |
| PropertyTagThumbnailCopyRight | &h503B |
| PropertyTagLuminanceTable | &h5090 |
| PropertyTagChrominanceTable | &h5091 |
| PropertyTagFrameDelay | &h5100 |
| PropertyTagLoopCount | &h5101 |
| PropertyTagGlobalPalette | &h5102 |
| PropertyTagIndexBackground | &h5103 |
| PropertyTagIndexTransparent | &h5104 |
| PropertyTagPixelUnit | &h5110 |
| PropertyTagPixelPerUnitX | &h5111 |
| PropertyTagPixelPerUnitY | &h5112 |
| PropertyTagPaletteHistogram | &h5113 |
| PropertyTagCopyright | &h8298 |
| PropertyTagExifExposureTime | &h829A |
| PropertyTagExifFNumber | &h829D |
| PropertyTagExifIFD | &h8769 |
| PropertyTagICCProfile | &h8773 |
| PropertyTagExifExposureProg | &h8822 |
| PropertyTagExifSpectralSense | &h8824 |
| PropertyTagGpsIFD | &h8825 |
| PropertyTagExifISOSpeed | &h8827 |
| PropertyTagExifOECF | &h8828 |
| PropertyTagExifVer | &h9000 |
| PropertyTagExifDTOrig | &h9003 |
| PropertyTagExifDTDigitized | &h9004 |
| PropertyTagExifCompConfig | &h9101 |
| PropertyTagExifCompBPP | &h9102 |
| PropertyTagExifShutterSpeed | &h9201 |
| PropertyTagExifAperture | &h9202 |
| PropertyTagExifBrightness | &h9203 |
| PropertyTagExifExposureBias | &h9204 |
| PropertyTagExifMaxAperture | &h9205 |
| PropertyTagExifSubjectDist | &h9206 |
| PropertyTagExifMeteringMode | &h9207 |
| PropertyTagExifLightSource | &h9208 |
| PropertyTagExifFlash | &h9209 |
| PropertyTagExifFocalLength | &h920A |
| PropertyTagExifMakerNote | &h927C |
| PropertyTagExifUserComment | &h9286 |
| PropertyTagExifDTSubsec | &h9290 |
| PropertyTagExifDTOrigSS | &h9291 |
| PropertyTagExifDTDigSS | &h9292 |
| PropertyTagExifFPXVer | &hA000 |
| PropertyTagExifColorSpace | &hA001 |
| PropertyTagExifPixXDim | &hA002 |
| PropertyTagExifPixYDim | &hA003 |
| PropertyTagExifRelatedWav | &hA004 |
| PropertyTagExifInterop | &hA005 |
| PropertyTagExifFlashEnergy | &hA20B |
| PropertyTagExifSpatialFR | &hA20C |
| PropertyTagExifFocalXRes | &hA20E |
| PropertyTagExifFocalYRes | &hA20F |
| PropertyTagExifFocalResUnit | &hA210 |
| PropertyTagExifSubjectLoc | &hA214 |
| PropertyTagExifExposureIndex | &hA215 |
| PropertyTagExifSensingMethod | &hA217 |
| PropertyTagExifFileSource | &hA300 |
| PropertyTagExifSceneType | &hA301 |
| PropertyTagExifCfaPattern | &hA302 |

**See also**:

[Property Tags in Alphabetical Order](https://docs.microsoft.com/en-us/windows/desktop/gdiplus/-gdiplus-constant-property-tags-in-alphabetical-order)

[Property Item Descriptions](https://docs.microsoft.com/en-us/windows/desktop/gdiplus/-gdiplus-constant-property-item-descriptions)

[Image File Format Specification](https://docs.microsoft.com/en-us/windows/desktop/gdiplus/-gdiplus-constant-image-file-format-specifications)
