# CGpImage Class

The **CGpImage** class provides methods for loading and saving raster images (bitmaps) and vector images (metafiles). An Image object encapsulates a bitmap or a metafile and stores attributes that you can retrieve by calling various Get methods. You can construct **Image** objects from a variety of file types including BMP, ICON, GIF, JPEG, Exif, PNG, TIFF, WMF, and EMF.

**Inherits from**: CGpBase.<br>
**Include file**: CGpBitmap.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorsImage) | Creates an **Image** object based on a file or stream. |
| [Clone](#CloneImage) | Copies the contents of the existing **Image** object into a new Image object. |
| [FindFirstItem](#FindFirstItem) | Retrieves the description and the data size of the first metadata item in this Image object. |
| [FindNextItem](#FindNextItem) | Retrieves the description and the data size of the next metadata item in this Image object. |
| [GetAllPropertyItems](#GetAllPropertyItems) | Gets all the property items (metadata) stored in this Image object. |
| [GetBounds](#GetBounds) | Gets the bounding rectangle for this image. |
| [GetEncoderParameterList](#GetEncoderParameterList) | Gets a list of the parameters supported by a specified image encoder. |
| [GetEncoderParameterListSize](#GetEncoderParameterListSize) | Gets the size, in bytes, of the parameter list for a specified image encoder. |
| [GetFlags](#GetFlags) | Gets a set of flags that indicate certain attributes of this Image object. |
| [GetFrameCount](#GetFrameCount) | Gets the number of frames in a specified dimension of this Image object. |
| [GetFrameDimensionsCount](#GetFrameDimensionsCount) | Gets the number of frame dimensions in this Image object. |
| [GetFrameDimensionsList](#GetFrameDimensionsList) | Gets the identifiers for the frame dimensions of this Image object. |
| [GetHeight](#GetHeight) | Gets the image height, in pixels, of this image. |
| [GetHorizontalResolution](#GetHorizontalResolution) | Gets the horizontal resolution, in dots per inch, of this image. |
| [GetItemData](#GetItemData) | Gets one piece of metadata from this Image object. |
| [GetPalette](#GetPalette) | Gets the **ColorPalette** of this Image object. |
| [GetPaletteSize](#GetPaletteSize) | Gets the size, in bytes, of the color palette of this Image object. |
| [GetPhysicalDimension](#GetPhysicalDimension) | Gets the width and height of this image. |
| [GetPixelFormat](#GetPixelFormat) | Gets the pixel format of this Image object. |
| [GetPropertyCount](#GetPropertyCount) | Gets the number of properties (pieces of metadata) stored in this Image object. |
| [GetPropertyIdList](#GetPropertyIdList) | Gets a list of the property identifiers used in the metadata of this Image object. |
| [GetPropertyItem](#GetPropertyItem) | Gets a specified property item (piece of metadata) from this Image object. |
| [GetPropertyItemSize](#GetPropertyItemSize) | Gets the size, in bytes, of a specified property item of this Image object. |
| [GetPropertySize](#GetPropertySize) | Gets the total size, in bytes, of all the property items stored in this Image object. |
| [GetRawFormat](#GetRawFormat) | Gets a globally unique identifier ( GUID) that identifies the format of this Image object. |
| [GetThumbnailImage](#GetThumbnailImage) | Gets a thumbnail image from this Image object. |
| [GetType](#GetType) | Gets the type (bitmap or metafile) of this Image object. |
| [GetVerticalResolution](#GetVerticalResolution) | Gets the vertical resolution, in dots per inch, of this image. |
| [GetWidth](#GetWidth) | Gets the width, in pixels, of this image. |
| [RemovePropertyItem](#RemovePropertyItem) | Removes a property item (piece of metadata) from this Image object. |
| [RotateFlip](#RotateFlip) | Rotates and flips this image. |
| [Save](#Save) | Saves this image to a file. |
| [SaveAdd](#SaveAdd) | Adds a frame to a file or stream specified in a previous call to the Save method. |
| [SelectActiveFrame](#SelectActiveFrame) | Selects the frame in this Image object specified by a dimension and an index. |
| [SetPalette](#SetPalette) | Sets the color palette of this Image object. |
| [SetPropertyItem](#SetPropertyItem) | Sets a property item (piece of metadata) for this Image object. |

# CGpBitmap Class

Extends the **CGpImage** class. The **Bitmap** object expands on the capabilities of the **Image** object by providing additional methods for creating and manipulating raster images.

**Inherits from**: CGpImage.<br>
**Include file**: CGpBitmap.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorsBitmap) | Creates a **Bitmap** object based on an icon or resource file. |
| [Clone](#CloneBitmap) | Creates a new Bitmap object by copying a portion of this bitmap. |
| [ConvertFormat](#ConvertFormat) | Converts a bitmap to a specified pixel format. |
| [GetHBITMAP](#GetHBITMAP) | Creates a Windows Graphics Device Interface (GDI) bitmap from this Bitmap object. |
| [GetHICON](#GetHICON) | Creates an icon from this **Bitmap** object. |
| [GetHistogram](#GetHistogram) | Returns one or more histograms for specified color channels of this **Bitmap** object. |
| [GetHistogramSize](#GetHistogramSize) | Returns the number of elements (in an array of DWORDs) that you must allocate before you call the GetHistogram method of a Bitmap object. |
| [GetPixel](#GetPixel) | Gets the color of a specified pixel in this bitmap. |
| [InitializePalette](#InitializePalette) | Initializes a standard, optimal, or custom color palette. |
| [LockBits](#LockBits) | Locks a rectangular portion of this bitmap and provides a temporary buffer that you can use to read or write pixel data in a specified format. |
| [SetPixel](#SetPixel) | Sets the color of a specified pixel in this bitmap. |
| [SetResolution](#SetResolution) | Sets the resolution of this Bitmap object. |
| [UnlockBits](#UnlockBits) | Unlocks a portion of this bitmap that was previously locked by a call to LockBits. |

# CGpMetafile Class

Extends the **CGpImage** class. The **Metafile** object defines a graphic metafile. A metafile contains records that describe a sequence of graphics API calls. Metafiles can be recorded (constructed) and played back (displayed).

**Inherits from**: CGpImage.<br>
**Include file**: CGpBitmap.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorsMetafile) | Creates a Windows GDI+ Metafile. |
| [EmfToWmfBits](#EmfToWmfBits) | Converts an enhanced-format metafile to a Windows Metafile Format (WMF) metafile and stores the converted records in a specified buffer. |
| [GetDownLevelRasterizationLimit](#GetDownLevelRasterizationLimit) | Gets the rasterization limit currently set for this metafile. |
| [GetHENHMETAFILE](#GetHENHMETAFILE) | Gets a Windows handle to an Enhanced Metafile (EMF) file. |
| [GetMetafileHeader](#GetMetafileHeader) | Gets the metafile header of this metafile. |
| [PlayRecord](#PlayRecord) | Plays a metafile record. |
| [SetDownLevelRasterizationLimit](#SetDownLevelRasterizationLimit) | Sets the resolution for certain brush bitmaps that are stored in this metafile. |

# CGpCachedBitmap Class

Creates a **CachedBitmap** object based on a **Bitmap** object and a **Graphics** object. The cached bitmap takes the pixel data from the **Bitmap** object and stores it in a format that is optimized for the display device associated with the **Graphics** object.

**Inherits from**: CGpBase.<br>
**Include file**: CGpBitmap.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorCachedBitmap) | Creates a **CachedBitmap** object based on a **Bitmap** object and a **Graphics** object. |

# CGpImageAttributes Class

Creates an **ImageAttributes** object that contains information about how bitmap and metafile colors are manipulated during rendering. An **ImageAttributes** object maintains several color-adjustment settings, including color-adjustment matrices, grayscale-adjustment matrices, gamma-correction values, color-map tables, and color-threshold values.

**Inherits from**: CGpBase.<br>
**Include file**: CGpImageAttributes.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#ConstructorImageAttributes) | Creates an **ImageAttributes** object. |
| [ClearBrushRemapTable](#ClearBrushRemapTable) | Clears the brush color-remap table of this **ImageAttributes** object. |
| [ClearColorKey](#ClearColorKey) | Clears the color key (transparency range) for a specified category. |
| [ClearColorMatrices](#ClearColorMatrices) | Clears the color-adjustment matrix and the grayscale-adjustment matrix for a specified category. |
| [ClearColorMatrix](#ClearColorMatrix) | Clears the color-adjustment matrix for a specified category. |
| [ClearGamma](#ClearGamma) | Disables gamma correction for a specified category. |
| [ClearNoOp](#ClearNoOp) | Clears the **NoOp** setting for a specified category. |
| [ClearOutputChannel](#ClearOutputChannel) | Clears the cyan-magenta-yellow-black (CMYK) output channel setting for a specified category. |
| [ClearOutputChannelColorProfile](#ClearOutputChannelColorProfile) | Clears the output channel color profile setting for a specified category. |
| [ClearRemapTable](#ClearRemapTable) | Clears the color-remap table for a specified category. |
| [ClearThreshold](#ClearThreshold) | Clears the threshold value for a specified category. |
| [Clone](#Clone) | Copies the contents of the existing **ImageAttributes** object into a new **ImageAttributes** object. |
| [GetAdjustedPalette](#GetAdjustedPalette) | Adjusts the colors in a palette according to the adjustment settings of a specified category. |
| [Reset](#Reset) | Clears all color- and grayscale-adjustment settings for a specified category. |
| [SetBrushRemapTable](#SetBrushRemapTable) | Sets the color remap table for the brush category. |
| [SetColorKey](#SetColorKey) | Sets the color key (transparency range) for a specified category. |
| [SetColorMatrices](#SetColorMatrices) | Sets the color-adjustment matrix and the grayscale-adjustment matrix for a specified category. |
| [SetColorMatrix](#SetColorMatrix) | Sets the color-adjustment matrix for a specified category. |
| [SetGamma](#SetGamma) | Sets the gamma value for a specified category. |
| [SetNoOp](#SetNoOp) | Turns off color adjustment for a specified category. |
| [SetOutputChannel](#SetOutputChannel) | Sets the CMYK output channel for a specified category. |
| [SetOutputChannelColorProfile](#SetOutputChannelColorProfile) | Sets the output channel color-profile file for a specified category. |
| [SetRemapTable](#SetRemapTable) | Sets the color-remap table for a specified category. |
| [SetThreshold](#SetThreshold) | Sets the threshold (transparency range) for a specified category. |
| [SetToIdentity](#SetToIdentity) | Sets the color-adjustment matrix of a specified category to identity matrix. |
| [SetWrapMode](#SetWrapMode) | Sets the the wrap mode of this **ImageAttributes** object. |

# <a name="ConstructorsImage"></a>Constructors (CGpImage)

Creates an **Image** object from another **Image** object.

```
CONSTRUCTOR CGpImage (BYVAL pImage AS CGpImage PTR)
```

Creates an **Image** object based on a file.

```
CONSTRUCTOR CGpImage (BYVAL pwszFileName AS WSTRING PTR,  BYVAL useicm AS BOOLEAN = FALSE)
```

Create a **Image** object based on an **IStream** interface.

```
CONSTRUCTOR CGpImage (BYVAL pStream AS IStream PTR, BYVAL useicm AS BOOLEAN = FALSE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszFileName* | Pointer to a null-terminated string that specifies the path name of the image file. The graphics file formats supported by GDI+ are BMP, GIF, JPEG, PNG, TIFF, Exif, WMF, and EMF. |
| *pStream* | Pointer to an **IStream** interface. |
| *useicm* | Optional. Boolean value that specifies whether the new Image object applies color correction according to color management information that is embedded in the image file. Embedded information can include International Color Consortium (ICC) profiles, gamma values, and chromaticity information. TRUE specifies that color correction is enabled, and FALSE specifies that color correction is not enabled. The default value is FALSE. |

# <a name="CloneImage"></a>Clone (CGpImage)

Copies the contents of the existing **Image** object into a new **Image** object.

```
FUNCTION Clone (BYVAL pCloneImage AS CGpImage PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pCloneImage* | Pointer to the **Image** object where to copy the contents of the existing object. |

#### Example

```
' ========================================================================================
' The following example creates an Image object based on a JPEG file. The code creates a
' second Image object by cloning the first. Then the code calls the DrawImage method twice
' to draw the two images.
' ========================================================================================
SUB Example_Clone (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96

   ' // Create an Image object, and then clone it.
   DIM image1 AS CGpImage = "climber.jpg"
   DIM pImage2 AS CGpImage
   image1.Clone(@pImage2)
   ' // You can also use:
   ' DIM pImage2 AS CGpImage = @image1

   ' // Draw the original image and the cloned image.
   graphics.DrawImage(@image1, 20 * rxRatio, 20 * ryRatio)
   graphics.DrawImage(@pImage2, 230 * rxRatio, 20 * ryRatio)

END SUB
' ========================================================================================
```

# <a name="FindFirstItem"></a>FindFirstItem (CGpImage)

Retrieves the description and the data size of the first metadata item in this **Image** object.

```
FUNCTION FindFirstItem (BYVAL pitem AS ImageItemData PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pitem* | In, Out. Pointer to an **ImageItemData** structure. On input, the **Desc** member points to a buffer (allocated by the caller) large enough to hold the metadata description (1 byte for JPEG, 4 bytes for PNG, 11 bytes for GIF), and the **DescSize** member specifies the size (1, 4, or 6) of the buffer pointed to by **Desc**. On output, the buffer pointed to by **Desc** receives the metadata description, and the **DataSize** member receives the size, in bytes, of the metadata itself. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

Use **FindFirstItem** along with **FindNextItem** to enumerate the metadata items, including custom metadata, stored in an image. **FindFirstItem** and **FindNextItem** do not enumerate the metadata items stored by the **SetPropertyItem** method.

# <a name="FindNextItem"></a>FindNextItem (CGpImage)

The **FindNextItem** method is used along with the **FindFirstItem** method to enumerate the metadata items stored in this **Image** object. The **FindNextItem** method retrieves the description and the data size of the next metadata item in this **Image** object. 

```
FUNCTION FindNextItem (BYVAL pitem AS ImageItemData PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pitem* | In, Out. Pointer to an **ImageItemData** structure. On input, the **Desc** member points to a buffer (allocated by the caller) large enough to hold the metadata description (1 byte for JPEG, 4 bytes for PNG, 11 bytes for GIF), and the **DescSize** member specifies the size (1, 4, or 6) of the buffer pointed to by **Desc**. On output, the buffer pointed to by **Desc** receives the metadata description, and the **DataSize** member receives the size, in bytes, of the metadata itself. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

**FindFirstItem** and **FindNextItem** do not enumerate the metadata items stored by the **SetPropertyItem** method.

# <a name="GetAllPropertyItems"></a>GetAllPropertyItems (CGpImage)

Gets all the property items (metadata) stored in this **Image** object.

```
FUNCTION GetAllPropertyItems (BYVAL totalBufferSize AS UINT, BYVAL numProperties AS UINT, _
   BYVAL allItems AS PropertyItem PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *totalBufferSize* | Integer that specifies the size, in bytes, of the allItems buffer. Call the **GetPropertySize** method to obtain the required size.  |
| *numProperties* | Integer that specifies the number of properties in the image. Call the **GetPropertySize** method to obtain this number. |
| *allItems* | Pointer to an array of **PropertyItem** structures that receives the property items. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

Some image files contain metadata that you can read to determine features of the image. For example, a digital photograph might contain metadata that you can read to determine the make and model of the camera used to capture the image.

GDI+ stores an individual piece of metadata in a PropertyItem object. The **GetAllPropertyItems** method returns an array of PropertyItem objects. Before you call **GetAllPropertyItems**, you must allocate a buffer large enough to receive that array. You can call the **GetPropertySize** method of an **Image** object to get the size, in bytes, of the required buffer. The **GetPropertySize** method also gives you the number of properties (pieces of metadata) in the image.

Several enumerations and constants related to image metadata are defined in Gdiplusimaging.inc.

# <a name="GetBounds"></a>GetBounds (CGpImage)

Gets the bounding rectangle for this image.

```
FUNCTION GetBounds (BYVAL srcRect AS GpRectF PTR, BYVAL srcUnit AS GpUnit PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *srcRect* | Pointer to a **GpRectF** object that receives the bounding rectangle. |
| *srcUnit* | Pointer to a variable that receives an element of the **GpUnit** enumeration that indicates the unit of measure for the bounding rectangle. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

The bounding rectangle for a metafile does not necessarily have (0, 0) as its upper-left corner. The coordinates of the upper-left corner can be negative or positive, depending on the drawing commands that were issued during the recording of the metafile. For example, suppose a metafile consists of a single ellipse that was recorded with the following statement:

```
DrawEllipse(pen, 200, 100, 80, 40)
```

Then the bounding rectangle for the metafile will enclose that one ellipse. The upper-left corner of the bounding rectangle will not be (0, 0); rather, it will be a point near (200, 100).

#### Example

```
' ========================================================================================
' The following example creates an Image object based on a metafile and then draws the image.
' Next, the code calls the Image.GetBounds method to get the bounding rectangle for the image
' and redraws the a 75 per cent of the image.
' ========================================================================================
SUB Example_GetBounds (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM pImage AS CGpImage = "climber.emf"
   graphics.DrawImage(@pImage, 0, 0)

   ' // Get the bounding rectangle for the image (metafile).
   DIM boundsRect AS RectF
   DIM nUnit AS GpUnit
   pImage.GetBounds(@boundsRect, @nUnit)

   ' // Draw 75 percent of the image.
   graphics.DrawImage(@pImage, 230.0, 0.0, boundsRect.X, boundsRect.Y, 0.75 * boundsRect.Width, boundsRect.Height, UnitPixel)

END SUB
' ========================================================================================
```

# <a name="GetEncoderParameterList"></a>GetEncoderParameterList (CGpImage)

Gets a list of the parameters supported by a specified image encoder.

```
FUNCTION GetEncoderParameterList (BYVAL clsidEncoder AS GUID PTR, BYVAL uSize AS UINT, _
   BYVAL buffer AS EncoderParameters PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *clsidEncoder* | Pointer to a **CLSID** that specifies the encoder. |
| *uSize* | Integer that specifies the size, in bytes, of the buffer array. Call the **GetEncoderParameterListSize** method to obtain the required size. |
| *buffer* | Pointer to an **EncoderParameters** object that receives the list of supported parameters. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

The **GetEncoderParameterList** method returns an array of **EncoderParameter** objects. Before you call **GetEncoderParameterList**, you must allocate a buffer large enough to receive that array, which is part of an **EncoderParameters** object. You can call the **GetEncoderParameterListSize** method to get the size, in bytes, of the required **EncoderParameters** object.

# <a name="GetEncoderParameterListSize"></a>GetEncoderParameterListSize (CGpImage)

Gets the size, in bytes, of the parameter list for a specified image encoder.

```
FUNCTION GetEncoderParameterListSize (BYVAL clsidEncoder AS GUID PTR) AS UINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *clsidEncoder* | Pointer to a **CLSID** that specifies the encoder. |

# <a name="GetFlags"></a>GetFlags (CGpImage)

Gets a set of flags that indicate certain attributes of this **Image** object.

```
FUNCTION GetFlags () AS UINT
```

#### Return value

This method returns a value that holds a set of single-bit flags. The flags are defined in the **ImageFlags** enumeration.

# <a name="GetFrameCount"></a>GetFrameCount (CGpImage)

Gets the number of frames in a specified dimension of this **Image** object.

```
FUNCTION GetFrameCount (BYVAL dimensionID AS GUID PTR) AS UINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *dimensionID* | Pointer to a GUID that specifies the dimension. GUIDs that identify various dimensions are defined in Gdiplusimaging.inc. |

# <a name="GetFrameDimensionsCount"></a>GetFrameDimensionsCount (CGpImage)

Gets the number of frame dimensions in this **Image** object.

```
FUNCTION GetFrameDimensionsCount () AS UINT
```

#### Remarks

This method returns information about multiple-frame images, which come in two styles: multiple page and multiple resolution.

A multiple-page image is an image that contains more than one image. Each page contains a single image (or frame). These pages (or images, or frames) are typically displayed in succession to produce an animated sequence, such as in an animated GIF file.

A multiple-resolution image is an image that contains more than one copy of an image at different resolutions.

Windows GDI+ can support an arbitrary number of pages (or images, or frames), as well as an arbitrary number of resolutions.

# <a name="GetFrameDimensionsList"></a>GetFrameDimensionsList (CGpImage)

Gets the identifiers for the frame dimensions of this Image object.

```
FUNCTION GetFrameDimensionsList (BYVAL dimensionIDs AS GUID PTR, BYVAL count AS UINT) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *dimensionIDs* | Pointer to an array that receives the identifiers. GUIDs that identify various dimensions are defined in Gdiplusimaging.inc. |
| *count* | Integer that specifies the number of elements in the dimensionIDs array. Call the **GetFrameDimensionsCount** method to determine this number. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

This method returns information about multiple-frame images, which come in two styles: multiple page and multiple resolution.

A multiple-page image is an image that contains more than one image. Each page contains a single image (or frame). These pages (or images, or frames) are typically displayed in succession to produce an animated sequence, such as in an animated GIF file.

A multiple-resolution image is an image that contains more than one copy of an image at different resolutions.

Windows GDI+ can support an arbitrary number of pages (or images, or frames), as well as an arbitrary number of resolutions.

# <a name="GetHeight"></a>GetHeight (CGpImage)

Gets the image height, in pixels, of this image.

```
FUNCTION GetHeight () AS UINT
```

# <a name="GetHorizontalResolution"></a>GetHorizontalResolution (CGpImage)

Gets the horizontal resolution, in dots per inch, of this image.

```
FUNCTION GetHorizontalResolution () AS SINGLE
```

# <a name="GetItemData"></a>GetItemData (CGpImage)

Gets one piece of metadata from this **Image** object.

```
FUNCTION GetItemData (BYVAL pitem AS ImageItemData PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pitem* | Pointer to an **ImageItemData** object that specifies the item to be retrieved. The **Data** member of the **ImageItemData** object points to a buffer that receives the custom metadata. If the **Data** member is set to NULL, this method returns the size of the required buffer in the **DataSize** member of the **ImageItemData** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="GetPalette"></a>GetPalette (CGpImage)

Gets the **ColorPalette** of this **Image** object.

```
FUNCTION GetPalette (BYVAL pal AS ColorPalette PTR, BYVAL nSize AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pal* | Pointer to a **ColorPalette** structure that receives the palette. |
| *nSize* | Integer that specifies the size, in bytes, of the palette. Call the **GetPaletteSize** method to determine the size. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="GetPaletteSize"></a>GetPaletteSize (CGpImage)

Gets the size, in bytes, of the color palette of this **Image** object.

```
FUNCTION GetPaletteSize () AS INT_
```

# <a name="GetPhysicalDimension"></a>GetPhysicalDimension (CGpImage)

Gets the width and height of this **Image** object.

```
FUNCTION GetPhysicalDimension (BYVAL psize AS SizeF PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *psize* | Pointer to a **GpSizeF** object that receives the width and height of this image. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="GetPixelFormat"></a>GetPixelFormat (CGpImage)

Gets the pixel format of this **Image** object.

```
FUNCTION GetPixelFormat () AS PixelFormat
```

#### Return value

This method returns an integer that indicates the pixel format of this **Image** object. The **PixelFormat** data type and constants that represent various pixel formats are defined in Gdipluspixelformats.inc.

# <a name="GetPropertyCount"></a>GetPropertyCount (CGpImage)

Gets the number of properties (pieces of metadata) stored in this **Image** object.

```
FUNCTION GetPropertyCount () AS UINT
```

# <a name="GetPropertyIdList"></a>GetPropertyIdList (CGpImage)

Gets a list of the property identifiers used in the metadata of this **Image** object.

```
FUNCTION GetPropertyIdList (BYVAL numOfProperty AS UINT, BYVAL list AS PROPID PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *numOfProperty* | Integer that specifies the number of elements in the list array. Call the **GetPropertyCount** method to determine this number. |
| *list* | Pointer to an array that receives the property identifiers. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

The **GetPropertyIdList** method returns an array of **PROPIDs**. Before you call **GetPropertyIdList**, you must allocate a buffer large enough to receive that array. You can call the **GetPropertyCount** method of an **Image** object to determine the size of the required buffer. The size of the buffer should be the return value of **GetPropertyCount** multiplied by **SIZEOF(PROPID)**.

# <a name="GetPropertyItem"></a>GetPropertyItem (CGpImage)

Gets a specified property item (piece of metadata) from this **Image** object.

```
FUNCTION GetPropertyItem (BYVAL propId AS PROPID, BYVAL propSize AS UINT, _
   BYVAL buffer AS PropertyItem PTR= AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *propId* | Integer that identifies the property item to be retrieved. |
| *propSize* | Integer that specifies the size, in bytes, of the property item to be retrieved. Call the **GetPropertyItemSize** method to determine the size. |
| *buffer* | Pointer to a **PropertyItem** structure that receives the property item. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="GetPropertyItemSize"></a>GetPropertyItemSize (CGpImage)

Gets the size, in bytes, of a specified property item of this **Image** object.

```
FUNCTION GetPropertyItemSize (BYVAL propId AS PROPID) AS UINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *propId* | Integer that identifies the property item. |

# <a name="GetPropertySize"></a>GetPropertySize (CGpImage)

Gets the total size, in bytes, of all the property items stored in this **Image** object. The **GetPropertySize** method also gets the number of property items stored in this Image object.

```
FUNCTION GetPropertySize (BYVAL totalBufferSize AS UINT PTR, BYVAL numProperties AS UINT PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *totalBufferSize* | Pointer to a UINT that receives the total size, in bytes, of all the property items. |
| *numProperties* | Pointer to a UINT that receives the number of property items. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

Windows GDI+ stores an individual piece of metadata in a **PropertyItem** structure. The **GetAllPropertyItems** method returns an array of **PropertyItem** structures. Before you call **GetAllPropertyItems**, you must allocate a buffer large enough to receive that array. You can call the **GetPropertySize** method of an **Image** object to get the size, in bytes, of the required buffer. The **GetPropertySize** method also gives you the number of properties (pieces of metadata) in the image.

# <a name="GetRawFormat"></a>GetRawFormat (CGpImage)

Gets a globally unique identifier ( GUID) that identifies the format of this **Image** object.

```
FUNCTION GetRawFormat (BYVAL guidformat AS GUID PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *guidformat* | Pointer to a GUID that receives the format identifier. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="GetThumbnailImage"></a>GetThumbnailImage (CGpImage)

Gets a thumbnail image from this **Image** object.

```
FUNCTION GetThumbnailImage (BYVAL thumbWidth AS UINT, BYVAL thumbHeight AS UINT, _
   BYVAL pThumbnail AS CGpImage PTR, BYVAL pCallback AS ANY PTR = NULL, _
   BYVAL callbackData AS ANY PTR = NULL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *thumbWidth* | Width, in pixels, of the requested thumbnail image. |
| *thumbHeight* | Height, in pixels, of the requested thumbnail image. |
| *pThumbnail* | Pointer to an **Image** object that will receive the requested thumbnail image. |
| *pCallback* | Optional. Callback function that you provide. During the process of creating or retrieving the thumbnail image, GDI+ calls this function to give you the opportunity to abort the process. The default value is NULL. |
| *callbackData* | Optional. Pointer to a block of memory that contains data to be used by the callback function. The default value is NULL. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A thumbnail image is a small copy of an image. Some image files have a thumbnail image embedded in the file. In such cases, this method retrieves the embedded thumbnail image. If there is no embedded thumbnail image, this method creates a thumbnail image by scaling the main image to the size specified in the *thumbWidth* and *thumbHeight* parameters. If both of those parameters are 0, a system-defined size is used.

#### Example

```
' ========================================================================================
' The following example creates an Image object based on a metafile and then draws the image.
' Next, the code calls the Image.GetBounds method to get the bounding rectangle for the image.
' The code makes two attempts to display 75 percent of the image. The first attempt fails
' because it specifies (0, 0) for the upper-left corner of the source rectangle. The second
' attempt succeeds because it uses the X and Y data members returned by Image.GetBounds to
' specify the upper-left corner of the source rectangle.
' ========================================================================================
SUB Example_GetThumbnailImage (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM pImage AS CGpImage = "climber.jpg"
   DIM pThumbnail AS CGpImage
   pImage.GetThumbnailImage(40, 40, @pThumbnail)

   graphics.DrawImage(@pImage, 10, 10, pImage.GetWidth, pImage.GetHeight)
   graphics.DrawImage(@pThumbnail, 220, 10, pThumbnail.GetWidth, pThumbnail.GetHeight)

END SUB
' ========================================================================================
```

# <a name="GetType"></a>GetType (CGpImage)

Gets the type (bitmap or metafile) of this **Image** object.

```
FUNCTION GetType () AS ImageType
```

#### Return value

This method returns an element of the ImageType enumeration that indicates the image type.

# <a name="GetVerticalResolution"></a>GetVerticalResolution (CGpImage)

Gets the vertical resolution, in dots per inch, of this image.

```
FUNCTION GetVerticalResolution () AS SINGLE
```

# <a name="GetWidth"></a>GetWidth (CGpImage)

Gets the width, in pixels, of this image

```
FUNCTION GetWidth () AS UINT
```

# <a name="RemovePropertyItem"></a>RemovePropertyItem (CGpImage)

Removes a property item (piece of metadata) from this **Image** object.

```
FUNCTION RemovePropertyItem (BYVAL propId AS PROPID) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *propId* | Integer that identifies the property item. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

The **RemovePropertyItem** method removes a specified property from an **Image** object, but that property item is not removed from the file or stream that was used to construct the Image object. To save the image (with the property item removed) to a new JPEG file or stream, call the **Save** method of the **Image** object.

# <a name="RotateFlip"></a>RotateFlip (CGpImage)

Rotates and flips this image.

```
FUNCTION RotateFlip (BYVAL rotateFlipType AS RotateFlipType) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *rotateFlipType* | Element of the **RotateFlipType** enumeration that specifies the type of rotation and the type of flip. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates an Image object based on a JPEG file. The code creates a
' second Image object by cloning the first. Then the code calls the DrawImage method twice
' to draw the two images.
' ========================================================================================
SUB Example_RotateFlip (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM pImage AS CGpImage = "climber.jpg"
   graphics.DrawImage(@pImage, 10, 10, pImage.GetWidth, pImage.GetHeight)
   pImage.RotateFlip(Rotate90FlipY)
   graphics.DrawImage(@pImage, 220, 10, pImage.GetWidth, pImage.GetHeight)

END SUB
' ========================================================================================
```

# <a name="Save"></a>Save (CGpImage)

Saves this image to a file.

```
FUNCTION Save (BYVAL pwszFileName AS WSTRING PTR, BYVAL clsidEncoder AS GUID PTR, _
   BYVAL encoderParams AS EncoderParameters PTR) AS GpStatus
FUNCTION Save (BYVAL pStream AS IStream PTR, BYVAL clsidEncoder AS GUID PTR, _
   BYVAL encoderParams AS EncoderParameters PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszFileName* | Pointer to a null-terminated string that specifies the path name for the saved image. |
| *pStream* | Pointer to an **IStream** interface. The implementation of **IStream** must include the **Seek**, **Read**, **Write**, and **Stat** methods. |
| *clsidEncoder* | Pointer to a **CLSID** that specifies the encoder to use to save the image. |
| *encoderParams* | Optional. Pointer to an **EncoderParameters** structure that holds parameters used by the encoder. The default value is NULL. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

GDI+ does not allow you to save an image to the same file that you used to construct the image. The following code creates an **Image** object by passing the file name MyImage.jpg to an **Image** constructor. That same file name is passed to the **Save** method of the **Image** object, so the **Save** method fails.

# <a name="SaveAdd"></a>SaveAdd (CGpImage)

Adds a frame to a file or stream specified in a previous call to the **Save** method.

```
FUNCTION SaveAdd (BYVAL encoderParams AS EncoderParameters PTR) AS GpStatus
FUNCTION SaveAdd (BYVAL pNewImage AS CGpImage PTR, BYVAL encoderParams AS EncoderParameters PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *encoderParams* | Pointer to an **Image** object that holds the frame to be added. |
| *pNewImage* | Pointer to an **EncoderParameters** structure that holds parameters required by the image encoder used by the save-add operation.  |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="SelectActiveFrame"></a>SelectActiveFrame (CGpImage)

Selects the frame in this **Image** object specified by a dimension and an index.

```
FUNCTION SelectActiveFrame (BYVAL dimensionID AS GUID PTR, BYVAL frameIndex AS UINT) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *dimensionID* | Pointer to a GUID that specifies the frame dimension. GUIDs that identify various frame dimensions are defined in Gdiplusimaging.inc. |
| *frameIndex* | Integer that specifies the index of the frame within the specified frame dimension. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

When you call the **SelectActiveFrame** method, all changes that you made to the previously active frame are discarded. If you want to retain changes that you make to a frame, call the Save method before you switch to a different frame.

Among all the image formats currently supported by GDI+, the only formats that support multiple-frame images are GIF and TIFF. When you call the **SelectActiveFrame** method on a GIF image, you should use **FrameDimensionTime**. When you call the SelectActiveFrame method on a TIFF image, you should use **FrameDimensionPage**.

# <a name="SetPalette"></a>SetPalette (CGpImage)

Sets the color palette of this **Image** object.

```
FUNCTION SetPalette (BYVAL pal AS ColorPalette PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *palette* | Pointer to a **ColorPalette** structure that specifies the palette. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="SetPropertyItem"></a>SetPropertyItem (CGpImage)

Sets a property item (piece of metadata) for this Image object. If the item already exists, then its contents are updated; otherwise, a new item is added.

```
FUNCTION SetPropertyItem (BYVAL pitem AS PropertyItem PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pitem* | Pointer to a **PropertyItem** object that specifies the property item to be set. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="ConstructorsBitmap"></a>Constructors (CGpBitmap)

Creates a **Bitmap** object based on another **Bitmap** obejct (clones it).

```
CONSTRUCTOR CGpBitmap (BYVAL pBitmap AS CGpBitmap PTR)
```

Creates a **Bitmap** object based on an icon.

```
CONSTRUCTOR CGpBitmap (BYVAL hicon AS HICON)
```

Creates a **Bitmap** object based on an application or DLL instance handle and the name of a bitmap resource.

```
CONSTRUCTOR CGpBitmap (BYVAL hInstance AS HINSTANCE, BYVAL pwszBitmapName AS WSTRING PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hIcon* | Handle to a GDI icon. |
| *hInstance* | Handle to an instance of a module whose executable file contains a bitmap resource. |
| *pwszBitmapName* | Pointer to a null-terminated string that specifies the path name of the bitmap resource to be loaded. Alternatively, this parameter can consist of the resource identifier in the low-order word and zero in the high-order word. You can use the MAKEINTRESOURCE macro to create this value. |

# <a name="CloneBitmap"></a>Clone (CGpBitmap)

Creates a new **Bitmap** object from another **Bitmap** object (clones it).

```
FUNCTION Clone (BYVAL pCloneBitmap AS CGpBitmap PTR) AS GpStatus
```

Creates a new **Bitmap** object by copying a portion of this bitmap.

```
FUNCTION Clone (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, _
   BYVAL nHeight AS SINGLE, BYVAL pxFormat AS PixelFormat, BYVAL pCloneBitmap AS CGpBitmap PTR) AS GpStatus
FUNCTION Clone (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, _
   BYVAL pxFormat AS PixelFormat, BYVAL pCloneBitmap AS CGpBitmap PTR) AS GpStatus
FUNCTION Clone (BYVAL rc AS GpRectF PTR, BYVAL pxFormat AS PixelFormat, _
   BYVAL pCloneBitmap AS CGpBitmap PTR) AS GpStatus
FUNCTION Clone (BYVAL rc AS GpRect PTR, BYVAL pxFormat AS PixelFormat, _
   BYVAL pCloneBitmap AS CGpBitmap PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | The x-coordinate of the upper-left corner of the rectangle that specifies the portion of this bitmap to copy. |
| *y* | The width of the rectangle that specifies the portion of this bitmap to copy. |
| *nWidth* | The y-coordinate of the upper-left corner of the rectangle that specifies the portion of this bitmap to copy. |
| *nHeight* | The height of the rectangle that specifies the portion of this image to copy. |
| *pxFormat* | Integer that specifies the pixel format of the new bitmap. The **PixelFormat** data type and constants that represent various pixel formats are defined in Gdipluspixelformats.inc. |
| *pCloneBitmap* | Pointer to the **Bitmap** object where to copy the contents of the existing object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Bitmap from an image file, clones the upper-left
' portion of the image, and then draws the cloned image.
' ========================================================================================
SUB Example_CloneArea (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96

   ' // Create a Bitmap object from a JPEG file.
   DIM myBitmap AS CGpBitmap = "climber.jpg"

   ' // Clone a portion of the bitmap.
   DIM cloneBitmap AS CGpBitmap
   myBitmap.Clone(0, 0, 100, 100, PixelFormatDontCare, @cloneBitmap)

   ' // Draw the clone.
   graphics.DrawImage(@cloneBitmap, 0, 0)

END SUB
' ========================================================================================
```

# <a name="ConvertFormat"></a>ConvertFormat (CGpBitmap)

Converts a bitmap to a specified pixel format. The original pixel data in the bitmap is replaced by the new pixel data.

```
FUNCTION ConvertFormat (BYVAL pxFormat AS PixelFormat, BYVAL nDitherType AS DitherType, _
   BYVAL nPaletteType AS PaletteType, BYVAL colourPalette AS ColorPalette PTR, _
   BYVAL alphaThresholdPercent AS SINGLE) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pxFormat* | Pixel format constant that specifies the new pixel format. |
| *dithertype* | Element of the **DitherType** enumeration that specifies the dithering algorithm. In cases where the conversion does not reduce the bit depth of the pixel data, pass **DitherTypeNone**. |
| *nPaletteType* | Element of the **PaletteType** enumeration that specifies a standard palette to be used for dithering. If you are converting to a non-indexed format, this parameter is ignored. In that case, pass any element of the **PaletteType** enumeration, say **PaletteTypeCustom**. |
| *colourPalette* | Pointer to a ColorPalette structure that specifies the palette whose indexes are stored in the pixel data of the converted bitmap. This palette (called the actual palette) does not have to have the type specified by the *palettetype* parameter. The *palettetype* parameter specifies a standard palette that can be used by any of the ordered or spiral dithering algorithms. If the actual palette has a type other than that specified by the palettetype parameter, then the **ConvertFormat** method performs a nearest-color conversion from the standard palette to the actual palette. |
| *alphaThresholdPercent* | Simple precision number in the range 0 through 100 that specifies which pixels in the source bitmap will map to the transparent color in the converted bitmap. A value of 0 specifies that none of the source pixels map to the transparent color. A value of 100 specifies that any pixel that is not fully opaque will map to the transparent color. A value of t specifies that any source pixel less than t percent of fully opaque will map to the transparent color. Note that for the alpha threshold to be effective, the palette must have a transparent color. If the palette does not have a transparent color, pixels with alpha values below the threshold will map to color that most closely matches (0, 0, 0, 0), usually black. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="GetHBITMAP"></a>GetHBITMAP (CGpBitmap)

Creates a Windows Graphics Device Interface (GDI) bitmap from this **Bitmap** object.

```
FUNCTION GetHBITMAP (BYVAL colorBackground AS ARGB, BYVAL hbmReturn AS HBITMAP PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *colorBackground* | ARGB color that specifies the background color. This parameter is ignored if the bitmap is totally opaque. |
| *hbmReturn* | Pointer to an HBITMAP that receives a handle to the GDI bitmap. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="GetHICON"></a>GetHICON (CGpBitmap)

Creates an icon from this **Bitmap** object.

```
FUNCTION GetHICON (BYVAL hiconReturn AS HICON PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *hIconReturn* | Pointer to an HICON that receives a handle to the icon. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="GetHistogram"></a>GetHistogram (CGpBitmap)

Returns one or more histograms for specified color channels of this **Bitmap** object.

```
FUNCTION GetHistogram (BYVAL nFormat AS HistogramFormat, BYVAL NumberOfEntries AS UINT, _
   BYVAL channel0 AS UINT PTR, BYVAL channel1 AS UINT PTR, BYVAL channel2 AS UINT PTR, _
   BYVAL channel3 AS UINT PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nFormat* | Element of the **HistogramFormat** enumeration that specifies the channels for which histograms will be created. |
| *NumberOfEntries* | Integer that specifies the number of elements (of type UINT) in each of the arrays pointed to by *channel0*, *channel1*, *channel2*, and *channel3*. You must allocate memory for those arrays before you call **GetHistogram**. To determine the required number of elements, call **GetHistogramSize**. |
| *channel0* | Pointer to an array of DWORDs that receives the first histogram. |
| *channel1* | Pointer to an array of DWORDs that receives the second histogram if there is a second histogram. Pass NULL if there is no second histogram. |
| *channel2* | Pointer to an array of DWORDs that receives the third histogram if there is a third histogram. Pass NULL if there is no third histogram. |
| *channel3* | Pointer to an array of DWORDs that receives the fourth histogram if there is a fourth histogram. Pass NULL if there is no fourth histogram. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

The number of histograms returned depends on the HistogramFormat enumeration element passed to the format parameter. For example, if format is equal to **HistogramFormatRGB**, then three histograms are returned: one each for the red, green, and blue channels. In that case, *channel0* points to the array that receives the red-channel histogram, *channel1* points to the array that receives the green-channel histogram, and *channel2* points to the array that receives the blue-channel histogram. For **HistogramFormatRGB**, *channel3* must be set to NULL because there is no fourth histogram. For more details, see the **HistogramFormat** enumeration.

# <a name="GetHistogramSize"></a>GetHistogramSize (CGpBitmap)

Returns the number of elements (in an array of DWORDs) that you must allocate before you call the **GetHistogram** method of a **Bitmap** object.

```
FUNCTION GetHistogramSize (BYVAL nFormat AS HistogramFormat, BYVAL NumberOfEntries AS UINT PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nFormat* | Element of the **HistogramFormat** enumeration that specifies the pixel format of the bitmap. |
| *NumberOfEntries* | Pointer to a DWORD that receives the number of entries. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="GetPixel"></a>GetPixel (CGpBitmap)

Gets the color of a specified pixel in this bitmap.

```
FUNCTION GetPixel (BYVAL x AS LONG, BYVAL y AS LONG, BYVAL colour AS ARGB PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | Integer that specifies the x-coordinate (column) of the pixel. |
| *y* | Integer that specifies the y-coordinate (row) of the pixel. |
| *colour* | Pointer to a DWORD that receives the color of the specified pixel. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

Depending on the format of the bitmap, **GetPixel** might not return the same value as was set by **SetPixel**. For example, if you call **SetPixel** on a **Bitmap** object whose pixel format is 32bppPARGB, the pixel's RGB components are premultiplied. A subsequent call to **GetPixel** might return a different value because of rounding. Also, if you call **SetPixel** on a **Bitmap** object whose color depth is 16 bits per pixel, information could be lost during the conversion from 32 to 16 bits, and a subsequent call to **GetPixel** might return a different value.

# <a name="InitializePalette"></a>InitializePalette (CGpBitmap)

Initializes a standard, optimal, or custom color palette.

```
FUNCTION InitializePalette (BYVAL colourPalette AS ColorPalette PTR, BYVAL nPaletteType AS PaletteType, _
   BYVAL optimalColors AS INT_, BYVAL useTransparentColor AS BOOL, BYVAL pBitmap AS CGpBitmap PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *colourPalette* | Pointer to a buffer that contains a **ColorPalette** structure followed by an array of ARGB values. The Entries member of a **ColorPalette** structure is an array of one ARGB value. You must allocate memory for the **ColorPalette** structure and for the additional ARGB values in the palette. For example, if the palette has 36 ARGB values, allocate a buffer as follows: HeapAlloc(SIZEOF(ColorPalette) + 35\*4). |
| *palettetype* | Element of the **PaletteType** enumeration that specifies the palette type. The palette can have one of several standard types, or it can be a custom palette that you define. Also, the **InitializePalette** method can create an optimal palette based on a specified bitmap. |
| *optimalColors* | Integer that specifies the number of colors you want to have in an optimal palette based on a specified bitmap. If this parameter is greater than 0, the *palettetype* parameter must be set to **PaletteTypeOptimal**, and the bitmap parameter must point to a **Bitmap** object. If you are creating a standard or custom palette rather than an optimal palette, set this parameter to 0. |
| *useTransparentColor* | Boolean value that specifies whether to include the transparent color in the palette. Set to TRUE to include the transparent color; otherwise FALSE. |
| *pBitmap* | Pointer to a **Bitmap** object for which an optimal palette will be created. If palettetype is set to **PaletteTypeOptimal** and optimalColors is set to a positive integer, set this parameter to the address of a **Bitmap** object. Otherwise, set this parameter to NULL. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="LockBits"></a>LockBits (CGpBitmap)

Locks a rectangular portion of this bitmap and provides a temporary buffer that you can use to read or write pixel data in a specified format. Any pixel data that you write to the buffer is copied to the **Bitmap** object when you call **UnlockBits**.

```
FUNCTION LockBits (BYVAL rc AS GpRect PTR, BYVAL flags AS UINT, BYVAL pxFormat AS PixelFormat, _
   BYVAL lockedBitmapData AS BitmapData PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *rc* | Pointer to a rectangle that specifies the portion of the bitmap to be locked. |
| *flags* | Set of flags that specify whether the locked portion of the bitmap is available for reading or for writing and whether the caller has already allocated a buffer. Individual flags are defined in the **ImageLockMode** enumeration. |
| *pxFormat* |  Integer that specifies the format of the pixel data in the temporary buffer. The pixel format of the temporary buffer does not have to be the same as the pixel format of this **Bitmap** object. The **PixelFormat** data type and constants that represent various pixel formats are defined in Gdipluspixelformats.inc. GDI+ version 1.0 does not support processing of 16-bits-per-channel images, so you should not set this parameter equal to PixelFormat48bppRGB, PixelFormat64bppARGB, or PixelFormat64bppPARGB. |
| *lockedBitmapData* | Pointer to a **BitmapData** object. If the **ImageLockModeUserInputBuf** flag of the flags parameter is cleared, then *lockedBitmapData* serves only as an output parameter. In that case, the *Scan0* data member of the **BitmapData** object receives a pointer to a temporary buffer, which is filled with the values of the requested pixels. The other data members of the **BitmapData** object receive attributes (width, height, format, and stride) of the pixel data in the temporary buffer. If the pixel data is stored bottom-up, the *Stride* data member is negative. If the pixel data is stored top-down, the *Stride* data member is positive. If the **ImageLockModeUserInputBuf** flag of the *flags* parameter is set, then *lockedBitmapData* serves as an input parameter (and possibly as an output parameter). In that case, the caller must allocate a buffer for the pixel data that will be read or written. The caller also must create a **BitmapData** object, set the *Scan0* data member of that **BitmapData** object to the address of the buffer, and set the other data members of the **BitmapData** object to specify the attributes (width, height, format, and stride) of the buffer. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="SetPixel"></a>SetPixel (CGpBitmap)

Sets the color of a specified pixel in this bitmap.

```
FUNCTION SetPixel (BYVAL x AS LONG, BYVAL y AS LONG, BYVAL colour AS ARGB) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | Integer that specifies the x-coordinate (column) of the pixel. |
| *y* | Integer that specifies the y-coordinate (row) of the pixel. |
| *colour* | Pointer to a variable that receives the color to set.|

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

Depending on the format of the bitmap, **GetPixel** might not return the same value as was set by **SetPixel**. For example, if you call SetPixel on a Bitmap object whose pixel format is 32bppPARGB, the RGB components are premultiplied. A subsequent call to **GetPixel** might return a different value because of rounding. Also, if you call **SetPixel** on a **Bitmap** whose color depth is 16 bits per pixel, information could be lost in the conversion from 32 to 16 bits, and a subsequent call to **GetPixel** return a different value.

# <a name="SetResolution"></a>SetResolution (CGpBitmap)

Sets the resolution of this **Bitmap** object.

```
FUNCTION SetResolution (BYVAL xdpi AS SINGLE, BYVAL ydpi AS SINGLE) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *xdpi* | The horizontal resolution in dots per inch. |
| *ydpi* | The vertical resolution in dots per inch. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="UnlockBits"></a>UnlockBits (CGpBitmap)

Unlocks a portion of this bitmap that was previously locked by a call to **LockBits**.

```
FUNCTION UnlockBits (BYVAL lockedBitmapData AS BitmapData PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *lockedBitmapData* | Pointer to a **BitmapData** object that was previously passed to **LockBits**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

**LockBits** and **UnlockBits** must be used as a pair. A call to the **LockBits** method of a **Bitmap** object establishes a temporary buffer that you can use to read or write pixel data in a specified format. After you write to the temporary buffer, a call to **UnlockBits** copies the pixel data in the buffer to the **Bitmap** object. If the pixel format of the temporary buffer is different from the pixel format of the **Bitmap** object, the pixel data is converted appropriately.

# <a name="ConstructorsMetafile"></a>Constructors (CGpMetafile)

Creates a Windows GDI+ Metafile object for playback based on a Windows Graphics Device Interface (GDI) Enhanced Metafile (EMF) file.

```
CONSTRUCTOR CGpMetafile (BYVAL hEmf AS HENHMETAFILE, BYVAL deleteEmf AS BOOL = FALSE)
```

Creates a **Metafile** object from an **IStream** interface for playback.

```
CONSTRUCTOR CGpMetafile (BYVAL pStream AS IStream PTR)
```

Creates a **Metafile** object for playback.

```
CONSTRUCTOR CGpMetafile (BYVAL pwszFileName AS WSTRING PTR)
CONSTRUCTOR CGpMetafile (BYVAL pwszFileName AS WSTRING PTR, BYVAL wmfPFH AS WmfPlaceableFileHeader PTR)
CONSTRUCTOR CGpMetafile (BYVAL hWmf AS HMETAFILE, BYVAL wmfPFH AS WmfPlaceableFileHeader PTR, _
   BYVAL deleteWmf AS BOOL = FALSE)
```

Creates a **Metafile** object for recording.

```
CONSTRUCTOR CGpMetafile (BYVAL referenceHdc AS HDC, BYVAL nType AS EmfType = EmfTypeEmfPlusDual, _
   BYVAL description AS WSTRING PTR = NULL)
CONSTRUCTOR CGpMetafile (BYVAL referenceHdc AS HDC, BYVAL frmRect AS GpRectF PTR, _
   BYVAL frameUnit AS MetafileFrameUnit = MetafileFrameUnitGdi, _
   BYVAL nType AS EmfType = EmfTypeEmfPlusDual, BYVAL description AS WSTRING PTR = NULL)
CONSTRUCTOR CGpMetafile (BYVAL referenceHdc AS HDC, BYVAL frmRect AS GpRect PTR, _
   BYVAL frameUnit AS MetafileFrameUnit = MetafileFrameUnitGdi, _
   BYVAL nType AS EmfType = EmfTypeEmfPlusDual, BYVAL description AS WSTRING PTR = NULL)
```

Creates a **Metafile** object for recording.

```
CONSTRUCTOR CGpMetafile (BYVAL pwszFileName AS WSTRING PTR, BYVAL referenceHdc AS HDC, _
   BYVAL nType AS EmfType = EmfTypeEmfPlusDual, BYVAL description AS WSTRING PTR = NULL)
CONSTRUCTOR CGpMetafile (BYVAL pwszFileName AS WSTRING PTR, BYVAL referenceHdc AS HDC, _
   BYVAL frmRect AS GpRectF PTR, BYVAL frameUnit AS MetafileFrameUnit = MetafileFrameUnitGdi, _
   BYVAL nType AS EmfType = EmfTypeEmfPlusDual, BYVAL description AS WSTRING PTR = NULL)
CONSTRUCTOR CGpMetafile (BYVAL pwszFileName AS WSTRING PTR, BYVAL referenceHdc AS HDC, _
   BYVAL frmRect AS GpRect PTR, BYVAL frameUnit AS MetafileFrameUnit = MetafileFrameUnitGdi, _
   BYVAL nType AS EmfType = EmfTypeEmfPlusDual, BYVAL description AS WSTRING PTR = NULL)
```

Creates a **Metafile** object from an **IStream** interface for recording.

```
CONSTRUCTOR CGpMetafile (BYVAL pStream AS IStream PTR, BYVAL referenceHdc AS HDC, _
   BYVAL nType AS EmfType, BYVAL description AS WSTRING PTR)
CONSTRUCTOR CGpMetafile (BYVAL pStream AS IStream PTR, BYVAL referenceHdc AS HDC, _
   BYVAL frmRect AS GpRectF PTR, BYVAL frameUnit AS MetafileFrameUnit = MetafileFrameUnitGdi, _
   BYVAL nType AS EmfType, BYVAL description AS WSTRING PTR)
CONSTRUCTOR CGpMetafile (BYVAL pStream AS IStream PTR, BYVAL referenceHdc AS HDC, _
   BYVAL frmRect AS GpRect PTR, BYVAL frameUnit AS MetafileFrameUnit = MetafileFrameUnitGdi, _
   BYVAL nType AS EmfType, BYVAL description AS WSTRING PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hEmf* | Windows handle to an enhanced-format metafile. |
| *deleteEmf* | Optional. Boolean value that specifies whether the Windows handle to a metafile is deleted when the **Metafile** object is deleted. TRUE specifies that the *hEmf* Windows handle is deleted, and FALSE specifies that the *hEmf* Windows handle is not | *pStream* | Pointer to a **IStream** interface that points to a data stream in a file. If the **Metafile** has been created for recording, when the commands are recorded, they will be saved to this stream. |
| *pwszFileName* | Pointer to a wide-character string that specifies the name of an existing disk file used to create the **Metafile** object for playback.  |
| *hWmf* | Windows handle to an windows metafile format. |
| *wmfPFH* | Pointer to a **WmfPlaceableFileHeader** structure that specifies a preheader preceding the metafile header. |
deleted. The default value is FALSE. |
| *deleteWmf* | Optional. Boolean value that specifies whether the Windows handle to a metafile is deleted when the **Metafile** object is deleted. TRUE specifies that the *hWmf* Windows handle is deleted, and FALSE specifies that the *hWmf* Windows handle is not | *referenceHdc* | Windows handle to a metafile. |
| *frameRect* | Reference to a rectangle that bounds the metafile display. |
| *frameUnit* | Optional. Element of the **MetafileFrameUnit** enumeration that specifies the unit of measure for **frameRect**. The default value is **MetafileFrameUnitGdi**. |
| *nType* | Optional. Element of the **EmfType** enumeration that specifies the type of metafile that will be recorded. The default value is **EmfTypeEmfPlusDual**. |
| *description* | Optional. Pointer to a wide-character string that specifies the descriptive name of the metafile. The default value is NULL. |

#### Remarks

When recording to a file, the file must be writable, and Windows GDI+ must be able to obtain an exclusive lock on the file.

# <a name="EmfToWmfBits"></a>EmfToWmfBits (CGpMetafile)

Converts an enhanced-format metafile to a Windows Metafile Format (WMF) metafile and stores the converted records in a specified buffer.

```
FUNCTION EmfToWmfBits (BYVAL hEmf AS HENHMETAFILE, BYVAL cbData16 AS UINT, BYVAL pData16 AS BYTE PTR, _
   BYVAL iMapMode AS INT_ = MM_ANISOTROPIC, BYVAL eFlags AS INT_ = EmfToWmfBitsFlagsDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *hEmf* | Windows handle to an enhanced-format metafile. |
| *cbData16* | Unsigned integer that specifies the number of bytes in the buffer pointed to by the *pData16* parameter. |
| *pData16* | Pointer to a buffer that receives the converted records. If *pData16* is NULL, **EmfToWmfBits** returns the number of bytes required to store the converted metafile records. |
| *iMapMode* | Optional. Specifies the mapping mode to use in the converted metafile. For a list of possible mapping modes, see **SetMapMode**. The default value is MM_ANISOTROPIC. |
| *eFlags* | Optional. Element of the **EmfToWmfBitsFlags** enumeration that specifies options for the conversion. The default value is **EmfToWmfBitsFlagsDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

This method replaces the records originally in the Metafile object with the converted records. To retain a copy of the original **Metafile** object, call the **Clone** method.

If you set the *emfType* parameter to **EmfTypeEmfPlusDual**, the converted metafile contains an Enhanced Metafile (EMF) representation and an EMF+ representation. The EMF representation is the original set of EMF records rather than EMF records converted back from the newly created EMF+ records.

It is possible for the return value to be **Ok** and the value returned in *conversionSuccess* to be FALSE. Sometimes the overall conversion is considered to be successful even if a few individual records failed to convert with complete accuracy. For example, the original metafile might have records or operations that are not supported by Windows GDI+ (or EMF+), in which case those records or operations are emulated.

# <a name="GetDownLevelRasterizationLimit"></a>GetDownLevelRasterizationLimit (CGpMetafile)

Gets the rasterization limit currently set for this metafile. The rasterization limit is the resolution used for certain brush bitmaps that are stored in the metafile. For a detailed explanation of the rasterization limit, see **SetDownLevelRasterizationLimit**.

```
FUNCTION GetDownLevelRasterizationLimit () AS UINT
```
#### Return value

This method returns the rasterization limit in dpi.

# <a name="GetHENHMETAFILE"></a>GetHENHMETAFILE (CGpMetafile)

Gets a Windows handle to an Enhanced Metafile (EMF) file.

```
FUNCTION GetHENHMETAFILE () AS HENHMETAFILE
```
#### Return value

This method sets the **Metafile** object to an invalid state. The user is responsible for calling **DeleteEnhMetafile**, to delete the Windows handle.

# <a name="GetMetafileHeader"></a>GetMetafileHeader (CGpMetafile)

Gets the metafile header of this metafile.

```
FUNCTION GetMetafileHeader (BYVAL mh AS MetafileHeader PTR) AS GpStatus
FUNCTION GetMetafileHeader (BYVAL hEmf AS HENHMETAFILE, BYVAL mh AS MetafileHeader PTR) AS GpStatus
FUNCTION GetMetafileHeader (BYVAL hWmf AS HMETAFILE, BYVAL mh AS MetafileHeader PTR) AS GpStatus
FUNCTION GetMetafileHeader (BYVAL pwsFileName AS WSTRING PTR, BYVAL mh AS MetafileHeader PTR) AS GpStatus
FUNCTION GetMetafileHeader (BYVAL pStream AS IStream PTR, BYVAL mh AS MetafileHeader PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *mf* | Pointer to a MetafileHeader structure that receives the header, which contains the attributes of the metafile. |
| *hEmf* | Windows handle to an enhanced-format metafile. |
| *hWmf* | Windows handle to an windows metafile format. |
| *wmfPFH* | Pointer to a placeable metafile header. |
| *pwszFileName* | Pointer to a wide-character string that specifies the name of an existing **Metafile** object that contains the header. |
| *pStream* | Pointer to a **IStream** interface that points to a data stream in a file. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="PlayRecord"></a>PlayRecord (CGpMetafile)

Plays a metafile record.

```
FUNCTION PlayRecord (BYVAL recordType AS EmfPlusRecordType, BYVAL flags AS UINT, _
   BYVAL dataSize AS UINT, BYVAL pData AS BYTE PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *recordType* | Element of the **EmfPlusRecordType** enumeration that specifies the type of metafile record to be played. |
| *flags* | Set of flags that specify attributes of the record to be played. |
| *dataSize* | Integer that specifies the number of bytes contained in the record data. |
| *pData* | Pointer to an array of bytes that contains the record data. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="SetDownLevelRasterizationLimit"></a>SetDownLevelRasterizationLimit (CGpMetafile)

Sets the resolution for certain brush bitmaps that are stored in this metafile.

```
FUNCTION SetDownLevelRasterizationLimit (BYVAL metafileRasterizationLimitDpi AS UINT) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *metafileRasterizationLimitDpi* | Non-negative integer that specifies the resolution in dpi. If you set this parameter equal to 0, the resolution is set to match the resolution of the device context handle that was passed to the **Metafile** constructor. If you set this parameter to a value greater than 0 but less than 10, the resolution is left unchanged. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

The purpose of this method is to prevent metafiles from becoming too large as a result of texture and gradient brushes being stored at high resolution. Suppose you construct a **Metafile** object (for recording an **EmfTypeEmfOnly** metafile) based on the device context of a printer that has a resolution of 600 dpi. Also suppose you create a path gradient brush or a texture brush based on a **Bitmap** object that has a resolution of 96 dpi. If the bitmap that represents that brush is stored in the metafile with a resolution of 96 dpi, it will require much less space than if it is stored with a resolution of 600 dpi.

The default rasterization limit for metafiles is 96 dpi. So if you do not call this method at all, path gradient brush and texture brush bitmaps are stored with a resolution of 96 dpi.

The rasterization limit has an effect on metafiles of type **EmfTypeEmfOnly** and **EmfTypeEmfPlusDual**, but it has no effect on metafiles of type **EmfTypeEmfPlusOnly**.

# <a name="ConstructorCachedBitmap"></a>Constructor (CGpCachedBitmap)

Creates a **CachedBitmap** object based on a **Bitmap** object and a **Graphics** object. The cached bitmap takes the pixel data from the **Bitmap** object and stores it in a format that is optimized for the display device associated with the **Graphics** object.

```
CONSTRUCTOR CGpCachedBitmap (BYVAL pBitmap AS CGpBitmap PTR, BYVAL pGraphics AS CGpGraphics PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pBitmap* | Pointer to a **Bitmap** object that contains the pixel data to be optimized.  |
| *pGraphics* | Pointer to a **Graphics** object that is associated with a display device for which the image will be optimized. |

#### Remarks

You can display a cached bitmap by passing the address of a **CachedBitmap** object to the **DrawCachedBitmap** method of a Graphics object. Use the **Graphics** object that was passed to the **CachedBitmap** constructor or another **Graphics** object that represents the same device.

# <a name="ConstructorImageAttributes"></a>Constructor (CGpImageAttributes)

Creates a new **ImageAttributes** object. This constructor is the default constructor and has no parameters.

```
CONSTRUCTOR CGpImageAttributes
```

Creates and initializes an **ImageAttributes** object from another **ImageAttributes** object.

```
CONSTRUCTOR CGpImageAttributes (BYVAL pImgAttr AS CGpImageAttributes PTR)
```

# <a name="ClearBrushRemapTable"></a>ClearBrushRemapTable (CGpImageAttributes)

Clears the brush color-remap table of this **ImageAttributes** object.

```
FUNCTION ClearBrushRemapTable () AS GpStatus
```

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="ClearColorKey"></a>ClearColorKey (CGpImageAttributes)

Clears the color key (transparency range) for a specified category.

```
FUNCTION ClearColorKey (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Remarks

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify one color key for the default category, a different color key for the bitmap category, and still a different color key for the pen category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the pen category, then the default settings apply to the pen category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify a default color key that makes any color with a red component from 200 through 255 transparent and you specify a default gamma value of 1.8. If you set the color key of the pen category by calling **SetColorKey**, then the default color key and the default gamma value will not apply to pens. If you later clear the pen color key by calling **ClearColorKey**, the pen category will not revert to the default color key; rather, the pen category will have no color key. Similarly, the pen category will not revert to the default gamma value; rather, the pen category will have no gamma value.

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="ClearColorMatrices"></a>ClearColorMatrices (CGpImageAttributes)

Clears the color-adjustment matrix and the grayscale-adjustment matrix for a specified category.

```
FUNCTION ClearColorMatrices (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify a pair (color and grayscale) of adjustment matrices for the default category, a different pair of adjustment matrices for the bitmap category, and still a different pair of adjustment matrices for the pen category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the pen category, then the default settings apply to the pen category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify a pair (color and grayscale) of adjustment matrices and a gamma value for the default category. If you set a pair of adjustment matrices for the pen category by calling **SetColorMatrices**, then the default adjustment matrices will not apply to pens. If you later clear the pen adjustment matrices by calling **ClearColorMatrices**, the pen category will not revert to the default adjustment matrices; rather, the pen category will have no adjustment matrices. Similarly, the pen category will not revert to the default gamma value; rather, the pen category will have no gamma value.

# <a name="ClearColorMatrix"></a>ClearColorMatrix (CGpImageAttributes)

Clears the color-adjustment matrix for a specified category.

```
FUNCTION ClearColorMatrix (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify a color-adjustment matrix for the default category, a different color-adjustment matrix for the bitmap category, and still a different color-adjustment matrix for the pen category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the pen category, then the default settings apply to the pen category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify a color-adjustment matrix and a gamma value for the default category. If you set a color-adjustment matrix for the pen category by calling **SetColorMatrix**, then the default color-adjustment matrix will not apply to pens. If you later clear the pen color-adjustment matrix by calling **ClearColorMatrix**, the pen category will not revert to the default adjustment matrix; rather, the pen category will have no adjustment matrix. Similarly, the pen category will not revert to the default gamma value; rather, the pen category will have no gamma value.

# <a name="ClearGamma"></a>ClearGamma (CGpImageAttributes)

Disables gamma correction for a specified category.

```
FUNCTION ClearGamma (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify a gamma value for the default category, a different gamma value for the bitmap category, and still a different gamma value for the pen category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the pen category, then the default settings apply to the pen category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify a gamma value and a color-adjustment matrix for the default category. If you set the gamma value for the pen category by calling **SetGamma**, then the default gamma value will not apply to pens. If you later clear the pen gamma value by calling **ClearGamma**, the pen category will not revert to the default gamma value; rather, the pen category will have no gamma value. Similarly, the pen category will not revert to the default color-adjustment matrix; rather, the pen category will have no color-adjustment matrix.

# <a name="ClearNoOp"></a>ClearNoOp (CGpImageAttributes)

Clears the **NoOp** setting for a specified category.

```
FUNCTION ClearNoOp (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

You can disable color adjustment for a certain object type by calling the **SetNoOp** method. Later, you can reinstate color adjustment for that object type by calling the **ClearNoOp** method. For example, the following statement disables color adjustment for brushes:

```
myImageAttributes.SetNoOp(ColorAdjustTypeBrush)
```

The following statement reinstates the brush color adjustment that was in place before the call to **SetNoOp**:

```
myImageAttributes.ClearNoOp(ColorAdjustTypeBrush)
```

# <a name="ClearOutputChannel"></a>ClearOutputChannel (CGpImageAttributes)

Clears the cyan-magenta-yellow-black (CMYK) output channel setting for a specified category.

```
FUNCTION ClearOutputChannel (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify an output channel for the default category and a different output channel for the bitmap category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the bitmap category, then the default settings apply to the bitmap category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify an output channel and an adjustment matrix for the default category. If you set the output channel for the bitmap category by calling **SetOutputChannel**, then the default output channel will not apply to bitmaps. If you later clear the bitmap output channel by calling **ClearOutputChannel**, the bitmap category will not revert to the default output channel; rather, the bitmap category will have no output channel setting. Similarly, the bitmap category will not revert to the default color-adjustment matrix; rather, the bitmap category will have no color-adjustment matrix.

# <a name="ClearOutputChannelColorProfile"></a>ClearOutputChannelColorProfile (CGpImageAttributes)

Clears the output channel color profile setting for a specified category.

```
FUNCTION ClearOutputChannelColorProfile (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify an output channel profile for the default category and a different output channel profile for the bitmap category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the bitmap category, then the default settings apply to the bitmap category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify an output channel profile and an adjustment matrix for the default category. If you set the output channel profile for the bitmap category by calling **SetOutputChannelColorProfile**, then the default output channel profile will not apply to bitmaps. If you later clear the bitmap output channel profile by calling **ClearOutputChannelColorProfile**, the bitmap category will not revert to the default output channel profile; rather, the bitmap category will have no output channel profile setting. Similarly, the bitmap category will not revert to the default color-adjustment matrix; rather, the bitmap category will have no color-adjustment matrix.

# <a name="ClearRemapTable"></a>ClearRemapTable (CGpImageAttributes)

Clears the color-remap table for a specified category.

```
FUNCTION ClearRemapTable (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify a remap table for the default category, a different remap table for the bitmap category, and still a different remap table for the pen category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the pen category, then the default settings apply to the pen category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify a remap table and a gamma value for the default category. If you set the remap table for the pen category by calling **SetRemapTable**, then the default remap table will not apply to pens. If you later clear the pen remap table by calling **ClearRemapTable**, the pen category will not revert to the default remap table; rather, the pen category will have no remap table. Similarly, the pen category will not revert to the default gamma value; rather, the pen category will have no gamma value.

# <a name="ClearThreshold"></a>ClearThreshold (CGpImageAttributes)

Clears the threshold value for a specified category.

```
FUNCTION ClearThreshold (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

The threshold is a value from 0 through 1 that specifies a cutoff point for each color component. For example, suppose the threshold is set to 0.7, and suppose you are rendering a color whose red, green, and blue components are 230, 50, and 220. The red component, 230, is greater than 0.7×255, so the red component will be changed to 255 (full intensity). The green component, 50, is less than 0.7×255, so the green component will be changed to 0. The blue component, 220, is greater than 0.7×255, so the blue component will be changed to 255.

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify a threshold for the default category, a different threshold for the bitmap category, and still a different threshold for the pen category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the pen category, then the default settings apply to the pen category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify a threshold and a gamma value for the default category. If you set the threshold for the pen category by calling **SetThreshold**, then the default threshold will not apply to pens. If you later clear the pen threshold by calling **ClearThreshold**, the pen category will not revert to the default threshold; rather, the pen category will have no threshold. Similarly, the pen category will not revert to the default gamma value; rather, the pen category will have no gamma value.

# <a name="Clone"></a>Clone (CGpImageAttributes)

Copies the contents of the existing ImageAttributes object into a new **ImageAttributes** object.

```
FUNCTION Clone (BYVAL pCloneImgAttr AS CGpImageAttributes PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pCloneImgAttr* | Pointer to the **ImageAttributes** object where to copy the contents of the existing object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="GetAdjustedPalette"></a>GetAdjustedPalette (CGpImageAttributes)

Adjusts the colors in a palette according to the adjustment settings of a specified category.

```
FUNCTION GetAdjustedPalette (BYVAL pColorPalette AS ColorPalette PTR, BYVAL colorAdjustType AS LONG) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pColorPalette* | Pointer to a **ColorPalette** structure that on input, contains the palette to be adjusted and, on output, receives the adjusted palette. |
| *colorAdjustType* | Element of the **ColorAdjustType** enumeration that specifies the category whose adjustment settings will be applied to the palette. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify a color-remap table for the default category, a different color-remap table for the bitmap category, and still a different color-remap table for the pen category.

When you call **GetAdjustedPalette**, you can specify the adjustment category that is used to adjust the palette colors. For example, if you pass **ColorAdjustTypeBitmap** to the **GetAdjustedPalette** method, then the adjustment settings of the bitmap category are used to adjust the palette colors.

# <a name="Reset"></a>Reset (CGpImageAttributes)

Clears all color- and grayscale-adjustment settings for a specified category.

```
FUNCTION Reset (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify a gamma value for the default category, a different gamma value for the bitmap category, and still a different gamma value for the pen category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the pen category, then the default settings apply to the pen category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. You can reinstate the default settings for that category by calling **Reset**. For example, suppose you specify a gamma value for the default category. If you set the gamma value for the pen category by calling **SetGamma**, then the default gamma value will not apply to pens. If you later pass **ColorAdjustTypePen** to the **Reset** method, the pen category will revert to the default gamma value.

# <a name="SetBrushRemapTable"></a>SetBrushRemapTable (CGpImageAttributes)

Sets the color remap table for the brush category.

```
FUNCTION SetBrushRemapTable (BYVAL mapSize AS UINT, BYVAL map AS ANY PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *mapSize* | The number of elements in the map array. |
| *map* | Pointer to an array of **ColorMap** structures. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A color-remap table is an array of **ColorMap** structures. Each **ColorMap** structure has two colors: one that specifies an old color and one that specifies a corresponding new color. During rendering, any color that matches one of the old colors in the remap table is changed to the corresponding new color.

Calling the **SetBrushRemapTable** method has the same effect as passing **ColorAdjustTypeBrush** to the **SetRemapTable**. The specified remap table applies to items in metafiles that are filled with a brush.

# <a name="SetColorKey"></a>SetColorKey (CGpImageAttributes)

Sets the color key (transparency range) for a specified category.

```
FUNCTION SetColorKey (BYVAL colorLow AS ARGB, BYVAL colorHigh AS ARGB, _
   BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *colorLow* | ARGB color that specifies the low color-key value. |
| *colorHigh* | ARGB color that specifies the high color-key value. |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

This method sets the high and low color-key values so that arange of colors can be made transparent. Any color that has each of its three components (red, green, blue) between the corresponding components of the high and low color keys is made transparent.

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify a color key for the default category, a different color key for the bitmap category, and still a different color key for the pen category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the pen category, then the default settings apply to the pen category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify a collection of adjustment settings for the default category. If you set the color key for the pen category by passing **ColorAdjustTypePen** to the **SetColorKey** method, then none of the default adjustment settings will apply to pens.

# <a name="SetColorMatrices"></a>SetColorMatrices (CGpImageAttributes)

Sets the color-adjustment matrix and the grayscale-adjustment matrix for a specified category.

```
FUNCTION SetColorMatrices (BYVAL pColorMatrix AS ColorMatrix PTR, _
   BYVAL pGrayMatrix AS ColorMatrix PTR, _
   BYVAL nMode AS ColorAdjustType = ColorMatrixFlagsDefault, _
   BYVAL nType AS LONG = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pColorMatrix* | Pointer to a 5×5 color-adjustment matrix. |
| *pGrayMatrix* | Pointer to a 5×5 grayscale-adjustment matrix. |
| *nMode* | Optional. Element of the **ColorMatrixFlags** enumeration that specifies the type of image and color that will be affected by the color-adjustment and grayscale-adjustment matrices. The default value is **ColorMatrixFlagsDefault**. |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify adjustment matrices for the default category, different adjustment matrices for the bitmap category, and still different adjustment matrices for the pen category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the pen category, then the default settings apply to the pen category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify a collection of adjustment settings for the default category. If you set the color-adjustment and grayscale-adjustment matrices for the pen category by passing **ColorAdjustTypePen** to the **SetColorMatrices** method, then none of the default adjustment settings will apply to pens.

# <a name="SetColorMatrix"></a>SetColorMatrix (CGpImageAttributes)

Sets the color-adjustment matrix for a specified category.

```
FUNCTION SetColorMatrix (BYVAL pColorMatrix AS ColorMatrix PTR, _
   BYVAL nMode AS ColorMatrixFlags = ColorMatrixFlagsDefault, _
   BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pColorMatrix* | Pointer to a 5×5 color-adjustment matrix. |
| *nMode* | Optional. Element of the **ColorMatrixFlags** enumeration that specifies the type of image and color that will be affected by the color-adjustment and grayscale-adjustment matrices. The default value is **ColorMatrixFlagsDefault**. |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify a color-adjustment matrix for the default category, a different color-adjustment matrix for the bitmap category, and still a different color-adjustment matrix for the pen category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the pen category, then the default settings apply to the pen category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify a collection of adjustment settings for the default category. If you set the color-adjustment matrix for the pen category by passing **ColorAdjustTypePen** to the **SetColorMatrix** method, then none of the default adjustment settings will apply to pens.

# <a name="SetGamma"></a>SetGamma (CGpImageAttributes)

Sets the gamma value for a specified category.

```
FUNCTION SetGamma (BYVAL gamma AS SINGLE, BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *gamma* | The gamma value. |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify a gamma value for the default category, a different gamma value for the bitmap category, and still a different gamma value for the pen category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the pen category, then the default settings apply to the pen category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify a collection of adjustment settings for the default category. If you set the gamma value for the pen category by passing **ColorAdjustTypePen** to the **SetGamma** method, then none of the default adjustment settings will apply to pens.

# <a name="SetNoOp"></a>SetNoOp (CGpImageAttributes)

Turns off color adjustment for a specified category. You can call **ClearNoOp** to reinstate the color-adjustment settings that were in place before the call to **SetNoOp**.

```
FUNCTION SetNoOp (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *gamma* | The gamma value. |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="SetOutputChannel"></a>SetOutputChannel (CGpImageAttributes)

Sets the CMYK output channel for a specified category.

```
FUNCTION SetOutputChannel (BYVAL channelFlags AS LONG, _
   BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *channelFlags* | Element of the **ColorChannelFlags** enumeration that specifies the output channel. |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

You can use the **SetOutputChannel** method to convert an image to a cyan-magenta-yellow-black (CMYK) color space and examine the intensities of one of the CMYK color channels. For example, suppose you create an **ImageAttributes** object and set its bitmap output channel to **ColorChannelFlagsC**. If you pass the address of that **ImageAttributes** object to the **DrawImage** method, the cyan component of each pixel is calculated, and each pixel in the rendered image is a shade of gray that indicates the intensity of its cyan channel. Similarly, you can render images that indicate the intensities of the magenta, yellow, and black channels.

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify an output channel for the default category and a different output channel for the bitmap category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the bitmap category, then the default settings apply to the bitmap category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify a collection of adjustment settings for the default category. If you set the output channel for the bitmap category by passing **ColorAdjustTypeBitmap** to the **SetOutputChannel** method, then none of the default adjustment settings will apply to bitmaps.

# <a name="SetOutputChannelColorProfile"></a>SetOutputChannelColorProfile (CGpImageAttributes)

Sets the output channel color-profile file for a specified category.

```
FUNCTION SetOutputChannelColorProfile ( BYVAL pwszColorProfileFilename AS WSTRING PTR, _
   BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszColorProfileFilename* | Path name of a color-profile file. If the color-profile file is in the %SystemRoot%\System32\Spool\Drivers\Color directory, then this parameter can be the file name. Otherwise, this parameter must be the fully-qualified path name. |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

You can use the **SetOutputChannel** and **SetOutputChannelColorProfile** methods to convert an image to a cyan-magenta-yellow-black (CMYK) color space and examine the intensities of one of the CMYK color channels. For example, suppose you write code that performs the following steps:

* Create an **Image** object.
* Create an **ImageAttributes** object.
* Pass **ColorChannelFlagsC** to the **SetOutputChannel** method of the **ImageAttributes** object.
* Pass the path name of a color profile file to the **SetOutputChannelColorProfile** method of the **ImageAttributes** object.
* Pass the addresses of the Image and **ImageAttributes** objects to the **DrawImage** method.

Windows GDI+ will use the color-profile file to calculate the cyan component of each pixel in the image, and each pixel in the rendered image will be a shade of gray that indicates the intensity of its cyan channel.

An ImageAttributes object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify an output channel color-profile file for the default category and a different output channel color-profile file for the bitmap category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the bitmap category, then the default settings apply to the bitmap category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify a collection of adjustment settings for the default category. If you set the output channel color-profile file for the bitmap category by passing **ColorAdjustTypeBitmap** to the **SetOutputChannelColorProfile** method, then none of the default adjustment settings will apply to bitmaps.

# <a name="SetRemapTable"></a>SetRemapTable (CGpImageAttributes)

Sets the color-remap table for a specified category.

```
FUNCTION SetRemapTable (BYVAL mapSize AS UINT, BYVAL map AS ANY PTR, _
   BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *mapSize* | The number of elements in the map array. |
| *map* | Pointer to an array of **ColorMap** structures that defines the color map. |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A color-remap table is an array of **ColorMap** structures. Each **ColorMap** structure has two colors: one that specifies an old color and one that specifies a corresponding new color. During rendering, any color that matches one of the old colors in the remap table is changed to the corresponding new color.

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify a color remap for the default category, a color-remap table for the bitmap category, and still a different color-remap table for the pen category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the pen category, then the default settings apply to the pen category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify a collection of adjustment settings for the default category. If you set the color-remap table for the pen category by passing **ColorAdjustTypePen** to the **SetRemapTable** method, then none of the default adjustment settings will apply to pens.

# <a name="SetThreshold"></a>SetThreshold (CGpImageAttributes)

Sets the threshold (transparency range) for a specified category.

```
FUNCTION SetThreshold (BYVAL threshold AS SINGLE, _
   BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *threshold* | The threshold value. |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

The threshold is a value from 0 through 1 that specifies a cutoff point for each color component. For example, suppose the threshold is set to 0.7, and suppose you are rendering a color whose red, green, and blue components are 230, 50, and 220. The red component, 230, is greater than 0.7×255, so the red component will be changed to 255 (full intensity). The green component, 50, is less than 0.7×255, so the green component will be changed to 0. The blue component, 220, is greater than 0.7×255, so the blue component will be changed to 255.

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify a threshold for the default category, a threshold for the bitmap category, and still a different threshold for the pen category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the pen category, then the default settings apply to the pen category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify a collection of adjustment settings for the default category. If you set the threshold for the pen category by passing **ColorAdjustTypePen** to the **SetThreshold** method, then none of the default adjustment settings will apply to pens.

# <a name="SetToIdentity"></a>SetToIdentity (CGpImageAttributes)

Sets the color-adjustment matrix of a specified category to identity matrix.

```
FUNCTION SetToIdentity (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="SetWrapMode"></a>SetWrapMode (CGpImageAttributes)

Sets the the wrap mode of this **ImageAttributes** object.

```
FUNCTION SetWrapMode (BYVAL nWrap AS WrapMode, BYVAL colour AS ARGB = &HFF000000, _
   BYVAL clamp AS BOOL = FALSE) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nWrap* | Element of the **WrapMode** enumeration that specifies how repeated copies of an image are used to tile an area. |
| *colour* | Optional. ARGB color that specifies the color of pixels outside of a rendered image. This color is visible if the wrap parameter is set to **WrapModeClamp** and the source rectangle passed to DrawImage is larger than the image itself. |
| *clamp* | Optional. This parameter has no effect in GDI+ version 1.0. Set this parameter to FALSE. The default value is FALSE. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.
