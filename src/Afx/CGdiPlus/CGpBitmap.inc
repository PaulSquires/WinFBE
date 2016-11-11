' ########################################################################################
' Microsoft Windows
' File: CGpBitmap.inc
' Contents: GDI+ image/bitmap classes.
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

' ########################################################################################
' CGpImage Class
' ########################################################################################

' ========================================================================================
' * Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CGpImage
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' * Creates an Image object based on a file.
' ========================================================================================
PRIVATE CONSTRUCTOR CGpImage (BYVAL pwszFileName AS WSTRING PTR, BYVAL useEmbeddedColorManagement AS BOOLEAN = FALSE)
   IF useEmbeddedColorManagement THEN
      m_Status = GdipLoadImageFromFileICM(pwszFileName, @m_pImage)
   ELSE
      m_Status = GdipLoadImageFromFile(pwszFileName, @m_pImage)
   END IF
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates an Image object based on a stream.
' ========================================================================================
PRIVATE CONSTRUCTOR CGpImage (BYVAL pStream AS IStream PTR, BYVAL useEmbeddedColorManagement AS BOOLEAN = FALSE)
   IF useEmbeddedColorManagement THEN
      m_Status = GdipLoadImageFromStreamICM(pStream, @m_pImage)
   ELSE
      m_Status = GdipLoadImageFromStream(pStream, @m_pImage)
   END IF
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' * Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR CGpImage
   IF m_pImage THEN GdipDisposeImage(m_pImage)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last status code.
' ========================================================================================
PRIVATE FUNCTION CGpImage.GetLastStatus () AS GpStatus
   RETURN m_Status
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last status code.
' ========================================================================================
PRIVATE FUNCTION CGpImage.SetStatus (BYVAL status AS GpStatus) AS GpStatus
   m_Status = status
   RETURN m_Status
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the size, in bytes, of the parameter list for a specified image encoder
' ========================================================================================
PRIVATE FUNCTION CGpImage.GetEncoderParameterListSize (BYVAL clsidEncoder AS GUID PTR) AS UINT
   DIM uSize AS UINT
   m_Status = GdipGetEncoderParameterListSize(m_pImage, clsidEncoder, @uSize)
   RETURN uSize
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a list of the parameters supported by a specified image encoder.
' ========================================================================================
PRIVATE FUNCTION CGpImage.GetEncoderParameterList (BYVAL clsidEncoder AS GUID PTR, BYVAL uSize AS UINT, BYVAL buffer AS EncoderParameters PTR) AS GpStatus
   RETURN SetStatus(GdipGetEncoderParameterList(m_pImage, clsidEncoder, uSize, buffer))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Saves this image to a file.
' ========================================================================================
PRIVATE FUNCTION CGpImage.Save (BYVAL pwszFileName AS WSTRING PTR, BYVAL clsidEncoder AS GUID PTR, BYVAL encoderParams AS EncoderParameters PTR) AS GpStatus
   RETURN SetStatus(GdipSaveImageToFile(m_pImage, pwszFileName, clsidEncoder, encoderParams))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Saves this image to a stream.
' ========================================================================================
PRIVATE FUNCTION CGpImage.Save (BYVAL pStream AS IStream PTR, BYVAL clsidEncoder AS GUID PTR, BYVAL encoderParams AS EncoderParameters PTR) AS GpStatus
   RETURN SetStatus(GdipSaveImageToStream(m_pImage, pStream, clsidEncoder, encoderParams))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a frame to a file or stream specified in a previous call to the SaveToFile or
' SaveToStream methods. Use this method to save selected frames from a multiple-frame
' image to another multiple-frame image.
' ========================================================================================
PRIVATE FUNCTION CGpImage.SaveAdd (BYVAL encoderParams AS EncoderParameters PTR) AS GpStatus
   RETURN SetStatus(GdipSaveAdd(m_pImage, encoderParams))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Adds a frame to a file or stream specified in a previous call to the SaveToFile or
' SaveToStream methods.
' ========================================================================================
PRIVATE FUNCTION CGpImage.SaveAdd (BYVAL pNewImage AS CGpImage PTR, BYVAL encoderParams AS EncoderParameters PTR) AS GpStatus
   RETURN SetStatus(GdipSaveAddImage(m_pImage, pNewImage->m_pImage, encoderParams))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the type (bitmap or metafile) of this Image object.
' ========================================================================================
PRIVATE FUNCTION CGpImage.GetType () AS ImageType
   DIM nType AS ImageType
   m_Status = GdipGetImageType(m_pImage, @nType)
   RETURN nType
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the width and height of this image.
' ========================================================================================
PRIVATE FUNCTION CGpImage.GetPhysicalDimension (BYVAL psize AS SizeF PTR) AS GpStatus
   RETURN SetStatus(GdipGetImageDimension(m_pImage, @psize->Width, @psize->Height))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the bounding rectangle for this image.
' ========================================================================================
PRIVATE FUNCTION CGpImage.GetBounds (BYVAL srcRect AS RectF PTR, BYVAL srcUnit AS GpUnit PTR) AS GpStatus
   RETURN SetStatus(GdipGetImageBounds(m_pImage, srcRect, srcUnit))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the width, in pixels, of this image.
' ========================================================================================
PRIVATE FUNCTION CGpImage.GetWidth () AS UINT
   DIM uWidth AS UINT
   m_Status = GdipGetImageWidth(m_pImage, @uWidth)
   RETURN uWidth
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the height, in pixels, of this image.
' ========================================================================================
PRIVATE FUNCTION CGpImage.GetHeight () AS UINT
   DIM uHeight AS UINT
   m_Status = GdipGetImageHeight(m_pImage, @uHeight)
   RETURN uHeight
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the height, in pixels, of this image.
' ========================================================================================
PRIVATE FUNCTION CGpImage.GetHorizontalResolution () AS SINGLE
   DIM resolution AS SINGLE
   m_Status = GdipGetImageHorizontalResolution(m_pImage, @resolution)
   RETURN resolution
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the width, in pixels, of this image.
' ========================================================================================
PRIVATE FUNCTION CGpImage.GetVerticalResolution () AS SINGLE
   DIM resolution AS SINGLE
   m_Status = GdipGetImageVerticalResolution(m_pImage, @resolution)
   RETURN resolution
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a set of flags that indicate certain attributes of this Image object.
' ========================================================================================
PRIVATE FUNCTION CGpImage.GetFlags () AS UINT
   DIM flags AS UINT
   m_Status = GdipGetImageFlags(m_pImage, @flags)
   RETURN flags
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a globally unique identifier ( GUID) that identifies the format of this Image
' object. GUIDs that identify various file formats are defined in Gdiplusimaging.inc.
' ========================================================================================
PRIVATE FUNCTION CGpImage.GetRawFormat (BYVAL guidformat AS GUID PTR) AS GpStatus
   RETURN SetStatus(GdipGetImageRawFormat(m_pImage, guidformat))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the pixel format of this Image object.
' ========================================================================================
PRIVATE FUNCTION CGpImage.GetPixelFormat () AS PixelFormat
   DIM pxformat AS PixelFormat
   m_Status = GdipGetImagePixelFormat(m_pImage, @pxformat)
   RETURN pxformat
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the pixel format of this Image object.
' ========================================================================================
PRIVATE FUNCTION CGpImage.GetPaletteSize () AS INT_
   DIM nSize AS INT_
   m_Status = GdipGetImagePaletteSize(m_pImage, @nSize)
   RETURN nSize
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the ColorPalette of this Image object.
' ========================================================================================
PRIVATE FUNCTION CGpImage.GetPalette (BYVAL pal AS ColorPalette PTR, BYVAL nSize AS INT_) AS GpStatus
   RETURN SetStatus(GdipGetImagePalette(m_pImage, pal, nSize))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the ColorPalette of this Image object.
' ========================================================================================
PRIVATE FUNCTION CGpImage.SetPalette (BYVAL pal AS ColorPalette PTR) AS GpStatus
   RETURN SetStatus(GdipSetImagePalette(m_pImage, pal))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the number of frame dimensions in this Image object.
' ========================================================================================
PRIVATE FUNCTION CGpImage.GetFrameDimensionsCount () AS UINT
   DIM nCount AS UINT
   m_Status = GdipImageGetFrameDimensionsCount(m_pImage, @nCount)
   RETURN nCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the identifiers for the frame dimensions of this Image object.
' ========================================================================================
PRIVATE FUNCTION CGpImage.GetFrameDimensionsList (BYVAL dimensionIDs AS GUID PTR, BYVAL nCount AS UINT) AS GpStatus
   RETURN SetStatus(GdipImageGetFrameDimensionsList(m_pImage, dimensionIDs, nCount))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the number of frames in a specified dimension of this Image object.
' ========================================================================================
PRIVATE FUNCTION CGpImage.GetFrameCount (BYVAL dimensionID AS GUID PTR) AS UINT
   DIM nCount AS UINT
   m_Status = GdipImageGetFrameCount(m_pImage, dimensionID, @nCount)
   RETURN nCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Selects the frame in this Image object specified by a dimension and an index.
' ========================================================================================
PRIVATE FUNCTION CGpImage.SelectActiveFrame (BYVAL dimensionID AS GUID PTR, BYVAL frameIndex AS UINT) AS GpStatus
   RETURN SetStatus(GdipImageSelectActiveFrame(m_pImage, dimensionID, frameIndex))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Rotates and flips this image.
' ========================================================================================
PRIVATE FUNCTION CGpImage.RotateFlip (BYVAL rotateFlipType AS RotateFlipType) AS GpStatus
   RETURN SetStatus(GdipImageRotateFlip(m_pImage, rotateFlipType))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the number of properties (pieces of metadata) stored in this Image object.
' ========================================================================================
PRIVATE FUNCTION CGpImage.GetPropertyCount () AS UINT
   DIM numProperty AS UINT
   m_Status = GdipGetPropertyCount(m_pImage, @numProperty)
   RETURN numProperty
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the number of properties (pieces of metadata) stored in this Image object.
' ========================================================================================
PRIVATE FUNCTION CGpImage.GetPropertyIdList (BYVAL numOfProperty AS UINT, BYVAL list AS PROPID PTR) AS GpStatus
   RETURN SetStatus(GdipGetPropertyIdList(m_pImage, numOfProperty, list))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the size, in bytes, of a specified property item of this Image object.
' ========================================================================================
PRIVATE FUNCTION CGpImage.GetPropertyItemSize (BYVAL propId AS PROPID) AS UINT
   DIM nSize AS UINT
   m_Status = GdipGetPropertyItemSize(m_pImage, propId, @nSize)
   RETURN nSize
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a specified property item (piece of metadata) from this Image object.
' ========================================================================================
PRIVATE FUNCTION CGpImage.GetPropertyItem (BYVAL propId AS PROPID, BYVAL propSize AS UINT, BYVAL buffer AS PropertyItem PTR) AS GpStatus
   RETURN SetStatus(GdipGetPropertyItem(m_pImage, propId, propSize, buffer))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the total size, in bytes, of all the property items stored in this Image object.
' The GetPropertySize method also gets the number of property items stored in this Image object.
' ========================================================================================
PRIVATE FUNCTION CGpImage.GetPropertySize (BYVAL totalBufferSize AS UINT PTR, BYVAL numProperties AS UINT PTR) AS GpStatus
   RETURN SetStatus(GdipGetPropertySize(m_pImage, totalBufferSize, numProperties))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets all the property items (metadata) stored in the specified Image object.
' ========================================================================================
PRIVATE FUNCTION CGpImage.GetAllPropertyItems (BYVAL totalBufferSize AS UINT, BYVAL numProperties AS UINT, BYVAL allItems AS PropertyItem PTR) AS GpStatus
   RETURN SetStatus(GdipGetAllPropertyItems(m_pImage, totalBufferSize, numProperties, allItems))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets all the property items (metadata) stored in the specified Image object.
' ========================================================================================
PRIVATE FUNCTION CGpImage.RemovePropertyItem (BYVAL propId AS PROPID) AS GpStatus
   RETURN SetStatus(GdipRemovePropertyItem(m_pImage, propId))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets a property item (piece of metadata) for this Image object. If the item already
' exists, then its contents are updated; otherwise, a new item is added.
' ========================================================================================
PRIVATE FUNCTION CGpImage.SetPropertyItem (BYVAL pitem AS PropertyItem PTR) AS GpStatus
   RETURN SetStatus(GdipSetPropertyItem(m_pImage, pitem))
END FUNCTION
' ========================================================================================

#ifdef __FB_64BIT__
' // The following functions don't compile with the 32-bit compiler
'#if (GDIPVER >= &h0110)

' ========================================================================================
' Retrieves the description and the data size of the first metadata item in this Image object.
' ========================================================================================
PRIVATE FUNCTION CGpImage.FindFirstItem (BYVAL pitem AS ImageItemData PTR) AS GpStatus
   RETURN SetStatus(GdipFindFirstImageItem(m_pImage, pitem))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The GdipFindNextImageItem function is used along with the GdipFindFirstImageItem
' function to enumerate the metadata items stored in this Image object. The
' GdipFindNextImageItem function retrieves the description and the data size of the
' next metadata item in the Image object.
' ========================================================================================
PRIVATE FUNCTION CGpImage.FindNextItem (BYVAL pitem AS ImageItemData PTR) AS GpStatus
   RETURN SetStatus(GdipFindNextImageItem(m_pImage, pitem))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets one piece of metadata from the specified Image object.
' ========================================================================================
PRIVATE FUNCTION CGpImage.GetItemData (BYVAL pitem AS ImageItemData PTR) AS GpStatus
   RETURN SetStatus(GdipGetImageItemData(m_pImage, pitem))
END FUNCTION
' ========================================================================================

'#endif   ' //(GDIPVER >= 0x0110)
#endif

' ========================================================================================
' Copies the contents of the existing Image object into a new Image object.
' ========================================================================================
PRIVATE FUNCTION CGpImage.Clone (BYVAL pCloneImage AS CGpImage PTR) AS GpStatus
   IF pCloneImage = NULL THEN RETURN SetStatus(InvalidParameter)
   IF pCloneImage->m_pImage THEN GdipDisposeImage(pCloneImage->m_pImage)
   RETURN SetStatus(GdipCloneImage(m_pImage, @pCloneImage->m_pImage))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a thumbnail image from this Image object.
' ========================================================================================
PRIVATE FUNCTION CGpImage.GetThumbnailImage (BYVAL thumbWidth AS UINT, BYVAL thumbHeight AS UINT, BYVAL pThumbnail AS CGpImage PTR, BYVAL pCallback AS ANY PTR = NULL, BYVAL callbackData AS ANY PTR = NULL) AS GpStatus
   IF pThumbnail->m_pImage THEN GdipDisposeImage(pThumbnail->m_pImage)
   RETURN SetStatus(GdipGetImageThumbnail(m_pImage, thumbWidth, thumbHeight, _
      IIF(pThumbnail, @pThumbnail->m_pImage, NULL), pCallback, callbackData))
END FUNCTION
' ========================================================================================


' ########################################################################################
' CGpBitmap class
' ########################################################################################

' ========================================================================================
' * Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CGpBitmap
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates an Bitmap object based on a file.
' ========================================================================================
PRIVATE CONSTRUCTOR CGpBitmap (BYVAL pwszFileName AS WSTRING PTR, BYVAL useEmbeddedColorManagement AS BOOLEAN = FALSE)
   IF useEmbeddedColorManagement THEN
      m_Status = GdipCreateBitmapFromFileICM(pwszFileName, @cast(GpBitmap PTR, m_pImage))
   ELSE
      m_Status = GdipCreateBitmapFromFile(pwszFileName, @cast(GpBitmap PTR, m_pImage))
   END IF
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates an Bitmap object based on a stream.
' ========================================================================================
PRIVATE CONSTRUCTOR CGpBitmap (BYVAL pStream AS IStream PTR, BYVAL useEmbeddedColorManagement AS BOOLEAN = FALSE)
   IF useEmbeddedColorManagement THEN
      m_Status = GdipCreateBitmapFromStreamICM(pStream, @cast(GpBitmap PTR, m_pImage))
   ELSE
      m_Status = GdipCreateBitmapFromStream(pStream, @cast(GpBitmap PTR, m_pImage))
   END IF
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a Bitmap object based on an array of bytes along with size and format information.
' ========================================================================================
PRIVATE CONSTRUCTOR CGpBitmap (BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL stride AS INT_, BYVAL pxFormat AS PixelFormat, BYVAL scan0 AS UBYTE PTR)
   m_Status = GdipCreateBitmapFromScan0(nWidth, nHeight, stride, pxFormat, scan0, @cast(GpBitmap PTR, m_pImage))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a Bitmap object based on size and format information.
' ========================================================================================
PRIVATE CONSTRUCTOR CGpBitmap (BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL pxFormat AS PixelFormat)
   m_Status = GdipCreateBitmapFromScan0(nWidth, nHeight, 0, pxFormat, NULL, @cast(GpBitmap PTR, m_pImage))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a Bitmap object based on a Graphics object, a width, and a height.
' ========================================================================================
PRIVATE CONSTRUCTOR CGpBitmap (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG, BYVAL pTarget AS CGpGraphics PTR)
   m_Status = GdipCreateBitmapFromGraphics(nWidth, nHeight, pTarget->m_pGraphics, @cast(GpBitmap PTR, m_pImage))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a Bitmap object based on a DirectDraw surface. The Bitmap object maintains a
' reference to the DirectDraw surface until the Bitmap object is deleted or goes out of scope.
' ========================================================================================
#ifdef IDirectDrawSurface7
PRIVATE CONSTRUCTOR CGpBitmap (BYVAL surface AS IDirectDrawSurface7 PTR)
   m_Status = GdipCreateBitmapFromDirectDrawSurface(surface, @cast(GpBitmap PTR, m_pImage))
END CONSTRUCTOR
#endif
' ========================================================================================

' ========================================================================================
' Creates a Bitmap object based on a BITMAPINFO structure and an array of pixel data.
' ========================================================================================
PRIVATE CONSTRUCTOR CGpBitmap (BYVAL gdiBitmapInfo AS BITMAPINFO PTR, BYVAL gdiBitmapData AS ANY PTR)
   m_Status = GdipCreateBitmapFromGdiDib(gdiBitmapInfo, gdiBitmapData, @cast(GpBitmap PTR, m_pImage))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a Bitmap object based on a handle to a Windows Windows Graphics Device
' Interface (GDI) bitmap and a handle to a GDI palette.
' ========================================================================================
PRIVATE CONSTRUCTOR CGpBitmap (BYVAL hbm AS HBITMAP, BYVAL hPal AS HPALETTE)
   m_Status = GdipCreateBitmapFromHBITMAP(hbm, hpal, @cast(GpBitmap PTR, m_pImage))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a Bitmap object based on an icon.
' ========================================================================================
PRIVATE CONSTRUCTOR CGpBitmap (BYVAL hicon AS HICON)
   m_Status = GdipCreateBitmapFromHICON(hicon, @cast(GpBitmap PTR, m_pImage))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a Bitmap object based on an application or DLL instance handle and the name of a bitmap.
' ========================================================================================
PRIVATE CONSTRUCTOR CGpBitmap (BYVAL hInstance AS HINSTANCE, BYVAL pwszBitmapName AS WSTRING PTR)
   m_Status = GdipCreateBitmapFromResource(hInstance, pwszBitmapName, @cast(GpBitmap PTR, m_pImage))
END CONSTRUCTOR
' ========================================================================================

' =====================================================================================
' Creates a Windows Graphics Device Interface (GDI) bitmap from this Bitmap object.
' =====================================================================================
PRIVATE FUNCTION CGpBitmap.GetHBITMAP (BYVAL colorBackground AS ARGB, BYVAL hbmReturn AS HBITMAP PTR) AS GpStatus
   RETURN SetStatus(GdipCreateHBITMAPFromBitmap(cast(GpBitmap PTR, m_pImage), hbmReturn, colorBackground))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Creates an icon from this Bitmap object.
' =====================================================================================
PRIVATE FUNCTION CGpBitmap.GetHICON (BYVAL hiconReturn AS HICON PTR) AS GpStatus
   RETURN SetStatus(GdipCreateHICONFromBitmap(cast(GpBitmap PTR, m_pImage), hiconReturn))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Creates a new Bitmap object by copying a portion of this bitmap.
' =====================================================================================
PRIVATE FUNCTION CGpBitmap.Clone (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE, BYVAL pxFormat AS PixelFormat, BYVAL pCloneBitmap AS CGpBitmap PTR) AS GpStatus
   IF pCloneBitmap = NULL THEN RETURN SetStatus(InvalidParameter)
   IF pCloneBitmap->m_pImage THEN GdipDisposeImage(pCloneBitmap->m_pImage)
   RETURN SetStatus(GdipCloneBitmapArea(x, y, nWidth, nHeight, pxFormat, cast(GpBitmap PTR, m_pImage), _
          @cast(GpBitmap PTR, pCloneBitmap->m_pImage)))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpBitmap.Clone (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL pxFormat AS PixelFormat, BYVAL pCloneBitmap AS CGpBitmap PTR) AS GpStatus
   IF pCloneBitmap = NULL THEN RETURN SetStatus(InvalidParameter)
   IF pCloneBitmap->m_pImage THEN GdipDisposeImage(pCloneBitmap->m_pImage)
   RETURN SetStatus(GdipCloneBitmapAreaI(x, y, nWidth, nHeight, pxFormat, cast(GpBitmap PTR, m_pImage), _
          @cast(GpBitmap PTR, pCloneBitmap->m_pImage)))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpBitmap.Clone (BYVAL rcf AS GpRectF PTR, BYVAL pxFormat AS PixelFormat, BYVAL pCloneBitmap AS CGpBitmap PTR) AS GpStatus
   IF pCloneBitmap = NULL THEN RETURN SetStatus(InvalidParameter)
   IF pCloneBitmap->m_pImage THEN GdipDisposeImage(pCloneBitmap->m_pImage)
   RETURN SetStatus(GdipCloneBitmapArea(rcf->x, rcf->y, rcf->Width, rcf->Height, pxFormat, cast(GpBitmap PTR, m_pImage), _
          @cast(GpBitmap PTR, pCloneBitmap->m_pImage)))
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpBitmap.Clone (BYVAL rc AS GpRect PTR, BYVAL pxFormat AS PixelFormat, BYVAL pCloneBitmap AS CGpBitmap PTR) AS GpStatus
   IF pCloneBitmap = NULL THEN RETURN SetStatus(InvalidParameter)
   IF pCloneBitmap->m_pImage THEN GdipDisposeImage(pCloneBitmap->m_pImage)
   RETURN SetStatus(GdipCloneBitmapArea(rc->x, rc->y, rc->Width, rc->Height, pxFormat, cast(GpBitmap PTR, m_pImage), _
          @cast(GpBitmap PTR, pCloneBitmap->m_pImage)))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Locks a rectangular portion of this bitmap and provides a temporary buffer that you
' can use to read or write pixel data in a specified format. Any pixel data that you
' write to the buffer is copied to the Bitmap object when you call UnlockBits.
' =====================================================================================
PRIVATE FUNCTION CGpBitmap.LockBits (BYVAL rc AS GpRect PTR, BYVAL flags AS UINT, BYVAL pxFormat AS PixelFormat, BYVAL lockedBitmapData AS BitmapData PTR) AS GpStatus
   RETURN SetStatus(GdipBitmapLockBits(cast(GpBitmap PTR, m_pImage), rc, flags, pxFormat, lockedBitmapData))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Unlocks a portion of this bitmap that was previously locked by a call to LockBits.
' =====================================================================================
PRIVATE FUNCTION CGpBitmap.UnlockBits (BYVAL lockedBitmapData AS BitmapData PTR) AS GpStatus
   RETURN SetStatus(GdipBitmapUnlockBits(cast(GpBitmap PTR, m_pImage), lockedBitmapData))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the color of a specified pixel in this bitmap.
' =====================================================================================
PRIVATE FUNCTION CGpBitmap.GetPixel (BYVAL x AS LONG, BYVAL y AS LONG, BYVAL colour AS ARGB PTR) AS GpStatus
   RETURN SetStatus(GdipBitmapGetPixel(cast(GpBitmap PTR, m_pImage), x, y, colour))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the color of a specified pixel in this bitmap.
' =====================================================================================
PRIVATE FUNCTION CGpBitmap.SetPixel (BYVAL x AS LONG, BYVAL y AS LONG, BYVAL colour AS ARGB) AS GpStatus
   RETURN SetStatus(GdipBitmapSetPixel(cast(GpBitmap PTR, m_pImage), x, y, colour))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the resolution of this Bitmap object.
' =====================================================================================
PRIVATE FUNCTION CGpBitmap.SetResolution (BYVAL xdpi AS SINGLE, BYVAL ydpi AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipBitmapSetResolution(cast(GpBitmap PTR, m_pImage), xdpi, ydpi))
END FUNCTION
' =====================================================================================

#ifdef __FB_64BIT__
' // The following functions don't compile with the 32-bit compiler
'#if (GDIPVER >= &H0110)

' =====================================================================================
' Initializes a standard, optimal, or custom color palette.
' =====================================================================================
PRIVATE FUNCTION CGpBitmap.InitializePalette (BYVAL colourPalette AS ColorPalette PTR, BYVAL nPaletteType AS PaletteType, BYVAL optimalColors AS INT_, BYVAL useTransparentColor AS BOOL, BYVAL pBitmap AS CGpBitmap PTR) AS GpStatus
   RETURN SetStatus(GdipInitializePalette(colourPalette, nPaletteType, optimalColors, useTransparentColor, cast(GpBitmap PTR, pBitmap->m_pImage)))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts a bitmap to a specified pixel format. The original pixel data in the bitmap
' is replaced by the new pixel data.
' =====================================================================================
PRIVATE FUNCTION CGpBitmap.ConvertFormat (BYVAL pxFormat AS PixelFormat, BYVAL nDitherType AS DitherType, BYVAL nPaletteType AS PaletteType, BYVAL colourPalette AS ColorPalette PTR, BYVAL alphaThresholdPercent AS SINGLE) AS GpStatus
   RETURN SetStatus(GdipBitmapConvertFormat(cast(GpBitmap PTR, m_pImage), pxFormat, nDitherType, nPaletteType, colourPalette, alphaThresholdPercent))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns one or more histograms for specified color channels of this Bitmap object.
' =====================================================================================
PRIVATE FUNCTION CGpBitmap.GetHistogram (BYVAL nFormat AS HistogramFormat, BYVAL NumberOfEntries AS UINT, BYVAL channel0 AS UINT PTR, BYVAL channel1 AS UINT PTR, BYVAL channel2 AS UINT PTR, BYVAL channel3 AS UINT PTR) AS GpStatus
   RETURN SetStatus(GdipBitmapGetHistogram(cast(GpBitmap PTR, m_pImage), nFormat, NumberOfEntries, channel0, channel1, channel2, channel3))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the number of elements (in an array of UINTs) that you must allocate before you call the
' GetHistogram method of a Bitmap object.
' =====================================================================================
PRIVATE FUNCTION CGpBitmap.GetHistogramSize (BYVAL nFormat AS HistogramFormat, BYVAL NumberOfEntries AS UINT PTR) AS GpStatus
   RETURN SetStatus(GdipBitmapGetHistogramSize(nFormat, NumberOfEntries))
END FUNCTION
' =====================================================================================

'#endif ' (GDIPVER >= &H0110)
#endif

' ########################################################################################
' CGpCachedBitmap Class
' ########################################################################################

' =====================================================================================
' * Creates a CachedBitmap object based on a Bitmap object and a Graphics object. The
' cached bitmap takes the pixel data from the Bitmap object and stores it in a format
' that is optimized for the display device associated with the Graphics object.
' =====================================================================================
PRIVATE CONSTRUCTOR CGpCachedBitmap (BYVAL pBitmap AS CGpBitmap PTR, BYVAL pGraphics AS CGpGraphics PTR)
   IF m_pCachedBitmap THEN GdipDeleteCachedBitmap(m_pCachedBitmap)
   m_Status = GdipCreateCachedBitmap(IIF(pBitmap, cast(GpBitmap PTR, pBitmap->m_pImage), NULL), IIF(pGraphics, pGraphics->m_pGraphics, NULL), @m_pCachedBitmap)
END CONSTRUCTOR
' =====================================================================================

' =====================================================================================
' Cleanup
' =====================================================================================
PRIVATE DESTRUCTOR CGpCachedBitmap
   IF m_pCachedBitmap THEN GdipDeleteCachedBitmap(m_pCachedBitmap)
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
' Returns the last status code.
' ========================================================================================
PRIVATE FUNCTION CGpCachedBitmap.GetLastStatus () AS GpStatus
   RETURN m_Status
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last status code.
' ========================================================================================
PRIVATE FUNCTION CGpCachedBitmap.SetStatus (BYVAL status AS GpStatus) AS GpStatus
   m_Status = status
   RETURN m_Status
END FUNCTION
' ========================================================================================

' ########################################################################################
' CGpMetafile Class
' ########################################################################################

' =====================================================================================
' Creates a Windows GDI+ Metafile.Metafile object for playback based on a Windows
' Graphics Device Interface (GDI) Enhanced Metafile (EMF) file.
' =====================================================================================
PRIVATE CONSTRUCTOR CGpMetafile (BYVAL hEmf AS HENHMETAFILE, BYVAL deleteEmf AS BOOL = FALSE)
   m_Status = GdipCreateMetafileFromEmf(hEmf, deleteEmf, @cast(GpMetafile PTR, m_pImage))
END CONSTRUCTOR
' =====================================================================================

' =====================================================================================
' Creates a Metafile object for playback.
' =====================================================================================
PRIVATE CONSTRUCTOR CGpMetafile (BYVAL pwszFileName AS WSTRING PTR)
   m_Status = GdipCreateMetafileFromFile(pwszFileName, @cast(GpMetafile PTR, m_pImage))
END CONSTRUCTOR
' =====================================================================================

' =====================================================================================
' Creates a Metafile object for playback.
' =====================================================================================
PRIVATE CONSTRUCTOR CGpMetafile (BYVAL pwszFileName AS WSTRING PTR, BYVAL wmfPFH AS WmfPlaceableFileHeader PTR)
   m_Status = GdipCreateMetafileFromWmfFile(pwszFileName, wmfPFH, @cast(GpMetafile PTR, m_pImage))
END CONSTRUCTOR
' =====================================================================================

' =====================================================================================
' Creates a Metafile object from an IStream interface for playback.
' =====================================================================================
PRIVATE CONSTRUCTOR CGpMetafile (BYVAL pStream AS IStream PTR)
   m_Status = GdipCreateMetafileFromStream(pStream, @cast(GpMetafile PTR, m_pImage))
END CONSTRUCTOR
' =====================================================================================

' =====================================================================================
' Creates a Windows GDI+ Metafile object for recording. The format will be placeable metafile.
' =====================================================================================
PRIVATE CONSTRUCTOR CGpMetafile (BYVAL hWmf AS HMETAFILE, BYVAL wmfPFH AS WmfPlaceableFileHeader PTR, BYVAL deleteWmf AS BOOL = FALSE)
   m_Status = GdipCreateMetafileFromWmf(hWmf, deleteWmf, wmfPFH,  @cast(GpMetafile PTR, m_pImage))
END CONSTRUCTOR
' =====================================================================================

' =====================================================================================
' Creates a Metafile object for recording
' =====================================================================================
PRIVATE CONSTRUCTOR CGpMetafile (BYVAL referenceHdc AS HDC, BYVAL nType AS EmfType = EmfTypeEmfPlusDual, BYVAL description AS WSTRING PTR = NULL)
   m_Status = GdipRecordMetafile(referenceHdc, nType, NULL, MetafileFrameUnitGdi, description, @cast(GpMetafile PTR, m_pImage))
END CONSTRUCTOR
' =====================================================================================

' =====================================================================================
' Creates a Metafile object for recording
' =====================================================================================
PRIVATE CONSTRUCTOR CGpMetafile (BYVAL referenceHdc AS HDC, BYVAL frmRect AS GpRectF PTR, BYVAL frameUnit AS MetafileFrameUnit = MetafileFrameUnitGdi, BYVAL nType AS EmfType = EmfTypeEmfPlusDual, BYVAL description AS WSTRING PTR = NULL)
   m_Status = GdipRecordMetafile(referenceHdc, nType, frmRect, frameUnit, description, @cast(GpMetafile PTR, m_pImage))
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
PRIVATE CONSTRUCTOR CGpMetafile (BYVAL referenceHdc AS HDC, BYVAL frmRect AS GpRect PTR, BYVAL frameUnit AS MetafileFrameUnit = MetafileFrameUnitGdi, BYVAL nType AS EmfType = EmfTypeEmfPlusDual, BYVAL description AS WSTRING PTR = NULL)
   m_Status = GdipRecordMetafileI(referenceHdc, nType, frmRect, frameUnit, description, @cast(GpMetafile PTR, m_pImage))
END CONSTRUCTOR
' =====================================================================================

' =====================================================================================
' Creates a Metafile object for recording
' =====================================================================================
PRIVATE CONSTRUCTOR CGpMetafile (BYVAL pwszFileName AS WSTRING PTR, BYVAL referenceHdc AS HDC, BYVAL nType AS EmfType = EmfTypeEmfPlusDual, BYVAL description AS WSTRING PTR = NULL)
   m_Status = GdipRecordMetafileFileName(pwszFileName, referenceHdc, nType, NULL, MetafileFrameUnitGdi, description, @cast(GpMetafile PTR, m_pImage))
END CONSTRUCTOR
' =====================================================================================

' =====================================================================================
' Creates a Metafile object for recording
' =====================================================================================
PRIVATE CONSTRUCTOR CGpMetafile (BYVAL pwszFileName AS WSTRING PTR, BYVAL referenceHdc AS HDC, BYVAL frmRect AS GpRectF PTR, _
BYVAL frameUnit AS MetafileFrameUnit = MetafileFrameUnitGdi, BYVAL nType AS EmfType = EmfTypeEmfPlusDual, BYVAL description AS WSTRING PTR = NULL)
   m_Status = GdipRecordMetafileFileName(pwszFileName, referenceHdc, nType, frmRect, frameUnit, description, @cast(GpMetafile PTR, m_pImage))
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
PRIVATE CONSTRUCTOR CGpMetafile (BYVAL pwszFileName AS WSTRING PTR, BYVAL referenceHdc AS HDC, BYVAL frmRect AS GpRect PTR, _
BYVAL frameUnit AS MetafileFrameUnit = MetafileFrameUnitGdi, BYVAL nType AS EmfType = EmfTypeEmfPlusDual, BYVAL description AS WSTRING PTR = NULL)
   m_Status = GdipRecordMetafileFileNameI(pwszFileName, referenceHdc, nType, frmRect, frameUnit, description, @cast(GpMetafile PTR, m_pImage))
END CONSTRUCTOR
' =====================================================================================

' =====================================================================================
' Creates a Metafile object for recording to an IStream interface.
' =====================================================================================
PRIVATE CONSTRUCTOR CGpMetafile (BYVAL pStream AS IStream PTR, BYVAL referenceHdc AS HDC, BYVAL nType AS EmfType, BYVAL description AS WSTRING PTR)
   m_Status = GdipRecordMetafileStream(pStream, referenceHdc, nType, NULL, MetafileFrameUnitGdi, description, @cast(GpMetafile PTR, m_pImage))
END CONSTRUCTOR
' =====================================================================================

' =====================================================================================
' Creates a Metafile object for recording to an IStream interface.
' =====================================================================================
PRIVATE CONSTRUCTOR CGpMetafile (BYVAL pStream AS IStream PTR, BYVAL referenceHdc AS HDC, BYVAL frmRect AS GpRectF PTR, _
BYVAL frameUnit AS MetafileFrameUnit = MetafileFrameUnitGdi, BYVAL nType AS EmfType = EmfTypeEmfPlusDual, BYVAL description AS WSTRING PTR = NULL)
   m_Status = GdipRecordMetafileStream(pStream, referenceHdc, nType, frmRect, frameUnit, description, @cast(GpMetafile PTR, m_pImage))
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
PRIVATE CONSTRUCTOR CGpMetafile (BYVAL pStream AS IStream PTR, BYVAL referenceHdc AS HDC, BYVAL frmRect AS GpRect PTR, _
BYVAL frameUnit AS MetafileFrameUnit = MetafileFrameUnitGdi, BYVAL nType AS EmfType = EmfTypeEmfPlusDual, BYVAL description AS WSTRING PTR = NULL)
   m_Status = GdipRecordMetafileStreamI(pStream, referenceHdc, nType, frmRect, frameUnit, description, @cast(GpMetafile PTR, m_pImage))
END CONSTRUCTOR
' =====================================================================================

' =====================================================================================
' The GetMetafileHeader method gets the metafile header of this metafile.
' GdipGetMetafileHeaderFromWmf isn't declared in the 64-bit headers.
' =====================================================================================
'PRIVATE FUNCTION CGpMetafile.GetMetafileHeader (BYVAL hWmf AS HMETAFILE, BYVAL wmfPFH AS WmfPlaceableFileHeader PTR, BYVAL mh AS MetafileHeader PTR) AS GpStatus
'   RETURN SetStatus(GdipGetMetafileHeaderFromWmf(hWmf, wmfPFH, mh))
'END FUNCTION
' =====================================================================================
' =====================================================================================
' Gets the metafile header of the specified metafile.
' =====================================================================================
PRIVATE FUNCTION CGpMetafile.GetMetafileHeader (BYVAL hEmf AS HENHMETAFILE, BYVAL mh AS MetafileHeader PTR) AS GpStatus
   RETURN SetStatus(GdipGetMetafileHeaderFromEmf(hEmf, mh))
END FUNCTION
' =====================================================================================
' =====================================================================================
' Gets the metafile header of the specified file.
' =====================================================================================
PRIVATE FUNCTION CGpMetafile.GetMetafileHeader (BYVAL pwsFileName AS WSTRING PTR, BYVAL mh AS MetafileHeader PTR) AS GpStatus
   RETURN SetStatus(GdipGetMetafileHeaderFromFile(pwsFileName, mh))
END FUNCTION
' =====================================================================================
' =====================================================================================
' Gets the metafile header of the specified stream.
' =====================================================================================
PRIVATE FUNCTION CGpMetafile.GetMetafileHeader (BYVAL pStream AS IStream PTR, BYVAL mh AS MetafileHeader PTR) AS GpStatus
   RETURN SetStatus(GdipGetMetafileHeaderFromStream(pStream, mh))
END FUNCTION
' =====================================================================================
' =====================================================================================
' Gets the metafile header of this metafile. [Default method.]
' =====================================================================================
PRIVATE FUNCTION CGpMetafile.GetMetafileHeader (BYVAL mh AS MetafileHeader PTR) AS GpStatus
   RETURN SetStatus(GdipGetMetafileHeaderFromMetafile(cast(GpMetafile PTR, m_pImage), mh))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets a Windows handle to an Enhanced Metafile (EMF) file.
' // Once this method is called, the Metafile object is in an invalid state and can no
' // longer be used.  It is the responsiblity of the caller to invoke DeleteEnhMetaFile
' // to delete this hEmf.
' =====================================================================================
PRIVATE FUNCTION CGpMetafile.GetHENHMETAFILE () AS HENHMETAFILE
   DIM hEmf AS HENHMETAFILE
   m_Status = GdipGetHemfFromMetafile(cast(GpMetafile PTR, m_pImage), @hEmf)
   RETURN hEmf
END FUNCTION
' =====================================================================================

' =====================================================================================
' Plays a metafile record.
' // Used in conjuction with EnumerateMetafile to play an EMF+
' // The data must be DWORD aligned if it's an EMF or EMF+.  It must be
' // WORD aligned if it's a WMF.
' =====================================================================================
PRIVATE FUNCTION CGpMetafile.PlayRecord (BYVAL recordType AS EmfPlusRecordType, BYVAL flags AS UINT, BYVAL dataSize AS UINT, BYVAL pData AS BYTE PTR) AS GpStatus
   RETURN SetStatus(GdipPlayMetafileRecord(cast(GpMetafile PTR, m_pImage), recordType, flags, dataSize, pData))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Sets the resolution for certain brush bitmaps that are stored in this metafile.
' // If you're using a printer HDC for the metafile, but you want the
' // metafile rasterized at screen resolution, then use this API to set
' // the rasterization dpi of the metafile to the screen resolution,
' // e.g. 96 dpi or 120 dpi.
' =====================================================================================
PRIVATE FUNCTION CGpMetafile.SetDownLevelRasterizationLimit (BYVAL metafileRasterizationLimitDpi AS UINT) AS GpStatus
   RETURN SetStatus(GdipSetMetafileDownLevelRasterizationLimit(cast(GpMetafile PTR, m_pImage), metafileRasterizationLimitDpi))
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the rasterization limit currently set for this metafile. The rasterization limit
' is the resolution used for certain brush bitmaps that are stored in the metafile.
' =====================================================================================
PRIVATE FUNCTION CGpMetafile.GetDownLevelRasterizationLimit () AS UINT
   DIM metafileRasterizationLimitDpi AS UINT
   m_Status = GdipGetMetafileDownLevelRasterizationLimit(cast(GpMetafile PTR, m_pImage), @metafileRasterizationLimitDpi)
   RETURN metafileRasterizationLimitDpi
END FUNCTION
' =====================================================================================

' =====================================================================================
' Converts an enhanced-format metafile to a Windows Metafile Format (WMF) metafile and
' stores the converted records in a specified buffer.
' =====================================================================================
PRIVATE FUNCTION CGpMetafile.EmfToWmfBits (BYVAL hEmf AS HENHMETAFILE, BYVAL cbData16 AS UINT, BYVAL pData16 AS BYTE PTR, BYVAL iMapMode AS INT_ = MM_ANISOTROPIC, BYVAL eFlags AS INT_ = EmfToWmfBitsFlagsDefault) AS GpStatus
   RETURN SetStatus(GdipEmfToWmfBits(hEmf, cbData16, pData16, iMapMode, eFlags))
END FUNCTION
' =====================================================================================
