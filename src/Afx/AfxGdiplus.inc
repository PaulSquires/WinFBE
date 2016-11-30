' ########################################################################################
' Microsoft Windows
' File: AfxGdiplus.inc
' Content: Gdi Plus wrapper functions
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

#include once "windows.bi"
#ifdef __FB_64BIT__
    #inclib "gdiplus"
    #include once "win/gdiplus-c.bi"
#else
    #include once "win/ddraw.bi"
    #include once "win/gdiplus.bi"
    using gdiplus
#endif
#include once "win/winspool.bi"
#include once "Afx/AfxWin.inc"

NAMESPACE Afx

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

UNION GDIP_BGRA
   color AS COLORREF
   TYPE
      blue  AS UBYTE
      green AS UBYTE
      red   AS UBYTE
      alpha AS UBYTE
   END TYPE
END UNION

' // Another mess in the FB headers: declared as PathData, both as an structure and as a class.
' // And GpPathData data declared AS ANY PTR.
TYPE GDIP_PATHDATA
   Count AS LONG
   Points AS GpPointF PTR
   Types AS UBYTE PTR
END TYPE

' ========================================================================================
' Returns an ARGB color value initialized with the specified values for the alpha, red,
' green, and blue components.
' ========================================================================================
PRIVATE FUNCTION GDIP_ARGB (BYVAL a AS UBYTE, BYVAL r AS UBYTE, BYVAL g AS UBYTE, BYVAL b AS UBYTE) AS COLORREF
   DIM clr AS GDIP_BGRA
   clr.alpha = a : clr.red   = r : clr.green = g : clr.blue  = b
   FUNCTION  = clr.color
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GDIP_COLOR (BYVAL a AS UBYTE, BYVAL r AS UBYTE, BYVAL g AS UBYTE, BYVAL b AS UBYTE) AS COLORREF
   DIM clr AS GDIP_BGRA
   clr.alpha = a : clr.red   = r : clr.green = g : clr.blue  = b
   FUNCTION  = clr.color
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns an XRGB color value initialized with the specified values for the red, green,
' and blue components.
' ========================================================================================
PRIVATE FUNCTION GDIP_XRGB (BYVAL r AS UBYTE, BYVAL g AS UBYTE, BYVAL b AS UBYTE) AS COLORREF
   FUNCTION = GDIP_ARGB(&HFF, r, g, b)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a BGRA color value initialized with the specified values for the  blue, green,
' red and alpha components.
' ========================================================================================
PRIVATE FUNCTION GDIP_BGRA (BYVAL b AS UBYTE, BYVAL g AS UBYTE, BYVAL r AS UBYTE, BYVAL a AS UBYTE) AS COLORREF
   DIM clr AS GDIP_BGRA
   clr.blue  = b : clr.green = g : clr.red   = r : clr.alpha = a
   FUNCTION  = clr.color
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns an RGBA color value initialized with the specified values for the red, green,
' blue and alpha components.
' ========================================================================================
PRIVATE FUNCTION GDIP_RGBA (BYVAL r AS UBYTE, BYVAL g AS UBYTE, BYVAL b AS UBYTE, BYVAL a AS UBYTE) AS COLORREF
   FUNCTION = GDIP_ARGB(a, r, g, b)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the alpha component of an ARGB color.
' ========================================================================================
PRIVATE FUNCTION GDIP_GetAlpha (BYVAL argbcolor AS COLORREF) AS BYTE
   DIM clr AS GDIP_BGRA
   clr.color = argbcolor
   RETURN clr.alpha
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the red component of an ARGB color.
' ========================================================================================
PRIVATE FUNCTION GDIP_GetRed (BYVAL argbcolor AS COLORREF) AS BYTE
   DIM clr AS GDIP_BGRA
   clr.color = argbcolor
   RETURN clr.red
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the green component of an ARGB color.
' ========================================================================================
PRIVATE FUNCTION GDIP_GetGreen (BYVAL argbcolor AS COLORREF) AS BYTE
   DIM clr AS GDIP_BGRA
   clr.color = argbcolor
   RETURN clr.green
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the blue component of an ARGB color.
' ========================================================================================
PRIVATE FUNCTION GDIP_GetBlue (BYVAL argbcolor AS COLORREF) AS BYTE
   DIM clr AS GDIP_BGRA
   clr.color = argbcolor
   RETURN clr.blue
END FUNCTION
' ========================================================================================

' ========================================================================================
' Fills a RECTF structure
' ========================================================================================
PRIVATE FUNCTION GDIP_RECTF (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS GpRectF
   DIM rcf AS GpRectF
   rcf.x = x : rcf.y = y : rcf.Width = nWidth : rcf.Height = nHeight
   RETURN rcf
END FUNCTION
' ========================================================================================

' ========================================================================================
' Fills a RECT structure
' ========================================================================================
PRIVATE FUNCTION GDIP_RECT (BYVAL x AS LONG, BYVAL y AS LONG, BYVAL nWidth AS LONG, BYVAL nHeight AS LONG) AS GpRect
   DIM rc AS GpRect
   rc.x = x : rc.y = y : rc.Width = nWidth : rc.Height = nHeight
   RETURN rc
END FUNCTION
' ========================================================================================

' =====================================================================================
' Fills a POINTF structure
' =====================================================================================
PRIVATE FUNCTION GDIP_POINTF (BYVAL x AS SINGLE, BYVAL y AS SINGLE) AS GpPointF
   DIM pt AS GpPointF
   pt.x = x : pt.y = y
   RETURN pt
END FUNCTION
' =====================================================================================

' =====================================================================================
' Fills a POINTF structure
' =====================================================================================
PRIVATE FUNCTION GDIP_POINT (BYVAL x AS LONG, BYVAL y AS LONG) AS GpPoint
   DIM pt AS GpPoint
   pt.x = x : pt.y = y
   RETURN pt
END FUNCTION
' =====================================================================================

' ========================================================================================
' Returns the version of Gdiplus.dll, e.g. 601 for version 6.01.
' ========================================================================================
PRIVATE FUNCTION AfxGdipDllVersion () AS LONG
   FUNCTION = AfxGetFileVersion("GDIPLUS.DLL")
END FUNCTION
' ========================================================================================

' =====================================================================================
' Returns the description of a GdiPlus status code.
' =====================================================================================
PRIVATE FUNCTION AfxGdipStatusStr (BYVAL status AS GpStatus) AS CWSTR

   DIM wszMsg AS WSTRING * 260 = "Unknown"

   SELECT CASE status
      CASE Ok                        : wszMsg = "Ok"
      CASE GenericError              : wszMsg = "Generic error"
      CASE InvalidParameter          : wszMsg = "Invalid parameter"
      CASE OutOfMemory               : wszMsg = "Out of memory"
      CASE ObjectBusy                : wszMsg = "Object busy"
      CASE InsufficientBuffer        : wszMsg = "Insufficient buffer"
      CASE NotImplemented            : wszMsg = "Not implemented"
      CASE Win32Error                : wszMsg = "Win 32 error"
      CASE WrongState                : wszMsg = "Wrong state"
      CASE Aborted                   : wszMsg = "Aborted"
      CASE FileNotFound              : wszMsg = "File not found"
      CASE ValueOverflow             : wszMsg = "Value overflow"
      CASE AccessDenied              : wszMsg = "Access denied"
      CASE UnknownImageFormat        : wszMsg = "Unknown image format"
      CASE FontFamilyNotFound        : wszMsg = "Font family not found"
      CASE FontStyleNotFound         : wszMsg = "Font style not found"
      CASE NotTrueTypeFont           : wszMsg = "Not TrueType font"
      CASE UnsupportedGdiplusVersion : wszMsg = "Unsupported GdiPlus version"
      CASE GdiplusNotInitialized     : wszMsg = "GdiPlus not initialized"
      CASE PropertyNotFound          : wszMsg = "Property not found"
      CASE PropertyNotSupported      : wszMsg = "Property not supported"
'   #IF (GDIPVER >= &H0110)
'      CASE ProfileNotFound           : cws = "Profile not found"
'   #ENDIF   ' //(GDIPVER >= &H0110)
   END SELECT

   RETURN wszMsg

END FUNCTION
' =====================================================================================

' ========================================================================================
' Initilizes GDI+
' Returns a token or 0 if it fails.
' Pass the token to AfxGdipShutdown when you have finished using GDI+.
' ========================================================================================
PRIVATE FUNCTION AfxGdipInit (BYVAL version AS UINT32 = 1) AS ULONG_PTR
   DIM token AS ULONG_PTR, StartupInput AS GdiplusStartupInput
   StartupInput.GdiplusVersion = version
   GdiplusStartup(@token, @StartupInput, NULL)
   FUNCTION = token
END FUNCTION
' ========================================================================================

' ========================================================================================
' Cleans up resources used by Windows GDI+. Each call to AfxGdipInit should be paired with
' a call to AfxGdipShutdown
' ========================================================================================
PRIVATE SUB AfxGdipShutdown (BYVAL token AS ULONG_PTR)
   IF token THEN GdiplusShutdown(token)
END SUB
' ========================================================================================

' ========================================================================================
' Returns the size of the image.
' Parameters:
' - wszFileName = [in] Filename path.
' - nWidth      = [out] Width, in pixels, of the image.
' - nHeight     = [out] Height, in pixels, of the image.
' Return Values:
'   If the function succeeds, it returns Ok, which is an element of the Status enumeration.
'   If the function fails, it returns one of the other elements of the Status enumeration.
' ========================================================================================
PRIVATE FUNCTION AfxGdipGetImageSizeFromFile (BYREF wszFileName AS WSTRING, BYVAL nWidth AS DWORD PTR, BYVAL nHeight AS DWORD PTR) AS LONG
   IF LEN(wszFileName) = 0 THEN RETURN E_INVALIDARG
   ' // Initialize Gdiplus
   DIM token AS ULONG_PTR = AfxGdipInit
   IF token = NULL THEN EXIT FUNCTION
   DIM hr AS LONG, pImage AS GpImage PTR
   ' // Load the image
   hr = GdipLoadImageFromFile(wszFileName, @pImage)
   IF hr = 0 THEN
      hr = GdipGetImageWidth(pImage, nWidth)
      hr = GdipGetImageHeight(pImage, nHeight)
      hr = GdipDisposeImage(pImage)
   END IF
   FUNCTION = hr
   ' // Shutdown Gdiplus
   GdiplusShutdown token
END FUNCTION
' ========================================================================================

' ========================================================================================
' Loads an image from a file using GDI+, converts it to an icon or bitmap and returns the handle.
' Parameters:
' - wszFileName   = [in] Path of the image to load and convert.
' - dimPercent    = Percent of dimming (1-99)
' - bGrayScale    = TRUE or FALSE. Convert to gray scale.
' - imageType     = IMAGE_ICON or IMAGE_BITMAP.
' - clrBackground = [in] The background color. This parameter is ignored if the image type
'                   is IMAGE_ICON or the bitmap is totally opaque.
' Return Value:
'   If the function succeeds, the return value is the handle of the created icon or bitmap.
'   If the function fails, the return value is NULL.
' Remarks:
'   A quirk in the GDI+ GdipLoadImageFromFile function causes that dim gray images (often
'   used for disabled icons) are converted to darker shades of gray. Therefore, is better
'   to use AfxGdipImageFromFile.
' ========================================================================================
PRIVATE FUNCTION AfxGdipImageFromFile2 (BYREF wszFileName AS WSTRING, _
   BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE, _
   BYVAL imageType AS LONG = IMAGE_ICON, BYVAL clrBackground AS ARGB = 0) AS HANDLE

   DIM pImage AS GpImage PTR, hImage AS HANDLE
   DIM ImageWidth AS LONG, ImageHeight AS LONG, x AS LONG, y AS LONG
   DIM pixColor AS GDIP_BGRA, iColor AS LONG, rFactor AS SINGLE

   ' // Initialize Gdiplus
   DIM token AS ULONG_PTR = AfxGdipInit
   IF token = NULL THEN EXIT FUNCTION
   ' // Load the image from file
   GdipLoadImageFromFile(wszFileName, @pImage)
   IF pImage = NULL THEN EXIT FUNCTION
   ' // Get the image width and height
   GdipGetImageWidth(pImage, @ImageWidth)
   GdipGetImageHeight(pImage, @ImageHeight)
   ' // Dim or/and gray the image
   IF dimPercent > 0 AND dimPercent < 100 THEN rFactor = dimPercent / 100
   IF rFactor <> 0 OR bGrayScale <> 0 THEN
      FOR y = 0 TO ImageHeight - 1
         FOR x = 0 TO ImageWidth - 1
            ' // Get the pixel color
            GdipBitmapGetPixel(CAST(GpBitmap PTR, pImage), x, y, @pixColor.color)
            IF dimPercent > 0 THEN
               pixColor.red   = (255 - pixColor.red) * rFactor + pixColor.red
               pixColor.green = (255 - pixColor.green) * rFactor + pixColor.green
               pixColor.blue  = (255 - pixColor.blue) * rFactor + pixColor.blue
            END IF
            IF bGrayScale THEN
               ' Note: The sum of the percentages for the three colors should add to up 1
               iColor = 0.299 * pixColor.red + 0.587 * pixColor.green + 0.114 * pixColor.blue
               pixColor.Color = GDIP_BGRA (iColor, iColor, iColor, pixColor.alpha)
            ELSE
               pixColor.color = GDIP_ARGB(pixColor.alpha, pixColor.red, pixColor.green, pixColor.Blue)
            END IF
            ' // Set the modified pixel color
            GdipBitmapSetPixel(CAST(GpBitmap PTR, pImage), x, y, pixColor.color)
         NEXT
      NEXT
   END IF
   ' // Create icon from image
   IF imageType = IMAGE_ICON THEN
      GdipCreateHICONFromBitmap(CAST(GpBitmap PTR, pImage), @hImage)
   ELSE
      GdipCreateHBITMAPFromBitmap(CAST(GpBitmap PTR, pImage), @hImage, clrBackground)
   END IF
   ' // Free the image
   IF pImage THEN GdipDisposeImage pImage
   ' // Shutdown Gdiplus
   GdiplusShutdown token
   ' // Return the handle of the icon
   FUNCTION = hImage

END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts an image stored in a buffer into an icon or bitmap and returns the handle.
' Parameters:
' - pBuffer       = [in] Pointer to the buffer
' - bufferSize    = Size of the buffer
' - dimPercent    = Percent of dimming (1-99)
' - bGrayScale    = TRUE or FALSE. Convert to gray scale.
' - imageType     = IMAGE_ICON or IMAGE_BITMAP.
' - clrBackground = [in] The background color. This parameter is ignored if the image type
'                   is IMAGE_ICON or the bitmap is totally opaque.
' Return Value:
'   If the function succeeds, the return value is the handle of the created icon or bitmap.
'   If the function fails, the return value is NULL.
' Usage example:
'   DIM wszFileName AS WSTRING * MAX_PATH
'   wszFileName = ExePath & "\arrow_left_256.png"
'   DIM bufferSize AS SIZE_T_
'   DIM nFile AS LONG
'   nFile = FREEFILE
'   OPEN wszFileName FOR BINARY AS nFile
'   IF ERR THEN EXIT FUNCTION
'   bufferSize = LOF(nFile)
'   DIM pBuffer AS UBYTE PTR
'   pBuffer = CAllocate(1, bufferSize)
'   GET #nFile, , *pBuffer, bufferSize
'   CLOSE nFile
'   IF pBuffer THEN
'      ImageList_ReplaceIcon(hImageList, -1, AfxGdipIconFromBuffer(pBuffer, ImageSize))
'      DeAllocate(pBuffer)
'   END IF
' ========================================================================================
PRIVATE FUNCTION AfxGdipImageFromBuffer (BYVAL pBuffer AS ANY PTR, BYVAL bufferSize AS SIZE_T_, _
   BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE, _
   BYVAL imageType AS LONG = IMAGE_ICON, BYVAL clrBackground AS ARGB = 0) AS HANDLE

   DIM pImage AS GpImage PTR, hImage AS HANDLE
   DIM pImageStream AS IStream PTR, hGlobal AS HGLOBAL, pGlobalBuffer AS LPVOID
   DIM ImageWidth AS LONG, ImageHeight AS LONG, x AS LONG, y AS LONG
   DIM pixColor AS GDIP_BGRA, iColor AS LONG, rFactor AS SINGLE

   ' // Initialize Gdiplus
   DIM token AS ULONG_PTR = AfxGdipInit
   IF token = NULL THEN EXIT FUNCTION
   ' // Allocate memory to hold the image
   hGlobal = GlobalAlloc(GMEM_MOVEABLE, bufferSize)
   IF hGlobal THEN
      ' // Lock the memory
      pGlobalBuffer = GlobalLock(hGlobal)
      IF pGlobalBuffer THEN
         ' // Copy the image from the binary string file to global memory
         CopyMemory(pGlobalBuffer, pBuffer, bufferSize)
         ' // Create an stream in global memory
         IF CreateStreamOnHGlobal(hGlobal, FALSE, @pImageStream) = S_OK THEN
            IF pImageStream THEN
               ' // Create a bitmap from the data contained in the stream
               GdipCreateBitmapFromStream(pImageStream, CAST(GpBitmap PTR PTR, @pImage))
               ' // Get the image width and height
               GdipGetImageWidth(pImage, @ImageWidth)
               GdipGetImageHeight(pImage, @ImageHeight)
               ' // Dim or/and gray the image
               IF dimPercent > 0 AND dimPercent < 100 THEN rFactor = dimPercent / 100
               IF rFactor <> 0 OR bGrayScale <> 0 THEN
                  FOR y = 0 TO ImageHeight - 1
                     FOR x = 0 TO ImageWidth - 1
                        ' // Get the pixel color
                        GdipBitmapGetPixel(CAST(GpBitmap PTR, pImage), x, y, @pixColor.color)
                        IF dimPercent > 0 THEN
                           pixColor.red   = (255 - pixColor.red) * rFactor + pixColor.red
                           pixColor.green = (255 - pixColor.green) * rFactor + pixColor.green
                           pixColor.blue  = (255 - pixColor.blue) * rFactor + pixColor.blue
                        END IF
                        IF bGrayScale THEN
                           ' Note: The sum of the percentages for the three colors should add to up 1
                           iColor = 0.299 * pixColor.red + 0.587 * pixColor.green + 0.114 * pixColor.blue
                           pixColor.Color = GDIP_BGRA (iColor, iColor, iColor, pixColor.alpha)
                        ELSE
                           pixColor.color = GDIP_ARGB(pixColor.alpha, pixColor.red, pixColor.green, pixColor.Blue)
                        END IF
                        ' // Set the modified pixel color
                        GdipBitmapSetPixel(CAST(GpBitmap PTR, pImage), x, y, pixColor.color)
                     NEXT
                  NEXT
               END IF
               ' // Create icon from image
               IF imageType = IMAGE_ICON THEN
                  GdipCreateHICONFromBitmap(CAST(GpBitmap PTR, pImage), @hImage)
               ELSE
                  GdipCreateHBITMAPFromBitmap(CAST(GpBitmap PTR, pImage), @hImage, clrBackground)
               END IF
               ' // Free the image
               IF pImage THEN GdipDisposeImage pImage
               pImageStream->lpVtbl->Release(pImageStream)
            END IF
         END IF
         ' // Unlock the memory
         GlobalUnlock pGlobalBuffer
      END IF
      ' // Free the memory
      GlobalFree hGlobal
   END IF

   ' // Shutdown Gdiplus
   GdiplusShutdown token
   ' // Return the handle of the icon
   FUNCTION = hImage

END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxGdipIconFromBuffer (BYVAL pBuffer AS ANY PTR, BYVAL bufferSize AS SIZE_T_, _
   BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE) AS HANDLE
   FUNCTION = AfxGdipImageFromBuffer(pBuffer, bufferSize, dimPercent, bGrayScale, IMAGE_ICON, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxGdipBitmapFromBuffer (BYVAL pBuffer AS ANY PTR, BYVAL bufferSize AS SIZE_T_, _
   BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE, BYVAL clrBackground AS ARGB = 0) AS HANDLE
   FUNCTION = AfxGdipImageFromBuffer(pBuffer, bufferSize, dimPercent, bGrayScale, IMAGE_BITMAP, clrBackground)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Loads an image from a file, converts it to an icon or bitmap and returns the handle.
' Parameters:
' - wszFileName   = [in] Path of the image to load and convert.
' - dimPercent    = Percent of dimming (1-99)
' - bGrayScale    = TRUE or FALSE. Convert to gray scale.
' - imageType     = IMAGE_ICON or IMAGE_BITMAP.
' - clrBackground = [in] The background color. This parameter is ignored if the image type
'                   is IMAGE_ICON or the bitmap is totally opaque.
' Return Value:
'   If the function succeeds, the return value is the handle of the created icon or bitmap.
'   If the function fails, the return value is NULL.
' Remarks:
' A quirk in the GDI+ GdipLoadImageFromFile function causes that dim gray images (often used
' for disabled icons) are converted to darker shades of gray. Therefore, instead of using it
' I'm getting here the image data opening the file with OPEN in binary mode.
' ========================================================================================
PRIVATE FUNCTION AfxGdipImageFromFile (BYREF wszFileName AS WSTRING, _
   BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE, _
   BYVAL imageType AS LONG = IMAGE_ICON, BYVAL clrBackground AS ARGB = 0) AS HANDLE

   DIM fd AS WIN32_FIND_DATAW

   ' // Check for the existence of the file
   IF LEN(wszFileName) = 0 THEN EXIT FUNCTION
   DIM hFind AS HANDLE = FindFirstFileW(@wszFileName, @fd)
   IF hFind = INVALID_HANDLE_VALUE THEN EXIT FUNCTION
   FindClose hFind
   ' // Make sure that is not a directory or a temporary file
   IF (fd.dwFileAttributes AND FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY OR _
      (fd.dwFileAttributes AND FILE_ATTRIBUTE_TEMPORARY) = FILE_ATTRIBUTE_TEMPORARY THEN
      EXIT FUNCTION
   END IF

   ' // Open the file and store its contents into a buffer
   DIM nFile AS LONG, bufferSize AS SIZE_T_
   nFile = FREEFILE
   OPEN wszFileName FOR BINARY AS nFile
   IF ERR THEN EXIT FUNCTION
   bufferSize = LOF(nFile)
   DIM pBuffer AS UBYTE PTR
   pBuffer = CAllocate(1, bufferSize)
   GET #nFile, , *pBuffer, bufferSize
   CLOSE nFile
   IF pBuffer THEN
      FUNCTION = AfxGdipImageFromBuffer(pBuffer, bufferSize, dimPercent, bGrayScale, imageType, clrBackground)
      DeAllocate(pBuffer)
   END IF

END FUNCTION
' ========================================================================================
' ========================================================================================
' Loads an image from a file, converts it to an icon and returns the handle.
' ========================================================================================
PRIVATE FUNCTION AfxGdipIconFromFile (BYREF wszFileName AS WSTRING, _
   BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE) AS HANDLE
   FUNCTION = AfxGdipImageFromFile(wszFileName, dimPercent, bGrayScale, IMAGE_ICON, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Loads an image from a file, converts it to a bitmap and returns the handle.
' ========================================================================================
PRIVATE FUNCTION AfxGdipBitmapFromFile (BYREF wszFileName AS WSTRING, _
   BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE, BYVAL clrBackground AS ARGB = 0) AS HANDLE
   FUNCTION = AfxGdipImageFromFile(wszFileName, dimPercent, bGrayScale, IMAGE_BITMAP, clrBackground)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Loads an image from a file, converts it to an icon and adds it to specified image list.
' Parameters:
' - hIml        = A handle to the image list.
' - wszFileName = [in] Path of the image to load and convert.
' - dimPercent  = Percent of dimming (1-99)
' - bGrayScale  = TRUE or FALSE. Convert to gray scale.
' Return value:
'   Returns the index of the image if successful, or -1 otherwise.
' ========================================================================================
PRIVATE FUNCTION AfxGdipAddIconFromFile (BYVAL hIml AS HIMAGELIST, BYREF wszFileName AS WSTRING, _
   BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE) AS LONG
   DIM hIcon AS HICON = AfxGdipImageFromFile(wszFileName, dimPercent, bGrayScale, IMAGE_ICON, 0)
   IF hIcon THEN FUNCTION = ImageList_ReplaceIcon(hIml, -1, hIcon)
   IF hIcon THEN DestroyIcon(hIcon)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Loads an image from a resource, converts it to an icon or bitmap and returns the handle.
' Parameter:
' - hInstance     = [in] A handle to the module whose portable executable file or an accompanying
'                   MUI file contains the resource. If this parameter is NULL, the function searches
'                   the module used to create the current process.
' - wszImageName  = [in] Name of the image in the resource file (.RES). If the image resource uses
'                   an integral identifier, wszImage should begin with a number symbol (#)
'                   followed by the identifier in an ASCII format, e.g., "#998". Otherwise,
'                   use the text identifier name for the image. Only images embedded as raw data
'                   (type RCDATA) are valid. These must be icons in format .png, .jpg, .gif, .tiff.
' - dimPercent    = Percent of dimming (1-99)
' - bGrayScale    = TRUE or FALSE. Convert to gray scale.
' - imageType     = IMAGE_ICON or IMAGE_BITMAP.
' - clrBackground = [in] The background color. This parameter is ignored if the image type
'                   is IMAGE_ICON or the bitmap is totally opaque.
' Return Value:
'   If the function succeeds, the return value is the handle of the created icon or bitmap.
'   If the function fails, the return value is NULL.
' ========================================================================================
PRIVATE FUNCTION AfxGdipImageFromRes (BYVAL hInstance AS HINSTANCE, BYREF wszImageName AS WSTRING, _
   BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE, _
   BYVAL imageType AS LONG = IMAGE_ICON, BYVAL clrBackground AS ARGB = 0) AS HANDLE

   DIM pImage AS GpImage PTR, hImage AS HANDLE
   DIM hRes AS HRSRC, pResData AS HRSRC, wID AS WORD, dwID AS DWORD, imageSize AS DWORD
   DIM pImageStream AS IStream PTR, hGlobal AS HGLOBAL, pGlobalBuffer AS LPVOID
   DIM ImageWidth AS LONG, ImageHeight AS LONG, x AS LONG, y AS LONG
   DIM pixColor AS GDIP_BGRA, iColor AS LONG, rFactor AS SINGLE

   ' // Initialize Gdiplus
   DIM token AS ULONG_PTR = AfxGdipInit
   IF token = NULL THEN EXIT FUNCTION

   ' // Find the resource
   IF LEFT(wszImageName, 1) = "#" THEN
      wID = VAL(MID(wszImageName, 2))
      dwID = MAKELONG(wID, 0)
      hRes = FindResourceW(hInstance, MAKEINTRESOURCEW(dwID), CAST(LPCWSTR, RT_RCDATA))
   ELSE
      hRes = FindResourceW(hInstance, wszImageName, CAST(LPCWSTR, RT_RCDATA))
   END IF
   IF hRes THEN
      ' // Retrieve the size of the image
      imageSize = SizeofResource(hInstance, hRes)
      IF imageSize THEN
         ' // Load the resource and get a pointer to the resource data.
         ' // Note: LockResource does not actually lock memory; it is just used
         ' // to obtain a pointer to the memory containing the resource data.
         pResData = LockResource(LoadResource(hInstance, hRes))
         IF pResData THEN
            ' // Allocate memory to hold the image
            hGlobal = GlobalAlloc(GMEM_MOVEABLE, imageSize)
            IF hGlobal THEN
               ' // Lock the memory
               pGlobalBuffer = GlobalLock(hGlobal)
               IF pGlobalBuffer THEN
                  ' // Copy the image from the binary string file to global memory
                  CopyMemory(pGlobalBuffer, pResData, imageSize)
                  ' // Create an stream in global memory
                  IF CreateStreamOnHGlobal(hGlobal, FALSE, @pImageStream) = S_OK THEN
                     IF pImageStream THEN
                        ' // Create a bitmap from the data contained in the stream
                        GdipCreateBitmapFromStream(pImageStream, CAST(GpBitmap PTR PTR, @pImage))
                        ' // Get the image width and height
                        GdipGetImageWidth(pImage, @ImageWidth)
                        GdipGetImageHeight(pImage, @ImageHeight)
                        ' // Dim or/and gray the image
                        IF dimPercent > 0 AND dimPercent < 100 THEN rFactor = dimPercent / 100
                        IF rFactor <> 0 OR bGrayScale <> 0 THEN
                           FOR y = 0 TO ImageHeight - 1
                              FOR x = 0 TO ImageWidth - 1
                                 ' // Get the pixel color
                                 GdipBitmapGetPixel(CAST(GpBitmap PTR, pImage), x, y, @pixColor.color)
                                 IF dimPercent > 0 THEN
                                    pixColor.red   = (255 - pixColor.red) * rFactor + pixColor.red
                                    pixColor.green = (255 - pixColor.green) * rFactor + pixColor.green
                                    pixColor.blue  = (255 - pixColor.blue) * rFactor + pixColor.blue
                                 END IF
                                 IF bGrayScale THEN
                                    ' Note: The sum of the percentages for the three colors should add tp up 1
                                    iColor = 0.299 * pixColor.red + 0.587 * pixColor.green + 0.114 * pixColor.blue
                                    pixColor.Color = GDIP_BGRA (iColor, iColor, iColor, pixColor.alpha)
                                 ELSE
                                    pixColor.color = GDIP_ARGB(pixColor.alpha, pixColor.red, pixColor.green, pixColor.Blue)
                                 END IF
                                 ' // Set the modified pixel color
                                 GdipBitmapSetPixel(CAST(GpBitmap PTR, pImage), x, y, pixColor.color)
                              NEXT
                           NEXT
                        END IF
                        ' // Create icon from image
                        IF imageType = IMAGE_ICON THEN
                           GdipCreateHICONFromBitmap(CAST(GpBitmap PTR, pImage), @hImage)
                        ELSE
                           GdipCreateHBITMAPFromBitmap(CAST(GpBitmap PTR, pImage), @hImage, clrBackground)
                        END IF
                        ' // Free the image
                        IF pImage THEN GdipDisposeImage pImage
                        pImageStream->lpVtbl->Release(pImageStream)
                     END IF
                  END IF
                  ' // Unlock the memory
                  GlobalUnlock pGlobalBuffer
               END IF
               ' // Free the memory
               GlobalFree hGlobal
            END IF
         END IF
      END IF
   END IF

   ' // Shutdown Gdiplus
   GdiplusShutdown token
   ' // Return the handle of the icon
   FUNCTION = hImage

END FUNCTION
' ========================================================================================
' ========================================================================================
' Loads an image from a resource file, converts it to an icon and returns the handle.
' ========================================================================================
PRIVATE FUNCTION AfxGdipIconFromRes (BYVAL hInstance AS HINSTANCE, BYREF wszImageName AS WSTRING, _
   BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE) AS HANDLE
   FUNCTION = AfxGdipImageFromRes(hInstance, wszImageName, dimPercent, bGrayScale, IMAGE_ICON, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Loads an image from a resource file, converts it to a bitmap and returns the handle.
' ========================================================================================
PRIVATE FUNCTION AfxGdipBitmapFromRes (BYVAL hInstance AS HINSTANCE, BYREF wszImageName AS WSTRING, _
   BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE, BYVAL clrBackground AS ARGB = 0) AS HANDLE
   FUNCTION = AfxGdipImageFromRes(hInstance, wszImageName, dimPercent, bGrayScale, IMAGE_BITMAP, clrBackground)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Loads an image from a resource file, converts it to an icon and adds it to specified image list.
' Parameters:
' - hIml         = A handle to the image list.
' - hInstance    = [in] A handle to the module whose portable executable file or an accompanying
'                  MUI file contains the resource. If this parameter is NULL, the function searches
'                  the module used to create the current process.
' - wszImageName = [in] Name of the image in the resource file (.RES). If the image resource uses
'                  an integral identifier, wszImage should begin with a number symbol (#)
'                  followed by the identifier in an ASCII format, e.g., "#998". Otherwise,
'                  use the text identifier name for the image. Only images embedded as raw data
'                  (type RCDATA) are valid. These must be icons in format .png, .jpg, .gif, .tiff.
' - dimPercent   = Percent of dimming (1-99)
' - bGrayScale   = TRUE or FALSE. Convert to gray scale.
' Return value:
'   Returns the index of the image if successful, or -1 otherwise.
' ========================================================================================
PRIVATE FUNCTION AfxGdipAddIconFromRes (BYVAL hIml AS HIMAGELIST, BYVAL hInstance AS HINSTANCE, _
   BYREF wszImageName AS WSTRING, BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE) AS LONG
   DIM hIcon AS HICON = AfxGdipImageFromRes(hInstance, wszImageName, dimPercent, bGrayScale, IMAGE_ICON, 0)
   IF hIcon THEN FUNCTION = ImageList_ReplaceIcon(hIml, -1, hIcon)
   IF hIcon THEN DestroyIcon(hIcon)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Helper function to retrieve the encoder's clsid
' Parameter:
' - wszMimeType = Mime type.
' Return value:
'   The encoder clsid.
' ========================================================================================
PRIVATE FUNCTION AfxGdipGetEncoderClsid (BYREF wszMimeType AS WSTRING) AS GUID
   ' // Initialize Gdiplus
   DIM token AS ULONG_PTR = AfxGdipInit
   IF token = 0 THEN EXIT FUNCTION
   ' // The GdipGetImageEncodersSize function gets the number of available image encoders
   ' // and the total size of the array of ImageCodecInfo structures
   DIM numEncoders AS DWORD, nSize AS DWORD
   GdipGetImageEncodersSize(@numEncoders, @nSize)
   DIM pImageCodecInfo AS ImageCodecInfo PTR
   pImageCodecInfo = CAllocate(nSize, 1)
   ' // Gets an array of ImageCodecInfo structures that contain information about the
   ' // available image encoders.
   DIM hr AS LONG = GdipGetImageEncoders(numEncoders, nSize, pImageCodecInfo)
   IF hr = 0 THEN
      DIM i AS LONG
      FOR i = 0 TO numEncoders - 1
         IF UCASE(*CAST(WSTRING PTR, pImageCodecInfo[i].MimeType)) = UCASE(wszMimeType) THEN
            FUNCTION = pImageCodecInfo[i].Clsid
            EXIT FOR
         END IF
      NEXT
   END IF
   ' // Shutdown Gdiplus
   GdiplusShutdown token
END FUNCTION
' ========================================================================================

' ========================================================================================
' Saves a GDI+ image to file.
' Parameters:
' - pImage = Pointer to the GDI+ image to save.
' - wszFileName = Path name for the image to be saved.
' - wszMimeType = Mime type.
'   "image/bmp" = Bitmap (.bmp)
'   "image/gif" = GIF (.gif)
'   "image/jpeg" = JPEG (.jpg)
'   "image/png" = PNG (.png)
'   "image/tiff" = TIFF (.tiff)
' Return value:
' If the method succeeds, it returns Ok, which is an element of the Status enumeration.
' If the method fails, it returns one of the other elements of the Status enumeration.
' ========================================================================================
PRIVATE FUNCTION AfxGdipSaveImageToFile (BYVAL pImage AS GpImage PTR, BYREF wszFileName AS WSTRING, BYREF wszMimeType AS WSTRING) AS LONG
   IF pImage = NULL OR LEN(TRIM(wszFileName)) = 0 THEN RETURN InvalidParameter
   ' // Initialize Gdiplus
   DIM token AS ULONG_PTR = AfxGdipInit
   IF token = 0 THEN EXIT FUNCTION
   DIM clsidEncoder AS GUID = AfxGdipGetEncoderClsid(wszMimeType)
   FUNCTION = GdipSaveImageToFile(pImage, @wszFileName, @clsidEncoder, NULL)
   ' // Shutdown Gdiplus
   GdiplusShutdown token
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxGdipSaveImageToBmp (BYVAL pImage AS GpImage PTR, BYREF wszFileName AS WSTRING) AS LONG
   FUNCTION = AfxGdipSaveImageToFile(pImage, wszFilename, "image/bmp")
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxGdipSaveImageToJpeg (BYVAL pImage AS GpImage PTR, BYREF wszFileName AS WSTRING) AS LONG
   FUNCTION = AfxGdipSaveImageToFile(pImage, wszFilename, "image/jpeg")
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxGdipSaveImageToPng (BYVAL pImage AS GpImage PTR, BYREF wszFileName AS WSTRING) AS LONG
   FUNCTION = AfxGdipSaveImageToFile(pImage, wszFilename, "image/png")
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxGdipSaveImageToGif (BYVAL pImage AS GpImage PTR, BYREF wszFileName AS WSTRING) AS LONG
   FUNCTION = AfxGdipSaveImageToFile(pImage, wszFilename, "image/gif")
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxGdipSaveImageToTiff (BYVAL pImage AS GpImage PTR, BYREF wszFileName AS WSTRING) AS LONG
   FUNCTION = AfxGdipSaveImageToFile(pImage, wszFilename, "image/tiff")
END FUNCTION
' ========================================================================================

' ========================================================================================
' Saves a Windows bitmap to file.
' Parameters:
' - hbmp = Handle to the Windows bitmap.
' - wszFileName = Path name for the image to be saved.
' - wszMimeType = Mime type (default: "image/bmp").
'   "image/bmp" = Bitmap (.bmp)
'   "image/gif" = GIF (.gif)
'   "image/jpeg" = JPEG (.jpg)
'   "image/png" = PNG (.png)
'   "image/tiff" = TIFF (.tiff)
' Return value:
' If the method succeeds, it returns Ok, which is an element of the Status enumeration.
' If the method fails, it returns one of the other elements of the Status enumeration.
' ========================================================================================
PRIVATE FUNCTION AfxGdipSaveHBITMAPToFile (BYVAL hbmp AS HBITMAP, BYREF wszFileName AS WSTRING, BYREF wszMimeType AS WSTRING = "image/bmp") AS LONG
   IF hbmp = NULL OR LEN(TRIM(wszFileName)) = 0 THEN RETURN InvalidParameter
   ' // Initialize Gdiplus
   DIM token AS ULONG_PTR = AfxGdipInit
   ' // Create a Bitmap object from an HBITMAP
   DIM pBitmap AS GpBitmap PTR
   GdipCreateBitmapFromHBITMAP(hbmp, NULL, @pBitmap)
   IF pBitmap THEN
      DIM clsidEncoder AS GUID = AfxGdipGetEncoderClsid(wszMimeType)
      FUNCTION = GdipSaveImageToFile(CAST(GpImage PTR, pBitmap), @wszFileName, @clsidEncoder, NULL)
   END IF
   IF token = 0 THEN EXIT FUNCTION
   ' // Shutdown Gdiplus
   GdiplusShutdown token
END FUNCTION
' ========================================================================================

' ========================================================================================
' Prints a Windows bitmap in the default printer.
' Parameters:
' - hbmp         = Handle to the bitmap.
' - bStretch     = Stretch the image.
' - nStretchMode = Stretching mode. Default value = InterpolationModeHighQualityBicubic.
'   InterpolationModeLowQuality = 1
'   InterpolationModeHighQuality = 2
'   InterpolationModeBilinear = 3
'   InterpolationModeBicubic = 4
'   InterpolationModeNearestNeighbor = 5
'   InterpolationModeHighQualityBilinear = 6
'   InterpolationModeHighQualityBicubic = 7
' Return value: Returns TRUE if the bitmap has been printed successfully, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION AfxGdipPrintHBITMAP (BYVAL hbmp AS HBITMAP, BYVAL bStretch AS BOOLEAN = FALSE, BYVAL nStretchMode AS LONG = InterpolationModeHighQualityBicubic) AS LONG

   ' // Initialize Gdiplus
   DIM token AS ULONG_PTR = AfxGdipInit
   IF token = NULL THEN EXIT FUNCTION
   DIM hPrinter AS HANDLE
   DIM pbufferDoc AS UBYTE PTR

   ' // Note: Using a DO... LOOP, with an EXIT DO before LOOP,
   ' // to avoid the use of GOTO.
   DO
      ' // Get the name of the default printer
      DIM wszPrinterName AS WSTRING * MAX_PATH
      GetProfileStringW("WINDOWS", "DEVICE", "", wszPrinterName, SIZEOF(wszPrinterName))
      DIM nPos AS LONG = INSTR(wszPrinterName, ",")
      IF nPos THEN wszPrinterName = LEFT(wszPrinterName, nPos - 1)
      IF wszPrinterName = "" THEN EXIT DO
      ' // Open the printer
      IF OpenPrinterW(wszPrinterName, @hPrinter, NULL) = 0 THEN EXIT DO
      ' // Allocate a buffer of the correct size
      DIM dwNeeded AS DWORD = DocumentPropertiesW(NULL, hPrinter, wszPrinterName, NULL, NULL, 0)
      pbufferDoc = CAllocate(1, dwNeeded)
      ' // Retrieve the printer configuration data
      DIM nRet AS LONG = DocumentPropertiesW(NULL, hPrinter, wszPrinterName, CAST(PDEVMODEW, pbufferDoc), NULL, DM_OUT_BUFFER)
      IF nRet <> IDOK THEN EXIT DO
      ' // Cast it to a DEVMODEW structure
      DIM pDevMode AS DEVMODEW PTR = CAST(DEVMODEW PTR, pbufferDoc)
      ' // Create a device context for the printer
      DIM hDC AS .HDC = CreateDCW(wszPrinterName, wszPrinterName, NULL, pDevMode)
      IF hDC = NULL THEN EXIT DO
      ' // Create a graphics object from the printer DC
      DIM pGraphics AS GpGraphics PTR
      GdipCreateFromHDC(hDC, @pGraphics)
      IF pGraphics = NULL THEN EXIT DO
      ' // Create a Bitmap object from an HBITMAP
      DIM pBitmap AS GpBitmap PTR
      GdipCreateBitmapFromHBITMAP(hbmp, NULL, @pBitmap)
      IF pBitmap = NULL THEN EXIT DO
      ' // Stretching
      DIM cx AS SINGLE, cy AS SINGLE
      IF bStretch THEN GdipSetInterpolationMode(pGraphics, nStretchMode)
      ' // Get the DPIs of the printer
      DIM dpiX AS SINGLE, dpiY AS SINGLE
      GdipGetDpiX(pGraphics, @dpiX)
      GdipGetDpiY(pGraphics, @dpiY)
      ' // Calculate the width and height according to the DPIs of the printer
      cx = GetDeviceCaps(hdc, HORZRES) / (dpiX / 100)
      cy = GetDeviceCaps(hdc, VERTRES) / (dpiY / 100)
      ' // Print the bitmap
      DIM di AS DOCINFOW
      DIM wszDocName AS WSTRING * 260 = "Printing bitmap"
      di.cbSize = SIZEOF(DOCINFOW)
      di.lpszDocName = @wszDocName
      DIM hr AS LONG = StartDocW(hDC, @di)
      IF hr <= 0 THEN EXIT DO
      DIM i AS LONG
      FOR i = 1 TO pDevMode->dmCopies
         IF StartPage(hDC) THEN
            ' // Draw the image
            IF bStretch THEN
               GdipDrawImageRect(pGraphics, CAST(GpImage PTR, pBitmap), 0, 0, cx, cy)
            ELSE
               GdipDrawImage(pGraphics, CAST(GpImage PTR, pBitmap), 0, 0)
            END IF
            EndPage(hDC)
         END IF
      NEXT
      EndDoc(hDC)
      FUNCTION = TRUE
      EXIT DO
   LOOP

   ' // Finished with the printer
   IF hPrinter THEN ClosePrinter(hPrinter)
   ' // Cleanup
   IF pbufferDoc THEN Delete(pbufferDoc)
   ' // Shutdown Gdiplus
   GdiplusShutdown token

END FUNCTION
' ========================================================================================

END NAMESPACE
