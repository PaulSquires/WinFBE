' ########################################################################################
' Microsoft Windows
' File: CCMemBmp.inc
' Contents: Memory bitmap class.
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "Afx/AfxGdiplus.inc"

NAMESPACE Afx

' ========================================================================================
' CMemBmp class
' ========================================================================================
TYPE CMemBmp

   Public:
      m_hBmp AS HBITMAP                     ' // Bitmap handle
      m_hOldBmp AS HBITMAP                  ' // Old bitmap handle
      m_hMemDc AS HDC                       ' // Memory compatible device context handle

   Public:
      DECLARE CONSTRUCTOR (BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, BYVAL hdc AS HDC = NULL, BYVAL clrBkg AS COLORREF = BGR(255, 255, 255))
      DECLARE CONSTRUCTOR (BYREF wszFileName AS WSTRING, BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE, BYVAL clrBkg AS ARGB = 0)
      DECLARE CONSTRUCTOR (BYVAL hInst AS HINSTANCE, BYREF wszImageName AS WSTRING, BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE, BYVAL clrBkg AS ARGB = 0)
      DECLARE DESTRUCTOR
      DECLARE FUNCTION GetMemDC () AS HDC
      DECLARE FUNCTION GethBmp () AS HBITMAP
      DECLARE FUNCTION GetWidth () AS LONG
      DECLARE FUNCTION GetHeight () AS LONG
      DECLARE FUNCTION GetWidthBytes () AS LONG
      DECLARE FUNCTION GetPlanes () AS LONG
      DECLARE FUNCTION GetBitsPixel () AS LONG
      DECLARE FUNCTION GetPixel (BYVAL x AS LONG, BYVAL y AS LONG) AS COLORREF
      DECLARE FUNCTION SetPixel (BYVAL x AS LONG, BYVAL y AS LONG, BYVAL crColor AS COLORREF) AS COLORREF
'      DECLARE SUB DrawBitmap (BYVAL pBitmap AS HBITMAP, BYVAL x AS LONG = 0, BYVAL y AS LONG = 0)
      DECLARE FUNCTION DrawBitmap (BYVAL hbmp AS HBITMAP, BYVAL x AS SINGLE = 0, BYVAL y AS SINGLE = 0, _
              BYVAL nRight AS SINGLE = 0, BYVAL nBottom AS SINGLE = 0) AS GpStatus
      DECLARE FUNCTION DrawBitmap (BYVAL pBitmap AS GpBitmap PTR, BYVAL x AS SINGLE = 0, BYVAL y AS SINGLE = 0, _
              BYVAL nRight AS SINGLE = 0, BYVAL nBottom AS SINGLE = 0) AS GpStatus
      DECLARE FUNCTION SaveBitmap (BYREF wszFileName AS WSTRING, BYREF wszMimeType AS WSTRING = "image/bmp") AS LONG
      DECLARE FUNCTION SaveBitmapAsBmp (BYREF wszFileName AS WSTRING) AS LONG
      DECLARE FUNCTION SaveBitmapAsJpeg (BYREF wszFileName AS WSTRING) AS LONG
      DECLARE FUNCTION SaveBitmapAsPng (BYREF wszFileName AS WSTRING) AS LONG
      DECLARE FUNCTION SaveBitmapAsGif (BYREF wszFileName AS WSTRING) AS LONG
      DECLARE FUNCTION SaveBitmapAsTiff (BYREF wszFileName AS WSTRING) AS LONG
      DECLARE FUNCTION PrintBitmap (BYVAL bStretch AS BOOLEAN = FALSE, BYVAL nStretchMode AS LONG = InterpolationModeHighQualityBicubic) AS BOOLEAN

END TYPE
' ========================================================================================

' ========================================================================================
' CMemBmp constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CMemBmp (BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, BYVAL hdc AS .HDC = NULL, BYVAL clrBkg AS COLORREF = BGR(255, 255, 255))
   IF hdc THEN
      ' // Create a compatible bitmap
      m_hBmp = CreateCompatibleBitmap(hdc, nWidth, nHeight)
      ' // Create a compatible device context
      m_hMemDc = CreateCompatibleDC(hdc)
      ' // Select the bitmap into the compatible device context
      IF m_hMemDc THEN m_hOldBmp = SelectObject(m_hMemDc, m_hBmp)
   ELSE
      DIM hdcScreen AS .HDC = CreateDCW("DISPLAY", "", "", NULL)
      IF hdcScreen THEN
         ' // Create a compatible bitmap
         m_hBmp = CreateCompatibleBitmap(hdcScreen, nWidth, nHeight)
         ' // Create a compatible device context
         m_hMemDc = CreateCompatibleDC(hdcScreen)
         ' // Select the bitmap into the compatible device context
         IF m_hMemDc THEN m_hOldBmp = SelectObject(m_hMemDc, m_hBmp)
         DeleteDC hdcScreen
      END IF
   END IF
   ' // Paint the background of the bitmap
   DIM coord AS RECT
   SetRect(@coord, 0, 0, nWidth, nHeight)
   DIM hBrush AS HBRUSH = CreateSolidBrush(clrBkg)
   IF hBRush THEN
      FillRect(m_hMemDc, @coord, hBrush)
      DeleteObject hBrush
   END IF
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CMemBmp (BYREF wszFileName AS WSTRING, BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE, BYVAL clrBkg AS ARGB = 0)
   DIM nWidth AS LONG, nHeight AS LONG, bm AS BITMAP
   ' // Load the image as a bitmap
   DIM hbmp AS HBITMAP = AfxGdipBitmapFromFile(wszFileName, dimPercent, bGrayScale, clrBkg)
   IF hbmp THEN
      ' // Retrieve the width and height of the bitmap
      IF GetObject(hbmp, SIZEOF(BITMAP), @bm) THEN
         nWidth = bm.bmWidth
         nHeight = bm.bmHeight
      END IF
   END IF
   ' // Create a device context
   DIM hdcScreen AS .HDC = CreateDCW("DISPLAY", "", "", NULL)
   IF hdcScreen THEN
      ' // Create a compatible bitmap
      m_hBmp = CreateCompatibleBitmap(hdcScreen, nWidth, nHeight)
      ' // Create a compatible device context
      m_hMemDc = CreateCompatibleDC(hdcScreen)
      ' // Select the bitmap into the compatible device context
      IF m_hMemDc THEN m_hOldBmp = SelectObject(m_hMemDc, m_hBmp)
      DeleteDC hdcScreen
   END IF
   ' // Draw the bitmap
   IF hbmp THEN
      this.DrawBitmap(hbmp)
      DeleteObject hbmp
   END IF
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CMemBmp (BYVAL hInst AS HINSTANCE, BYREF wszImageName AS WSTRING, BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE, BYVAL clrBkg AS ARGB = 0)
   DIM nWidth AS LONG, nHeight AS LONG, bm AS BITMAP
   ' // Load the image as a bitmap
   DIM hbmp AS HBITMAP = AfxGdipBitmapFromRes(hInst, wszImageName, dimPercent, bGrayScale, clrBkg)
   IF hbmp THEN
      ' // Retrieve the width and height of the bitmap
      IF GetObject(hbmp, SIZEOF(BITMAP), @bm) THEN
         nWidth = bm.bmWidth
         nHeight = bm.bmHeight
      END IF
   END IF
   ' // Create a device context
   DIM hdcScreen AS .HDC = CreateDCW("DISPLAY", "", "", NULL)
   IF hdcScreen THEN
      ' // Create a compatible bitmap
      m_hBmp = CreateCompatibleBitmap(hdcScreen, nWidth, nHeight)
      ' // Create a compatible device context
      m_hMemDc = CreateCompatibleDC(hdcScreen)
      ' // Select the bitmap into the compatible device context
      IF m_hMemDc THEN m_hOldBmp = SelectObject(m_hMemDc, m_hBmp)
      DeleteDC hdcScreen
   END IF
   ' // Draw the bitmap
   IF hbmp THEN
      this.DrawBitmap(hbmp)
      DeleteObject hbmp
   END IF
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CMemBmp destructor
' ========================================================================================
PRIVATE DESTRUCTOR CMemBmp
   ' // Restore the original bitmap
   SelectObject(m_hMemDc, m_hOldBmp)
   ' // Destroys the bitmap and the memory device context
   IF m_hBmp THEN DeleteObject m_hBmp
   IF m_hMemDc THEN DeleteDC m_hMemDc
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the handle of the memory compatible device context.
' ========================================================================================
PRIVATE FUNCTION CMemBmp.GetMemDC () AS HDC
   FUNCTION = m_hMemDc
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the handle of the compatible bitmap.
' ========================================================================================
PRIVATE FUNCTION CMemBmp.GethBmp () AS HBITMAP
   FUNCTION = m_hBmp
END FUNCTION
' ========================================================================================

' ========================================================================================
' Draws a bitmap at the specified coodinates in the memory bitmap.
' Parameters:
' - x, y = Coordinates.
' Note: Does not work sometimes with handles returned by GdipCreateHBITMAPFromBitmap
' or even CMemBmp.
' ========================================================================================
'SUB CMemBmp.DrawBitmap (BYVAL pBitmap AS HBITMAP, BYVAL x AS LONG = 0, BYVAL y AS LONG = 0)
'   IF pBitmap = NULL THEN RETURN
'   DIM bm AS BITMAP, hMemDC AS HDC
'   hMemDC = CreateCompatibleDC(m_hMemDC)
'   IF hMemDC THEN
'      IF GetObject(pBitmap, SIZEOF(BITMAP), @bm) THEN
'         DIM hOldBmp AS HBITMAP = SelectObject(hMemDC, pBitmap)
'         BitBlt m_hMemDC, x, y, bm.bmWidth, bm.bmHeight, hMemDC, 0, 0, SRCCOPY
'         SelectObject(hMemDC, hOldBmp)
'      END IF
'      DeleteDC hMemDC
'   END IF
'END SUB
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CMemBmp.DrawBitmap (BYVAL hbmp AS HBITMAP, BYVAL x AS SINGLE = 0, BYVAL y AS SINGLE = 0, _
   BYVAL nRight AS SINGLE = 0, BYVAL nBottom AS SINGLE = 0) AS GpStatus

   DIM nStatus AS GpStatus, pGraphics AS GpGraphics PTR, pBitmap AS GpBitmap PTR
   DIM nWidth AS DWORD, nHeight AS DWORD
   DO
      nStatus = GdipCreateBitmapFromHBITMAP (hbmp, NULL, @pBitmap)
      IF pBitmap = NULL THEN EXIT DO
      GdipGetImageWidth(cast(GpImage PTR, pBitmap), @nWidth)
      GdipGetImageHeight(cast(GpImage PTR, pBitmap), @nHeight)
      nStatus = GdipCreateFromHDC(m_hMemDC, @pGraphics)
      IF pGraphics = NULL THEN EXIT DO
      IF nRight = 0 OR nBottom = 0 THEN
         GdipDrawImage(pGraphics, cast(GpImage PTR, pBitmap), x, y)
      ELSE
         GdipDrawImageRect(pGraphics, cast(GpImage PTR, pBitmap), x, y, nRight, nBottom)
      END IF
      EXIT DO
   LOOP
   IF pBitmap THEN GdipDisposeImage(cast(GpImage PTR, pBitmap))
   IF pGraphics THEN GdipDeleteGraphics(pGraphics)
   RETURN nStatus

END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CMemBmp.DrawBitmap (BYVAL pBitmap AS GpBitmap PTR, BYVAL x AS SINGLE = 0, BYVAL y AS SINGLE = 0, _
   BYVAL nRight AS SINGLE = 0, BYVAL nBottom AS SINGLE = 0) AS GpStatus

   IF pBitmap = NULL THEN RETURN 2   ' Status InvalidParameter
   DIM nStatus AS GpStatus, pGraphics AS GpGraphics PTR
   DIM nWidth AS DWORD, nHeight AS DWORD
   DO
      GdipGetImageWidth(cast(GpImage PTR, pBitmap), @nWidth)
      GdipGetImageHeight(cast(GpImage PTR, pBitmap), @nHeight)
      nStatus = GdipCreateFromHDC(m_hMemDC, @pGraphics)
      IF pGraphics = NULL THEN EXIT DO
      IF nRight = 0 OR nBottom = 0 THEN
         GdipDrawImage(pGraphics, cast(GpImage PTR, pBitmap), x, y)
      ELSE
         GdipDrawImageRect(pGraphics, cast(GpImage PTR, pBitmap), x, y, nRight, nBottom)
      END IF
      EXIT DO
   LOOP
   IF pGraphics THEN GdipDeleteGraphics(pGraphics)
   RETURN nStatus

END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the width of the bitmap in pixels.
' ========================================================================================
PRIVATE FUNCTION CMemBmp.GetWidth () AS LONG
   DIM bm AS BITMAP
   IF GetObject(m_hBmp, SIZEOF(bm), @bm) THEN FUNCTION = bm.bmWidth
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the height of the bitmap in pixels.
' ========================================================================================
PRIVATE FUNCTION CMemBmp.GetHeight () AS LONG
   DIM bm AS BITMAP
   IF GetObject(m_hBmp, SIZEOF(bm), @bm) THEN FUNCTION = bm.bmHeight
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the number of bytes in each scan line.
' ========================================================================================
PRIVATE FUNCTION CMemBmp.GetWidthBytes () AS LONG
   DIM bm AS BITMAP
   IF GetObject(m_hBmp, SIZEOF(bm), @bm) THEN FUNCTION = bm.bmWidthBytes
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the count of color planes.
' ========================================================================================
PRIVATE FUNCTION CMemBmp.GetPlanes () AS LONG
   DIM bm AS BITMAP
   IF GetObject(m_hBmp, SIZEOF(bm), @bm) THEN FUNCTION = bm.bmPlanes
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the number of bits required to indicate the color of a pixel.
' ========================================================================================
PRIVATE FUNCTION CMemBmp.GetBitsPixel () AS LONG
   DIM bm AS BITMAP
   IF GetObject(m_hBmp, SIZEOF(bm), @bm) THEN FUNCTION = bm.bmBitsPixel
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the red, green, blue (RGB) color value of the pixel at the specified coordinates.
' ========================================================================================
PRIVATE FUNCTION CMemBmp.GetPixel (BYVAL x AS LONG, BYVAL y AS LONG) AS COLORREF
   FUNCTION = .GetPixel(m_hMemDC, x, y)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the pixel at the specified coordinates to the specified color.
' ========================================================================================
PRIVATE FUNCTION CMemBmp.SetPixel (BYVAL x AS LONG, BYVAL y AS LONG, BYVAL crColor AS COLORREF) AS COLORREF
   FUNCTION = .SetPixel(m_hMemDC, x, y, crColor)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Saves the bitmap to a file.
' Parameters:
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
PRIVATE FUNCTION CMemBmp.SaveBitmap (BYREF wszFileName AS WSTRING, BYREF wszMimeType AS WSTRING = "image/bmp") AS LONG
   FUNCTION = AfxGdipSaveHBITMAPToFile(m_hBmp, wszFileName, wszMimeType)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CMemBmp.SaveBitmapAsBmp (BYREF wszFileName AS WSTRING) AS LONG
   FUNCTION = AfxGdipSaveHBITMAPToFile(m_hBmp, wszFileName, "image/bmp")
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CMemBmp.SaveBitmapAsJpeg (BYREF wszFileName AS WSTRING) AS LONG
   FUNCTION = AfxGdipSaveHBITMAPToFile(m_hBmp, wszFileName, "image/jpeg")
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CMemBmp.SaveBitmapAsPng (BYREF wszFileName AS WSTRING) AS LONG
   FUNCTION = AfxGdipSaveHBITMAPToFile(m_hBmp, wszFileName, "image/png")
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CMemBmp.SaveBitmapAsGif (BYREF wszFileName AS WSTRING) AS LONG
   FUNCTION = AfxGdipSaveHBITMAPToFile(m_hBmp, wszFileName, "image/gif")
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CMemBmp.SaveBitmapAsTiff (BYREF wszFileName AS WSTRING) AS LONG
   FUNCTION = AfxGdipSaveHBITMAPToFile(m_hBmp, wszFileName, "image/tiff")
END FUNCTION
' ========================================================================================

' ========================================================================================
' Prints the bitmap in the default printer.
' Parameters:
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
PRIVATE FUNCTION CMemBmp.PrintBitmap (BYVAL bStretch AS BOOLEAN = FALSE, BYVAL nStretchMode AS LONG = InterpolationModeHighQualityBicubic) AS BOOLEAN
   FUNCTION = AfxGdipPrintHBITMAP(m_hBmp, bStretch, nStretchMode)
END FUNCTION
' ========================================================================================

END NAMESPACE
