' ########################################################################################
' Microsoft Windows
' File: CGdiPlus.inc
' Contents: GDI+ wrapper classes master file
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "win/ddraw.bi"
#include once "Afx/AfxGdiPlus.inc"

TYPE CGpAdjustableArrowCap_ AS CGpAdjustableArrowCap
TYPE CGpBitmap_ AS CGpBitmap
TYPE CGpBrush_ AS CGpBrush
TYPE CGpCachedBitmap_ AS CGpCachedBitmap
TYPE CGpCustomLineCap_ AS CGpCustomLineCap
TYPE CGpFont_ AS CGpFont
TYPE CGpFontCollection_ AS CGpFontCollection
TYPE CGpFontFamily_ AS CGpFontFamily
TYPE CGpGraphics_ AS CGpGraphics
TYPE CGpGraphicsPath_ AS CGpGraphicsPath
TYPE CGpGraphicsPathIterator_ AS CGpGraphicsPathIterator
TYPE CGpImage_ AS CGpImage
TYPE CGpImageAttributes_ AS CGpImageAttributes
TYPE CGpMatrix_ AS CGpMatrix
TYPE CGpMetafile_ AS CGpMetafile
TYPE CGpPen_ AS CGpPen
TYPE CGpRegion_ AS CGpRegion
TYPE CGpStringFormat_ AS CGpStringFormat

' ########################################################################################
' CGpMatrix class
' ########################################################################################
TYPE CGpMatrix

Public:
   m_Status AS GpStatus
   m_pMatrix AS GpMatrix PTR

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL m11 AS SINGLE, BYVAL m12 AS SINGLE, BYVAL m21 AS SINGLE, BYVAL m22 AS SINGLE, BYVAL dx AS SINGLE, BYVAL dy AS SINGLE)
   DECLARE CONSTRUCTOR (BYVAL rcf AS GpRectF PTR, BYVAL dstplg AS GpPointF PTR)
   DECLARE CONSTRUCTOR (BYVAL rc AS GpRect PTR, BYVAL dstplg AS GpPoint PTR)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastStatus () AS GpStatus
   DECLARE FUNCTION SetStatus (BYVAL status AS GpStatus) AS GpStatus
   DECLARE FUNCTION Clone (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
   DECLARE FUNCTION GetElements (BYVAL m AS SINGLE PTR) AS GpStatus
   DECLARE FUNCTION SetElements (BYVAL m11 AS SINGLE, BYVAL m12 AS SINGLE, BYVAL m21 AS SINGLE, BYVAL m22 AS SINGLE, BYVAL dx AS SINGLE, BYVAL dy AS SINGLE) AS GpStatus
   DECLARE FUNCTION OffsetX () AS SINGLE
   DECLARE FUNCTION OffsetY () AS SINGLE
   DECLARE FUNCTION Reset () AS GpStatus
   DECLARE FUNCTION Multiply (BYVAL pMatrix AS CGpMatrix PTR, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   DECLARE FUNCTION Translate (BYVAL nOffsetX AS SINGLE, BYVAL nOffsetY AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   DECLARE FUNCTION Scale (BYVAL scaleX AS SINGLE, BYVAL scaleY AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   DECLARE FUNCTION Rotate (BYVAL angle AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   DECLARE FUNCTION RotateAt (BYVAL angle AS SINGLE, BYVAL center AS GpPointF PTR, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   DECLARE FUNCTION Shear (BYVAL shearX AS SINGLE, BYVAL shearY AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   DECLARE FUNCTION Invert () AS GpStatus
   DECLARE FUNCTION TransformPoints (BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_ = 1) AS GpStatus
   DECLARE FUNCTION TransformPoints (BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_ = 1) AS GpStatus
   DECLARE FUNCTION TransformVectors (BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_ = 1) AS GpStatus
   DECLARE FUNCTION TransformVectors (BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_ = 1) AS GpStatus
   DECLARE FUNCTION IsInvertible () AS BOOLEAN
   DECLARE FUNCTION IsIdentity () AS BOOLEAN
   DECLARE FUNCTION Equals (BYVAL pMatrix AS CGpMatrix PTR) AS BOOLEAN

END TYPE
' ########################################################################################

' ########################################################################################
' CGpImageAttributes class
' ########################################################################################
TYPE CGpImageAttributes

Public:
   m_Status AS GpStatus
   m_pImgAttr AS GpImageAttributes PTR

Public:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastStatus () AS GpStatus
   DECLARE FUNCTION SetStatus (BYVAL status AS GpStatus) AS GpStatus
   DECLARE FUNCTION Clone (BYVAL pCloneImgAttr AS CGpImageAttributes PTR) AS GpStatus
   DECLARE FUNCTION SetToIdentity (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   DECLARE FUNCTION Reset (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   DECLARE FUNCTION SetColorMatrix (BYVAL pColorMatrix AS ColorMatrix PTR, BYVAL nMode AS ColorMatrixFlags = ColorMatrixFlagsDefault, BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   DECLARE FUNCTION ClearColorMatrix (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   DECLARE FUNCTION SetColorMatrices (BYVAL pColorMatrix AS ColorMatrix PTR, BYVAL pGrayMatrix AS ColorMatrix PTR, BYVAL nMode AS ColorAdjustType = ColorMatrixFlagsDefault, BYVAL nType AS LONG = ColorAdjustTypeDefault) AS GpStatus
   DECLARE FUNCTION ClearColorMatrices (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   DECLARE FUNCTION SetThreshold (BYVAL threshold AS SINGLE, BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   DECLARE FUNCTION ClearThreshold (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   DECLARE FUNCTION SetGamma (BYVAL gamma AS SINGLE, BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   DECLARE FUNCTION ClearGamma (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   DECLARE FUNCTION SetNoOp (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   DECLARE FUNCTION ClearNoOp (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   DECLARE FUNCTION SetColorKey (BYVAL colorLow AS ARGB, BYVAL colorHigh AS ARGB, BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   DECLARE FUNCTION ClearColorKey (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   DECLARE FUNCTION SetOutputChannel (BYVAL channelFlags AS LONG, BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   DECLARE FUNCTION ClearOutputChannel (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   DECLARE FUNCTION SetOutputChannelColorProfile (BYVAL pwszColorProfileFilename AS WSTRING PTR, BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   DECLARE FUNCTION ClearOutputChannelColorProfile (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   ' // map parameter declared as ANY PTR in the two functions below because in the FB headers
   ' // it is declared as ColorMap PTR for 32 bit and ColorMap_ PTR for 64 bit
   DECLARE FUNCTION SetRemapTable (BYVAL mapSize AS UINT, BYVAL map AS ANY PTR, BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   DECLARE FUNCTION SetBrushRemapTable (BYVAL mapSize AS UINT, BYVAL map AS ANY PTR) AS GpStatus
   DECLARE FUNCTION ClearRemapTable (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
   DECLARE FUNCTION ClearBrushRemapTable () AS GpStatus
   DECLARE FUNCTION SetWrapMode (BYVAL nWrap AS WrapMode, BYVAL colour AS ARGB = &HFF000000, BYVAL clamp AS BOOL = FALSE) AS GpStatus
   DECLARE FUNCTION GetAdjustedPalette (BYVAL pColorPalette AS ColorPalette PTR, BYVAL colorAdjustType AS LONG) AS GpStatus

END TYPE
' ########################################################################################

' ########################################################################################
' CGpImage class
' ########################################################################################
TYPE CGpImage

Public:
   m_Status AS GpStatus
   m_pImage AS GpImage PTR

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pwszFileName AS WSTRING PTR, BYVAL useEmbeddedColorManagement AS BOOLEAN = FALSE)
   DECLARE CONSTRUCTOR (BYVAL pStream AS IStream PTR, BYVAL useEmbeddedColorManagement AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastStatus () AS GpStatus
   DECLARE FUNCTION SetStatus (BYVAL status AS GpStatus) AS GpStatus
   DECLARE FUNCTION FromFile (BYVAL pwszFileName AS WSTRING PTR, BYVAL useEmbeddedColorManagement AS BOOLEAN) AS CGpImage
   DECLARE FUNCTION FromStream (BYVAL pStream AS IStream PTR, BYVAL useEmbeddedColorManagement AS BOOLEAN) AS CGpImage
   DECLARE FUNCTION GetEncoderParameterListSize (BYVAL clsidEncoder AS GUID PTR) AS UINT
   DECLARE FUNCTION GetEncoderParameterList (BYVAL clsidEncoder AS GUID PTR, BYVAL uSize AS UINT, BYVAL buffer AS EncoderParameters PTR) AS GpStatus
   DECLARE FUNCTION Save (BYVAL pwszFileName AS WSTRING PTR, BYVAL clsidEncoder AS GUID PTR, BYVAL encoderParams AS EncoderParameters PTR) AS GpStatus
   DECLARE FUNCTION Save (BYVAL pStream AS IStream PTR, BYVAL clsidEncoder AS GUID PTR, BYVAL encoderParams AS EncoderParameters PTR) AS GpStatus
   DECLARE FUNCTION SaveAdd (BYVAL encoderParams AS EncoderParameters PTR) AS GpStatus
   DECLARE FUNCTION SaveAdd (BYVAL pNewImage AS CGpImage PTR, BYVAL encoderParams AS EncoderParameters PTR) AS GpStatus
   DECLARE FUNCTION GetType () AS ImageType
   DECLARE FUNCTION GetPhysicalDimension (BYVAL psize AS SizeF PTR) AS GpStatus
   DECLARE FUNCTION GetBounds (BYVAL srcRect AS RectF PTR, BYVAL srcUnit AS GpUnit PTR) AS GpStatus
   DECLARE FUNCTION GetWidth () AS UINT
   DECLARE FUNCTION GetHeight () AS UINT
   DECLARE FUNCTION GetHorizontalResolution () AS SINGLE
   DECLARE FUNCTION GetVerticalResolution () AS SINGLE
   DECLARE FUNCTION GetFlags () AS UINT
   DECLARE FUNCTION GetRawFormat (BYVAL guidformat AS GUID PTR) AS GpStatus
   DECLARE FUNCTION GetPixelFormat () AS PixelFormat
   DECLARE FUNCTION GetPaletteSize () AS INT_
   DECLARE FUNCTION GetPalette (BYVAL pal AS ColorPalette PTR, BYVAL nSize AS INT_) AS GpStatus
   DECLARE FUNCTION SetPalette (BYVAL pal AS ColorPalette PTR) AS GpStatus
   DECLARE FUNCTION GetFrameDimensionsCount () AS UINT
   DECLARE FUNCTION GetFrameDimensionsList (BYVAL dimensionIDs AS GUID PTR, BYVAL nCount AS UINT) AS GpStatus
   DECLARE FUNCTION GetFrameCount (BYVAL dimensionID AS GUID PTR) AS UINT
   DECLARE FUNCTION SelectActiveFrame (BYVAL dimensionID AS GUID PTR, BYVAL frameIndex AS UINT) AS GpStatus
   DECLARE FUNCTION RotateFlip (BYVAL rotateFlipType AS RotateFlipType) AS GpStatus
   DECLARE FUNCTION GetPropertyCount () AS UINT
   DECLARE FUNCTION GetPropertyIdList (BYVAL numOfProperty AS UINT, BYVAL list AS PROPID PTR) AS GpStatus
   DECLARE FUNCTION GetPropertyItemSize (BYVAL propId AS PROPID) AS UINT
   DECLARE FUNCTION GetPropertyItem (BYVAL propId AS PROPID, BYVAL propSize AS UINT, BYVAL buffer AS PropertyItem PTR) AS GpStatus
   DECLARE FUNCTION GetPropertySize (BYVAL totalBufferSize AS UINT PTR, BYVAL numProperties AS UINT PTR) AS GpStatus
   DECLARE FUNCTION GetAllPropertyItems (BYVAL totalBufferSize AS UINT, BYVAL numProperties AS UINT, BYVAL allItems AS PropertyItem PTR) AS GpStatus
   DECLARE FUNCTION RemovePropertyItem (BYVAL propId AS PROPID) AS GpStatus
   DECLARE FUNCTION SetPropertyItem (BYVAL pitem AS PropertyItem PTR) AS GpStatus
   DECLARE FUNCTION Clone (BYVAL pImage AS CGpImage PTR) AS GpStatus
   DECLARE FUNCTION GetThumbnailImage (BYVAL thumbWidth AS UINT, BYVAL thumbHeight AS UINT, BYVAL pThumbnail AS CGpImage PTR, BYVAL pCallback AS ANY PTR = NULL, BYVAL callbackData AS ANY PTR = NULL) AS GpStatus

#ifdef __FB_64BIT__
' // The following functions don't compile with the 32-bit compiler
'#if (GDIPVER >= &h0110)
   DECLARE FUNCTION FindFirstItem (BYVAL pitem AS ImageItemData PTR) AS GpStatus
   DECLARE FUNCTION FindNextItem (BYVAL pitem AS ImageItemData PTR) AS GpStatus
   DECLARE FUNCTION GetItemData (BYVAL pitem AS ImageItemData PTR) AS GpStatus
'#endif
#endif

END TYPE
' ########################################################################################

' ########################################################################################
' CGpBitmap class
' ########################################################################################
TYPE CGpBitmap EXTENDS CGpImage

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pwszFileName AS WSTRING PTR, BYVAL useEmbeddedColorManagement AS BOOLEAN = FALSE)
   DECLARE CONSTRUCTOR (BYVAL pStream AS IStream PTR, BYVAL useEmbeddedColorManagement AS BOOLEAN = FALSE)
   DECLARE CONSTRUCTOR (BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL stride AS INT_, BYVAL pxFormat AS PixelFormat, BYVAL scan0 AS UBYTE PTR)
   DECLARE CONSTRUCTOR (BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL pxFormat AS PixelFormat)
   DECLARE CONSTRUCTOR (BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL pTarget AS CGpGraphics_ PTR)
#ifdef IDirectDrawSurface7
   DECLARE CONSTRUCTOR (BYVAL surface AS IDirectDrawSurface7 PTR)
#endif
   DECLARE CONSTRUCTOR (BYVAL gdiBitmapInfo AS BITMAPINFO PTR, BYVAL gdiBitmapData AS ANY PTR)
   DECLARE CONSTRUCTOR (BYVAL hbm AS HBITMAP, BYVAL hPal AS HPALETTE)
   DECLARE CONSTRUCTOR (BYVAL hicon AS HICON)
   DECLARE CONSTRUCTOR (BYVAL hInstance AS HINSTANCE, BYVAL pwszbitmapName AS WSTRING PTR)
   DECLARE FUNCTION GetHBITMAP (BYVAL colorBackground AS ARGB, BYVAL hbmReturn AS HBITMAP PTR) AS GpStatus
   DECLARE FUNCTION GetHICON (BYVAL hiconReturn AS HICON PTR) AS GpStatus
   DECLARE FUNCTION Clone (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE, BYVAL pxFormat AS PixelFormat, BYVAL pCloneBitmap AS CGpBitmap PTR) AS GpStatus
   DECLARE FUNCTION Clone (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL pxFormat AS PixelFormat, BYVAL pCloneBitmap AS CGpBitmap PTR) AS GpStatus
   DECLARE FUNCTION Clone (BYVAL rcf AS GpRectF PTR, BYVAL pxFormat AS PixelFormat, BYVAL pCloneBitmap AS CGpBitmap PTR) AS GpStatus
   DECLARE FUNCTION Clone (BYVAL rc AS GpRect PTR, BYVAL pxFormat AS PixelFormat, BYVAL pCloneBitmap AS CGpBitmap PTR) AS GpStatus
   DECLARE FUNCTION LockBits (BYVAL rc AS GpRect PTR, BYVAL flags AS UINT, BYVAL pxFormat AS PixelFormat, BYVAL lockedBitmapData AS BitmapData PTR) AS GpStatus
   DECLARE FUNCTION UnlockBits (BYVAL lockedBitmapData AS BitmapData PTR) AS GpStatus
   DECLARE FUNCTION GetPixel (BYVAL x AS LONG, BYVAL y AS LONG, BYVAL colour AS ARGB PTR) AS GpStatus
   DECLARE FUNCTION SetPixel (BYVAL x AS LONG, BYVAL y AS LONG, BYVAL colour AS ARGB) AS GpStatus
   DECLARE FUNCTION SetResolution (BYVAL xdpi AS SINGLE, BYVAL ydpi AS SINGLE) AS GpStatus
#ifdef __FB_64BIT__
' // The following functions don't compile with the 32-bit compiler
   DECLARE FUNCTION ConvertFormat (BYVAL pxFormat AS PixelFormat, BYVAL nDitherType AS DitherType, BYVAL nPaletteType AS PaletteType, BYVAL colourPalette AS ColorPalette PTR, BYVAL alphaThresholdPercent AS SINGLE) AS GpStatus
   DECLARE FUNCTION InitializePalette (BYVAL colourPalette AS ColorPalette PTR, BYVAL nPaletteType AS PaletteType, BYVAL optimalColors AS INT_, BYVAL useTransparentColor AS BOOL, BYVAL pBitmap AS CGpBitmap PTR) AS GpStatus
   DECLARE FUNCTION GetHistogram (BYVAL nFormat AS HistogramFormat, BYVAL NumberOfEntries AS UINT, BYVAL channel0 AS UINT PTR, BYVAL channel1 AS UINT PTR, BYVAL channel2 AS UINT PTR, BYVAL channel3 AS UINT PTR) AS GpStatus
   DECLARE FUNCTION GetHistogramSize (BYVAL nFormat AS HistogramFormat, BYVAL NumberOfEntries AS UINT PTR) AS GpStatus
#endif

END TYPE
' ########################################################################################

' ########################################################################################
' CGpMetafile class
' ########################################################################################
TYPE CGpMetafile EXTENDS CGpImage

Public:
   DECLARE CONSTRUCTOR (BYVAL hWmf AS HMETAFILE, BYVAL wmfPFH AS WmfPlaceableFileHeader PTR, BYVAL deleteWmf AS BOOL = FALSE)
   DECLARE CONSTRUCTOR (BYVAL hEmf AS HENHMETAFILE, BYVAL deleteEmf AS BOOL = FALSE)
   DECLARE CONSTRUCTOR (BYVAL pwszFileName AS WSTRING PTR)
   DECLARE CONSTRUCTOR (BYVAL pwszFileName AS WSTRING PTR, BYVAL wmfPFH AS WmfPlaceableFileHeader PTR)
   DECLARE CONSTRUCTOR (BYVAL pStream AS IStream PTR)
   DECLARE CONSTRUCTOR (BYVAL referenceHdc AS HDC, BYVAL nType AS EmfType = EmfTypeEmfPlusDual, BYVAL description AS WSTRING PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL referenceHdc AS HDC, BYVAL frmRect AS GpRectF PTR, BYVAL frameUnit AS MetafileFrameUnit = MetafileFrameUnitGdi, BYVAL nType AS EmfType = EmfTypeEmfPlusDual, BYVAL description AS WSTRING PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL referenceHdc AS HDC, BYVAL frmRect AS GpRect PTR, BYVAL frameUnit AS MetafileFrameUnit = MetafileFrameUnitGdi, BYVAL nType AS EmfType = EmfTypeEmfPlusDual, BYVAL description AS WSTRING PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL pwszFileName AS WSTRING PTR, BYVAL referenceHdc AS HDC, BYVAL nType AS EmfType = EmfTypeEmfPlusDual, BYVAL description AS WSTRING PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL pwszFileName AS WSTRING PTR, BYVAL referenceHdc AS HDC, BYVAL frmRect AS GpRectF PTR, _
           BYVAL frameUnit AS MetafileFrameUnit = MetafileFrameUnitGdi, BYVAL nType AS EmfType = EmfTypeEmfPlusDual, BYVAL description AS WSTRING PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL pwszFileName AS WSTRING PTR, BYVAL referenceHdc AS HDC, BYVAL frmRect AS GpRect PTR, _
           BYVAL frameUnit AS MetafileFrameUnit = MetafileFrameUnitGdi, BYVAL nType AS EmfType = EmfTypeEmfPlusDual, BYVAL description AS WSTRING PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL pStream AS IStream PTR, BYVAL referenceHdc AS HDC, BYVAL nType AS EmfType, BYVAL description AS WSTRING PTR)
   DECLARE CONSTRUCTOR (BYVAL pStream AS IStream PTR, BYVAL referenceHdc AS HDC, BYVAL frmRect AS GpRectF PTR, _
           BYVAL frameUnit AS MetafileFrameUnit = MetafileFrameUnitGdi, BYVAL nType AS EmfType = EmfTypeEmfPlusDual, BYVAL description AS WSTRING PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL pStream AS IStream PTR, BYVAL referenceHdc AS HDC, BYVAL frmRect AS GpRect PTR, _
           BYVAL frameUnit AS MetafileFrameUnit = MetafileFrameUnitGdi, BYVAL nType AS EmfType = EmfTypeEmfPlusDual, BYVAL description AS WSTRING PTR = NULL)
'   DECLARE FUNCTION GetMetafileHeader (BYVAL hWmf AS HMETAFILE, BYVAL wmfPFH AS WmfPlaceableFileHeader PTR, BYVAL mh AS MetafileHeader PTR) AS GpStatus
   DECLARE FUNCTION GetMetafileHeader (BYVAL hEmf AS HENHMETAFILE, BYVAL mh AS MetafileHeader PTR) AS GpStatus
   DECLARE FUNCTION GetMetafileHeader (BYVAL pwsFileName AS WSTRING PTR, BYVAL mh AS MetafileHeader PTR) AS GpStatus
   DECLARE FUNCTION GetMetafileHeader (BYVAL pStream AS IStream PTR, BYVAL mh AS MetafileHeader PTR) AS GpStatus
   DECLARE FUNCTION GetMetafileHeader (BYVAL mh AS MetafileHeader PTR) AS GpStatus
   DECLARE FUNCTION GetHENHMETAFILE () AS HENHMETAFILE
   DECLARE FUNCTION PlayRecord (BYVAL recordType AS EmfPlusRecordType, BYVAL flags AS UINT, BYVAL dataSize AS UINT, BYVAL pData AS BYTE PTR) AS GpStatus
   DECLARE FUNCTION SetDownLevelRasterizationLimit (BYVAL metafileRasterizationLimitDpi AS UINT) AS GpStatus
   DECLARE FUNCTION GetDownLevelRasterizationLimit () AS UINT
   DECLARE FUNCTION EmfToWmfBits (BYVAL hEmf AS HENHMETAFILE, BYVAL cbData16 AS UINT, BYVAL pData16 AS BYTE PTR, BYVAL iMapMode AS INT_ = MM_ANISOTROPIC, BYVAL eFlags AS INT_ = EmfToWmfBitsFlagsDefault) AS GpStatus

END TYPE
' ########################################################################################

' ########################################################################################
' CGpCachedBitmap class
' ########################################################################################
TYPE CGpCachedBitmap

Public:
   m_Status AS GpStatus
   m_pCachedBitmap AS GpCachedBitmap PTR

Public:
   DECLARE CONSTRUCTOR (BYVAL pBitmap AS CGpBitmap PTR, BYVAL pGraphics AS CGpGraphics_ PTR)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastStatus () AS GpStatus
   DECLARE FUNCTION SetStatus (BYVAL status AS GpStatus) AS GpStatus

END TYPE
' ########################################################################################

' ########################################################################################
' CGpBrush class
' ########################################################################################
TYPE CGpBrush

Public:
   m_Status AS GpStatus
   m_pBrush AS GpBrush PTR

Public:
   ' // CGpBrush
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastStatus () AS GpStatus
   DECLARE FUNCTION SetStatus (BYVAL status AS GpStatus) AS GpStatus
   DECLARE FUNCTION Clone (BYVAL pBrush AS CGpBrush PTR) AS GpStatus
   DECLARE FUNCTION GetType () AS GpBrushType

END TYPE

' // CGpSolidBrush
TYPE CGpSolidBrush EXTENDS CGpBrush
   DECLARE CONSTRUCTOR (BYVAL colour AS ARGB = &hFF000000)
   DECLARE FUNCTION GetColor (BYVAL colour AS ARGB PTR) AS GpStatus
   DECLARE FUNCTION GetColor () AS ARGB
   DECLARE FUNCTION SetColor (BYVAL colour AS ARGB) AS GpStatus
END TYPE

' // CGpTextureBrush
TYPE CGpTextureBrush EXTENDS CGpBrush
   DECLARE CONSTRUCTOR (BYVAL pImage AS CGpImage PTR, BYVAL nWrapMode AS wrapMode = WrapModeTile)
   DECLARE CONSTRUCTOR (BYVAL pImage AS CGpImage PTR, BYVAL dstRect AS GpRectF PTR, BYVAL pImageAttributes AS CGpImageAttributes_ PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL pImage AS CGpImage PTR, BYVAL dstRect AS GpRect PTR, BYVAL pImageAttributes AS CGpImageAttributes_ PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL pImage AS CGpImage PTR, BYVAL nWrapMode AS wrapMode, BYVAL dstRect AS GpRect PTR)
   DECLARE CONSTRUCTOR (BYVAL pImage AS CGpImage PTR, BYVAL nWrapMode AS wrapMode, BYVAL dstRect AS GpRectF PTR)
   DECLARE CONSTRUCTOR (BYVAL pImage AS CGpImage PTR, BYVAL nWrapMode AS wrapMode, BYVAL dstX AS SINGLE, BYVAL dstY AS SINGLE, BYVAL dstWidth AS SINGLE, BYVAL dstHeight AS SINGLE)
   DECLARE CONSTRUCTOR (BYVAL pImage AS CGpImage ptR, BYVAL nWrapMode AS wrapMode, BYVAL dstX AS INT_, BYVAL dstY AS INT_, BYVAL dstWidth AS INT_, BYVAL dstHeight AS INT_)
   DECLARE FUNCTION SetTransform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
   DECLARE FUNCTION GetTransform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
   DECLARE FUNCTION ResetTransform () AS GpStatus
   DECLARE FUNCTION MultiplyTransform (BYVAL pMatrix AS CGpMatrix PTR, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   DECLARE FUNCTION TranslateTransform (BYVAL dx AS SINGLE, BYVAL dy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   DECLARE FUNCTION ScaleTransform (BYVAL sx AS SINGLE, BYVAL sy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   DECLARE FUNCTION RotateTransform (BYVAL angle AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   DECLARE FUNCTION SetWrapMode (BYVAL nWrapMode AS wrapMode) AS GpStatus
   DECLARE FUNCTION GetWrapMode () AS wrapMode
   DECLARE FUNCTION GetImage (BYVAL pImage AS CGpImage PTR) AS GpStatus
END TYPE

' // CGpLinearGradientBrush
TYPE CGpLinearGradientBrush EXTENDS CGpBrush
   DECLARE CONSTRUCTOR (BYVAL point1 AS GpPointF PTR, BYVAL point2 AS GpPointF PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB)
   DECLARE CONSTRUCTOR (BYVAL point1 AS GpPoint PTR, BYVAL point2 AS GpPoint PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB)
   DECLARE CONSTRUCTOR (BYVAL rc AS GpRectF PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, BYVAL nMode AS GpWrapMode)
   DECLARE CONSTRUCTOR (BYVAL rc AS GpRect PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, BYVAL nMode AS GpWrapMode)
   DECLARE CONSTRUCTOR (BYVAL rc AS GpRectF PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, BYVAL angle AS SINGLE, BYVAL isAngleScalable AS BOOL = FALSE)
   DECLARE CONSTRUCTOR (BYVAL rc AS GpRect PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, BYVAL angle AS SINGLE, BYVAL isAngleScalable AS BOOL = FALSE)
   DECLARE FUNCTION SetLinearColors (BYVAL color1 AS ARGB, BYVAL color2 AS ARGB) AS GpStatus
   DECLARE FUNCTION GetLinearColors (BYVAL colors AS ARGB PTR) AS GpStatus
   DECLARE FUNCTION GetRectangle (BYVAL rc AS GpRectF PTR) AS GpStatus
   DECLARE FUNCTION GetRectangle (BYVAL rc AS GpRect PTR) AS GpStatus
   DECLARE FUNCTION SetGammaCorrection (BYVAL useGammaCorrection AS BOOL) AS GpStatus
   DECLARE FUNCTION GetGammaCorrection () AS BOOLEAN
   DECLARE FUNCTION GetBlendCount () AS INT_
   DECLARE FUNCTION SetBlend (BYVAL blendFactors AS SINGLE PTR, BYVAL blendPositions AS SINGLE PTR, BYVAL nCount AS LONG) AS GpStatus
   DECLARE FUNCTION GetBlend (BYVAL blendFactors AS SINGLE PTR, BYVAL blendPositions AS SINGLE PTR, BYVAL nCount AS LONG) AS GpStatus
   DECLARE FUNCTION GetInterpolationColorCount () AS INT_
   DECLARE FUNCTION SetInterpolationColors (BYVAL presetColors AS ARGB PTR, BYVAL blendPositions AS SINGLE PTR, BYVAL nCount AS LONG) AS GpStatus
   DECLARE FUNCTION GetInterpolationColors (BYVAL presetColors AS ARGB PTR, BYVAL blendPositions AS SINGLE PTR, BYVAL nCount AS LONG) AS GpStatus
   DECLARE FUNCTION SetBlendBellShape (BYVAL focus AS SINGLE, BYVAL scale AS SINGLE = 1.0) AS GpStatus
   DECLARE FUNCTION SetBlendTriangularShape (BYVAL focus AS SINGLE, BYVAL scale AS SINGLE = 1.0) AS GpStatus
   DECLARE FUNCTION SetTransform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
   DECLARE FUNCTION GetTransform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
   DECLARE FUNCTION ResetTransform () AS GpStatus
   DECLARE FUNCTION MultiplyTransform (BYVAL pMatrix AS CGpMatrix PTR, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   DECLARE FUNCTION TranslateTransform (BYVAL dx AS SINGLE, BYVAL dy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   DECLARE FUNCTION ScaleTransform (BYVAL sx AS SINGLE, BYVAL sy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   DECLARE FUNCTION RotateTransform (BYVAL angle AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   DECLARE FUNCTION SetWrapMode (BYVAL nWrapMode AS WrapMode) AS GpStatus
   DECLARE FUNCTION GetWrapMode () AS WrapMode
END TYPE

' // CGpPathGradientBrush
TYPE CGpPathGradientBrush EXTENDS CGpBrush
   DECLARE CONSTRUCTOR (BYVAL pts AS GpPointF PTR, BYVAL nCount AS LONG, BYVAL nWrapMode AS WrapMode = WrapModeClamp)
   DECLARE CONSTRUCTOR (BYVAL pts AS GpPoint PTR, BYVAL nCount AS LONG, BYVAL nWrapMode AS WrapMode = WrapModeClamp)
   DECLARE CONSTRUCTOR (BYVAL pGraphPath AS CGpGraphicsPath_ PTR)
   DECLARE FUNCTION GetCenterColor (BYVAL colour AS ARGB PTR) AS GpStatus
   DECLARE FUNCTION SetCenterColor (BYVAL colour AS ARGB) AS GpStatus
   DECLARE FUNCTION GetPointCount () AS INT_
   DECLARE FUNCTION GetSurroundColorCount () AS INT_
   DECLARE FUNCTION GetSurroundColors (BYVAL colors AS ARGB PTR, BYVAL nCount AS INT_ PTR) AS GpStatus
   DECLARE FUNCTION SetSurroundColors (BYVAL colors AS ARGB PTR, BYVAL nCount AS INT_ PTR) AS GpStatus
   DECLARE FUNCTION GetGraphicsPath (BYVAL pGrapgPath AS CGpGraphicsPath_ PTR) AS GpStatus
   DECLARE FUNCTION SetGraphicsPath (BYVAL pGraphPath AS CGpGraphicsPath_ PTR) AS GpStatus
   DECLARE FUNCTION GetCenterPoint (BYVAL pt AS GpPointF PTR) AS GpStatus
   DECLARE FUNCTION GetCenterPoint (BYVAL pt AS GpPoint PTR) AS GpStatus
   DECLARE FUNCTION SetCenterPoint (BYVAL pt AS GpPointF PTR) AS GpStatus
   DECLARE FUNCTION SetCenterPoint (BYVAL pt AS GpPoint PTR) AS GpStatus
   DECLARE FUNCTION GetRectangle (BYVAL rc AS GpRectF PTR) AS GpStatus
   DECLARE FUNCTION GetRectangle (BYVAL rc AS GpRect PTR) AS GpStatus
   DECLARE FUNCTION SetGammaCorrection (BYVAL useGammaCorrection AS BOOL) AS GpStatus
   DECLARE FUNCTION GetGammaCorrection () AS BOOLEAN
   DEClARE FUNCTION GetBlendCount () AS INT_
   DECLARE FUNCTION GetBlend (BYVAL blendFactors AS SINGLE PTR, BYVAL blendPositions AS SINGLE PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION SetBlend (BYVAL blendFactors AS SINGLE PTR, BYVAL blendPositions AS SINGLE PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION GetInterpolationColorCount () AS INT_
   DECLARE FUNCTION SetInterpolationColors (BYVAL presetColors AS ARGB PTR, BYVAL blendPositions AS SINGLE PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION GetInterpolationColors (BYVAL presetColors AS ARGB PTR, BYVAL blendPositions AS SINGLE PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION SetBlendBellShape (BYVAL focus AS SINGLE, BYVAL scale AS SINGLE = 1.0) AS GpStatus
   DECLARE FUNCTION SetBlendTriangularShape (BYVAL focus AS SINGLE, BYVAL scale AS SINGLE = 1.0) AS GpStatus
   DECLARE FUNCTION GetTransform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
   DECLARE FUNCTION SetTransform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
   DECLARE FUNCTION ResetTransform () AS GpStatus
   DECLARE FUNCTION MultiplyTransform (BYVAL pMatrix AS CGpMatrix PTR, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   DECLARE FUNCTION TranslateTransform (BYVAL dx AS SINGLE, BYVAL dy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   DECLARE FUNCTION ScaleTransform (BYVAL sx AS SINGLE, BYVAL sy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   DECLARE FUNCTION RotateTransform (BYVAL angle AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   DECLARE FUNCTION GetFocusScales (BYVAL xScale AS SINGLE PTR, BYVAL yScale AS SINGLE PTR) AS GpStatus
   DECLARE FUNCTION SetFocusScales (BYVAL xScale AS SINGLE, BYVAL yScale AS SINGLE) AS GpStatus
   DECLARE FUNCTION GetWrapMode () AS WrapMode
   DECLARE FUNCTION SetWrapMode (BYVAL nWrapMode AS WrapMode) AS GpStatus
END TYPE

' // CGpHatchBrush
TYPE CGpHatchBrush EXTENDS CGpBrush
   DECLARE CONSTRUCTOR (BYVAL nHatchStyle AS HatchStyle, BYVAL foreColor AS ARGB, BYVAL backColor AS ARGB = &HFF000000)
   DECLARE FUNCTION GetHatchStyle () AS HatchStyle
   DECLARE FUNCTION GetForegroundColor (BYVAL colour AS ARGB PTR) AS GpStatus
   DECLARE FUNCTION GetBackgroundColor (BYVAL colour AS ARGB PTR) AS GpStatus
END TYPE
' ########################################################################################

' ########################################################################################
' CGpPen class
' ########################################################################################
TYPE CGpPen

Public:
   m_Status AS GpStatus
   m_pPen AS GpPen PTR

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL colour AS ARGB, BYVAL nWidth AS SINGLE = 1.0)
   DECLARE CONSTRUCTOR (BYVAL pBrush AS CGpBrush PTR, BYVAL nWidth AS SINGLE = 1.0)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastStatus () AS GpStatus
   DECLARE FUNCTION SetStatus (BYVAL status AS GpStatus) AS GpStatus
   DECLARE FUNCTION Clone (BYVAL pPen AS CGpPen PTR) AS GpStatus
   DECLARE FUNCTION SetWidth (BYVAL nWidth AS SINGLE) AS GpStatus
   DECLARE FUNCTION GetWidth () AS SINGLE
   DECLARE FUNCTION SetLineCap (BYVAL startCap AS LineCap, BYVAL endCap AS LineCap, BYVAL nDashCap AS DashCap) AS GpStatus
   DECLARE FUNCTION SetStartCap (BYVAL startCap AS LineCap) AS GpStatus
   DECLARE FUNCTION SetEndCap (BYVAL endCap AS LineCap) AS GpStatus
   DECLARE FUNCTION SetDashCap (BYVAL nDashCap AS DashCap) AS GpStatus
   DECLARE FUNCTION GetStartCap () AS LineCap
   DECLARE FUNCTION GetEndCap () AS LineCap
   DECLARE FUNCTION GetDashCap () AS DashCap
   DECLARE FUNCTION SetLineJoin (BYVAL nLineJoin AS LineJoin) AS GpStatus
   DECLARE FUNCTION GetLineJoin () AS LineJoin
   DECLARE FUNCTION SetCustomStartCap (BYVAL pCustomLineCap AS CGpCustomLineCap_ PTR) AS GpStatus
   DECLARE FUNCTION GetCustomStartCap (BYVAL pCustomLineCap AS CGpCustomLineCap_ PTR) AS GpStatus
   DECLARE FUNCTION SetCustomEndCap (BYVAL pCustomLineCap AS CGpCustomLineCap_ PTR) AS GpStatus
   DECLARE FUNCTION GetCustomEndCap (BYVAL pCustomLineCap AS CGpCustomLineCap_ PTR) AS GpStatus
   DECLARE FUNCTION SetMiterLimit (BYVAL miterLimit AS SINGLE) AS GpStatus
   DECLARE FUNCTION GetMiterLimit () AS SINGLE
   DECLARE FUNCTION SetAlignment (BYVAL nPenAlignment AS PenAlignment) AS GpStatus
   DECLARE FUNCTION GetAlignment () AS PenAlignment
   DECLARE FUNCTION SetTransform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
   DECLARE FUNCTION GetTransform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
   DECLARE FUNCTION ResetTransform () AS GpStatus
   DECLARE FUNCTION MultiplyTransform (BYVAL pMatrix AS CGpMatrix PTR, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   DECLARE FUNCTION TranslateTransform (BYVAL dx AS SINGLE, BYVAL dy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   DECLARE FUNCTION ScaleTransform (BYVAL sx AS SINGLE, BYVAL sy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   DECLARE FUNCTION RotateTransform (BYVAL angle AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   DECLARE FUNCTION GetPenType () AS PenType
   DECLARE FUNCTION SetColor (BYVAL colour AS ARGB) AS GpStatus
   DECLARE FUNCTION SetBrush (BYVAL pBrush AS CGpBrush PTR) AS GpStatus
   DECLARE FUNCTION GetColor (BYVAL colour AS ARGB PTR) AS GpStatus
   DECLARE FUNCTION GetBrush (BYVAL pBrush AS CGpBrush PTR) AS GpStatus
   DECLARE FUNCTION GetDashStyle () AS DashStyle
   DECLARE FUNCTION SetDashStyle (BYVAL nDashStyle AS DashStyle) AS GpStatus
   DECLARE FUNCTION GetDashOffset () AS SINGLE
   DECLARE FUNCTION SetDashOffset (BYVAL dashOffset AS SINGLE) AS GpStatus
   DECLARE FUNCTION SetDashPattern (BYVAL dashArray AS SINGLE PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION GetDashPatternCount () AS INT_
   DECLARE FUNCTION GetDashPattern (BYVAL dashArray AS SINGLE PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION SetCompoundArray (BYVAL compoundArray AS SINGLE PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION GetCompoundArrayCount () AS GpStatus
   DECLARE FUNCTION GetCompoundArray (BYVAL compoundArray AS SINGLE PTR, BYVAL nCount AS INT_) AS GpStatus

END TYPE
' ########################################################################################

' ########################################################################################
' CGpCustomLineCap class
' ########################################################################################
TYPE CGpCustomLineCap

Public:
   m_Status AS GpStatus
   m_pCustomLineCap AS GpCustomLineCap PTR

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pFillPath AS CGpGraphicsPath_ PTR, BYVAL pStrokePath AS CGpGraphicsPath_ PTR, BYVAL baseCap AS LineCap = LinecapFLat, BYVAL baseInset AS SINGLE = 0.0)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastStatus () AS GpStatus
   DECLARE FUNCTION SetStatus (BYVAL status AS GpStatus) AS GpStatus
   DECLARE FUNCTION SetStrokeCaps (BYVAL startCap AS LineCap, BYVAL endCap AS LineCap) AS GpStatus
   DECLARE FUNCTION GetStrokeCaps (BYVAL startCap AS LineCap PTR, BYVAL endCap AS LineCap PTR) AS GpStatus
   DECLARE FUNCTION SetStrokeCap (BYVAL strokeCap AS LineCap) AS GpStatus
   DECLARE FUNCTION SetStrokeJoin (BYVAL nLineJoin AS LineJoin) AS GpStatus
   DECLARE FUNCTION GetStrokeJoin () AS LineJoin
   DECLARE FUNCTION SetBaseCap (BYVAL baseCap AS LineCap) AS GpStatus
   DECLARE FUNCTION GetBaseCap () AS LineCap
   DECLARE FUNCTION SetBaseInset (BYVAL inset AS SINGLE) AS GpStatus
   DECLARE FUNCTION GetBaseInset () AS SINGLE
   DECLARE FUNCTION SetWidthScale (BYVAL widthScale AS SINGLE) AS GpStatus
   DECLARE FUNCTION GetWidthScale () AS SINGLE
   DECLARE FUNCTION Clone (BYVAL pCustomLineCap AS CGpCustomLineCap PTR) AS GpStatus

END TYPE

TYPE CGpAdjustableArrowCap EXTENDS CGpCustomLineCap

   DECLARE CONSTRUCTOR (BYVAL nHeight AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL isFilled AS BOOL = CTRUE)
   DECLARE FUNCTION SetHeight (BYVAL nHeight AS SINGLE) AS GpStatus
   DECLARE FUNCTION GetHeight () AS SINGLE
   DECLARE FUNCTION SetWidth (BYVAL nWidth AS SINGLE) AS GpStatus
   DECLARE FUNCTION GetWidth () AS SINGLE
   DECLARE FUNCTION SetMiddleInset (BYVAL middleInset AS SINGLE) AS GpStatus
   DECLARE FUNCTION GetMiddleInset () AS SINGLE
   DECLARE FUNCTION SetFillState (BYVAL isFilled AS BOOL) AS GpStatus
   DECLARE FUNCTION IsFilled () AS BOOLEAN

END TYPE
' ########################################################################################

' ########################################################################################
' CGpRegion class
' ########################################################################################
TYPE CGpRegion

Public:
   m_Status AS GpStatus
   m_pRegion AS GpRegion PTR

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL rcf AS GpRectF PTR)
   DECLARE CONSTRUCTOR (BYVAL rc AS GpRect PTR)
   DECLARE CONSTRUCTOR (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE)
   DECLARE CONSTRUCTOR (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_)
   DECLARE CONSTRUCTOR (BYVAL pPath AS CGpGraphicsPath_ PTR)
   DECLARE CONSTRUCTOR (BYVAL regionData AS BYTE PTR, BYVAL nSize AS LONG)
   DECLARE CONSTRUCTOR (BYVAL hRgn AS HRGN)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastStatus () AS GpStatus
   DECLARE FUNCTION SetStatus (BYVAL status AS GpStatus) AS GpStatus
   DECLARE FUNCTION Clone (BYVAL pRegion AS CGpRegion PTR) AS GpStatus
   DECLARE FUNCTION MakeInfinite () AS GpStatus
   DECLARE FUNCTION MakeEmpty () AS GpStatus
   DECLARE FUNCTION Intersect (BYVAL rc AS GpRectF PTR) AS GpStatus
   DECLARE FUNCTION Intersect (BYVAL rc AS GpRect PTR) AS GpStatus
   DECLARE FUNCTION Intersect (BYVAL pPath AS CGpGraphicsPath_ PTR) AS GpStatus
   DECLARE FUNCTION Intersect (BYVAL pRegion AS CGpRegion PTR) AS GpStatus
   DECLARE FUNCTION Union_ (BYVAL rc AS GpRectF PTR) AS GpStatus
   DECLARE FUNCTION Union_ (BYVAL rc AS GpRect PTR) AS GpStatus
   DECLARE FUNCTION Union_ (BYVAL pPath AS CGpGraphicsPath_ PTR) AS GpStatus
   DECLARE FUNCTION Union_ (BYVAL pRegion AS CGpRegion PTR) AS GpStatus
   DECLARE FUNCTION Xor_ (BYVAL rc AS GpRectF PTR) AS GpStatus
   DECLARE FUNCTION Xor_ (BYVAL rc AS GpRect PTR) AS GpStatus
   DECLARE FUNCTION Xor_ (BYVAL gpPath AS CGpGraphicsPath_ PTR) AS GpStatus
   DECLARE FUNCTION Xor_ (BYVAL gpRegion AS CGpRegion PTR) AS GpStatus
   DECLARE FUNCTION Exclude (BYVAL rc AS GpRectF PTR) AS GpStatus
   DECLARE FUNCTION Exclude (BYVAL rc AS GpRect PTR) AS GpStatus
   DECLARE FUNCTION Exclude (BYVAL pPath AS CGpGraphicsPath_ PTR) AS GpStatus
   DECLARE FUNCTION Exclude (BYVAL pRegion AS CGpRegion PTR) AS GpStatus
   DECLARE FUNCTION Complement (BYVAL rc AS GpRectF PTR) AS GpStatus
   DECLARE FUNCTION Complement (BYVAL rc AS GpRect PTR) AS GpStatus
   DECLARE FUNCTION Complement (BYVAL pPath AS CGpGraphicsPath_ PTR) AS GpStatus
   DECLARE FUNCTION Complement (BYVAL pRegion AS CGpRegion pTR) AS GpStatus
   DECLARE FUNCTION Translate (BYVAL dx AS SINGLE, BYVAL dy AS SINGLE) AS GpStatus
   DECLARE FUNCTION Translate (BYVAL dx AS INT_, BYVAL dy AS INT_) AS GpStatus
   DECLARE FUNCTION Transform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
   DECLARE FUNCTION GetBounds (BYVAL rc AS GpRectF PTR, BYVAL pGraphics AS CGpGraphics_ PTR) AS GpStatus
   DECLARE FUNCTION GetBounds (BYVAL rc AS GpRect PTR, BYVAL pGraphics AS CGpGraphics_ pTR) AS GpStatus
   DECLARE FUNCTION GetHRGN (BYVAL pGraphics AS CGpGraphics_ PTR) AS HRGN
   DECLARE FUNCTION IsEmpty (BYVAL pGraphics AS CGpGraphics_ PTR) AS BOOLEAN
   DECLARE FUNCTION IsInfinite (BYVAL pGraphics AS CGpGraphics_ PTR) AS BOOLEAN
   DECLARE FUNCTION Equals (BYVAL pRegion AS CGpRegion PTR, BYVAL pGraphics AS CGpGraphics_ PTR) AS BOOLEAN
   DECLARE FUNCTION GetDataSize () AS UINT
   DECLARE FUNCTION GetData (BYVAL buffer AS BYTE PTR, BYVAL bufferSize AS UINT, BYVAL sizeFilled AS UINT PTR) AS GpStatus
   DECLARE FUNCTION IsVisible (BYVAL pt AS GpPointF PTR, BYVAL pGraphics AS CGpGraphics_ PTR) AS GpStatus
   DECLARE FUNCTION IsVisible (BYVAL pt AS GpPoint PTR, BYVAL pGraphics AS CGpGraphics_ PTR) AS GpStatus
   DECLARE FUNCTION IsVisible (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL pGraphics AS CGpGraphics_ PTR) AS GpStatus
   DECLARE FUNCTION IsVisible (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL gpGraphics AS CGpGraphics_ PTR) AS GpStatus
   DECLARE FUNCTION IsVisible (BYVAL rc AS GpRectF PTR, BYVAL pGraphics AS CGpGraphics_ PTR) AS GpStatus
   DECLARE FUNCTION IsVisible (BYVAL rc AS GpRect PTR, BYVAL pGraphics AS CGpGraphics_ PTR) AS GpStatus
   DECLARE FUNCTION IsVisible (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE, BYVAL pGraphics AS CGpGraphics_ PTR) AS GpStatus
   DECLARE FUNCTION IsVisible (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL pGraphics AS CGpGraphics_ PTR) AS GpStatus
   DECLARE FUNCTION GetRegionScansCount (BYVAL pMatrix AS CGpMatrix PTR) AS UINT
   DECLARE FUNCTION GetRegionScans (BYVAL pMatrix AS CGpMatrix PTR, BYVAL rects AS GpRectF PTR, BYVAL nCount AS INT_ PTR) AS GpStatus

END TYPE
' ########################################################################################

' ########################################################################################
' CGpFont class
' ########################################################################################
TYPE CGpFont

Public:
   m_Status AS GpStatus
   m_pFont AS GpFont PTR

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL hdc AS HDC)
   DECLARE CONSTRUCTOR (BYVAL hdc AS HDC, BYVAL hFont AS HFONT)
#ifdef __FB_64BIT__
   DECLARE CONSTRUCTOR (BYVAL hdc AS HDC, BYVAL plogfont AS LOGFONTW PTR)
   DECLARE CONSTRUCTOR (BYVAL hdc AS HDC, BYVAL plogfont AS LOGFONTA PTR)
#endif
   DECLARE CONSTRUCTOR (BYVAL pFamily AS CGpFontFamily_ PTR, BYVAL emSize AS SINGLE, BYVAL nStyle AS INT_ = 0, BYVAL nUnit AS GpUnit = 0)
   DECLARE CONSTRUCTOR (BYVAL pwszFamilyName AS WSTRING PTR, BYVAL emSize AS SINGLE, BYVAL nStyle AS INT_ = 0, BYVAL nUnit AS GpUnit = 0, BYVAL pFontCollection AS CGpFontCollection_ PTR = NULL)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastStatus () AS GpStatus
   DECLARE FUNCTION SetStatus (BYVAL status AS GpStatus) AS GpStatus
#ifdef __FB_64BIT__
   DECLARE FUNCTION GetLogFontA (BYVAL pGraphics AS CGpGraphics_ PTR, BYVAL pfA AS LOGFONTA PTR) AS GpStatus
   DECLARE FUNCTION GetLogFontW (BYVAL pGraphics AS CGpGraphics_ PTR, BYVAL pfW AS LOGFONTW PTR) AS GpStatus
#endif
   DECLARE FUNCTION Clone (BYVAL pFont AS CGpFont PTR) AS GpStatus
   DECLARE FUNCTION IsAvailable () AS BOOLEAN
   DECLARE FUNCTION GetFamily (BYVAL pFamily AS CGpFontFamily_ PTR) AS GpStatus
   DECLARE FUNCTION GetStyle () AS INT_
   DECLARE FUNCTION GetSize () AS SINGLE
   DECLARE FUNCTION GetUnit () AS GpUnit
   DECLARE FUNCTION GetHeight (BYVAL pGraphics AS CGpGraphics_ PTR) AS SINGLE
   DECLARE FUNCTION GetHeight (BYVAL dpi AS SINGLE) AS SINGLE

END TYPE
' ########################################################################################

' ########################################################################################
' CGpFontFamily class
' ########################################################################################
TYPE CGpFontFamily

Public:
   m_Status AS GpStatus
   m_pFontFamily AS GpFontFamily PTR

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pwszName AS WSTRING PTR, BYVAL pFontCollection AS CGpFontCollection_ PTR = NULL)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastStatus () AS GpStatus
   DECLARE FUNCTION SetStatus (BYVAL status AS GpStatus) AS GpStatus
   DECLARE FUNCTION Clone (BYVAL pFontFamily AS CGpFontFamily PTR) AS GpStatus
   DECLARE FUNCTION GenericSansSerif (BYVAL pFontFamily AS CGpFontFamily PTR) AS GpStatus
   DECLARE FUNCTION GenericSerif (BYVAL pFontFamily AS CGpFontFamily PTR) AS GpStatus
   DECLARE FUNCTION GenericMonospace (BYVAL pFontFamily AS CGpFontFamily PTR) AS GpStatus
   DECLARE FUNCTION GetFamilyName (BYVAL pwszName AS WSTRING PTR, BYVAL language AS LANGID = LANG_NEUTRAL) AS GpStatus
   DECLARE FUNCTION IsAvailable () AS BOOLEAN
   DECLARE FUNCTION IsStyleAvailable (BYVAL nStyle AS INT_) AS BOOLEAN
   DECLARE FUNCTION GetEmHeight (BYVAL nStyle AS INT_) AS UINT16
   DECLARE FUNCTION GetCellAscent (BYVAL nStyle AS INT_) AS UINT16
   DECLARE FUNCTION GetCellDescent (BYVAL nStyle AS INT_) AS UINT16
   DECLARE FUNCTION GetLineSpacing (BYVAL nStyle AS INT_) AS UINT16

END TYPE
' ########################################################################################

' ########################################################################################
' CGpFontCollection class
' ########################################################################################
TYPE CGpFontCollection

Public:
   m_Status AS GpStatus
   m_pFontCollection AS GpFontCollection PTR

Public:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastStatus () AS GpStatus
   DECLARE FUNCTION SetStatus (BYVAL status AS GpStatus) AS GpStatus
   DECLARE FUNCTION GetFamilyCount () AS INT_
   DECLARE FUNCTION GetFamilies (BYVAL numSought AS INT_, BYVAL gpfamilies AS CGpFontFamily PTR, BYVAL numFound AS inT_ PTR) AS GpStatus

END TYPE

TYPE CGpInstalledFontCollection EXTENDS CGpFontCollection

Public:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

END TYPE

TYPE CGpPrivateFontCollection EXTENDS CGpFontCollection

Public:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE FUNCTION AddFontFile (BYVAL pwszFileName AS WSTRING PTR) AS GpStatus
   DECLARE FUNCTION AddMemoryFont (BYVAL pMemory AS ANY PTR, BYVAL length AS INT_) AS GpStatus

END TYPE
' ########################################################################################

' ########################################################################################
' CGpStringFormat class
' ########################################################################################
TYPE CGpStringFormat

Public:
   m_Status AS GpStatus
   m_pStringFormat AS GpStringFormat PTR

Public:
   DECLARE CONSTRUCTOR (BYVAL formatFlags AS INT_ = 0, BYVAL language AS LANGID = LANG_NEUTRAL)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastStatus () AS GpStatus
   DECLARE FUNCTION SetStatus (BYVAL status AS GpStatus) AS GpStatus
   DECLARE FUNCTION Clone (BYVAL pStringFormat AS CGpStringFormat PTR) AS GpStatus
   DECLARE FUNCTION GenericDefault (BYVAL pStringFormat AS CGpStringFormat PTR) AS GpStatus
   DECLARE FUNCTION GenericTypographic (BYVAL pStringFormat AS CGpStringFormat PTR) AS GpStatus
   DECLARE FUNCTION SetFormatFlags (BYVAL flags AS INT_) AS GpStatus
   DECLARE FUNCTION GetFormatFlags () AS INT_
   DECLARE FUNCTION SetAlignment (BYVAL nAlign AS StringAlignment) AS GpStatus
   DECLARE FUNCTION GetAlignment () AS StringAlignment
   DECLARE FUNCTION SetLineAlignment (BYVAL nAlign AS StringAlignment) AS GpStatus
   DECLARE FUNCTION GetLineAlignment () AS StringAlignment
   DECLARE FUNCTION SetHotkeyPrefix (BYVAL nHotkeyPrefix AS HotkeyPrefix) AS GpStatus
   DECLARE FUNCTION GetHotkeyPrefix () AS HotkeyPrefix
   DECLARE FUNCTION SetTabStops (BYVAL firstTabOffset AS SINGLE, BYVAL nCount AS INT_, BYVAL tabStops AS SINGLE PTR) AS GpStatus
   DECLARE FUNCTION GetTabStopCount () AS INT_
   DECLARE FUNCTION GetTabStops (BYVAL nCount AS INT_, BYVAL firstTabOffset AS SINGLE PTR, BYVAL tabStops AS SINGLE PTR) AS GpStatus
   DECLARE FUNCTION SetDigitSubstitution (BYVAL language AS LANGID, BYVAL substitute AS StringDigitSubstitute) AS GpStatus
   DECLARE FUNCTION GetDigitSubstitutionLanguage () AS LANGID
   DECLARE FUNCTION GetDigitSubstitutionMethod () AS StringDigitSubstitute
   DECLARE FUNCTION SetTrimming (BYVAL trimming AS StringTrimming) AS GpStatus
   DECLARE FUNCTION GetTrimming () AS StringTrimming
   DECLARE FUNCTION SetMeasurableCharacterRanges (BYVAL rangeCount AS INT_, BYVAL ranges AS CharacterRange PTR) AS GpStatus
   DECLARE FUNCTION GetMeasurableCharacterRangeCount () AS INT_

END TYPE
' ########################################################################################

' ########################################################################################
' CGpGraphics class
' ########################################################################################
TYPE CGpGraphics

Public:
   m_Status AS GpStatus
   m_pGraphics AS GpGraphics PTR

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL hdc AS HDC)
   DECLARE CONSTRUCTOR (BYVAL hdc AS HDC, BYVAL hDevice AS HANDLE)
   DECLARE CONSTRUCTOR (BYVAL hwnd AS HWND, BYVAL icm AS BOOLEAN = FALSE)
   DECLARE CONSTRUCTOR (BYVAL pImage AS CGpImage PTR)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastStatus () AS GpStatus
   DECLARE FUNCTION SetStatus (BYVAL status AS GpStatus) AS GpStatus
   DECLARE FUNCTION FromHDC (BYVAL hdc AS HDC) AS GpStatus
   DECLARE FUNCTION FromHDC (BYVAL hdc AS HDC, BYVAL hDevice AS HANDLE) AS GpStatus
   DECLARE FUNCTION FromHWND (BYVAL hwnd AS HWND, BYVAL icm AS BOOLEAN = FALSE) AS GpStatus
   DECLARE FUNCTION FromImage (BYVAL pImage AS CGpImage PTR) AS GpStatus
   DECLARE SUB Flush (BYVAL intention AS FlushIntention = FlushIntentionFlush)
   DECLARE FUNCTION GetHDC () AS HDC
   DECLARE SUB ReleaseHDC (BYVAL hdc AS HDC)
   DECLARE FUNCTION SetRenderingOrigin (BYVAL x AS INT_, BYVAL y AS INT_) AS GpStatus
   DECLARE FUNCTION GetRenderingOrigin (BYVAL x AS LONG PTR, BYVAL y AS LONG PTR) AS GpStatus
   DECLARE FUNCTION SetCompositingMode (BYVAL nCompositingMode AS CompositingMode) AS GpStatus
   DECLARE FUNCTION GetCompositingMode () AS CompositingMode
   DECLARE FUNCTION SetCompositingQuality (BYVAL nQuality AS CompositingQuality) AS GpStatus
   DECLARE FUNCTION GetCompositingQuality () AS CompositingQuality
   DECLARE FUNCTION SetTextRenderingHint (BYVAL newMode AS TextRenderingHint) AS GpStatus
   DECLARE FUNCTION GetTextRenderingHint () AS TextRenderingHint
   DECLARE FUNCTION SetTextContrast (BYVAL contrast AS UINT) AS GpStatus
   DECLARE FUNCTION GetTextContrast () AS UINT
   DECLARE FUNCTION GetInterpolationMode () AS InterpolationMode
   DECLARE FUNCTION SetInterpolationMode (BYVAL interpolationMode AS LONG) AS GpStatus
   DECLARE FUNCTION GetSmoothingMode () AS SmoothingMode
   DECLARE FUNCTION SetSmoothingMode (BYVAL smoothingMode AS LONG) AS GpStatus
   DECLARE FUNCTION GetPixelOffsetMode () AS PixelOffsetMode
   DECLARE FUNCTION SetPixelOffsetMode (BYVAL nMode AS PixelOffsetMode) AS GpStatus
   DECLARE FUNCTION SetTransform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
   DECLARE FUNCTION ResetTransform () AS GpStatus
   DECLARE FUNCTION MultiplyTransform (BYVAL pMatrix AS CGpMatrix PTR, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   DECLARE FUNCTION TranslateTransform (BYVAL dx AS SINGLE, BYVAL dy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   DECLARE FUNCTION ScaleTransform (BYVAL sx AS SINGLE, BYVAL sy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   DECLARE FUNCTION RotateTransform (BYVAL angle AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   DECLARE FUNCTION GetTransform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
   DECLARE FUNCTION SetPageUnit (BYVAL unit AS GpUnit) AS GpStatus
   DECLARE FUNCTION SetPageScale (BYVAL nScale AS SINGLE) AS GpStatus
   DECLARE FUNCTION GetPageUnit () AS GpUnit
   DECLARE FUNCTION GetPageScale () AS SINGLE
   DECLARE FUNCTION GetDpiX () AS SINGLE
   DECLARE FUNCTION GetDpiY () AS SINGLE
   DECLARE FUNCTION TransformPoints (BYVAL destSpace AS CoordinateSpace, BYVAL srcSpace AS CoordinateSpace, BYVAL pts AS GpPointF PTR, BYVAL nCount AS LONG) AS GpStatus
   DECLARE FUNCTION TransformPoints (BYVAL destSpace AS CoordinateSpace, BYVAL srcSpace AS CoordinateSpace, BYVAL pts AS GpPoint PTR, BYVAL nCount AS LONG) AS GpStatus
   DECLARE FUNCTION GetNearestColor (BYVAL colour AS ARGB PTR) AS GpStatus
   DECLARE FUNCTION DrawLine (BYVAL pPen AS CGpPen PTR, BYVAL x1 AS SINGLE, BYVAL y1 AS SINGLE, BYVAL x2 AS SINGLE, BYVAL y2 AS SINGLE) AS GpStatus
   DECLARE FUNCTION DrawLine (BYVAL pPen AS CGpPen PTR, BYVAL x1 AS INT_, BYVAL y1 AS INT_, BYVAL x2 AS INT_, BYVAL y2 AS INT_) AS GpStatus
   DECLARE FUNCTION DrawLine (BYVAL pPen AS CGpPen PTR, BYVAL pt1 AS GpPointF PTR, BYVAL pt2 AS GpPointF PTR) AS GpStatus
   DECLARE FUNCTION DrawLine (BYVAL pPen AS CGpPen PTR, BYVAL pt1 AS GpPoint PTR, BYVAL pt2 AS GpPoint PTR) AS GpStatus
   DECLARE FUNCTION DrawLines (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, BYVAL nCount AS LONG) AS GpStatus
   DECLARE FUNCTION DrawLines (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPoint PTR, BYVAL nCount AS LONG) AS GpStatus
   DECLARE FUNCTION DrawArc (BYVAL pPen AS CGpPen PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   DECLARE FUNCTION DrawArc (BYVAL pPen AS CGpPen PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   DECLARE FUNCTION DrawArc (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRectF PTR, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   DECLARE FUNCTION DrawArc (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRect PTR, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   DECLARE FUNCTION DrawBezier (BYVAL pPen AS CGpPen PTR, BYVAL x1 AS SINGLE, BYVAL y1 AS SINGLE, BYVAL x2 AS SINGLE, BYVAL y2 AS SINGLE, BYVAL x3 AS SINGLE, BYVAL y3 AS SINGLE, BYVAL x4 AS SINGLE, BYVAL y4 AS SINGLE) AS GpStatus
   DECLARE FUNCTION DrawBezier (BYVAL pPen AS CGpPen PTR, BYVAL x1 AS INT_, BYVAL y1 AS INT_, BYVAL x2 AS INT_, BYVAL y2 AS INT_, BYVAL x3 AS INT_, BYVAL y3 AS INT_, BYVAL x4 AS INT_, BYVAL y4 AS INT_) AS GpStatus
   DECLARE FUNCTION DrawBezier (BYVAL pPen AS CGpPen PTR, BYVAL pt1 AS GpPointF PTR, BYVAL pt2 AS GpPointF PTR, BYVAL pt3 AS GpPointF PTR, BYVAL pt4 AS GpPointF PTR) AS GpStatus
   DECLARE FUNCTION DrawBezier (BYVAL pPen AS CGpPen PTR, BYVAL pt1 AS GpPoint PTR, BYVAL pt2 AS GpPoint PTR, BYVAL pt3 AS GpPoint PTR, BYVAL pt4 AS GpPoint PTR) AS GpStatus
   DECLARE FUNCTION DrawBeziers (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION DrawBeziers (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION DrawRectangle (BYVAL pPen AS CGpPen PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS GpStatus
   DECLARE FUNCTION DrawRectangle (BYVAL pPen AS CGpPen PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS GpStatus
   DECLARE FUNCTION DrawRectangle (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRectF PTR) AS GpStatus
   DECLARE FUNCTION DrawRectangle (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRect PTR) AS GpStatus
   DECLARE FUNCTION DrawRectangles (BYVAL pPen AS CGpPen PTR, BYVAL rects AS GpRectF PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION DrawRectangles (BYVAL pPen AS CGpPen PTR, BYVAL rects AS GpRect PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION DrawEllipse (BYVAL pPen AS CGpPen PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS GpStatus
   DECLARE FUNCTION DrawEllipse (BYVAL pPen AS CGpPen PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS GpStatus
   DECLARE FUNCTION DrawEllipse (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRectF PTR) AS GpStatus
   DECLARE FUNCTION DrawEllipse (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRect PTR) AS GpStatus
   DECLARE FUNCTION DrawPie (BYVAL pPen AS CGpPen PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   DECLARE FUNCTION DrawPie (BYVAL pPen AS CGpPen PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   DECLARE FUNCTION DrawPie (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRectF PTR, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   DECLARE FUNCTION DrawPie (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRect PTR, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   DECLARE FUNCTION DrawPolygon (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, BYVAL nCount AS LONG) AS GpStatus
   DECLARE FUNCTION DrawPolygon (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPoint PTR, BYVAL nCount AS LONG) AS GpStatus
   DECLARE FUNCTION DrawPath (BYVAL pPen AS CGpPen PTR, BYVAL pPath AS CGpGraphicsPath_ PTR) AS GpStatus
   DECLARE FUNCTION DrawCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION DrawCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION DrawCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_, BYVAL tension AS SINGLE) AS GpStatus
   DECLARE FUNCTION DrawCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_, BYVAL tension AS SINGLE) AS GpStatus
   DECLARE FUNCTION DrawCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_, BYVAL offset AS INT_, BYVAL numberOfSegments AS INT_, BYVAL tension AS SINGLE = 0.5) AS GpStatus
   DECLARE FUNCTION DrawCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_, BYVAL offset AS INT_, BYVAL numberOfSegments AS INT_, BYVAL tension AS SINGLE = 0.5) AS GpStatus
   DECLARE FUNCTION DrawClosedCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION DrawClosedCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION DrawClosedCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_, BYVAL tension AS SINGLE) AS GpStatus
   DECLARE FUNCTION DrawClosedCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_, BYVAL tension AS SINGLE) AS GpStatus
   DECLARE FUNCTION Clear (BYVAL colour AS ARGB) AS GpStatus
   DECLARE FUNCTION FillRectangle (BYVAL pBrush AS CGpBrush PTR, BYVAL rc AS GpRectF PTR) AS GpStatus
   DECLARE FUNCTION FillRectangle (BYVAL pBrush AS CGpBrush PTR, BYVAL rc AS GpRect PTR) AS GpStatus
   DECLARE FUNCTION FillRectangle (BYVAL pBrush AS CGpBrush PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS GpStatus
   DECLARE FUNCTION FillRectangle (BYVAL pBrush AS CGpBrush PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS GpStatus
   DECLARE FUNCTION FillRectangles (BYVAL pBrush AS CGpBrush PTR, BYVAL rects AS GpRectF PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION FillRectangles (BYVAL pBrush AS CGpBrush PTR, BYVAL rects AS GpRect PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION FillPolygon (BYVAL pBrush AS CGpBrush PTR, BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION FillPolygon (BYVAL pBrush AS CGpBrush PTR, BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION FillPolygon (BYVAL pBrush AS CGpBrush PTR, BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_, BYVAL nFillMode AS FillMode) AS GpStatus
   DECLARE FUNCTION FillPolygon (BYVAL pBrush AS CGpBrush PTR, BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_, BYVAL nFillMode AS FillMode) AS GpStatus
   DECLARE FUNCTION FillEllipse (BYVAL pBrush AS CGpBrush PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS GpStatus
   DECLARE FUNCTION FillEllipse (BYVAL pBrush AS CGpBrush PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS GpStatus
   DECLARE FUNCTION FillEllipse (BYVAL pBrush AS CGpBrush PTR, BYVAL rc AS GpRectF PTR) AS GpStatus
   DECLARE FUNCTION FillEllipse (BYVAL pBrush AS CGpBrush PTR, BYVAL rc AS GpRect PTR) AS GpStatus
   DECLARE FUNCTION FillPie (BYVAL pBrush AS CGpBrush PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   DECLARE FUNCTION FillPie (BYVAL pBrush AS CGpBrush PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   DECLARE FUNCTION FillPie (BYVAL pBrush AS CGpBrush PTR, BYVAL rc AS GpRectF PTR, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   DECLARE FUNCTION FillPie (BYVAL pBrush AS CGpBrush PTR, BYVAL rc AS GpRect PTR, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   DECLARE FUNCTION FillPath (BYVAL pBrush AS CGpBrush PTR, BYVAL pPath AS CGpGraphicsPath_ PTR) AS GpStatus
   DECLARE FUNCTION FillClosedCurve (BYVAL pBrush AS CGpBrush PTR, BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION FillClosedCurve (BYVAL pBrush AS CGpBrush PTR, BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION FillClosedCurve (BYVAL pBrush AS CGpBrush PTR, BYVAL pts AS GpPointF PTR, BYVAL nCount AS LONG, BYVAL nFillMode AS FillMode, BYVAL tension AS SINGLE = 0.5) AS GpStatus
   DECLARE FUNCTION FillClosedCurve (BYVAL pBrush AS CGpBrush PTR, BYVAL pts AS GpPoint PTR, BYVAL nCount AS LONG, BYVAL nFillMode AS FillMode, BYVAL tension AS SINGLE = 0.5) AS GpStatus
   DECLARE FUNCTION FillRegion (BYVAL pBrush AS CGpBrush PTR, BYVAL pRegion AS CGpRegion PTR) AS GpStatus
   DECLARE FUNCTION DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFont AS CGpFont PTR, BYVAL rc AS GpRectF PTR, BYVAL pBrush AS CGpBrush PTR) AS GpStatus
   DECLARE FUNCTION DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFont AS CGpFont PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE, BYVAL pBrush AS CGpBrush PTR) AS GpStatus
   DECLARE FUNCTION DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFont AS CGpFont PTR, BYVAL layoutRect AS GpRectF PTR, BYVAL pStringFormat AS CGpStringFormat PTR, BYVAL pBrush AS CGpBrush PTR) AS GpStatus
   DECLARE FUNCTION DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFont AS CGpFont PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE, BYVAL pStringFormat AS CGpStringFormat PTR, BYVAL pBrush AS CGpBrush PTR) AS GpStatus
   DECLARE FUNCTION DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFont AS CGpFont PTR, BYVAL origin as GpPointF PTR, BYVAL pBrush AS CGpBrush PTR) AS GpStatus
   DECLARE FUNCTION DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFont AS CGpFont PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL pBrush AS CGpBrush PTR) AS GpStatus
   DECLARE FUNCTION DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFont AS CGpFont PTR, BYVAL origin AS GpPointF PTR, BYVAL pStringFormat AS CGpStringFormat PTR, BYVAL pBrush AS CGpBrush PTR) AS GpStatus
   DECLARE FUNCTION DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFont AS CGpFont PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL pStringFormat AS CGpStringFormat PTR, BYVAL pBrush AS CGpBrush PTR) AS GpStatus
   DECLARE FUNCTION MeasureString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFont AS CGpFont PTR, BYVAL layoutRect AS GpRectF PTR, BYVAL pStringFormat AS CGpStringFormat PTR, BYVAL boundingBox AS GpRectF PTR, BYREF codepointsFitted AS INT_ = 0, BYREF linesFilled AS INT_ = 0) AS GpStatus
   DECLARE FUNCTION MeasureString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFont AS CGpFont PTR, BYVAL layoutRectSize AS GpSizeF PTR, BYVAL pStringFormat AS CGpStringFormat PTR, BYVAL size AS GpSizeF PTR, BYREF codepointsFitted AS INT_ = 0, BYREF linesFilled AS INT_ = 0) AS GpStatus
   DECLARE FUNCTION MeasureString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFont AS CGpFont PTR, BYVAL origin AS GpPointF PTR, BYVAL pStringFormat AS CGpStringFormat PTR, BYVAL boundingBox AS GpRectF PTR) AS GpStatus
   DECLARE FUNCTION MeasureString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFont AS CGpFont PTR, BYVAL layoutRect AS GpRectF PTR, BYVAL pStringFormat AS CGpStringFormat PTR, BYVAL boundingBox AS GpRectF PTR) AS GpStatus
   DECLARE FUNCTION MeasureString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFont AS CGpFont PTR, BYVAL layoutRect AS GpRectF PTR, BYVAL boundingBox AS GpRectF PTR) AS GpStatus
   DECLARE FUNCTION MeasureString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFont AS CGpFont PTR, BYVAL origin AS GpPointF PTR, BYVAL boundingBox AS GpRectF PTR) AS GpStatus
   DECLARE FUNCTION MeasureCharacterRanges (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFont AS CGpFont PTR, _
      BYVAL layoutRect AS GpRectF PTR, BYVAL pStringFormat AS CGpStringFormat PTR, BYVAL regionCount AS INT_, BYVAL regions AS CGpRegion PTR) AS GpStatus
   DECLARE FUNCTION DrawDriverString (BYVAL rgText AS UINT16 PTR, BYVAL length AS LONG, BYVAL pFont AS CGpFont PTR, BYVAL pBrush AS CGpBrush PTR, BYVAL positions AS ANY PTR, BYVAL flags AS LONG, BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
   DECLARE FUNCTION MeasureDriverString (BYVAL pText AS UINT16 PTR, BYVAL length AS LONG, BYVAL pFont AS CGpFont PTR, _
      BYVAL positions AS ANY PTR, BYVAL flags AS LONG, BYVAL pMatrix AS CGpMatrix PTR, BYVAL boundingBox AS GpRectF PTR) AS GpStatus
   DECLARE FUNCTION DrawCachedBitmap (BYVAL gpCachedBitmap AS CGpCachedBitmap PTR, BYVAL x AS LONG, BYVAL y AS LONG) AS GpStatus
   DECLARE FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL pt AS GpPointF PTR) AS GpStatus
   DECLARE FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE) AS GpStatus
   DECLARE FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS GpStatus
   DECLARE FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL rcf AS GpRectF PTR) AS GpStatus
   DECLARE FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL pt AS GpPoint PTR) AS GpStatus
   DECLARE FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL x AS INT_, BYVAL y AS INT_) AS GpStatus
   DECLARE FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS GpStatus
   DECLARE FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL rc AS GpRect PTR) AS GpStatus
   DECLARE FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destPoints AS GpPointF PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destPoints AS GpPoint PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL srcx AS SINGLE, BYVAL srcy AS SINGLE, BYVAL srcwidth AS SINGLE, BYVAL srcheight AS SINGLE, BYVAL srcUnit AS gpUnit) AS GpStatus
   DECLARE FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL srcx AS INT_, BYVAL srcy AS INT_, BYVAL srcwidth AS INT_, BYVAL srcheight AS INT_, BYVAL srcUnit AS gpUnit) AS GpStatus
   DECLARE FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destRect AS GpRectF PTR, _
      BYVAL srcx AS SINGLE, BYVAL srcy AS SINGLE, BYVAL srcwidth AS SINGLE, BYVAL srcheight AS SINGLE, BYVAL srcUnit AS GpUnit, _
      BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL, BYVAL pcallback AS DrawImageAbort = NULL, BYVAL pcallbackData AS ANY PTR = NULL) AS GpStatus
   DECLARE FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destRect AS GpRect PTR, _
      BYVAL srcx AS INT_, BYVAL srcy AS INT_, BYVAL srcwidth AS INT_, BYVAL srcheight AS INT_, BYVAL srcUnit AS GpUnit, _
      BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL, BYVAL pcallback AS DrawImageAbort = NULL, BYVAL pcallbackData AS ANY PTR = NULL) AS GpStatus
   DECLARE FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destPoints AS GpPointF PTR, BYVAL nCount AS INT_, _
      BYVAL srcx AS SINGLE, BYVAL srcy AS SINGLE, BYVAL srcwidth AS SINGLE, BYVAL srcheight AS SINGLE, BYVAL srcUnit AS GpUnit, _
      BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL, BYVAL pcallback AS DrawImageAbort = NULL, BYVAL pcallbackData AS ANY PTR = NULL) AS GpStatus
   DECLARE FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destPoints AS GpPoint PTR, BYVAL nCount AS INT_, _
      BYVAL srcx AS INT_, BYVAL srcy AS INT_, BYVAL srcwidth AS INT_, BYVAL srcheight AS INT_, BYVAL srcUnit AS GpUnit, _
      BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL, BYVAL pcallback AS DrawImageAbort = NULL, BYVAL pcallbackData AS ANY PTR = NULL) AS GpStatus
   DECLARE FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destx AS SINGLE, _
      BYVAL desty AS SINGLE, BYVAL destWidth AS SINGLE, BYVAL destHeight AS SINGLE, _
      BYVAL srcx AS SINGLE, BYVAL srcy AS SINGLE, BYVAL srcwidth AS SINGLE, BYVAL srcheight AS SINGLE, BYVAL srcUnit AS GpUnit, _
      BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL, BYVAL pcallback AS DrawImageAbort = NULL, BYVAL pcallbackData AS ANY PTR = NULL) AS GpStatus
   DECLARE FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destx AS INT_, BYVAL desty AS INT_, BYVAL destWidth AS INT_, BYVAL destHeight AS INT_, _
      BYVAL srcx AS INT_, BYVAL srcy AS INT_, BYVAL srcwidth AS INT_, BYVAL srcheight AS INT_, BYVAL srcUnit AS GpUnit, _
      BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL, BYVAL pcallback AS DrawImageAbort = NULL, BYVAL pcallbackData AS ANY PTR = NULL) AS GpStatus
   DECLARE FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destRect AS GpRectF PTR, _
      BYVAL sourceRect AS GpRectF PTR, BYVAL srcUnit AS GpUnit, BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL) AS GpStatus
   DECLARE FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destRect AS GpRect PTR, _
      BYVAL sourceRect AS GpRect PTR, BYVAL srcUnit AS GpUnit, BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL) AS GpStatus
   DECLARE FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destx AS SINGLE, BYVAL desty AS SINGLE, BYVAL destWidth AS SINGLE, BYVAL destHeight AS SINGLE, _
      BYVAL srcx AS SINGLE, BYVAL srcy AS SINGLE, BYVAL srcWidth AS SINGLE, BYVAL srcHeight AS SINGLE, BYVAL srcUnit AS GpUnit, BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL) AS GpStatus
   DECLARE FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destx AS INT_, BYVAL desty AS INT_, BYVAL destWidth AS INT_, BYVAL destHeight AS INT_, _
      BYVAL srcx AS INT_, BYVAL srcy AS INT_, BYVAL srcWidth AS INT_, BYVAL srcHeight AS INT_, BYVAL srcUnit AS GpUnit, BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL) AS GpStatus
   DECLARE FUNCTION SetClip (BYVAL pGraphics AS CGpGraphics PTR, BYVAL nCombineMode AS CombineMode = CombineModeReplace) AS GpStatus
   DECLARE FUNCTION SetClip (BYVAL rc AS GpRectF PTR, BYVAL nCombineMode AS CombineMode = CombineModeReplace) AS GpStatus
   DECLARE FUNCTION SetClip (BYVAL rc AS GpRect PTR, BYVAL nCombineMode AS CombineMode = CombineModeReplace) AS GpStatus
   DECLARE FUNCTION SetClip (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE, BYVAL nCombineMode AS CombineMode = CombineModeReplace) AS GpStatus
   DECLARE FUNCTION SetClip (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL nCombineMode AS CombineMode = CombineModeReplace) AS GpStatus
   DECLARE FUNCTION SetClip (BYVAL pPath AS CGpGraphicsPath_ PTR, BYVAL nCombineMode AS CombineMode = CombineModeReplace) AS GpStatus
   DECLARE FUNCTION SetClip (BYVAL pRegion AS CGpRegion PTR, BYVAL nCombineMode AS CombineMode = CombineModeReplace) AS GpStatus
   DECLARE FUNCTION SetClip (BYVAL hRgn AS HRGN, BYVAL nCombineMode AS CombineMode = CombineModeReplace) AS GpStatus
   DECLARE FUNCTION IntersectClip (BYVAL rcf AS GpRectF PTR) AS GpStatus
   DECLARE FUNCTION IntersectClip (BYVAL rc AS GpRect PTR) AS GpStatus
   DECLARE FUNCTION IntersectClip (BYVAL pRegion AS CGpRegion PTR) AS GpStatus
   DECLARE FUNCTION IntersectClip (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS GpStatus
   DECLARE FUNCTION IntersectClip (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS GpStatus
   DECLARE FUNCTION ExcludeClip (BYVAL rcf AS GpRectF PTR) AS GpStatus
   DECLARE FUNCTION ExcludeClip (BYVAL rc AS GpRect PTR) AS GpStatus
   DECLARE FUNCTION ExcludeClip (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS GpStatus
   DECLARE FUNCTION ExcludeClip (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS GpStatus
   DECLARE FUNCTION ExcludeClip (BYVAL pRegion AS CGpRegion PTR) AS GpStatus
   DECLARE FUNCTION ResetClip () AS GpStatus
   DECLARE FUNCTION TranslateClip (BYVAL dx AS SINGLE, BYVAL dy AS SINGLE) AS GpStatus
   DECLARE FUNCTION GetClip (BYVAL pRegion AS CGpRegion PTR) AS GpStatus
   DECLARE FUNCTION GetClipBounds (BYVAL rc AS GpRectF PTR) AS GpStatus
   DECLARE FUNCTION GetClipBounds (BYVAL rc AS GpRect PTR) AS GpStatus
   DECLARE FUNCTION IsClipEmpty () AS BOOLEAN
   DECLARE FUNCTION GetVisibleClipBounds (BYVAL rc AS GpRectF PTR) AS GpStatus
   DECLARE FUNCTION GetVisibleClipBounds (BYVAL rc AS GpRect PTR) AS GpStatus
   DECLARE FUNCTION IsVisibleClipEmpty () AS BOOLEAN
   DECLARE FUNCTION IsVisible (BYVAL pt AS GpPointF PTR) AS BOOLEAN
   DECLARE FUNCTION IsVisible (BYVAL pt AS GpPoint PTR) AS BOOLEAN
   DECLARE FUNCTION IsVisible (BYVAL x AS SINGLE, BYVAL y AS SINGLE) AS BOOLEAN
   DECLARE FUNCTION IsVisible (BYVAL x AS INT_, BYVAL y AS INT_) AS BOOLEAN
   DECLARE FUNCTION IsVisible (BYVAL rcf AS GpRectF PTR) AS BOOLEAN
   DECLARE FUNCTION IsVisible (BYVAL rc AS GpRect PTR) AS BOOLEAN
   DECLARE FUNCTION IsVisible (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS BOOLEAN
   DECLARE FUNCTION IsVisible (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS BOOLEAN
   DECLARE FUNCTION Save () AS GraphicsState
   DECLARE FUNCTION Restore (BYVAL gstate AS GraphicsState) AS GpStatus
   DECLARE FUNCTION BeginContainer () AS GraphicsContainer
   DECLARE FUNCTION BeginContainer (BYVAL destrect AS GpRectF PTR, BYVAL srcrect AS GpRectF PTR, BYVAL nUnit AS GpUnit) AS GraphicsContainer
   DECLARE FUNCTION BeginContainer (BYVAL destrect AS GpRect PTR, BYVAL srcrect AS GpRect PTR, BYVAL nUnit AS GpUnit) AS GraphicsContainer
   DECLARE FUNCTION EndContainer (BYVAL state AS GraphicsContainer) AS GpStatus
   DECLARE FUNCTION AddMetafileComment (BYVAL pdata AS BYTE PTR, BYVAL sizeData AS UINT) AS GpStatus
   DECLARE FUNCTION GetHalftonePalette () AS HPALETTE

END TYPE
' ########################################################################################

' ########################################################################################
' CGpGraphicsPath class
' ########################################################################################
TYPE CGpGraphicsPath

Public:
   m_Status AS GpStatus
   m_pPath AS GpPath PTR

Public:
   DECLARE CONSTRUCTOR (BYVAL nFillMode AS FillMode = FillModeAlternate)
   DECLARE CONSTRUCTOR (BYVAL pts AS GpPointF PTR, BYVAL types AS BYTE PTR, BYVAL nCount AS LONG, BYVAL nFillMode AS FillMode = FillModeAlternate)
   DECLARE CONSTRUCTOR (BYVAL pts AS GpPoint PTR, BYVAL types AS BYTE PTR, BYVAL nCount AS LONG, BYVAL nFillMode AS FillMode = FillModeAlternate)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastStatus () AS GpStatus
   DECLARE FUNCTION SetStatus (BYVAL status AS GpStatus) AS GpStatus
   DECLARE FUNCTION Clone (BYVAL pCloneGraphicsPath AS CGpGraphicsPath PTR) AS GpStatus
   DECLARE FUNCTION Reset () AS GpStatus
   DECLARE FUNCTION GetFillMode () AS FillMode
   DECLARE FUNCTION SetFillMode (BYVAL nFillmode AS FillMode) AS GpStatus
   DECLARE FUNCTION GetPathData (BYVAL pPathData AS ANY PTR) AS GpStatus
   DECLARE FUNCTION StartFigure () AS GpStatus
   DECLARE FUNCTION CloseFigure () AS GpStatus
   DECLARE FUNCTION CloseAllFigures () AS GpStatus
   DECLARE FUNCTION SetMarker () AS GpStatus
   DECLARE FUNCTION ClearMarkers () AS GpStatus
   DECLARE FUNCTION Reverse () AS GpStatus
   DECLARE FUNCTION GetLastPoint (BYVAL lastPoint AS GpPointF PTR) AS GpStatus
   DECLARE FUNCTION AddLine (BYVAL x1 AS SINGLE, BYVAL y1 AS SINGLE, BYVAL x2 AS SINGLE, BYVAL y2 AS SINGLE) AS GpStatus
   DECLARE FUNCTION AddLine (BYVAL x1 AS INT_, BYVAL y1 AS INT_, BYVAL x2 AS INT_, BYVAL y2 AS INT_) AS GpStatus
   DECLARE FUNCTION AddLine (BYVAL pt1 AS GpPointF PTR, BYVAL pt2 AS GpPointF PTR) AS GpStatus
   DECLARE FUNCTION AddLine (BYVAL pt1 AS GpPoint PTR, BYVAL pt2 AS GpPoint PTR) AS GpStatus
   DECLARE FUNCTION AddLines (BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION AddLines (BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION AddArc (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL width_ AS SINGLE, BYVAL height_ AS SINGLE, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   DECLARE FUNCTION AddArc (BYVAL rcf AS GpRectF PTR, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   DECLARE FUNCTION AddArc (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL width_ AS INT_, BYVAL height_ AS INT_, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   DECLARE FUNCTION AddArc (BYVAL rc AS GpRect PTR, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   DECLARE FUNCTION AddBezier (BYVAL x1 AS SINGLE, BYVAL y1 AS SINGLE, BYVAL x2 AS SINGLE, BYVAL y2 AS SINGLE, BYVAL x3 AS SINGLE, BYVAL y3 AS SINGLE, BYVAL x4 AS SINGLE, BYVAL y4 AS SINGLE) AS GpStatus
   DECLARE FUNCTION AddBezier (BYVAL x1 AS INT_, BYVAL y1 AS INT_, BYVAL x2 AS INT_, BYVAL y2 AS INT_, BYVAL x3 AS INT_, BYVAL y3 AS INT_, BYVAL x4 AS INT_, BYVAL y4 AS INT_) AS GpStatus
   DECLARE FUNCTION AddBezier (BYVAL pt1 AS GpPointF PTR, BYVAL pt2 AS GpPointF PTR, BYVAL pt3 AS GpPointF PTR, BYVAL pt4 AS GpPointF PTR) AS GpStatus
   DECLARE FUNCTION AddBezier (BYVAL pt1 AS GpPoint PTR, BYVAL pt2 AS GpPoint PTR, BYVAL pt3 AS GpPoint PTR, BYVAL pt4 AS GpPoint PTR) AS GpStatus
   DECLARE FUNCTION AddBeziers (BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION AddBeziers (BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION AddCurve (BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION AddCurve (BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION AddCurve (BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_, BYVAL tension AS SINGLE) AS GpStatus
   DECLARE FUNCTION AddCurve (BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_, BYVAL tension AS SINGLE) AS GpStatus
   DECLARE FUNCTION AddCurve (BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_, BYVAL offset AS INT_, BYVAL numberOfSegments AS INT_, BYVAL tension AS SINGLE) AS GpStatus
   DECLARE FUNCTION AddCurve (BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_, BYVAL offset AS INT_, BYVAL numberOfSegments AS INT_, BYVAL tension AS SINGLE) AS GpStatus
   DECLARE FUNCTION AddClosedCurve (BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION AddClosedCurve (BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_, BYVAL tension AS SINGLE) AS GpStatus
   DECLARE FUNCTION AddClosedCurve (BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION AddClosedCurve (BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_, BYVAL tension AS SINGLE) AS GpStatus
   DECLARE FUNCTION AddRectangle (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS GpStatus
   DECLARE FUNCTION AddRectangle (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS GpStatus
   DECLARE FUNCTION AddRectangle (BYVAL rcr AS GpRectF PTR) AS GpStatus
   DECLARE FUNCTION AddRectangle (BYVAL rcr AS GpRect PTR) AS GpStatus
   DECLARE FUNCTION AddRectangles (BYVAL rects AS GpRectF PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION AddRectangles (BYVAL rects AS GpRect PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION AddEllipse (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS GpStatus
   DECLARE FUNCTION AddEllipse (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS GpStatus
   DECLARE FUNCTION AddEllipse (BYVAL rcf AS GpRectF PTR) AS GpStatus
   DECLARE FUNCTION AddEllipse (BYVAL rc AS GpRect PTR) AS GpStatus
   DECLARE FUNCTION AddPie (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   DECLARE FUNCTION AddPie (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   DECLARE FUNCTION AddPie (BYVAL rcf AS GpRectF PTR, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   DECLARE FUNCTION AddPie (BYVAL rcf AS GpRect PTR, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
   DECLARE FUNCTION AddPolygon (BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION AddPolygon (BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION AddPath (BYVAL pAddingPath AS CGpGraphicsPath PTR, BYVAL bConnect AS BOOL) AS GpStatus
   DECLARE FUNCTION AddString (BYVAL pString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFamily AS CGpFontFamily_ PTR, BYVAL nStyle AS INT_, BYVAL emSize AS SINGLE, BYVAL layoutRect AS GpRectF PTR, BYVAL pFormat AS CGpStringFormat_ PTR) AS GpStatus
   DECLARE FUNCTION AddString (BYVAL pString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFamily AS CGpFontFamily_ PTR, BYVAL nStyle AS INT_, BYVAL emSize AS SINGLE, BYVAL layoutRect AS GpRect PTR, BYVAL pFormat AS CGpStringFormat_ PTR) AS GpStatus
   DECLARE FUNCTION AddString (BYVAL pString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFamily AS CGpFontFamily_ PTR, BYVAL nStyle AS INT_, BYVAL emSize AS SINGLE, BYVAL origin AS GpPointF PTR, BYVAL pFormat AS CGpStringFormat_ PTR) AS GpStatus
   DECLARE FUNCTION AddString (BYVAL pString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFamily AS CGpFontFamily_ PTR, BYVAL nStyle AS INT_, BYVAL emSize AS SINGLE, BYVAL origin AS GpPoint PTR, BYVAL pFormat AS CGpStringFormat_ PTR) AS GpStatus
   DECLARE FUNCTION TransForm (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
   DECLARE FUNCTION GetBounds (BYVAL bounds AS GpRectF PTR, BYVAL pMatrix AS CGpMatrix PTR, BYVAL pPen AS CGpPen PTR) AS GpStatus
   DECLARE FUNCTION GetBounds (BYVAL bounds AS GpRect PTR, BYVAL pMatrix AS CGpMatrix PTR, BYVAL pPen AS CGpPen PTR) AS GpStatus
   DECLARE FUNCTION Flatten (BYVAL pMatrix AS CGpMatrix PTR = NULL, BYVAL flatness AS SINGLE = FlatnessDefault) AS GpStatus
   DECLARE FUNCTION Widen (BYVAL pPen AS CGpPen PTR, BYVAL pMatrix AS CGpMatrix PTR = NULL, BYVAL flatness AS SINGLE = FlatnessDefault) AS GpStatus
   DECLARE FUNCTION Outline (BYVAL pMatrix AS CGpMatrix PTR = NULL, BYVAL flatness AS SINGLE = FlatnessDefault) AS GpStatus
   DECLARE FUNCTION Warp (BYVAL destPoints AS GpPointF PTR, BYVAL nCount AS INT_, BYVAL srcRect AS GpRectF PTR, BYVAL pMatrix AS CGpMatrix PTR = NULL, BYVAL nWarpMode AS WarpMode = WarpModePerspective, BYVAL Flatness AS SINGLE = FlatnessDefault) AS GpStatus
   DECLARE FUNCTION GetPointCount () AS INT_
   DECLARE FUNCTION GetPathTypes (BYVAL types AS BYTE PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION GetPathPoints (BYREF pts AS GpPointF PTR, BYVAL nCount AS INT_) AS GpStatus
   DECLARE FUNCTION IsVisible (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL pGraphics AS CGpGraphics_ PTR = NULL) AS BOOLEAN
   DECLARE FUNCTION IsVisible (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL pGraphics AS CGpGraphics_ PTR = NULL) AS BOOLEAN
   DECLARE FUNCTION IsVisible (BYVAL pt AS GpPointF PTR, BYVAL pGraphics AS CGpGraphics_ PTR = NULL) AS BOOLEAN
   DECLARE FUNCTION IsVisible (BYVAL pt AS GpPoint PTR, BYVAL pGraphics AS CGpGraphics_ PTR = NULL) AS BOOLEAN
   DECLARE FUNCTION IsOutlineVisible (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL pPen AS CGpPen PTR, BYVAL pGraphics AS CGpGraphics_ PTR = NULL) AS BOOLEAN
   DECLARE FUNCTION IsOutlineVisible (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL pPen AS CGpPen PTR, BYVAL pGraphics AS CGpGraphics_ PTR = NULL) AS BOOLEAN
   DECLARE FUNCTION IsOutlineVisible (BYVAL pt AS GpPointF PTR, BYVAL pPen AS CGpPen PTR, BYVAL pGraphics AS CGpGraphics_ PTR = NULL) AS BOOLEAN
   DECLARE FUNCTION IsOutlineVisible (BYVAL pt AS GpPoint PTR, BYVAL pPen AS CGpPen PTR, BYVAL pGraphics AS CGpGraphics_ PTR = NULL) AS BOOLEAN

END TYPE
' ########################################################################################

' ########################################################################################
' CGpGraphicsPathIterator class
' ########################################################################################
TYPE CGpGraphicsPathIterator

Public:
   m_Status AS GpStatus
   m_pPathIretator AS GpPathIterator PTR

Public:
   DECLARE CONSTRUCTOR (BYVAL pPath AS CGpGraphicsPath PTR)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastStatus () AS GpStatus
   DECLARE FUNCTION SetStatus (BYVAL status AS GpStatus) AS GpStatus
   DECLARE FUNCTION NextSubpath (BYVAL startIndex AS INT_ PTR, BYVAL endIndex AS INT_ PTR, BYVAL isClosed AS BOOL PTR) AS INT_
   DECLARE FUNCTION NextSubpathPath (BYVAL pPath AS CGpGraphicsPath PTR, BYVAL isClosed AS BOOL PTR) AS INT_
   DECLARE FUNCTION NextPathType (BYVAL pathType AS BYTE PTR, BYVAL startIndex AS INT_ PTR, BYVAL endIndex AS INT_ PTR) AS INT_
   DECLARE FUNCTION NextMarker (BYVAL startIndex AS INT_ PTR, BYVAL endIndex AS INT_ PTR) AS INT_
   DECLARE FUNCTION NextMarker (BYVAL pPath AS CGpGraphicsPath PTR) AS INT_
   DECLARE FUNCTION GetCount () AS INT_
   DECLARE FUNCTION GetSubpathCount () AS INT_
   DECLARE FUNCTION HasCurve () AS BOOLEAN
   DECLARE SUB Rewind
   DECLARE FUNCTION Enumerate (BYVAL pts AS GpPointF PTR, BYVAL types AS BYTE PTR, BYVAL nCount AS INT_) AS INT_
   DECLARE FUNCTION CopyData (BYVAL pts AS GpPointF PTR, BYVAL types AS BYTE PTR, BYVAL startIndex AS INT_, BYVAL endIndex AS INT_) AS INT_

END TYPE
' ########################################################################################

#include once "Afx/CGdiPlus/CGpBitmap.inc"
#include once "Afx/CGdiPlus/CGpBrush.inc"
#include once "Afx/CGdiPlus/CGpFont.inc"
#include once "Afx/CGdiPlus/CGpGraphics.inc"
#include once "Afx/CGdiPlus/CGpImageAttributes.inc"
#include once "Afx/CGdiPlus/CGpLineCaps.inc"
#include once "Afx/CGdiPlus/CGpMatrix.inc"
#include once "Afx/CGdiPlus/CGpPath.inc"
#include once "Afx/CGdiPlus/CGpPen.inc"
#include once "Afx/CGdiPlus/CGpRegion.inc"
#include once "Afx/CGdiPlus/CGpStringFormat.inc"

' // Not defined in the FB headers for 32-bit
#ifndef __FB_64BIT__
#define GetImageDecoders(numDecoders, size, decoders) cast(GpStatus, GdipGetImageDecoders((numDecoders), (size), (decoders)))
#define GetImageDecodersSize(numDecoders, size) cast(GpStatus, GdipGetImageDecodersSize((numDecoders), (size)))
#define GetImageEncoders(numEncoders, size, encoders) cast(GpStatus, GdipGetImageEncoders((numEncoders), (size), (encoders)))
#define GetImageEncodersSize(numEncoders, size) cast(GpStatus, GdipGetImageEncodersSize((numEncoders), (size)))
#endif
