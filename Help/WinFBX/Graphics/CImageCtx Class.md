# CImageCtx Class

**CImageCtx** is an image control based on GDI+. The images can be rendered using its actual size or can be resized to fit the width or the height of the window. Resizing is done automatically by the control.

To use the control, include the **CImageCtx.inc** file in your program and use the namespace **Afx**.

Include file: CImageCtx.inc

### Constructor

Registers the class name and creates an instance of the control.

```
CONSTRUCTOR CImageCtx (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, _
   BYREF wszTitle AS CONST WSTRING = "", BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, _
   BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, BYVAL dwStyle AS DWORD = 0, _
   BYVAL dwExStyle AS DWORD = 0, BYVAL lpParam AS LONG_PTR = 0)
```
| Parameter  | Description |
| ---------- | ----------- |
| *pWindow* | Pointer to the **CWindow** class of the parent window. |
| *cID* | Control identifier. |
| *wszTitle* | The window caption. If "OPENGL" is used as the caption, support for OpenGL is added to the control. |
| *x* | The x-coordinate of the upper-left corner of the window relative to the upper-left corner of the parent window's client area. |
| *y* | The initial y-coordinate of the upper-left corner of the window relative to the upper-left corner of the parent window's client area. |
| *nWidth* | The width of the control. |
| *nHeight* | The height of the control. |
| *dwStyle* | The style of the window being created. Pass -1 to use the default styles.<br>Default styles: WS_VISIBLE OR WS_CHILD OR WS_TABSTOP. |
| *dwExStyle* | The extended window style of the control being created. Pass -1 to use the default styles. |
| *lpParam* | Pointer to custom data. |

#### Return value

A pointer to the new instance of the class.

#### Usage example:

```
' // Create the main window
DIM pWindow AS CWindow
pWindow.Create(NULL, "CWindow with an image control", @WndProc)
pWindow.SetClientSize(600, 350)
pWindow.Center

' // Add an image control
DIM pImageCtx AS CImageCtx = CImageCtx(@pWindow, IDC_IMAGECTX, "", 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
' // Load an image from disk
pImageCtx.LoadImageFromFile ExePath & "\image.jpg"
```

| Name       | Description |
| ---------- | ----------- |
| [Demo](#Demo) | CWindow CImage Control demo. |


### Helper Procedure: AfxCImageCtxPtr

Returns a pointer to the CIMageCtx class given the handle of its associated window.

```
FUNCTION AfxCImageCtxPtr (BYVAL hwnd AS HWND) AS CImageCtx PTR
FUNCTION AfxCImageCtxPtr (BYVAL hParent AS HWND, BYVAL cID AS LONG) AS CImageCtx PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle of the window associated with the graphic control. Call the **hWindow** method of the **CImageCtx** class to retrieve it. |
| *hParent* | The handle of the parent window of the control. |
| *cID* | The identifier of the control. |

#### Return value

A pointer to the CImageCtx class.

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [Clear](#Clear) | Clears the contents of the controlr. |
| [GetBkColor](#GetBkColor) | Gets the background RGB color used by the **CImageCtx** control. |
| [GetBkColorHot](#GetBkColorHot) | Gets the background hot RGB color used by the **CImageCtx** control. |
| [GetImageAdjustment](#GetImageAdjustment) | Gets the image adjustment setting used by the control. |
| [GetImageHeight](#GetImageHeight) | Gets the height of the image, in pixels, currently loaded in the **CImageCtx** control. |
| [GetImagePtr](#GetImagePtr) | Gets a pointer to the GDI+ GpImage object used by the control to render the loaded image. |
| [GetImageWidth](#GetImageWidth) | Gets the width of the image, in pixels, currently loaded in the control. |
| [GetInterpolationMode](#GetInterpolationMode) | Gets the interpolation mode used by GDI+. |
| [hWindow](#hWindow) | Returns the handle of the control. |
| [LoadBitmapFromResource](#LoadBitmapFromResource) | Loads a bitmap from a resource file into the control. |
| [LoadImageFromFile](#LoadImageFromFile) | Loads an image from disk into the control. |
| [LoadImageFromResource](#LoadImageFromResource) | Loads an image from a resource file into the control. |
| [Redraw](#Redraw) | Redraws the **CImageCtx** control. |
| [SetBkColor](#SetBkColor) | Sets the background RGB color used by the **CImageCtx** control. |
| [SetBkColorHot](#SetBkColorHot) | Sets the background hot RGB color used by the **CImageCtx** control. |
| [SetImageAdjustment](#SetImageAdjustment) | Sets the image adjustment setting used by the control. |
| [SetImageHeight](#SetImageHeight) | Sets the height of the image. |
| [SetImageSize](#SetImageSize) | Sets the size of the image. |
| [SetImageWidth](#SetImageWidth) | Sets the width of the image. |
| [SetInterpolationMode](#SetInterpolationMode) | Sets the interpolation mode used by GDI+. |

### Notification Messages

| Name       | Description |
| ---------- | ----------- |
| [NM_CLICK](#NM_CLICK) | Sent by the control when the user clicks it with the left mouse button. |
| [NM_DBLCLK](#NM_DBLCLK) | Sent by the control when the user double clicks it with the left mouse button. |
| [NM_KILLFOCUS](#NM_KILLFOCUS) | Notifies a control's parent window that the control has lost the input focus. |
| [NM_RCLICK](#NM_RCLICK) | Sent by the control when the user clicks it with the right mouse button. |
| [NM_RDBLCLK](#NM_RDBLCLK) | Sent by the control when the user double clicks it with the right mouse button. |
| [NM_SETFOCUS](#NM_SETFOCUS) | Notifies a control's parent window that the control has received the input focus. |

# <a name="NM_CLICK"></a>NM_CLICK Notification Message

Sent by the control when the user clicks it with the left mouse button. This notification code is sent in the form of a WM_NOTIFY message.

```
CASE WM_NOTIFY
   DIM phdr AS NMHDR PTR = CAST(NMHDR PTR, lParam)
   IF wParam = IDC_IMGCTX THEN
      SELECT CASE phdr->code
         CASE NM_CLICK
            ' Left button clicked
      END SELECT
   END IF
END IF
```

#### Remarks

IDC_IMGCTX is the constant value used as identifier of the control. Change it if needed.

# <a name="NM_DBLCLK"></a>NM_DBLCLK Notification Message

Sent by the control when the user double clicks it with the left mouse button. This notification code is sent in the form of a WM_NOTIFY message.

```
CASE WM_NOTIFY
   DIM phdr AS NMHDR PTR = CAST(NMHDR PTR, lParam)
   IF wParam = IDC_IMGCTX THEN
      SELECT CASE phdr->code
         CASE NM_DBLCLK
            ' Left button double clicked
      END SELECT
   END IF
END IF
```

#### Remarks

IDC_IMGCTX is the constant value used as identifier of the control. Change it if needed.

# <a name="NM_KILLFOCUS"></a>NM_KILLFOCUS Notification Message

Notifies a control's parent window that the control has lost the input focus. This notification code is sent in the form of a WM_NOTIFY message. 

```
CASE WM_NOTIFY
   DIM phdr AS NMHDR PTR = CAST(NMHDR PTR, lParam)
   IF wParam = IDC_IMGCTX THEN
      SELECT CASE phdr->code
         CASE NM_KILLFOCUS
            ' The control has lost focus
      END SELECT
   END IF
END IF
```

#### Remarks

IDC_IMGCTX is the constant value used as identifier of the control. Change it if needed.

# <a name="NM_RCLICK"></a>NM_RCLICK Notification Message

Notifies a control's parent window that the control has lost the input focus. This notification code is sent in the form of a WM_NOTIFY message. 

```
CASE WM_NOTIFY
   DIM phdr AS NMHDR PTR = CAST(NMHDR PTR, lParam)
   IF wParam = IDC_IMGCTX THEN
      SELECT CASE phdr->code
         CASE NM_RCLICK
            ' Right button clicked
      END SELECT
   END IF
END IF
```

#### Remarks

IDC_IMGCTX is the constant value used as identifier of the control. Change it if needed.

# <a name="NM_RDBLCLK"></a>NM_RDBLCLK Notification Message

Sent by the control when the user double clicks it with the right mouse button. This notification code is sent in the form of a WM_NOTIFY message.

```
CASE WM_NOTIFY
   DIM phdr AS NMHDR PTR = CAST(NMHDR PTR, lParam)
   IF wParam = IDC_IMGCTX THEN
      SELECT CASE phdr->code
         CASE NM_RDBLCLK
            ' Right button double clicked
      END SELECT
   END IF
END IF
```

#### Remarks

IDC_IMGCTX is the constant value used as identifier of the control. Change it if needed.

# <a name="NM_SETFOCUS"></a>NM_SETFOCUS Notification Message

Notifies a control's parent window that the control has received the input focus. This notification code is sent in the form of a WM_NOTIFY message. 

```
CASE WM_NOTIFY
   DIM phdr AS NMHDR PTR = CAST(NMHDR PTR, lParam)
   IF wParam = IDC_IMGCTX THEN
      SELECT CASE phdr->code
         CASE NM_SETFOCUS
            ' The control has gained focus
      END SELECT
   END IF
END IF
```

#### Remarks

IDC_IMGCTX is the constant value used as identifier of the control. Change it if needed.

# <a name="Clear"></a>Clear

Clears the contents of the control.

```
SUB Clear
```

# <a name="GetBkColor"></a>GetBkColor

Gets the background RGB color used by the **CImageCtx** control.

```
FUNCTION GetBkColor () AS LONG
```

# <a name="GetBkColorHot"></a>GetBkColorHot

Gets the background hot RGB color used by the **CImageCtx** control (when the mouse is over the control).

```
FUNCTION GetBkColorHot () AS LONG
```

# <a name="GetImageAdjustment"></a>GetImageAdjustment

Gets the image adjustment setting used by the control.

```
FUNCTION GetImageAdjustment () AS LONG
```

#### Return value

The current adjustment setting, that can be one of the following values:

| Value      | Description |
| ---------- | ----------- |
| **GDIP_IMAGECTX_AUTOSIZE** | Autoadjusts the image to the width or height of the control. |
| **GDIP_IMAGECTX_ACTUALSIZE** | Shows the image with its actual size. |
| **GDIP_IMAGECTX_FITTOWIDTH** | Adjusts the image to the width of the control. |
| **GDIP_IMAGECTX_FITTOHEIGHT** | Adjusts the image to the height of the control. |
| **GDIP_IMAGECTX_STRETCH** | Adjusts the image to the height and width of the control. |

# <a name="GetImageHeight"></a>GetImageHeight

Gets the height of the image, in pixels, currently loaded in the **CImageCtx** control.

```
FUNCTION GetImageHeight () AS LONG
```

# <a name="GetImagePtr"></a>GetImagePtr

Gets a pointer to the GDI+ **GpImage** object used by the control to render the loaded image.

```
FUNCTION GetImagePtr () AS GpImage PTR
```

#### Return value

A copy of the pointer to the **GpImage** object used by the control to render the loaded image.

### Remarks

The returned pointer can be used to call GDI+ Image functions.

Don't dispose the returned pointer. The control keeps a copy of it and calls **GdipDisposeImage** when it is destroyed.

# <a name="GetImageWidth"></a>GetImageWidth

Gets the with of the image, in pixels, currently loaded in the **CImageCtx** control.

```
FUNCTION GetImageWidth () AS LONG
```

# <a name="GetInterpolationMode"></a>GetInterpolationMode

Gets the interpolation mode used by GDI+.

```
FUNCTION GetInterpolationMode () AS LONG
```

#### Return value

The current interpolation mode value.

```
InterpolationModeInvalid = -1
InterpolationModeDefault = 0
InterpolationModeLowQuality = 1
InterpolationModeHighQuality = 2
InterpolationModeBilinear = 3
InterpolationModeBicubic = 4
InterpolationModeNearestNeighbor = 5
InterpolationModeHighQualityBilinear = 6
InterpolationModeHighQualityBicubic = 7
```

# <a name="hWindow"></a>hWindow

Returns the handle of the control.

```
FUNCTION hWindow () AS HWND
```

# <a name="LoadBitmapFromResource"></a>LoadBitmapFromResource

Loads a bitmap from a resource file into the control.
**Note**: Works with .bmp and .png files, but not with .jpg and .tif files.

```
FUNCTION LoadBitmapFromResource (BYVAL hInstance AS HINSTANCE, BYREF wszResourceName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *hInstance* | Handle to the instance of the module whose executable file contains the resource. A value of NULL specifies the module handle associated with the image file that the operating system used to create the current process.  |
| *wszResourceName* | The name of the resource. |

#### Return value

Ok (0) or an error code.

# <a name="LoadImageFromFile"></a>LoadImageFromFile

Loads an image from disk into the control.

```
FUNCTION LoadImageFromFile (BYREF wszFileName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileName* | Fully qualified path and filename of the image file to load. |

#### Return value

Ok (0) or an error code.

# <a name="LoadImageFromResource"></a>LoadImageFromResource

Loads an image from a resource file into the control.
**Note**: Works with .bmp and .png files, but not with .jpg and .tif files.

```
FUNCTION LoadImageFromResource (BYVAL hInstance AS HINSTANCE, BYREF wszResourceName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *hInstance* | Handle to the instance of the module whose executable file contains the resource. A value of NULL specifies the module handle associated with the image file that the operating system used to create the current process.  |
| *wszResourceName* | The name of the resource. |

#### Return value

Ok (0) or an error code.

# <a name="Redraw"></a>Redraw

Redraws the **CImageCtx** control.

```
SUB Redraw
```

# <a name="SetBkColor"></a>SetBkColor

Sets the background RGB color used by the **CImageCtx** control.

```
FUNCTION SetBkColor (BYVAL clr AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *clr* | The new RGB background color. |
| *fRedraw* | Optional. TRUE to redraw the control to reflect the changes. |

# <a name="SetBkColorHot"></a>SetBkColorHot

Sets the background hot RGB color used by the **CImageCtx** control (when the mouse i over the control).

```
FUNCTION SetBkColorHot (BYVAL clr AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *clr* | The new RGB background color. |
| *fRedraw* | Optional. TRUE to redraw the control to reflect the changes. |

#### Return value

The previous background color.

# <a name="SetImageAdjustment"></a>SetImageAdjustment

Sets the image adjustment setting used by the control.

```
FUNCTION SetImageAdjustment (BYVAL ImageAdjustment AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *ImageAdjustment* | The setting to set. Can be one of the values listed below. |
| *fRedraw* | Optional. TRUE to redraw the control to reflect the changes. |

### Image adjustment constants

| Value      | Description |
| ---------- | ----------- |
| **GDIP_IMAGECTX_AUTOSIZE** | Autoadjusts the image to the width or height of the control. |
| **GDIP_IMAGECTX_ACTUALSIZE** | Shows the image with its actual size. |
| **GDIP_IMAGECTX_FITTOWIDTH** | Adjusts the image to the width of the control. |
| **GDIP_IMAGECTX_FITTOHEIGHT** | Adjusts the image to the height of the control. |
| **GDIP_IMAGECTX_STRETCH*** | Adjusts the image to the height and width of the control. |

#### Return value

The previous setting value.

# <a name="SetImageHeight"></a>SetImageHeight

Sets the height of the image.

```
SUB SetImageHeight (BYVAL nHeight AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nHeight* | Height of the image, in pixels. |
| *fRedraw* | Optional. TRUE to redraw the control to reflect the changes. |

# <a name="SetImageSize"></a>SetImageSize

Sets the size of the image.

```
SUB SetImageSize (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nWidth* | Width of the image, in pixels. |
| *nHeight* | Height of the image, in pixels. |
| *fRedraw* | Optional. TRUE to redraw the control to reflect the changes. |

# <a name="SetImageWidth"></a>SetImageWidth

Sets the width of the image.

```
SUB SetImageWidth (BYVAL nWidth AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nWidth* | Height of the image, in pixels. |
| *fRedraw* | Optional. TRUE to redraw the control to reflect the changes. |

# <a name="SetInterpolationMode"></a>SetInterpolationMode

Sets the interpolation mode used by GDI+. The interpolation mode determines the algorithm that is used when images are scaled or rotated. 

```
SUB SetInterpolationMode (BYVAL InterpolationMode AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *InterpolationMode* | The interpolation mode. Should be one of the following values: |

### Interpolation modes

```
InterpolationModeInvalid = -1
InterpolationModeDefault = 0
InterpolationModeLowQuality = 1
InterpolationModeHighQuality = 2
InterpolationModeBilinear = 3
InterpolationModeBicubic = 4
InterpolationModeNearestNeighbor = 5
InterpolationModeHighQualityBilinear = 6
InterpolationModeHighQualityBicubic = 7
```

#### Return value

The previous setting value.

# <a name="Demo"></a>Demo

```
' ########################################################################################
' Microsoft Windows
' File: ImageCtx_Demo
' Contents: Image control demo
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define _WIN32_WINNT &h0602
#INCLUDE ONCE "Afx/CWindow.inc"
#INCLUDE ONCE "Afx/CImageCtx.inc"

USING Afx

CONST IDC_LOAD              = 100
CONST IDC_IMAGECTX          = 101
CONST IDC_RADIO_AUTOSIZE    = 102
CONST IDC_RADIO_ACTUALSIZE  = 103
CONST IDC_RADIO_FITTOHEIGHT = 104
CONST IDC_RADIO_FITTOWIDTH  = 105
CONST IDC_RADIO_STRETCH     = 106

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG

   END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

' ========================================================================================
' Displays the File Open Dialog and loads the selected file in the image control.
' ========================================================================================
SUB AfxLoadFileDialog (BYVAL hwndOwner AS HWND, BYVAL pImageCtx AS CImageCtx PTR, BYVAL sigdnName AS SIGDN = SIGDN_FILESYSPATH)

   ' // Create an instance of the FileOpenDialog interface
   DIM hr AS LONG
   DIM pofd AS IFileOpenDialog PTR
   hr = CoCreateInstance(@CLSID_FileOpenDialog, NULL, CLSCTX_INPROC_SERVER, @IID_IFileOpenDialog, @pofd)
   IF pofd = NULL THEN EXIT SUB

   ' // Set the file types
   DIM rgFileTypes(1 TO 5) AS COMDLG_FILTERSPEC
   rgFileTypes(1).pszName = @WSTR("JPG Files (*.jpg)")
   rgFileTypes(2).pszName = @WSTR("BMP Files (*.bmp)")
   rgFileTypes(3).pszName = @WSTR("PNG Files (*.png)")
   rgFileTypes(4).pszName = @WSTR("TIF Files (*.tif)")
   rgFileTypes(5).pszName = @WSTR("All files")
   rgFileTypes(1).pszSpec = @WSTR("*.jpg")
   rgFileTypes(2).pszSpec = @WSTR("*.bmp")
   rgFileTypes(3).pszSpec = @WSTR("*.png")
   rgFileTypes(4).pszSpec = @WSTR("*.tif")
   rgFileTypes(5).pszSpec = @WSTR("*.*")
   pofd->lpVtbl->SetFileTypes(pofd, 5, @rgFileTypes(1))

   ' // Set the title of the dialog
   hr = pofd->lpVtbl->SetTitle(pofd, "A Single-Selection Dialog")
   ' // Display the dialog
   hr = pofd->lpVtbl->Show(pofd, hwndOwner)

   ' // Set the default folder
   DIM pFolder AS IShellItem PTR
   SHCreateItemFromParsingName (CURDIR, NULL, @IID_IShellItem, @pFolder)
   IF pFolder THEN
      pofd->lpVtbl->SetDefaultFolder(pofd, pFolder)
      pFolder->lpVtbl->Release(pFolder)
   END IF

   ' // Get the result
   DIM pItem AS IShellItem PTR
   DIM pwszName AS WSTRING PTR
   IF SUCCEEDED(hr) THEN
      hr = pofd->lpVtbl->GetResult(pofd, @pItem)
      IF SUCCEEDED(hr) THEN
         hr = pItem->lpVtbl->GetDisplayName(pItem, sigdnName, @pwszName)
         pImageCtx->LoadImageFromFile(*pwszName)
         CoTaskMemFree(pwszName)
      END IF
   END IF

   ' // Cleanup
   IF pItem THEN pItem->lpVtbl->Release(pItem)
   IF pofd THEN pofd->lpVtbl->Release(pofd)

END SUB
' ========================================================================================

' ========================================================================================
' Main
' ========================================================================================
FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                  BYVAL hPrevInstance AS HINSTANCE, _
                  BYVAL szCmdLine AS ZSTRING PTR, _
                  BYVAL nCmdShow AS LONG) AS LONG

   ' // Initialize the COM library
   CoInitialize(NULL)

   ' // Set process DPI aware
   ' // The recommended way is to use a manifest
'   AfxSetProcessDPIAware

   ' // Create the main window
   DIM pWindow AS CWindow
   pWindow.Create(NULL, "CWindow with an image control", @WndProc)
   pWindow.SetClientSize(600, 350)
   pWindow.Center

   ' // Add a button without coordinates (it will be reiszed in WM_SIZE, below)
   DIM hCtl AS HWND = pWindow.AddControl("RadioButton", , IDC_RADIO_AUTOSIZE, "Autosize")
   SendMessageW hCtl, BM_SETCHECK, BST_CHECKED, 0
   SetFocus hCtl
   pWindow.AddControl("RadioButton", , IDC_RADIO_ACTUALSIZE, "Actual size")
   pWindow.AddControl("RadioButton", , IDC_RADIO_FITTOWIDTH, "Fit to width")
   pWindow.AddControl("RadioButton", , IDC_RADIO_FITTOHEIGHT, "Fit to height")
   pWindow.AddControl("RadioButton", , IDC_RADIO_STRETCH, "Stretch")
   pWindow.AddControl("Button", , IDC_LOAD, "Load")
   pWindow.AddControl("Button", , IDCANCEL, "Close")

   ' // Add an image control
   DIM pImageCtx AS CImageCtx = CImageCtx(@pWindow, IDC_IMAGECTX, , 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   ' // Load an image from disk
   pImageCtx.LoadImageFromFile ".\Images\Ciutat_de_les_Arts_i_de_les_Ciencies_01.jpg"


   ' // Process Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

   ' // Uninitialize the COM library
   CoUninitialize

END FUNCTION
' ========================================================================================

' ========================================================================================
' Window procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   DIM pWindow AS CWindow PTR, pImageCtx AS CImageCtx PTR

   SELECT CASE uMsg

      CASE WM_COMMAND
         SELECT CASE GET_WM_COMMAND_ID(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application sending an WM_CLOSE message
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF
            CASE IDC_LOAD
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  ' // Load the picture
                  AfxLoadFileDialog hwnd, AfxCImageCtxPtr(hwnd, IDC_IMAGECTX)
               END IF
            CASE IDC_RADIO_AUTOSIZE
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  AfxCImageCtxPtr(hwnd, IDC_IMAGECTX)->SetImageAdjustment GDIP_IMAGECTX_AUTOSIZE, CTRUE
                  ' // Alternate way: Get a pointer to the CImageCtx class
'                  pImageCtx = CAST(CImageCtx PTR, GetWindowLongPtr(GetDlgItem(hwnd, IDC_IMAGECTX), 0))
'                  pImageCtx->SetImageAdjustment GDIP_IMAGECTX_AUTOSIZE, CTRUE
                  EXIT FUNCTION
               END IF
            CASE IDC_RADIO_ACTUALSIZE
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  AfxCImageCtxPtr(hwnd, IDC_IMAGECTX)->SetImageAdjustment GDIP_IMAGECTX_ACTUALSIZE, CTRUE
                  ' // Alternate way: Get a pointer to the CImageCtx class
'                  pImageCtx = CAST(CImageCtx PTR, GetWindowLongPtr(GetDlgItem(hwnd, IDC_IMAGECTX), 0))
'                  pImageCtx->SetImageAdjustment GDIP_IMAGECTX_ACTUALSIZE, CTRUE
                  EXIT FUNCTION
               END IF
            CASE IDC_RADIO_FITTOWIDTH
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  AfxCImageCtxPtr(hwnd, IDC_IMAGECTX)->SetImageAdjustment GDIP_IMAGECTX_FITTOWIDTH, CTRUE
                  ' // Alternate way: Get a pointer to the CImageCtx class
'                  pImageCtx = CAST(CImageCtx PTR, GetWindowLongPtr(GetDlgItem(hwnd, IDC_IMAGECTX), 0))
'                  pImageCtx->SetImageAdjustment GDIP_IMAGECTX_FITTOWIDTH, CTRUE
                  EXIT FUNCTION
               END IF
            CASE IDC_RADIO_FITTOHEIGHT
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  AfxCImageCtxPtr(hwnd, IDC_IMAGECTX)->SetImageAdjustment GDIP_IMAGECTX_FITTOHEIGHT, CTRUE
                  ' // Alternate way: Get a pointer to the CImageCtx class
'                  pImageCtx = CAST(CImageCtx PTR, GetWindowLongPtr(GetDlgItem(hwnd, IDC_IMAGECTX), 0))
'                  pImageCtx->SetImageAdjustment GDIP_IMAGECTX_FITTOHEIGHT, CTRUE
                  EXIT FUNCTION
               END IF
            CASE IDC_RADIO_STRETCH
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  AfxCImageCtxPtr(hwnd, IDC_IMAGECTX)->SetImageAdjustment GDIP_IMAGECTX_STRETCH, CTRUE
                  ' // Alternate way: Get a pointer to the CImageCtx class
'                  pImageCtx = CAST(CImageCtx PTR, GetWindowLongPtr(GetDlgItem(hwnd, IDC_IMAGECTX), 0))
'                  pImageCtx->SetImageAdjustment GDIP_IMAGECTX_STRETCH, CTRUE
                  EXIT FUNCTION
               END IF
         END SELECT

      CASE WM_NOTIFY
         ' // Process notification messages
         DIM phdr AS NMHDR PTR
         phdr = cast(NMHDR PTR, lParam)
         IF wParam = IDC_IMAGECTX THEN
            SELECT CASE phdr->code
               CASE NM_CLICK : MessageBox hwnd, "NM_CLICK", "", MB_OK
               CASE NM_DBLCLK
               CASE NM_RCLICK
               CASE NM_RDBLCLK
               CASE NM_SETFOCUS
               CASE NM_KILLFOCUS
            END SELECT
         END IF

      CASE WM_SIZE
         pWindow = AfxCWindowPtr(hwnd)
         ' // Alternate way:
'         pWindow = CAST(CWindow PTR, GetWindowLongPtr(hwnd, 0))
         ' // If the window isn't minimized, resize it
         IF wParam <> SIZE_MINIMIZED THEN
            pWindow->MoveWindow GetDlgItem(hwnd, IDC_IMAGECTX), 10, 10, pWindow->ClientWidth - 20, pWindow->ClientHeight - 65, CTRUE
            pWindow->MoveWindow GetDlgItem(hwnd, IDC_LOAD), pWindow->ClientWidth - 185, pWindow->ClientHeight - 35, 75, 23, CTRUE
            pWindow->MoveWindow GetDlgItem(hwnd, IDCANCEL), pWindow->ClientWidth - 95, pWindow->ClientHeight - 35, 75, 23, CTRUE
            pWindow->MoveWindow GetDlgItem(hwnd, IDC_RADIO_AUTOSIZE), 15, pWindow->ClientHeight - 49, 100, 21, CTRUE
            pWindow->MoveWindow GetDlgItem(hwnd, IDC_RADIO_ACTUALSIZE), 15, pWindow->ClientHeight - 28, 100, 21, CTRUE
            pWindow->MoveWindow GetDlgItem(hwnd, IDC_RADIO_FITTOWIDTH), 150, pWindow->ClientHeight - 49, 100, 21, CTRUE
            pWindow->MoveWindow GetDlgItem(hwnd, IDC_RADIO_FITTOHEIGHT), 150, pWindow->ClientHeight - 28, 100, 21, CTRUE
            pWindow->MoveWindow GetDlgItem(hwnd, IDC_RADIO_STRETCH), 275, pWindow->ClientHeight - 28, 60, 21, CTRUE
         END IF

    	CASE WM_DESTROY
         ' // Ends the application by sending a WM_QUIT message
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hWnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
```
