# CGraphCtx Class

`CGraphCtx` is a graphic control for pictures, text and graphics. You can use both GDI and GDI+ to draw graphics and text and to load and manipulate images. Optionally, you can add support for OpenGL by passing "OPENGL" in the caption.

This control features persistence and uses a virtual buffer (set initially equal to the size of the control when it is created) to allow to display images bigger than the size of the control. Scrollbars are automatically added when the size of the virtual buffer is bigger than the size of the control and removed when unneeded. It also features keyboard navigation and sends command messages to the parent window when the return or Escape keys are pressed, and notification messages for mouse clicks.

To use the control, include the CGraphCtx.inc file in your program and use the namespace **Afx**.

**Include file**: CGraphCtx.inc

### Constructor

Registers the class name and creates an instance of the control.

```
CONSTRUCTOR CGraphCtx (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, _
   BYREF wszTitle AS WSTRING = "", BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, _
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

| Name       | Description |
| ---------- | ----------- |
| [Example 1](#Skeleton) | CWindow Graphic Control Skeleton. |
| [Example 2](#OpenGLSkeleton) | CWindow OpenGL Graphic Control Skeleton. |

### Helper Procedure: AfxCGraphPtr

Returns a pointer to the CGraphCtx class given the handle of its associated window.

```
FUNCTION AfxCGraphPtr (BYVAL hwnd AS HWND) AS CGraphCtx PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle of the window associated with the graphic control. Call the **hWindow** method of the **CGraphCtx** class to retrieve it. |

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [Clear](#Clear) | Clears the graphic control with the specified RGB color. |
| [CreateBitmapFromFile](#CreateBitmapFromFile) | Loads and displays the specified image in the Graphic Control. It also allows to convert the image to gray scale and/or dim the image. |
| [DrawBitmap](#DrawBitmap) | Draws a bitmap in the Graphic Control. |
| [GetBits](#GetBits) | Returns the location of the DIB bit values. |
| [GethBmp](#GethBmp) | Returns the handle of the compatible bitmap. |
| [GethRC](#GethRC) | If OpenGL is enabled, it returns the handle of the rendering context of the control. |
| [GetMemDC](#GetMemDC) | Handle to  the memory device context of the control. |
| [GetVirtualBufferHeight](#GetVirtualBufferHeight) | Returns the height of the virtual buffer. |
| [GetVirtualBufferWidth](#GetVirtualBufferWidth) | Returns the width of the virtual buffer. |
| [hWindow](#hWindow) | Returns the handle of the control. |
| [LoadImageFromFile](#LoadImageFromFile) | Loads and displays the specified image in the Graphic Control. |
| [LoadImageFromRes](#LoadImageFromRes) | Loads the specified image from a resource file in the Graphic Control. It also allows to convert the image to gray scale and/or dim the image. |
| [MakeCurrent](#MakeCurrent) | Redirects OpenGL calls are directed to the correct rendering context. |
| [PrintImage](#PrintImage) | Prints the image in the default printer. |
| [Resizable](#Resizable) | Gets/sets the value of the Resizable property. |
| [SaveImage](#SaveImage) | Saves the image to a file. |
| [SetVirtualBufferSize](#SetVirtualBufferSize) | Sets the size of the virtual buffer. |
| [Stretchable](#Stretchable) | Gets/sets the value of the Stretchable property. |
| [StretchMode](#StretchMode) | Gets/sets the value of the StretchMode property. |

### Notification Messages

| Name       | Description |
| ---------- | ----------- |
| [NM_CLICK](#NM_CLICK) | Sent by the control when the user clicks it with the left mouse button. |
| [NM_DBLCLK](#NM_DBLCLK) | Sent by the control when the user double clicks it with the left mouse button. |
| [NM_KILLFOCUS](#NM_KILLFOCUS) | Notifies a control's parent window that the control has lost the input focus. |
| [NM_RCLICK](#NM_RCLICK) | Sent by the control when the user clicks it with the right mouse button. |
| [NM_RDBLCLK](#NM_RDBLCLK) | Sent by the control when the user double clicks it with the right mouse button. |
| [NM_SETFOCUS](#NM_SETFOCUS) | Notifies a control's parent window that the control has received the input focus. |

### <a name="NM_CLICK"></a>NM_CLICK Notification Message

Sent by the control when the user clicks it with the left mouse button. This notification code is sent in the form of a WM_NOTIFY message.

```
CASE WM_NOTIFY
   DIM phdr AS NMHDR PTR = CAST(NMHDR PTR, lParam)
   IF wParam = IDC_GRCTX THEN
      SELECT CASE phdr->code
         CASE NM_CLICK
            ' Left button clicked
      END SELECT
   END IF
END IF
```

#### Remarks

IDC_GRCTX is the constant value used as identifier of the control. Change it if needed.

### <a name="NM_DBLCLK"></a>NM_DBLCLK Notification Message

Sent by the control when the user double clicks it with the left mouse button. This notification code is sent in the form of a WM_NOTIFY message.

```
CASE WM_NOTIFY
   DIM phdr AS NMHDR PTR = CAST(NMHDR PTR, lParam)
   IF wParam = IDC_GRCTX THEN
      SELECT CASE phdr->code
         CASE NM_DBLCLK
            ' Left button double clicked
      END SELECT
   END IF
END IF
```

#### Remarks

IDC_GRCTX is the constant value used as identifier of the control. Change it if needed.

### <a name="NM_KILLFOCUS"></a>NM_KILLFOCUS Notification Message

Notifies a control's parent window that the control has lost the input focus. This notification code is sent in the form of a WM_NOTIFY message. 

```
CASE WM_NOTIFY
   DIM phdr AS NMHDR PTR = CAST(NMHDR PTR, lParam)
   IF wParam = IDC_GRCTX THEN
      SELECT CASE phdr->code
         CASE NM_KILLFOCUS
            ' The control has lost focus
      END SELECT
   END IF
END IF
```

#### Remarks

IDC_GRCTX is the constant value used as identifier of the control. Change it if needed.

### <a name="NM_RCLICK"></a>NM_RCLICK Notification Message

Notifies a control's parent window that the control has lost the input focus. This notification code is sent in the form of a WM_NOTIFY message. 

```
CASE WM_NOTIFY
   DIM phdr AS NMHDR PTR = CAST(NMHDR PTR, lParam)
   IF wParam = IDC_GRCTX THEN
      SELECT CASE phdr->code
         CASE NM_RCLICK
            ' Right button clicked
      END SELECT
   END IF
END IF
```

#### Remarks

IDC_GRCTX is the constant value used as identifier of the control. Change it if needed.

### <a name="NM_RDBLCLK"></a>NM_RDBLCLK Notification Message

Sent by the control when the user double clicks it with the right mouse button. This notification code is sent in the form of a WM_NOTIFY message.

```
CASE WM_NOTIFY
   DIM phdr AS NMHDR PTR = CAST(NMHDR PTR, lParam)
   IF wParam = IDC_GRCTX THEN
      SELECT CASE phdr->code
         CASE NM_RDBLCLK
            ' Right button double clicked
      END SELECT
   END IF
END IF
```

#### Remarks

IDC_GRCTX is the constant value used as identifier of the control. Change it if needed.

### <a name="NM_SETFOCUS"></a>NM_SETFOCUS Notification Message

Notifies a control's parent window that the control has received the input focus. This notification code is sent in the form of a WM_NOTIFY message. 

```
CASE WM_NOTIFY
   DIM phdr AS NMHDR PTR = CAST(NMHDR PTR, lParam)
   IF wParam = IDC_GRCTX THEN
      SELECT CASE phdr->code
         CASE NM_SETFOCUS
            ' The control has gained focus
      END SELECT
   END IF
END IF
```

#### Remarks

IDC_GRCTX is the constant value used as identifier of the control. Change it if needed.

### <a name="Clear"></a>Clear

Clears the graphic control with the specified RGB color.

```
FUNCTION Clear (BYVAL RGBColor AS COLORREF) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *RGBColor* | RGB color used to fill the control. |

#### Return value

TRUE or FALSE.

### <a name="CreateBitmapFromFile"></a>CreateBitmapFromFile

Loads and displays the specified image in the Graphic Control. It also allows to convert the image to gray scale and/or dim the image.

```
SUB CreateBitmapFromFile (BYREF wszFileName AS WSTRING, _
   BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileName* | Absolute path to the file. |
| *dimPercent* | Optional. Percent of dimming (1-99). |
| *bGrayScale* | TRUE or FALSE. Convert to gray scale. |

#### Return value

TRUE or FALSE.

### <a name="DrawBitmap"></a>DrawBitmap

Draws a bitmap in the Graphic Control.

```
SUB DrawBitmap (BYVAL hbmp AS HBITMAP, BYVAL x AS SINGLE = 0, BYVAL y AS SINGLE = 0, _
   BYVAL nRight AS SINGLE = 0, BYVAL nBottom AS SINGLE = 0) AS GpStatus
```
```
SUB DrawBitmap (BYVAL pBitmap AS GpBitmap PTR, BYVAL x AS SINGLE = 0, BYVAL y AS SINGLE = 0, _
   BYVAL nRight AS SINGLE = 0, BYVAL nBottom AS SINGLE = 0) AS GpStatus
```
```
SUB DrawBitmap (BYREF pMemBmp AS CMemBmp, BYVAL x AS SINGLE = 0, BYVAL y AS SINGLE = 0, _
   BYVAL nRight AS SINGLE = 0, BYVAL nBottom AS SINGLE = 0) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *hbmp* | A handle to the bitmap. |
| *pBitmap* | Pointer to a GDI+ bitmap. |
| *pMemBmp* | Reference to a memory bitmap (CMemBmp class). |
| *x* | Optional. Left position. |
| *y* | Optional. Top position. |
| *nRight* | Optional. Right position.. |
| *nBottom* | Optional. Bottom position.. |

#### Return value

Returns OK (0) on success or a GdiPlus status code on failure.

#### Remarks

If both *x* and *y* are ommited, the image is draw starting at position 0, 0.

If *nRight* and *nBottom* are specified, the image is draw stretched in the bounding rectangle formed by *x*, *y*, *nRight* and *nBottom*.

#### Sample code

```
DIM pMemBmp AS CMemBmp = CMemBmp($"C:\Users\Pepe\Pictures\Cole_Kyla_01.jpg")
Rectangle pMemBmp.GetMemDC, 10, 10, 150, 150
LineTo pMemBmp.GetMemDC, 30, 180
pGraphCtx.DrawBitmap pMemBmp
```

### <a name="GetBits"></a>GetBits

Returns the location of the DIB bit values.

```
FUNCTION GetBits () AS ANY PTR
```

#### Return value

Pointer to the location of the DIB bit values.

### <a name="GethBmp"></a>GethBmp

Returns the handle of the compatible bitmap.

```
FUNCTION GethBmp () AS HBITMAP
```

### <a name="GethRC"></a>GethRC

If OpenGL is enabled, it returns the handle of the rendering context of the control.

```
FUNCTION GethRC () AS HGLRC
```

### <a name="GetMemDC"></a>GetMemDC

Returns the handle of the memory device context of the control.

```
FUNCTION GetMemDC () AS HDC
```

### <a name="GetVirtualBufferHeight"></a>GetVirtualBufferHeight

Returns the height of the virtual buffer.

```
FUNCTION GetVirtualBufferHeight () AS LONG
```

### <a name="GetVirtualBufferWidth"></a>GetVirtualBufferWidth

Returns the width of the virtual buffer.

```
FUNCTION GetVirtualBufferWidth () AS LONG
```

### <a name="hWindow"></a>hWindow

Returns the handle of the control.

```
FUNCTION hWindow () AS HWND
```

### <a name="LoadImageFromFile"></a>LoadImageFromFile

Loads and displays the specified image in the Graphic Control.

```
SUB LoadImageFromFile (BYREF wszFileName AS WSTRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileName* | Absolute path to the file. |

#### Remarks

A quirk in the GDI+ **GdipCreateBitmapFromFile** function causes that black and white images are loaded with increased contrast. Therefore, it's better to use the **CreateBitmapFromFile** for black and white images.

### <a name="LoadImageFromRes"></a>LoadImageFromRes

Loads the specified image from a resource file in the Graphic Control. It also allows to convert the image to gray scale and/or dim the image.

```
SUB LoadImageFromRes (BYVAL hInst AS HINSTANCE, BYREF wszImageName AS WSTRING, _
   BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE, _
   BYVAL clrBackground AS ARGB = 0)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hInst* | A handle to the module whose portable executable file or an accompanying MUI file contains the resource. If this parameter is NULL, the function searches the module used to create the current process. |
| *wszImageName* | Name of the image in the resource file (.RES). If the image resource uses an integral identifier, wszImage should begin with a number symbol (#) followed by the identifier in an ASCII format, e.g., "#998". Otherwise, use the text identifier name for the image. Only images embedded as raw data (type RCDATA) are valid. These must be icons in format .png, .jpg, .gif, .tiff. |
| *dimPercent* | Percent of dimming (1-99). |
| *bGrayScale* | TRUE or FALSE. Convert to gray scale. |
| *clrBackground* | The background color. This parameter is ignored if the bitmap is totally opaque. |

### <a name="MakeCurrent"></a>MakeCurrent

As more than one instance of this control can be used on a form, we need to make sure that OpenGL calls are directed to the correct rendering context. This is achieved by calling the **MakeCurrent** method.

```
FUNCTION MakeCurrent () AS BOOLEAN
```
#### Return value

TRUE or FALSE.

### <a name="PrintImage"></a>PrintImage

Prints the image in the default printer.

```
FUNCTION PrintImage (BYVAL bStretch AS BOOLEAN = FALSE, _
   BYVAL nStretchMode AS LONG = InterpolationModeHighQualityBicubic) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *bStretch* | Optional. Stretch the image. |
| *nStretchMode* | Optional. Stretching mode.<br>Default value = InterpolationModeHighQualityBicubic. |

### InterpolationMode Enumeration

| Constant   | Description |
| ---------- | ----------- |
| **InterpolationModeDefault** | Specifies the default interpolation mode. |
| **InterpolationModeLowQuality** | Specifies a low-quality mode. |
| **InterpolationModeHighQuality** | Specifies a high-quality mode. |
| **InterpolationModeBilinear** | Specifies bilinear interpolation. No prefiltering is done. This mode is not suitable for shrinking an image below 50 percent of its original size. |
| **InterpolationModeBicubic** | Specifies bicubic interpolation. No prefiltering is done. This mode is not suitable for shrinking an image below 25 percent of its original size. |
| **InterpolationModeNearestNeighbor** | Specifies nearest-neighbor interpolation. |
| **InterpolationModeHighQualityBilinear** | Specifies high-quality, bilinear interpolation. Prefiltering is performed to ensure high-quality shrinking. |
| **InterpolationModeHighQualityBicubic** | Specifies high-quality, bicubic interpolation. Prefiltering is performed to ensure high-quality shrinking. This mode produces the highest quality transformed images. |

#### Return value

Returns TRUE if the bitmap has been printed successfully, or FALSE otherwise.

### <a name="Resizable"></a>Resizable

Gets/sets the value of the **Resizable** property.

```
PROPERTY Resizable () AS BOOLEAN
PROPERTY Resizable (BYVAL bResizable AS BOOLEAN)
```

#### Return value

TRUE or FALSE.

##### Remarks

If resizable, the virtual buffer is set to the size of the control. If the control is made smaller and then bigger, part of the contents
are lost. Therefore, the caller must redraw it. Resizable and stretchable are mutually exclusive.

### <a name="SaveImage"></a>SaveImage

Saves the image to a file.

```
FUNCTION SaveImage (BYREF wszFileName AS WSTRING, BYREF wszMimeType AS WSTRING) AS LONG
```
```
FUNCTION SaveImageAsBmp (BYREF wszFileName AS WSTRING) AS LONG
FUNCTION SaveImageAsJpeg (BYREF wszFileName AS WSTRING) AS LONG
FUNCTION SaveImageAsPng (BYREF wszFileName AS WSTRING) AS LONG
FUNCTION SaveImageAsGif (BYREF wszFileName AS WSTRING) AS LONG
FUNCTION SaveImageAsTiff (BYREF wszFileName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileName* | Absolute path name of the file. |
| *wszMimeType* | The mime type.<br>"image/bmp" = Bitmap (.bmp)<br>"image/gif" = GIF (.gif)<br>"image/jpeg" = JPEG (.jpg)<br>"image/png" = PNG (.png)<br>"image/tiff" = TIFF (.tiff) |

#### Return value

If the method succeeds, it returns Ok, which is an element of the GDI+ Status enumeration.

If the method fails, it returns one of the other elements of the GDI+ Status enumeration.

### <a name="SetVirtualBufferSize"></a>SetVirtualBufferSize

Sets the size of the virtual buffer.

```
FUNCTION SetVirtualBufferSize (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nWidth* | Width, in pixels, of the virtual buffer. |
| *nHeight* | Height, in pixels, of the virtual buffer. |

#### Return value

If the function succeeds, the return value is S_OK (0).

### <a name="Stretchable"></a>Stretchable

Gets/sets the value of the **Stretchable** property.

```
PROPERTY Stretchable () AS BOOLEAN
PROPERTY Stretchable (BYVAL bStretchable AS BOOLEAN)
```

| Parameter  | Description |
| ---------- | ----------- |
| *bStretchable* | TRUE or FALSE. |

#### Return value

TRUE or FALSE.

### <a name="StretchMode"></a>StretchMode

Gets/sets the value of the **StretchMode** property.

```
PROPERTY StretchMode () AS LONG
PROPERTY StretchMode (BYVAL nStretchMode AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nStretchMode* | The stretching mode. |

| Mode       | Description |
| ---------- | ----------- |
| **BLACKONWHITE** | Performs a Boolean AND operation using the color values for the eliminated and existing pixels. If the bitmap is a monochrome bitmap, this mode preserves black pixels at the expense of white pixels. |
| **COLORONCOLOR** | Deletes the pixels. This mode deletes all eliminated lines of pixels without trying to preserve their information. |
| **HALFTONE** | Maps pixels from the source rectangle into blocks of pixels in the destination rectangle. The average color over the destination block of pixels approximates the color of the source pixels. After setting the HALFTONE stretching mode, an application must call the **SetBrushOrgEx** function to set the brush origin. If it fails to do so, brush misalignment occurs. |
| **STRETCH_ANDSCANS** | Same as BLACKONWHITE. |
| **STRETCH_DELETESCANS** | Same as COLORONCOLOR. |
| **STRETCH_HALFTONE** | Same as HALFTONE. |
| **STRETCH_ORSCANS** | Same as WHITEONBLACK. |
| **WHITEONBLACK** | Performs a Boolean OR operation using the color values for the eliminated and existing pixels. If the bitmap is a monochrome bitmap, this mode preserves white pixels at the expense of black pixels. |

#### Return value

The previous value of the property.

### <a name="Skeleton"></a>CWindow Graphic Control Skeleton

```
' ########################################################################################
' Microsoft Windows
' Contents: CWindow Graphic Control Skeleton
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#INCLUDE ONCE "windows.bi"
#INCLUDE ONCE "Afx/CWindow.inc"
#INCLUDE ONCE "Afx/CGraphCtx.inc"
USING Afx

CONST IDC_GRCTX = 1001

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG

   END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

' ========================================================================================
' Window procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hWnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   SELECT CASE uMsg

      CASE WM_CREATE
         EXIT FUNCTION

      CASE WM_COMMAND
         SELECT CASE LOWORD(wParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application sending an WM_CLOSE message
               IF HIWORD(wParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF
         END SELECT

      CASE WM_SIZE
         ' // If the window isn't minimized, resize it
         IF wParam <> SIZE_MINIMIZED THEN
            DIM pWindow AS CWindow PTR = CAST(CWindow PTR, GetWindowLongPtrW(hwnd, 0))
            pWindow->MoveWindow GetDlgItem(hwnd, IDC_GRCTX), 0, 0, pWindow->ClientWidth, pWindow->ClientHeight, CTRUE
         END IF

    	CASE WM_DESTROY
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   FUNCTION = DefWindowProcW(hWnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Main
' ========================================================================================
FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                  BYVAL hPrevInstance AS HINSTANCE, _
                  BYVAL szCmdLine AS ZSTRING PTR, _
                  BYVAL nCmdShow AS LONG) AS LONG

   ' // Set process DPI aware
   AfxSetProcessDPIAware

   DIM pWindow AS CWindow
   pWindow.Create(NULL, "CWindow Graphic Control Skeleton", @WndProc)
   pWindow.SetClientSize(600, 350)
   pWindow.Center

   ' // Add a graphic control
   DIM pGraphCtx AS CGraphCtx = CGraphCtx(@pWindow, IDC_GRCTX, "", 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   pGraphCtx.Clear BGR(255, 255, 255)

   ' // Capture the desktop window and display it in the control
   DIM hBitmap AS HBITMAP = AfxCaptureDisplay
   pGraphCtx.SetVirtualBufferSize(AfxGetBitmapWidth(hBitmap), AfxGetBitmapHeight(hBitmap))
   AfxDrawBitmap(pGraphCtx.GetMemDC, 0, 0, hBitmap)
   DeleteObject hBitmap

   ' // Process Windows events
   FUNCTION = pWindow.DoEvents(nCmdShow)

END FUNCTION
' ========================================================================================
```

### <a name="OpenGLSkeleton"></a>CWindow OpenGL Graphic Control Skeleton

The following example demonstrates the use of OpenGL and GDI+ together.

```
' ########################################################################################
' Microsoft Windows
' Contents: CWindow OpenGL Graphic Control Skeleton
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "Afx/CWindow.inc"
#INCLUDE ONCE "Afx/AfxGdiPlus.inc"
#INCLUDE ONCE "Afx/CGraphCtx.inc"
USING Afx

CONST IDC_GRCTX = 1001

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG

   END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

' ========================================================================================
' The following sample code draws a line.
' Change it with your own code.
' ========================================================================================
SUB GDIP_Render (BYVAL hDC AS HDC)

   DIM hStatus AS GpStatus
   DIM pGraphics AS GpGraphics PTR
   DIM pPen AS GpPen PTR

   ' // Create a graphics object
   hStatus = GdipCreateFromHDC(hDC, @pGraphics)
   ' // Create a Pen
   hStatus = GdipCreatePen1(GDIP_ARGB(255, 255, 0, 0), AfxScaleX(1), UnitPixel, @pPen)
   ' // Draw the line
   GdipDrawLineI pGraphics, pPen, 0, 0, AfxScaleX(200), AfxScaleY(100)

   ' // Cleanup
   IF pPen THEN GdipDeletePen(pPen)
   IF pGraphics THEN GdipDeleteGraphics(pGraphics)

END SUB
' ========================================================================================

' =======================================================================================
' All the setup goes here
' =======================================================================================
SUB RenderScene (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG)

   ' // Specify clear values for the color buffers
   glClearColor 0.0, 0.0, 0.0, 0.0
   ' // Specify the clear value for the depth buffer
   glClearDepth 1.0
   ' // Specify the value used for depth-buffer comparisons
   glDepthFunc GL_LESS
   ' // Enable depth comparisons and update the depth buffer
   glEnable GL_DEPTH_TEST
   ' // Select smooth shading
   glShadeModel GL_SMOOTH

   ' // Prevent divide by zero making height equal one
   IF nHeight = 0 THEN nHeight = 1
   ' // Reset the current viewport
   glViewport 0, 0, nWidth, nHeight
   ' // Select the projection matrix
   glMatrixMode GL_PROJECTION
   ' // Reset the projection matrix
   glLoadIdentity
   ' // Calculate the aspect ratio of the window
   gluPerspective 45.0!, nWidth / nHeight, 0.1!, 100.0!
   ' // Select the model view matrix
   glMatrixMode GL_MODELVIEW
   ' // Reset the model view matrix
   glLoadIdentity

   ' // Clear the screen buffer
   glClear GL_COLOR_BUFFER_BIT OR GL_DEPTH_BUFFER_BIT
   ' // Reset the view
   glLoadIdentity

   glTranslatef -1.5!, 0.0!, -6.0!       ' Move left 1.5 units and into the screen 6.0

   glBegin GL_TRIANGLES                 ' Drawing using triangles
      glColor3f   1.0!, 0.0!, 0.0!       ' Set the color to red
      glVertex3f  0.0!, 1.0!, 0.0!       ' Top
      glColor3f   0.0!, 1.0!, 0.0!       ' Set the color to green
      glVertex3f  1.0!,-1.0!, 0.0!       ' Bottom right
      glColor3f   0.0!, 0.0!, 1.0!       ' Set the color to blue
      glVertex3f -1.0!,-1.0!, 0.0!       ' Bottom left
   glEnd                                 ' Finished drawing the triangle

   glTranslatef 3.0!,0.0!,0.0!           ' Move right 3 units

   glColor3f 0.5!, 0.5!, 1.0!            ' Set the color to blue one time only
   glBegin GL_QUADS                     ' Draw a quad
      glVertex3f -1.0!, 1.0!, 0.0!       ' Top left
      glVertex3f  1.0!, 1.0!, 0.0!       ' Top right
      glVertex3f  1.0!,-1.0!, 0.0!       ' Bottom right
      glVertex3f -1.0!,-1.0!, 0.0!       ' Bottom left
   glEnd                                 ' Done drawing the quad

   ' // Required: force execution of GL commands in finite time
   glFlush

END SUB
' =======================================================================================

' ========================================================================================
' Main
' ========================================================================================
FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                  BYVAL hPrevInstance AS HINSTANCE, _
                  BYVAL szCmdLine AS ZSTRING PTR, _
                  BYVAL nCmdShow AS LONG) AS LONG

   ' // Set process DPI aware
   AfxSetProcessDPIAware

   ' // Create the main window
   DIM pWindow AS CWindow
   pWindow.Create(NULL, "CWindow OpenGL Graphic Control Skeleton", @WndProc)
   pWindow.Brush = CAST(HBRUSH, COLOR_WINDOW + 1)
   pWindow.SetClientSize(400, 250)
   pWindow.Center

   ' // Initialize GDI+
   DIM token AS UINT_PTR = AfxGdipInit

   ' // Add a graphic control with OPENGL enabled
   DIM pGraphCtx AS CGraphCtx = CGraphCtx(@pWindow, IDC_GRCTX, "OPENGL", 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   ' // Make the control stretchable
   pGraphCtx.Stretchable = TRUE
   ' // Make current the rendering context
   pGraphCtx.MakeCurrent
   ' // Render the OpenGL scene
   RenderScene pGraphCtx.GetVirtualBufferWidth, pGraphCtx.GetVirtualBufferHeight
   ' // Draw the GDI+ graphics
   GDIP_Render pGraphCtx.GetMemDC

   ' // Dispatch Windows events
   FUNCTION = pWindow.DoEvents(nCmdShow)

   ' // Shutdown GDI+
   IF token THEN GdiPlusShutdown token

END FUNCTION
' ========================================================================================

' ========================================================================================
' Main window callback procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   SELECT CASE uMsg

      CASE WM_COMMAND
         SELECT CASE GET_WM_COMMAND_ID(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF
         END SELECT

      CASE WM_SIZE
         ' // If the window isn't minimized, resize it
         IF wParam <> SIZE_MINIMIZED THEN
            DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
            IF pWindow THEN pWindow->MoveWindow GetDlgItem(hwnd, IDC_GRCTX), 0, 0, pWindow->ClientWidth, pWindow->ClientHeight, CTRUE
         END IF

    	CASE WM_DESTROY
         ' // End the application
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hWnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
```
