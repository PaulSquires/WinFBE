# CGpGraphics Class

The **CGpGraphics** class provides methods for drawing lines, curves, figures, images, and text. A **Graphics** object stores attributes of the display device and attributes of the items to be drawn.

**Inherits from**: CGpBase.<br>
**Imclude file**: CGpGraphics.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorsGraphics) | Creates a **Graphics** object. |
| [AddMetafileComment](#AddMetafileComment) | Adds a text comment to an existing metafile. |
| [BeginContainer](#BeginContainer) | Begins a new graphics container. |
| [Clear](#Clear) | Clears a **Graphics** object to a specified color. |
| [DrawArc](#DrawArc) | Draws an arc. The arc is part of an ellipse. |
| [DrawBezier](#DrawBezier) | Draws a Bézier spline. |
| [DrawBeziers](#DrawBeziers) | Draws a sequence of connected Bézier splines. |
| [DrawCachedBitmap](#DrawCachedBitmap) | Draws the image stored in a **CachedBitmap** object. |
| [DrawClosedCurve](#DrawClosedCurve) | Draws a closed cardinal spline. |
| [DrawCurve](#DrawCurve) | Draws a cardinal spline. |
| [DrawDriverString](#DrawDriverString) | Draws characters at the specified positions. |
| [DrawEllipse](#DrawEllipse) | Draws an ellipse. |
| [DrawImage](#DrawImage) | Draws an image. |
| [DrawImageFX](#DrawImageFX) | Draws a portion of an image after applying a specified effect. |
| [DrawLine](#DrawLine) | Draws a line that connects two points. |
| [DrawLines](#DrawLines) | Draws a sequence of connected lines. |
| [DrawPath](#DrawPath) | Draws a sequence of lines and curves defined by a **GraphicsPath** object. |
| [DrawPie](#DrawPie) | Draws a pie. |
| [DrawPolygon](#DrawPolygon) | Draws a polygon. |
| [DrawRectangle](#DrawRectangle) | Draws a rectangle. |
| [DrawRectangles](#DrawRectangles) | Draws a sequence of rectangles. |
| [DrawString](#DrawString) | Draws a string based on a font and an origin for the string. |
| [EndContainer](#EndContainer) | Closes a graphics container that was previously opened by the **BeginContainer** method. |
| [EnumerateMetafile](#EnumerateMetafile) | Calls an application-defined callback function for each record in a specified metafile. |
| [ExcludeClip](#ExcludeClip) | Updates the clipping region to the portion of itself that does not intersect the specified rectangle. |
| [FillClosedCurve](#FillClosedCurve) | Creates a closed cardinal spline from an array of points and uses a brush to fill the interior of the spline. |
| [FillEllipse](#FillEllipse) | Uses a brush to fill the interior of an ellipse that is specified by coordinates and dimensions. |
| [FillPath](#FillPath) | Uses a brush to fill the interior of a path. |
| [FillPie](#FillPie) | Uses a brush to fill the interior of a pie. |
| [FillPolygon](#FillPolygon) | Uses a brush to fill the interior of a polygon. |
| [FillRectangle](#FillRectangle) | Uses a brush to fill the interior of a rectangle. |
| [FillRectangles](#FillRectangles) | Uses a brush to fill the interior of a sequence of rectangles. |
| [FillRegion](#FillRegion) | Uses a brush to fill a specified region. |
| [Flush](#Flush) | Flushes all pending graphics operations. |
| [FromHDC](#FromHDC) | Creates a **Graphics** object that is associated with a specified device context. |
| [FromHWND](#FromHWND) | Creates a **Graphics** object that is associated with a specified window. |
| [FromImage](#FromImage) | Creates a **Graphics** object that is associated with a specified device context. |
| [GetClip](#GetClip) | Gets the clipping region of this **Graphics** object. |
| [GetClipBounds](#GetClipBounds) | Gets a rectangle that encloses the clipping region of this **Graphics** object. |
| [GetCompositingMode](#GetCompositingMode) | Gets the compositing mode currently set for this **Graphics** object. |
| [GetCompositingQuality](#GetCompositingQuality) | Gets the compositing quality currently set for this **Graphics** object. |
| [GetDpiX](#GetDpiX) | Gets the horizontal resolution, in dots per inch, of the display device associated with this **Graphics** object. |
| [GetDpiY](#GetDpiY) | Gets the vertical resolution, in dots per inch, of the display device associated with this **Graphics** object. |
| [GetHalftonePalette](#GetHalftonePalette) | Gets a Windows halftone palette. |
| [GetHDC](#GetHDC) | Gets a handle to the device context associated with this **Graphics** object. |
| [GetInterpolationMode](#GetInterpolationMode) | Gets the interpolation mode currently set for this **Graphics** object. |
| [GetNearestColor](#GetNearestColor) | Gets the nearest color to the color that is passed in. |
| [GetPageScale](#GetPageScale) | Gets the scaling factor currently set for the page transformation of this **Graphics** object. |
| [GetPageUnit](#GetPageUnit) | Gets the unit of measure currently set for this **Graphics** object. |
| [GetPixelOffsetMode](#GetPixelOffsetMode) | Gets the pixel offset mode currently set for this **Graphics** object. |
| [GetRenderingOrigin](#GetRenderingOrigin) | Gets the rendering origin currently set for this **Graphics** object. |
| [GetSmoothingMode](#GetSmoothingMode) | Determines whether smoothing (antialiasing) is applied to the **Graphics** object. |
| [GetTextContrast](#GetTextContrast) | Gets the contrast value currently set for this **Graphics** object. |
| [GetTextRenderingHint](#GetTextRenderingHint) | Returns the text rendering mode currently set for this **Graphics** object. |
| [GetTransform](#GetTransform) | Gets the world transformation matrix of this **Graphics** object. |
| [GetVisibleClipBounds](#GetVisibleClipBounds) | Gets a rectangle that encloses the visible clipping region of this **Graphics** object. |
| [IntersectClip](#IntersectClip) | Updates the clipping region of this **Graphics** object to the portion of the specified rectangle that intersects with the current clipping region of this **Graphics** object. |
| [IsClipEmpty](#IsClipEmpty) | Determines whether the clipping region of this **Graphics** object is empty. |
| [IsVisible](#IsVisible) | Determines whether the specified point is inside the visible clipping region of this **Graphics** object. |
| [IsVisibleClipEmpty](#IsVisibleClipEmpty) | Determines whether the visible clipping region of this **Graphics** object is empty. |
| [MeasureCharacterRanges](#MeasureCharacterRanges) | Gets a set of regions each of which bounds a range of character positions within a string. |
| [MeasureDriverString](#MeasureDriverString) | Measures the bounding box for the specified characters and their corresponding positions. |
| [MeasureString](#MeasureString) | Measures the extent of the string in the specified font, format, and layout rectangle. |
| [MultiplyTransform](#MultiplyTransform) | Updates this **Graphics** object's world transformation matrix with the product of itself and another matrix. |
| [ReleaseHDC](#ReleaseHDC) | Releases a device context handle obtained by a previous call to the **GetHDC** method of this **Graphics** object. |
| [ResetClip](#ResetClip) | Sets the clipping region of this **Graphics** object to an infinite region. |
| [ResetTransform](#ResetTransform) | Sets the world transformation matrix of this Graphics object to the identity matrix. |
| [Restore](#Restore) | Sets the state of this **Graphics** object to the state stored by a previous call to the **Save** method of this Graphics object. |
| [RotateTransform](#RotateTransform) | Updates the world transformation matrix of this **Graphics** object with the product of itself and a rotation matrix. |
| [Save](#Save) | Saves the current state (transformations, clipping region, and quality settings) of this **Graphics** object. |
| [ScaleTransform](#ScaleTransform) | Updates this **Graphics** object's world transformation matrix with the product of itself and a scaling matrix. |
| [SetClip](#SetClip) | Updates the clipping region of this **Graphics** object to a region that is the combination of itself and the clipping region of another **Graphics** object. |
| [SetCompositingMode](#SetCompositingMode) | Sets the compositing mode of this **Graphics** object. |
| [SetCompositingQuality](#SetCompositingQuality) | Sets the compositing quality of this **Graphics** object. |
| [SetInterpolationMode](#SetInterpolationMode) | Sets the interpolation mode of this **Graphics** object. |
| [SetPageScale](#SetPageScale) | Sets the scaling factor for the page transformation of this **Graphics** object. |
| [SetPageUnit](#SetPageUnit) | Sets the unit of measure for this **Graphics** object. |
| [SetPixelOffsetMode](#SetPixelOffsetMode) | Sets the pixel offset mode of this **Graphics** object. |
| [SetRenderingOrigin](#SetRenderingOrigin) | Sets the rendering origin of this **Graphics** object. |
| [SetSmoothingMode](#SetSmoothingMode) | Sets the rendering quality of the **Graphics** object. |
| [SetTextContrast](#SetTextContrast) | Sets the contrast value of this **Graphics** object. |
| [SetTextRenderingHint](#SetTextRenderingHint) | Sets the text rendering mode of this **Graphics** object. |
| [SetTransform](#SetTransform) | Sets the world transformation of this **Graphics** object. |
| [TransformPoints](#TransformPoints) | Converts an array of points from one coordinate space to another. |
| [TranslateClip](#TranslateClip) | Translates the clipping region of this **Graphics** object. |
| [TranslateTransform](#TranslateTransform) | Updates this **Graphics** object's world transformation matrix with the product of itself and a translation matrix. |

# CGpGraphicsPath Class

The **CGpGraphicsPath** allows the creation of **GraphicPath** objects. A **GraphicsPath** object stores a sequence of lines, curves, and shapes. You can draw the entire sequence by calling the **DrawPath** method of a **Graphics** object. You can partition the sequence of lines, curves, and shapes into figures, and with the help of a **GraphicsPathIterator** object, you can draw selected figures. You can also place markers in the sequence, so that you can draw selected portions of the path.

**Inherits from**: CGpBase.<br>
**Imclude file**: CGpPath.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorsGraphicsPath) | Creates a **Graphics** object. |
| [AddArc](#AddArc) | Adds an elliptical arc to the current figure of this path. |
| [AddBezier](#AddBezier) | Adds a Bézier spline to the current figure of this path. |
| [AddBeziers](#AddBeziers) | Adds a sequence of connected Bézier splines to the current figure of this path. |
| [AddClosedCurve](#AddClosedCurve) | Adds a closed cardinal spline to this path. |
| [AddCurve](#AddCurve) | Adds a cardinal spline to this path. |
| [AddEllipse](#AddEllipse) | Adds an ellipse to this path. |
| [AddLine](#AddLine) | Adds a line to the current figure of this path. |
| [AddLines](#AddLines) | Adds a sequence of connected lines to the current figure of this path. |
| [AddPath](#AddPath) | Adds a path to this path. |
| [AddPie](#AddPie) | Adds a pie to this path. An arc is a portion of an ellipse, and a pie is a portion of the area enclosed by an ellipse. |
| [AddPolygon](#AddPolygon) | Adds a polygon to this path. |
| [AddRectangle](#AddRectangle) | Adds a rectangle to this path. |
| [AddRectangles](#AddRectangles) | Adds a sequence of rectangles to this path. |
| [AddString](#AddString) | Adds the outline of a string to this path. |
| [ClearMarkers](#ClearMarkers) | Clears the markers from this path. |
| [Clone](#Clone) | Copies the contents of the existing **GraphicsPath** object into a new **GraphicsPath** object. |
| [CloseAllFigures](#CloseAllFigures) | Clears the markers from this path. |
| [CloseFigure](#CloseFigure) | Clears the current figure of this path. |
| [Flatten](#Flatten) | Applies a transformation to this path and converts each curve in the path to a sequence of connected lines. |
| [GetBounds](#GetBounds) | Gets a bounding rectangle for this path. |
| [GetFillMode](#GetFillMode) | Gets the fill mode of this path. |
| [GetLastPoint](#GetLastPoint) | Gets the ending point of the last figure in this path. |
| [GetPathData](#GetPathData) | Gets an array of points and an array of point types from this path. |
| [GetPathPoints](#GetPathPoints) | Gets this path's array of points. |
| [GetPathTypes](#GetPathTypes) | Gets this path's array of point types. |
| [GetPointCount](#GetPointCount) | Gets the number of points in this path's array of data points. |
| [IsOutlineVisible](#IsOutlineVisible) | Determines whether a specified point touches the outline of this path when the path is drawn by a specified Graphics object and a specified pen. |
| [IsVisible](#IsVisible) | Determines whether a specified point lies in the area that is filled when this path is filled by a specified **Graphics** object. |
| [Outline](#Outline) | Transforms and flattens this path, and then converts this path's data points so that they represent only the outline of the path. |
| [Reset](#Reset) | Empties the path and sets the fill mode to **FillModeAlternate**. |
| [Reverse](#Reverse) | Reverses the order of the points that define this path's lines and curves. |
| [SetFillMode](#SetFillMode) | Sets the fill mode of this path. |
| [SetMarker](#SetMarker) | Designates the last point in this path as a marker point. |
| [StartFigure](#StartFigure) | Starts a new figure without closing the current figure. |
| [Transform](#Transform) | Multiplies each of this path's data points by a specified matrix. |
| [Warp](#Warp) | Applies a warp transformation to this path. |
| [Widen](#Widen) | Replaces this path with curves that enclose the area that is filled when this path is drawn by a specified pen. |

# CGpGraphicsPathIterator Class

The **CGpGraphicsPathIterator** class provides methods for isolating selected subsets of the path stored in a **GraphicsPath** object. A path consists of one or more figures. You can use a **GraphicsPathIterator** to isolate one or more of those figures. A path can also have markers that divide the path into sections. You can use a **GraphicsPathIterator** object to isolate one or more of those sections.

**Inherits from**: CGpBase.<br>
**Imclude file**: CGpPath.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#ConstructorGraphicsPathIterator) | Creates a new **GraphicsPathIterator** object and associates it with a **GraphicsPath** object. |
| [CopyData](#CopyData) | Copies a subset of the path's data points to a **GpPointF** array and copies a subset of the path's point types to a byte array. |
| [Enumerate](#Enumerate) | Copies the path's data points to a **GpPointF** array and copies the path's point types to a byte array. |
| [GetCount](#GetCount) | Returns the number of data points in the path. |
| [GetSubpathCount](#GetSubpathCount) | Returns the number of subpaths (also called figures) in the path. |
| [HasCurve](#HasCurve) | Determines whether the path has any curves. |
| [NextMarker](#NextMarker) | Gets the starting index and the ending index of the next marker-delimited section in this iterator's associated path. |
| [NextPathType](#NextPathType) | Gets the starting index and the ending index of the next group of data points that all have the same type. |
| [NextSubpath](#NextSubpath) | Gets the starting index and the ending index of the next subpath (figure) in this iterator's associated path. |
| [Rewind](#Rewind) | Rewinds this iterator to the beginning of its associated path. |

# <a name="ConstructorsGraphics"></a>Constructors (CGpGraphics)

Creates a **Graphics** object that is associated with a specified device context. When you use this method to create a **Graphics** object, make sure that the **Graphics** object is deleted before the device context is released.

```
CONSTRUCTOR CGpGraphics (BYVAL hdc AS HDC)
```
Creates a **Graphics** object that is associated with a specified device context and a specified device. When you use this constructor to create a **Graphics** object, make sure that the Graphics object is deleted or goes out of scope before the device context is released.

```
CONSTRUCTOR CGpGraphics (BYVAL hdc AS HDC, BYVAL hDevice AS HANDLE)
```

Creates a **Graphics** object that is associated with a specified window.

```
CONSTRUCTOR CGpGraphics (BYVAL hwnd AS HWND, BYVAL icm AS BOOLEAN = FALSE)
```

Creates a **Graphics** object that is associated with an **Image** object. This constructor fails if the **Image** object is based on a metafile that was opened for reading. The Image(file) and Metafile(file) constructors open a metafile for reading. To open a metafile for recording, use a **Metafile** constructor that receives a device context handle.

```
CONSTRUCTOR CGpGraphics (BYVAL pImage AS CGpImage PTR)
```

This constructor also fails if the image uses one of the following pixel formats:

PixelFormatUndefined<br>
PixelFormatDontCare<br>
PixelFormat1bppIndexed<br>
PixelFormat4bppIndexed<br>
PixelFormat8bppIndexed<br>
PixelFormat16bppGrayScale<br>
PixelFormat16bppARGB1555

```
CONSTRUCTOR CGpGraphics (BYVAL pImage AS CGpImage PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hdc* | Handle to the device context that will be associated with the new **Graphics** object. |
| *hDevice* | Handle to a device that will be associated with the new **Graphics** object. |
| *hwnd* | Handle to a window that will be associated with the new **Graphics** object. |
| *icm* | Optional. Boolean value that specifies whether the new **Graphics** object applies color adjustment according to the ICC profile associated with the display device. TRUE specifies that color adjustment is applied, and FALSE specifies that color adjustment is not applied. The default value is FALSE. |
| *pImage* | Pointer to an **Image** object that will be associated with the new **Graphics** object. |


# <a name="AddMetafileComment"></a>AddMetafileComment (CGpGraphics)

Adds a text comment to an existing metafile.

```
FUNCTION AddMetafileComment (BYVAL pdata AS BYTE PTR, BYVAL sizeData AS UINT) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pData* | Pointer to a buffer that contains the comment. |
| *sizeData* | Integer that specifies the number of bytes in the value of the *data* parameter.  |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.


# <a name="BeginContainer"></a>BeginContainer (CGpGraphics)

Begins a new graphics container.

```
FUNCTION BeginContainer () AS GraphicsContainer
FUNCTION BeginContainer (BYVAL destrect AS GpRectF PTR, BYVAL srcrect AS GpRectF PTR, _
   BYVAL nUnit AS GpUnit) AS GraphicsContainer
FUNCTION BeginContainer (BYVAL destrect AS GpRect PTR, BYVAL srcrect AS GpRect PTR, _
   BYVAL nUnit AS GpUnit) AS GraphicsContainer
```

| Parameter  | Description |
| ---------- | ----------- |
| *destrect* | Reference to a rectangle that, together with *srcrect*, specifies a transformation for the container. |
| *srcrect* | Reference to a rectangle that, together with *dstrect*, specifies a transformation for the container. |

#### Return value

This method returns a value that identifies the container.

#### Remarks

Use this method to create nested graphics containers. Graphics containers are used to retain graphics state, such as transformations, clipping regions, and various rendering properties.

The **BeginContainer** method returns a value of type **GraphicsContainer**. When you have finished using a container, pass that value to the **EndContainer** method. The **GraphicsContainer** data type is defined in Gdiplusenums.bi.

When you call the **BeginContainer** method of a **Graphics** object, an information block that holds the state of the **Graphics** object is put on a stack. The **BeginContainer** method returns a value that identifies that information block. When you pass the identifying value to the EndContainer method, the information block is removed from the stack and is used to restore the **Graphics** object to the state it was in at the time of the **BeginContainer** call.

Containers can be nested; that is, you can call the **BeginContainer** method several times before you call the **EndContainer** method. Each time you call the **BeginContainer** method, an information block is put on the stack, and you receive an identifier for the information block. When you pass one of those identifiers to the **EndContainer** method, the **Graphics** object is returned to the state it was in at the time of the **BeginContainer** call that returned that particular identifier. The information block placed on the stack by that **BeginContainer** call is removed from the stack, and all information blocks placed on that stack after that **BeginContainer** call are also removed.

Calls to the **Save** method place information blocks on the same stack as calls to the **BeginContainer** method. Just as an **EndContainer** call is paired with a **BeginContainer** call, a **Restore** call is paired with a **Save** call.

**Caution**: When you call **EndContainer**, all information blocks placed on the stack (by **Save** or by **BeginContainer**) after the corresponding call to **BeginContainer** are removed from the stack. Likewise, when you call **Restore**, all information blocks placed on the stack (by Save or by **BeginContainer**) after the corresponding call to Save are removed from the stack.

For more information about graphics containers, see [Nested Graphics Containers](https://docs.microsoft.com/en-us/windows/desktop/gdiplus/-gdiplus-nested-graphics-containers-use)


# <a name="Clear"></a>Clear (CGpGraphics)

Clears a **Graphics** object to a specified color.

```
FUNCTION Clear (BYVAL colour AS ARGB) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *Colour* | The color to paint the background. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.


# <a name="DrawArc"></a>DrawArc (CGpGraphics)

Draws an arc. The arc is part of an ellipse.

```
FUNCTION DrawArc (BYVAL pPen AS CGpPen PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, _
   BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE, BYVAL startAngle AS SINGLE, _
   BYVAL sweepAngle AS SINGLE) AS LONG
FUNCTION DrawArc (BYVAL pPen AS CGpPen PTR, BYVAL x AS INT_, BYVAL y AS INT_, _
   BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL startAngle AS SINGLE, _
   BYVAL sweepAngle AS SINGLE) AS LONG
FUNCTION DrawArc (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRectF PTR, _
   BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS LONG
FUNCTION DrawArc (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRect PTR, _
   BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pPen* | Pointer to a pen that is used to draw the arc. |
| *x* | The x-coordinate of the upper-left corner of the bounding rectangle for the ellipse that contains the arc. |
| *y* | The y-coordinate of the upper-left corner of the bounding rectangle for the ellipse that contains the arc. |
| *nWidth* | The width of the ellipse that contains the arc. |
| *nHeight* | The height of the ellipse that contains the arc. |
| *startAngle* | The angle between the x-axis and the starting point of the arc.  |
| *sweepAngle* | The angle between the starting and ending points of the arc. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example draws a 90-degree arc.
' ========================================================================================
SUB Example_DrawArc (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Draw the arc
   DIM redPen AS CGpPen = CGpPen(GDIP_ARGB(255, 255, 0, 0), 3)
   graphics.DrawArc(@redPen, 0, 0, 200, 100, 0.0, 90.0)

END SUB
' ========================================================================================
```

# <a name="DrawBezier"></a>DrawBezier (CGpGraphics)

Draws a Bézier spline.

```
FUNCTION DrawBezier (BYVAL pPen AS CGpPen PTR, BYVAL x1 AS SINGLE, BYVAL y1 AS SINGLE, _
   BYVAL x2 AS SINGLE, BYVAL y2 AS SINGLE, BYVAL x3 AS SINGLE, BYVAL y3 AS SINGLE, _
   BYVAL x4 AS SINGLE, BYVAL y4 AS SINGLE) AS GpStatus
FUNCTION DrawBezier (BYVAL pPen AS CGpPen PTR, BYVAL x1 AS INT_, BYVAL y1 AS INT_, _
   BYVAL x2 AS INT_, BYVAL y2 AS INT_, BYVAL x3 AS INT_, BYVAL y3 AS INT_, _
   BYVAL x4 AS INT_, BYVAL y4 AS INT_) AS GpStatus
FUNCTION DrawBezier (BYVAL pPen AS CGpPen PTR, BYVAL pt1 AS GpPointF) AS GpStatus
FUNCTION DrawBezier (BYVAL pPen AS CGpPen PTR, BYVAL pt1 AS GpPoint) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pPen* | Pointer to a pen that is used to draw the Bézier spline. |
| *x1* | The x-coordinate of the starting point of the Bézier spline. |
| *y1* | The y-coordinate of the starting point of the Bézier spline. |
| *x2* | The x-coordinate of the first control point of the Bézier spline. |
| *y2* | The y-coordinate of the first control point of the Bézier spline. |
| *x3* | The x-coordinate of the second control point of the Bézier spline. |
| *y3* | The y-coordinate of the second control point of the Bézier spline. |
| *x4* | The x-coordinate of the ending point of the Bézier spline. |
| *y4* | The y-coordinate of the ending point of the Bézier spline. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example draws a Bézier curve.
' ========================================================================================
SUB Example_DrawBezier (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Draw the curve.
   DIM greenPen AS CGpPen = GDIP_ARGB(255, 0, 255, 0)
   graphics.DrawBezier(@greenPen, 100.0, 100.0, 200.0, 10.0, 350.0, 50.0, 500.0, 100.0)

   ' // Draw the end points and control points.
   DIM redBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   DIM blueBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 255)
   graphics.FillEllipse(@redBrush, 100 - 5, 100 - 5, 10, 10)
   graphics.FillEllipse(@redBrush, 500 - 5, 100 - 5, 10, 10)
   graphics.FillEllipse(@blueBrush, 200 - 5, 10 - 5, 10, 10)
   graphics.FillEllipse(@blueBrush, 350 - 5, 50 - 5, 10, 10)

END SUB
' ========================================================================================
```


# <a name="DrawBeziers"></a>DrawBeziers (CGpGraphics)

Draws a sequence of connected Bézier splines.

```
FUNCTION DrawBeziers (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pPen* | Pointer to a pen that is used to draw the Bézier spline. |
| *pts* | Pointer to an array of GpPointF objects that specify the starting, ending, and control points of the Bézier splines. |
| *count* | Integer that specifies the number of elements in the points array. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A Bézier spline does not pass through its control points. The control points act as magnets, pulling the curve in certain directions to influence the way a Bézier spline bends. Each Bézier spline requires a starting point and an ending point. Each ending point is the starting point for the next Bézier spline.

#### Example

```
' ========================================================================================
' The following example draws a pair of Bézier curves.
' ========================================================================================
SUB Example_DrawBeziers (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Define a Pen object and an array of PointF objects.
   DIM greenPen AS CGpPen = GDIP_ARGB(255, 0, 255, 0)
   DIM startPoint AS GpPointF : startPoint.x = 100.0 : startPoint.y = 100.0
   DIM ctrlPoint1 AS GpPointF : ctrlPoint1.x = 200.0 : ctrlPoint1.y = 50.0
   DIM ctrlPoint2 AS GpPointF : ctrlPoint2.x = 400.0 : ctrlPoint2.y = 10.0
   DIM endPoint1  AS GpPointF : endPoint1.x  = 500.0 : endPoint1.y  = 100.0
   DIM ctrlPoint3 AS GpPointF : ctrlPoint3.x = 600.0 : ctrlPoint3.y = 200.0
   DIM ctrlPoint4 AS GpPointF : ctrlPoint4.x = 700.0 : ctrlPoint4.y = 400.0
   DIM endPoint2  AS GpPointF : endPoint2.x  = 500.0 : endPoint2.y  = 500.0

   DIM curvePoints(6) AS GpPointF
   curvePoints(0) = startPoint
   curvePoints(1) = ctrlPoint1
   curvePoints(2) = ctrlPoint2
   curvePoints(3) = endPoint1
   curvePoints(4) = ctrlPoint3
   curvePoints(5) = ctrlPoint4
   curvePoints(6) = endPoint2

   ' // Draw the Bezier curves.
   graphics.DrawBeziers(@greenPen, @curvePoints(0), 7)

   ' // Draw the control and end points.
   DIM redBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   graphics.FillEllipse(@redBrush, 100 - 5, 100 - 5, 10, 10)
   graphics.FillEllipse(@redBrush, 500 - 5, 100 - 5, 10, 10)
   graphics.FillEllipse(@redBrush, 500 - 5, 500 - 5, 10, 10)
   DIM blueBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 255)
   graphics.FillEllipse(@blueBrush, 200 - 5, 50 - 5, 10, 10)
   graphics.FillEllipse(@blueBrush, 400 - 5, 10 - 5, 10, 10)
   graphics.FillEllipse(@blueBrush, 600 - 5, 200 - 5, 10, 10)
   graphics.FillEllipse(@blueBrush, 700 - 5, 400 - 5, 10, 10)

END SUB
' ========================================================================================
```

# <a name="DrawCachedBitmap"></a>DrawCachedBitmap (CGpGraphics)

Draws the image stored in a **CachedBitmap** object.

```
FUNCTION DrawCachedBitmap (BYVAL pCachedBitmap AS CGpCachedBitmap PTR, _
   BYVAL x AS LONG, BYVAL y AS LONG) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pCachedBitmap* | Pointer to a **CachedBitmap** object that contains the image to be drawn. |
| *x* | The x-coordinate of the upper-left corner of the image. |
| *y* | The y-coordinate of the upper-left corner of the image. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A **CachedBitmap** object stores an image in a format that is optimized for a particular display screen. You cannot draw a cached bitmap to a printer or to a metafile.

Cached bitmaps will not work with any transformations other than translation.

When you construct a **CachedBitmap** object, you must pass the address of a **Graphics** object to the constructor. If the screen associated with that Graphics object has its bit depth changed after the cached bitmap is constructed, then the **DrawCachedBitmap** method will fail, and you should reconstruct the cached bitmap. Alternatively, you can hook the display change notification message and reconstruct the cached bitmap at that time.


# <a name="DrawClosedCurve"></a>DrawClosedCurve (CGpGraphics)

Draws a closed cardinal spline.

```
FUNCTION DrawClosedCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, BYVAL count AS INT_) AS GpStatus
FUNCTION DrawClosedCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPoint PTR, BYVAL count AS INT_) AS GpStatus
FUNCTION DrawClosedCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, BYVAL count AS INT_, _
   BYVAL tension AS SINGLE) AS GpStatus
FUNCTION DrawClosedCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPoint PTR, BYVAL count AS INT_, _
   BYVAL tension AS SINGLE) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pPen* | Pointer to a pen that is used to draw the closed cardinal spline. |
| *pts* | Pointer to an array of **GpPointF** objects that specify the coordinates of the closed cardinal spline. The array of **GpPointF** objects must contain a minimum of three elements. |
| *count* | Integer that specifies the number of elements in the points array. |
| *tension* | Single precision number that specifies how tightly the curve bends through the coordinates of the closed cardinal spline. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

Each ending point is the starting point for the next cardinal spline. In a closed cardinal spline, the curve continues through the last point in the points array and connects with the first point in the array.

#### Example

```
' ========================================================================================
' The following example draws a closed cardinal spline.
' ========================================================================================
SUB Example_DrawClosedCurve (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Define a Pen object and an array of PointF objects.
   DIM greenPen AS CGpPen = GDIP_ARGB(255, 0, 255, 0)
   DIM point1 AS GpPointF : point1.x = 100.0 : point1.y = 100.0
   DIM point2 AS GpPointF : point2.x = 200.0 : point2.y = 50.0
   DIM point3 AS GpPointF : point3.x = 400.0 : point3.y = 10.0
   DIM point4 AS GpPointF : point4.x = 500.0 : point4.y = 100.0
   DIM point5 AS GpPointF : point5.x = 600.0 : point5.y = 200.0
   DIM point6 AS GpPointF : point6.x = 700.0 : point6.y = 400.0
   DIM point7 AS GpPointF : point7.x = 500.0 : point7.y = 500.0

   DIM curvePoints(6) AS GpPointF
   curvePoints(0) = point1
   curvePoints(1) = point2
   curvePoints(2) = point3
   curvePoints(3) = point4
   curvePoints(4) = point5
   curvePoints(5) = point6
   curvePoints(6) = point7

   ' // Draw the closed curve.
   graphics.DrawClosedCurve(@greenPen, @curvePoints(0), 7, 1.0)

   ' // Draw the points in the curve.
   DIM redBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   graphics.FillEllipse(@redBrush, 95, 95, 10, 10)
   graphics.FillEllipse(@redBrush, 495, 95, 10, 10)
   graphics.FillEllipse(@redBrush, 495, 495, 10, 10)
   graphics.FillEllipse(@redBrush, 195, 45, 10, 10)
   graphics.FillEllipse(@redBrush, 395, 5, 10, 10)
   graphics.FillEllipse(@redBrush, 595, 195, 10, 10)
   graphics.FillEllipse(@redBrush, 695, 395, 10, 10)

END SUB
' ========================================================================================
```


# <a name="DrawCurve"></a>DrawCurve (CGpGraphics)

Draws a cardinal spline.

```
FUNCTION DrawCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, BYVAL count AS INT_) AS GpStatus
FUNCTION DrawCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPoint PTR, BYVAL count AS INT_) AS GpStatus
FUNCTION DrawCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, _
   BYVAL count AS INT_, BYVAL tension AS SINGLE) AS GpStatus
FUNCTION DrawCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPoint PTR, _
   BYVAL count AS INT_, BYVAL tension AS SINGLE) AS GpStatus
FUNCTION DrawCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, _
   BYVAL count AS INT_, BYVAL tension AS SINGLE) AS GpStatus
FUNCTION DrawCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, BYVAL count AS INT_, _
   BYVAL offset AS INT_, BYVAL numberOfSegments AS INT_, BYVAL tension AS SINGLE) AS GpStatus
FUNCTION DrawCurve (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPoint PTR, BYVAL count AS INT_, _
   BYVAL offset AS INT_, BYVAL numberOfSegments AS INT_, BYVAL tension AS SINGLE) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pPen* | Pointer to a pen that is used to draw the cardinal spline. |
| *pts* | Pointer to an array of GpPointF objects that specify the coordinates that the cardinal spline passes through. |
| *count* | Integer that specifies the number of elements in the points array. |
| *offset* | Integer that specifies the element in the points array that specifies the point at which the cardinal spline begins. |
| *numberOfSegments* | Integer that specifies the number of segments in the cardinal spline. |
| *tension* | Single precision number that specifies how tightly the curve bends through the coordinates of the cardinal spline. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A segment is defined as a curve that connects two consecutive points in the cardinal spline. The ending point of each segment is the starting point for the next.

#### Example

```
' ========================================================================================
' The following example draws a cardinal spline.
' ========================================================================================
SUB Example_DrawCurve (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM greenPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 255, 0), 3)

   DIM point1 AS GpPointF : point1.x = 100.0 : point1.y = 100.0
   DIM point2 AS GpPointF : point2.x = 200.0 : point2.y = 50.0
   DIM point3 AS GpPointF : point3.x = 400.0 : point3.y = 10.0
   DIM point4 AS GpPointF : point4.x = 500.0 : point4.y = 100.0

   DIM curvePoints(3) AS GpPointF
   curvePoints(0) = point1
   curvePoints(1) = point2
   curvePoints(2) = point3
   curvePoints(3) = point4

   ' // Draw the curve.
   graphics.DrawCurve(@greenPen, @curvePoints(0), 4)

   ' // Draw the points in the curve.
   DIM redBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   graphics.FillEllipse(@redBrush, 95, 95, 10, 10)
   graphics.FillEllipse(@redBrush, 195, 45, 10, 10)
   graphics.FillEllipse(@redBrush, 395, 5, 10, 10)
   graphics.FillEllipse(@redBrush, 495, 95, 10, 10)

END SUB
' ========================================================================================
```


# <a name="DrawDriverString"></a>DrawDriverString (CGpGraphics)

Draws characters at the specified positions. The method gives the client complete control over the appearance of text. The method assumes that the client has already set up the format and layout to be applied.

```
FUNCTION DrawDriverString (BYVAL pText AS UINT16 PTR, BYVAL length AS INT_, BYVAL pFont AS CGpFont PTR, _
   BYVAL pBrush AS CGpBrush PTR, BYVAL positions AS ANY PTR, BYVAL flags AS INT_, _
   BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pText* | Pointer to an array of 16-bit values. If the **DriverStringOptionsCmapLookup** flag is set, each value specifies a Unicode character to be displayed. Otherwise, each value specifies an index to a font glyph that defines a character to be displayed. |
| *length* | Integer that specifies the number of values in the text array. The length parameter can be set to –1 if the string is null terminated. |
| *pFont* | Pointer to a **Font** object that specifies the family name, size, and style of the font that is to be applied to the string. |
| *pBrush* | Pointer to a **Brush** object that is used to fill the string. |
| *positions* | If the **DriverStringOptionsRealizedAdvance** flag is set, positions is a pointer to a **GpPointF** structure that specifies the position of the first glyph. Otherwise, positions is an array of **GpPointF** structures, each of which specifies the origin of an individual glyph. |
| *flags* | Integer that specifies the options for the appearance of the string. This value must be an element of the **DriverStringOptions** enumeration or the result of a bitwise OR applied to two or more of these elements. |
| *pMatrix* | Pointer to a **Matrix** object that specifies the transformation matrix to apply to each value in the text array. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A segment is defined as a curve that connects two consecutive points in the cardinal spline. The ending point of each segment is the starting point for the next. The numberOfSegments parameter must not be greater than the count parameter minus the offset parameter plus one.


# <a name="DrawEllipse"></a>DrawEllipse (CGpGraphics)

Draws an ellipse.

```
FUNCTION DrawEllipse (BYVAL pPen AS CGpPen PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, _
   BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS GpStatus
FUNCTION DrawEllipse (BYVAL pPen AS CGpPen PTR, BYVAL x AS INT_, BYVAL y AS INT_, _
   BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS GpStatus
FUNCTION DrawEllipse (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRectF) AS GpStatus
FUNCTION DrawEllipse (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRect) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pPen* | Pointer to a pen that is used to draw the ellipse. |
| *x* | The x-coordinate of the upper-left corner of the rectangle that bounds the ellipse. |
| *y* | The y-coordinate of the upper-left corner of the rectangle that bounds the ellipse. |
| *nWidth* | The width of the rectangle that bounds the ellipse. |
| *nHeight* | The height of the rectangle that bounds the ellipse. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example draws an ellipse.
' ========================================================================================
SUB Example_DrawEllipse (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Draw the ellipse
   DIM bluePen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 3)
   graphics.DrawEllipse(@bluePen, 100.0, 70.0, 200.0, 100.0)

END SUB
' ========================================================================================
```


# <a name="DrawImage"></a>DrawImage (CGpGraphics)

Draws an image.

```
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL rc AS GpRectF PTR) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL rc AS GpRect PTR) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, _
   BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL x AS INT_, BYVAL y AS INT_ _
   BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL pt AS GpPointF PTR) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL pt AS GpPoint PTR) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL x AS INT_, BYVAL y AS INT_) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destPoints AS GpPointF, BYVAL count AS INT_) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destPoints AS GpPoint, BYVAL count AS INT_) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, _
   BYVAL srcx AS SINGLE, BYVAL srcy AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE, _
   BYVAL srcUnit AS GpUnit) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL srcx AS INT_, _
   BYVAL srcy AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL srcUnit AS GpUnit) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destRect AS GpRectF PTR, _
   BYVAL sourceRect AS GpRectF PTR, BYVAL srcUnit AS GpUnit, _
   BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destRect AS GpRect PTR, _
   BYVAL sourceRect AS GpRect PTR, BYVAL srcUnit AS GpUnit, _
    BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destx AS SINGLE, BYVAL desty AS SINGLE, _
   BYVAL destWidth AS SINGLE, BYVAL destHeight AS SINGLE, BYVAL sourcex AS SINGLE, _
   BYVAL sourcey AS SINGLE, BYVAL sourceWidth AS SINGLE, BYVAL sourceHeight AS SINGLE, _
   BYVAL srcUnit AS GpUnit, BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destx AS INT_, BYVAL desty AS INT_, _
   BYVAL destWidth AS INT_, BYVAL destHeight AS INT_, BYVAL sourcex AS INT_, BYVAL sourcey AS INT_, _
   BYVAL sourceWidth AS INT_, BYVAL sourceHeight AS INT_, BYVAL srcUnit AS GpUnit, _
   BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destRect AS GpRectF PTR, _
   BYVAL srcx AS SINGLE, BYVAL srcy AS SINGLE, BYVAL srcwidth AS SINGLE, _
   BYVAL srcheight AS SINGLE, BYVAL srcUnit AS GpUnit, _
   BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL, _
   BYVAL pcallback AS DrawImageAbort = NULL, BYVAL pcallbackData AS ANY PTR = NULL) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destRect AS GpRect PTR, _
   BYVAL srcx AS INT_, BYVAL srcy AS INT_, BYVAL srcwidth AS INT_, BYVAL srcheight AS INT_, _
   BYVAL srcUnit AS GpUnit, BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL, _
   BYVAL pcallback AS DrawImageAbort = NULL, BYVAL pcallbackData AS ANY PTR = NULL) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destx AS SINGLE, BYVAL desty AS SINGLE, _
   BYVAL destWidth AS SINGLE, BYVAL destHeigh AS SINGLE, BYVAL srcx AS SINGLE, _
   BYVAL srcy AS SINGLE, BYVAL srcwidth AS SINGLE, BYVAL srcheight AS SINGLE, _
   BYVAL srcUnit AS GpUnit, BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL, _
   BYVAL pcallback AS DrawImageAbort = NULL, BYVAL pcallbackData AS ANY PTR = NULL) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destx AS INT_, BYVAL desty AS INT_, _
   BYVAL destWidth AS INT_, BYVAL destHeigh AS INT_, BYVAL srcx AS INT_, BYVAL srcy AS INT_, _
   BYVAL srcwidth AS INT_, BYVAL srcheight AS INT_, BYVAL srcUnit AS GpUnit, _
   BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL, _
   BYVAL pcallback AS DrawImageAbort = NULL, BYVAL pcallbackData AS ANY PTR = NULL) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destPoints AS GpPointF PTR, _
   BYVAL nCount AS INT_, BYVAL srcx AS SINGLE, BYVAL srcy AS SINGLE, BYVAL srcwidth AS SINGLE, _
   BYVAL srcheight AS SINGLE, BYVAL srcUnit AS GpUnit, _
   BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL, _
   BYVAL pcallback AS DrawImageAbort = NULL, BYVAL pcallbackData AS ANY PTR = NULL) AS GpStatus
FUNCTION DrawImage (BYVAL pImage AS CGpImage PTR, BYVAL destPoints AS GpPoint PTR, _
   BYVAL nCount AS INT_, BYVAL srcx AS INT_, BYVAL srcy AS INT_, BYVAL srcwidth AS INT_, _
   BYVAL srcheight AS INT_, BYVAL srcUnit AS GpUnit, _
   BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL, _
   BYVAL pcallback AS DrawImageAbort = NULL, BYVAL pcallbackData AS ANY PTR = NULL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pImage* | Pointer to an **Image** object that specifies the source image. |
| *x* | The x-coordinate of the upper-left corner of the destination rectangle at which to draw the image. |
| *y* | The y-coordinate of the upper-left corner of the destination rectangle at which to draw the image. |
| *nWidth* | Optional. The width of the destination rectangle at which to draw the image. |
| *nHeight* | Optional. The height of the destination rectangle at which to draw the image. |
| *pt* | Reference to a **GpPointF** object that specifies the coordinates of the upper-left corner of the destination position at which to draw the image. |
| *srcx* | The x-coordinate of the upper-left corner of the portion of the source image to be drawn. |
| *srcy* | The y-coordinate of the upper-left corner of the portion of the source image to be drawn. |
| *srcwidth* | The width of the portion of the source image to be drawn. |
| *srcheight* | The height of the portion of the source image to be drawn. |
| *srcUnit* | Element of the Unit enumeration that specifies the unit of measure for the image. The default value is **UnitPixel**. |
| *destPoints* | Pointer to an array of **GpPointF** objects that specify the area, in a parallelogram, in which to draw the image. |
| *nCount* | Integer that specifies the number of elements in the *destPoints* array. |
| *pImageAttributes* | Pointer to an **ImageAttributes** object that specifies the color and size attributes of the image to be drawn. The default value is NULL. |
| *pCallback* | Callback method used to cancel the drawing in progress. The default value is NULL. |
| *pCallbackData* | Pointer to additional data used by the method specified by the *pCallback* parameter. The default value is NULL. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example draws an image.
' ========================================================================================
SUB Example_DrawImage (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set scaling
   graphics.SetPageUnit(UnitPixel)
   graphics.SetPageScale(rxRatio)

   ' // Create an Image object.
   DIM pImage AS CGpImage = "climber.jpg"

   ' // Draw the original source image.
   graphics.DrawImage(@pImage, 10, 10)

   ' // Draw the scaled image.
   graphics.DrawImage(@pImage, 200, 50, 150, 75)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' The following example draws a portion of an image. The portion of the source image to be
' drawn is scaled to fit a specified parallelogram.
' ========================================================================================
SUB Example_DrawImage (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set scaling
   graphics.SetPageUnit(UnitPixel)
   graphics.SetPageScale(rxRatio)

   ' // Create an Image object.
   DIM pImage AS CGpImage = "climber.jpg"

   ' // Draw the original source image.
   graphics.DrawImage(@pImage, 10, 10)

   ' // Draw the scaled image.
   graphics.DrawImage(@pImage, 200.0, 30.0, 70.0, 20.0, 100.0, 200.0, UnitPixel)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' The following example draws an image.
' ========================================================================================
SUB Example_DrawImage (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set scaling
   graphics.SetPageUnit(UnitPixel)
   graphics.SetPageScale(rxRatio)

   ' // Create an Image object.
   DIM pImage AS CGpImage = "climber.jpg"

   ' // Create an array of PointF objects that specify the destination of the image.
   DIM destPoints(0 TO 2) AS GpPointF
   destPoints(0).x =  30 : destPoints(0).y =  30
   destPoints(1).x = 250 : destPoints(1).y =  50
   destPoints(2).x = 175 : destPoints(2).y = 120

   ' // Draw the image.
   graphics.DrawImage(@pImage, @destPoints(0), 3)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' The following example draws the original source image and then draws a portion of the
' image in a specified parallelogram.
' ========================================================================================
SUB Example_DrawImageRectRect (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set scaling
   graphics.SetPageUnit(UnitPixel)
   graphics.SetPageScale(rxRatio)

   ' // Create an Image object.
   DIM pImage AS CGpImage = "pattern.png"

   ' // Draw the original source image.
   graphics.DrawImage(@pImage, 10, 10)

   ' // Define the portion of the image to draw.
   DIM srcX AS SINGLE = 70.0
   DIM srcY AS SINGLE = 20.0
   DIM srcWidth AS SINGLE = 100.0
   DIM srcHeight AS SINGLE = 100.0

   ' // Create an array of Point objects that specify the destination of the cropped image.
   DIM destPoints(0 TO 2) AS GpPointF
   destPoints(0).x = 230 : destPoints(0).y = 30
   destPoints(1).x = 350 : destPoints(1).y = 50
   destPoints(2).x = 275 : destPoints(2).y = 120

   ' Yet another mess of the FB GdiPlus declares.
'#ifdef __FB_64BIT__
'   DIM redToBlue AS ColorMap_
'   redToBlue.oldColor.value = GDIP_ARGB(255, 255, 0, 0)
'   redToBlue.newColor.value = GDIP_ARGB(255, 0, 0, 255)
'#else
'   DIM redToBlue AS ColorMap
'   redToBlue.from = GDIP_ARGB(255, 255, 0, 0)
'   redToBlue.to = GDIP_ARGB(255, 0, 0, 255)
'#endif

   ' // GDIP_COLORMAP is an union that solves the 32/64-bit incompatibility
   DIM redToBlue AS GDIP_COLORMAP = (GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255))

   ' // Create an ImageAttributes object that specifies a recoloring from red to blue.
   DIM remapAttributes AS CGpImageAttributes
   RemapAttributes.SetRemapTable(1, @redToBlue)

   ' // Draw the cropped image
   graphics.DrawImage(@pImage, @destPoints(0), 3, srcX, srcY, srcWidth, srcHeight, _
                     UnitPixel, @remapAttributes, NULL, NULL)


END SUB
' ========================================================================================
```


# <a name="DrawImageFX"></a>DrawImageFX (CGpGraphics)

Draws a portion of an image after applying a specified effect.

**Note**: Not yet implemented.

```
FUNCTION DrawImageFX (BYVAL pImage AS CGpImage PTR, BYREF sourceRect AS GpRectF, _
   BYVAL pMatrix AS CGpMatrix PTR, BYREF gpEffect AS CGpEffect, _
   BYREF gpImageAttributes AS CGpImageAttributes, BYVAL srcUnit AS GpUnit) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pImage* | Pointer to an **Image** object that specifies the image to be drawn. |
| *sourceRect* | Pointer to a **GpRectF** structure that specifies the portion of the image to be drawn. |
| *pMatrix* | Pointer to a **Matrix** object that specifies the parallelogram in which the image portion is rendered. The destination parallelogram is calculated by applying the affine transformation stored in the matrix to the source rectangle. |
| *pEffect* | Pointer to a instance of a descendant of the **CGpEffect** class. The descendant specifies an effect or adjustment (for example, a change in contrast) that is applied to the image before rendering. The image is not permanently altered by the effect. |
| *imageAttributes* | Pointer to an **ImageAttributes** object that specifies color adjustments to be applied when the image is rendered. Can be NULL. |
| *srcUnit* | Element of the **GpUnit** enumeration that specifies the unit of measure for the source rectangle. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.


# <a name="DrawLine"></a>DrawLine (CGpGraphics)

Draws a line that connects two points.

```
FUNCTION DrawLine (BYVAL pPen AS CGpPen PTR, BYVAL x1 AS SINGLE, BYVAL y1 AS SINGLE, _
   BYVAL x2 AS SINGLE, BYVAL y2 AS SINGLE) AS GpStatus
FUNCTION DrawLine (BYVAL pPen AS CGpPen PTR, BYVAL x1 AS INT_, BYVAL y1 AS INT_, _
   BYVAL x2 AS INT_, BYVAL y2 AS INT_) AS GpStatus
FUNCTION DrawLine (BYVAL pPen AS CGpPen PTR, BYVAL pt1 AS GpPointF PTR, BYVAL pt2 AS GpPointF PTR) AS GpStatus
FUNCTION DrawLine (BYVAL pPen AS CGpPen PTR, BYVAL pt1 AS GpPoint PTR, BYVAL pt2 AS GpPoint PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pPen* | Pointer to a pen that is used to draw the line. |
| *x1* | The x-coordinate of the starting point of the line. |
| *y1* | The y-coordinate of the starting point of the line. |
| *x2* | The x-coordinate of the ending point of the line. |
| *y2* | The y-coordinate of the ending point of the line. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example draws a line.
' ========================================================================================
SUB Example_DrawLine (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Draw the line
   DIM blackPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 0), 3)
   graphics.DrawLine(@blackPen, 100.0, 100.0, 500.0, 100.0)

END SUB
' ========================================================================================
```


# <a name="DrawLines"></a>DrawLines (CGpGraphics)

Draws a sequence of connected lines.

```
FUNCTION DrawLines (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, BYVAL count AS LONG) AS GpStatus
FUNCTION DrawLines (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPoint PTR, BYVAL count AS LONG) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pPen* | Pointer to a pen that is used to draw the lines. |
| *pts* | Pointer to an array of **GpPointF** structures that specify the starting and ending points of the lines. |
| *nCount* | Integer that specifies the number of elements in the points array. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example draws a sequence of connected lines.
' ========================================================================================
SUB Example_DrawLines (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Pen object
   DIM blackPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 0), 3)

   ' // Create an array of PointF objects that define the lines to draw
   DIM point1 AS GpPointF : point1.x =  10 : point1.y = 10
   DIM point2 AS GpPointF : point2.x =  10 : point2.y = 100
   DIM point3 AS GpPointF : point3.x = 200 : point3.y = 50
   DIM point4 AS GpPointF : point4.x = 250 : point4.y = 300

   DIM pts(0 TO 3) AS GpPointF
   pts(0) = point1
   pts(1) = point2
   pts(2) = point3
   pts(3) = point4

   ' // Draw the lines
   graphics.DrawLines(@blackPen, @pts(0), 4)

END SUB
' ========================================================================================
```


# <a name="DrawPath"></a>DrawPath (CGpGraphics)

Draws a sequence of lines and curves defined by a **GraphicsPath** object.

```
FUNCTION DrawPath (BYVAL pPen AS CGpPen PTR, BYVAL pPath AS CGpGraphicsPath PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pPen* | Pointer to a pen that is used to draw the path. |
| *pPath* | Pointer to a **GraphicsPath** object that specifies the sequence of lines and curves that make up the path. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example draws a GraphicsPath object.
' ========================================================================================
SUB Example_DrawPath (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a GraphicsPath, and add an ellipse
   DIM ellipsePath AS CGpGraphicsPath
   ellipsePath.AddEllipse(100, 70, 200, 100)

   ' // Create a Pen object
   DIM blackPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 0), 3)

   ' // Draw ellipsePath.
   graphics.DrawPath(@blackPen, @ellipsePath)

END SUB
' ========================================================================================
```


# <a name="DrawPie"></a>DrawPie (CGpGraphics)

Draws a pie.

```
FUNCTION DrawPie (BYVAL pPen AS CGpPen PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, _
   BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE, BYVAL startAngle AS SINGLE, _
   BYVAL sweepAngle AS SINGLE) AS GpStatus
FUNCTION DrawPie (BYVAL pPen AS CGpPen PTR, BYVAL x AS INT_, BYVAL y AS INT_, _
   BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL startAngle AS INT_, _
   BYVAL sweepAngle AS INT_) AS GpStatus
FUNCTION DrawPie (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRectF, _
   BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
FUNCTION DrawPie (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRect, _
   BYVAL startAngle AS INT_, BYVAL sweepAngle AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pPen* | Pointer to a pen that is used to draw the pie. |
| *x* | The x-coordinate of the upper-left corner of the rectangle that bounds the ellipse in which to draw the pie. |
| *y* | The y-coordinate of the upper-left corner of the rectangle that bounds the ellipse in which to draw the pie. |
| *nWidth* | The width of the rectangle that bounds the ellipse in which to draw the pie. |
| *nHeight* | The height of the rectangle that bounds the ellipse in which to draw the pie. |
| *startAngle* | The angle, in degrees, between the x-axis and the starting point of the arc that defines the pie. A positive value specifies clockwise rotation. |
| *sweepAngle* | The angle, in degrees, between the starting and ending points of the arc that defines the pie. A positive value specifies clockwise rotation. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example draws a pie.
' ========================================================================================
SUB Example_DrawPie (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Draw the pie
   DIM blackPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 0), 3)
   graphics.DrawPie(@blackPen, 0.0, 0.0, 200.0, 100.0, 0.0, 45.0)

END SUB
' ========================================================================================
```


# <a name="DrawPolygon"></a>DrawPolygon (CGpGraphics)

Draws a polygon.

```
FUNCTION DrawPolygon (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPointF PTR, BYVAL count AS LONG) AS GpStatus
FUNCTION DrawPolygon (BYVAL pPen AS CGpPen PTR, BYVAL pts AS GpPoint PTR, BYVAL count AS LONG) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pPen* | Pointer to a pen that is used to draw the polygon. |
| *pts* | Pointer to an array of **GpPointF** objects that specify the vertices of the polygon. |
| *nCount* | Integer that specifies the number of elements in the points array. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example draws a sequence of connected lines.
' ========================================================================================
SUB Example_DrawPolygons (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Pen object
   DIM blackPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 0), 3)

   ' // Create an array of GpPoint objects that define the lines to draw
   DIM point1 AS GpPoint : point1.x = 100 : point1.y = 100
   DIM point2 AS GpPoint : point2.x = 200 : point2.y = 130
   DIM point3 AS GpPoint : point3.x = 150 : point3.y = 200
   DIM point4 AS GpPoint : point4.x =  50 : point4.y = 200
   DIM point5 AS GpPoint : point5.x =   0 : point5.y = 130

   DIM pts(0 TO 4) AS GpPoint
   pts(0) = point1
   pts(1) = point2
   pts(2) = point3
   pts(3) = point4
   pts(4) = point5

   ' // Draw the polygon
   graphics.DrawPolygon(@blackPen, @pts(0), 5)

END SUB
' ========================================================================================
```


# <a name="DrawRectangle"></a>DrawRectangle (CGpGraphics)

Draws a rectangle.

```
FUNCTION DrawRectangle (BYVAL pPen AS CGpPen PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, _
   BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS GpStatus
FUNCTION DrawRectangle (BYVAL pPen AS CGpPen PTR, BYVAL x AS INT_, BYVAL y AS INT_, _
   BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS GpStatus
FUNCTION DrawRectangle (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRectF PTR) AS GpStatus
FUNCTION DrawRectangle (BYVAL pPen AS CGpPen PTR, BYVAL rc AS GpRect PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pPen* | Pointer to a pen that is used to draw the rectangle. |
| *x* | The x-coordinate of the upper-left corner of the rectangle. |
| *y* | The y-coordinate of the upper-left corner of the rectangle. |
| *nWidth* | The width of the rectangle. |
| *nHeight* | The height of the rectangle. |
| *rc* | The x, y, width and height of the rectangle as a **GpRect** or **GpRectF** structure. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example draws a rectangle.
' ========================================================================================
SUB Example_DrawRectangle (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Draw the rectangle
   DIM blackPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 0), 3)
   graphics.DrawRectangle(@blackPen, 0, 0, 200, 200)

END SUB
' ========================================================================================
```


# <a name="DrawRectangles"></a>DrawRectangles (CGpGraphics)

Draws a sequence of rectangles.

```
FUNCTION DrawRectangles (BYVAL pPen AS CGpPen PTR, BYVAL rects AS GpRectF PTR, BYVAL count AS INT_) AS GpStatus
FUNCTION DrawRectangles (BYVAL pPen AS CGpPen PTR, BYVAL rects AS GpRect PTR, BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pPen* | Pointer to a pen that is used to draw the rectangles. |
| *rects* | Pointer to an array of **GpRectF** or **GpRect** objects that specify the coordinates of the rectangles to be drawn. |
| *count* | Integer that specifies the number of elements in the rects array. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example draws a group of rectangles.
' ========================================================================================
SUB Example_DrawRectangles (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a pen object
   DIM blackPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 0), 3)

   ' // Create an array of GpRect objects
   DIM rect1 AS GpRect : rect1.x =   0 : rect1.y =   0 : rect1.Width = 100 : rect1.Height = 200
   DIM rect2 AS GpRect : rect2.x = 100 : rect2.y = 200 : rect2.Width = 250 : rect2.Height = 50
   DIM rect3 AS GpRect : rect3.x = 300 : rect3.y =   0 : rect3.Width =   50 : rect3.Height = 100
   DIM rects(2) AS GpRect
   rects(0) = rect1
   rects(1) = rect2
   rects(2) = rect3

   ' // Draw the rectangles
   graphics.DrawRectangles(@blackPen, @rects(0), 3)

END SUB
' ========================================================================================
```


# <a name="DrawString"></a>DrawString (CGpGraphics)

Draws a string based on a font and an origin for the string.

```
FUNCTION DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFont AS CGpFont PTR, _
   BYVAL rc AS GpRectF PTR, BYVAL pBrush AS CGpBrush PTR) AS GpStatus
FUNCTION DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, _
   BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE, BYVAL pBrush AS CGpBrush PTR) AS GpStatus
FUNCTION DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL rc AS GpRect PTR, BYVAL pBrush AS CGpBrush PTR) AS GpStatus
FUNCTION DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, _
   BYVAL nHeight AS INT_, BYVAL pBrush AS CGpBrush PTR) AS GpStatus
FUNCTION DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL rc AS GpPointF PTR, BYVAL pBrush AS CGpBrush PTR) AS GpStatus
FUNCTION DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, _
   BYVAL pBrush AS CGpBrush PTR) AS GpStatus
FUNCTION DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL rc AS GpPoint PTR, BYVAL pBrush AS CGpBrush PTR) AS GpStatus
FUNCTION DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL pBrush AS CGpBrush PTR) AS GpStatus
FUNCTION DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL rc AS GpRectF PTR, BYVAL pBrush AS CGpBrush PTR; _
   BYVAL pStringFormat AS CGpStringFormat PTR) AS GpStatus
FUNCTION DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, _
   BYVAL nHeight AS SINGLE, BYVAL pBrush AS CGpBrush PTR, _
   BYVAL pStringFormat AS CGpStringFormat PTR) AS GpStatus
FUNCTION DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL rc AS GpRect PTR, BYVAL pBrush AS CGpBrush PTR; _
   BYVAL pStringFormat AS CGpStringFormat PTR) AS GpStatus
FUNCTION DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, _
   BYVAL nHeight AS INT_, BYVAL pBrush AS CGpBrush PTR, _
   BYVAL pStringFormat AS CGpStringFormat PTR) AS GpStatus
FUNCTION DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL rc AS GpPointF PTR, BYVAL pBrush AS CGpBrush PTR; _
   BYVAL pStringFormat AS CGpStringFormat PTR) AS GpStatus
FUNCTION DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL pBrush AS CGpBrush PTR, _
   BYVAL pStringFormat AS CGpStringFormat PTR) AS GpStatus
FUNCTION DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL rc AS GpPoint PTR, BYVAL pBrush AS CGpBrush PTR; _
   BYVAL pStringFormat AS CGpStringFormat PTR) AS GpStatus
FUNCTION DrawString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL x AS INT_, BYVAL y AS INT_, _
   BYVAL pBrush AS CGpBrush PTR, BYVAL pStringFormat AS CGpStringFormat PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszString* | Pointer to a wide-character string to be drawn. |
| *pFont* | Pointer to a Font object that specifies the font attributes (the family name, the size, and the style of the font) to use. |
| *x* | The x-coordinate of the upper-left corner of the destination position at which to draw the string. |
| *y* | The y-coordinate of the upper-left corner of the destination position at which to draw the string. |
| *pStringFormat* | Pointer to a **StringFormat** object that specifies text layout information and display manipulations to be applied to the string. |
| *pBrush* | Pointer to a **Brush** object that is used to fill the string. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.


# <a name="EndContainer"></a>EndContainer (CGpGraphics)

Closes a graphics container that was previously opened by the **BeginContainer** method.

```
FUNCTION EndContainer (BYVAL state AS GraphicsContainer) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *state* | Value (previously returned by **BeginContainer**) that identifies the container to be closed. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

When you call the **BeginContainer** method of a **Graphics** object, an information block that holds the state of the **Graphics** object is put on a stack. The **BeginContainer** method returns a value that identifies that information block. When you pass the identifying value to the **EndContainer** method, the information block is removed from the stack and is used to restore the **Graphics** object to the state it was in at the time of the **BeginContainer** call.

Containers can be nested; that is, you can call the **BeginContainer** method several times before you call the **EndContainer** method. Each time you call the **BeginContainer** method, an information block is put on the stack, and you receive an identifier for the information block. When you pass one of those identifiers to the **EndContainer** method, the **Graphics** object is returned to the state it was in at the time of the **BeginContainer** call that returned that particular identifier. The information block placed on the stack by that BeginContainer call is removed from the stack, and all information blocks placed on that stack after that **BeginContainer** call are also removed.

Calls to the **Save** method place information blocks on the same stack as calls to the **BeginContainer** method. Just as an **EndContainer** call is paired with a **BeginContainer** call, a **Restore** call is paired with a **Save** call.

Caution  When you call **EndContainer**, all information blocks placed on the stack (by **Save** or by **BeginContainer**) after the corresponding call to **BeginContainer** are removed from the stack. Likewise, when you call **Restore**, all information blocks placed on the stack (by **Save** or by **BeginContainer**) after the corresponding call to **Save** are removed from the stack.


# <a name="EnumerateMetafile"></a>EnumerateMetafile (CGpGraphics)

Calls an application-defined callback function for each record in a specified metafile. You can use this method to display a metafile by calling **PlayRecord** in the callback function.

**Note**: Not yet implemented.

```
FUNCTION EnumerateMetafileDestPoint (BYVAL pMetafile AS CGpMetafile PTR, BYVAL destPoint AS GpPointF PTR, _
   BYVAL pCallback AS EnumerateMetafileProc, BYVAL pCallbackData AS ANU PTR = NULL, _
   BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL) AS GpStatus
FUNCTION EnumerateMetafileDestPoints (BYVAL pMetafile AS CGdipMetafile PTR, BYVAL destPoints AS PointF PTR, _
   BYVAL nCount AS INT_, BYVAL pCallback AS EnumerateMetafileProc, BYVAL pCallbackData AS ANY PTR = NULL, _
   BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL) AS GpStatus
FUNCTION EnumerateMetafileDestRect (BYVAL pMetafile AS CGdMetafile PTR, BYVAL destRect AS RectF PTR, _
   BYVAL pCallback AS EnumerateMetafileProc, BYVAL pCallbackData AS ANY PTR = NULL, _
   BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL) AS GpStatus
FUNCTION EnumerateMetafileSrcRectDestPoint (BYVAL pMetafile AS CGpMetafile PTR, BYVAL destPoint AS PointF PTR, _
   BYVAL srcRect AS RectF PTR, BYVAL srcUnit AS GpUnit, BYVAL pCallback AS EnumerateMetafileProc, _
   BYVAL pCallbackData AS ANY PTR = NULL, BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL) AS GpStatus
FUNCTION EnumerateMetafileSrcRectDestPoint (BYVAL pMetafile AS IGdipMetafile PTR, BYVAL destPoints AS PointF PTR, _
   BYVAL nCount AS INT_, BYVAL srcRect AS RectF PTR, BYVAL srcUnit AS GpUnit, _
   BYVAL pCallback AS EnumerateMetafileProc, BYVAL pCallbackData AS ANY PTR = NULL, _
   BYVAL pImageAttributes AS IGdipImageAttributes PTR = NULL) AS LONG
FUNCTION EnumerateMetafileSrcRectDestRect (BYVAL pMetafile AS CGpMetafile PTR, BYVAL destRect AS RectF PTR, _
   BYVAL nCount AS INT_, BYVAL srcRect AS RectF PTR, BYVAL srcUnit AS GpUnit, _
   BYVAL pCallback AS EnumerateMetafileProc, BYVAL pCallbackData AS ANY PTR = NULL, _
   BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMetafile* | Pointer to a metafile to be enumerated. |
| *destPoint* | Reference to a point that specifies the upper-left corner of the displayed metafile. |
| *pCallback* | Optional. Pointer to an application-defined callback function. The prototype for the callback function is given in Gdiplustypes.inc. |
| *pCallbackData* | Optional. Pointer to a block of data that is passed to the callback function. The default value is NULL. |
| *pImageAttributes* | Optional. Pointer to an **ImageAttributes** object that specifies color adjustments for the displayed metafile. The default value is NULLL. |
| *destPoints* | Pointer to an array of destination points. This is an array of three points that defines the destination parallelogram for the displayed metafile. |
| *nCount* | Integer that specifies the number of points in the *destPoints* array. |
| *destRect* | Reference to a **GpRectF** object that specifies the rectangle in which the metafile is displayed. |
| *srcRect* | Reference to a rectangle that specifies the portion of the metafile that is displayed. |
| *srcUnit* | Element of the **GpUnit** enumeration that specifies the unit of measure for the source rectangle. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.


# <a name="ExcludeClip"></a>ExcludeClip (CGpGraphics)

Updates the clipping region to the portion of itself that does not intersect the specified rectangle.

```
FUNCTION ExcludeClipRect (BYVAL rc AS GpRectF PTR) AS GpStatus
FUNCTION ExcludeClipRect (BYVAL rc AS GpRect PTR) AS GpStatus
FUNCTION ExcludeClipRect (BYVAL x AS SINGLE, BYVAL y AS SINGLE, _
   BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS GpStatus
FUNCTION ExcludeClipRect (BYVAL x AS INT_, BYVAL y AS INT_, _
   BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS GpStatus
```

Updates the clipping region with the portion of itself that does not overlap the specified region.

```
FUNCTION ExcludeClipRect (BYVAL pRegion AS CGpRegion PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *rc* | Reference to a rectangle to use to update the clipping region. |
| *x* | The x-coordinate of the upper-left corner of the rectangle. |
| *y* | The y-coordinate of the upper-left corner of the rectangle. |
| *nWidth* | The width of the rectangle. |
| *nHeight* | The height of the rectangle. |
| *pRegion* | Pointer to a **Region** object. |

#### Example

```
' ========================================================================================
' The following example uses a rectangle to update a clipping region and then draws a
' rectangle that demonstrates the updated clipping region.
' ========================================================================================
SUB Example_ExcludeClip (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a GpRect object, and set the clipping region to its exclusion.
   DIM excludeRect AS GpRect
   excludeRect.x = 125 : excludeRect.y = 50 : excludeRect.Width = 150 : excludeRect.Height = 150
   graphics.ExcludeClip(@excludeRect)

   ' // Fill a rectangle to demonstrate the clipping region.
   graphics.FillRectangle(@CGpSolidBrush(GDIP_ARGB(255, 0, 0, 255)), 0, 0, 400, 250)

END SUB
' ========================================================================================
```


# <a name="FillClosedCurve"></a>FillClosedCurve (CGpGraphics)

Creates a closed cardinal spline from an array of points and uses a brush to fill the interior of the spline.

```
FUNCTION FillClosedCurve (BYVAL pBrush AS CGpBrush PTR, BYVAL pts AS GpPointF PTR, BYVAL count AS INT_) AS GpStatus
FUNCTION FillClosedCurve (BYVAL pBrush AS CGpBrush PTR, BYVAL pts AS GpPoint PTR, BYVAL count AS INT_) AS GpStatus
FUNCTION FillClosedCurve (BYVAL pBrush AS CGpBrush PTR, BYVAL pts AS GpPointF PTR, BYVAL count AS INT_, _
   BYVAL nFillMode AS FillMode, BYVAL tension AS SINGLE = 0.5) AS GpStatus
FUNCTION FillClosedCurve (BYVAL pBrush AS CGpBrush PTR, BYVAL pts AS GpPoint PTR, BYVAL count AS INT_, _
   BYVAL nFillMode AS FillMode, BYVAL tension AS SINGLE = 0.5) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pBrush* | Pointer to a **Brush** object that is used to paint the interior of the spline. |
| *pts* | Pointer to an array of points that this method uses to create a closed cardinal spline. Each point in the array is a point on the spline. |
| *count* | Integer that specifies the number of points in the points array. |
| *fillMode* | Element of the **FillMode** enumeration that specifies how to fill a closed area that is created when the curve passes over itself. |
| *tension* | Optional. Nonnegative real number that specifies how tightly the spline bends as it passes through the points. A value of 0 specifies that the spline is a sequence of straight lines. As the value increases, the curve becomes fuller. The default value is 0.5!. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example draws a closed cardinal spline.
' ========================================================================================
SUB Example_FillClosedCurve (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Define a Brush object and an array of Point objects.
   DIM blackBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 0)
   DIM point1 AS GpPoint : point1.x = 100 : point1.y = 100
   DIM point2 AS GpPoint : point2.x = 200 : point2.y = 50
   DIM point3 AS GpPoint : point3.x = 250 : point3.y = 200
   DIM point4 AS GpPoint : point4.x =  50 : point4.y = 150

   DIM pts(3) AS GpPoint
   pts(0) = point1
   pts(1) = point2
   pts(2) = point3
   pts(3) = point4

   ' //Fill the curve.
   graphics.FillClosedCurve(@blackBrush, @pts(0), 4)

END SUB
' ========================================================================================
```


# <a name="FillEllipse"></a>FillEllipse (CGpGraphics)

Uses a brush to fill the interior of an ellipse that is specified by coordinates and dimensions.

```
FUNCTION FillEllipse (BYVAL pBrush AS CGpBrush PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, _
   BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS GpStatus
FUNCTION FillEllipse (BYVAL pBrush AS CGpBrush PTR, BYVAL x AS INT_, BYVAL y AS INT_, _
   BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS GpStatus
FUNCTION FillEllipse (BYVAL pBrush AS CGpBrush PTR, BYVAL rcAS GpRectF) AS GpStatus
FUNCTION FillEllipse (BYVAL pBrush AS CGpBrush PTR, BYVAL rcAS GpRect) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pBrush* | Pointer to a **Brush** object that is used to paint the interior of the ellipse. |
| *x* | The x-coordinate of the upper-left corner of the rectangle that specifies the boundaries of the ellipse. |
| *y* | The y-coordinate of the upper-left corner of the rectangle that specifies the boundaries of the ellipse. |
| *nWidth* | The width of the rectangle that specifies the boundaries of the ellipse. |
| *nHeight* | The height of the rectangle that specifies the boundaries of the ellipse. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example fills an ellipse that is defined by coordinates and dimensions.
' ========================================================================================
SUB Example_FillEllipse (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a SolidBrush object
   DIM blackBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 0)

   ' // Fill the ellipse.
   graphics.FillEllipse(@blackBrush, 0, 0, 200.1, 100.4)

END SUB
' ========================================================================================
```


# <a name="FillPath"></a>FillPath (CGpGraphics)

Uses a brush to fill the interior of a path. If a figure in the path is not closed, this method treats the nonclosed figure as if it were closed by a straight line that connects the figure's starting and ending points.

```
FUNCTION FillPath (BYVAL pBrush AS CGpBrush PTR, BYVAL pPath AS CGpGraphicsPath PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pBrush* | Pointer to a brush that is used to paint the interior of the path. |
| *pPath* | Pointer to a **GraphicsPath** object that specifies the path. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example fills a path.
' ========================================================================================
SUB Example_FillPath (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a GraphicsPath and add an ellipse
   DIM ellipsePath AS CGpGraphicsPath
   ellipsePath.AddEllipse(0, 0, 200, 100)

   ' // Create a SolidBrush object
   DIM blackBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 0)

   ' // Draw ellipsePath
   graphics.FillPath(@blackBrush, @ellipsePath)

END SUB
' ========================================================================================

```


# <a name="FillPie"></a>FillPie (CGpGraphics)

Uses a brush to fill the interior of a pie.

```
FUNCTION FillPie (BYVAL pBrush AS CGpBrush PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, _
   BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE, BYVAL startAngle AS SINGLE, _
   BYVAL sweepAngle AS SINGLE) AS GpStatus
FUNCTION FillPie (BYVAL pBrush AS CGpBrush PTR, BYVAL x AS INT_, BYVAL y AS INT_, _
   BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL startAngle AS SINGLE, _
   BYVAL sweepAngle AS SINGLE) AS GpStatus
FUNCTION FillPie (BYVAL pBrush AS CGpBrush PTR, BYVAL rc AS GpRectF PTR) AS GpStatus
FUNCTION FillPie (BYVAL pBrush AS CGpBrush PTR, BYVAL rc AS GpRect PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pBrush* | Pointer to a brush that is used to paint the interior of the pie. |
| *x* | The x-coordinate of the upper-left corner of the rectangle that bounds the ellipse. A curved portion of the ellipse is the arc of the pie. |
| *y* | The y-coordinate of the upper-left corner of the rectangle that bounds the ellipse. |
| *nWidth* | The width of the rectangle that bounds the ellipse. |
| *nHeight* | The height of the rectangle that bounds the ellipse. |
| *startAngle* | The angle, in degrees, between the x-axis and the starting point of the pie's arc. |
| *swepAngle* | The angle, in degrees, between the starting and ending points of the pie's arc. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example defines a pie and then fills it.
' ========================================================================================
SUB Example_FillPie (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a GraphicsPath and add an ellipse
   DIM ellipsePath AS CGpGraphicsPath
   ellipsePath.AddEllipse(0, 0, 200, 100)

   ' // Create a SolidBrush object
   DIM blackBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 0)

   ' // Fill the pie.
   graphics.FillPie(@blackBrush, 0, 0, 200, 100, 0, 45)

END SUB
' ========================================================================================
```


# <a name="FillPolygon"></a>FillPolygon (CGpGraphics)

Uses a brush to fill the interior of a polygon.

```
FUNCTION FillPolygon (BYVAL pBrush AS CGpBrush PTR, BYVAL pts AS GpPointF PTR, BYVAL count AS INT_) AS GpStatus
FUNCTION FillPolygon (BYVAL pBrush AS CGpBrush PTR, BYVAL pts AS GpPoint PTR, BYVAL count AS INT_) AS GpStatus
FUNCTION FillPolygon (BYVAL pBrush AS CGpBrush PTR, BYVAL pts AS GpPointF PTR, _
   BYVAL count AS INT_, BYVAL nFillMode AS FillMode) AS GpStatus
FUNCTION FillPolygon (BYVAL pBrush AS CGpBrush PTR, BYVAL pts AS GpPoint PTR, _
   BYVAL count AS INT_, BYVAL nFillMode AS FillMode) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pBrush* | Pointer to a brush that is used to paint the interior of the polygon. |
| *pts* | Pointer to an array of points that make up the vertices of the polygon. The first two points in the array specify the first side of the polygon. Each additional point specifies a new side, the vertices of which include the point and the previous point. If the last point and the first point do not coincide, they specify the last side of the polygon. |
| *nCount* | Integer that specifies the number of points in the points array. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example uses a brush to fill the interior of a polygon.
' ========================================================================================
SUB Example_FillPolygon (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a SolidBrush object
   DIM blackBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 0)

   ' // Create an array of GpPointF objects that define the polygon
   DIM rgPoints(4) AS GpPointF
   rgPoints(0).x = 100.0 : rgPoints(0).y = 200.0
   rgPoints(1).x = 200.0 : rgPoints(1).y = 130.0
   rgPoints(2).x = 150.0 : rgPoints(2).y = 200.0
   rgPoints(3).x =  50.0 : rgPoints(3).y = 200.0
   rgPoints(4).x =   0.0 : rgPoints(4).y = 130.0

   ' // Fill the polygon
   graphics.FillPolygon(@blackBrush, @rgPoints(0), 5)

END SUB
' ========================================================================================
```

# <a name="FillRectangle"></a>FillRectangle (CGpGraphics)

Uses a brush to fill the interior of a rectangle.

```
FUNCTION FillRectangle (BYVAL pBrush AS CGpBrush PTR, BYVAL x AS SINGLE, BYVAL y AS SINGLE, _
   BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS GpStatus
FUNCTION FillRectangle (BYVAL pBrush AS CGpBrush PTR,  BYVAL x AS INT_, BYVAL y AS INT_, _
   BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS GpStatus
FUNCTION FillRectangle (BYVAL pBrush AS CGpBrush PTR, BYVAL rc AS GpRectF PTR) AS GpStatus
FUNCTION FillRectangle (BYVAL pBrush AS CGpBrush PTR, BYVAL rc AS GpRect PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pBrush* | Pointer to a brush that is used to paint the interior of the rectangle. |
| *x* | The x-coordinate of the upper-left corner of the rectangle to be filled. |
| *y* | The y-coordinate of the upper-left corner of the rectangle to be filled. |
| *nWidth* | The width of the rectangle to be filled. |
| *nHeight* | The height of the rectangle to be filled. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example uses a brush to fill the interior of a rectangle.
' ========================================================================================
SUB Example_FillRectangle (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a SolidBrush object
   DIM blackBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 0)

   ' // Fill the rectangle
   graphics.FillRectangle(@blackBrush, 0, 0, 100, 100)

END SUB
' ========================================================================================
```


# <a name="FillRectangles"></a>FillRectangles (CGpGraphics)

Uses a brush to fill the interior of a sequence of rectangles.

```
FUNCTION FillRectangles (BYVAL pBrush AS CGpBrush PTR, BYVAL rects AS GpRectF PTR, BYVAL count AS INT_) AS GpStatus
FUNCTION FillRectangles (BYVAL pBrush AS CGpBrush PTR, BYVAL rects AS GpRect PTR, BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pBrush* | Pointer to a brush that is used to paint the interior of each rectangle. |
| *rects* | Pointer to an array of rectangles to be filled. |
| *count* | Integer that specifies the number of rectangles in the rects array. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example fills a sequence of rectangles.
' ========================================================================================
SUB Example_FillRectangles (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a pen object
   DIM blackBrush AS CGpSolidBrush = CGpSolidBrush(GDIP_ARGB(255, 0, 0, 0))

   ' // Create an array of RectF objects.
   DIM rects(0 TO 2) AS GpRectF
   rects(0).x =   0.0 : rects(0).y =   0.0 : rects(0).Width = 100.0 : rects(0).Height = 200.0
   rects(1).x = 100.5 : rects(1).y = 200.5 : rects(1).Width = 200.5 : rects(1).Height = 50.5
   rects(2).x = 300.8 : rects(2).y =   0.8 : rects(2).Width =  50.8 : rects(2).Height = 150.8

   ' // Draw the rectangles
   graphics.FillRectangles(@blackBrush, @rects(0), 3)

END SUB
' ========================================================================================
```


# <a name="FillRegion"></a>FillRegion (CGpGraphics)

Uses a brush to fill a specified region.

```
FUNCTION FillRegion (BYVAL pBrush AS CGpBrush PTR, BYVAL pRegion AS CGpRegion PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pBrush* | Pointer to a brush that is used to paint the region. |
| *pRegion* | Pointer to a region to be filled. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

Because a region describes a set of pixels, a pixel is considered either fully inside, or fully outside the region. Consequently, FillRegion does not antialias the edges of the region.

#### Example

```
' ========================================================================================
' The following example creates a region from a rectangle and then fills the region.
' ========================================================================================
SUB Example_FillRegion (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a SolidBrush object
   DIM blackBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 0)

   ' // Create a Region object from a rectangle.
   DIM ellipseRegion AS CGpRegion = CGpRegion(0, 0, 200, 100)

   ' // Fill the region.
   graphics.FillRegion(@blackBrush, @ellipseRegion)

END SUB
' ========================================================================================
```


# <a name="Flush"></a>Flush (CGpGraphics)

Flushes all pending graphics operations.

```
SUB Flush (BYVAL intention AS FlushIntention = FlushIntentionFlush)
```

| Parameter  | Description |
| ---------- | ----------- |
| *intention* | Element of the **FlushIntention** enumeration that specifies whether pending operations are flushed immediately (not executed) or executed as soon as possible. |


# <a name="FromHDC"></a>FromHDC (CGpGraphics)

Creates a **Graphics** object that is associated with a specified device context.

```
FUNCTION FromHDC (BYVAL hdc AS HDC) AS GpStatus
FUNCTION FromHDC (BYVAL hdc AS HDC, BYVAL hDevice AS HANDLE) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *hdc* | Handle to the device context that will be associated with the new **Graphics** object. |
| *hDevice* | Handle to a device that will be associated with the new **Graphics** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

When you use these methods to create a Graphics object, make sure that the Graphics object is deleted before the device context is released.

#### Example

```
' ========================================================================================
' The following example draws a rectangle.
' ========================================================================================
SUB Example_FromHDC (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Draw the rectangle
   DIM redPen AS CGpPen = CGpPen(GDIP_ARGB(255, 255, 0, 0), 1)
   graphics.DrawRectangle(@redPen, 10, 10, 200, 100)

END SUB
' ========================================================================================
```


# <a name="FromHWND"></a>FromHWND (CGpGraphics)

Creates a **Graphics** object that is associated with a specified window.

```
FUNCTION FromHWND (BYVAL hwnd AS HWND, BYVAL icm AS BOOLEAN = FALSE) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to a window that will be associated with the new **Graphics** object. |
| *icm* | Optional. Boolean value that specifies whether the new **Graphics** object applies color adjustment according to the ICC profile associated with the display device. TRUE specifies that color adjustment is applied, and FALSE specifies that color adjustment is not applied. The default value is FALSE. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.


# <a name="FromImage"></a>FromImage (CGpGraphics)

Creates a **Graphics** object that is associated with a specified device context.

```
FUNCTION FromImage (BYVAL pImage AS CGpImage PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pImage* | Pointer to an Image object that will be associated with the new **Graphics** object. |

This method fails if the **Image** object is based on a metafile that was opened for reading. The Image(file) and Metafile(file) constructors open a metafile for reading. To open a metafile for recording, use a **Metafile** constructor that receives a device context handle.

This method also fails if the image uses one of the following pixel formats:

PixelFormatUndefined<br>
PixelFormatDontCare<br>
PixelFormat1bppIndexed<br>
PixelFormat4bppIndexed<br>
PixelFormat8bppIndexed<br>
PixelFormat16bppGrayScale<br>
PixelFormat16bppARGB1555

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.


# <a name="GetClip"></a>GetClip (CGpGraphics)

Gets the clipping region of this **Graphics** object.

```
FUNCTION GetClip (BYVAL pRegion AS CGpRegion PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pRegion* | Pointer to a **Region** object that receives the clipping region. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example sets a clipping region. Next, the code gets the clipping region,
' stores it in a Region object, and then uses the stored object to fill the region.
' ========================================================================================
SUB Example_GetClip (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Set a clipping region.
   graphics.SetClip(0, 0, 200, 100)

   ' // Get the clipping region.
   DIM clipRegion AS CGpRegion
   graphics.GetClip(@clipRegion)

   ' // Fill the clipping region of the graphics object.
   graphics.FillRegion(@CGpSolidBrush(GDIP_ARGB(255, 0, 0, 255)), @clipRegion)

END SUB
' ========================================================================================
```


# <a name="GetClipBounds"></a>GetClipBounds (CGpGraphics)

Gets a rectangle that encloses the clipping region of this **Graphics** object.

```
FUNCTION GetClipBounds (BYVAL rc AS GpRectF PTR) AS GpStatus
FUNCTION GetClipBounds (BYVAL rc AS GpRect PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *rc* | Pointer to a **GpRectF** or **GpRect** object that receives the rectangle that encloses the clipping region. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

The world transformation is applied to the clipping region and then the enclosing rectangle is calculated.

If you do not explicitly set the clipping region of a **Graphics** object, its clipping region is infinite. When the clipping region is infinite, **GetClipBounds** returns a large rectangle. The **X** and **Y** data members of that rectangle are large negative numbers, and the **Width** and **Height** data members are large positive numbers.

#### Example

```
' ========================================================================================
' The following example sets a clipping region, gets the rectangle that encloses the
' clipping region, and then fills the rectangle.
' ========================================================================================
SUB Example_GetClipBounds (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a region using a rectangle
   DIM myRegion AS CGpRegion = CGpRegion(25, 25, 100, 50)

   ' // Modify the region by using a rectangle
   DIM rcf AS GpRectF : rcf.x = 40 : rcf.y = 50 : rcf.Width = 100 : rcf.Height = 50
   myRegion.Union_(@rcf)

   ' // Set the clipping region of the graphics object
   graphics.SetClip(@myRegion)

   ' // Now, get the clipping region, and fill it
   DIM gRegion AS CGpRegion
   graphics.GetClip(@gRegion)
   DIM blueBrush AS CGpSolidBrush = GDIP_ARGB(100, 0, 0, 255)
   graphics.FillRegion(@blueBrush, @gRegion)

   ' // Get a rectangle that encloses the clipping region, and draw the enclosing rectangle
   DIM enclosingRect AS GpRectF
   graphics.GetClipBounds(@enclosingRect)
   graphics.ResetClip
   DIM greenPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 255, 0), 1.5)
   graphics.DrawRectangle(@greenPen, @enclosingRect)

END SUB
' ========================================================================================
```


# <a name="GetCompositingMode"></a>GetCompositingMode (CGpGraphics)

Gets the compositing mode currently set for this **Graphics** object.

```
FUNCTION GetCompositingMode () AS CompositingMode
```

#### Return value

This method returns an element of the **CompositingMode** enumeration that indicates the compositing mode currently set for this **Graphics** object.

#### Remarks

Suppose you create a **SolidBrush** object based on a color that has an alpha component of 192, which is about 75 percent of 255. If your **Graphics** object has its compositing mode set to **CompositingModeSourceOver**, then areas filled with the solid brush are a blend that is 75 percent brush color and 25 percent background color. If your **Graphics** object has its compositing mode set to **CompositingModeSourceCopy**, then the background color is not blended with the brush color. However, the color rendered by the brush has an intensity that is 75 percent of what it would be if the alpha component were 255.

#### Example

```
' ========================================================================================
' The following example creates a Graphics object and sets its compositing mode to
' CompositingModeSourceCopy. The code creates a SolidBrush object based on a color with an
' alpha component of 128. The code passes the address of that brush to the Graphics.FillRectangle
' method of the Graphics object to fill a rectangle with a color that is not blended with
' the background color. The call to the Graphics.CompositingMode method of the Graphics
' object demonstrates how to obtain the compositing mode (which is already known in this
' case). The code determines whether the compositing mode is CompositingModeSourceCopy and
' if so, changes it to CompositingModeSourceOver. Then the code calls Graphics.FillRectangle
' a second time to fill a rectangle with a color that is a half-and-half blend of the brush
' color and the background color.
' ========================================================================================
SUB Example_CompositingMode (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   graphics.SetCompositingMode(CompositingModeSourceCopy)
   DIM alphaBrush AS CGpSolidBrush = GDIP_ARGB(128, 255, 0, 0)
   graphics.FillRectangle(@alphaBrush, 0, 0, 100, 100)

   ' // Get the compositing mode.
   DIM compMode AS CompositingMode
   compMode = graphics.GetCompositingMode

   ' // Change the compositing mode if it is CompositingModeSourceCopy.
   IF compMode = CompositingModeSourceCopy THEN
      graphics.SetCompositingMode(CompositingModeSourceOver)
   END IF

   graphics.FillRectangle(@alphaBrush, 0, 100, 100, 100)

END SUB
' ========================================================================================
```


# <a name="GetCompositingQuality"></a>GetCompositingQuality (CGpGraphics)

Gets the compositing quality currently set for this **Graphics** object.

```
FUNCTION GetCompositingQuality () AS CompositingQuality
```

#### Return value

This method returns an element of the **CompositingQuality** enumeration that indicates the compositing quality currently set for this **Graphics** object.

#### Example

The following example creates a **Graphics** object and sets its compositing quality to **CompositingQualityHighQuality**. The code creates a **SolidBrush** object based on a color with an alpha component of 128. The code passes the address of that brush to the FillRectangle method of the **Graphics** object. The call to the **GetCompositingQuality** method of the **Graphics** object demonstrates how to obtain the compositing quality (which is already known in this case). The code determines whether the compositing quality is **CompositingQualityHighQuality** and if so, changes it to **CompositingQualityHighSpeed**. Then the code calls the **FillRectangle** method a second time. The second rectangle is filled with the same brush that was used to fill the first rectangle, but the result is different because of the compositing quality setting.

```
' ========================================================================================
' The following example creates a Graphics object and sets its compositing quality to
' CompositingQualityHighQuality. The code creates a SolidBrush object based on a color
' with an alpha component of 128. The code passes the address of that brush to the
' Graphics.FillRectangle method of the Graphics object. The call to the Graphics.CompositingQuality
' method of the Graphics object demonstrates how to obtain the compositing quality (which
' is already known in this case). The code determines whether the compositing quality is
' CompositingQualityHighQuality and if so, changes it to CompositingQualityHighSpeed.
' Then the code calls the Graphics.FillRectangle method a second time. The second rectangle
' is filled with the same brush that was used to fill the first rectangle, but the result
' is different because of the compositing quality setting.
' ========================================================================================
SUB Example_CompositingQuality (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   graphics.SetCompositingQuality(CompositingQualityHighQuality)
   DIM alphaBrush AS CGpSolidBrush = GDIP_ARGB(128, 255, 0, 0)
   graphics.FillRectangle(@alphaBrush, 0, 0, 100, 100)

   ' // Get the compositing mode.
   DIM compQuality AS CompositingMode
   compQuality = graphics.GetCompositingQuality

   ' // Change the compositing quality if it is CompositingQualityHighQuality
   IF compQuality = CompositingQualityHighQuality THEN
      graphics.SetCompositingQuality(CompositingQualityHighSpeed)
   END IF

   graphics.FillRectangle(@alphaBrush, 0, 100, 100, 100)

END SUB
' ========================================================================================
```


# <a name="GetDpiX"></a>GetDpiX (CGpGraphics)

Gets the horizontal resolution, in dots per inch, of the display device associated with this **Graphics** object.

```
FUNCTION GetDpiX () AS SINGLE
```

#### Return value

This method returns the horizontal resolution, in dots per inch, of the display device associated with this **Graphics** object.

#### Example

```
' ========================================================================================
' The following example gets the horizontal resolution, in dots per inch, of the display
' device and uses that value to convert pixels to inches. Then the code draws two rectangles
' that have the same width: one measured in inches and one measured in pixels.
' ========================================================================================
SUB Example_GetDpiX (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc

   ' // Get the vertical resolution, in dots per inch, of the display device
   DIM dpiX AS SINGLE = graphics.GetDpiX

   ' // Set the unit of measure for graphics to inches
   graphics.SetPageUnit(UnitInch)
   graphics.SetPageScale(dpiX / 96)

   ' // Use dpiX to convert pixels to inches, and draw a
   DIM side AS SINGLE = 100.0 / dpiX
   DIM bluePen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 0)
   graphics.DrawRectangle(@bluePen, 0.0, 0.0, side, side)

   ' // Set the unit of measure for graphics to pixels
   graphics.SetPageUnit(UnitPixel)
   graphics.SetPageScale(dpiX / 96)

   ' // Draw a 100-pixel square.
   DIM redPen AS CGpPen = CGpPen(GDIP_ARGB(255, 255, 0, 0), 1)
   graphics.DrawRectangle(@redPen, 120, 0, 100, 100)

END SUB
' ========================================================================================
```


# <a name="GetDpiY"></a>GetDpiY (CGpGraphics)

Gets the vertical resolution, in dots per inch, of the display device associated with this Graphics object.

```
FUNCTION GetDpiY () AS SINGLE
```

#### Return value

This method returns the vertical resolution, in dots per inch, of the display device associated with this **Graphics** object.

#### Example

```
' ========================================================================================
' The following example gets the vertical resolution, in dots per inch, of the display
' device and uses that value to convert pixels to inches. Then the code draws two rectangles
' that have the same height: one measured in inches and one measured in pixels.
' ========================================================================================
SUB Example_GetDpiY (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc

   ' // Get the vertical resolution, in dots per inch, of the display device
   DIM dpiY AS SINGLE = graphics.GetDpiY

   ' // Set the unit of measure for graphics to inches
   graphics.SetPageUnit(UnitInch)
   graphics.SetPageScale(dpiY / 96)

   ' // Use dpiY to convert pixels to inches, and draw a
   ' // rectangle that has a width of 100 pixels
   DIM side AS SINGLE = 100.0 / dpiY
   DIM bluePen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 0)
   graphics.DrawRectangle(@bluePen, 0.0, 0.0, side, side)

   ' // Set the unit of measure for graphics to pixels
   graphics.SetPageUnit(UnitPixel)
   graphics.SetPageScale(dpiY / 96)

   ' // Draw a 100-pixel square.
   DIM redPen AS CGpPen = CGpPen(GDIP_ARGB(255, 255, 0, 0), 1)
   graphics.DrawRectangle(@redPen, 120, 0, 100, 100)

END SUB
' ========================================================================================
```


# <a name="GetHalftonePalette"></a>GetHalftonePalette (CGpGraphics)

Gets a Windows halftone palette.

```
FUNCTION GetHalftonePalette () AS HPALETTE
```

#### Remarks

The purpose of the **GetHalftonePalette** method is to enable GDI+ to produce a better quality halftone when the display uses 8 bits per pixel. To display an image using the halftone palette, use the following procedure:

1. Call **GetHalftonePalette** to get a GDI+ halftone palette.
2. Select the halftone palette into a device context.
3. Realize the palette by calling the **RealizePalette** function.
4. Construct a **Graphics** object from a handle to the device context.
5. Call the **DrawImage** method of the **Graphics** object.

Be sure to delete the palette when you have finished using it. If you do not follow the preceding procedure, then on an 8-bits-per-pixel-display device, the default, 16-color process is used, which results in a lesser quality halftone.

#### Example

```
' ========================================================================================
' The following example draws the same image twice. Before the image is drawn the second
' time, the code gets a halftone palette, selects the palette into a device context, and
' realizes the palette.
' ========================================================================================
SUB Example_GetHalfTonePalette (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96

   ' // Create an Image object.
   DIM pImage AS CGpImage = "climber.jpg"

   DIM graphics1 AS CGpGraphics = hdc
   graphics1.DrawImage(@pImage, 10 * rxRatio, 10 * rxRatio)

   DIM hPalette AS HPALETTE
   hPalette = graphics1.GetHalfTonePalette
   SelectPalette(hdc, hPalette, FALSE)
   RealizePalette(hdc)
   DIM graphics2 AS CGpGraphics = hdc
   graphics2.DrawImage(@pImage, 210 *rxRatio, 10 * rxRatio)
   DeleteObject(hPalette)

END SUB
' ========================================================================================
```

# <a name="GetHDC"></a>GetHDC (CGpGraphics)

Gets a handle to the device context associated with this **Graphics** object.

```
FUNCTION GetHDC () AS HDC
```

#### Remarks

Each call to the **GetHDC** method of a **Graphics** object should be paired with a call to the **ReleaseHDC** method of that same Graphics object. Do not call any methods of the **Graphics** object between the calls to GetHDC and **ReleaseHDC**. If you attempt to call a method of the **Graphics** object between **GetHDC** and **ReleaseHDC**, the method will fail and will return **ObjectBusy**.

Any state changes you make to the device context between **GetHDC** and **ReleaseHDC** will be ignored by GDI+ and will not be reflected in rendering done by GDI+.

#### Example

```
' ========================================================================================
' The following function uses GDI+ to draw an ellipse, then uses GDI to draw a rectangle,
' and finally uses GDI+ to draw a line. The function's one parameter is a pointer to a GDI+
' Graphics object. The code calls the Graphics.DrawEllipse method of that Graphics object
' to draw an ellipse. Next, the code calls the Graphics.GetHDC method to obtain a handle
' to the device context associated with the Graphics object. The code draws a rectangle by
' passing the device context handle to the GDI Rectangle function. The code calls the
' Graphics.ReleaseHDC method of the Graphics object and then uses the Graphics object to
' draw a line.
' ========================================================================================
SUB Example_GetHDC (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM bluePen AS CGpPen = GDIP_ARGB(255, 0, 0, 255)
   graphics.DrawEllipse(@bluepen, 10, 10, 100, 50)   ' // GDI+

   DIM hdc2 AS HDC = graphics.GetHDC

  ' // Make GDI calls, but don't call any methods
  ' // on graphics until after the call to ReleaseHDC.
   Rectangle(hdc2, 120 * rxRatio, 10 * rxRatio, 220 * rxRatio, 60 * rxRatio)   ' // GDI
   graphics.ReleaseHDC(hdc)

   ' // Ok to call methods on g again.
   graphics.DrawLine(@bluePen, 240, 10, 340, 60)

END SUB
' ========================================================================================
```


# <a name="GetInterpolationMode"></a>GetInterpolationMode (CGpGraphics)

Gets the interpolation mode currently set for this Graphics object. The interpolation mode determines the algorithm that is used when images are scaled or rotated.

```
FUNCTION GetInterpolationMode () AS InterpolationMode
```


# <a name="GetNearestColor"></a>GetNearestColor (CGpGraphics)

Gets the nearest color to the color that is passed in. This method works on 8-bits per pixel or lower display devices for which there is an 8-bit color palette.

```
FUNCTION GetNearestColor (BYVAL colour AS ARGB PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *colour* | Pointer to a DWORD variable that, on input, specifies the color to be tested and, on output, receives the nearest color found in the color palette. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Color object and fills a rectangle with that color. It
' then gets the nearest 8-bit color and fills a second rectangle with that color.
' ========================================================================================
SUB Example_GetNearestColor (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Color object, and fill a rectangle with that color.
   DIM colour AS ARGB = &hFFA53F88
   graphics.FillRectangle(@CGpSolidBrush(colour), 0, 0, 100, 100)

   ' // Get the nearest 8-bit color, and fill a second rectangle with that color.
   graphics.GetNearestColor(@colour)
   graphics.FillRectangle(@CGpSolidBrush(colour), 100, 0, 100, 100)

END SUB
' ========================================================================================
```


# <a name="GetPageScale"></a>GetPageScale (CGpGraphics)

Gets the scaling factor currently set for the page transformation of this **Graphics** object. The page transformation converts page coordinates to device coordinates.

```
FUNCTION GetPageScale () AS SINGLE
```

#### Example

```
' ========================================================================================
' The following example sets the world and page transformations of a Graphics object. The
' page scale and page unit both belong to the page transformation. The code demonstrates
' how to obtain the page scale (which is already known in this case). The code determines
' whether the page scale is greater than 1 and, if so, sets the page unit to UnitMillimeter.
' The call to Graphics.DrawRectangle draws a rectangle that has a width of 3 centimeters
' (UnitMillimeter along with a scaling factor of 10) and a height of 2 centimeters.
' ========================================================================================
SUB Example_GetPageScale (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc

   DIM blackPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 0), 0)

   '// Set the world transformation of the Graphics object.
   graphics.TranslateTransform(4, 1)

   ' // Set the page transformation of the Graphics object.
   graphics.SetPageUnit(UnitInch)
   graphics.SetPageScale(10)

   DIM pageScale AS SINGLE = graphics.GetPageScale
   IF pageScale > 1 THEN graphics.SetPageUnit(UnitMillimeter)

   graphics.DrawRectangle(@blackPen, 0, 0, 3, 2)

END SUB
' ========================================================================================
```


# <a name="GetPageUnit"></a>GetPageUnit (CGpGraphics)

Gets the unit of measure currently set for this **Graphics** object.

```
FUNCTION GetPageUnit () AS GpUnit
```

#### Example

```
' ========================================================================================
' The following example creates a Graphics object and sets its unit of measure to UnitPixel.
' Then the code draws a square that is 100 pixels on each side. The call to Graphics.GetPageUnit
' demonstrates how to obtain the unit of measure (which is already known in this case) for
' the Graphics object. The code determines whether the unit of measure is UnitPixel and if
' so, changes the unit of measure to UnitInch. Then the code draws a square that is 1 inch
' on each side.
' ========================================================================================
SUB Example_GetPageUnit (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96

   DIM blackPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 0), 0)
   DIM redPen AS CGpPen = CGpPen(GDIP_ARGB(255, 255, 0, 0), 0)

   graphics.SetPageUnit(UnitPixel)
   graphics.DrawRectangle(@blackPen, 0, 0, 100 * rxRatio, 100 * rxRatio)

   ' // Get the page unit.
   DIM pageUnit AS GpUnit
   pageUnit = graphics.GetPageUnit

   IF pageUnit = UnitPixel THEN
      graphics.SetPageUnit(UnitInch)
   END IF

   graphics.DrawRectangle(@redPen, 2, 0, 1, 1)

END SUB
' ========================================================================================
```


# <a name="GetPixelOffsetMode"></a>GetPixelOffsetMode (CGpGraphics)

Gets the pixel offset mode currently set for this **Graphics** object.

```
FUNCTION GetPixelOffsetMode () AS PixelOffsetMode
```

#### PixelOffsetMode Enumeration

The **PixelOffsetMode** enumeration specifies the pixel offset mode. This enumeration is used by the **GetPixelOffsetMode** and **SetPixelOffsetMode** methods of the Graphics class.

| Cpmstant   | Description |
| ---------- | ----------- |
| **PixelOffsetModeInvalid** | Used internally.  |
| **PixelOffsetModeDefault** | Equivalent to **PixelOffsetModeNone**. |
| **PixelOffsetModeHighSpeed** | Equivalent to **PixelOffsetModeNone**. |
| **PixelOffsetModeHighQuality** | Equivalent to **PixelOffsetModeHalf**. |
| **PixelOffsetModeNone** | Indicates that pixel centers have integer coordinates. |
| **PixelOffsetModeHalf** | Indicates that pixel centers have coordinates that are half way between integer values. |

#### Remarks

Consider the pixel in the upper-left corner of an image with address (0, 0). With **PixelOffsetModeNone**, the pixel covers the area between 0.5 and 0.5 in both the x and y directions; that is, the pixel center is at (0, 0). With **PixelOffsetModeHalf**, the pixel covers the area between 0 and 1 in both the x and y directions; that is, the pixel center is at (0.5, 0.5).


# <a name="GetRenderingOrigin"></a>GetRenderingOrigin (CGpGraphics)

Gets the rendering origin currently set for this **Graphics** object. The rendering origin is used to set the dither origin for 8-bits per pixel and 16-bits per pixel dithering and is also used to set the origin for hatch brushes.

```
FUNCTION GetRenderingOrigin (BYVAL x AS LONG PTR, BYVAL y AS LONG PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | Pointer to an LONG that receives the x-coordinate of the rendering origin |
| *y* | Pointer to an LONG that receives the y-coordinate of the rendering origin. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.



# <a name="GetSmoothingMode"></a>GetSmoothingMode (CGpGraphics)

Determines whether smoothing (antialiasing) is applied to the **Graphics** object.

```
FUNCTION GetSmoothingMode () AS SmoothingMode
```

#### SmoothingMode Enumeration

The **SmoothingMode** enumeration specifies the type of smoothing (antialiasing) that is applied to lines and curves. This enumeration is used by the **GetSmoothingMode** and **SetSmoothingMode** functions.

| Constant   | Description |
| ---------- | ----------- |
| **SmoothingModeInvalid** | Reserved. |
| **SmoothingModeDefault** | Specifies that smoothing is not applied. |
| **SmoothingModeHighSpeed** | Specifies that smoothing is not applied. |
| **SmoothingModeHighQuality** | Specifies that smoothing is applied using an 8 X 4 box filter. |
| **SmoothingModeNone** | Specifies that smoothing is not applied. |
| **SmoothingModeAntiAlias8x4** | Specifies that smoothing is applied using an 8 X 4 box filter. |
| **SmoothingModeAntiAlias** | Specifies that smoothing is applied using an 8 X 4 box filter. |
| **SmoothingModeAntiAlias8x8** | Specifies that smoothing is applied using an 8 X 8 box filter. |

#### Remarks

Smoothing performed by an 8 X 4 box filter gives better results for nearly vertical lines than it does for nearly horizontal lines. Smoothing performed by an 8 X 8 box filter gives equally good results for nearly vertical and nearly horizontal lines. The 8x8 algorithm produces higher quality smoothing but is slower than the 8 X 4 algorithm.


#### Return value

If smoothing (antialiasing) is applied to this **Graphics** object, this method returns **SmoothingModeAntiAlias**. If smoothing (antialiasing) is not applied to this **Graphics** object, this method returns **SmoothingModeAntiAlias** and **SmoothingModeNone** are elements of the **SmoothingMode** enumeration.


# <a name="GetTextContrast"></a>GetTextContrast (CGpGraphics)

Gets the contrast value currently set for this **Graphics** object. The contrast value is used for antialiasing text.

```
FUNCTION GetTextContrast () AS UINT
```


# <a name="GetTextRenderingHint"></a>GetTextRenderingHint (CGpGraphics)

Returns the text rendering mode currently set for this **Graphics** object.

```
FUNCTION GetTextRenderingHint () AS TextRenderingHint
```

#### TextRenderingHint Enumeration

Specifies the process used to render text. The process affects the quality of the text.

| Constant   | Description |
| ---------- | ----------- |
| **TextRenderingHintSystemDefault** | Specifies that a character is drawn using the currently selected system font smoothing mode (also called a rendering hint). |
| **TextRenderingHintSingleBitPerPixelGridFit** | Specifies that a character is drawn using its glyph bitmap and hinting to improve character appearance on stems and curvature. |
| **TextRenderingHintSingleBitPerPixel** | Specifies that a character is drawn using its glyph bitmap and no hinting. This results in better performance at the expense of quality. |
| **TextRenderingHintAntiAliasGridFit** | Specifies that a character is drawn using its antialiased glyph bitmap and hinting. This results in much better quality due to antialiasing at a higher performance cost. |
| **TextRenderingHintAntiAlias**| Specifies that a character is drawn using its antialiased glyph bitmap and no hinting. Stem width differences may be noticeable because hinting is turned off. |
| **TextRenderingHintClearTypeGridFit** | Specifies that a character is drawn using its glyph Microsoft ClearType bitmap and hinting. This type of text rendering cannot be used along with **CompositingModeSourceCopy**.<br>Microsoft Windows XP and Windows Server 2003 only: ClearType rendering is supported only on Windows XP and Windows Server 2003. Therefore, **TextRenderingHintClearTypeGridFit** is ignored on other operating systems even though Windows GDI+ is supported on those operating systems. |

#### Remarks

The quality associated with each process varies according to the circumstances. **TextRenderingHintClearTypeGridFit** provides the best quality for most LCD monitors and relatively small font sizes. **TextRenderingHintAntiAlias** provides the best quality for rotated text. Generally, a process that produces higher quality text is slower than a process that produces lower quality text.


# <a name="GetTransform"></a>GetTransform (CGpGraphics)

Gets the world transformation matrix of this **Graphics** object.

```
FUNCTION GetTransform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMatrix* | Pointer to a **Matrix** object that receives the transformation matrix. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.



# <a name="GetVisibleClipBounds"></a>GetVisibleClipBounds (CGpGraphics)

Gets a rectangle that encloses the visible clipping region of this **Graphics** object. The visible clipping region is the intersection of the clipping region of this **Graphics** object and the clipping region of the window.

```
FUNCTION GetVisibleClipBounds (BYVAL rc AS GpRectF PTR) AS GpStatus
FUNCTION GetVisibleClipBounds (BYVAL rc AS GpRect PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *rc* | Pointer to a **GpRectF** or **GpRect** object that receives the rectangle that encloses the visible clipping region. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example sets the clipping region for the Graphics object. It then gets a
' rectangle that encloses the visible clipping region and fills that rectangle.
' ========================================================================================
SUB Example_GetVisibleClipBounds (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   '// Set the clipping region.
   DIM rc AS GpRect : rc.x = 100 : rc.y = 100 : rc.Width = 200 : rc.Height = 100
   graphics.SetClip(@rc)

   ' // Get a bounding rectangle for the clipping region.
   DIM boundRect AS GpRect
   graphics.GetVisibleClipBounds(@boundRect)

   ' // Fill the bounding rectangle.
   graphics.FillRectangle(@CGpSolidBrush(GDIP_ARGB(255, 0, 0, 0)), @boundRect)

END SUB
' ========================================================================================
```


# <a name="IntersectClip"></a>IntersectClip (CGpGraphics)

Updates the clipping region of this **Graphics** object to the portion of the specified rectangle that intersects with the current clipping region of this **Graphics** object.

```
FUNCTION IntersectClip (BYVAL rc AS GpRectF PTR) AS GpStatus
FUNCTION IntersectClip (BYVAL rc AS GpRect PTR) AS GpStatus
FUNCTION IntersectClip (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, _
   BYVAL nHeight AS SINGLE) AS GpStatus
FUNCTION IntersectClip (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS GpStatus
FUNCTION IntersectClip (BYVAL pRegion AS CGpRegion PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *rc* | Reference to a rectangle that is used to update the clipping region. |
| *pRegion* | Pointer to a region that is used to update the clipping region of this **Graphics** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example sets a clipping region and updates the clipping region. It then
' draws rectangles to demonstrate the effective clipping region.
' ========================================================================================
SUB Example_IntersectClip (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Set the clipping region.
   DIM clipRect AS GpRectF
   clipRect.x = 0.5 : clipRect.y = 0.5 : clipRect.Width = 200.5 : clipRect.Height = 200.5
   graphics.SetClip(@clipRect)

   ' // Update the clipping region to the portion of the rectangle that
   ' // intersects with the current clipping region.
   DIM rcIntersect AS GpRectF
   rcIntersect.x = 100.5 : rcIntersect.y = 100.5 : rcIntersect.Width = 200.5 : rcIntersect.Height = 200.5
   graphics.IntersectClip(@rcIntersect)

   ' // Fill a rectangle to demonstrate the effective clipping region.
   graphics.FillRectangle(@CGpSolidBrush(GDIP_ARGB(255, 0, 0, 255)), 0, 0, 500, 500)

   ' // Reset the clipping region to infinite.
   graphics.ResetClip

   ' // Draw clipRect and intersectRect.
   graphics.DrawRectangle(@CGpPen(GDIP_ARGB(255, 0, 0, 0)), @clipRect)
   graphics.DrawRectangle(@CGpPen(GDIP_ARGB(255, 255, 0, 0)), @rcIntersect)

END SUB
' ========================================================================================
```


# <a name="IsClipEmpty"></a>IsClipEmpty (CGpGraphics)

Determines whether the clipping region of this **Graphics** object is empty.

```
FUNCTION IsClipEmpty () AS BOOLEAN
```

####Return value

If the clipping region of a **Graphics** object is empty, this method returns TRUE; otherwise, it returns FALSE.

#### Remarks

If the clipping region of a **Graphics** object is empty, there is no area left in which to draw. Consequently, nothing will be drawn when the clipping region is empty.

#### Example

```
' ========================================================================================
' The following example determines whether the clipping region is empty and, if it isn't,
' draws a rectangle.
' ========================================================================================
SUB Example_IsClipEmpty (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // If the clipping region is not empty, draw a rectangle.
   IF graphics.IsClipEmpty = FALSE THEN
      graphics.DrawRectangle(@CGpPen(GDIP_ARGB(255, 0, 0, 0), 3), 0, 0, 100, 100)
   END IF

END SUB
' ========================================================================================
```


# <a name="IsVisible"></a>IsVisible (CGpGraphics)

Determines whether the specified point is inside the visible clipping region of this **Graphics** object. The visible clipping region is the intersection of the clipping region of this **Graphics** object and the clipping region of the window.

```
FUNCTION IsVisible (BYVAL pt AS GpPointF PTR) AS BOOLEAN
FUNCTION IsVisible (BYVAL pt AS GpPoint PTR) AS BOOLEAN
FUNCTION IsVisible (BYVAL x AS SINGLE, BYVAL y AS SINGLE) AS BOOLEAN
FUNCTION IsVisible (BYVAL x AS INT_, BYVAL y AS INT_) AS BOOLEAN
FUNCTION IsVisible (BYVAL rc AS GpRectF PTR) AS BOOLEAN
FUNCTION IsVisible (BYVAL rc AS GpRect PTR) AS BOOLEAN
FUNCTION IsVisible (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS BOOLEAN
FUNCTION IsVisible (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *pt* | Reference to a point to be tested to see whether it is inside the visible clipping region. |
| *x, y* | Reference to a point to be tested to see whether it is inside the visible clipping region. |
| *rc* | Reference to a rectangle to be tested to see whether it intersects the visible clipping region. |
| *x, y, nWidth, nHeight* | Reference to a rectangle to be tested to see whether it intersects the visible clipping region. |

#### Return value

If the specified point is inside the visible clipping region, this method returns TRUE; otherwise, it returns FALSE.

#### Example

```
' ========================================================================================
' The following example determines whether the clipping region is empty and, if it isn't,
' draws a rectangle.
' ========================================================================================
SUB Example_IsVisible (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Determine whether the point is visible and, if it is, draw an ellipse.
   IF graphics.IsVisible(100, 100) THEN
      graphics.FillEllipse(@CGpSolidBrush(GDIP_ARGB(255, 0, 0, 0)), 100, 100, 5, 5)
   END IF

END SUB
' ========================================================================================
```


# <a name="IsVisibleClipEmpty"></a>IsVisibleClipEmpty (CGpGraphics)

Determines whether the visible clipping region of this **Graphics** object is empty. The visible clipping region is the intersection of the clipping region of this **Graphics** object and the clipping region of the window.

```
FUNCTION IsVisibleClipEmpty () AS BOOLEAN
```

#### Return value

If the visible clipping region of this Graphics object is empty, this method returns TRUE; otherwise, it returns FALSE.

#### Remarks

If the visible clipping region of a **Graphics** object is empty, there is no area left in which to draw. Consequently, nothing will be drawn when the visible clipping region is empty.

#### Example

```
' ========================================================================================
' The following example determines whether the clipping region is empty and, if it isn't,
' draws a rectangle.
' ========================================================================================
SUB Example_IsVisibleClipEmpty (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // If the clipping region is not empty, draw a rectangle
   IF graphics.IsVisibleClipEmpty = FALSE THEN
      graphics.DrawRectangle(@CGpPen(GDIP_ARGB(255, 0, 0, 0), 3), 0, 0, 100, 100)
   END IF

END SUB
' ========================================================================================
```


# <a name="MeasureCharacterRanges"></a>MeasureCharacterRanges (CGpGraphics)

Gets a set of regions each of which bounds a range of character positions within a string.

```
FUNCTION MeasureCharacterRanges (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL layoutRect AS GpRectF PTR, _
   BYVAL pStringFormat AS CGpStringFormat PTR, BYVAL regionCount AS INT_, _
   BYVAL regions AS CGpRegion PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszString* | Pointer to a wide-character string. Important  For bidirectional: languages, such as Arabic, the string length must not exceed 2046 characters. |
| *pFont* | Pointer to a **Font** object that specifies the font characteristics (the family name, size, and style of the font) to be applied to the string. |
| *layoutRect* | Reference to a rectangle that bounds the string. |
| *stringFormat* | Pointer to a **StringFormat** object that specifies the character ranges and layout information, such as alignment, trimming, tab stops, and so forth. |
| *regionCount* | Integer that specifies the number of regions that are expected to be received into the regions array. This number should be equal to the number of character ranges currently in the **StringFormat** object. |
| *regions* | Pointer to an array of **Region** objects that receives the regions, each of which bounds a range of text. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.


# <a name="MeasureDriverString"></a>MeasureDriverString (CGpGraphics)

Measures the bounding box for the specified characters and their corresponding positions.

```
FUNCTION MeasureDriverString (BYVAL pText AS UINT16 PTR, BYVAL length AS LONG, _
   BYVAL pFont AS CGpFont PTR, BYVAL positions AS ANY PTR, BYVAL flags AS LONG, _
   BYVAL pMatrix AS CGpMatrix PTR, BYVAL boundingBox AS GpRectF PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pText* | Pointer to an array of 16-bit values. If the **DriverStringOptionsCmapLookup** flag is set, each value specifies a Unicode character to be displayed. Otherwise, each value specifies an index to a font glyph that defines a character to be displayed. |
| *length* | Integer that specifies the number of values in the text array. The length parameter can be set to –1 if the string is null terminated. |
| *pFont* | Pointer to a **Font** object that specifies the family name, size, and style of the font to be applied to the string. |
| *positions* | If the **DriverStringOptionsRealizedAdvance** flag is set, positions is a pointer to a **GpPointF** object that specifies the position of the first glyph. Otherwise, positions is an array of **GpPointF** objects, each of which specifies the origin of an individual glyph. |
| *flags* | Integer that specifies the options for the appearance of the string. This value must be an element of the **DriverStringOptions** enumeration or the result of a bitwise **OR** applied to two or more of these elements. |
| *matrix* | Pointer to a **Matrix** object that specifies the transformation matrix to apply to each value in the text array. |
| *boundingBox* | Pointer to a **GpRectF** or **GpRect** object that receives the rectangle that bounds the string. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.


# <a name="MeasureString"></a>MeasureString (CGpGraphics)

Measures the extent of the string in the specified font, format, and layout rectangle.

```
FUNCTION MeasureString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL layoutRect AS GpRectF PTR, _
   BYVAL boundingBox AS GpRectF PTR) AS GpStatus
FUNCTION MeasureString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, BYVAL pFont AS CGpFont PTR, _
   BYVAL layoutRect AS GpPointF PTR, BYVAL boundingBox AS GpRectF PTR) AS GpStatus
FUNCTION MeasureString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL layoutRect AS GpRectF PTR, _
   BYVAL pStringFormat AS CGpStringFormat PTR, BYVAL boundingBox AS GpRectF PTR) AS GpStatus
FUNCTION MeasureString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL layoutRect AS GpRectF PTR, _
   BYVAL pStringFormat AS CGpStringFormat PTR, BYVAL boundingBox AS GpPointF PTR) AS GpStatus
FUNCTION MeasureString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL layoutRect AS GpRectF PTR, _
   BYVAL pStringFormat AS CGpStringFormat PTR, BYVAL boundingBox AS GpRectF PTR, _
   BYREF codepointsFitted AS INT_ = 0, BYREF linesFilled AS INT_ = 0) AS GpStatus
FUNCTION MeasureString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFont AS CGpFont PTR, BYVAL layoutRect AS GpSizeF PTR, _
   BYVAL pStringFormat AS CGpStringFormat PTR, BYVAL boundingBox AS GpSizeF PTR, _
   BYREF codepointsFitted AS INT_ = 0, BYREF linesFilled AS INT_ = 0) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszString* | Pointer to a wide-character string to be measured. Important: For bidirectional languages, such as Arabic, the string length must not exceed 2046 characters. |
| *pFont* | Pointer to a **Font** object that specifies the family name, size, and style of the font to be applied to the string. |
| *origin* | Reference to the point at which the string starts. |
| *layoutRect* | Reference to the rectangle that bounds the string. |
| *stringFormat* | Pointer to a **StringFormat** object that specifies the layout information, such as alignment, trimming, tab stops, and so forth. |
| *boundingBox* | Pointer to a **GpRectF** object that receives the rectangle that bounds the string. |
| *size* | Pointer to a **GpSizeF** object that receives the width and height of the rectangle that bounds the string. |
| *codepointsFitted* | Pointer to an LONG that receives the number of characters that actually fit into the layout rectangle. The default value is a NULL pointer. |
| *linesFilled* | Pointer to an LONG that receives the number of lines that fit into the layout rectangle. The default value is a NULL pointer. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.


# <a name="MultiplyTransform"></a>MultiplyTransform (CGpGraphics)

Updates this **Graphics** object's world transformation matrix with the product of itself and another matrix.

```
FUNCTION MultiplyTransform (BYVAL pMatrix AS CGpMatrix PTR, _
   BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMatrix* | Pointer to a matrix that will be multiplied by the world transformation matrix of this **Graphics** object. |
| *order* | Optional. Element of the **MatrixOrder** enumeration that specifies the order of multiplication. **MatrixOrderPrepend** specifies that the passed matrix is on the left, and **MatrixOrderAppend** specifies that the passed matrix is on the right. The default value is **MatrixOrderPrepend**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example calls the Graphics.RotateTransform method of a Graphics object to
' fill its world transformation matrix with the elements that represent a 30-degree rotation.
' Then the code calls the MultiplyTransform method to replace the world transformation matrix
' (which represents the 30-degree rotation) of the Graphics object with the product of itself
' and a translation matrix. At that point, the world transformation matrix of the Graphics
' object represents a composite transformation: first rotate, then translate. Finally, the
' code calls the DrawEllipse method to draw an ellipse that is rotated and translated.
' ========================================================================================
SUB Example_MultiplyTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM matrix AS CGpMatrix
   matrix.Translate(150 * rxRatio, 100 * rxRatio)

   graphics.RotateTransform(30)   ' // first rotate
   graphics.MultiplyTransform(@matrix, MatrixOrderAppend)   ' // then translate

   DIM bluePen AS CGpPen = GDIP_ARGB(255, 0, 0, 255)
   graphics.DrawEllipse(@bluePen, -80, -40, 160, 80)

END SUB
' ========================================================================================
```

# <a name="ReleaseHDC"></a>ReleaseHDC (CGpGraphics)

Releases a device context handle obtained by a previous call to the **GetHDC** method of this **Graphics** object.

```
SUB ReleaseHDC (BYVAL hdc AS HDC)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hdc* | Handle to a device context obtained by a previous call to the **GetHDC** method of this **Graphics** object. |


# <a name="ResetClip"></a>ResetClip (CGpGraphics)

Sets the clipping region of this **Graphics** object to an infinite region.

```
FUNCTION ResetClip () AS GpStatus
```

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.



# <a name="ResetTransform"></a>ResetTransform (CGpGraphics)

Sets the world transformation matrix of this **Graphics** object to the identity matrix.

```
FUNCTION ResetTransform () AS GpStatus
```

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

The identity matrix represents a transformation that does nothing. If the world transformation matrix of a **Graphics** object is the identity matrix, then no world transformation is applied to items drawn by that **Graphics** object.


# <a name="Restore"></a>Restore (CGpGraphics)

Sets the state of this **Graphics** object to the state stored by a previous call to the **Save** method of this **Graphics** object.

```
FUNCTION Restore (BYVAL gstate AS GraphicsState) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *gstate* | 32-bit value (returned by a previous call to the **Save** method of this Graphics object) that identifies a block of saved state. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.


# <a name="RotateTransform"></a>RotateTransform (CGpGraphics)

Updates the world transformation matrix of this **Graphics** object with the product of itself and a rotation matrix.

```
FUNCTION RotateTransform ( BYVAL angle AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *angle* | The angle, in degrees, of rotation. |
| *order* | Optional. Element of the **MatrixOrder** enumeration that specifies the order of multiplication. **MatrixOrderPrepend** specifies that the passed matrix is on the left, and **MatrixOrderAppend** specifies that the passed matrix is on the right. The default value is **MatrixOrderPrepend**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.


# <a name="Save"></a>Save (CGpGraphics)

Saves the current state (transformations, clipping region, and quality settings) of this **Graphics** object. You can restore the state later by calling the **Restore** method.

```
FUNCTION Save () AS GraphicsState
```

#### Return value

This method returns a value that identifies the saved state. Pass this value to the **Restore** method when you want to restore the state. The **GraphicsState** data type is defined in Gdiplusenums.bi.


# <a name="ScaleTransform"></a>ScaleTransform (CGpGraphics)

Updates this **Graphics** object's world transformation matrix with the product of itself and a scaling matrix.

```
FUNCTION ScaleTransform (BYVAL sx AS SINGLE, BYVAL sy AS SINGLE, _
   BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *sx* | The horizontal scaling factor in the scaling matrix. |
| *sy* | The vertical scaling factor in the scaling matrix. |
| *order* | Optional. Element of the **MatrixOrder** enumeration that specifies the order of multiplication. **MatrixOrderPrepend** specifies that the passed matrix is on the left, and **MatrixOrderAppend** specifies that the passed matrix is on the right. The default value is **MatrixOrderPrepend**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.


# <a name="SetClip"></a>SetClip (CGpGraphics)

Updates the clipping region of this **Graphics** object to a region that is the combination of itself and the clipping region of another **Graphics** object.

```
FUNCTION SetClip (BYVAL pGraphics AS CGpGraphics PTR, _
   BYVAL nCombineMode AS CombineMode = CombineModeReplace) AS GpStatus
FUNCTION SetClip (BYVAL rc AS GpRectF PTR, _
   BYVAL nCombineMode AS CombineMode = CombineModeReplace) AS GpStatus
FUNCTION SetClip (BYVAL rc AS GpRect PTR, _
   BYVAL nCombineMode AS CombineMode = CombineModeReplace) AS GpStatus
FUNCTION SetClip (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE, _
   BYVAL nCombineMode AS CombineMode = CombineModeReplace) AS GpStatus
FUNCTION SetClip (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, _
   BYVAL nCombineMode AS CombineMode = CombineModeReplace) AS GpStatus
FUNCTION SetClip (BYVAL pPath AS CGpGraphicsPath PTR, _
   BYVAL nCombineMode AS CombineMode = CombineModeReplace) AS GpStatus
FUNCTION SetClip (BYVAL pRegion AS CGpRegion PTR, _
   BYVAL nCombineMode AS CombineMode = CombineModeReplace) AS GpStatus
FUNCTION SetClip (BYVAL hRgn AS HRGN, _
   BYVAL nCombineMode AS CombineMode = CombineModeReplace) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pGraphics* | Pointer to a **Graphics** object that contains the clipping region to be combined with the clipping region of this **Graphics** object. |
| *combineMode* | Optional. Element of the **CombineMode** enumeration that specifies how the specified region is combined with the clipping region of this **Graphics** object. The default value is **CombineModeReplace**. |
| *hRgn* | Handle to a GDI region to be combined with the clipping region of this **Graphics** object. This is provided for legacy code. New applications should pass a **Region** object as the first parameter. |
| *pRegion* | Pointer to a **Region** object that specifies the region to be combined with the clipping region of this **Graphics** object. |
| *pPath* | Pointer to a **GraphicsPath** object that specifies the region to be combined with the clipping region of this **Graphics** object. |
| *rc* | Reference to a rectangle to be combined with the clipping region of this **Graphics** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example updates the clipping region with a Rect structure and then fills a rectangle.
' ========================================================================================
SUB Example_SetClip (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Rect object.
   DIM clipRect AS GpRectF = GDIP_RECTF(0, 0, 200, 100)

   ' // Set the clipping region
   graphics.SetClip(@clipRect)

   ' // Fill a rectangle to demonstrate the clipping region.
   graphics.FillRectangle(@CGpSolidBrush(GDIP_ARGB(255, 0, 0, 0)), 0, 0, 500, 500)

END SUB
' ========================================================================================
```

# <a name="SetCompositingMode"></a>SetCompositingMode (CGpGraphics)

Sets the compositing mode of this **Graphics** object.

```
FUNCTION SetCompositingMode (BYVAL nCompositingMode AS CompositingMode) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nCompositingMode* | Element of the **CompositingMode** enumeration that specifies the compositing mode. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Graphics object and sets its compositing mode to
' CompositingModeSourceCopy. The code creates a SolidBrush object based on a color with an
' alpha component of 128. The code passes the address of that brush to the Graphics.FillRectangle
' method of the Graphics object to fill a rectangle with a color that is not blended with
' the background color. The call to the Graphics::CompositingMode method of the Graphics
' object demonstrates how to obtain the compositing mode (which is already known in this
' case). The code determines whether the compositing mode is CompositingModeSourceCopy and
' if so, changes it to CompositingModeSourceOver. Then the code calls Graphics.FillRectangle
' a second time to fill a rectangle with a color that is a half-and-half blend of the brush
' color and the background color.
' ========================================================================================
SUB Example_CompositingMode (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   graphics.SetCompositingMode(CompositingModeSourceCopy)
   DIM alphaBrush AS CGpSolidBrush = GDIP_ARGB(128, 255, 0, 0)
   graphics.FillRectangle(@alphaBrush, 0, 0, 100, 100)

   ' // Get the compositing mode.
   DIM compMode AS CompositingMode
   compMode = graphics.GetCompositingMode

   ' // Change the compositing mode if it is CompositingModeSourceCopy.
   IF compMode = CompositingModeSourceCopy THEN
      graphics.SetCompositingMode(CompositingModeSourceOver)
   END IF

   graphics.FillRectangle(@alphaBrush, 0, 100, 100, 100)

END SUB
' ========================================================================================
```


# <a name="SetCompositingQuality"></a>SetCompositingQuality (CGpGraphics)

Sets the compositing quality of this **Graphics** object.

```
FUNCTION SetCompositingQuality (BYVAL nQuality AS CompositingQuality) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nQuality* | Element of the **CompositingQuality** enumeration that specifies the compositing quality. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Graphics object and sets its compositing quality to
' CompositingQualityHighQuality. The code creates a SolidBrush object based on a color
' with an alpha component of 128. The code passes the address of that brush to the
' Graphics.FillRectangle method of the Graphics object. The call to the Graphics.CompositingQuality
' method of the Graphics object demonstrates how to obtain the compositing quality (which
' is already known in this case). The code determines whether the compositing quality is
' CompositingQualityHighQuality and if so, changes it to CompositingQualityHighSpeed.
' Then the code calls the Graphics.FillRectangle method a second time. The second rectangle
' is filled with the same brush that was used to fill the first rectangle, but the result
' is different because of the compositing quality setting.
' ========================================================================================
SUB Example_CompositingQuality (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   graphics.SetCompositingQuality(CompositingQualityHighQuality)
   DIM alphaBrush AS CGpSolidBrush = GDIP_ARGB(128, 255, 0, 0)
   graphics.FillRectangle(@alphaBrush, 0, 0, 100, 100)

   ' // Get the compositing mode.
   DIM compQuality AS CompositingMode
   compQuality = graphics.GetCompositingQuality

   ' // Change the compositing quality if it is CompositingQualityHighQuality
   IF compQuality = CompositingQualityHighQuality THEN
      graphics.SetCompositingQuality(CompositingQualityHighSpeed)
   END IF

   graphics.FillRectangle(@alphaBrush, 0, 100, 100, 100)

END SUB
' ========================================================================================
```


# <a name="SetInterpolationMode"></a>SetInterpolationMode (CGpGraphics)

Sets the interpolation mode of this **Graphics** object. The interpolation mode determines the algorithm that is used when images are scaled or rotated.

```
FUNCTION SetInterpolationMode (BYVAL interpolationMode AS LONG) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *interpolationMode* | Element of the **InterpolationMode** enumeration that specifies the interpolation mode. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.


# <a name="SetPageScale"></a>SetPageScale (CGpGraphics)

Sets the scaling factor for the page transformation of this **Graphics** object. The page transformation converts page coordinates to device coordinates.

```
FUNCTION SetPageScale (BYVAL nScale AS SINGLE) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nScale* | The scaling factor. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example sets the world transformation and the page transformation of a
' Graphics object. The page unit and the page scale both belong to the page transformation.
' The code sets the page unit to millimeters and sets the page scale to 10. The call to the
' Graphics.DrawRectangle method draws a rectangle that has a width of 3 centimeters
' (UnitMillimeter along with a scaling factor of 10) and a height of 2 centimeters. The 
' rectangle is translated 4 centimeters to the right and 1 centimeter down by the world
' transformation.
' ========================================================================================
SUB Example_PageScale (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc

   '// Set the world transformation of the Graphics object.
   graphics.TranslateTransform(4, 1)

   ' // Set the page transformation of the Graphics object.
   graphics.SetPageUnit(UnitMillimeter)
   graphics.SetPageScale(10)

   DIM blackPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 0), 0)
   graphics.DrawRectangle(@blackPen, 0, 0, 3, 2)

END SUB
' ========================================================================================
```


# <a name="SetPageUnit"></a>SetPageUnit (CGpGraphics)

Sets the unit of measure for this **Graphics** object. The page unit belongs to the page transformation, which converts page coordinates to device coordinates.

```
FUNCTION SetPageUnit (BYVAL unit AS GpUnit) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *unit* | Element of the **GpUnit** enumeration that specifies the unit of measure for this **Graphics** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example draws two rectangles: one measured in pixels and one measured in inches.
' ========================================================================================
SUB Example_SetPageUnit (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96

   ' // Set the page units to pixels, and draw a rectangle.
   graphics.SetPageUnit(UnitPixel)
   DIM blackPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 0), 0.0)
   graphics.DrawRectangle(@blackPen, 0, 0, 100 * rxRatio, 100 * rxRatio)

   ' // Set the page units to inches, and draw a rectangle.
   graphics.SetPageUnit(UnitInch)
   DIM bluePen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 0.0)
   graphics.DrawRectangle(@bluePen, 2, 0, 1, 1)

END SUB
' ========================================================================================
```


# <a name="SetPixelOffsetMode"></a>SetPixelOffsetMode (CGpGraphics)

Sets the pixel offset mode of this **Graphics** object.

```
FUNCTION SetPixelOffsetMode (BYVAL nMode AS PixelOffsetMode) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nMode* | Element of the **PixelOffsetMode** enumeration that specifies the pixel offset mode. |

#### PixelOffsetMode Enumeration

The **PixelOffsetMode** enumeration specifies the pixel offset mode. This enumeration is used by the **GetPixelOffsetMode** and **SetPixelOffsetMode** methods of the Graphics class.

| Cpmstant   | Description |
| ---------- | ----------- |
| **PixelOffsetModeInvalid** | Used internally.  |
| **PixelOffsetModeDefault** | Equivalent to **PixelOffsetModeNone**. |
| **PixelOffsetModeHighSpeed** | Equivalent to **PixelOffsetModeNone**. |
| **PixelOffsetModeHighQuality** | Equivalent to **PixelOffsetModeHalf**. |
| **PixelOffsetModeNone** | Indicates that pixel centers have integer coordinates. |
| **PixelOffsetModeHalf** | Indicates that pixel centers have coordinates that are half way between integer values. |

#### Remarks

Consider the pixel in the upper-left corner of an image with address (0, 0). With **PixelOffsetModeNone**, the pixel covers the area between 0.5 and 0.5 in both the x and y directions; that is, the pixel center is at (0, 0). With **PixelOffsetModeHalf**, the pixel covers the area between 0 and 1 in both the x and y directions; that is, the pixel center is at (0.5, 0.5).

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.


# <a name="SetRenderingOrigin"></a>SetRenderingOrigin (CGpGraphics)

Sets the rendering origin of this **Graphics** object. The rendering origin is used to set the dither origin for 8-bits-per-pixel and 16-bits-per-pixel dithering and is also used to set the origin for hatch brushes.

```
FUNCTION SetRenderingOrigin (BYVAL x AS INT_, BYVAL y AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | Integer that specifies the x-coordinate of the rendering origin. |
| *y* | Integer that specifies the y-coordinate of the rendering origin. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example sets the text-rendering hint to low quality and draws text. It then
' gets the text-rendering hint, changes it to high quality, and draws more text to
' demonstrate the difference.
' ========================================================================================
SUB Example_SetRenderingOrigin (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM hatchBrush AS CGpHatchBrush = CGpHatchBrush(HatchStyleDiagonalCross, GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 255, 255))
   graphics.FillRectangle(@hatchBrush, 0, 0, 100, 50)
   graphics.SetRenderingOrigin(3, 0)
   graphics.FillRectangle(@hatchBrush, 0, 50, 100, 50)

END SUB
' ========================================================================================
```


# <a name="SetSmoothingMode"></a>SetSmoothingMode (CGpGraphics)

Sets the rendering quality of the **Graphics** object.

```
FUNCTION SetSmoothingMode (BYVAL smoothingMode AS LONG) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *smoothingMode* | Element of the **SmoothingMode** enumeration that specifies whether smoothing (antialiasing) is applied to lines and curves. |

#### SmoothingMode Enumeration

The **SmoothingMode** enumeration specifies the type of smoothing (antialiasing) that is applied to lines and curves. This enumeration is used by the **GetSmoothingMode** and **SetSmoothingMode** functions.

| Constant   | Description |
| ---------- | ----------- |
| **SmoothingModeInvalid** | Reserved. |
| **SmoothingModeDefault** | Specifies that smoothing is not applied. |
| **SmoothingModeHighSpeed** | Specifies that smoothing is not applied. |
| **SmoothingModeHighQuality** | Specifies that smoothing is applied using an 8 X 4 box filter. |
| **SmoothingModeNone** | Specifies that smoothing is not applied. |
| **SmoothingModeAntiAlias8x4** | Specifies that smoothing is applied using an 8 X 4 box filter. |
| **SmoothingModeAntiAlias** | Specifies that smoothing is applied using an 8 X 4 box filter. |
| **SmoothingModeAntiAlias8x8** | Specifies that smoothing is applied using an 8 X 8 box filter. |

#### Remarks

Smoothing performed by an 8 X 4 box filter gives better results for nearly vertical lines than it does for nearly horizontal lines. Smoothing performed by an 8 X 8 box filter gives equally good results for nearly vertical and nearly horizontal lines. The 8x8 algorithm produces higher quality smoothing but is slower than the 8 X 4 algorithm.


#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example sets the smoothing mode to high speed and draws an ellipse. It then
' gets the smoothing mode, changes it to high quality, and draws a second ellipse to
' demonstrate the difference.
' ========================================================================================
SUB Example_SmoothingMode (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Set the smoothing mode to SmoothingModeHighSpeed.
   graphics.SetSmoothingMode(SmoothingModeHighSpeed)

   ' // Draw an ellipse.
   graphics.DrawEllipse(@CGpPen(GDIP_ARGB(255, 0, 0, 0), 3), 10, 50, 200, 100)

   ' // Get the smoothing mode.
   DIM nMode AS SmoothingMode = graphics.GetSmoothingMode

   ' // Test mode to see whether smoothing has been set for the Graphics object.
   IF nMode <> SmoothingModeHighQuality THEN
      graphics.SetSmoothingMode(SmoothingModeHighQuality)
   END IF

   ' // Draw an ellipse to demonstrate the difference.
   graphics.DrawEllipse(@CGpPen(GDIP_ARGB(255, 255, 0, 0), 3), 220, 50, 200, 100)

END SUB
' ========================================================================================
```


# <a name="SetTextContrast"></a>SetTextContrast (CGpGraphics)

Sets the contrast value of this **Graphics** object. The contrast value is used for antialiasing text.

```
FUNCTION SetTextContrast (BYVAL contrast AS UINT) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *contrast* | The contrast value for antialiasing text. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.


# <a name="SetTextRenderingHint"></a>SetTextRenderingHint (CGpGraphics)

Sets the text rendering mode of this **Graphics** object.

```
FUNCTION SetTextRenderingHint (BYVAL newMode AS TextRenderingHint) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *newMode* | Element of the **TextRenderingHint** enumeration that specifies the process currently used by this **Graphics** object to render text. |


#### TextRenderingHint Enumeration

Specifies the process used to render text. The process affects the quality of the text.

| Constant   | Description |
| ---------- | ----------- |
| **TextRenderingHintSystemDefault** | Specifies that a character is drawn using the currently selected system font smoothing mode (also called a rendering hint). |
| **TextRenderingHintSingleBitPerPixelGridFit** | Specifies that a character is drawn using its glyph bitmap and hinting to improve character appearance on stems and curvature. |
| **TextRenderingHintSingleBitPerPixel** | Specifies that a character is drawn using its glyph bitmap and no hinting. This results in better performance at the expense of quality. |
| **TextRenderingHintAntiAliasGridFit** | Specifies that a character is drawn using its antialiased glyph bitmap and hinting. This results in much better quality due to antialiasing at a higher performance cost. |
| **TextRenderingHintAntiAlias**| Specifies that a character is drawn using its antialiased glyph bitmap and no hinting. Stem width differences may be noticeable because hinting is turned off. |
| **TextRenderingHintClearTypeGridFit** | Specifies that a character is drawn using its glyph Microsoft ClearType bitmap and hinting. This type of text rendering cannot be used along with **CompositingModeSourceCopy**.<br>Microsoft Windows XP and Windows Server 2003 only: ClearType rendering is supported only on Windows XP and Windows Server 2003. Therefore, **TextRenderingHintClearTypeGridFit** is ignored on other operating systems even though Windows GDI+ is supported on those operating systems. |

#### Remarks

The quality associated with each process varies according to the circumstances. **TextRenderingHintClearTypeGridFit** provides the best quality for most LCD monitors and relatively small font sizes. **TextRenderingHintAntiAlias** provides the best quality for rotated text. Generally, a process that produces higher quality text is slower than a process that produces lower quality text.

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.


# <a name="SetTransform"></a>SetTransform (CGpGraphics)

Sets the world transformation of this **Graphics** object.

```
FUNCTION SetTransform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMatrix* | Pointer to a **Matrix** object that specifies the world transformation. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a rotation matrix and passes the address of that matrix to
' the Graphics.SetTransform method of a Graphics object. The code calls the Graphics.DrawRectangle
' method of the Graphics object to draw a rotated rectangle.
' ========================================================================================
SUB Example_SetTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96

   ' // Create a rotation matrix.
   DIM transformMatrix AS CGpMatrix
   transformMatrix.Rotate(45.0!)
   ' // Adjust to DPI
   transformMatrix.Scale(rxRatio, rxRatio)

   ' // Set the transformation matrix of the Graphics object.
   graphics.SetTransform(@transformMatrix)

   ' // Draw a rotated rectangle.
   DIM blackPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 0))
   graphics.DrawRectangle(@blackPen, 120, 0, 100, 50)

END SUB
' ========================================================================================
```


# <a name="TransformPoints"></a>TransformPoints (CGpGraphics)

Converts an array of points from one coordinate space to another. The conversion is based on the current world and page transformations of this **Graphics** object.

```
FUNCTION TransformPoints (BYVAL destSpace AS CoordinateSpace, BYVAL srcSpace AS CoordinateSpace, _
   BYVAL pts AS GpPointF PTR, BYVAL count AS LONG) AS GpStatus
FUNCTION TransformPoints (BYVAL destSpace AS CoordinateSpace, BYVAL srcSpace AS CoordinateSpace, _
   BYVAL pts AS GpPoint PTR, BYVAL count AS LONG) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *destSpace* | Element of the **CoordinateSpace** enumeration that specifies the destination coordinate space. |
| *srcSpace* | Element of the **CoordinateSpace** enumeration that specifies the source coordinate space. |
| *pts* | Pointer to an array that, on input, holds the points to be converted and, on output, holds the converted points. |
| *count* | Integer that specifies the number of elements in the *pts* array. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Graphics object and sets its world transformation to a
' translation 40 units right and 30 units down. Then the code creates an array of points
' and passes the address of that array to the Graphics::TransformPoints method of the same
' Graphics object. The points in the array are transformed by the world transformation of
' the Graphics object. The code calls the Graphics.DrawLine method twice: once to connect
' the two points before the transformation and once to connect the two points after the
' transformation.
' ========================================================================================
SUB Example_TransformPoints (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

'   Pen pen(Color(255, 0, 0, 255));
   DIM bluePen AS CGpPen = GDIP_ARGB(255, 0, 0, 255)

   ' // Create an array of two Point objects.
   DIM rgPoints(0 TO 1) AS GpPoint
   rgPoints(1).x = 100 : rgPoints(1).y = 50

   ' // Draw a line that connects the two points.
   ' // No transformation has been performed yet.
   graphics.DrawLine(@bluePen, @rgPoints(0), @rgPoints(1))

   ' // Set the world transformation of the Graphics object.
   graphics.TranslateTransform(40, 30)

   ' // Transform the points in the array from world to page coordinates.
   graphics.TransformPoints(CoordinateSpacePage, CoordinateSpaceWorld, @rgPoints(0), 2)

   ' // It is the world transformation that takes points from world
   ' // space to page space. Because the world transformation is a
   ' // translation 40 to the right and 30 down, the
   ' // points in the array are now (40, 30) and (140, 80).

   ' // Draw a line that connects the transformed points.
   graphics.ResetTransform
   DIM bluePen2 AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), rxRatio)
   graphics.DrawLine(@bluePen2, @rgPoints(0), @rgPoints(1))

END SUB
' ========================================================================================
```


# <a name="TranslateClip"></a>TranslateClip (CGpGraphics)

Translates the clipping region of this **Graphics** object.

```
FUNCTION TranslateClip (BYVAL dx AS SINGLE, BYVAL dy AS SINGLE) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *dx* | The horizontal component of the translation. |
| *dy* | The vertical component of the translation. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Graphics object and sets its clipping region. The code
' translates the clipping region 100 units to the right and then fills a large rectangle
' that is clipped by the translated region.
' ========================================================================================
SUB Example_TranslateClip (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Set the clipping region.
   graphics.SetClip(0, 0, 100, 50)

   ' // Translate the clipping region.
   graphics.TranslateClip(40, 30)

   ' // Fill an ellipse that is clipped by the translated clipping region.
   DIM redBrush AS CGpSolidBrush = CGpSolidBrush(GDIP_ARGB(255, 255, 0, 0))
   graphics.FillEllipse(@redBrush, 20, 40, 100, 80)

   ' // Draw the outline of the clipping region (rectangle).
   DIM blackPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 0), 2)
   graphics.DrawRectangle(@blackPen, 40, 30, 100, 50)

END SUB
' ========================================================================================
```


# <a name="TranslateTransform"></a>TranslateTransform (CGpGraphics)

Updates this **Graphics** object's world transformation matrix with the product of itself and a translation matrix.

```
FUNCTION TranslateTransform (BYVAL dx AS SINGLE, BYVAL dy AS SINGLE, _
   BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *dx* | The horizontal component of the translation. |
| *dy* | The vertical component of the translation. |
| *order* | Optional. Element of the MatrixOrder enumeration that specifies the order of multiplication. **MatrixOrderPrepend** specifies that the passed matrix is on the left, and **MatrixOrderAppend** specifies that the passed matrix is on the right. The default value is **MatrixOrderPrepend**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

Not implemented in the C++ classes.

#### Example

```
' ========================================================================================
' The following example sets the world transformation of a Graphics object to a rotation.
' The call to Graphics.TranslateTransform multiplies the Graphics object's existing world
' transformation matrix (rotation) by a translation matrix. The MatrixOrderAppend argument
' specifies that the multiplication is done with the translation matrix on the right. At
' that point, the world transformation matrix of the Graphics object represents a composite
' transformation: first rotate, then translate. The call to DrawEllipse draws a rotated
' and translated ellipse.
' ========================================================================================
SUB Example_TranslateTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM redPen AS CGpPen = CGpPen(GDIP_ARGB(255, 255, 0, 0))
   graphics.RotateTransform(30)
   graphics.TranslateTransform(100, 50, MatrixOrderAppend)
   graphics.DrawEllipse(@redPen, 0, 0, 200, 80)

END SUB
' ========================================================================================
```

# <a name="ConstructorsGraphicsPath"></a>Constructors (CGpGraphicsPath)

Creates a **GraphicsPath** object.

```
CONSTRUCTOR GraphicsPath (OPTIONAL BYVAL fillMode AS LONG) AS LONG
CONSTRUCTOR CGpGraphicsPath (BYVAL pts AS GpPointF PTR, _
   BYVAL types AS BYTE PTR, BYVAL nCount AS LONG, BYVAL nFillMode AS FillMode = FillModeAlternate)
CONSTRUCTOR CGpGraphicsPath (BYVAL pts AS GpPoint PTR, _
  BYVAL types AS BYTE PTR, BYVAL nCount AS LONG, BYVAL nFillMode AS FillMode = FillModeAlternate)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pts* | Pointer to an array of points that specifies the endpoints and control points of the lines and Bézier splines that are used to draw the path. |
| *types* | Pointer to an array of bytes that holds the point type and a set of flags for each point in the points array. The point type is stored in the three least significant bits, and the flags are stored in the four most significant bits. Possible point types and flags are listed in the PathPointType enumeration. |
| *nCount* | Integer that specifies the number of elements in the points array. This is the same as the number of elements in the types array. |
| *nCount* | Optional. Element of the **FillMode** enumeration that specifies how areas are filled if the path intersects itself. The default value is **FillModeAlternate**. |
| *fillMode* | Optional. Element of the **FillMode** enumeration that specifies how areas are filled if the path intersects itself. The default value is **FillModeAlternate**. |


# <a name="AddArc"></a>AddArc (CGpGraphicsPath)

Adds an elliptical arc to the current figure of this path.

```
FUNCTION AddArc (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE, _
   BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
FUNCTION AddArc (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, _
   BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
FUNCTION AddArc (BYVAL rc AS GpRectF, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
FUNCTION AddArc (BYVAL rc AS GpRect, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | The x-coordinate of the upper-left corner of the bounding rectangle for the ellipse that contains the arc. |
| *y* | The y-coordinate of the upper-left corner of the bounding rectangle for the ellipse that contains the arc. |
| *nWidth* | The width of the bounding rectangle for the ellipse that contains the arc. |
| *nHeight* | The height of the bounding rectangle for the ellipse that contains the arc. |
| *startAngle* | The clockwise angle, in degrees, between the horizontal axis of the ellipse and the starting point of the arc. |
| *sweepAngle* | The clockwise angle, in degrees, between the starting point (startAngle) and the ending point of the arc. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object path, adds an arc to path, closes
' the arc, and then draws path.
' ========================================================================================
SUB Example_AddArc (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   DIM rc AS GpRect = GDIP_RECT(20, 20, 50, 100)

   DIM path AS CGpGraphicsPath
   path.AddArc(@rc, 0.0, 180.0)
   path.CloseFigure
   
   ' // Draw the path.
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawPath(@pen, @path)

END SUB
' ========================================================================================
```


# <a name="AddBezier"></a>AddBezier (CGpGraphicsPath)

Adds a Bézier spline to the current figure of this path.

```
FUNCTION AddBezier (BYVAL x1 AS SINGLE, BYVAL y1 AS SINGLE, BYVAL x2 AS SINGLE, BYVAL y2 AS SINGLE, _
   BYVAL x3 AS SINGLE, BYVAL y3 AS SINGLE, BYVAL x4 AS SINGLE, BYVAL y4 AS SINGLE) AS GpStatus
FUNCTION AddBezier (BYVAL x1 AS INT_, BYVAL y1 AS INT_, BYVAL x2 AS INT_, BYVAL y2 AS INT_, _
   BYVAL x3 AS INT_, BYVAL y3 AS INT_, BYVAL x4 AS INT_, BYVAL y4 AS INT_) AS GpStatus
FUNCTION AddBezier (BYVAL pt1 AS GpPointF, BYVAL pt2 AS GpPointF) AS GpStatus
FUNCTION AddBezier (BYVAL pt1 AS GpPoint, BYVAL pt2 AS GpPoint) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | The x-coordinate of the starting point of the Bézier spline. |
| *y* | The y-coordinate of the starting point of the Bézier spline. |
| *x2* | The x-coordinate of the first control point of the Bézier spline. |
| *y2* | The y-coordinate of the first control point of the Bézier spline. |
| *x3* | The x-coordinate of the second control point of the Bézier spline. |
| *y3* | The y-coordinate of the second control point of the Bézier spline. |
| *x4* | The x-coordinate of the ending point of the Bézier spline. |
| *y4* | The y-coordinate of the ending point of the Bézier spline. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.


# <a name="AddBeziers"></a>AddBeziers (CGpGraphicsPath)

Adds a sequence of connected Bézier splines to the current figure of this path.

```
FUNCTION AddBeziers (BYVAL pts AS GpPointF PTR, BYVAL count AS INT_) AS GpStatus
FUNCTION AddBeziers (BYVAL pts AS GpPoint PTR, BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pts* | Pointer to an array of starting points, ending points, and control points for the connected splines. The first spline is constructed from the first point through the fourth point in the array and uses the second and third points as control points. Each subsequent spline in the sequence needs exactly three more points: the ending point of the previous spline is used as the starting point, the next two points in the sequence are control points, and the third point is the ending point. |
| *count* | Integer that specifies the number of elements in the *pts* array. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object path, adds a sequence of two connected
' Bézier splines to path, closes the current figure (the only figure in this case), and
' then draws path.
' ========================================================================================
SUB Example_AddBeziers (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   DIM pts(0 TO 6) AS GpPoint = {GDIP_POINT(50, 50), GDIP_POINT(60, 20), GDIP_POINT(70, 100), GDIP_POINT(80, 50), GDIP_POINT(120, 40), GDIP_POINT(150, 80), GDIP_POINT(170, 30)}
   DIM path AS CGpGraphicsPath
   path.AddBeziers(@pts(0), 7)
   path.CloseFigure
   
   ' // Draw the path
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawPath(@pen, @path)
   
END SUB
' ========================================================================================
```


# <a name="AddClosedCurve"></a>AddClosedCurve (CGpGraphicsPath)

Adds a closed cardinal spline to this path.

```
FUNCTION AddClosedCurve (BYVAL pts AS GpPointF PTR, BYVAL count AS INT_) AS GpStatus
FUNCTION AddClosedCurve (BYVAL pts AS GpPoint PTR, BYVAL count AS INT_) AS GpStatus
FUNCTION AddClosedCurve (BYVAL pts AS GpPointF PTR, BYVAL count AS INT_, BYVAL tension AS SINGLE) AS GpStatus
FUNCTION AddClosedCurve (BYVAL pts AS GpPoint PTR, BYVAL count AS INT_, BYVAL tension AS SINGLE) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pts* | Pointer to an array of points that define the cardinal spline. The cardinal spline is a curve that passes through each point in the array. |
| *count* | Integer that specifies the number of elements in the *pts* array. |
| *tension* | Nonnegative real number that controls the length of the curve and how the curve bends. A value of 0 specifies that the spline is a sequence of straight line segments. As the value increases, the curve becomes fuller. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

You should keep a copy of the points array if those points will be needed later. The **GraphicsPath** object does not store the points passed to the **AddClosedCurve** method; instead, it converts the cardinal spline to a sequence of Bézier splines and stores the points that define those Bézier splines. You cannot retrieve the original array of points from the **GraphicsPath** object.

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object path, adds a closed cardinal spline
' to path, and then draws path.
' ========================================================================================
SUB Example_AddClosedCurve (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   DIM pts(0 TO 3) AS GpPoint = {GDIP_POINT(50, 50), GDIP_POINT(60, 20), GDIP_POINT(70, 100), GDIP_POINT(80, 50)}
   DIM path AS CGpGraphicsPath
   path.AddClosedCurve(@pts(0), 4)

   ' // Draw the path
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawPath(@pen, @path)

END SUB
' ========================================================================================
```


# <a name="AddCurve"></a>AddCurve (CGpGraphicsPath)

Adds a cardinal spline to this path.

```
FUNCTION AddCurve (BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_) AS GpStatus
FUNCTION AddCurve (BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_) AS GpStatus
FUNCTION AddCurve (BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_, BYVAL tension AS SINGLE) AS GpStatus
FUNCTION AddCurve (BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_, BYVAL tension AS SINGLE) AS GpStatus
FUNCTION AddCurve (BYVAL pts AS GpPointF PTR, BYVAL nCount AS INT_, BYVAL offset AS INT_, _
   BYVAL numberOfSegments AS INT_, BYVAL tension AS SINGLE) AS GpStatus
FUNCTION AddCurve (BYVAL pts AS GpPoint PTR, BYVAL nCount AS INT_, BYVAL offset AS INT_, _
   BYVAL numberOfSegments AS INT_, BYVAL tension AS SINGLE) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pts* | Pointer to an array of points that define the cardinal spline. The cardinal spline is a curve that passes through each point in the array. |
| *count* | Integer that specifies the number of elements in the *pts* array. |
| *offset* | Integer that specifies the index of the array element that is used as the first point of the cardinal spline. |
| *numberOfSegments* | Integer that specifies the number of segments in the cardinal spline. Segments are the curves that connect consecutive points in the array. |
| *tension* | Nonnegative single precision number that controls the length of the curve and how the curve bends. A value of 0 specifies that the spline is a sequence of straight line segments. As the value increases, the curve becomes fuller. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

```
' ========================================================================================
' The following example creates a GraphicsPath object path, adds a cardinal spline to path,
' and then draws path.
' ========================================================================================
SUB Example_AddCurve (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   DIM pts(0 TO 3) AS GpPoint = {GDIP_POINT(50, 50), GDIP_POINT(60, 20), GDIP_POINT(70, 100), GDIP_POINT(80, 50)}
   DIM path AS CGpGraphicsPath
   path.AddCurve(@pts(0), 4)

   ' // Draw the path
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawPath(@pen, @path)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object path, adds a cardinal spline to path,
' and then draws path.
' ========================================================================================
SUB Example_AddCurve (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   DIM path AS CGpGraphicsPath
   DIM pts(0 TO 7) AS GpPoint = {GDIP_POINT(50, 50), GDIP_POINT(70, 80), GDIP_POINT(100, 100), _
      GDIP_POINT(130, 40), GDIP_POINT(150, 90), GDIP_POINT(180, 30), GDIP_POINT(210, 120), GDIP_POINT(240, 80)}
   path.AddCurve(@pts(0), 8, 2, 4, 1.0)
   
   DIM pen AS CGpPen = GDIP_ARGB(255, 0, 0, 255)
   graphics.DrawPath(@pen, @path)

   ' // Draw all eight points in the array.
   DIM brush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   
   FOR j AS LONG = 0 TO 7
      graphics.FillEllipse(@brush, pts(j).x - 3, pts(j).y - 3, 6, 6)
   NEXT

END SUB
' ========================================================================================
```


# <a name="AddEllipse"></a>AddEllipse (CGpGraphicsPath)

Adds an ellipse to this path.

```
FUNCTION AddEllipse (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, _
   BYVAL nHeight AS SINGLE) AS GpStatus
FUNCTION AddEllipse (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, _
   BYVAL nHeight AS INT_) AS GpStatus
FUNCTION AddEllipse (BYVAL rc AS GpRectF PTR) AS GpStatus
FUNCTION AddEllipse (BYVAL rc AS GpRect PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | The x-coordinate of the upper-left corner of the bounding rectangle for the ellipse. |
| *y* | The y-coordinate of the upper-left corner of the bounding rectangle for the ellipse. |
| *nWidth* | The width of the bounding rectangle for the ellipse. |
| *nHeight* | The height of the bounding rectangle for the ellipse. |
| *rc* | Pointer to a **GpRectF** or **GpRect** structure specifying the dimensions of the rectagle. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object path, adds an ellipse to path, and
' then draws path.
' ========================================================================================
SUB Example_AddEllipse (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   DIM path AS CGpGraphicsPath
   path.AddEllipse(20, 20, 200, 100)

   ' // Draw the path
   DIM pen AS CgpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawPath(@pen, @path)

END SUB
' ========================================================================================
```


# <a name="AddLine"></a>AddLine (CGpGraphicsPath)

Adds a line to the current figure of this path.

```
FUNCTION AddLine (BYVAL x1 AS SINGLE, BYVAL y1 AS SINGLE, BYVAL x2 AS SINGLE, BYVAL y2 AS SINGLE) AS GpStatus
FUNCTION AddLine (BYVAL x1 AS INT_, BYVAL y1 AS INT_, BYVAL x2 AS INT_, BYVAL y2 AS INT_) AS GpStatus
FUNCTION AddLine (BYVAL pt1 AS GpPointF PTR, BYVAL pt2 AS GpPointF PTR) AS GpStatus
FUNCTION AddLine (BYVAL pt1 AS GpPoint PTR, BYVAL pt2 AS GpPoint PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *x1* | The x-coordinate of the starting point of the line. |
| *y1* | The y-coordinate of the starting point of the line. |
| *x2* | The x-coordinate of the ending point of the line. |
| *y2* | The y-coordinate of the ending point of the line. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object path, adds a line to path, and then draws path.
' ========================================================================================
SUB Example_AddLine (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   DIM path AS CGpGraphicsPath
   path.AddLine(20, 20, 200, 100)

   ' // Draw the path
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawPath(@pen, @path)

END SUB
' ========================================================================================
```


# <a name="AddLines"></a>AddLines (CGpGraphicsPath)

Adds a sequence of connected lines to the current figure of this path.

```
FUNCTION AddLines (BYVAL pts AS GpPointF PTR, BYVAL count AS INT_) AS GpStatus
FUNCTION AddLines (BYVAL pts AS GpPoint PTR, BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pts* | Pointer to an array of points that specify the starting and ending points of the lines. The first point in the array is the starting point of the first line, and the last point in the array is the ending point of the last line. Each of the other points serves as ending point for one line and starting point for the next line. |
| *count* | Integer that specifies the number of elements in the *pts* array. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object path, adds a sequence of four
' connected lines to path, and then draws path.
' ========================================================================================
SUB Example_AddLines (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   DIM pts(0 TO 4) AS GpPoint = {GDIP_POINT(20, 20), GDIP_POINT(30, 30), GDIP_POINT(40, 24), _
      GDIP_POINT(50, 30), GDIP_POINT(60, 20)}
   DIM path AS CGpGraphicsPath
   path.AddLines(@pts(0), 5)

   ' // Draw the path
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawPath(@pen, @path)

END SUB
' ========================================================================================
```


# <a name="AddPath"></a>AddPath (CGpGraphicsPath)

Adds a path to this path.

```
FUNCTION AddPath (BYVAL pAddingPath AS CGpGraphicsPath PTR, BYVAL bConnect AS BOOL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pAddingPath* | Pointer to the path to be added. |
| *bConnect* | BOOL value that specifies whether the first figure in the added path is part of the last figure in this path.<br>**CTRUE**: Specifies that (if possible) the first figure in the added path is part of the last figure in this path.<br>**FALSE**: Specifies that the first figure in the added path is separate from the last figure in this path. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

Even if the value of the connect parameter is CTRUE, this method might not be able to make the first figure of the added path part of the last figure of this path. If either of those figures is closed, then they must remain separate figures.

#### Example

```
' ========================================================================================
' The following example creates two GraphicsPath objects: path1 and path2. The code adds
' an open figure consisting of an arc and a Bézier spline to each path. The code calls the
' GraphicsPath.AddPath method of path1 to add path2 to path1. The second argument is TRUE,
' which specifies that all four items (two arcs and two Bézier splines) belong to the same
' figure.
' ========================================================================================
SUB Example_AddPath (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   DIM path1 AS CGpGraphicsPath
   path1.AddArc(10, 10, 50, 20, 0.0, 150.0)
   path1.AddBezier(10, 50, 60, 50, 10, 80, 60, 80)
   
   DIM path2 AS CGpGraphicsPath
   path2.AddArc(10, 110, 50, 20, 0.0, 150.0)
   path2.AddBezier(10, 150, 60, 150, 10, 180, 60, 180)

   path1.AddPath(@path2, TRUE)

   DIM pen AS CGpPen = GDIP_ARGB(255, 0, 0, 255)
   graphics.DrawPath(@pen, @path1)

END SUB
' ========================================================================================
```


# <a name="AddPie"></a>AddPie (CGpGraphicsPath)

Adds a pie to this path. An arc is a portion of an ellipse, and a pie is a portion of the area enclosed by an ellipse. A pie is bounded by an arc and two lines (edges) that go from the center of the ellipse to the endpoints of the arc.

```
FUNCTION AddPie (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE, _
   BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
FUNCTION AddPie (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, _
   BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
FUNCTION AddPie (BYVAL rc AS GpRectF PTR, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
FUNCTION AddPie (BYVAL rc AS GpRect PTR, BYVAL startAngle AS SINGLE, BYVAL sweepAngle AS SINGLE) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | The x-coordinate of the upper-left corner of the rectangle that bounds the ellipse that bounds the pie. |
| *y* | The y-coordinate of the upper-left corner of the rectangle that bounds the ellipse that bounds the pie. |
| *nWidth* | The width of the rectangle that bounds the ellipse that bounds the pie. |
| *nHeight* | The height of the rectangle that bounds the ellipse that bounds the pie. |
| *startAngle* | The clockwise angle, in degrees, between the horizontal axis of the ellipse and the starting point of the arc that defines the pie. |
| *sweepAngle* | The clockwise angle, in degrees, between the starting and ending points of the arc that defines the pie. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

Even if the value of the connect parameter is CTRUE, this method might not be able to make the first figure of the added path part of the last figure of this path. If either of those figures is closed, then they must remain separate figures.

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object path, adds a pie to path, and then draws path.
' ========================================================================================
SUB Example_AddPie (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   DIM path AS CGpGraphicsPath
   path.AddPie(50, 50, 100, 100, 20.0, 45.0)

   ' // Draw the path
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawPath(@pen, @path)

END SUB
' ========================================================================================
```


# <a name="AddPolygon"></a>AddPolygon (CGpGraphicsPath)

Adds a polygon to this path.

```
FUNCTION AddPolygon (BYVAL pts AS GpPointF PTR, BYVAL count AS INT_) AS GpStatus
FUNCTION AddPolygon (BYVAL pts AS GpPoint PTR, BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pts* | Pointer to an array of points that specifies the vertices of the polygon. |
| *count* | Integer that specifies the number of elements in the *pts* array. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object path, adds a polygon to path, and then draws path.
' ========================================================================================
SUB Example_AddPolygon (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   DIM pts(0 TO 2) AS GpPoint = {GDIP_POINT(20, 20), GDIP_POINT(120, 20), GDIP_POINT(120, 70)}
   DIM path AS CGpGraphicsPath
   path.AddPolygon(@pts(0), 3)

   ' // Draw the path
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawPath(@pen, @path)

END SUB
' ========================================================================================
```


# <a name="AddRectangle"></a>AddRectangle (CGpGraphicsPath)

Adds a rectangle to this path.

```
FUNCTION AddRectangle (BYVAL x AS SINGLE, BYVAL y AS SINGLE, _
   BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE) AS GpStatus
FUNCTION AddRectangle (BYVAL x AS INT_, BYVAL y AS INT_, _
   BYVAL nWidth AS INT_, BYVAL nHeight AS INT_) AS GpStatus
FUNCTION AddRectangle (BYVAL rc AS GpRectF PTR) AS GpStatus
FUNCTION AddRectangle (BYVAL rc AS GpRect PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | The x-coordinate of the upper-left corner of the rectangle. |
| *y* | The y-coordinate of the upper-left corner of the rectangle. |
| *nWidth* | The width of the rectangle. |
| *nWidth* | The height of the rectangle. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object path, adds a rectangle to path, and then draws path.
' ========================================================================================
SUB Example_AddRectangle (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   DIM rc AS GpRect = GDIP_RECT(20, 20, 100, 50)
   DIM path AS CGpGraphicsPath
   path.AddRectangle(@rc)

   ' // Draw the path
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawPath(@pen, @path)

END SUB
' ========================================================================================
```


# <a name="AddRectangles"></a>AddRectangles (CGpGraphicsPath)

Adds a sequence of rectangles to this path.

```
FUNCTION AddRectangles (BYVAL rects AS GpRectF PTR, BYVAL nCount AS INT_) AS GpStatus
FUNCTION AddRectangles (BYVAL rects AS GpRect PTR, BYVAL nCount AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *rects* | Pointer to an array of rectangles that will be added to the path. |
| *nCount* | Integer that specifies the number of elements in the rects array. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object path, adds two rectangles to path, and then draws path.
' ========================================================================================
SUB Example_AddRectangles (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   DIM rects(0 TO 1) AS GpRect = {GDIP_RECT(20, 20, 100, 50), GDIP_RECT(30, 30, 50, 100)}
   DIM path AS CGpGraphicsPath
   path.AddRectangles(@rects(0), 2)

   ' // Draw the path
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawPath(@pen, @path)

END SUB
' ========================================================================================
```

# <a name="AddString"></a>AddString (CGpGraphicsPath)

Adds the outline of a string to this path.

```
FUNCTION AddString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFamily AS CGpFontFamily PTR, BYVAL style AS INT_, BYVAL emSize AS SINGLE, _
   BYVAL layoutRect AS GpRectF PTR, BYVAL pFormat AS CGpStringFormat PTR) AS GpStatus
FUNCTION AddString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFamily AS CGpFontFamily PTR, BYVAL style AS INT_, BYVAL emSize AS SINGLE, _
   BYVAL layoutRect AS GpRect PTR, BYVAL pFormat AS CGpStringFormat PTR) AS GpStatus
FUNCTION AddString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFamily AS CGpFontFamily PTR, BYVAL style AS INT_, BYVAL emSize AS SINGLE, _
   BYVAL layoutRect AS GpPointF PTR, BYVAL pFormat AS CGpStringFormat PTR) AS GpStatus
FUNCTION AddString (BYVAL pwszString AS WSTRING PTR, BYVAL length AS INT_, _
   BYVAL pFamily AS CGpFontFamily PTR, BYVAL style AS INT_, BYVAL emSize AS SINGLE, _
   BYVAL layoutRect AS GpPoint PTR, BYVAL pFormat AS CGpStringFormat PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszString* | Pointer to a wide-character string. |
| *pFamily* | Pointer to a **FontFamily** object that specifies the font family for the string. |
| *style* | Integer that specifies the style of the typeface. This value must be an element of the **FontStyle** enumeration or the result of a bitwise **OR** applied to two or more of these elements. For example, **FontStyleBold OR FontStyleUnderline OR FontStyleStrikeout** sets the style as a combination of the three styles. |
| *emSize* | The em size, in world units, of the string characters. |
| *layoutRect* | Reference to a **GpRectF** or **GpRect** object that specifies, in world units, the bounding rectangle for the string. |
| *pFormat* | Pointer to a **StringFormat** object that specifies layout information (alignment, trimming, tab stops, and the like) for the string. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

Note that GDI+ does not support PostScript fonts or OpenType fonts which do not have TrueType outlines. 

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object path, adds a NULL-terminated string
' to path, and then draws path.
' ========================================================================================
SUB Example_AddString (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   DIM fontFamily AS CGpFontFamily = "Times New Roman"
   DIM path AS CGpGraphicsPath
   DIM rc AS GpRect = GDIP_RECT(50, 50, 150, 100)
   path.AddString("Hello World", -1, @fontFamily, FontStyleRegular, 48, @rc, NULL)

   ' // Draw the path
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawPath(@pen, @path)

END SUB
' ========================================================================================
```

# <a name="ClearMarkers"></a>ClearMarkers (CGpGraphicsPath)

Clears the markers from this path.

```
FUNCTION ClearMarkers () AS GpStatus
```

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.


# <a name="Clone"></a>Clone (CGpGraphicsPath)

Copies the contents of the existing **GraphicsPath** object into a new **GraphicsPath** object.

```
FUNCTION Clone (BYVAL pCloneGraphicsPath AS CGpGraphicsPath PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pCloneGraphicsPath* | Pointer to a variable that will receive a pointer to the cloned **GraphicsPath** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.


# <a name="CloseAllFigures"></a>CloseAllFigures (CGpGraphicsPath)

Closes all open figures in this path.

```
FUNCTION CloseAllFigures () AS GpStatus
```

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object path, adds two figures to path (arcs),
' closes all figures, and then draws path.
' ========================================================================================
SUB Example_CloseAllFigures (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   DIM rect1 AS GpRect = GDIP_RECT(20, 20, 50, 100)
   DIM rect2 AS GpRect = GDIP_RECT(40, 40, 50, 100)

   DIM path AS CGpGraphicsPath
   path.AddArc(@rect1, 0, 180)   '  // first figure
   path.StartFigure
   path.AddArc(@rect2, 0, 180)   '  // second figure
   path.CloseAllFigures

   ' // Draw the path
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawPath(@pen, @path)

END SUB
' ========================================================================================
```

# <a name="CloseFigure"></a>CloseFigure (CGpGraphicsPath)

Clears the current figure of this path.

```
FUNCTION CloseFigure () AS GpStatus
```

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object path and adds two figures to path.
' The code closes the first figure, but leaves the second figure open.
' ========================================================================================
SUB Example_CloseFigure (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   DIM rect1 AS GpRect = GDIP_RECT(20, 20, 50, 100)
   DIM rect2 AS GpRect = GDIP_RECT(40, 40, 50, 100)

   DIM path AS CGpGraphicsPath
   path.AddArc(@rect1, 0, 180)   ' // first figure
   path.CloseFigure              ' // first figure closed
   path.AddArc(@rect2, 0, 180)   ' // second figure

   ' // Draw the path
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawPath(@pen, @path)

END SUB
' ========================================================================================
```

# <a name="Flatten"></a>Flatten (CGpGraphicsPath)

Applies a transformation to this path and converts each curve in the path to a sequence of connected lines.

```
FUNCTION Flatten (BYVAL pMatrix AS CGpMatrix PTR = NULL, BYVAL flatness AS SINGLE = FlatnessDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMatrix* | Optional. Pointer to a **Matrix** object that specifies the transformation to be applied to the path's data points. The default value is NULL, which specifies that no transformation is to be applied. |
| *flatness* | Optional. The maximum error between the path and its flattened approximation. Reducing the flatness increases the number of line segments in the approximation. The default value is **FlatnessDefault**, which is a constant defined in Gdiplusenums.bi. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object and then adds a curve (cardinal spline),
' an ellipse, and a Bézier spline to the path. The code draws the path in blue, flattens
' the path, and then draws the flattened path in green. The call to GraphicsPath.GetPathData
' gets the data points stored by the flattened path. The code draws each of those points by
' filling a small ellipse with red.
' ========================================================================================
SUB Example_Flatten (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   DIM pts(0 TO 3) AS GpPoint = {GDIP_POINT(20, 50), GDIP_POINT(40, 70), GDIP_POINT(60, 10), GDIP_POINT(80, 50)}
   DIM path AS CGpGraphicsPath
   path.AddCurve(@pts(0), 4)
   path.AddEllipse(20, 100, 150, 80)
   path.AddBezier(20, 200, 20, 250, 50, 210, 100, 260)

   ' // Draw the path before flattening
   DIM pen AS CGpPen = GDIP_ARGB(255, 0, 0, 255)
   graphics.DrawPath(@pen, @path)

   path.Flatten(NULL, 8.0)

   ' // Draw the flattened path
   pen.SetColor(GDIP_ARGB(255, 0, 255, 0))
   graphics.DrawPath(@pen, @path)

   ' // Get the path data from the flattened path.
   DIM pathData AS GDIP_PATHDATA
   path.GetPathData(@pathData)

   ' // Draw the data points of the flattened path
   DIM brush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   FOR j AS LONG = 0 TO pathData.Count - 1
      graphics.FillEllipse(@brush, pathData.Points[j].x - 3.0, _
         pathData.Points[j].y - 3.0, 6.0, 6.0)
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetBounds"></a>GetBounds (CGpGraphicsPath)

Gets a bounding rectangle for this path.

```
FUNCTION GetBounds (BYVAL bounds AS GpRectF PTR, BYVAL pMatrix AS CGpMatrix PTR, _
   BYVAL pPen AS CGpPen PTR) AS GpStatus
FUNCTION GetBounds (BYVAL bounds AS GpRect PTR, BYVAL pMatrix AS CGpMatrix PTR, _
   BYVAL pPen AS CGpPen PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *bounds* | Pointer to a **GpRectF** or **GpRect** structures that receives the bounding rectangle. |
| *pMatrix* | Optional. Pointer to a Matrix object that specifies a transformation to be applied to this path before the bounding rectangle is calculated. This path is not permanently transformed; the transformation is used only during the process of calculating the bounding rectangle. The default value is NULL. |
| *pPen* | Optional. Pointer to a **Pen** object that influences the size of the bounding rectangle. The bounding rectangle received in bounds will be large enough to enclose this path when the path is drawn with the pen specified by this parameter. This ensures that the path is enclosed by the bounding rectangle even if the path is drawn with a wide pen. The default value is NULL. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a path that has one curve and one ellipse. The code draws
' the path with a thick yellow pen and a thin black pen. The GraphicsPath.GetBounds method
' receives the address of the thick yellow pen and calculates a bounding rectangle for the
' path. Then the code draws the bounding rectangle.
' ========================================================================================
SUB Example_GetBounds (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   DIM blackPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 0), 1)
   DIM yellowPen AS CGpPen = CGpPen(GDIP_ARGB(255, 255, 255, 0), 10)
   DIM redPen AS CGpPen = CGpPen(GDIP_ARGB(255, 255, 0, 0), 1)

   DIM pts(0 TO 3) AS GpPoint = {GDIP_POINT(120, 120), GDIP_POINT(200, 130), GDIP_POINT(150, 200), GDIP_POINT(130, 180)}

   ' // Create a path that has one curve and one ellipse.
   DIM path AS CGpGraphicsPath
   path.AddClosedCurve(@pts(0), 4)
   path.AddEllipse(120, 220, 100, 40)

   ' // Draw the path with a thick yellow pen and a thin black pen.
   graphics.DrawPath(@yellowPen, @path)
   graphics.DrawPath(@blackPen, @path)
 
   ' // Get the path's bounding rectangle.
   DIM rc AS GpRect
   path.GetBounds(@rc, NULL, @yellowPen)
   graphics.DrawRectangle(@redPen, @rc)  

END SUB
' ========================================================================================
```

# <a name="GetFillMode"></a>GetFillMode (CGpGraphicsPath)

Gets the fill mode of this path.

```
FUNCTION GetFillMode () AS FillMode
```

#### Return value

This method returns an element of the **FillMode** enumeration.


# <a name="GetLastPoint"></a>GetLastPoint (CGpGraphicsPath)

Gets the ending point of the last figure in this path.

```
FUNCTION GetLastPoint (BYVAL lastPoint AS GpPointF PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *lastPoint* | Pointer to a **GpPointF** sreucture that receives the last point. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.


# <a name="GetPathData"></a>GetPathData (CGpGraphicsPath)

Gets an array of points and an array of point types from this path. Together, these two arrays define the lines, curves, figures, and markers of this path.

```
FUNCTION GetPathData (BYVAL pPathData AS GpPathData PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pPathData* | Pointer to a **PathData** structure that receives the path data. The Points data member of the **PathData** structure receives a pointer to an array of **GpPointF** objects that contains the path points. The Types data member of the **PathData** structure receives a pointer to an array of bytes that contains the point types. The Count data member of the **PathData** structure receives an integer that indicates the number of elements in the Points array. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates and draws a path that has a line, a rectangle, an ellipse,
' and a curve. The code gets the path's points and types by passing the address of a PathData
' structure to the GraphicsPath.GetPathData method. Then the code draws each of the path's data points.
' ========================================================================================
SUB Example_GetPathData (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   DIM points(0 TO 4) AS GpPoint = {GDIP_POINT(200, 200), GDIP_POINT(250, 240), _
      GDIP_POINT(200, 300), GDIP_POINT(300, 310), GDIP_POINT(250, 350)}
   DIM path AS CGpGraphicsPath
   path.AddLine(20, 100, 150, 200)
   path.AddRectangle(40, 30, 80, 60)
   path.AddEllipse(200, 30, 200, 100)
   path.AddCurve(@points(0), 5)

    ' // Draw the path
   DIM pen AS CGpPen = GDIP_ARGB(255, 0, 0, 255)
   graphics.DrawPath(@pen, @path)

   ' // Get the path data
   DIM pathData AS GDIP_PATHDATA
   path.GetPathData(@pathData)

   ' // Draw the path's data points
   DIM brush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   FOR j AS LONG = 0 TO pathData.Count - 1
      graphics.FillEllipse(@brush, pathData.Points[j].x - 3.0, _
         pathData.Points[j].y - 3.0, 6.0, 6.0)
   NEXT

END SUB
' ========================================================================================
```


# <a name="GetPathPoints"></a>GetPathPoints (CGpGraphicsPath)

Gets this path's array of points. The array contains the endpoints and control points of the lines and Bézier splines that are used to draw the path.

```
FUNCTION GetPathPoints (BYREF pts AS GpPointF PTR, BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pts* | Pointer to an array of **GpPointF** objects that receives the data points. You must allocate memory for this array. You can call the **GetPointCount** method to determine the required size of the array. The size, in bytes, should be the return value of **GetPointCount** multiplied by **SIZEOF(GpPointF)**. |

#### Remarks

A **GraphicsPath** object has an array of points and an array of types. Each element in the array of types is a byte that specifies the point type and a set of flags for the corresponding element in the array of points. Possible point types and flags are listed in the **PathPointType** enumeration.

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates and draws a path that has a line, a rectangle, an ellipse,
' and a curve. The code calls the path's GraphicsPath::GetPointCount method to determine
' the number of data points that are stored in the path. The code allocates a buffer large
' enough to receive the array of data points and passes the address of that buffer to the
' GraphicsPath.GetPathPoints method. Finally, the code draws each of the path's data points.
' ========================================================================================
SUB Example_GetPathPoints (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   ' // Create a path that has a line, a rectangle, an ellipse, and a curve.
   DIM path AS CGpGraphicsPath
   DIM points(0 TO 4) AS GpPoint = {GDIP_POINT(200, 200), GDIP_POINT(250, 240), _
      GDIP_POINT(200, 300), GDIP_POINT(300, 310), GDIP_POINT(250, 350)}

   path.AddLine(20, 100, 150, 200)
   path.AddRectangle(40, 30, 80, 60)
   path.AddEllipse(200, 30, 200, 100)
   path.AddCurve(@points(0), 5)

   ' // Draw the path
   DIM pen AS CGpPen = GDIP_ARGB(255, 0, 0, 255)
   graphics.DrawPath(@pen, @path)

   ' // Get the path points.
   DIM nCount AS LONG = path.GetPointCount
   DIM dataPoints(nCount -1) AS GpPointF
   path.GetPathPoints(@dataPoints(0), nCount)

   ' // Draw the path's data points
   DIM brush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   FOR j AS LONG = 0 TO nCount - 1
      graphics.FillEllipse(@brush, dataPoints(j).x - 3.0, _
         dataPoints(j).y - 3.0, 6.0, 6.0)
   NEXT

END SUB
' ========================================================================================
```


# <a name="GetPathTypes"></a>GetPathTypes (CGpGraphicsPath)

Gets this path's array of point types.

```
FUNCTION GetPathTypes (BYVAL types AS BYTE PTR, BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *types* | Pointer to an array that receives the point types. You must allocate memory for this array. You can call the GetPointCount method to determine the required size of the array. |
| *count* | Integer that specifies the number of elements in the types array. Set this parameter equal to the return value of the GetPointCount method. |

#### Remarks

A **GraphicsPath** object has an array of points and an array of types. Each element in the array of types is a byte that specifies the point type and a set of flags for the corresponding element in the array of points. Possible point types and flags are listed in the **PathPointType** enumeration.

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.


# <a name="GetPointCount"></a>GetPointCount (CGpGraphicsPath)

Gets the number of points in this path's array of data points. This is the same as the number of types in the path's array of point types.

```
FUNCTION GetPointCount () AS INT_
```

#### Remarks

A **GraphicsPath** object has an array of points and an array of types. Each element in the array of types is a byte that specifies the point type and a set of flags for the corresponding element in the array of points. Possible point types and flags are listed in the **PathPointType** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a path that has one ellipse and one line. The code calls
' the GraphicsPath.GetPointCount method to determine the number of data points stored in
' the path. Then the code calls the GraphicsPath::GetPathPoints method to retrieve those
' data points. Finally, the code fills a small ellipse at each of the data points.
' ========================================================================================
SUB Example_GetPointCount (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   ' // Create a path that has one ellipse and one line.
   DIM path AS CGpGraphicsPath
   path.AddEllipse(10, 10, 200, 100)
   path.AddLine(220, 120, 300, 160)

   ' // Find out how many data points are stored in the path.
   DIM nCount AS LONG = path.GetPointCount

   ' // Draw the path points
   DIM RedBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   DIM points(nCount - 1) AS GpPointF
   path.GetPathPoints(@points(0), nCount)

   FOR j AS LONG = 0 TO nCount - 1
      graphics.FillEllipse(@redBrush, points(j).x - 3.0, points(j).y - 3.0, 6.0, 6.0)
   NEXT

END SUB
' ========================================================================================
```


# <a name="IsOutlineVisible"></a>IsOutlineVisible (CGpGraphicsPath)

Determines whether a specified point touches the outline of this path when the path is drawn by a specified **Graphics** object and a specified pen.

```
FUNCTION IsOutlineVisible (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL pPen AS CGpPen PTR, _
   BYVAL pGraphics AS CGpGraphics PTR = NULL) AS BOOLEAN
FUNCTION IsOutlineVisible (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL pPen AS CGpPen PTR, _
   BYVAL pGraphics AS CGpGraphics PTR = NULL) AS BOOLEAN
FUNCTION IsOutlineVisible (BYVAL pt AS GpPointF PTR, BYVAL pPen AS CGpPen PTR, _
   BYVAL pGraphics AS CGpGraphics PTR = NULL) AS BOOLEAN
FUNCTION IsOutlineVisible (BYVAL pt AS GpPoint PTR, BYVAL pPen AS CGpPen PTR, _
   BYVAL pGraphics AS CGpGraphics PTR = NULL) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | The x-coordinate of the point to be tested. |
| *y* | The x-coordinate of the point to be tested. |
| *pPen* | Pointer to a Pen object. This method determines whether the test point touches the path outline that would be drawn by this pen. More points will touch an outline drawn by a wide pen than will touch an outline drawn by a narrow pen. |
| *pGraphics* | Optional. Pointer to a **Graphics** object that specifies a world-to-device transformation. If the value of this parameter is NULL, the test is done in world coordinates; otherwise, the test is done in device coordinates. The default value is NULL.  |

#### Return value

If the test point touches the outline of this path, this method returns TRUE; otherwise, it returns FALSE.

#### Example

```
' ========================================================================================
' The following example creates an elliptical path and draws that path with a wide yellow
' pen. Then the code tests each point in an array to see whether the point touches the
' outline (as it would be drawn by the wide yellow pen) of the path. Points that touch the
' outline are painted green, and points that don't touch the outline are painted red.
' ========================================================================================
SUB Example_IsOutlineVisible (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   DIM yellowPen AS CGpPen = CGpPen(GDIP_ARGB(255, 255, 255, 0), 20)
   DIM brush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0,  0)

   ' // Create and draw a path
   DIM path AS CGpGraphicsPath
   path.AddEllipse(50, 50, 200, 100)
   graphics.DrawPath(@yellowPen, @path)
   
   ' // Create an array of three points, and determine whether each
   ' // point in the array touches the outline of the path.
   ' // If a point touches the outline, paint it green.
   ' // If a point does not touch the outline, paint it red.
   DIM points(0 TO 2) AS GpPoint = {GDIP_POINT(230, 138), GDIP_POINT(100, 120), GDIP_POINT(150, 170)}
   FOR j AS LONG = 0 TO 2
      IF path.IsOutlineVisible(points(j).x, points(j).y, @yellowPen, NULL) THEN
         brush.SetColor(GDIP_ARGB(255, 0, 255,  0))
      ELSE
         brush.SetColor(GDIP_ARGB(255, 255, 0,  0))
      END IF
      graphics.FillEllipse(@brush, points(j).x - 3, points(j).y - 3, 6, 6)
   NEXT

END SUB
' ========================================================================================
```


# <a name="IsVisible"></a>IsVisible (CGpGraphicsPath)

Determines whether a specified point lies in the area that is filled when this path is filled by a specified **Graphics** object.

```
FUNCTION IsVisible (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL pGraphics AS CGpGraphics PTR = NULL) AS BOOLEAN
FUNCTION IsVisible (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL pGraphics AS CGpGraphics PTR = NULL) AS BOOLEAN
FUNCTION IsVisible (BYVAL rc AS GpRectF PTR, BYVAL pGraphics AS CGpGraphics PTR = NULL) AS BOOLEAN
FUNCTION IsVisible (BYVAL rc AS GpRect PTR, BYVAL pGraphics AS CGpGraphics PTR = NULL) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | The x-coordinate of the point to be tested. |
| *y* | The y-coordinate of the point to be tested. |
| *pGraphics* | Optional. Pointer to a **Graphics** object that specifies a world-to-device transformation. If the value of this parameter is NULL, the test is done in world coordinates; otherwise, the test is done in device coordinates. The default value is NULL.  |

#### Return value

If the test point touches the outline of this path, this method returns TRUE; otherwise, it returns FALSE.

#### Example

```
' ========================================================================================
' The following example creates an elliptical path and draws that path with a narrow black
' pen. Then the code tests each point in an array to see whether the point lies in the
' interior of the path. Points that lie in the interior are painted green, and points that
' do not lie in the interior are painted red.
' ========================================================================================
SUB Example_IsVisible (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   DIM blackPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 0), 1)
   DIM brush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0,  0)

   ' // Create and draw a path
   DIM path AS CGpGraphicsPath
   path.AddEllipse(50, 50, 200, 100)
   graphics.DrawPath(@blackPen, @path)

   ' // Create an array of four points, and determine whether each
   ' // point in the array touches the outline of the path.
   ' // If a point touches the outline, paint it green.
   ' // If a point does not touch the outline, paint it red.
   DIM points(0 TO 3) AS GpPoint = {GDIP_POINT(50, 100), GDIP_POINT(250, 100), _
       GDIP_POINT(150, 170), GDIP_POINT(180, 60)}
   FOR j AS LONG = 0 TO 3
      IF path.IsVisible(points(j).x, points(j).y, @graphics) THEN
         brush.SetColor(GDIP_ARGB(255, 0, 255,  0))
      ELSE
         brush.SetColor(GDIP_ARGB(255, 255, 0,  0))
      END IF
      graphics.FillEllipse(@brush, points(j).x - 3, points(j).y - 3, 6, 6)
   NEXT

END SUB
' ========================================================================================
```

# <a name="Outline"></a>Outline (CGpGraphicsPath)

Transforms and flattens this path, and then converts this path's data points so that they represent only the outline of the path.

```
FUNCTION Outline (pMatrix AS CGpMatrix PTR = NULL, BYVAL flatness AS SINGLE = FlatnessDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMatrix* | Optional. Pointer to a **Matrix** object that specifies the transformation. If this parameter is NULL, no transformation is applied. The default value is NULL. |
| *flatness* | Optional. Single precision number that specifies the maximum error between the path and its flattened approximation. Reducing the flatness increases the number of line segments in the approximation. The default value is **FlatnessDefault**, which is a constant defined in Gdiplusenums.bi. |

#### Return value

If the test point touches the outline of this path, this method returns TRUE; otherwise, it returns FALSE.

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object and calls the GraphicsPath.AddClosedCurve
' method to add a closed cardinal spline to the path. The code calls the GraphicsPath.Widen
' method to widen the path and then draws the path. Next, the code calls the path's Outline
' method. The code calls the TranslateTransform method of a Graphics object so that the
' outlined path drawn by the subsequent call to DrawPath sits to the right of the first path.
' ========================================================================================
SUB Example_Outline (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   DIM bluePen AS CGpPen = GDIP_ARGB(255, 0, 0, 255)
   DIM greenPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 255, 0), 10)

   DIM points(0 TO 3) AS GpPoint = {GDIP_POINT(20, 20), GDIP_POINT(160, 100), _
       GDIP_POINT(140, 60), GDIP_POINT(60, 100)}
  
   DIM path AS CGpGraphicsPath
   path.AddClosedCurve(@points(0), 4)

   path.Widen(@greenPen)
   graphics.DrawPath(@bluePen, @path)

   path.Outline

   graphics.TranslateTransform(180, 0)
   graphics.DrawPath(@bluePen, @path)

END SUB
' ========================================================================================
```

# <a name="Reset"></a>Reset (CGpGraphicsPath)

Empties the path and sets the fill mode to **FillModeAlternate**.

```
FUNCTION Reset () AS GpStatus
```

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="Reverse"></a>Reverse (CGpGraphicsPath)

Reverses the order of the points that define this path's lines and curves.

```
FUNCTION Reverse () AS GpStatus
```

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="SetFillMode"></a>SetFillMode (CGpGraphicsPath)

Sets the fill mode of this path.

```
FUNCTION SetFillMode (BYVAL nFillmode AS FillMode) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nFillmode* | Element of the **FillMode** enumeration that specifies how to fill areas when the path intersects itself. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="SetMarker"></a>SetMarker (CGpGraphicsPath)

Designates the last point in this path as a marker point.

```
FUNCTION SetMarker () AS GpStatus
```

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A **GraphicsPath** object has an array of points and an array of types. Each element in the array of types is a byte that specifies the point type and a set of flags for the corresponding element in the array of points. Possible point types and flags are listed in the **PathPointType** enumeration.

Each time you add a line, curve, or shape to a path, the point array and the type array are expanded. When you call **SetMarker**, a marker flag is placed in the last byte of the type array. That flag designates the last point of the point array as a marker point.

Markers divide a path into sections. You can use a **GraphicsPathIterator** object to draw selected sections of a path.


# <a name="StartFigure"></a>StartFigure (CGpGraphicsPath)

Starts a new figure without closing the current figure. Subsequent points added to this path are added to the new figure.

```
FUNCTION StartFigure () AS GpStatus
```

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a path and adds two figures to that path. The first figure
' has three arcs, and the second figure has two arcs. The arcs within a figure are connected
' by straight lines, but there is no connecting line between the last arc in the first
' figure and the first arc in the second figure.
' ========================================================================================
SUB Example_StartFigure (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   DIM pen AS CGpPen = GDIP_ARGB(255, 0, 0, 255)
   DIM rc AS GpRect = GDIP_RECT(0, 0, 100, 50)
   DIM path AS CGpGraphicsPath

   path.AddArc(0, 0, 100, 50, 0.0, 180.0)
   path.AddArc(0, 60, 100, 50, 0.0, 180.0)
   path.AddArc(0, 120, 100, 50, 0.0, 180.0)

   ' // Start a new figure (subpath).
   ' // Do not close the current figure.
   path.StartFigure
   path.AddArc(0, 180, 100, 50, 0.0, 180.0)
   path.AddArc(0, 240, 100, 50, 0.0, 180.0)

   graphics.DrawPath(@pen, @path)

END SUB
' ========================================================================================
```


# <a name="Transform"></a>Transform (CGpGraphicsPath)

Multiplies each of this path's data points by a specified matrix.

```
FUNCTION Transform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMatrix* | Pointer to a **Matrix** object that specifies the transformation. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a path and adds two figures to that path. The first figure
' has three arcs, and the second figure has two arcs. The arcs within a figure are connected
' by straight lines, but there is no connecting line between the last arc in the first
' figure and the first arc in the second figure.
' ========================================================================================
SUB Example_Transform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   DIM path AS CGpGraphicsPath
   path.AddRectangle(40, 10, 200, 50)

   ' // Draw the path in blue before applying a transformation
   graphics.DrawPath(@CGpPen(GDIP_ARGB(255, 0, 0, 255)), @path)

   ' // Transform the path
   DIM matrix AS CGpMatrix
   matrix.Rotate(30.0)
   path.Transform(@matrix)

   ' // Draw the transformed path in red.
   graphics.DrawPath(@CGpPen(GDIP_ARGB(255, 255, 0,  0)), @path)

END SUB
' ========================================================================================
```


# <a name="Warp"></a>Warp (CGpGraphicsPath)

Applies a warp transformation to this path. The **Warp** method also flattens (converts to a sequence of straight lines) the path.

```
FUNCTION Warp (BYVAL destPoints AS GpPointF PTR, BYVAL count AS INT_, BYVAL srcRect AS GpRectF PTR, _
   BYVAL pMatrix AS CGpMatrix PTR = NULL, BYVAL nWarpMode AS WarpMode = WarpModePerspective, _
   BYVAL flatness AS SINGLE = FlatnessDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *destPoints* | Pointer to an array of points that, along with the *srcRect* parameter, defines the warp transformation. |
| *count* | Integer that specifies the number of points in the *destPoints* array. The value of this parameter must be 3 or 4. |
| *srcRect* | Reference to a rectangle that, along with the *destPoints* parameter, defines the warp transformation. |
| *pMatrix* | Optional. Pointer to a Matrix object that represents a transformation to be applied along with the warp. If this parameter is NULL, no transformation is applied. The default value is NULL. |
| *nWarpMode* | Optional. Element of the **WarpMode** enumeration that specifies the kind of warp to be applied. The default value is **WarpModePerspective**. |
| *flatness* | Optional. Real number that influences the number of line segments that are used to approximate the original path. Small values specify that many line segments are used, and large values specify that few line segments are used. The default value is **FlatnessDefault**, which is a constant defined in Gdiplusenums.bi. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object and adds a closed figure to the path.
' The code defines a warp transformation by specifying a source rectangle and an array of
' four destination points. The source rectangle and destination points are passed to the
' Warp method. The code draws the path twice: once before it has been warped and once after
' it has been warped.
' ========================================================================================
SUB Example_Warp (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   ' // Create a path.
   DIM points(0 TO 7) AS GpPointF = {GDIP_POINTF(20.0, 60.0), GDIP_POINTF(30.0, 90.0), _
      GDIP_POINTF(15.0, 110.0), GDIP_POINTF(15.0, 145.0), GDIP_POINTF(55.0, 145.0), _
      GDIP_POINTF(55.0, 110.0), GDIP_POINTF(40.0, 90.0), GDIP_POINTF(50.0, 60.0)}
   DIM path AS CGpGraphicsPath
   path.AddLines(@points(0), 8)
   path.CloseFigure

   ' // Draw the path before applying a warp transformation.
   DIM bluePen AS CGpPen = GDIP_ARGB(255, 0, 0, 255)
   graphics.DrawPath(@bluePen, @path)

   ' // Define a warp transformation, and warp the path.
   DIM srcRect AS GpRectF = GDIP_RECTF(10.0, 50.0, 50.0, 100.0)
   DIM destPts(0 TO 3) AS GpPointF = {GDIP_POINTF(220.0, 10.0), GDIP_POINTF(280.0, 10.0), _
       GDIP_POINTF(100.0, 150.0), GDIP_POINTF(400.0, 150.0)}
   path.Warp(@destPts(0), 4, @srcRect)

   ' // Draw the warped path.
   graphics.DrawPath(@bluePen, @path)

   ' // Draw the source rectangle and the destination polygon.
   DIM blackPen AS CGpPen = GDIP_ARGB(255, 0, 0, 0)
   graphics.DrawRectangle(@blackPen, @srcRect)
   graphics.DrawLine(@blackPen, @destPts(0), @destPts(1))
   graphics.DrawLine(@blackPen, @destPts(0), @destPts(2))
   graphics.DrawLine(@blackPen, @destPts(1), @destPts(3))
   graphics.DrawLine(@blackPen, @destPts(2), @destPts(3))

END SUB
' ========================================================================================
```

# <a name="Widen"></a>Widen (CGpGraphicsPath)

Replaces this path with curves that enclose the area that is filled when this path is drawn by a specified pen. The **Widen** method also flattens the path.

```
FUNCTION Widen (BYVAL pPen AS CGpPen PTR, BYVAL pMatrix AS CGpMatrix PTR = NULL, _
   BYVAL flatness AS SINGLE = FlatnessDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pPen* | Pointer to a **Pen** object. The path is made as wide as it would be when drawn by this pen. |
| *pMatrix* | Optional. Pointer to a **Matrix** object that specifies a transformation to be applied along with the widening. If this parameter is NULL, no transformation is applied. The default value is NULL. |
| *flatness* | Optional. Real number that influences the number of line segments that are used to approximate the original path. Small values specify that many line segments are used, and large values specify that few line segments are used. The default value is **FlatnessDefault**, which is a constant defined in Gdiplusenums.bi. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object and adds a closed curve to the path.
' The code creates a green pen that has a width of 10 and passes the address of that pen
' to the GraphicsPath.Widen method. Then the code draws the path with a blue pen of width 1.
' ========================================================================================
SUB Example_Widen (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   DIM bluePen AS CGpPen = GDIP_ARGB(255, 0, 0, 255)
   DIM greenPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 255,  0), 10)
   DIM points(0 TO 3) AS GpPoint = {GDIP_POINT(20, 20), GDIP_POINT(160, 100), _
       GDIP_POINT(140, 60), GDIP_POINT(60, 100)}

   DIM path AS CGpGraphicsPath
   path.AddClosedCurve(@points(0), 4)
   path.Widen(@greenPen)
   graphics.DrawPath(@bluePen, @path)

END SUB
' ========================================================================================
```

# <a name="ConstructorGraphicsPathIterator"></a>Constructor (CGpGraphicsPathIterator)

Creates a new **GraphicsPathIterator** object and associates it with a **GraphicsPath** object.

```
CONSTRUCTOR CGpGraphicsPathIterator (BYVAL pPath AS CGpGraphicsPath PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pPath* | Pointer to a **GraphicsPath** object that will be associated with this **GraphicsPathIterator** object. |


# <a name="CopyData"></a>CopyData (CGpGraphicsPathIterator)

Copies a subset of the path's data points to a **GpPointF** array and copies a subset of the path's point types to a byte array.

```
FUNCTION CopyData (BYVAL pts AS GpPointF PTR, BYVAL types AS BYTE PTR, _
   BYVAL startIndex AS INT_, BYVAL endIndex AS INT_) AS INT_
```

| Parameter  | Description |
| ---------- | ----------- |
| *pts* | Pointer to an array that receives a subset of the path's data points. |
| *types* | Pointer to an array that receives a subset of the path's point types. |
| *startIndex* | Integer that specifies the starting index of the points and types to be copied. |
| *endIndex* | Integer that specifies the ending index of the points and types to be copied. |

#### Return value

This method returns the number of points copied. This is the same as the number of types copied.

# <a name="Enumerate"></a>Enumerate (CGpGraphicsPathIterator)

Copies the path's data points to a **GpPointF** array and copies the path's point types to a byte array.

```
FUNCTION Enumerate (BYVAL pts AS GpPointF PTR, BYVAL types AS BYTE PTR, BYVAL count AS INT_) AS INT_
```

| Parameter  | Description |
| ---------- | ----------- |
| *pts* | Pointer to an array that receives the path's data points. |
| *types* | Pointer to an array that receives the path's point types. |
| *count* | Integer that specifies the number of elements in the points array. This is the same as the number of elements in the types array. |

#### Return value

This method returns the number of points retrieved.

#### Remarks

This **GraphicsPathIterator** object is associated with a **GraphicsPath** object. That **GraphicsPath** object has an array of points and an array of types. Each element in the array of types is a byte that specifies the point type and a set of flags for the corresponding element in the array of points. Possible point types and flags are listed in the **PathPointType** enumeration.

You can call the **GetCount** method to determine the number of data points in the path. The points parameter points to a buffer that receives the data points, and the types parameter points to a buffer that receives the types. Before you call the **Enumerate** method, you must allocate memory for those buffers. The size of the points buffer should be the return value of **GetCount** multiplied by **SIZEOF(GpPointF)**. The size of the types buffer should be the  return value of **GetCount**.

# <a name="GetCount"></a>GetCount (CGpGraphicsPathIterator)

Returns the number of data points in the path.

```
FUNCTION GetCount () AS INT_
```

#### Remarks

This **GraphicsPathIterator** object is associated with a **GraphicsPath** object. That **GraphicsPath** object has an array of points and an array of types. Each element in the array of types is a byte that specifies the point type and a set of flags for the corresponding element in the array of points. Possible point types and flags are listed in the **PathPointType** enumeration.

# <a name="GetSubpathCount"></a>GetSubpathCount (CGpGraphicsPathIterator)

Returns the number of subpaths (also called figures) in the path.

```
FUNCTION GetSubpathCount () AS INT_
```

# <a name="HasCurve"></a>HasCurve (CGpGraphicsPathIterator)

Determines whether the path has any curves.

```
FUNCTION HasCurve () AS BOOLEAN
```

#### Return value

If the path has at least one curve, this method returns TRUE; otherwise, it returns FALSE.

#### Remarks

All curves in a path are stored as sequences of Bézier splines. For example, when you add an ellipse to a path, you specify the upper-left corner, the width, and the height of the ellipse's bounding rectangle. Those numbers (upper-left corner, width, and height) are not stored in the path; instead; the ellipse is converted to a sequence of four Bézier splines. The path stores the endpoints and control points of those Bézier splines.

A path stores an array of data points, each of which belongs to a line or a Bézier spline. If some of the points in the array belong to Bézier splines, then **HasCurve** returns TRUE. If all points in the array belong to lines, then **HasCurve** returns FALSE.

Certain methods flatten a path, which means that all the curves in the path are converted to sequences of lines. After a path has been flattened, **HasCurve** will always return FALSE. Flattening happens when you call the **Flatten**, **Widen**, or **Warp** method of the **GraphicsPath** class.


# <a name="NextMarker"></a>NextMarker (CGpGraphicsPathIterator)

Gets the starting index and the ending index of the next marker-delimited section in this iterator's associated path.

```
FUNCTION NextMarker (BYVAL startIndex AS INT_ PTR, BYVAL endIndex AS INT_ PTR) AS INT_
```

Gets the next marker-delimited section of this iterator's associated path.

```
FUNCTION NextMarker (BYVAL pPath AS CGpGraphicsPath PTR) AS INT_
```

| Parameter  | Description |
| ---------- | ----------- |
| *startIndex* | Pointer to an LONG that receives the starting index. |
| *endIndex* | Pointer to an LONG that receives the ending index. |
| *pPath* | Pointer to a **GraphicsPath** object. This method sets the data points of this **GraphicsPath** object to match the data points of the retrieved section. |

#### Return value

This method returns the number of data points in the retrieved section. If there are no more marker-delimited sections to retrieve, this method returns 0.

#### Remarks

A path has an array of data points that define its lines and curves. You can call a path's **SetMarker** method to designate certain points in the array as markers. Those marker points divide the path into sections.

The first time you call the **NextMarker** method of an iterator, it gets the first marker-delimited section of that iterator's associated path. The second time, it gets the second section, and so on. Each time you call **NextSubpath**, it returns the number of data points in the retrieved section. When there are no sections remaining, it returns 0.

# <a name="NextPathType"></a>NextPathType (CGpGraphicsPathIterator)

Gets the starting index and the ending index of the next group of data points that all have the same type.

```
FUNCTION NextPathType (BYVAL pathType AS BYTE PTR, BYVAL startIndex AS INT_ PTR, _
   BYVAL endIndex AS INT_ PTR) AS INT_
```

| Parameter  | Description |
| ---------- | ----------- |
| *pathType* | Pointer to a BYTE that receives the point type shared by all points in the group. Possible values are **PathPointTypeLine** and **PathPointTypeBezier**, which are elements of the **PathPointType** enumeration. |
| *startIndex* | Pointer to an LONG that receives the starting index of the group of points. |
| *endIndex* | Pointer to an LONG that receives the ending index of the group of points. |

#### Return value

This method returns the number of data points in the group. If there are no more groups in the path, this method returns 0.

#### Remarks

A path has an array of data points that define its lines and curves. All curves in the path are represented as Bézier splines, so a given point in the array has one of two types: **PathPointTypeLine** or **PathPointTypeBezier**.

The first time you call the **NextSubpath** method of an iterator, it gets the starting and ending indices of the first group of points that all have the same type. The second time, it gets the second group, and so on. Each time you call **NextSubpath**, it returns the number of data points in the obtained group. When there are no groups remaining, it returns 0.

# <a name="NextSubpath"></a>NextSubpath (CGpGraphicsPathIterator)

Gets the starting index and the ending index of the next subpath (figure) in this iterator's associated path.

```
FUNCTION NextSubpath (BYVAL startIndex AS INT_ PTR, BYVAL endIndex AS INT_ PTR, _
   BYVAL isClosed AS BOOL PTR) AS INT_
FUNCTION NextSubpath (BYVAL startIndex AS INT_ PTR, BYVAL endIndex AS INT_ PTR, _
   BYVAL isClosed AS BOOL PTR) AS INT_
```

| Parameter  | Description |
| ---------- | ----------- |
| *startIndex* | Pointer to an LONG that receives the starting index of the group of points. |
| *endIndex* | Pointer to an LONG that receives the ending index of the group of points. |
| *isClosed* | Pointer to a BOOL that receives a value that indicates whether the obtained figure is closed. If the figure is closed, the received value is TRUE; otherwise, the received value is FALSE. |

#### Return value

This method returns the number of data points in the next figure. If there are no more figures in the path, this method returns 0.

#### Remarks

The first time you call the **NextSubpath** method of an iterator, it gets the starting and ending indices of the first group of points that all have the same type. The second time, it gets the second group, and so on. Each time you call **NextSubpath**, it returns the number of data points in the obtained group. When there are no groups remaining, it returns 0.

# <a name="Rewind"></a>Rewind (CGpGraphicsPathIterator)

Rewinds this iterator to the beginning of its associated path.

```
SUB Rewind
```

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

The first time you call the **NextSubpath** method of an iterator, it gets the first figure (subpath) of that iterator's associated path. The second time, it gets the second figure, and so on. When you call **Rewind**, the sequence starts over; that is, after you call **Rewind**, the next call to **NextSubpath** gets the first figure in the path. The **NextMarker** and **NextPathType** methods behave similarly.
