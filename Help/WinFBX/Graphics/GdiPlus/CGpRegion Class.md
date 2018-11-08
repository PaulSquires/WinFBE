# CGpRegion Class

Provides methods to manage **Regions**. **Regions** are used in clipping and hit-testing operations.

The **Region** object describes an area of the display surface. The area can be any shape. In other words, the boundary of the area can be a combination of curved and straight lines. Regions can also be created from the interiors of rectangles, paths, or a combination of these. Regions are used in clipping and hit-testing operations.

**Inherits from**: CGpBase.<br>
**Imclude file**: CGpRegion.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#Constructors) | Create a **Region** object. |
| [Clone](#Clone) | Copies the contents of the existing **Region** object into a new **Region** object. |
| [Complement](#Complement) | Updates this region to the portion of the specified rectangle's interior that does not intersect this region. |
| [Equals](#Equals) | Determines whether this region is equal to a specified region. |
| [Exclude](#Exclude) | Updates this region to the portion of itself that does not intersect the specified rectangle's interior. |
| [GetBounds](#GetBounds) | Gets a rectangle that encloses this region. |
| [GetData](#GetData) | Gets data that describes this region. |
| [GetDataSize](#GetDataSize) | Gets the number of bytes of data that describes this region. |
| [GetHRGN](#GetHRGN) | Creates a Windows Graphics Device Interface (GDI) region from this region. |
| [GetRegionScans](#GetRegionScans) | Gets an array of rectangles that approximate this region. |
| [GetRegionScansCount](#GetRegionScansCount) | Gets the number of rectangles that approximate this region. |
| [Intersect](#Intersect) | Updates this region to the portion of itself that intersects the specified rectangle's interior. |
| [IsEmpty](#IsEmpty) | Determines whether this region is empty. |
| [IsInfinite](#IsInfinite) | Determines whether this region is infinite. |
| [IsVisible](#IsVisible) | Determines whether a point is inside this region. |
| [MakeEmpty](#MakeEmpty) | Updates this region to an empty region. In other words, the region occupies no space on the display device. |
| [MakeInfinite](#MakeInfinite) | Updates this region to an infinite region. |
| [Transform](#Transform) | Transforms this region by multiplying each of its data points by a specified matrix. |
| [Translate](#Translate) | Offsets this region by specified amounts in the horizontal and vertical directions. |
| [Union_](#Union_) | Updates this region to all portions (intersecting and nonintersecting) of itself and all portions of the specified rectangle's interior. |
| [Xor_](#Xor_) | Updates this region to the nonintersecting portions of itself and the specified rectangle's interior. |

# <a name="Constructors"></a>Constructors

Creates a region that is defined by a rectangle.

```
CONSTRUCTOR CGpRegion
CONSTRUCTOR CGpRegion (BYVAL pRegion AS CGpRegion PTR)
CONSTRUCTOR CGpRegion (BYVAL rc AS GpRectF PTR)
CONSTRUCTOR CGpRegion (BYVAL rc AS GpRect PTR)
CONSTRUCTOR CGpRegion (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE)
CONSTRUCTOR CGpRegion (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_)
```

Creates a region that is defined by a path (a **GraphicsPath** object) and has a fill mode that is contained in the **GraphicsPath** object.

```
CONSTRUCTOR CGpRegion (BYVAL pPath AS CGpGraphicsPath PTR)
```

Creates a region that is defined by data obtained from another region.

```
CONSTRUCTOR CGpRegion (BYVAL regionData AS BYTE PTR, BYVAL nSize AS LONG)
```

Creates a region that is identical to the region that is specified by a handle to a Windows Graphics Device Interface (GDI) region.

```
CONSTRUCTOR CGpRegion (BYVAL hRgn AS HRGN)
```

| Parameter  | Description |
| ---------- | ----------- |
| *rc* | Reference to a rectangle. |
| *x, y, nWidth, nHeight* | left coordinate, top coordinate, width and height of a rectangle. |
| *pPath* | Pointer to a **GraphicsPath** object that specifies the path |
| *regionData* | Pointer to an array of bytes that specifies a region. The data contained in the bytes is obtained from another region by using the **GetData** method. |
| *nSize* | Integer that specifies the number of bytes in the *regionData* array. |
| *hRgn* | Handle to an existing GDI region.  |


# <a name="Clone"></a>Clone

Copies the contents of the existing **Region** object into a new **Region** object.

```
FUNCTION Clone (BYVAL pCloneRegion AS CGpRegion PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pCloneRegion* | Pointer to the **Region** object where to copy the contents of the existing object. |

#### Example

```
' ========================================================================================
' The following example creates two regions, one from a rectangle and the other from a path.
' Next, the code clones the region that was created from a path and uses a solid brush to
' fill the cloned region. Then, it forms the union of the two regions and fills it.
' ========================================================================================
SUB Example_CloneRegion (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   DIM alphaBrush AS CGpSolidBrush = GDIP_ARGB(128, 0, 0, 255)

   DIM pts(0 TO 5) AS GpPoint = {GDIP_POINT(110, 20), GDIP_POINT(120, 30), GDIP_POINT(100, 60), GDIP_POINT(120, 70), GDIP_POINT(150, 60), GDIP_POINT(140, 10)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(110, 20), (120, 30), (100, 60), (120, 70), (150, 60), (140, 10)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 5) AS GpPoint
'   pts(0).x = 110 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 30 : pts(2).x = 100 : pts(2).y = 60
'   pts(3).x = 120 : pts(3).y = 70 : pts(4).x = 150 : pts(4).y = 60 : pts(5).x = 140 : pts(5).y = 10
'#endif
   DIM pPath AS CGpGraphicsPath
   pPath.AddClosedCurve(@pts(0), 6)

   ' // Create a region from a rectangle.
   DIM rc AS GpRect = GDIP_RECT(65, 15, 70, 45)
'#ifdef __FB_64BIT__
'   DIM rc AS GpRect = (65, 15, 70, 45)
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpRect is defined as Rect in 64 bit and as Rect_ in 32 bit.
'   DIM rc AS GpRect : rc.x = 65 : rc.y = 15 : rc.Width = 70 : rc.Height = 45
'#endif
   DIM rectRegion AS CGpRegion = @rc

   ' // Create a region from a curved path.
   DIM pathRegion AS CGpRegion = @pPath

   ' // Make a copy (clone) of the curved region.
   DIM pClonedRegion AS CGpRegion
   pathRegion.Clone(@pClonedRegion)
   ' // You can also use:
   ' IM pClonedRegion AS CGpRegion = @pathRegion

   ' // Fill the cloned region with a red brush.
   graphics.FillRegion(@solidBrush, @pClonedRegion)

   ' // Form the union of the cloned region and the rectangular region.
   pClonedRegion.Union_(@rectRegion)

   ' // Fill the union of the two regions with a semitransparent blue brush.
   graphics.FillRegion(@alphaBrush, @pClonedRegion)

END SUB
' ========================================================================================
```

# <a name="Complement"></a>Complement

Updates this region to the portion of the specified rectangle's interior that does not intersect this region.

```
FUNCTION Complement (BYVAL rc AS GpRectF PTR) AS GpStatus
FUNCTION Complement (BYVAL rc AS GpRect PTR) AS GpStatus
```

Updates this region to the portion of the specified path's interior that does not intersect this region.

```
FUNCTION Complement (BYVAL pPath AS CGpGraphicsPath PTR) AS GpStatus
```

Updates this region to the portion of another region that does not intersect this region.

```
FUNCTION Complement (BYVAL pRegion AS CGpRegion PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *rc* | Reference to a rectangle to use to update this **Region** object. |
| *pPath* | Pointer to a **GraphicsPath** object that specifies the path to use to update this **Region** object. |
| *pRegion* | Pointer to a **Region** object to use to update this **Region** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a region from a path and then uses a rectangle to update the region.
' ========================================================================================
SUB Example_ComplementRegion (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)

   DIM pts(0 TO 5) AS GpPoint = {GDIP_POINT(110, 20), GDIP_POINT(120, 30), GDIP_POINT(100, 60), GDIP_POINT(120, 70), GDIP_POINT(150, 60), GDIP_POINT(140, 10)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(110, 20), (120, 30), (100, 60), (120, 70), (150, 60), (140, 10)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 5) AS GpPoint
'   pts(0).x = 110 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 30 : pts(2).x = 100 : pts(2).y = 60
'   pts(3).x = 120 : pts(3).y = 70 : pts(4).x = 150 : pts(4).y = 60 : pts(5).x = 140 : pts(5).y = 10
'#endif
   DIM pPath AS CGpGraphicsPath
   pPath.AddClosedCurve(@pts(0), 6)

   DIM rcf AS GpRectF = GDIP_RECTF(65.3, 15.1, 70.0, 45.8)
'#ifdef __FB_64BIT__
'   DIM rcf AS GpRectF = (65.3, 15.1, 70.0, 45.8)
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpRect is defined as Rect in 64 bit and as Rect_ in 32 bit.
'   DIM rcf AS GpRectF : rcf.x = 65.3 : rcf.y = 15.1 : rcf.Width = 70.0 : rcf.Height = 45.8
'#endif

   ' // Create a region from a path.
   DIM pRegion AS CGpRegion = @pPath

   ' // Update the region so that it consists of all points inside the
   ' // rectangle but outside the path region.
   pRegion.Complement(@rcf)
   graphics.FillRegion(@solidBrush, @pRegion)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' The following example creates a region from a path and then uses a rectangle to update the region.
' ========================================================================================
SUB Example_ComplementRegion (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)

   DIM pts(0 TO 5) AS GpPoint = {GDIP_POINT(110, 20), GDIP_POINT(120, 30), GDIP_POINT(100, 60), GDIP_POINT(120, 70), GDIP_POINT(150, 60), GDIP_POINT(140, 10)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(110, 20), (120, 30), (100, 60), (120, 70), (150, 60), (140, 10)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 5) AS GpPoint
'   pts(0).x = 110 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 30 : pts(2).x = 100 : pts(2).y = 60
'   pts(3).x = 120 : pts(3).y = 70 : pts(4).x = 150 : pts(4).y = 60 : pts(5).x = 140 : pts(5).y = 10
'#endif
   DIM pPath AS CGpGraphicsPath
   pPath.AddClosedCurve(@pts(0), 6)

   DIM rc AS GpRect = GDIP_RECT(65, 15, 70, 45)
'#ifdef __FB_64BIT__
'   DIM rc AS GpRect = (65, 15, 70, 45)
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpRect is defined as Rect in 64 bit and as Rect_ in 32 bit.
'   DIM rc AS GpRect : rc.x = 65 : rc.y = 15 : rc.Width = 70 : rc.Height = 45
'#endif

   ' // Create a region from a rectangle
   DIM pRegion AS CGpRegion = @rc

   ' // Update the region so that it consists of all points inside the
   ' // rectangle but outside the path region
   pRegion.Complement(@pPath)
   graphics.FillRegion(@solidBrush, @pRegion)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' The following example creates a region from a path and then uses a rectangle to update the region.
' ========================================================================================
SUB Example_ComplementRegion (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)

   DIM pts(0 TO 5) AS GpPoint = {GDIP_POINT(110, 20), GDIP_POINT(120, 30), GDIP_POINT(100, 60), GDIP_POINT(120, 70), GDIP_POINT(150, 60), GDIP_POINT(140, 10)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(110, 20), (120, 30), (100, 60), (120, 70), (150, 60), (140, 10)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 5) AS GpPoint
'   pts(0).x = 110 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 30 : pts(2).x = 100 : pts(2).y = 60
'   pts(3).x = 120 : pts(3).y = 70 : pts(4).x = 150 : pts(4).y = 60 : pts(5).x = 140 : pts(5).y = 10
'#endif
   DIM pPath AS CGpGraphicsPath
   pPath.AddClosedCurve(@pts(0), 6)

   DIM rc AS GpRect = GDIP_RECT(65, 15, 70, 45)
'#ifdef __FB_64BIT__
'   DIM rc AS GpRect = (65, 15, 70, 45)
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpRect is defined as Rect in 64 bit and as Rect_ in 32 bit.
'   DIM rc AS GpRect : rc.x = 65 : rc.y = 15 : rc.Width = 70 : rc.Height = 45
'#endif

   ' // Create a region from a rectangle
   DIM rectRegion AS CGpRegion = @rc
   ' // Create a region from a path
   DIM pathRegion AS CGpRegion = @pPath

   ' // Update the region so that it consists of all points inside the
   ' // rectangle but outside the path region
   rectRegion.Complement(@pathRegion)
   graphics.FillRegion(@solidBrush, @rectRegion)

END SUB
' ========================================================================================
```

# <a name="Equals"></a>Equals

Determines whether this region is equal to a specified region.

```
FUNCTION Equals (BYVAL pRegion AS CGpRegion PTR, BYVAL pGraphics AS CGpGraphics PTR) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *pRegion* | Pointer to a **Region** object to test against this Region object. |
| *pGraphics* | Pointer to a **Graphics** object that contains the world and page transformations required to calculate the device coordinates of this region and the region that is specified by the region parameter. |

#### Return value

If this region is identical to the specified region, this method returns TRUE; otherwise, it returns FALSE.

#### Example

```
' ========================================================================================
' The following example creates two regions, one from a rectangle and the other from a path.
' The code then tests to determine whether the two regions are the same and performs a task
' based on the result of the test.
' ========================================================================================
SUB Example_Equals (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create solid brushes
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   DIM alphaBrush AS CGpSolidBrush = GDIP_ARGB(128, 0, 0, 255)

   DIM pts(0 TO 3) AS GpPoint = {GDIP_POINT(20, 20), GDIP_POINT(120, 20), GDIP_POINT(120, 70), GDIP_POINT(20, 70)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(20, 20), (120, 20), (120, 70), (20, 70)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 3) AS GpPoint
'   pts(0).x = 20 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 20 : pts(2).x = 120 : pts(2).y = 70 : pts(3).x = 20 : pts(3).y = 70
'#endif

   DIM rc AS GpRect = GDIP_RECT(20, 20, 100, 50)
'#ifdef __FB_64BIT__
'   DIM rc AS GpRect = (20, 20, 100, 50)
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpRect is defined as Rect in 64 bit and as Rect_ in 32 bit.
'   DIM rc AS GpRect : rc.x = 20 : rc.y = 20 : rc.Width = 100 : rc.Height = 50
'#endif

   DIM path AS CGpGraphicsPath
   path.AddPolygon(@pts(0), 4)

   ' // Create a region from a rectangle.
   DIM rectRegion AS CGpRegion = @rc
   graphics.FillRegion(@solidBrush, @rectRegion)

   ' // Create a region from a path.
   DIM pathRegion AS CGpRegion = @path
   graphics.FillRegion(@alphaBrush, @pathRegion)

   IF pathRegion.Equals(@rectRegion, @graphics) THEN
      ' // The two regions are the same.
      ' // Perform some task.
      PRINT "Equal"
   ELSE
      ' // The two regions are not the same.
      ' // Perform a different task.
      PRINT "Not equal"
   END IF

END SUB
' ========================================================================================
```

# <a name="Exclude"></a>Exclude

Updates this region to the portion of itself that does not intersect the specified rectangle's interior.

```
FUNCTION Exclude (BYVAL rc AS GpRectF PTR) AS GpStatus
FUNCTION Exclude (BYVAL rc AS GpRect PTR) AS GpStatus
```

Updates this region to the portion of itself that does not intersect the specified path's interior.

```
FUNCTION Exclude (BYVAL pPath AS CGpGraphicsPath PTR) AS GpStatus
```

Updates this region to the portion of itself that does not intersect another region.

```
FUNCTION Exclude (BYVAL pRegion AS CGpRegion PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *rc* | Reference to a rectangle to use to update this **Region** object. |
| *pPath* | Pointer to a **GraphicsPath** object that specifies the path to use to update this **Region** object. |
| *pRegion* |  Pointer to a Region object to use to update this Region object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a region from a path and then uses a rectangle to update the region.
' ========================================================================================
SUB Example_ExcludeRegion (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)

   DIM pts(0 TO 5) AS GpPoint = {GDIP_POINT(110, 20), GDIP_POINT(120, 30), GDIP_POINT(100, 60), GDIP_POINT(120, 70), GDIP_POINT(150, 60), GDIP_POINT(140, 10)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(110, 20), (120, 30), (100, 60), (120, 70), (150, 60), (140, 10)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 5) AS GpPoint
'   pts(0).x = 110 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 30 : pts(2).x = 100 : pts(2).y = 60
'   pts(3).x = 120 : pts(3).y = 70 : pts(4).x = 150 : pts(4).y = 60 : pts(5).x = 140 : pts(5).y = 10
'#endif
   DIM pPath AS CGpGraphicsPath
   pPath.AddClosedCurve(@pts(0), 6)

   DIM rcf AS GpRectF = GDIP_RECTF(65.3, 15.1, 70.0, 45.8)
'#ifdef __FB_64BIT__
'   DIM rcf AS GpRectF = (65.3, 15.1, 70.0, 45.8)
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpRect is defined as Rect in 64 bit and as Rect_ in 32 bit.
'   DIM rcf AS GpRectF : rcf.x = 65.3 : rcf.y = 15.1 : rcf.Width = 70.0 : rcf.Height = 45.8
'#endif

   ' // Create a region from a path
   DIM pRegion AS CGpRegion = @pPath

   ' // Exclude a rectangle from the path region.
   pRegion.Exclude(@rcf)
   graphics.FillRegion(@solidBrush, @pRegion)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' The following example creates a region from a rectangle and then uses a path to update the region.
' ========================================================================================
SUB Example_ExcludeRegion (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)

   DIM pts(0 TO 5) AS GpPoint = {GDIP_POINT(110, 20), GDIP_POINT(120, 30), GDIP_POINT(100, 60), GDIP_POINT(120, 70), GDIP_POINT(150, 60), GDIP_POINT(140, 10)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(110, 20), (120, 30), (100, 60), (120, 70), (150, 60), (140, 10)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 5) AS GpPoint
'   pts(0).x = 110 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 30 : pts(2).x = 100 : pts(2).y = 60
'   pts(3).x = 120 : pts(3).y = 70 : pts(4).x = 150 : pts(4).y = 60 : pts(5).x = 140 : pts(5).y = 10
'#endif
   DIM pPath AS CGpGraphicsPath
   pPath.AddClosedCurve(@pts(0), 6)

   DIM rcf AS GpRectF = GDIP_RECTF(65.3, 15.1, 70.0, 45.8)
'#ifdef __FB_64BIT__
'   DIM rcf AS GpRectF = (65.3, 15.1, 70.0, 45.8)
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpRect is defined as Rect in 64 bit and as Rect_ in 32 bit.
'   DIM rcf AS GpRectF : rcf.x = 65.3 : rcf.y = 15.1 : rcf.Width = 70.0 : rcf.Height = 45.8
'#endif

   ' // Create a region from a rectangle
   DIM pRegion AS CGpRegion = @rcf

   ' // Exclude the intersecting portion of the path interior from the region.
   pRegion.Exclude(@pPath)
   graphics.FillRegion(@solidBrush, @pRegion)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' The following example creates two regions, one from a rectangle and the other from a path.
' Then the code uses the rectangular region to update the path region.
' ========================================================================================
SUB Example_ExcludeRegion (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)

   DIM pts(0 TO 5) AS GpPoint = {GDIP_POINT(110, 20), GDIP_POINT(120, 30), GDIP_POINT(100, 60), GDIP_POINT(120, 70), GDIP_POINT(150, 60), GDIP_POINT(140, 10)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(110, 20), (120, 30), (100, 60), (120, 70), (150, 60), (140, 10)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 5) AS GpPoint
'   pts(0).x = 110 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 30 : pts(2).x = 100 : pts(2).y = 60
'   pts(3).x = 120 : pts(3).y = 70 : pts(4).x = 150 : pts(4).y = 60 : pts(5).x = 140 : pts(5).y = 10
'#endif
   DIM pPath AS CGpGraphicsPath
   pPath.AddClosedCurve(@pts(0), 6)

   DIM rcf AS GpRectF = GDIP_RECTF(65.3, 15.1, 70.0, 45.8)
'#ifdef __FB_64BIT__
'   DIM rcf AS GpRectF = (65.3, 15.1, 70.0, 45.8)
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpRect is defined as Rect in 64 bit and as Rect_ in 32 bit.
'   DIM rcf AS GpRectF : rcf.x = 65.3 : rcf.y = 15.1 : rcf.Width = 70.0 : rcf.Height = 45.8
'#endif

   ' // Create a region from a rectangle
   DIM rectRegion AS CGpRegion = @rcf

  ' // Create a region from a path.
   DIM pathRegion AS CGpRegion = @pPath

   ' // Exclude a rectangle region from the path region.
   pathRegion.Exclude(@rectRegion)
   graphics.FillRegion(@solidBrush, @pathRegion)

END SUB
' ========================================================================================
```

# <a name="GetBounds"></a>GetBounds

Gets a rectangle that encloses this region.

```
FUNCTION GetBounds (BYVAL rc AS GpRectF PTR, BYVAL pGraphics AS CGpGraphics PTR) AS GpStatus
FUNCTION GetBounds (BYVAL rc AS GpRect PTR, BYVAL pGraphics AS CGpGraphics PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *rc* | Pointer to a **GpRectF** or **GpRect** structure that receives the enclosing rectangle. |
| *pGraphics* | Pointer to a **Graphics** object that contains the world and page transformations required to calculate the device coordinates of this region and the rectangle. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a region from a path, gets the region's enclosing rectangle,
' and then displays both.
' ========================================================================================
SUB Example_GetBoundsRect (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)
s
   DIM pts(0 TO 5) AS GpPoint = {GDIP_POINT(110, 20), GDIP_POINT(120, 30), GDIP_POINT(100, 60), GDIP_POINT(120, 70), GDIP_POINT(150, 60), GDIP_POINT(140, 10)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(110, 20), (120, 30), (100, 60), (120, 70), (150, 60), (140, 10)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 5) AS GpPoint
'   pts(0).x = 110 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 30 : pts(2).x = 100 : pts(2).y = 60
'   pts(3).x = 120 : pts(3).y = 70 : pts(4).x = 150 : pts(4).y = 60 : pts(5).x = 140 : pts(5).y = 10
'#endif
   DIM path AS CGpGraphicsPath
   path.AddClosedCurve(@pts(0), 6)

   ' // Create a region from a path
   DIM pathRegion AS CGpRegion = @path

   ' // Get the region's enclosing rectangle.
   DIM rc AS GpRect
   pathRegion.GetBounds(@rc, @graphics)

   ' // Show the region and the enclosing rectangle.
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   DIM pen AS CGpPen = GDIP_ARGB(255, 0, 0, 0)
   graphics.FillRegion(@solidBrush, @pathRegion)
   graphics.DrawRectangle(@pen, @rc)

END SUB
' ========================================================================================
```

# <a name="GetData"></a>GetData

Gets data that describes this region.

```
FUNCTION GetData (BYVAL buffer AS BYTE PTR, BYVAL bufferSize AS UINT, _
   BYVAL sizeFilled AS UINT PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *buffer* | Out. Pointer to an array of BYTE values that receives the region data. |
| *bufferSize* | Integer that specifies the size, in bytes, of the buffer array. The size of the buffer array can be greater than or equal to the number of bytes required to store the region data. The exact number of bytes required can be determined by calling the Region.GetDataSize method. |
| *sizeFilled* | Out. Pointer to a UINT variable that receives the number of bytes of data actually received by the buffer array. The default value is NULL. |

#### Example

```
' ========================================================================================
' The following example creates a region from a path and then gets the data that describes
' the region.
' ========================================================================================
SUB Example_GetData (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM pts(0 TO 5) AS GpPoint = {GDIP_POINT(110, 20), GDIP_POINT(120, 30), GDIP_POINT(100, 60), GDIP_POINT(120, 70), GDIP_POINT(150, 60), GDIP_POINT(140, 10)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(110, 20), (120, 30), (100, 60), (120, 70), (150, 60), (140, 10)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 5) AS GpPoint
'   pts(0).x = 110 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 30 : pts(2).x = 100 : pts(2).y = 60
'   pts(3).x = 120 : pts(3).y = 70 : pts(4).x = 150 : pts(4).y = 60 : pts(5).x = 140 : pts(5).y = 10
'#endif
   DIM path AS CGpGraphicsPath
   path.AddClosedCurve(@pts(0), 6)

   ' // Create a region from a path
   DIM pathRegion AS CGpRegion = @path

   ' // Get the region data.
   DIM bufferSize AS UINT
   DIM sizeFilled AS UINT
   DIM pData AS BYTE PTR
   
   bufferSize = pathRegion.GetDataSize
   pData = Allocate(bufferSize)
   pathRegion.GetData(pData, bufferSize, @sizeFilled)
   
   ' // Inspect or use the region data.
   ' ...
   Delete pData

END SUB
' ========================================================================================
```

# <a name="GetDataSize"></a>GetDataSize

Gets the number of bytes of data that describes this region.

```
FUNCTION GetDataSize () AS UINT
```

#### Example

```
' ========================================================================================
' The following example creates a region from a path and then gets the data that describes
' the region.
' ========================================================================================
SUB Example_GetDataSize (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM pts(0 TO 5) AS GpPoint = {GDIP_POINT(110, 20), GDIP_POINT(120, 30), GDIP_POINT(100, 60), GDIP_POINT(120, 70), GDIP_POINT(150, 60), GDIP_POINT(140, 10)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(110, 20), (120, 30), (100, 60), (120, 70), (150, 60), (140, 10)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 5) AS GpPoint
'   pts(0).x = 110 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 30 : pts(2).x = 100 : pts(2).y = 60
'   pts(3).x = 120 : pts(3).y = 70 : pts(4).x = 150 : pts(4).y = 60 : pts(5).x = 140 : pts(5).y = 10
'#endif
   DIM path AS CGpGraphicsPath
   path.AddClosedCurve(@pts(0), 6)

   ' // Create a region from a path
   DIM pathRegion AS CGpRegion = @path

   ' // Get the region data.
   DIM bufferSize AS UINT
   DIM sizeFilled AS UINT
   DIM pData AS BYTE PTR
   
   bufferSize = pathRegion.GetDataSize
   pData = Allocate(bufferSize)
   pathRegion.GetData(pData, bufferSize, @sizeFilled)
   
   ' // Inspect or use the region data.
   ' ...
   Delete pData

END SUB
' ========================================================================================
```

# <a name="GetHRGN"></a>GetHRGN

Creates a Windows Graphics Device Interface (GDI) region from this region.

```
FUNCTION GetHRGN (BYVAL pGraphics AS CGpGraphics PTR) AS HRGN
```

| Parameter  | Description |
| ---------- | ----------- |
| *pGraphics* | Pointer to a Graphics object that contains the world and page transformations required to calculate the device coordinates of this region. |

#### Return value

This method returns a Windows handle to a GDI region that is created from this region.

#### Remarks

It is the caller's responsibility to call the GDI function **DeleteObject** to free the GDI region when it is no longer needed.

#### Example

```
' ========================================================================================
' The following example creates a GDI+ region from a path and then uses the GDI+ region to
' create a GDI region. The code then uses a GDI function to display the GDI region.
' ========================================================================================
SUB Example_GetHRGN (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM pts(0 TO 5) AS GpPoint = {GDIP_POINT(110, 20), GDIP_POINT(120, 30), GDIP_POINT(100, 60), GDIP_POINT(120, 70), GDIP_POINT(150, 60), GDIP_POINT(140, 10)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(110, 20), (120, 30), (100, 60), (120, 70), (150, 60), (140, 10)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 5) AS GpPoint
'   pts(0).x = 110 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 30 : pts(2).x = 100 : pts(2).y = 60
'   pts(3).x = 120 : pts(3).y = 70 : pts(4).x = 150 : pts(4).y = 60 : pts(5).x = 140 : pts(5).y = 10
'#endif

   DIM pPath AS CGpGraphicsPath
   pPath.AddClosedCurve(@pts(0), 6)

   ' // Create a region from a path
   DIM pathRegion AS CGpRegion = @pPath

   ' // Get a handle to a GDI region.
   DIM hRegion AS HRGN = pathRegion.GetHRGN(@graphics)

   ' // Use GDI to display the region.
   DIM hBrush AS HBRUSH = CreateSolidBrush(BGR(255, 0, 0))
   FillRgn(hdc, hRegion, hBrush)

   IF hBrush THEN DeleteObject(hBrush)
   IF hRegion THEN DeleteObject(hRegion)

END SUB
' ========================================================================================
```

# <a name="GetRegionScans"></a>GetRegionScans

Gets an array of rectangles that approximate this region. The region is transformed by a specified matrix before the rectangles are calculated.

```
FUNCTION GetRegionScans (BYVAL pMatrix AS CGpMatrix PTR, BYVAL rects AS GpRectF PTR, _
   BYVAL count AS INT_ PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMatrix* | Pointer to a **Matrix** object that is used to transform the region. |
| *rects* | Out. Pointer to an array of **GpRectF** objects that receives the rectangles. |
| *count* | Out. Pointer to a LONG that receives a value that indicates the number of rectangles that approximate this region. The value is valid even if rects is a NULL pointer. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

The **GetRegionScansCount** method can be used first to determine the number of rectangles. Then, you can allocate a buffer that is the correct size and set the rects parameter to point to the buffer.

#### Example

```
' ========================================================================================
' The following example creates a region from a path and gets a set of rectangles that
' approximate the region. The code then draws each of the rectangles.
' ========================================================================================
SUB Example_GetRegionScans (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96

   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   DIM pen AS CGpPen = GDIP_ARGB(255, 0, 0, 0)
   DIM pPath AS CGpGraphicsPath
   DIM matrix AS CGpMatrix

   ' // Create a region from a path
   pPath.AddEllipse(10 * rxRatio, 10 * rxRatio, 50 * rxRatio, 300 * rxRatio)
   DIM pathRegion AS CGpRegion = @pPath
   graphics.FillRegion(@solidBrush, @pathRegion)

   ' // Get the rectangles
   graphics.GetTransform(@matrix)
   DIM nCount AS LONG = pathRegion.GetRegionScansCount(@matrix)
   DIM rects(nCount - 1) AS GpRectF
   pathRegion.GetRegionScans(@matrix, @rects(0), @nCount)

   ' // Draw the rectangles.
   FOR j AS LONG = 0 TO nCount - 1
      graphics.DrawRectangle(@pen, @rects(j))
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetRegionScansCount"></a>GetRegionScansCount

Gets the number of rectangles that approximate this region. The region is transformed by a specified matrix before the rectangles are calculated.

```
FUNCTION GetRegionScansCount (BYVAL pMatrix AS CGpMatrix PTR) AS UINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMatrix* | Pointer to a **Matrix** object that is used to transform the region. |

#### Remarks

The **GetRegionScansCount** method can be used before the **GetRegionScans** method to determine the number of rectangles. Then, you can allocate a buffer that is the correct size to store the rectangles that are obtained with the **GetRegionScans** method.

#### Example

```
' ========================================================================================
' The following example creates a region from a path and gets a set of rectangles that
' approximate the region. The code then draws each of the rectangles.
' ========================================================================================
SUB Example_GetRegionScansCount (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96

   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   DIM pen AS CGpPen = GDIP_ARGB(255, 0, 0, 0)
   DIM pPath AS CGpGraphicsPath
   DIM matrix AS CGpMatrix

   ' // Create a region from a path
   pPath.AddEllipse(10 * rxRatio, 10 * rxRatio, 50 * rxRatio, 300 * rxRatio)
   DIM pathRegion AS CGpRegion = @pPath
   graphics.FillRegion(@solidBrush, @pathRegion)

   ' // Get the rectangles
   graphics.GetTransform(@matrix)
   DIM nCount AS LONG = pathRegion.GetRegionScansCount(@matrix)
   DIM rects(nCount - 1) AS GpRectF
   pathRegion.GetRegionScans(@matrix, @rects(0), @nCount)

   ' // Draw the rectangles.
   FOR j AS LONG = 0 TO nCount - 1
      graphics.DrawRectangle(@pen, @rects(j))
   NEXT

END SUB
' ========================================================================================
```

# <a name="Intersect"></a>Intersect

Updates this region to the portion of itself that intersects the specified rectangle's interior.

```
FUNCTION Intersect (BYVAL rc AS GpRectF PTR) AS GpStatus
FUNCTION Intersect (BYVAL rc AS GpRect PTR) AS GpStatus
```

Updates this region to the portion of itself that intersects the specified path's interior.

```
FUNCTION Intersect (BYVAL pPath AS CGpGraphicsPath PTR) AS GpStatus
```

Updates this region to the portion of itself that intersects another region.

```
FUNCTION Intersect (BYVAL pRegion AS CGpRegion PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *rc* | Reference to a rectangle to use to update this **Region* object. |
| *pPath* | Pointer to a **GraphicsPath** object that specifies the path to use to update this **Region** object. |
| *pRegion* | Pointer to a **Region** object to use to update this **Region** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a region from a path and then uses a rectangle to update the region.
' ========================================================================================
SUB Example_IntersectRegion (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)

   DIM pts(0 TO 5) AS GpPoint = {GDIP_POINT(110, 20), GDIP_POINT(120, 30), GDIP_POINT(100, 60), GDIP_POINT(120, 70), GDIP_POINT(150, 60), GDIP_POINT(140, 10)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(110, 20), (120, 30), (100, 60), (120, 70), (150, 60), (140, 10)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 5) AS GpPoint
'   pts(0).x = 110 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 30 : pts(2).x = 100 : pts(2).y = 60
'   pts(3).x = 120 : pts(3).y = 70 : pts(4).x = 150 : pts(4).y = 60 : pts(5).x = 140 : pts(5).y = 10
'#endif
   DIM pPath AS CGpGraphicsPath
   pPath.AddClosedCurve(@pts(0), 6)

   DIM rcf AS GpRectF = GDIP_RECTF(65.2, 15.0, 70.9, 45.0)
'#ifdef __FB_64BIT__
'   DIM rcf AS GpRectF = (65.2, 15.0, 70.9, 45.0)
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpRect is defined as Rect in 64 bit and as Rect_ in 32 bit.
'   DIM rcf AS GpRectF : rcf.x = 65.2 : rcf.y = 15.0 : rcf.Width = 70.9 : rcf.Height = 45.0
'#endif

   ' // Create a region from a path
   DIM pRegion AS CGpRegion = @pPath

   ' // Update the region to the portion that intersects with the rectangle
   pRegion.Intersect(@rcf)
   graphics.FillRegion(@solidBrush, @pRegion)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' The following example creates a region from a rectangle and then uses a path to update the region.
' ========================================================================================
SUB Example_IntersectRegion (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)

   DIM pts(0 TO 5) AS GpPoint = {GDIP_POINT(110, 20), GDIP_POINT(120, 30), GDIP_POINT(100, 60), GDIP_POINT(120, 70), GDIP_POINT(150, 60), GDIP_POINT(140, 10)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(110, 20), (120, 30), (100, 60), (120, 70), (150, 60), (140, 10)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 5) AS GpPoint
'   pts(0).x = 110 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 30 : pts(2).x = 100 : pts(2).y = 60
'   pts(3).x = 120 : pts(3).y = 70 : pts(4).x = 150 : pts(4).y = 60 : pts(5).x = 140 : pts(5).y = 10
'#endif
   DIM pPath AS CGpGraphicsPath
   pPath.AddClosedCurve(@pts(0), 6)

   DIM rcf AS GpRectF = GDIP_RECTF(65.2, 15.0, 70.9, 45.0)
'#ifdef __FB_64BIT__
'   DIM rcf AS GpRectF = (65.2, 15.0, 70.9, 45.0)
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpRect is defined as Rect in 64 bit and as Rect_ in 32 bit.
'   DIM rcf AS GpRectF : rcf.x = 65.2 : rcf.y = 15.0 : rcf.Width = 70.9 : rcf.Height = 45.0
'#endif

   ' // Create a region from a rectangle
   DIM pRegion AS CGpRegion = @rcf

   ' // Update the region to the portion that intersects with the rectangle
   pRegion.Intersect(@pPath)
   graphics.FillRegion(@solidBrush, @pRegion)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' The following example creates two regions, one from a path and the other from a rectangle.
' The code then uses the rectangular region to update the path region.
' ========================================================================================
SUB Example_IntersectRegion (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)

   DIM pts(0 TO 5) AS GpPoint = {GDIP_POINT(110, 20), GDIP_POINT(120, 30), GDIP_POINT(100, 60), GDIP_POINT(120, 70), GDIP_POINT(150, 60), GDIP_POINT(140, 10)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(110, 20), (120, 30), (100, 60), (120, 70), (150, 60), (140, 10)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 5) AS GpPoint
'   pts(0).x = 110 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 30 : pts(2).x = 100 : pts(2).y = 60
'   pts(3).x = 120 : pts(3).y = 70 : pts(4).x = 150 : pts(4).y = 60 : pts(5).x = 140 : pts(5).y = 10
'#endif
   DIM pPath AS CGpGraphicsPath
   pPath.AddClosedCurve(@pts(0), 6)

   DIM rcf AS GpRectF = GDIP_RECTF(65.2, 15.0, 70.9, 45.0)
'#ifdef __FB_64BIT__
'   DIM rcf AS GpRectF = (65.2, 15.0, 70.9, 45.0)
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpRect is defined as Rect in 64 bit and as Rect_ in 32 bit.
'   DIM rcf AS GpRectF : rcf.x = 65.2 : rcf.y = 15.0 : rcf.Width = 70.9 : rcf.Height = 45.0
'#endif

   ' // Create a region from a rectangle
   DIM rectRegion AS CGpRegion = @rcf

   ' // Create a region from a path
   DIM pathRegion AS CGpRegion = @pPath

   ' // Update the region to the portion that intersects with another region
   pathRegion.Intersect(@rectRegion)
   graphics.FillRegion(@solidBrush, @pathRegion)

END SUB
' ========================================================================================
```

# <a name="IsEmpty"></a>IsEmpty

Determines whether this region is empty.

```
FUNCTION IsEmpty (BYVAL pGraphics AS CGpGraphics PTR) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *pGraphics* | Pointer to a **Graphics** object that contains the world and page transformations required to calculate the device coordinates of this region. |

#### Return value

If this region is empty, this method returns TRUE; otherwise, it returns FALSE.

#### Example

```
' ========================================================================================
' The following example creates two regions, one from a path and the other from a rectangle.
' The code then updates the path region with the rectangular region and tests the updated
' region to determine if it is empty.
' ========================================================================================
SUB Example_IsEmpty (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM pts(0 TO 5) AS GpPoint = {GDIP_POINT(110, 20), GDIP_POINT(120, 30), GDIP_POINT(100, 60), GDIP_POINT(120, 70), GDIP_POINT(150, 60), GDIP_POINT(140, 10)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(110, 20), (120, 30), (100, 60), (120, 70), (150, 60), (140, 10)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 5) AS GpPoint
'   pts(0).x = 110 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 30 : pts(2).x = 100 : pts(2).y = 60
'   pts(3).x = 120 : pts(3).y = 70 : pts(4).x = 150 : pts(4).y = 60 : pts(5).x = 140 : pts(5).y = 10
'#endif

   DIM path AS CGpGraphicsPath
   path.AddClosedCurve(@pts(0), 6)

   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   DIM rc AS GpRect = GDIP_RECT(5, 15, 70, 45)
'#ifdef __FB_64BIT__
'   DIM rc AS GpRect = (5, 15, 70, 45)
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpRect is defined as Rect in 64 bit and as Rect_ in 32 bit.
'   DIM rc AS GpRect : rcf.x = 5 : rc.y = 15 : rc.Width = 70 : rc.Height = 45
'#endif

   ' // Create a region from a path
   DIM pathRegion AS CGpRegion = @path
   graphics.FillRegion(@solidBrush, @pathRegion)

   ' // Create a region from a rectangle.
   DIM rectRegion AS CGpRegion = @rc
   graphics.FillRegion(@solidBrush, @rectRegion)

   ' // Update the path region to the portion that intersects the rectangular region.
   pathRegion.Intersect(@rectRegion)
   IF pathRegion.IsEmpty(@graphics) THEN
      '// The intersection is empty.
      PRINT "Region is empty"
   END IF
   
END SUB
' ========================================================================================
```

# <a name="IsInfinite"></a>IsInfinite

Determines whether this region is infinite.

```
FUNCTION IsInfinite (BYVAL pGraphics AS CGpGraphics PTR) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *pGraphics* | Pointer to a **Graphics** object that contains the world and page transformations required to calculate the device coordinates of this region. |

#### Return value

If this region is infinite, this method returns TRUE; otherwise, it returns FALSE.

#### Example

```
' ========================================================================================
' The following example uses the default constructor to create an infinite region. Then the
' code tests the region to determine if it is infinite and performs a task based on the result.
' ========================================================================================
SUB Example_IsInfinite (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96

   graphics.SetPageUnit(UnitInch)
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)

   DIM rc AS GpRect = GDIP_RECT(0, 0, 100, 50)
'#ifdef __FB_64BIT__
'   DIM rc AS GpRect = (0, 0, 100, 50)
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpRect is defined as Rect in 64 bit and as Rect_ in 32 bit.
'   DIM rc AS GpRect : rc.Width = 100 : rc.Height = 50
'#endif

   ' // The default constructor creates an infinite region.
   DIM region AS CGpRegion

   IF region.IsInfinite(@graphics) THEN
      ' // The region is infinite. Perform some task.
      PRINT "Region is infinite"
   ELSE
      ' // Perform a different task.
   END IF

   graphics.FillRegion(@solidBrush, @region)

END SUB
' ========================================================================================
```

# <a name="IsVisible"></a>IsVisible

Determines whether a point is inside this region.

```
FUNCTION IsVisible (BYVAL pt AS GpPointF PTR, BYVAL pGraphics AS CGpGraphics PTR) AS GpStatus
FUNCTION IsVisible (BYVAL pt AS GpPoint PTR, BYVAL pGraphics AS CGpGraphics PTR) AS GpStatus
FUNCTION IsVisible (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL pGraphics AS CGpGraphics PTR) AS GpStatus
FUNCTION IsVisible (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL pGraphics AS CGpGraphics PTR) AS GpStatus
FUNCTION IsVisible (BYVAL pt AS GpRectF PTR, BYVAL pGraphics AS CGpGraphics PTR) AS GpStatus
FUNCTION IsVisible (BYVAL pt AS GpRect PTR, BYVAL pGraphics AS CGpGraphics PTR) AS GpStatus
FUNCTION IsVisible (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, _
   BYVAL nHeight AS SINGLE, BYVAL pGraphics AS CGpGraphics PTR) AS GpStatus
FUNCTION IsVisible (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, _
   BYVAL nHeight AS INT_, BYVAL pGraphics AS CGpGraphics PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pt* | Reference to a point to test. |
| *rc* | Reference to a rectangle to test. |
| *x* | x-coordinate of the upper-left corner of the rectangle to test. |
| *x* | y-coordinate of the upper-left corner of the rectangle to test. |
| *nWidth* | The width of the rectangle to test |
| *nHeight* | The height of the rectangle to test |
| *pGraphics* | Optional. Pointer to a **Graphics** object that contains the world and page transformations required to calculate the device coordinates of this region and the point. The default value is NULL. |

#### Return value

If the point is inside this region, this method returns TRUE; otherwise, it returns FALSE.

#### Example

```
' ========================================================================================
' The following example creates a region from a path and then tests to determine whether a
' rectangle intersects the region.
' ========================================================================================
SUB Example_IsVisible (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM pts(0 TO 5) AS GpPoint = {GDIP_POINT(110, 20), GDIP_POINT(120, 30), GDIP_POINT(100, 60), GDIP_POINT(120, 70), GDIP_POINT(150, 60), GDIP_POINT(140, 10)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(110, 20), (120, 30), (100, 60), (120, 70), (150, 60), (140, 10)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 5) AS GpPoint
'   pts(0).x = 110 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 30 : pts(2).x = 100 : pts(2).y = 60
'   pts(3).x = 120 : pts(3).y = 70 : pts(4).x = 150 : pts(4).y = 60 : pts(5).x = 140 : pts(5).y = 10
'#endif

   DIM path AS CGpGraphicsPath
   path.AddClosedCurve(@pts(0), 6)

   ' // Create a region from a path
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   DIM pathRegion AS CGpRegion = @path
   graphics.FillRegion(@solidBrush, @pathRegion)

   ' // Check to see whether a rectangle intersects the region.
   DIM rc AS GpRect = GDIP_RECT(55, 25, 70, 30)
'#ifdef __FB_64BIT__
'   DIM rc AS GpRect = (65, 25, 70, 30)
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpRect is defined as Rect in 64 bit and as Rect_ in 32 bit.
'   DIM rc AS GpRect : rc.x = 65 : rc.y = 25 : rc.Width = 70 : rc.Height = 30
'#endif

   IF pathRegion.IsVisible(@rc, @graphics) THEN
      ' // All or part of the rectangle is in the region.
      PRINT "Is visible"
   END IF

   ' // Draw the test rectangle.
   DIM pen AS CGpPen = GDIP_ARGB(255, 0, 0, 0)
   graphics.DrawRectangle(@pen, @rc)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' The following example creates a RectF structure and determines whether the rectangle is
' visible or partially visible on the display device. If it is, it fills the rectangle.
' ========================================================================================
SUB Example_IsVisibleRect (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a RectF structure
   DIM visibleRect AS GpRectF = GDIP_RECTF(0.0, 0.0, 100.2, 100.2)

   ' // Determine whether the rectangle is visible and, if it is, fill it.
   IF graphics.IsVisible(@visibleRect) THEN
       graphics.FillRectangle(@CGpSolidBrush(GDIP_ARGB(255, 0, 0, 0)), @visibleRect)
   END IF

END SUB
' ========================================================================================
```

# <a name="MakeEmpty"></a>MakeEmpty

Updates this region to an empty region. In other words, the region occupies no space on the display device.

```
FUNCTION MakeEmpty () AS GpStatus
```

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a region from a rectangle and paints it with a brush. Then
' the code makes the region empty and paints the region with a different color to show that
' the region occupies no space on the display device.
' ========================================================================================
SUB Example_MakeEmptyRegion (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM rc AS GpRect = GDIP_RECT(65, 15, 70, 45)
'#ifdef __FB_64BIT__
'   DIM rc AS GpRect = (65, 15, 70, 45)
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpRect is defined as Rect in 64 bit and as Rect_ in 32 bit.
'   DIM rc AS GpRect : rc.x = 65 : rc.y = 15 : rc.Width = 70 : rc.Height = 45
'#endif

   ' // Create red and blue solid brushes
   DIM redBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   DIM blueBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 255)

   ' // Create a region, and fill it with a red brush.
   DIM rectRegion AS CGpRegion = @rc
   graphics.FillRegion(@redBrush, @rectRegion)

   ' // Make the region empty, and then fill it with a blue brush.
   ' // The color won't change because it is empty and can't be filled with FillRegion.
   rectRegion.MakeEmpty
   graphics.FillRegion(@blueBrush, @rectRegion)

END SUB
' ========================================================================================
```

# <a name="MakeInfinite"></a>MakeInfinite

Updates this region to an infinite region.

```
FUNCTION MakeInfinite () AS GpStatus
```

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a region from a rectangle, makes the region infinite, and
' fills the region to show the infinite region.
' ========================================================================================
SUB Example_MakeInfiniteRegion (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM rc AS GpRect = GDIP_RECT(65, 15, 70, 45)
'#ifdef __FB_64BIT__
'   DIM rc AS GpRect = (65, 15, 70, 45)
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpRect is defined as Rect in 64 bit and as Rect_ in 32 bit.
'   DIM rc AS GpRect : rc.x = 65 : rc.y = 15 : rc.Width = 70 : rc.Height = 45
'#endif

   ' // Create red and blue solid brushes
   DIM redBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   DIM blueBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 255)

   ' // Create a region, and fill it with a red brush
   DIM rectRegion AS CGpRegion = @rc
   graphics.FillRegion(@redBrush, @rectRegion)

   ' // Make the region infinite, and then fill it with a blue brush
   rectRegion.MakeInfinite
   graphics.FillRegion(@blueBrush, @rectRegion)

END SUB
' ========================================================================================
```

# <a name="Transform"></a>Transform

Transforms this region by multiplying each of its data points by a specified matrix.

```
FUNCTION Transform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMatrix* | Pointer to a matrix that specifies the transformation. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a region from a path and fills it. The code then transforms
' the region and fills the transformed region to show the transformation.
' ========================================================================================
SUB Example_TransformRegion (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM pts(0 TO 5) AS GpPoint = {GDIP_POINT(110, 20), GDIP_POINT(120, 30), GDIP_POINT(100, 60), GDIP_POINT(120, 70), GDIP_POINT(150, 60), GDIP_POINT(140, 10)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(110, 20), (120, 30), (100, 60), (120, 70), (150, 60), (140, 10)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 5) AS GpPoint
'   pts(0).x = 110 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 30 : pts(2).x = 100 : pts(2).y = 60
'   pts(3).x = 120 : pts(3).y = 70 : pts(4).x = 150 : pts(4).y = 60 : pts(5).x = 140 : pts(5).y = 10
'#endif

   DIM rcf AS GpRectF = GDIP_RECTF(65.3, 15.1, 70.0, 45.0)
'#ifdef __FB_64BIT__
'   DIM rcf AS GpRectF = (65.3, 15.1, 70.0, 45.0)
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpRect is defined as Rect in 64 bit and as Rect_ in 32 bit.
'   DIM rcf AS GpRectF : rcf.x = 65.3 : rcf.y = 15.1 : rcf.Width = 70.0 : rcf.Height = 45.0
'#endif

   DIM pPath AS CGpGraphicsPath
   pPath.AddClosedCurve(@pts(0), 6)

   ' // Create a region from a path.
   DIM pathRegion AS CGpRegion = @pPath

   ' // Transform the region.
   DIM matrix AS CGpMatrix = CGpMatrix(1, 0, 0, 3, 100, 0)   ' // vertical stretch, shift right
   pathRegion.Transform(@matrix)
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   graphics.FillRegion(@solidBrush, @pathRegion)

END SUB
' ========================================================================================
```

# <a name="Translate"></a>Translate

Offsets this region by specified amounts in the horizontal and vertical directions.

```
FUNCTION Translate (BYVAL dx AS SINGLE, BYVAL dy AS SINGLE) AS GpStatus
FUNCTION Translate (BYVAL dx AS INT_, BYVAL dy AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *dx* | The amount to shift the region in the x direction. |
| *dy* | The amount to shift the region in the y direction. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a region from a path and fills it. The code then
' translates the region and fills the translated region to show how the region has shifted.
' ========================================================================================
SUB Example_TranslateRegion (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM pts(0 TO 5) AS GpPoint = {GDIP_POINT(110, 20), GDIP_POINT(120, 30), GDIP_POINT(100, 60), GDIP_POINT(120, 70), GDIP_POINT(150, 60), GDIP_POINT(140, 10)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(110, 20), (120, 30), (100, 60), (120, 70), (150, 60), (140, 10)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 5) AS GpPoint
'   pts(0).x = 110 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 30 : pts(2).x = 100 : pts(2).y = 60
'   pts(3).x = 120 : pts(3).y = 70 : pts(4).x = 150 : pts(4).y = 60 : pts(5).x = 140 : pts(5).y = 10
'#endif

   DIM rcf AS GpRectF = GDIP_RECTF(65.3, 15.1, 70.0, 45.0)
'#ifdef __FB_64BIT__
'   DIM rcf AS GpRectF = (65.3, 15.1, 70.0, 45.0)
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpRect is defined as Rect in 64 bit and as Rect_ in 32 bit.
'   DIM rcf AS GpRectF : rcf.x = 65.3 : rcf.y = 15.1 : rcf.Width = 70.0 : rcf.Height = 45.0
'#endif

   DIM pPath AS CGpGraphicsPath
   pPath.AddClosedCurve(@pts(0), 6)

   ' // Create a region from a path.
   DIM pathRegion AS CGpRegion = @pPath

   ' // Translate the region.
   DIM matrix AS CGpMatrix = CGpMatrix(1, 0, 0, 3, 100, 0)   ' // vertical stretch, shift right
   pathRegion.Translate(100, 60)
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   graphics.FillRegion(@solidBrush, @pathRegion)

END SUB
' ========================================================================================
```

# <a name="Union_"></a>Union_

Updates this region to all portions (intersecting and nonintersecting) of itself and all portions of the specified rectangle's interior.

```
FUNCTION Union_ (BYVAL rc AS GpRectF PTR) AS GpStatus
FUNCTION Union_ (BYVAL rc AS GpRect PTR) AS GpStatus
```

Updates this region to all portions (intersecting and nonintersecting) of itself and all portions of the specified path's interior.

```
FUNCTION Union_ (BYVAL pPath AS CGpGraphicsPath PTR) AS GpStatus
```

Updates this region to all portions (intersecting and nonintersecting) of itself and all portions of another region.

```
FUNCTION Union_ (BYVAL pRegion AS CGpRegion PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *rc* | Reference to a rectangle to use to update this **Region* object. |
| *pPath* | Pointer to a **GraphicsPath** object that specifies the path to use to update this **Region** object. |
| *pRegion* | Pointer to a **Region** object to use to update this **Region** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a region from a path and then uses a rectangle to update
' the region. The updated region is the union of the path region and a rectangle.
' ========================================================================================
SUB Example_UnionRegion (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)

   DIM pts(0 TO 5) AS GpPoint = {GDIP_POINT(110, 20), GDIP_POINT(120, 30), GDIP_POINT(100, 60), GDIP_POINT(120, 70), GDIP_POINT(150, 60), GDIP_POINT(140, 10)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(110, 20), (120, 30), (100, 60), (120, 70), (150, 60), (140, 10)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 5) AS GpPoint
'   pts(0).x = 110 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 30 : pts(2).x = 100 : pts(2).y = 60
'   pts(3).x = 120 : pts(3).y = 70 : pts(4).x = 150 : pts(4).y = 60 : pts(5).x = 140 : pts(5).y = 10
'#endif

   DIM rcf AS GpRectF = GDIP_RECTF(65.3, 15.1, 70.0, 45.8)
'#ifdef __FB_64BIT__
'   DIM rcf AS GpRectF = (65.3, 15.1, 70.0, 45.8)
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpRect is defined as Rect in 64 bit and as Rect_ in 32 bit.
'   DIM rcf AS GpRectF : rcf.x = 65.3 : rcf.y = 15.1 : rcf.Width = 70.0 : rcf.Height = 45.8
'#endif

   DIM pPath AS CGpGraphicsPath
   pPath.AddClosedCurve(@pts(0), 6)

   ' // Create a region from a path.
   DIM pRegion AS CGpRegion = @pPath

   ' // Form the union of the region and a rectangle
   pRegion.Union_(@rcf)
   graphics.FillRegion(@solidBrush, @pRegion)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' The following example creates a region from a rectangle and then uses a path to update
' the region. The updated region is the union of the rectangular region and the path.
' ========================================================================================
SUB Example_UnionRegion (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)

   DIM pts(0 TO 5) AS GpPoint = {GDIP_POINT(110, 20), GDIP_POINT(120, 30), GDIP_POINT(100, 60), GDIP_POINT(120, 70), GDIP_POINT(150, 60), GDIP_POINT(140, 10)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(110, 20), (120, 30), (100, 60), (120, 70), (150, 60), (140, 10)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 5) AS GpPoint
'   pts(0).x = 110 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 30 : pts(2).x = 100 : pts(2).y = 60
'   pts(3).x = 120 : pts(3).y = 70 : pts(4).x = 150 : pts(4).y = 60 : pts(5).x = 140 : pts(5).y = 10
'#endif

   DIM rcf AS GpRectF = GDIP_RECTF(65.3, 15.1, 70.0, 45.8)
'#ifdef __FB_64BIT__
'   DIM rcf AS GpRectF = (65.3, 15.1, 70.0, 45.8)
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpRect is defined as Rect in 64 bit and as Rect_ in 32 bit.
'   DIM rcf AS GpRectF : rcf.x = 65.3 : rcf.y = 15.1 : rcf.Width = 70.0 : rcf.Height = 45.8
'#endif

   DIM pPath AS CGpGraphicsPath
   pPath.AddClosedCurve(@pts(0), 6)

   ' // Create a region from a rectangle
   DIM pRegion AS CGpRegion = @rcf

   ' // Exclude the intersecting portion of the path interior from the region.
   pRegion.Union_(@pPath)
   graphics.FillRegion(@solidBrush, @pRegion)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' The following example creates two regions, one from a path and the other from a rectangle.
' The code then uses the rectangular region to update the path region. The updated region
' is the union of the two regions.
' ========================================================================================
SUB Example_UnionRegion (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)

   DIM pts(0 TO 5) AS GpPoint = {GDIP_POINT(110, 20), GDIP_POINT(120, 30), GDIP_POINT(100, 60), GDIP_POINT(120, 70), GDIP_POINT(150, 60), GDIP_POINT(140, 10)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(110, 20), (120, 30), (100, 60), (120, 70), (150, 60), (140, 10)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 5) AS GpPoint
'   pts(0).x = 110 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 30 : pts(2).x = 100 : pts(2).y = 60
'   pts(3).x = 120 : pts(3).y = 70 : pts(4).x = 150 : pts(4).y = 60 : pts(5).x = 140 : pts(5).y = 10
'#endif

   DIM rcf AS GpRectF = GDIP_RECTF(65.3, 15.1, 70.0, 45.8)
'#ifdef __FB_64BIT__
'   DIM rcf AS GpRectF = (65.3, 15.1, 70.0, 45.8)
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpRect is defined as Rect in 64 bit and as Rect_ in 32 bit.
'   DIM rcf AS GpRectF : rcf.x = 65.3 : rcf.y = 15.1 : rcf.Width = 70.0 : rcf.Height = 45.8
'#endif

   DIM pPath AS CGpGraphicsPath
   pPath.AddClosedCurve(@pts(0), 6)

   ' // Create a region from a rectangle
   DIM rectRegion AS CGpRegion = @rcf

   ' // Create a region from a path.
   DIM pathRegion AS CGpRegion = @pPath

   ' // Exclude a rectangle region from the path region.
   pathRegion.Union_(@rectRegion)
   graphics.FillRegion(@solidBrush, @pathRegion)

END SUB
' ========================================================================================
```

# <a name="Xor_"></a>Xor_

Updates this region to the nonintersecting portions of itself and the specified rectangle's interior.

```
FUNCTION Xor_ (BYVAL rc AS GpRectF PTR) AS GpStatus
FUNCTION Xor_ (BYVAL rc AS GpRect PTR) AS GpStatus
```

Updates this region to the nonintersecting portions of itself and the specified path's interior.

```
FUNCTION Xor_ (BYVAL pPath AS CGpGraphicsPath PTR) AS GpStatus
```

Updates this region to the nonintersecting portions of itself and another region.

```
FUNCTION Xor_ (BYVAL pRegion AS CGpRegion PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *rc* | Reference to a rectangle to use to update this **Region* object. |
| *pPath* | Pointer to a **GraphicsPath** object that specifies the path to use to update this **Region** object. |
| *pRegion* | Pointer to a **Region** object to use to update this **Region** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' Updates a region to the nonintersecting portions of itself and the specified rectangle's interior.
' ========================================================================================
SUB Example_XorRegion (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)

   DIM pts(0 TO 5) AS GpPoint = {GDIP_POINT(110, 20), GDIP_POINT(120, 30), GDIP_POINT(100, 60), GDIP_POINT(120, 70), GDIP_POINT(150, 60), GDIP_POINT(140, 10)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(110, 20), (120, 30), (100, 60), (120, 70), (150, 60), (140, 10)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 5) AS GpPoint
'   pts(0).x = 110 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 30 : pts(2).x = 100 : pts(2).y = 60
'   pts(3).x = 120 : pts(3).y = 70 : pts(4).x = 150 : pts(4).y = 60 : pts(5).x = 140 : pts(5).y = 10
'#endif

   DIM rcf AS GpRectF = GDIP_RECTF(65.3, 15.1, 70.0, 45.8)
'#ifdef __FB_64BIT__
'   DIM rcf AS GpRectF = (65.3, 15.1, 70.0, 45.8)
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpRect is defined as Rect in 64 bit and as Rect_ in 32 bit.
'   DIM rcf AS GpRectF : rcf.x = 65.3 : rcf.y = 15.1 : rcf.Width = 70.0 : rcf.Height = 45.8
'#endif

   DIM pPath AS CGpGraphicsPath
   pPath.AddClosedCurve(@pts(0), 6)

   ' // Create a region from a path.
   DIM pRegion AS CGpRegion = @pPath

   ' // Perform an exclusive OR operation on the region and a rectangle.
   pRegion.Xor_(@rcf)
   graphics.FillRegion(@solidBrush, @pRegion)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' The following example creates a region from a path and then uses a rectangle to update the region.
' ========================================================================================
SUB Example_XorRegion (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)

   DIM pts(0 TO 5) AS GpPoint = {GDIP_POINT(110, 20), GDIP_POINT(120, 30), GDIP_POINT(100, 60), GDIP_POINT(120, 70), GDIP_POINT(150, 60), GDIP_POINT(140, 10)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(110, 20), (120, 30), (100, 60), (120, 70), (150, 60), (140, 10)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 5) AS GpPoint
'   pts(0).x = 110 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 30 : pts(2).x = 100 : pts(2).y = 60
'   pts(3).x = 120 : pts(3).y = 70 : pts(4).x = 150 : pts(4).y = 60 : pts(5).x = 140 : pts(5).y = 10
'#endif

   DIM rcf AS GpRectF = GDIP_RECTF(65.3, 15.1, 70.0, 45.8)
'#ifdef __FB_64BIT__
'   DIM rcf AS GpRectF = (65.3, 15.1, 70.0, 45.8)
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpRect is defined as Rect in 64 bit and as Rect_ in 32 bit.
'   DIM rcf AS GpRectF : rcf.x = 65.3 : rcf.y = 15.1 : rcf.Width = 70.0 : rcf.Height = 45.8
'#endif

   DIM pPath AS CGpGraphicsPath
   pPath.AddClosedCurve(@pts(0), 6)

   ' // Create a region from a rectangle
   DIM pRegion AS CGpRegion = @rcf

   ' // Perform an exclusive OR operation on the region and a rectangle.
   pRegion.Xor_(@pPath)
   graphics.FillRegion(@solidBrush, @pRegion)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' Updates a region to the nonintersecting portions of itself and the specified rectangle's interior.
' ========================================================================================
SUB Example_XorRegion (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)

   DIM pts(0 TO 5) AS GpPoint = {GDIP_POINT(110, 20), GDIP_POINT(120, 30), GDIP_POINT(100, 60), GDIP_POINT(120, 70), GDIP_POINT(150, 60), GDIP_POINT(140, 10)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 5) AS GpPoint = {(110, 20), (120, 30), (100, 60), (120, 70), (150, 60), (140, 10)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 5) AS GpPoint
'   pts(0).x = 110 : pts(0).y = 20 : pts(1).x = 120 : pts(1).y = 30 : pts(2).x = 100 : pts(2).y = 60
'   pts(3).x = 120 : pts(3).y = 70 : pts(4).x = 150 : pts(4).y = 60 : pts(5).x = 140 : pts(5).y = 10
'#endif

   DIM rcf AS GpRectF = GDIP_RECTF(65.3, 15.1, 70.0, 45.8)
'#ifdef __FB_64BIT__
'   DIM rcf AS GpRectF = (65.3, 15.1, 70.0, 45.8)
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpRect is defined as Rect in 64 bit and as Rect_ in 32 bit.
'   DIM rcf AS GpRectF : rcf.x = 65.3 : rcf.y = 15.1 : rcf.Width = 70.0 : rcf.Height = 45.8
'#endif

   DIM pPath AS CGpGraphicsPath
   pPath.AddClosedCurve(@pts(0), 6)

   ' // Create a region from a rectangle
   DIM rectRegion AS CGpRegion = @rcf

   ' // Create a region from a path
   DIM pathRegion AS CGpRegion = @pPath

   ' // Perform an exclusive OR operation on the region and a rectangle.
   pathRegion.Xor_(@rectRegion)
   graphics.FillRegion(@solidBrush, @pathRegion)

END SUB
' ========================================================================================
```
