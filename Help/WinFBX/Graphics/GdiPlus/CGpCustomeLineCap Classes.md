# CGpCustomeLineCap Class

The **CGpCustomLineCap** class encapsulates a custom line cap. A line cap defines the style of graphic used to draw the ends of a line. It can be various shapes, such as a square, circle, or diamond. A custom line cap is defined by the path that draws it. The path is drawn by using a **Pen** object to draw the outline of a shape or by using a **Brush** object to fill the interior. The cap can be used on either or both ends of the line. Spacing can be adjusted between the end caps and the line.

**Inherits from**: CGpBase.
**Imclude file**: CGpLineCaps.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorCustomLineCap) | Creates a **CustomLineCap** object. |
| [Clone](#Clone) | Copies the contents of the existing **CustomLineCap** object into a new **CustomLineCap** object. |
| [GetBaseCap](#GetBaseCap) | Gets the style of the base cap. |
| [GetBaseInset](#GetBaseInset) | Gets the distance between the base cap to the start of the line. |
| [GetStrokeCaps](#GetStrokeCaps) | Gets the end cap styles for both the start line cap and the end line cap. |
| [GetStrokeJoin](#GetStrokeJoin) | Returns the style of **LineJoin** used to join multiple lines in the same **GraphicsPath** object. |
| [GetWidthScale](#GetWidthScale) | Gets the value of the scale width. |
| [SetBaseCap](#SetBaseCap) | Sets the **LineCap** that appears as part of this **CustomLineCap** at the end of a line. |
| [SetBaseInset](#SetBaseInset) | Sets the base inset value of this custom line cap. |
| [SetStrokeCap](#SetStrokeCap) | Sets the **LineCap** object used to start and end lines within the **GraphicsPath** object that defines this CustomLineCap object. |
| [SetStrokeCaps](#SetStrokeCaps) | Sets the **LineCap** objects used to start and end lines within the **GraphicsPath** object that defines this **CustomLineCap** object. |
| [SetStrokeJoin](#SetStrokeJoin) | Sets the style of line join for the stroke. |
| [SetWidthScale](#SetWidthScale) | Sets the value of the scale width.  |

# CGpAdjustableArrowCap Class

The **CGpAdjustableArrowCap** object extends **CGpCustomLineCap**. This object builds a line cap that looks like an arrow.

**Inherits from**: CGpCustomLineCap.
**Imclude file**: CGpLineCaps.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorArrowCap) | Creates an adjustable arrow line cap with the specified height and width. |
| [GetHeight](#GetHeight) | Gets the height of the arrow cap. |
| [GetMiddleInset](#GetMiddleInset) | Gets the value of the inset. |
| [GetWidth](#GetWidth) | Gets the width of the arrow cap. |
| [IsFilled](#IsFilled) | Determines whether the arrow cap is filled. |
| [SetFillState](#SetFillState) | Sets the fill state of the arrow cap. |
| [SetHeight](#SetHeight) | Sets the height of the arrow cap. |
| [SetMiddleInset](#SetMiddleInset) | Sets the number of units that the midpoint of the base shifts towards the vertex. |
| [SetWidth](#SetWidth) | Sets the width of the arrow cap. |

# <a name="ConstructorCustomLineCap"></a>Constructors (CGpCustomLineCap)

Creates a **CustomLineCap** object.

```
CONSTRUCTOR CGpCustomLineCap
CONSTRUCTOR CGpCustomLineCap (BYVAL pCustomLineCap AS CGpCustomLineCap PTR)
CONSTRUCTOR CGpCustomLineCap (BYVAL pFillPath AS CGpGraphicsPath PTR, _
   BYVAL pStrokePath AS CGpGraphicsPath PTR, BYVAL baseCap AS LineCap = LinecapFLat, _
   BYVAL baseInset AS SINGLE = 0.0)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pFillPath* | Pointer to a path. |
| *pStrokePath* | Pointer to a path. |
| *baseCap* | Optional. Element of the **LineCap** enumeration that specifies the line cap that will be used. The default value is **LineCapFlat**.  |
| *baseInset* | Optional. The default value is 0.0. |

#### Remarks

The *pFillPath* and *pStrokePath* parameters cannot be used at the same time. You should pass NULL to one of those two parameters. If you pass nonnull values to both parameters, then *pFillPath* is ignored.

The **CustomLineCap** class uses the winding fill mode regardless of the fill mode that is set for the **GraphicsPath** object passed to the **CustomLineCap** constructor.

# <a name="Clone"></a>Clone (CGpCustomLineCap)

Copies the contents of the existing **CustomLineCap** object into a new **CustomLineCap** object.

```
FUNCTION Clone (BYVAL pCustomLineCap AS CGpCustomLineCap PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pCustomLineCap* | Pointer to the **CustomLineCap** object where to copy the contents of the existing object. |

#### Example

```
' ========================================================================================
' The following example creates a CustomLineCap object with a stroke path, creates a second
' CustomLineCap object by cloning the first, and then assigns the cloned cap to a Pen object.
' It then draws a line by using the Pen object.
' ========================================================================================
SUB Example_Clone (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Path object, and add two lines to it
   DIM pts(0 TO 2) AS GpPoint = {GDIP_POINT(-15, -15), GDIP_POINT(0, 0), GDIP_POINT(15, -15)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 2) AS GpPoint = {(-15, -15), (0, 0), (15, -15)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 2) AS GpPoint
'   pts(0).x = -15 : pts(0).y = -15 : pts(2).x = 15: pts(2).y = -15
'#endif

   DIM capPath AS CGpGraphicsPath = FillModeAlternate
   capPath.AddLines(@pts(0), 3)

   ' // Create a CustomLineCap object
   DIM firstCap AS CGpCustomLineCap = CGpCustomLineCap(NULL, @capPath)

  ' // Create a copy of firstCap
   DIM secondCap AS CGpCustomLineCap
   firstCap.Clone(@secondCap)

  ' // Create a Pen object, assign second cap as the custom end cap, and draw a line
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 0), 1)
   pen.SetCustomEndCap(@secondCap)
   graphics.DrawLine(@pen, 0, 0, 100, 100)

END SUB
' ========================================================================================
```

# <a name="GetBaseCap"></a>GetBaseCap (CGpCustomLineCap)

Gets the style of the base cap. The base cap is used as a cap at the end of a line along with this **CustomLineCap** object.

```
FUNCTION GetBaseCap () AS LineCap
```

#### Return value

This method returns the value of the line cap used at the base of the line.

#### Example

```
' ========================================================================================
' The following example creates a CustomLineCap object, sets its base cap, and then gets
' the base cap and assigns it to a Pen object. It then uses the Pen object to draw a line.
' ========================================================================================
SUB Example_GetBaseCap (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' //Create a Path object
   DIM capPath AS CGpGraphicsPath

   ' // Create a CustomLineCap object, and set its base cap to LineCapRound
   DIM custCap AS CGpCustomLineCap = CGpCustomLineCap(NULL, @capPath)
   custCap.SetBaseCap(LineCapRound)

   ' // Get the base cap of custCap
   DIM baseCap AS LineCap = custCap.GetBaseCap

   ' // Create a Pen object, assign baseCap as the end cap, and draw a line
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 255, 0), 10)
   pen.SetEndCap(baseCap)
   graphics.DrawLine(@pen, 0, 0, 100, 100)

END SUB
' ========================================================================================
```

# <a name="GetBaseInset"></a>GetBaseInset (CGpCustomLineCap)

Gets the distance between the base cap to the start of the line.

```
FUNCTION GetBaseInset () AS SINGLE
```

#### Return value

The base inset is used to separate the base cap from the start of the line. A value of 0 makes the base cap and the line touch. A value greater than 0 inserts a space (in units) between the line cap and the start of the line.

#### Example

```
' ========================================================================================
' The following example creates a CustomLineCap object, gets the base inset of the cap,
' and then creates a second CustomLineCap object that uses the same base inset.
' ========================================================================================
SUB Example_GetBaseInset (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Path object, and add two lines to it
   DIM pts(0 TO 2) AS GpPoint = {GDIP_POINT(-15, -15), GDIP_POINT(0, 0), GDIP_POINT(15, -15)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 2) AS GpPoint = {(-15, -15), (0, 0), (15, -15)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 2) AS GpPoint
'   pts(0).x = -15 : pts(0).y = -15 : pts(2).x = 15: pts(2).y = -15
'#endif

   DIM capPath AS CGpGraphicsPath = FillModeAlternate
   capPath.AddLines(@pts(0), 3)

   ' // Create a CustomLineCap object, and set its base cap to LineCapRound
   DIM custCap AS CGpCustomLineCap = CGpCustomLineCap(NULL, @capPath, LineCapRound, 5)

  ' // Get the base inset of custCap
   DIM baseInset AS SINGLE = custCap.GetBaseInset

  ' // Create a second CustomLineCap object with the same base inset as the first.
   DIM insetCap AS CGpCustomLineCap = CGpCustomLineCap(NULL, @capPath, LineCapRound, baseInset)

  ' // Create a Pen object and assign insetCap as the custom end cap. Then draw a line.
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 5)
   pen.SetCustomEndCap(@insetCap)
   graphics.DrawLine(@pen, 10, 10, 200, 200)

END SUB
' ========================================================================================
```

# <a name="GetStrokeCaps"></a>GetStrokeCaps (CGpCustomLineCap)

Gets the end cap styles for both the start line cap and the end line cap. Line caps are objects that end the individual lines within a path.

```
FUNCTION GetStrokeCaps (BYVAL startCap AS LineCap PTR, BYVAL endCap AS LineCap PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *startCap* | Out. Pointer to a variable that receives an element of the LineCap enumeration that indicates the line cap used at the start of the line to be drawn. |
| *endCap* | Out. Pointer to a variable that receives an element of the LineCap enumeration that indicates the line cap used at the end of the line to be drawn. |

#### Example

```
' ========================================================================================
' The following example creates a CustomLineCap object and sets its start and end line caps.
' It then gets the line caps and assigns them as the start and end caps of a Pen object
' that it then uses to draw a line.
' ========================================================================================
SUB Example_GetStrokeCaps (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Path object, and add two lines to it
   DIM pts(0 TO 2) AS GpPoint = {GDIP_POINT(-15, -15), GDIP_POINT(0, 0), GDIP_POINT(15, -15)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 2) AS GpPoint = {(-15, -15), (0, 0), (15, -15)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 2) AS GpPoint
'   pts(0).x = -15 : pts(0).y = -15 : pts(2).x = 15: pts(2).y = -15
'#endif

   DIM capPath AS CGpGraphicsPath = FillModeAlternate
   capPath.AddLines(@pts(0), 3)

   ' // Create a CustomLineCap object, and set its base cap to LineCapRound
   DIM custCap AS CGpCustomLineCap = CGpCustomLineCap(NULL, @capPath)

  ' // Set the start and end caps for custCap.
   custCap.SetStrokeCaps(LineCapTriangle, LineCapRound)

  ' // Get the start and end caps from custCap.
   DIM AS LineCap startStrokeCap, endStrokeCap
   custCap.GetStrokeCaps(@startStrokeCap, @endStrokeCap)

  ' // Create a Pen object, assign startStrokeCap and endStrokeCap as the
  ' // start and end caps, and draw a line.
   DIM strokeCapPen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 0), 15.2)
   strokeCapPen.SetLineCap(startStrokeCap, endStrokeCap, DashCapFlat)
   graphics.DrawLine(@strokeCapPen, 100, 100, 300, 100)

END SUB
' ========================================================================================
```

# <a name="GetStrokeJoin"></a>GetStrokeJoin (CGpCustomLineCap)

Returns the style of **LineJoin** used to join multiple lines in the same **GraphicsPath** object.

```
FUNCTION GetStrokeJoin () AS LineJoin
```

#### Return value

The **CustomLineCap** object uses a path and a stroke to define the end cap. The stroke is contained in a **GraphicsPath** object, which can contain more than one figure. If there is more than one figure in the **GraphicsPath** object, the stroke join determines how their joint is graphically displayed.

#### Example

```
' ========================================================================================
' The following example creates a CustomLineCap object with a stroke join. It then gets the
' stroke join and assigns it as the line join of a Pen object that it then uses to draw a line.
' ========================================================================================
SUB Example_GetStrokeJoin (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Path object, and add two lines to it
   DIM pts(0 TO 2) AS GpPoint = {GDIP_POINT(-15, -15), GDIP_POINT(0, 0), GDIP_POINT(15, -15)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 2) AS GpPoint = {(-15, -15), (0, 0), (15, -15)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 2) AS GpPoint
'   pts(0).x = -15 : pts(0).y = -15 : pts(2).x = 15: pts(2).y = -15
'#endif

   DIM capPath AS CGpGraphicsPath = FillModeAlternate
   capPath.AddLines(@pts(0), 3)

   ' // Create a CustomLineCap object, and set its base cap to LineCapRound
   DIM custCap AS CGpCustomLineCap = CGpCustomLineCap(NULL, @capPath)

  ' // Set the start and end caps for custCap
   custCap.SetStrokeJoin(LineJoinBevel)

   ' // Get the stroke join from custCap
   DIM strokeJoin AS LineJoin = custCap.GetStrokeJoin

  ' // Create a Pen object, assign strokeJoin as the line join,
  ' // and draw two joined lines in a path.
   DIM strokeJoinPen AS CGpPen = CGpPen(GDIP_ARGB(255, 255, 0, 0), 15.1)
   strokeJoinPen.SetLineJoin(strokeJoin)
   DIM joinPath AS CGpGraphicsPath
   joinPath.AddLine(10, 10, 10, 200)
   joinPath.AddLine(10, 200, 200, 200)
   graphics.DrawPath(@strokeJoinPen, @joinPath)

END SUB
' ========================================================================================
```

# <a name="GetWidthScale"></a>GetWidthScale (CGpCustomLineCap)

Gets the value of the scale width. This is the amount to scale the custom line cap relative to the width of the **Pen** object used to draw a line. The default value of 1.0 does not scale the line cap.

```
FUNCTION GetWidthScale () AS SINGLE
```

#### Return value

This method returns the value of the width-scaling factor.

#### Example

```
' ========================================================================================
' The following example creates a CustomLineCap object and sets the width scale. It then
' gets the width scale, assigns the custom cap to a Pen object, and draws a line.
' ========================================================================================
SUB Example_GetWidthScale (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Path object, and add two lines to it
   DIM pts(0 TO 2) AS GpPoint = {GDIP_POINT(-15, -15), GDIP_POINT(0, 0), GDIP_POINT(15, -15)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 2) AS GpPoint = {(-15, -15), (0, 0), (15, -15)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 2) AS GpPoint
'   pts(0).x = -15 : pts(0).y = -15 : pts(2).x = 15: pts(2).y = -15
'#endif

   DIM capPath AS CGpGraphicsPath = FillModeAlternate
   capPath.AddLines(@pts(0), 3)

   ' // Create a CustomLineCap object, and set its base cap to LineCapRound
   DIM custCap AS CGpCustomLineCap = CGpCustomLineCap(NULL, @capPath)

   ' // Set the width scale for custCap.
   custCap.SetWidthScale(3)

   ' // Get the width scale from custCap.
   DIM widthScale AS SINGLE = custCap.GetWidthScale

   ' // If the width scale is 3, assign custCap as the end cap of a Pen object and draw a line.
   IF widthScale = 3 THEN
      DIM widthScalePen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 255, 0), 1.0)
      widthScalePen.SetCustomEndCap(@custCap)
      graphics.DrawLine(@widthScalePen, 0, 0, 200, 200)
   END IF

END SUB
' ========================================================================================
```

# <a name="SetBaseCap"></a>SetBaseCap (CGpCustomLineCap)

Sets the **LineCap** that appears as part of this **CustomLineCap** at the end of a line.

```
FUNCTION SetBaseCap (BYVAL baseCap AS LineCap) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *baseCap* | Element of the **LineCap** enumeration that specifies the line cap used on the ends of the line to be drawn. |

#### Return value

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a CustomLineCap object and sets its base cap. It then
' assigns the custom cap as the end cap of a Pen object and draws a line.
' ========================================================================================
SUB Example_SetBaseCap (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Path object, and add two lines to it
   DIM pts(0 TO 2) AS GpPoint = {GDIP_POINT(-15, -15), GDIP_POINT(0, 0), GDIP_POINT(15, -15)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 2) AS GpPoint = {(-15, -15), (0, 0), (15, -15)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 2) AS GpPoint
'   pts(0).x = -15 : pts(0).y = -15 : pts(2).x = 15: pts(2).y = -15
'#endif

   DIM capPath AS CGpGraphicsPath = FillModeAlternate
   capPath.AddLines(@pts(0), 3)

   ' // Create a CustomLineCap object, and set its base cap to LineCapRound
   DIM custCap AS CGpCustomLineCap = CGpCustomLineCap(NULL, @capPath)
   custCap.SetBaseCap(LineCapRound)

   ' // Create a Pen object, assign baseCap as the end cap, and draw a line
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 5.3)
   pen.SetCustomEndCap(@custCap)
   graphics.DrawLine(@pen, 10, 10, 200, 200)

END SUB
' ========================================================================================
```

# <a name="SetBaseInset"></a>SetBaseInset (CGpCustomLineCap)

Sets the base inset value of this custom line cap. This is the distance between the end of a line and the base cap.

```
FUNCTION SetBaseInset (BYVAL baseinset AS SINGLE) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *baseinset* | Single precision number that specifies the distance, in units, from the base cap to the start of the line. If this value is greater than the length of the line, the behavior of this method is undefined. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a CustomLineCap object and sets the base inset of the cap.
' It then assigns the custom cap to a Pen object and draws a line.
' ========================================================================================
SUB Example_SetBaseInset (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Path object, and add two lines to it
   DIM pts(0 TO 2) AS GpPoint = {GDIP_POINT(-15, -15), GDIP_POINT(0, 0), GDIP_POINT(15, -15)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 2) AS GpPoint = {(-15, -15), (0, 0), (15, -15)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 2) AS GpPoint
'   pts(0).x = -15 : pts(0).y = -15 : pts(2).x = 15: pts(2).y = -15
'#endif

   DIM capPath AS CGpGraphicsPath = FillModeAlternate
   capPath.AddLines(@pts(0), 3)

   ' // Create a CustomLineCap object, and set its base cap to LineCapRound
   DIM custCap AS CGpCustomLineCap = CGpCustomLineCap(NULL, @capPath, LineCapRound)
   ' // Set the base inset
   custCap.SetBaseInset(5)

  ' // Create a Pen object, assign custCap as the custom end cap, and then draw a line.
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 0), 5.1)
   pen.SetCustomEndCap(@custCap)
   graphics.DrawLine(@pen, 10, 10, 200, 200)

END SUB
' ========================================================================================
```

# <a name="SetStrokeCap"></a>SetStrokeCap (CGpCustomLineCap)

Sets the **LineCap** object used to start and end lines within the **GraphicsPath** object that defines this **CustomLineCap** object.

```
FUNCTION SetStrokeCap (BYVAL strokeCap AS LineCap) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *strokeCap* | Element of the **LineCap** enumeration that specifies the line cap that will be used on the ends of the line to be drawn. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a CustomLineCap object and sets its start and end stroke caps.
' It then assigns the custom cap to a Pen object and draws a line.
' ========================================================================================
SUB Example_SetStrokeCap (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Path object, and add two lines to it
   DIM pts(0 TO 2) AS GpPoint = {GDIP_POINT(-15, -15), GDIP_POINT(0, 0), GDIP_POINT(15, -15)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 2) AS GpPoint = {(-15, -15), (0, 0), (15, -15)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 2) AS GpPoint
'   pts(0).x = -15 : pts(0).y = -15 : pts(2).x = 15: pts(2).y = -15
'#endif

   DIM capPath AS CGpGraphicsPath = FillModeAlternate
   capPath.AddLines(@pts(0), 3)

   ' // Create a CustomLineCap object, and set its base cap to LineCapRound
   DIM custCap AS CGpCustomLineCap = CGpCustomLineCap(NULL, @capPath)

   ' // Set the start and end caps for custCap to LineCapTriangle
   custCap.SetStrokeCap(LineCapTriangle)

   ' // Create a Pen object, assign startStrokeCap and endStrokeCap as the
   ' // start and end caps, and draw a line.
   DIM strokeCapPen AS CGpPen = CGpPen(GDIP_ARGB(255, 255, 0, 0), 4.8)
   strokeCapPen.SetCustomEndCap(@custCap)
   graphics.DrawLine(@strokeCapPen, 100, 100, 300, 100)

END SUB
' ========================================================================================
```

# <a name="SetStrokeCaps"></a>SetStrokeCaps (CGpCustomLineCap)

Sets the **LineCap** objects used to start and end lines within the **GraphicsPath** object that defines this **CustomLineCap** object.

```
FUNCTION SetStrokeCaps (BYVAL startCap AS LineCap, BYVAL endCap AS LineCap) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *startCap* | Element of the **LineCap** enumeration that specifies the line cap that will be used for the start of the line to be drawn. |
| *endCap* | Element of the **LineCap** enumeration that specifies the line cap that will be used for the end of the line to be drawn. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a CustomLineCap object and sets its start and end stroke caps.
' It then assigns the custom cap to a Pen object and draws a line.
' ========================================================================================
SUB Example_SetStrokeCaps (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Path object, and add two lines to it
   DIM pts(0 TO 2) AS GpPoint = {GDIP_POINT(-15, -15), GDIP_POINT(0, 0), GDIP_POINT(15, -15)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 2) AS GpPoint = {(-15, -15), (0, 0), (15, -15)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 2) AS GpPoint
'   pts(0).x = -15 : pts(0).y = -15 : pts(2).x = 15: pts(2).y = -15
'#endif

   DIM capPath AS CGpGraphicsPath = FillModeAlternate
   capPath.AddLines(@pts(0), 3)

   ' // Create a CustomLineCap object, and set its base cap to LineCapRound
   DIM custCap AS CGpCustomLineCap = CGpCustomLineCap(NULL, @capPath)

   ' // Set the start and end caps for custCap.
   custCap.SetStrokeCaps(LineCapTriangle, LineCapRound)

   ' // Create a Pen object, assign startStrokeCap and endStrokeCap as the
   ' // start and end caps, and draw a line.
   DIM strokeCapPen AS CGpPen = CGpPen(GDIP_ARGB(255, 255, 0, 255), 5.0!)
   strokeCapPen.SetCustomEndCap(@custCap)
   graphics.DrawLine(@strokeCapPen, 100, 100, 300, 100)

END SUB
' ========================================================================================
```

# <a name="SetStrokeJoin"></a>SetStrokeJoin (CGpCustomLineCap)

Sets the style of line join for the stroke. The line join specifies how two lines that intersect within the **GraphicsPath** object that makes up the custom line cap are joined.

```
FUNCTION SetStrokeJoin (BYVAL nLineJoin AS LineJoin) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *strokeJoin* | Element of the LineJoin enumeration that specifies the line join that will be used for two lines that are drawn by the same pen and that have overlapping ends. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a CustomLineCap object with a stroke join. It then gets the
' stroke join and assigns it as the line join of a Pen object that it then uses to draw a line.
' ========================================================================================
SUB Example_SetStrokeJoin (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Path object, and add two lines to it
   DIM pts(0 TO 2) AS GpPoint = {GDIP_POINT(-15, -15), GDIP_POINT(0, 0), GDIP_POINT(15, -15)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 2) AS GpPoint = {(-15, -15), (0, 0), (15, -15)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 2) AS GpPoint
'   pts(0).x = -15 : pts(0).y = -15 : pts(2).x = 15: pts(2).y = -15
'#endif

   DIM capPath AS CGpGraphicsPath = FillModeAlternate
   capPath.AddLines(@pts(0), 3)

   ' // Create a CustomLineCap object, and set its base cap to LineCapRound
   DIM custCap AS CGpCustomLineCap = CGpCustomLineCap(NULL, @capPath)

  ' // Set the start and end caps for custCap
   custCap.SetStrokeJoin(LineJoinBevel)

  ' // Create a Pen object, assign strokeJoin as the line join,
  ' // and draw two joined lines in a path.
   DIM strokeJoinPen AS CGpPen = CGpPen(GDIP_ARGB(255, 255, 0, 0), 15.1)
   strokeJoinPen.SetCustomEndCap(@custCap)
   graphics.DrawLine(@strokeJoinPen, 0, 0, 200, 200)

END SUB
' ========================================================================================
```

# <a name="SetWidthScale"></a>SetWidthScale (CGpCustomLineCap)

Sets the value of the scale width. This is the amount to scale the custom line cap relative to the width of the **Pen** used to draw lines. The default value of 1.0 does not scale the line cap.

```
FUNCTION SetWidthScale (BYVAL widthScale AS SINGLE) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *widthScale* | Single precision number that specifies the scaling factor that will be used to scale the line width. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a CustomLineCap object and sets the width scale. It then
' gets the width scale, assigns the custom cap to a Pen object, and draws a line.
' ========================================================================================
SUB Example_SetWidthScale (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Path object, and add two lines to it
   DIM pts(0 TO 2) AS GpPoint = {GDIP_POINT(-15, -15), GDIP_POINT(0, 0), GDIP_POINT(15, -15)}
'#ifdef __FB_64BIT__
'   DIM pts(0 TO 2) AS GpPoint = {(-15, -15), (0, 0), (15, -15)}
'#else
'   ' // With the 32-bit compiler, the above syntax can't be used because a mess in the
'   ' // FB headers for GdiPlus: GpPoint is defined as Point in 64 bit and as Point_ in 32 bit.
'   DIM pts(0 TO 2) AS GpPoint
'   pts(0).x = -15 : pts(0).y = -15 : pts(2).x = 15: pts(2).y = -15
'#endif

   DIM capPath AS CGpGraphicsPath = FillModeAlternate
   capPath.AddLines(@pts(0), 3)

   ' // Create a CustomLineCap object, and set its base cap to LineCapRound
   DIM custCap AS CGpCustomLineCap = CGpCustomLineCap(NULL, @capPath)

   ' // Set the width scale for custCap.
   custCap.SetWidthScale(3)

   ' // Assign custCap to a Pen object and draw a line.
   DIM widthScalePen AS CGpPen = CGpPen(GDIP_ARGB(255, 180, 0, 180), 1.7)
   widthScalePen.SetCustomEndCap(@custCap)
   graphics.DrawLine(@widthScalePen, 0, 0, 200, 200)

END SUB
' ========================================================================================
```

# <a name="ConstructorArrowCap"></a>Constructors (CGpAdjustableArrowCap)

Creates an adjustable arrow line cap with the specified height and width. The arrow line cap can be filled or nonfilled. The middle inset defaults to zero.

```
CONSTRUCTOR CGpAdjustableArrowCap (BYVAL pAdjustableArrowCap AS CGpAdjustableArrowCap PTR)
CONSTRUCTOR CGpAdjustableArrowCap (BYVAL nHeight AS SINGLE, BYVAL nWidth AS SINGLE, _
   BYVAL bIsFilled AS BOOL = CTRUE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nHeight* | The length, in units, of the arrow from its base to its point. |
| *nWidth* | The distance, in units, between the corners of the base of the arrow. |
| *bIsFilled* | Optional. Boolean value that specifies whether the arrow is filled. The default value is TRUE. |

#### Remarks

The middle inset is the number of units that the midpoint of the base shifts towards the vertex. A middle inset of zero results in no shift — the base is a straight line, giving the arrow a triangular shape. A positive (greater than zero) middle inset results in a shift the specified number of units toward the vertex — the base is an arrow shape that points toward the vertex, giving the arrow cap a V-shape. A negative (less than zero) middle inset results in a shift the specified number of units away from the vertex — the base becomes an arrow shape that points away from the vertex, giving the arrow either a diamond shape (if the absolute value of the middle inset is equal to the height) or distorted diamond shape. If the middle inset is equal to or greater than the height of the arrow cap, the cap does not appear at all. The value of the middle inset affects the arrow cap only if the arrow cap is filled. The middle inset defaults to zero when an **AdjustableArrowCap** object is constructed.

#### Example

```
' ========================================================================================
' The following example creates two AdjustableArrowCap objects, arrowCapStart and
' arrowCapEnd, and sets the fill mode to TRUE. The code then creates a Pen object and
' assigns arrowCapStart as the starting line cap for this Pen object and arrowCapEnd as
' the ending line cap. Next, draws a line.
' ========================================================================================
SUB Example_CreateAdjustableArrowCap (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   ' // Create an AdjustableArrowCap that is filled
   DIM arrowCapStart AS CGpAdjustableArrowCap = CGpAdjustableArrowCap(10, 10, CTRUE)
   ' // Adjust to DPI by setting the scale width
   arrowCapStart.SetWidthScale(rxRatio)

   ' // Create an AdjustableArrowCap that is not filled
   DIM arrowCapEnd AS CGpAdjustableArrowCap = CGpAdjustableArrowCap(15, 15, FALSE)
   ' // Adjust to DPI by setting the scale width
   arrowCapEnd.SetWidthScale(rxRatio)

   ' // Create a Pen
   DIM arrowPen AS CGpPen = GDIP_ARGB(255, 0, 0, 0)

   ' // Assign arrowStart as the start cap
   arrowPen.SetCustomStartCap(@arrowCapStart)

   ' // Assign arrowEnd as the end cap
   arrowPen.SetCustomEndCap(@arrowCapEnd)

   ' // Draw a line using arrowPen
   graphics.DrawLine(@arrowPen, 0, 0, 100, 100)

END SUB
' ========================================================================================
 ```

# <a name="GetHeight"></a>GetHeight (CGpAdjustableArrowCap)

Gets the height of the arrow cap. The height is the distance from the base of the arrow to its vertex.

```
FUNCTION GetHeight () AS SINGLE
```

#### Example

```
 ========================================================================================
' The following example creates an AdjustableArrowCap, myArrow, and sets the height of the
' cap. The code then creates a Pen, assigns myArrow as the ending line cap for the Pen,
' and draws a capped line. Next, the code gets the height of the arrow cap, creates a new
' arrow cap with height equal to the height of myArrow, assigns the new arrow cap as the
' ending line cap for the Pen, and draws another capped line.
' ========================================================================================
SUB Example_GetHeight (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   ' // Create an AdjustableArrowCap with a height of 10 pixels
   DIM myArrow AS CGpAdjustableArrowCap = CGpAdjustableArrowCap(10, 10, CTRUE)
   ' // Adjust to DPI by setting the scale width
   myArrow.SetWidthScale(rxRatio)

   ' // Create a Pen, and assign myArrow as the end cap
   DIM arrowPen AS CGpPen = GDIP_ARGB(255, 0, 0, 0)
   arrowPen.SetCustomEndCap(@myArrow)

   ' // Draw a line using arrowPen
   graphics.DrawLine(@arrowPen, 0, 20, 100, 20)

   ' // Create a second arrow cap using the height of the first one
   DIM AS CGpAdjustableArrowCap otherArrow = CGpAdjustableArrowCap(myArrow.GetHeight, 20, CTRUE)
   otherArrow.SetWidthScale(rxRatio)

   ' // Assign the new arrow cap as the end cap for arrowPen
   arrowPen.SetCustomEndCap(@otherArrow)

   ' // Draw a line using arrowPen
   graphics.DrawLine(@arrowPen, 0, 55, 100, 55)

END SUB
' ========================================================================================
```

# <a name="GetMiddleInset"></a>GetMiddleInset (CGpAdjustableArrowCap)

Gets the value of the inset. The middle inset is the number of units that the midpoint of the base shifts towards the vertex.

```
FUNCTION GetMiddleInset () AS SINGLE
```

#### Remarks

The middle inset is the number of units that the midpoint of the base shifts towards the vertex. A middle inset of zero results in no shift — the base is a straight line, giving the arrow a triangular shape. A positive (greater than zero) middle inset results in a shift the specified number of units toward the vertex — the base is an arrow shape that points toward the vertex, giving the arrow cap a V-shape. A negative (less than zero) middle inset results in a shift the specified number of units away from the vertex — the base becomes an arrow shape that points away from the vertex, giving the arrow either a diamond shape (if the absolute value of the middle inset is equal to the height) or distorted diamond shape. If the middle inset is equal to or greater than the height of the arrow cap, the cap does not appear at all. The value of the middle inset affects the arrow cap only if the arrow cap is filled. The middle inset defaults to zero when an **AdjustableArrowCap** object is constructed.

#### Example

```
' ========================================================================================
' The following example creates an AdjustableArrowCap object, myArrow, with middle
' inset set to zero (default value). The code then creates a Pen object, assigns
' myArrow as the ending line cap for this Pen object, and draws a capped line. Next,
' the code gets the middle inset, increments it, and draws another capped line.
' ========================================================================================
SUB Example_GetMiddleInset (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   ' // Create an AdjustableArrowCap with a height of 10 pixels
   DIM myArrow AS CGpAdjustableArrowCap = CGpAdjustableArrowCap(10, 10, CTRUE)
   ' // Adjust to DPI by setting the scale width
   myArrow.SetWidthScale(rxRatio)

   ' // Create a Pen, and assign myArrow as the end cap
   DIM arrowPen AS CGpPen = GDIP_ARGB(255, 0, 0, 0)
   arrowPen.SetCustomEndCap(@myArrow)

   ' // Draw a line using arrowPen
   graphics.DrawLine(@arrowPen, 0, 20, 100, 20)

   ' // Get the inset of the arrow
   DIM inset AS SINGLE = myArrow.GetMiddleInset

   ' // Increase inset by 5 pixels and draw another line
   myArrow.SetMiddleInset(inset + 5)
   arrowPen.SetCustomEndCap(@myArrow)
   graphics.DrawLine(@arrowPen, 0, 50, 100, 50)

END SUB
' ========================================================================================
```

# <a name="GetWidth"></a>GetWidth (CGpAdjustableArrowCap)

Gets the width of the arrow cap. The width is the distance between the endpoints of the base of the arrow.

```
FUNCTION GetWidth () AS SINGLE
```

#### Example

```
' ========================================================================================
' The following example creates an AdjustableArrowCap object, myArrow, and sets the width
' of the cap to 5 pixels. Next, the code creates a Pen object, assigns myArrow as the
' ending line cap for this Pen object, and draws a capped line. The code obtains the width
' using IGdipAdjustableArrowCap.GetWidth and sets the height equal to the width. The code then
' draws another capped line with the new cap.
' ========================================================================================
SUB Example_GetWidth (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   ' // Create an AdjustableArrowCap with a height of 10 pixels.
   DIM myArrow AS CGpAdjustableArrowCap = CGpAdjustableArrowCap(10, 5, CTRUE)
   ' // Adjust to DPI by setting the scale width
   myArrow.SetWidthScale(rxRatio)

   ' // Create a Pen, and assign myArrow as the end cap
   DIM arrowPen AS CGpPen = GDIP_ARGB(255, 0, 0, 0)
   arrowPen.SetCustomEndCap(@myArrow)

   ' // Get the width of the arrow
   DIM nWidth AS SINGLE = myArrow.GetWidth

   ' // Draw a line using arrowPen
   graphics.DrawLine(@arrowPen, 0, 0, 100, 100)

   ' // Set height equal to width and draw another line
   myArrow.SetHeight(nWidth)
   arrowPen.SetCustomEndCap(@myArrow)
   graphics.DrawLine(@arrowPen, 0, 40, 100, 140)

END SUB
' ========================================================================================
```

# <a name="IsFilled"></a>IsFilled (CGpAdjustableArrowCap)

Determines whether the arrow cap is filled.

```
FUNCTION IsFilled () AS BOOLEAN
```

#### Return value

If the arrow cap is filled, this method returns TRUE; otherwise, it returns FALSE.

#### Example

```
' ========================================================================================
' The following example creates an AdjustableArrowCap object, myArrow, and sets the fill
' mode to TRUE. The code then creates a Pen object and assigns myArrow as the ending line
' cap for this Pen object. Next, the code tests whether myArrow is filled and, if it is,
' draws a line.
' ========================================================================================
SUB Example_IsFilled (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   ' // Create an AdjustableArrowCap with a height of 10 pixels
   DIM myArrow AS CGpAdjustableArrowCap = CGpAdjustableArrowCap(10, 10, CTRUE)
   ' // Adjust to DPI by setting the scale width
   myArrow.SetWidthScale(rxRatio)

   ' // Create a Pen, and assign myArrow as the end cap
   DIM arrowPen AS CGpPen = GDIP_ARGB(255, 0, 0, 0)
   arrowPen.SetCustomEndCap(@myArrow)

   ' // If the cap is filled, draw a line using arrowPen
   IF myArrow.IsFilled THEN graphics.DrawLine(@arrowPen, 0, 0, 100, 100)

END SUB
' ========================================================================================
```

# <a name="SetFillState"></a>SetFillState (CGpAdjustableArrowCap)

Sets the fill state of the arrow cap. If the arrow cap is not filled, only the outline is drawn.

```
FUNCTION SetFillState (BYVAL bIsFilled AS BOOL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *bIsFilled* | Boolean value that specifies whether the arrow cap is filled. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates an AdjustableArrowCap object, myArrow, and sets the fill
' mode to FALSE. The code then creates a Pen object and assigns myArrow as the ending
' line cap for this Pen object. Next, the code draws a line.
' ========================================================================================
SUB Example_SetFillState (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   ' // Create an AdjustableArrowCap with a height of 10 pixels
   ' // Fill state defaults to TRUE when arrow cap is constructed
   DIM myArrow AS CGpAdjustableArrowCap = CGpAdjustableArrowCap(10, 10)
   ' // Adjust to DPI by setting the scale width
   myArrow.SetWidthScale(rxRatio)

   ' // Set fill state to FALSE
   myArrow.SetFillState(FALSE)

   ' // Create a Pen, and assign myArrow as the end cap
   DIM arrowPen AS CGpPen = GDIP_ARGB(255, 0, 0, 0)
   arrowPen.SetCustomEndCap(@myArrow)

   ' // Draw a line using arrowPen
   graphics.DrawLine(@arrowPen, 0, 0, 100, 100)

END SUB
' ========================================================================================
```

# <a name="SetHeight"></a>SetHeight (CGpAdjustableArrowCap)

Sets the height of the arrow cap. This is the distance from the base of the arrow to its vertex.

```
FUNCTION SetHeight (BYVAL nHeight AS SINGLE) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nHeight* | The height, in units, for the arrow cap. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates an AdjustableArrowCap, pMyArrowCap, and sets the height of the
' cap to 15 pixels. The code then creates a Pen and assigns pMyArrowCap as the ending line cap
' for this Pen. Next, the code draws a capped line.
' ========================================================================================
SUB Example_SetHeight (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   ' // Create an AdjustableArrowCap with a height of 10 pixels
   DIM myArrow AS CGpAdjustableArrowCap = CGpAdjustableArrowCap(10, 5, CTRUE)
   ' // Adjust to DPI by setting the scale width
   myArrow.SetWidthScale(rxRatio)

   ' // Create a Pen, and assign myArrow as the end cap
   DIM arrowPen AS CGpPen = GDIP_ARGB(255, 0, 0, 0)
   arrowPen.SetCustomEndCap(@myArrow)

   ' // Get the width of the arrow
   DIM nWidth AS SINGLE = myArrow.GetWidth

   ' // Draw a line using arrowPen
   graphics.DrawLine(@arrowPen, 0, 0, 100, 100)

   ' // Set height equal to width and draw another line
   myArrow.SetHeight(nWidth)
   arrowPen.SetCustomEndCap(@myArrow)
   graphics.DrawLine(@arrowPen, 0, 40, 100, 140)

END SUB
' ========================================================================================
```

# <a name="SetMiddleInset"></a>SetMiddleInset (CGpAdjustableArrowCap)

Sets the number of units that the midpoint of the base shifts towards the vertex.

```
FUNCTION SetMiddleInset (BYVAL middleInset AS SINGLE) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *middleInset* | The number of units that the midpoint of the base shifts towards the vertex. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates an AdjustableArrowCap object, myArrow, and sets the middle
' inset of the cap to 5 pixels. The code then creates a Pen object and assigns myArrow as
' the ending line cap for this Pen object. Next, the code draws a capped line.
' ========================================================================================
SUB Example_SetMiddleInset (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   ' // Create an AdjustableArrowCap with a height of 10 pixels
   DIM myArrow AS CGpAdjustableArrowCap = CGpAdjustableArrowCap(10, 10, CTRUE)
   ' // Adjust to DPI by setting the scale width
   myArrow.SetWidthScale(rxRatio)
   ' // Set the middle inset to 5
   myArrow.SetMiddleInset(5)

   ' // Create a Pen, and assign myArrow as the end cap
   DIM arrowPen AS CGpPen = GDIP_ARGB(255, 0, 0, 0)
   arrowPen.SetCustomEndCap(@myArrow)

   ' // Draw a line using arrowPen
   graphics.DrawLine(@arrowPen, 0, 0, 100, 100)

END SUB
' ========================================================================================
```

# <a name="SetWidth"></a>SetWidth (CGpAdjustableArrowCap)

Sets the width of the arrow cap. The width is the distance between the endpoints of the base of the arrow.

```
FUNCTION SetWidth (BYVAL nWidth AS SINGLE) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nWidth* | The width, in units, for the arrow cap. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates an AdjustableArrowCap, myArrow, and sets the width of
' the cap to 10 pixels. The code then creates a Pen, assigns myArrow as the ending
' line cap for this Pen, and draws a capped line. Next, the code sets the width to 15
' pixels, reassigns the arrow cap to the pen, and draws a new capped line.
' ========================================================================================
SUB Example_SetWidth (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   ' // Create an AdjustableArrowCap with a height of 10 pixels
   DIM myArrow AS CGpAdjustableArrowCap = CGpAdjustableArrowCap(10, 10, CTRUE)
   ' // Adjust to DPI by setting the scale width
   myArrow.SetWidthScale(rxRatio)

   ' // Create a Pen, and assign myArrow as the end cap
   DIM arrowPen AS CGpPen = GDIP_ARGB(255, 0, 0, 0)
   arrowPen.SetCustomEndCap(@myArrow)

   ' // Draw a line using arrowPen
   graphics.DrawLine(@arrowPen, 0, 0, 100, 100)

   ' // Set the cap to the new width, and reassign the arrow cap to the pen object.
   myArrow.SetWidth(15.7)
   arrowPen.SetCustomEndCap(@myArrow)

   ' // Draw a line with new cap
   graphics.DrawLine(@arrowPen, 0, 40, 100, 140)

END SUB
' ========================================================================================
```
