# CGpBrush Class

The **CGpBrush** class is a base interface that defines a **Brush** object. A **Brush** object is used to paint the interior of graphics shapes, such as rectangles, ellipses, pies, polygons, and paths. You must not instantiate the **CGpBrush** class, but to use one of its derived classes.

A closed shape, such as a rectangle or an ellipse, consists of an outline and an interior. The outline is drawn with a pen and the interior is filled with a brush. GDI+ provides several brush classes for filling the interiors of closed shapes: **CGpSolidBrush**, **CGpHatchBrush**, **CGpTextureBrush**, **CGpLinearGradientBrush**, and **CGpPathGradientBrush**. All of these classes inherit from the **CGpBrush** class.

**Inherits from**: CGpBase.<br>
**Include file**: CGpBrush.inc.

### Methods

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorsBrush) | Creates a **Brush** object from another **Brush** object. |
| [Clone](#CloneBrush) | Copies the contents of the existing **Brush** object into a new **Brush** object. |
| [GetType](#GetTypeBrush) | Gets the type of this brush. |

# CGpSolidBrush Class

The **SolidBrush** object defines a solid color Brush object. A **Brush** object is used to fill in shapes similar to the way a paint brush can paint the inside of a shape.

**Inherits from**: CGpBrush.<br>
**Include file**: CGpBrush.inc.

### Methods

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorSolidBrush) | Creates a **SolidBrush** object based on a color. |
| [GetColor](#GetColorSolidBrush) | Gets the color of this brush. |
| [SetColor](#SetColorSolidBrush) | Sets the color of this brush. |

# CGpHatchBrush Class

Creates a **HatchBrush** object based on a hatch style, a foreground color, and a background color.

**Inherits from**: CGpBrush.<br>
**Include file**: CGpBrush.inc.

### Methods

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorHatchBrush) | Creates a **HatchBrush** object based on a hatch style, a foreground color, and a background color. |
| [GetBackgroundColor](#GetBackgroundColor) | Gets the background color of this hatch brush. |
| [GetForegroundColor](#GetForegroundColor) | Gets the foreground color of this hatch brush. |
| [GetHatchStyle](#GetHatchStyle) | Gets the hatch style of this hatch brush. |

# CGpLinearGradientBrush Class

Defines a brush that paints a color gradient in which the color changes evenly from the starting boundary line of the linear gradient brush to the ending boundary line of the linear gradient brush. The boundary lines of a linear gradient brush are two parallel straight lines. The color gradient is perpendicular to the boundary lines of the linear gradient brush, changing gradually across the stroke from the starting boundary line to the ending boundary line. The color gradient has one color at the starting boundary line and another color at the ending boundary line.

**Inherits from**: CGpBrush.<br>
**Include file**: CGpBrush.inc.

### Constructors and Methods

| Name       | Description |
| ---------- | ----------- |
| [Constructorss](#ConstructorsLGBrush) | Creates a **LinearGradientBrush** object. |
| [GetBlend](#GetBlendLGBrush) | Gets the blend factors and their corresponding blend positions. |
| [GetBlendCount](#GetBlendCountLGBrush) | Gets the number of blend factors currently set. |
| [GetGammaCorrection](#GetGammaCorrectionLGBrush) | Determines whether gamma correction is enabled for this brush. |
| [GetInterpolationColorCount](#GetInterpolationColorCountLGBrush) | Gets the number of colors currently set to be interpolated. |
| [GetInterpolationColors](#GetInterpolationColorsLGBrush) | Gets the blend factors and their corresponding blend positions. |
| [GetLinearColors](#GetLinearColors) | Gets the starting color and ending color. |
| [GetRectangle](#GetRectangleLGBrush) | Gets the rectangle that defines the boundaries of the gradient. |
| [GetTransform](#GetTransformLGBrush) | Gets the transformation matrix. |
| [GetWrapMode](#GetWrapModeLGBrush) | Gets the wrap mode currently set for this brush. |
| [MultiplyTransform](#MultiplyTransformLGBrush) | Updates this brush's transformation matrix with the product of itself and another matrix. |
| [ResetTransform](#ResetTransformLGBrush) | Resets the transformation matrix to the identity matrix. |
| [RotateTransform](#RotateTransformLGBrush) | Updates this brush's current transformation matrix with the product of itself and a rotation matrix. |
| [ScaleTransform](#ScaleTransformLGBrush) | Updates this brush's current transformation matrix with the product of itself and a scaling matrix. |
| [SetBlend](#SetBlendLGBrush) | Sets the blend factors and the blend positions to create a custom blend. |
| [SetBlendBellShape](#SetBlendBellShapeLGBrush) | Sets the blend bell shape. |
| [SetBlendTriangularShape](#SetBlendTriangularShapeLGBrush) | Sets the blend triangular shape. |
| [SetGammaCorrection](#SetGammaCorrectionLGBrush) | Specifies whether gamma correction is enabled. |
| [SetInterpolationColors](#SetInterpolationColorsLGBrush) | Sets the colors to be interpolated and their corresponding blend positions. |
| [SetLinearColors](#SetLinearColors) | Sets the starting color and ending color. |
| [SetTransform](#SetTransformLGBrush) | Sets the transformation matrix. |
| [SetWrapMode](#SetWrapModeLGBrush) | Sets the wrap mode. |
| [TranslateTransform](#TranslateTransformLGBrush) | Updates this brush's current transformation matrix with the product of itself and a translation matrix. |

# CGpPathGradientBrush Class

A **PathGradientBrush** object stores the attributes of a color gradient that you can use to fill the interior of a path with a gradually changing color. A path gradient brush has a boundary path, a boundary color, a center point, and a center color. When you paint an area with a path gradient brush, the color changes gradually from the boundary color to the center color as you move from the boundary path to the center point.

**Inherits from**: CGpBrush.<br>
**Include file**: CGpBrush.inc.

### Constructors and Methods

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorsPGBrush) | Creates a **PathGradientBrush** object. |
| [GetBlend](#GetBlendPGBrush) | Gets the blend factors and their corresponding blend positions. |
| [GetBlendCount](#GetBlendCountPGBrush) | Gets the number of blend factors currently set. |
| [GetCenterColor](#GetCenterColor) | Gets center color of the brush. |
| [GetCenterPoint](#GetCenterPoint) | Gets the center point of the brush. |
| [GetFocusScales](#GetFocusScales) | Gets the focus scales of the brush. |
| [GetGammaCorrection](#GetGammaCorrectionPGBrush) | Determines whether gamma correction is enabled for this brush. |
| [GetInterpolationColorCount](#GetInterpolationColorCountPGBrush) | Gets the number of preset colors currently specified for this brush. |
| [GetInterpolationColors](#GetInterpolationColorsPGBrush) | Gets preset colors and blend positions currently specified for this brush. |
| [GetPointCount](#GetPointCount) | Gets the number of points in the array of points that defines this brush's boundary path. |
| [GetRectangle](#GetRectanglePGBrush) | Gets the smallest rectangle that encloses the boundary path of this brush. |
| [GetSurroundColorCount](#GetSurroundColorCount) | Gets the number of colors that have been specified for the boundary path of this brush. |
| [GetSurroundColors](#GetSurroundColors) | Gets the surround colors currently specified for this brush. |
| [GetTransform](#GetTransformPGBrush) | Gets the transformation matrix. |
| [GetWrapMode](#GetWrapModePGBrush) | Gets the wrap mode currently set for this brush. |
| [MultiplyTransform](#MultiplyTransformPGBrush) | Updates this brush's transformation matrix with the product of itself and another matrix. |
| [ResetTransform](#ResetTransformPGBrush) | Resets the transformation matrix to the identity matrix. |
| [RotateTransform](#RotateTransformPGBrush) | Updates this brush's current transformation matrix with the product of itself and a rotation matrix. |
| [ScaleTransform](#ScaleTransformPGBrush) | Updates this brush's current transformation matrix with the product of itself and a scaling matrix. |
| [SetBlend](#SetBlendPGBrush) | Sets the blend factors and the blend positions to create a custom blend. |
| [SetBlendBellShape](#SetBlendBellShapePGBrush) | Sets the blend bell shape. |
| [SetBlendTriangularShape](#SetBlendTriangularShapePGBrush) | Sets the blend triangular shape. |
| [SetCenterColor](#SetCenterColor) | Sets the center color of this brush. |
| [SetCenterPoint](#SetCenterPoint) | Sets the center point of this brush. |
| [SetFocusScales](#SetFocusScales) | Sets the focus scales of this brush. |
| [SetGammaCorrection](#SetGammaCorrectionPGBrush) | Specifies whether gamma correction is enabled. |
| [SetInterpolationColors](#SetInterpolationColorsPGBrush) | Sets the colors to be interpolated and their corresponding blend positions. |
| [SetSurroundColors](#SetSurroundColors) | Sets the surround colors of this brush. |
| [SetTransform](#SetTransformPGBrush) | Sets the transformation matrix. |
| [SetWrapMode](#SetWrapModePGBrush) | Sets the wrap mode. |
| [TranslateTransform](#TranslateTransformPGBrush) | Updates this brush's current transformation matrix with the product of itself and a translation matrix. |

# CGpTextureBrush Class

Defines a **Brush** object that contains an **Image** object that is used for the fill. The fill image can be transformed by using the local **Matrix** object contained in the **Brush** object.

**Inherits from**: CGpBrush.<br>
**Include file**: CGpBrush.inc.

### Constructors and Methods

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorsTBrush) | Creates a texture brush. |
| [GetImage](#GetImage) | Gets a pointer to the **Image** object that is defined by this brush. |
| [GetTransform](#GetTransformTBrush) | Gets the transformation matrix. |
| [GetWrapMode](#GetWrapModeTBrush) | Gets the wrap mode currently set for this brush. |
| [MultiplyTransform](#MultiplyTransformTBrush) | Updates this brush's transformation matrix with the product of itself and another matrix. |
| [ResetTransform](#ResetTransformTBrush) | Resets the transformation matrix to the identity matrix. |
| [RotateTransform](#RotateTransformTBrush) | Updates this brush's current transformation matrix with the product of itself and a rotation matrix. |
| [ScaleTransform](#ScaleTransformTBrush) | Updates this brush's current transformation matrix with the product of itself and a scaling matrix. |
| [SetTransform](#SetTransformTBrush) | Sets the transformation matrix. |
| [SetWrapMode](#SetWrapModeTBrush) | Sets the wrap mode. |
| [TranslateTransform](#TranslateTransformTBrush) | Updates this brush's current transformation matrix with the product of itself and a translation matrix. |

# <a name="ConstructorsBrush"></a>Constructors (CGpBrush)

Creates a **Brush** object.

```
CONSTRUCTOR CGpBrush
CONSTRUCTOR CGpBrush (BYVAL pBrush AS CGpBrush PTR)
```

# <a name="CloneBrush"></a>Clone (CGpBrush)

Copies the contents of the existing **Brush** object into a new **Brush** object.

```
FUNCTION Clone (BYVAL pBrush AS CGpBrush PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pBrush* | Pointer to a variable that will receive a pointer to the cloned **Brush** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a SolidBrush object, clones it, and then uses the clone
' to fill a rectangle.
' ========================================================================================
SUB Example_CloneBrush (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   ' // Create a SolidBrush object
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)

   ' // Create a clone of solidBrush
   DIM cloneBrush AS CGpSolidBrush
   solidBrush.Clone(@cloneBrush)
   ' // You can also use:
   ' DIM cloneBrush AS CGpSolidBrush = @solidBrush

   ' // Use cloneBrush to fill a rectangle
   graphics.FillRectangle(@cloneBrush, 0, 0, 100, 100)

END SUB
' ========================================================================================
```

# <a name="GetTypeBrush"></a>GetType (CGpBrush)

Gets the type of this brush.

```
FUNCTION GetType () AS GpBrushType
```

#### Return value

This method returns the type of this brush. The value returned is one of the elements of the **BrushType** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a SolidBrush object, checks the type of the object, and
' then, if the type is BrushTypeSolidColor, uses the brush to fill a rectangle.
' ========================================================================================
SUB Example_GetType (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a SolidBrush object
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 255)

   ' // Get the type of solidBrush
   DIM nType AS BrushType = solidBrush.GetType

   ' // If the type of solidBrush is BrushTypeSolidColor, use it to fill a rectangle
   IF nType = BrushTypeSolidColor THEN
      graphics.FillRectangle(@solidBrush, 0, 0, 100, 100)
   END IF

END SUB
' ========================================================================================
```

# <a name="ConstructorSolidBrush"></a>Constructors (CGpSolidBrush)

Creates a **SolidBrush** object based on a color.

```
CONSTRUCTOR CGpSolidBrush (BYVAL pSolidBrush AS CGpSolidBrush PTR)
CONSTRUCTOR CGpSolidBrush (BYVAL colour AS ARGB = &hFF000000)
```

| Parameter  | Description |
| ---------- | ----------- |
| *colour* | An ARGB color that specifies the initial color of this solid brush. |

# <a name="GetColorSolidBrush"></a>GetColor (CGpSolidBrush)

Gets the color of this solid brush.

```
FUNCTION GetColor (BYVAL colour AS ARGB PTR) AS GpStatus
FUNCTION GetColor () AS ARGB
```

| Parameter  | Description |
| ---------- | ----------- |
| *colour* | Pointer to a variable that receives the color of this solid brush. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

The second overloaded function returns the ARGB color as the result of the function.

# <a name="SetColorSolidBrush"></a>SetColor (CGpSolidBrush)

Sets the color of this solid brush.

```
FUNCTION SetColor (BYVAL colour AS ARGB) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *colour* | The color to be set in this solid brush. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a solid brush and uses it to fill a rectangle. The code
' uses GdipSetSolidFillColor to change the color of the solid brush and then paints a
' second rectangle the new color.
' ========================================================================================
SUB Example_SetColor (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   ' // Create a solid brush, and use it to fill a rectangle
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 255)
   graphics.FillRectangle(@solidBrush, 10, 10, 200, 100)

   ' // Change the color of the brush to red, and fill another rectangle
   solidBrush.SetColor(GDIP_ARGB(255, 255, 0, 0))
   graphics.FillRectangle(@solidBrush, 220, 10, 200, 100)

END SUB
' ========================================================================================
```

# <a name="ConstructorHatchBrush"></a>Constructors (CGpHatchBrush)

Creates a **HatchBrush** object based on a hatch style, a foreground color, and a background color.

```
CONSTRUCTOR CGpHatchBrush (BYVAL pHatchBrush AS CGpHatchBrush PTR)
FUNCTION HatchBrush (BYVAL hatchStyle AS HatchStyle, BYVAL foreColor AS ARGB, _
   BYVAL backColor AS ARGB = &HFF000000)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hatchStyle* | Element of the **HatchStyle** enumeration that specifies the pattern of hatch lines that will be used. |
| *foreColor* | Reference to a color to use for the hatch lines. |
| *backColor* | Optional. Reference to a color to use for the background. |

# <a name="GetBackgroundColor"></a>GetBackgroundColor (CGpHatchBrush)

Gets the background color of this hatch brush.

```
FUNCTION GetBackgroundColor (BYVAL colour AS ARGB PTR) AS GpStatus
FUNCTION GetBackgroundColor () AS ARGB
```

| Parameter  | Description |
| ---------- | ----------- |
| *colour* | Pointer to a variable that receives the background color. The background color defines the color over which the hatch lines are drawn. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

The second overloaded function returns the ARGB color as the result of the function.

#### Example

```
' ========================================================================================
' The following example sets up three Color objects: black, turquoise, and current
' (initialized to black). A rectangle is painted by using turquoise as the background
' color and black as the foreground color. Then the HatchBrush.GetBackgroundColor method
' is used to get the current color of the brush (which at the time is turquoise). The
' address of the current Color object (initialized to black) is passed as the return point
' for the call to HatchBrush,GetBackgroundColor. When the rectangle is painted again,
' note that the background color is again turquoise (not black). This shows that the call
' to HatchBrush.GetBackgroundColor was successful.
' ========================================================================================
SUB Example_HatchBrushGetBackgroundColor (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   ' // Set colors
   DIM black AS ARGB = GDIP_ARGB(255, 0, 0, 0)           ' // foreground
   DIM turquoise AS ARGB = GDIP_ARGB(255, 0, 255, 255)   ' // background
   DIM current AS ARGB = GDIP_ARGB(255, 0, 0, 0)         ' // new foreground

   ' // Set and then draw the first hatch style.
   DIM brush AS CGpHatchBrush = CGpHatchBrush(HatchStyleHorizontal, black, turquoise)
   graphics.FillRectangle(@brush, 20, 20, 100, 50)

   ' // Get the current background color of the brush.
   brush.GetBackgroundColor(@current)

   ' // Draw the rectangle again using the current color.
   DIM brush2 AS CGpHatchBrush = CGpHatchBrush(HatchStyleDiagonalCross, black, current)
   graphics.FillRectangle(@brush2, 130, 20, 100, 50)

END SUB
' ========================================================================================
```

# <a name="GetForegroundColor"></a>GetForegroundColor (CGpHatchBrush)

Gets the foreground color of this hatch brush.

```
FUNCTION GetForegroundColor (BYVAL colour AS ARGB PTR) AS GpStatus
FUNCTION GetForegroundColor () AS ARGB
```

| Parameter  | Description |
| ---------- | ----------- |
| *colour* | Pointer to a variable that receives the foreground color. The foreground color defines the color of the hatch lines. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

The second overloaded function returns the ARGB color as the result of the function.

#### Example

```
' ========================================================================================
' The following example sets up three Color objects: black, turquoise, and current
' (initialized to black). A rectangle is painted by using turquoise as the background
' color and black as the foreground color. Then the HatchBrush.GetBackgroundColor method
' is used to get the current color of the brush (which at the time is turquoise). The
' address of the current Color object (initialized to black) is passed as the return point
' for the call to HatchBrush,GetBackgroundColor. When the rectangle is painted again,
' note that the background color is again turquoise (not black). This shows that the call
' to HatchBrush.GetBackgroundColor was successful.
' ========================================================================================
SUB Example_HatchBrushGetForegroundColor (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   ' // Set colors
   DIM blue AS ARGB = GDIP_ARGB(255, 0, 0, 255)          ' // foreground
   DIM turquoise AS ARGB = GDIP_ARGB(255, 0, 255, 255)   ' // background
   DIM current AS ARGB = GDIP_ARGB(255, 0, 0, 0)         ' // new foreground

   ' // Set and then draw the first hatch style.
   DIM brush AS CGpHatchBrush = CGpHatchBrush(HatchStyleHorizontal, blue, turquoise)
   graphics.FillRectangle(@brush, 20, 20, 100, 50)

   ' // Get the current foreground color of the brush.
   brush.GetForegroundColor(@current)

   ' // Draw the rectangle again using the current color.
   DIM brush2 AS CGpHatchBrush = CGpHatchBrush(HatchStyleDiagonalCross, current, turquoise)
   graphics.FillRectangle(@brush2, 130, 20, 100, 50)

END SUB
' ========================================================================================
```

# <a name="GetHatchStyle"></a>GetHatchStyle (CGpHatchBrush)

Gets the hatch style of this hatch brush.

```
FUNCTION GetHatchStyle () AS HatchStyle
```

#### Return value

This method returns the hatch style, which is one of the elements of the **HatchStyle** enumeration.

#### Example

```
' ========================================================================================
' The following example sets up two hatch styles: horiz and current (initialized to
' HatchStyleDiagonalCross). A rectangle that uses horiz as the hatch style is painted.
' Then the HatchBrush.GetHatchStyle method is used to get the current hatch style of the
' brush (which at the time is HatchStyleHorizontal). The address of the current HatchStyle
' object (initialized to HatchStyleDiagonalCross) is passed as the return point for the
' call to GetHatchStyle. When the rectangle is painted again, notice that the hatch style
' is again HatchStyleHorizontal (not HatchStyleDiagonalCross). This shows that the call to
' HatchBrush.GetHatchStyle was successful. 
' ========================================================================================
SUB Example_HatchBrushGetHatchStyle (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Set colors
   DIM blue AS ARGB = GDIP_ARGB(255, 0, 0, 255)          ' // foreground
   DIM turquoise AS ARGB = GDIP_ARGB(255, 0, 255, 255)   ' // background

   ' // Set hatch styles
   DIM horiz AS HatchStyle = HatchStyleHorizontal
   DIM current AS HatchStyle = HatchStyleDiagonalCross

   ' // Set and then draw the first hatch style.
   DIM brush AS CGpHatchBrush = CGpHatchBrush(horiz, blue, turquoise)
   graphics.FillRectangle(@brush, 20, 20, 100, 50)

   ' // Get the current hatch style of the brush.
   current = brush.GetHatchStyle
   
   ' // Get the current background color of the brush.
   brush.GetBackgroundColor(@current)

   ' // Draw the rectangle again using the current hatch style.
   DIM brush2 AS CGpHatchBrush = CGpHatchBrush(current, blue, turquoise)
   graphics.FillRectangle(@brush2, 130, 20, 100, 50)

END SUB
' ========================================================================================
```
# <a name="ConstructorsLGBrush"></a>Constructors (CGpLinearGradientBrush)

Creates a **LinearGradientBrush** object from a set of boundary points and boundary colors.

```
CONSTRUCTOR CGpLinearGradientBrush (BYVAL pLinearGradientBrush AS CGpLinearGradientBrush PTR)
CONSTRUCTOR CGpLinearGradientBrush (BYVAL point1 AS GpPointF PTR, BYVAL point2 AS GpPointF PTR, _
   BYVAL color1 AS ARGB, BYVAL color2 AS ARGB)
CONSTRUCTOR CGpLinearGradientBrush (BYVAL point1 AS GpPoint PTR, BYVAL point2 AS GpPoint PTR, _
   BYVAL color1 AS ARGB, BYVAL color2 AS ARGB)
```

Creates a **LinearGradientBrush** object based on a rectangle and mode of direction.

```
CONSTRUCTOR CGpLinearGradientBrush (BYVAL rc AS GpRectF PTR, BYVAL color1 AS ARGB, _
   BYVAL color2 AS ARGB, BYVAL mode AS LinearGradientMode)
CONSTRUCTOR CGpLinearGradientBrush (BYVAL rc AS GpRect PTR, BYVAL color1 AS ARGB, _
   BYVAL color2 AS ARGB, BYVAL mode AS LinearGradientMode)
```

Creates a LinearGradientBrush object from a rectangle and angle of direction.

```
CONSTRUCTOR CGpLinearGradientBrush (BYVAL rc AS GpRectF PTR, BYVAL color1 AS ARGB, _
   BYVAL color2 AS ARGB, BYVAL angle AS SINGLE, BYVAL isAngleScalable AS BOOL)
CONSTRUCTOR CGpLinearGradientBrush (BYVAL rc AS GpRect PTR, BYVAL color1 AS ARGB, _
   BYVAL color2 AS ARGB, BYVAL angle AS SINGLE, BYVAL isAngleScalable AS BOOL)
```

| Parameter  | Description |
| ---------- | ----------- |
| *point1* | Reference to a **GpPoint** structure that specifies the starting point of the gradient. The starting boundary line passes through the starting point. |
| *point2* | Reference to a **GpPoint** structure that specifies the ending point of the gradient. The ending boundary line passes through the ending point. |
| *color1* | An ARGB value that specifies the color at the starting boundary line of this linear gradient brush. |
| *color2* | An ARGB value that specifies the color at the ending boundary line of this linear gradient brush. |
| *rc* | Reference to a rectangle that specifies the starting and ending points of the gradient. The upper-left corner of the rectangle is the starting point. The lower-right corner is the ending point. |
| *mode* | Element of the **LinearGradientMode** enumeration that specifies the direction of the gradient. |
| *angle* | Real number that, if *isAngleScalable* is TRUE, specifies the base angle from which the angle of the directional line is calculated, or that, if *isAngleScalable* is FALSE, specifies the angle of the directional line. The angle is measured from the top of the rectangle that is specified by rect and must be in degrees. The gradient follows the directional line. |
| *isAngleScalable* | BOOL value that specifies whether the angle is scalable. If *isAngleScalable* is TRUE, the angle of the directional line is scalable; otherwise, the angle is not scalable. |

# <a name="GetBlendLGBrush"></a>GetBlend (CGpLinearGradientBrush)

Gets the blend factors and their corresponding blend positions from a **LinearGradientBrush** object.

```
FUNCTION GetBlend (BYVAL blendFactors AS SINGLE PTR, BYVAL blendPositions AS SINGLE PTR, _
   BYVAL count AS LONG) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *blendFactors* | Pointer to an array that receives the blend factors. Each number in the array indicates a percentage of the ending color and is in the range from 0.0 through 1.0. |
| *blendPositions* | Pointer to an array that receives the blend positions. Each number in the array indicates a percentage of the distance between the starting boundary and the ending boundary and is in the range from 0.0 through 1.0, where 0.0 indicates the starting boundary of the gradient and 1.0 indicates the ending boundary. A blend position between 0.0 and 1.0 indicates a line, parallel to the boundary lines, that is a certain fraction of the distance from the starting boundary to the ending boundary. For example, a blend position of 0.7 indicates the line that is 70 percent of the distance from the starting boundary to the ending boundary. The color is constant on lines that are parallel to the boundary lines. |
| *count* | Integer that specifies the number of blend factors to retrieve. Before calling the **GetBlend** method of a **LinearGradientBrush** object, call the **GetBlendCount** method of that same **LinearGradientBrush** object to determine the current number of blend factors. The number of blend positions retrieved is the same as the number of blend factors retrieved. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A **LinearGradientBrush** object has two parallel boundaries: a starting boundary and an ending boundary. A color is associated with each of these two boundaries. Each boundary is a straight line that passes through a specified point — the starting boundary passes through the starting point; the ending boundary passes through the ending point — and is perpendicular to the direction of the linear gradient brush. The direction of the linear gradient brush follows the line that is defined by the starting and ending points. This line, the "directional line," may be horizontal, vertical, or diagonal. All points that lie on a line that is parallel to the boundaries are the same color. When you fill an area with a linear gradient brush, the color changes gradually from one line to the next as you move along the directional line from the starting boundary to the ending boundary. By default, the change in color is proportional to the change in distance; that is, a line 30 percent of the distance between the starting boundary and the ending boundary has a color that is 30 percent of the distance between the starting boundary color and the ending boundary color. The color pattern is repeated outside of the starting and ending boundaries.

You can call the **SetBlend** method of a LinearGradientBrush object to customize the relationship between color and distance. For example, suppose you set the blend positions to {0, 0.5, 1} and you set the blend factors to {0, 0.3, 1}. Then a line 50 percent of the distance between the starting boundary and the ending boundary will have a color that is 30 percent of the distance between the starting boundary color and the ending boundary color.

#### Example

```
' ========================================================================================
' The following example creates a linear gradient brush, sets its blend, and uses the brush
' to fill a rectangle. The code then gets the blend. The blend factors and positions can
' then be inspected or used in some way.
' ========================================================================================
SUB Example_GetBlend (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM factors(0 TO 3) AS SINGLE = {0.0, 0.4, 0.6, 1.0}
   DIM positions(0 TO 3) AS SINGLE = {0.0, 0.2, 0.8, 1.0}
   DIM rcf AS GpRectF = GDIP_RECTF(0, 0, 100, 50)

   DIM linGrBrush AS CGpLinearGradientBrush = CGpLinearGradientBrush(@rcf, GDIP_ARGB(255, 255, 0, 0), _
       GDIP_ARGB(255, 0, 0, 255), LinearGradientModeHorizontal)

   linGrBrush.SetBlend(@factors(0), @positions(0), 4)
   graphics.FillRectangle(@linGrBrush, @rcf)

   DIM blendCount AS LONG = linGrBrush.GetBlendCount
   DIM rgFactors(blendCount - 1) AS SINGLE
   DIM rgPositions(blendCount - 1) AS SINGLE

   linGrBrush.GetBlend(@rgFactors(0), @rgPositions(0), blendCount)

   FOR j AS LONG = 0 TO blendCount - 1
'      // Inspect or use the value in rgFactors(j)
'      // Inspect or use the value in rgPositions(j)
      OutputDebugString STR(rgFactors(j)) & STR(rgPositions(j))
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetBlendCountLGBrush"></a>GetBlendCount (CGpLinearGradientBrush)

Gets the number of blend factors currently set for this **LinearGradientBrush** object.

```
FUNCTION GetBlendCount () AS INT_
```

#### Return value

This method returns the number of blend factors currently set for this **LinearGradientBrush** object. If no custom blend has been set by using **SetBlend**, or if invalid positions were passed to **SetBlend**, then **GetBlend** returns 1.

#### Example

```
' ========================================================================================
' The following example creates a linear gradient brush, sets its blend, and uses the brush
' to fill a rectangle. The code then gets the blend. The blend factors and positions can
' then be inspected or used in some way.
' ========================================================================================
SUB Example_GetBlend (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM factors(0 TO 3) AS SINGLE = {0.0, 0.4, 0.6, 1.0}
   DIM positions(0 TO 3) AS SINGLE = {0.0, 0.2, 0.8, 1.0}

   DIM pt1 AS GpPoint = GDIP_POINT(0, 0)
   DIM pt2 AS GpPoint = GDIP_POINT(100, 0)
   DIM linGrBrush AS CGpLinearGradientBrush = CGpLinearGradientBrush(@pt1, @pt2, _
      GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255))
   linGrBrush.SetBlend(@factors(0), @positions(0), 4)

   ' // Use the linear gradient brush to fill a rectangle.
   graphics.FillRectangle(@linGrBrush, 0, 0, 100, 50)

   ' // Obtain information about the linear gradient brush.
   DIM blendCount AS LONG = linGrBrush.GetBlendCount
   DIM rgFactors(blendCount - 1) AS SINGLE
   DIM rgPositions(blendCount - 1) AS SINGLE

   linGrBrush.GetBlend(@rgFactors(0), @rgPositions(0), blendCount)

   FOR j AS LONG = 0 TO blendCount - 1
'      // Inspect or use the value in rgFactors(j)
'      // Inspect or use the value in rgPositions(j)
      OutputDebugString STR(rgFactors(j)) & STR(rgPositions(j))
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetGammaCorrectionLGBrush"></a>GetGammaCorrection (CGpLinearGradientBrush)

Gets the focus scales of this path gradient brush.

```
FUNCTION GetGammaCorrection () AS BOOL
```

#### Return value

If gamma correction is enabled, this method returns TRUE; otherwise, it returns FALSE.

# <a name="GetInterpolationColorCountLGBrush"></a>GetInterpolationColorCount (CGpLinearGradientBrush)

Gets the number of colors currently set to be interpolated for this linear gradient brush.

```
FUNCTION GetInterpolationColorCount () AS INT_
```

#### Return value

This method returns the number of colors to be interpolated for this linear gradient brush. If no colors have been set by using **SetInterpolationColors**, or if invalid positions were passed to **SetInterpolationColors**, then **GetInterpolationColorCount** returns 0.

#### Remarks

A simple linear gradient brush has two colors: a color at the starting boundary and a color at the ending boundary. When you paint with such a brush, the color changes gradually from the starting color to the ending color as you move from the starting boundary to the ending boundary. You can create a more complex gradient by using the **SetInterpolationColors** method to specify an array of colors and their corresponding blend positions to be interpolated for this linear gradient brush.

You can obtain the colors and blend positions currently set for a linear gradient brush by calling its **GetInterpolationColors** method. Before you call the **GetInterpolationColors** method, you must allocate two buffers: one to hold the array of colors and one to hold the array of blend positions. You can call the **GetInterpolationColorCount** method to determine the required size of those buffers. The size of the colors buffer is the return value of **GetInterpolationColorCount** multiplied by **sizeof(Color)**. The size of the blend positions buffer is the value of **GetInterpolationColorCount** multiplied by **sizeof(REAL)**.

#### Example

```
' ========================================================================================
' The following example sets the colors that are interpolated for this linear gradient
' brush to red, blue, and green and sets the blend positions to 0, 0.3, and 1. The code
' calls the LinearGradientBrush::GetInterpolationColorCount method of a LinearGradientBrush
' object to obtain the number of colors currently set to be interpolated for the brush.
' Next, the code gets the colors and their positions. Then, the code fills a small
' rectangle with each color.
' ========================================================================================
SUB Example_GetInterpolationColors (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a linear gradient brush, and set the colors to be interpolated.
   DIM colors(0 TO 2) AS ARGB = {GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255), GDIP_ARGB(255, 0, 255, 0)}
   DIM positions(0 TO 2) AS SINGLE = {0.0, 0.3, 1.0}

   DIM pt1 AS GpPoint = GDIP_POINT(0, 0)
   DIM pt2 AS GpPoint = GDIP_POINT(100, 0)

   DIM linGrBrush AS CGpLinearGradientBrush = CGpLinearGradientBrush(@pt1, @pt2, GDIP_ARGB(255, 0, 0, 0), GDIP_ARGB(255, 255, 255, 255))
   linGrBrush.SetInterpolationColors(@colors(0), @positions(0), 3)

   ' // Obtain information about the linear gradient brush.
   ' // How many colors have been specified to be interpolated for this brush?
   DIM colorCount AS LONG = linGrBrush.GetInterpolationColorCount

   ' // Allocate a buffer large enough to hold the set of colors.
   DIM rgcolors(0 TO colorCount - 1) AS ARGB

   ' // Allocate a buffer to hold the relative positions of the colors.
   DIM rgPositions(0 TO colorCount - 1) AS SINGLE

   ' // Get the colors and their relative positions.
   linGrBrush.GetInterpolationColors(@rgcolors(0), @rgPositions(0), colorCount)

   ' // Fill a small rectangle with each of the colors.
   DIM pSolidBrush AS CGpSolidBrush PTR
   FOR j AS LONG = 0 TO colorCount - 1
      pSolidBrush = NEW CGpSolidBrush(rgcolors(j))
      graphics.FillRectangle(pSolidBrush, 15 * j, 0, 10, 10)
      Delete pSolidBrush
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetInterpolationColorsLGBrush"></a>GetInterpolationColors (CGpLinearGradientBrush)

Gets the blend factors and their corresponding blend positions from a **LinearGradientBrush** object.

```
FUNCTION GetInterpolationColors (BYVAL presetColors AS ARGB PTR, _
   BYVAL blendPositions AS SINGLE PTR, BYVAL count AS LONG) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *presetColors* | Pointer to an array that receives the colors. A color of a given index in the presetColors array corresponds to the blend position of that same index in the *blendPositions* array. |
| *blendPositions* | Pointer to an array that receives the blend positions. Each number in the array indicates a percentage of the distance between the starting boundary and the ending boundary and is in the range from 0.0 through 1.0, where 0.0 indicates the starting boundary of the gradient and 1.0 indicates the ending boundary. A blend position between 0.0 and 1.0 indicates a line, parallel to the boundary lines, that is a certain fraction of the distance from the starting boundary to the ending boundary. For example, a blend position of 0.7 indicates the line that is 70 percent of the distance from the starting boundary to the ending boundary. The color is constant on lines that are parallel to the boundary lines. |
| *count* | Integer that specifies the number of elements in the presetColors array. This is the same as the number of elements in the blendPositions array. Before calling the **GetInterpolationColors** method of a **LinearGradientBrush** object, call the **GetInterpolationColorCount** method of that same **LinearGradientBrush** object to determine the current number of colors. The number of blend positions retrieved is the same as the number of colors retrieved. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example sets the colors that are interpolated for this linear gradient
' brush to red, blue, and green and sets the blend positions to 0, 0.3, and 1. The code
' calls the LinearGradientBrush::GetInterpolationColorCount method of a LinearGradientBrush
' object to obtain the number of colors currently set to be interpolated for the brush.
' Next, the code gets the colors and their positions. Then, the code fills a small
' rectangle with each color.
' ========================================================================================
SUB Example_GetInterpolationColors (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a linear gradient brush, and set the colors to be interpolated.
   DIM colors(0 TO 2) AS ARGB = {GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255), GDIP_ARGB(255, 0, 255, 0)}
   DIM positions(0 TO 2) AS SINGLE = {0.0, 0.3, 1.0}

   DIM pt1 AS GpPoint = GDIP_POINT(0, 0)
   DIM pt2 AS GpPoint = GDIP_POINT(100, 0)

   DIM linGrBrush AS CGpLinearGradientBrush = CGpLinearGradientBrush(@pt1, @pt2, GDIP_ARGB(255, 0, 0, 0), GDIP_ARGB(255, 255, 255, 255))
   linGrBrush.SetInterpolationColors(@colors(0), @positions(0), 3)

   ' // Obtain information about the linear gradient brush.
   ' // How many colors have been specified to be interpolated for this brush?
   DIM colorCount AS LONG = linGrBrush.GetInterpolationColorCount

   ' // Allocate a buffer large enough to hold the set of colors.
   DIM rgcolors(0 TO colorCount - 1) AS ARGB

   ' // Allocate a buffer to hold the relative positions of the colors.
   DIM rgPositions(0 TO colorCount - 1) AS SINGLE

   ' // Get the colors and their relative positions.
   linGrBrush.GetInterpolationColors(@rgcolors(0), @rgPositions(0), colorCount)

   ' // Fill a small rectangle with each of the colors.
   DIM pSolidBrush AS CGpSolidBrush PTR
   FOR j AS LONG = 0 TO colorCount - 1
      pSolidBrush = NEW CGpSolidBrush(rgcolors(j))
      graphics.FillRectangle(pSolidBrush, 15 * j, 0, 10, 10)
      Delete pSolidBrush
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetLinearColors"></a>GetLinearColors (CGpLinearGradientBrush)

Gets the starting color and ending color of this linear gradient brush.

```
FUNCTION GetLinearColors (BYVAL colors AS ARGB PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *colors* | Pointer to an array that receives the starting color and the ending color. The first color in the colors array is the color at the starting boundary line of the gradient; the second color in the colors array is the color at the ending boundary line. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a linear gradient brush and gets the boundary colors. Next,
' the code uses each of the two colors to create a solid brush. Then, the code fills a
' rectangle with each solid brush.
' ========================================================================================
SUB Example_GetLinearColors (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a linear gradient brush.
   DIM rc AS GpRect = GDIP_RECT(0, 0, 100, 50)
   DIM linGrBrush AS CGpLinearGradientBrush = CGpLinearGradientBrush(@rc, _
      GDIP_ARGB(255, 0, 0, 0), GDIP_ARGB(255, 0, 0, 255), LinearGradientModeHorizontal)

   ' // Obtain information about the linear gradient brush.
   DIM colors(0 TO 1) AS ARGB
   linGrBrush.GetLinearColors(@colors(0))

   ' // Fill a small rectangle with each of the two colors.
   DIM solidBrush0 AS CGpSolidBrush = colors(0)
   DIM solidBrush1 AS CGpSolidBrush = colors(1)
   graphics.FillRectangle(@solidBrush0, 0, 0, 20, 20)
   graphics.FillRectangle(@solidBrush1, 25, 0, 20, 20)

END SUB
' ========================================================================================
```

# <a name="GetRectangleLGBrush"></a>GetRectangle (CGpLinearGradientBrush)

Gets the rectangle that defines the boundaries of the gradient.

```
FUNCTION GetRectangle (BYVAL rc AS GpRectF PTR) AS GpStatus
FUNCTION GetRectangle (BYVAL rc AS GpRect PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *rc* | Pointer to a **GpRectF** or **GpRect** structure that receives the rectangle that defines the boundaries of the gradient. For example, if a linear gradient brush is constructed with a starting point at (20.2, 50.8) and an ending point at (60.5, 110.0), then the defining rectangle has its upper-left point at (20.2, 50.8), a width of 40.3, and a height of 59.2. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

The rectangle defines the boundaries of the gradient in the following ways: the right and left sides of the rectangle form the boundaries of a horizontal gradient; the top and bottom sides form the boundaries of a vertical gradient; two of the diagonally opposing corners lie on the boundaries of a diagonal gradient. In each of these cases, either side/corner may be on the starting boundary, depending on how the starting and ending points are passed to the constructor.

#### Example

```
' ========================================================================================
' The following example creates a linear gradient brush. Then the code gets the brush's
' rectangle and draws it.
' ========================================================================================
SUB Example_GetRectangle (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a linear gradient brush.
   DIM pt1 AS GpPoint = GDIP_POINT(20, 10)
   DIM pt2 AS GpPoint = GDIP_POINT(60, 110)
   DIM linGrBrush AS CGpLinearGradientBrush = CGpLinearGradientBrush(@pt1, @pt2, _
      GDIP_ARGB(255, 0, 0, 0), GDIP_ARGB(255, 0, 0, 255))

   ' // Obtain information about the linear gradient brush.
   DIM rc AS GpRect
   linGrBrush.GetRectangle(@rc)

   ' // Draw the retrieved rectangle.
   DIM pen AS CGpPen = GDIP_ARGB(255, 0, 0, 0)
   graphics.DrawRectangle(@pen, @rc)

END SUB
' ========================================================================================
```

# <a name="GetTransformLGBrush"></a>GetTransform (CGpLinearGradientBrush)

Gets the transformation matrix of this linear gradient brush. 

```
FUNCTION GetTransform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMatrix* | Pointer to a **Matrix** object that receives the transformation matrix. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A **LinearGradientBrush** object maintains a transformation matrix that can store any affine transformation. When you use a linear gradient brush to fill an area, GDI+ transforms the brush's boundary lines according to the brush's transformation matrix and then fills the area. The transformed boundaries exist only during rendering; the boundaries stored in the **LinearGradientBrush** object are not transformed

#### Example

```
' ========================================================================================
' The following example creates a linear gradient brush and sets its transformation matrix.
' Next, the code gets the brush's transformation matrix and proceeds to inspect or use the
' matrix elements.
' ========================================================================================
SUB Example_GetTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a linear gradient brush.
   DIM pt1 AS GpPoint = GDIP_POINT(0, 0)
   DIM pt2 AS GpPoint = GDIP_POINT(200, 0)
   DIM linGrBrush AS CGpLinearGradientBrush = CGpLinearGradientBrush(@pt1, @pt2, _
      GDIP_ARGB(255, 0, 0, 0), GDIP_ARGB(255, 0, 0, 255))

   DIM matrixSet AS CGpMatrix = CGpMatrix(0, 1, -1, 0, 0, 0)

   linGrBrush.SetTransform(@matrixSet)

   ' // Obtain information about the linear gradient brush.
   DIM matrixGet AS CGpMatrix
   DIM elements(0 TO 5) AS SINGLE

   linGrBrush.GetTransform(@matrixGet)
   matrixGet.GetElements(@elements(0))

   graphics.FillRectangle(@CGpSolidBrush(GDIP_ARGB(255, 0, 0, 0)), 0, 0, 20, 20)

   FOR j AS LONG = 0 TO 5
      ' // Inspect or use the value in elements[j].
      PRINT STR(elements(j))
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetWrapModeLGBrush"></a>GetWrapMode (CGpLinearGradientBrush)

Gets the wrap mode for this brush. The wrap mode determines how an area is tiled when it is painted with a brush.

```
FUNCTION GetWrapMode () AS WrapMode
```

#### Return value

This method returns one of the following elements of the **WrapMode** enumeration:

* WrapModeTile
* WrapModeTileFlipX
* WrapModeTileFlipY
* WrapModeTileFlipXY

#### Example

```
' ========================================================================================
' The following example creates a linear gradient brush and sets its wrap mode. Next, the
' code gets the brush's wrap mode and performs tasks based on the brush's current wrap mode.
' ========================================================================================
SUB Example_GetWrapMode (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a linear gradient brush.
   DIM rc AS GpRect = GDIP_RECT(0, 0, 100, 50)
   DIM linGrBrush AS CGpLinearGradientBrush = CGpLinearGradientBrush(@rc, _
      GDIP_ARGB(255, 0, 0, 0), GDIP_ARGB(255, 0, 0, 255), LinearGradientModeHorizontal)

   linGrBrush.SetWrapMode(WrapModeTileFlipX)

   ' // Obtain information about the linear gradient brush.
   DIM nWrapMode AS WrapMode
   nWrapMode = linGrBrush.GetWrapMode

   IF nWrapMode = WrapModeTileFlipX THEN
      ' // Do some task
   ELSEIF nWrapMode = WrapModeTileFlipY THEN
      ' // Do a different task
   END IF

END SUB
' ========================================================================================
```

# <a name="MultiplyTransformLGBrush"></a>MultiplyTransform (CGpLinearGradientBrush)

Updates this brush's transformation matrix with the product of itself and another matrix.

```
FUNCTION MultiplyTransform (BYVAL pMatrix AS CGpMatrix PTR, _
   BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMatrix* | Pointer to a matrix to be multiplied by the brush's current transformation matrix. |
| *order* | Optional. Element of the MatrixOrder enumeration that specifies the order of multiplication. **MatrixOrderPrepend** specifies that the passed matrix is on the left, and **MatrixOrderAppend** specifies that the passed matrix is on the right. The default value is **MatrixOrderPrepend**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A single 3 ×3 matrix can store any sequence of affine transformations. If you have several 3 ×3 matrices, each of which represents an affine transformation, the product of those matrices is a single 3 ×3 matrix that represents the entire sequence of transformations. The transformation represented by that product is called a composite transformation. For example, suppose matrix R represents a rotation, and matrix T represents a translation. If matrix M is the product RT, then matrix M represents a composite transformation: first rotate, then translate.

The order of matrix multiplication is important. In general, the matrix product RT is not the same as the matrix product TR. In the example given in the previous paragraph, the composite transformation represented by RT (first rotate, then translate) is not the same as the composite transformation represented by TR (first translate, then rotate).

#### Example

```
' ========================================================================================
' The following example creates a linear gradient brush and uses it to fill a rectangle.
' Next, the code sets the brush's transformation matrix, fills a rectangle with the
' transformed brush, modifies the brush's transformation matrix, and again fills a rectangle
' with the transformed brush.
' ========================================================================================
SUB Example_MultiplyTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM S AS CGpMatrix = CGpMatrix(2, 0, 0, 1, 0, 0)    ' // horizontal doubling
   DIM T AS CGpMatrix = CGpMatrix(1, 0, 0, 1, 50, 0)   '  // horizontal translation of 50 units

   DIM rc AS GpRect = GDIP_RECT(0, 0, 200, 100)
   DIM linGrBrush AS CGpLinearGradientBrush = CGpLinearGradientBrush(@rc, _
       GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255), LinearGradientModeHorizontal)

   ' // Fill a large area with the gradient brush (no transformation).
   graphics.FillRectangle(@linGrBrush, 0, 0, 800, 100)

   ' // Apply the scaling transformation.
   linGrBrush.SetTransform(@S)

   ' // Fill a large area with the scaled gradient brush.
   graphics.FillRectangle(@linGrBrush, 0, 150, 800, 100)

   ' // Form a composite transformation: first scale, then translate.
   linGrBrush.MultiplyTransform(@T, MatrixOrderAppend)

   ' // Fill a large area with the scaled and translated gradient brush.
   graphics.FillRectangle(@linGrBrush, 0, 300, 800, 100)

END SUB
' ========================================================================================
```

# <a name="ResetTransformLGBrush"></a>ResetTransform (CGpLinearGradientBrush)

Resets the transformation matrix of this linear gradient brush to the identity matrix. This means that no transformation takes place.

```
FUNCTION ResetTransform () AS GpStatus
```

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
 ========================================================================================
' The following example creates a linear gradient brush and uses it to fill a rectangle.
' Next, the code sets the brush's transformation matrix, fills a rectangle with the
' transformed brush, resets the brush's transformation matrix, and fills a rectangle with
' the untransformed brush.
' ========================================================================================
SUB Example_ResetTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM S AS CGpMatrix = CGpMatrix(2, 0, 0, 1, 0, 0)    ' // horizontal doubling

   DIM rc AS GpRect = GDIP_RECT(0, 0, 200, 100)
   DIM linGrBrush AS CGpLinearGradientBrush = CGpLinearGradientBrush(@rc, _
       GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255), LinearGradientModeHorizontal)

   ' // Fill a large area with the gradient brush (no transformation).
   graphics.FillRectangle(@linGrBrush, 0, 0, 800, 100)

   ' // Apply the scaling transformation.
   linGrBrush.SetTransform(@S)

   ' // Fill a large area with the scaled gradient brush.
   graphics.FillRectangle(@linGrBrush, 0, 150, 800, 100)

   ' // Reset the transformation
   linGrBrush.ResetTransform

   ' // Fill a large area with the gradient brush (no transformation)
   graphics.FillRectangle(@linGrBrush, 0, 300, 800, 100)

END SUB
' ========================================================================================
```

# <a name="RotateTransformLGBrush"></a>RotateTransform (CGpLinearGradientBrush)

Updates this brush's current transformation matrix with the product of itself and a rotation matrix.

```
FUNCTION RotateTransform (BYVAL angle AS SINGLE, _
   BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *angle* | Single precision number that specifies the angle of rotation in degrees. |
| *order* | Optional. Element of the MatrixOrder enumeration that specifies the order of multiplication. **MatrixOrderPrepend** specifies that the passed matrix is on the left, and **MatrixOrderAppend** specifies that the passed matrix is on the right. The default value is **MatrixOrderPrepend**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A single 3 ×3 matrix can store any sequence of affine transformations. If you have several 3 ×3 matrices, each of which represents an affine transformation, the product of those matrices is a single 3 ×3 matrix that represents the entire sequence of transformations. The transformation represented by that product is called a composite transformation. For example, suppose matrix T represents a translation, and matrix R represents a rotation. If matrix M is the product TR, then matrix M represents a composite transformation: first translate, then rotate.

The order of matrix multiplication is important. In general, the matrix product RT is not the same as the matrix product TR. In the example given in the previous paragraph, the composite transformation represented by RT (first rotate, then translate) is not the same as the composite transformation represented by TR (first translate, then rotate).

#### Example

```
' ========================================================================================
' The following example creates a linear gradient brush and uses it to fill a rectangle.
' Next, the code modifies the brush's transformation matrix, applying a composite transformation,
' and then fills a rectangle with the transformed brush.
' ========================================================================================
SUB Example_RotateTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM rc AS GpRect = GDIP_RECT(0, 0, 80, 40)
   DIM linGrBrush AS CGpLinearGradientBrush = CGpLinearGradientBrush(@rc, _
       GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255), LinearGradientModeHorizontal)

   ' // Fill a large area with the gradient brush (no transformation).
   graphics.FillRectangle(@linGrBrush, 0, 0, 800, 150)

   ' // Apply a composite transformation: first scale, then rotate.
   linGrBrush.ScaleTransform(2, 1)   '                 ' // horizontal doubling
   linGrBrush.RotateTransform(20, MatrixOrderAppend)   ' // 20-degree rotation

   ' // Fill a large area with the transformed linear gradient brush.
   graphics.FillRectangle(@linGrBrush, 0, 200, 800, 150)

END SUB
' ========================================================================================
```

# <a name="ScaleTransformLGBrush"></a>ScaleTransform (CGpLinearGradientBrush)

Updates this brush's current transformation matrix with the product of itself and a scaling matrix.

```
FUNCTION ScaleTransform (BYVAL sx AS SINGLE, BYVAL sy AS SINGLE, _
   BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *sx* | Single precision number that specifies the amount to scale in the x direction. |
| *sy* | Single precision number that specifies the amount to scale in the y direction. |
| *order* | Optional. Element of the **MatrixOrder** enumeration that specifies the order of multiplication. **MatrixOrderPrepend** specifies that the passed matrix is on the left, and **MatrixOrderAppend** specifies that the passed matrix is on the right. The default value is **MatrixOrderPrepend**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A single 3 ×3 matrix can store any sequence of affine transformations. If you have several 3 ×3 matrices, each of which represents an affine transformation, the product of those matrices is a single 3 ×3 matrix that represents the entire sequence of transformations. The transformation represented by that product is called a composite transformation. For example, suppose matrix T represents a translation, and matrix S represents a scaling. If matrix M is the product TS, then matrix M represents a composite transformation: first translate, then scale.

The order of matrix multiplication is important. In general, the matrix product RT is not the same as the matrix product TR. In the example given in the previous paragraph, the composite transformation represented by RT (first rotate, then translate) is not the same as the composite transformation represented by TR (first translate, then rotate).

#### Example

```
' ========================================================================================
' The following example creates a linear gradient brush and uses it to fill a rectangle.
' Next, the code modifies the brush's transformation matrix, applying a composite transformation,
' and then fills a rectangle with the transformed brush.
' ========================================================================================
SUB Example_ScaleTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM rc AS GpRect = GDIP_RECT(0, 0, 80, 40)
   DIM linGrBrush AS CGpLinearGradientBrush = CGpLinearGradientBrush(@rc, _
       GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255), LinearGradientModeHorizontal)

   ' // Fill a large area with the gradient brush (no transformation).
   graphics.FillRectangle(@linGrBrush, 0, 0, 800, 150)

   ' // Apply a composite transformation: first translate, then scale.
   linGrBrush.RotateTransform(60, 0)   ' // horizontal translation
   linGrBrush.ScaleTransform(2, 1)     ' // horizontal doubling

   ' // Fill a large area with the transformed linear gradient brush.
   graphics.FillRectangle(@linGrBrush, 0, 200, 800, 150)

END SUB
' ========================================================================================
```

# <a name="SetBlendLGBrush"></a>SetBlend (CGpLinearGradientBrush)

Sets the blend factors and the blend positions of this linear gradient brush to create a custom blend.

```
FUNCTION SetBlend (BYVAL blendFactors AS SINGLE PTR, BYVAL blendPositions AS SINGLE PTR, _
   BYVAL count AS LONG) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *blendFactors* | Pointer to an array of single precision numbers that specify blend factors. Each number in the array specifies a percentage of the ending color and should be in the range from 0.0 through 1.0. |
| *blendPositions* | Pointer to an array of single precision numbers that specify blend positions. Each number in the array indicates a percentage of the distance between the starting boundary and the ending boundary and is in the range from 0.0 through 1.0, where 0.0 indicates the starting boundary of the gradient and 1.0 indicates the ending boundary. There must be at least two positions specified: the first position, which is always 0.0f, and the last position, which is always 1.0f. Otherwise, the behavior is undefined. A blend position between 0.0 and 1.0 indicates a line, parallel to the boundary lines, that is a certain fraction of the distance from the starting boundary to the ending boundary. For example, a blend position of 0.7 indicates the line that is 70 percent of the distance from the starting boundary to the ending boundary. The color is constant on lines that are parallel to the boundary lines. |
| *count* | Optional. Integer that specifies the number of elements in the blendFactors array. This is the same as the number of elements in the *blendPositions* array. The blend factor at a given array index corresponds to the blend position at that same array index. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A **LinearGradientBrush** object has two boundaries. When you fill an area with a linear gradient brush, the color changes gradually as you move from the starting boundary to the ending boundary. By default, the color is linearly related to the distance, but you can customize the relationship between color and distance by calling the **SetBlend** method.

#### Example

```
' ========================================================================================
' The following example creates a linear gradient brush, sets a custom blend, and uses the
' brush to fill a rectangle.
' ========================================================================================
SUB Example_SetBlend (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM factors(0 TO 3) AS SINGLE = {0.0, 0.4, 0.6, 1.0}
   DIM positions(0 TO 3) AS SINGLE = {0.0, 0.2, 0.8, 1.0}
   DIM rcf AS GpRectF = GDIP_RECTF(0, 0, 100, 50)

   DIM linGrBrush AS CGpLinearGradientBrush = CGpLinearGradientBrush(@rcf, GDIP_ARGB(255, 255, 0, 0), _
       GDIP_ARGB(255, 0, 0, 255), LinearGradientModeHorizontal)

   linGrBrush.SetBlend(@factors(0), @positions(0), 4)
   graphics.FillRectangle(@linGrBrush, @rcf)

END SUB
' ========================================================================================
```

# <a name="SetBlendBellShapeLGBrush"></a>SetBlendBellShape (CGpLinearGradientBrush)

Sets the blend shape of this path gradient brush.

```
FUNCTION SetBlendBellShape (BYVAL focus AS SINGLE, BYVAL scale AS SINGLE = 1.0) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *focus* | Single precision number that specifies where the center color will be at its highest intensity. This number must be in the range 0 through 1. |
| *scale* | Single precision number that specifies the maximum intensity of center color that gets blended with the boundary color. This number must be in the range 0 through 1. The default value is 1. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

By default, the color changes gradually from the starting color (color at the starting boundary of the linear gradient brush) to the ending color (color at the ending boundary of the linear gradient brush) as you move from the starting boundary to the ending boundary. You can customize the positioning and blending of the starting and ending colors by using the SetBlendBellShape method.

The **SetBlendBellShape** method customizes the blend so that it follows a bell-shaped curve with the extremes of the bell's base at the gradient's boundaries. The starting color, which, in a default blend, is at the starting boundary of a linear gradient brush, appears at the starting and ending boundaries of the linear gradient brush when a bell-shaped blend is applied. The position of the ending color, which, in a default blend, is at the ending boundary, is somewhere between the boundaries and is determined by the value of the focus. In other words, the focus specifies the position of the peak of the bell. For example, a focus value of 0.7 places the peak at 70 percent of the distance between the starting and ending boundaries. The ending color appears at this peak.

The ending color in a bell-shaped blend is a percentage of the gamut between the gradient's default-blend starting color and default-blend ending color. For example, suppose a linear gradient brush is constructed with red as the starting color and blue as the ending color. If **SetBlendBellShape** is called with a scale value of 0.8, the ending color in the bell shaped blend is a hue that is 80 percent between red and blue (20 percent red, 80 percent blue). A scale value of 1.0 produces an ending color that is 100 percent blue.

#### Example

```
' ========================================================================================
' The following example creates a linear gradient brush, sets a bell-shaped blend, and uses
' the brush to fill a rectangle. Twice more, the code sets a bell-shaped blend with different
' values and, each time, uses the brush to fill a rectangle.
' ========================================================================================
SUB Example_SetBlendBellShape (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM pt1 AS GpPoint = GDIP_POINT(0, 0)
   DIM pt2 AS GpPoint = GDIP_POINT(500, 0)

   DIM linGrBrush AS CGpLinearGradientBrush = CGpLinearGradientBrush(@pt1, @pt2, _
      GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255))

   linGrBrush.SetBlendBellShape(0.5, 0.6)
   graphics.FillRectangle(@linGrBrush, 0, 0, 500, 50)

   linGrBrush.SetBlendBellShape(0.5, 0.8)
   graphics.FillRectangle(@linGrBrush, 0, 75, 500, 50)

   linGrBrush.SetBlendBellShape(0.5, 1.0)
   graphics.FillRectangle(@linGrBrush, 0, 150, 500, 50)

END SUB
' ========================================================================================
```

# <a name="SetBlendTriangularShapeLGBrush"></a>SetBlendTriangularShape (CGpLinearGradientBrush)

Sets the blend shape of this path gradient brush.

```
FUNCTION SetBlendTriangularShape (BYVAL focus AS SINGLE, BYVAL scale AS SINGLE = 1.0) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *focus* | Single precision number that specifies where the center color will be at its highest intensity. This number must be in the range 0 through 1. |
| *scale* | Single precision number that specifies the maximum intensity of center color that gets blended with the boundary color. This number must be in the range 0 through 1. The default value is 1. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

By default, the color changes gradually from the starting color (color at the starting boundary of the linear gradient brush) to the ending color (color at the ending boundary of the linear gradient brush) as you move from the starting boundary to the ending boundary. You can customize the positioning and blending of the starting and ending colors by using the **SetBlendTriangularShape** method.

The **SetBlendTriangularShape** method customizes the blend so that it follows a triangular shape with the extremes of the triangle's base at the gradient's boundaries. The starting color, which, in a default blend, is at the starting boundary of a linear gradient brush, appears at the starting and ending boundaries of the linear gradient brush when a triangular-shaped blend is applied. The position of the ending color, which, in a default blend, is at the ending boundary, is somewhere between the boundaries and is determined by the value of the focus. In other words, the focus specifies the position of the peak of the triangle. For example, a focus value of 0.5 places the peak half way between the starting and ending boundaries. The ending color appears at this peak.

The ending color in a triangular-shaped blend is a percentage of the gamut between the gradient's default-blend starting color and default-blend ending color. For example, suppose a linear gradient brush is constructed with red as the starting color and blue as the ending color. If **SetBlendTriangularShape** is called with a scale value of 0.3, the ending color in the triangular-shaped blend is a hue that is 30 percent between red and blue (70 percent red, 30 percent blue). A scale value of 1.0 produces an ending color that is 100 percent blue.

#### Example

```
' ========================================================================================
' The following example creates a linear gradient brush, sets a triangular-shaped blend,
' and uses the brush to fill a rectangle. Twice more, the code sets a triangular-shaped
' blend with different values and, each time, uses the brush to fill a rectangle.
' ========================================================================================
SUB Example_SetBlendTriangularShape (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM pt1 AS GpPoint = GDIP_POINT(0, 0)
   DIM pt2 AS GpPoint = GDIP_POINT(500, 0)

   DIM linGrBrush AS CGpLinearGradientBrush = CGpLinearGradientBrush(@pt1, @pt2, _
      GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255))

   linGrBrush.SetBlendTriangularShape(0.5, 0.6)
   graphics.FillRectangle(@linGrBrush, 0, 0, 500, 50)

   linGrBrush.SetBlendTriangularShape(0.5, 0.8)
   graphics.FillRectangle(@linGrBrush, 0, 75, 500, 50)

   linGrBrush.SetBlendTriangularShape(0.5, 1.0)
   graphics.FillRectangle(@linGrBrush, 0, 150, 500, 50)

END SUB
' ========================================================================================
```

# <a name="SetGammaCorrectionLGBrush"></a>SetGammaCorrection (CGpLinearGradientBrush)

Specifies whether gamma correction is enabled for this linear gradient brush.

```
FUNCTION SetGammaCorrection (BYVAL useGammaCorrection AS BOOL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *useGammaCorrection* | Boolean value that specifies whether gamma correction occurs during rendering. TRUE specifies that gamma correction is enabled, and FALSE specifies that gamma correction is not enabled. By default, gamma correction is disabled during construction of a **LinearGradientBrush** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

Gamma correction is often done to match the intensity contrast of the gradient to the ability of the human eye to perceive intensity changes.

# <a name="SetGammaCorrectionPGBrush"></a>SetGammaCorrection (CGpPathGradientBrush)

Specifies specifies whether gamma correction is enabled for this path gradient brush.

```
FUNCTION SetGammaCorrection (BYVAL useGammaCorrection AS BOOL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *useGammaCorrection* | Boolean value that specifies whether gamma correction is enabled. TRUE specifies that gamma correction is enabled, and FALSE specifies that gamma correction is not enabled. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

Gamma correction is often done to match the intensity contrast of the gradient to the ability of the human eye to perceive intensity changes.

# <a name="SetInterpolationColorsLGBrush"></a>SetInterpolationColors (CGpLinearGradientBrush)

Specifies whether gamma correction is enabled for this linear gradient brush.

```
FUNCTION SetInterpolationColors (BYVAL presetColors AS ARGB PTR, _
   BYVAL blendPositions AS SINGLE PTR, BYVAL count AS LONG) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *presetColors* | Pointer to an array of ARGB colors that specify the colors to be interpolated for this linear gradient brush. A color of a given index in the *presetColors* array corresponds to the blend position of that same index in the *blendPositions* array. |
| *blendPositions* | Pointer to an array of single precision numbers that specify the blend positions. Each number in the array specifies a percentage of the distance between the starting boundary and the ending boundary and is in the range from 0.0 through 1.0, where 0.0 indicates the starting boundary of the gradient and 1.0 indicates the ending boundary. There must be at least two positions specified: the first position, which is always 0.0f, and the last position, which is always 1.0f. Otherwise, the behavior is undefined. A blend position between 0.0 and 1.0 indicates the line, parallel to the boundary lines, that is a certain fraction of the distance from the starting boundary to the ending boundary. For example, a blend position of 0.7 indicates the line that is 70 percent of the distance from the starting boundary to the ending boundary. The color is constant on lines that are parallel to the boundary lines. |
| *count* | Integer that specifies the number of elements in the *presetColors* array. This is the same as the number of elements in the *blendPositions* array. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a linear gradient brush, sets the colors to be interpolated
' for the linear gradient brush, and fills a rectangle.
' ========================================================================================
SUB Example_SetInterpolationColors (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a linear gradient brush, and set the colors to be interpolated.
   DIM colors(0 TO 2) AS ARGB = {GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255), GDIP_ARGB(255, 0, 255, 0)}
   DIM positions(0 TO 2) AS SINGLE = {0.0, 0.3, 1.0}

   DIM pt1 AS GpPoint = GDIP_POINT(0, 0)
   DIM pt2 AS GpPoint = GDIP_POINT(100, 0)

   DIM linGrBrush AS CGpLinearGradientBrush = CGpLinearGradientBrush(@pt1, @pt2, GDIP_ARGB(255, 0, 0, 0), GDIP_ARGB(255, 255, 255, 255))
   linGrBrush.SetInterpolationColors(@colors(0), @positions(0), 3)

   graphics.FillRectangle(@linGrBrush, 0, 0, 100, 50)
   
END SUB
' ========================================================================================
```

# <a name="SetLinearColors"></a>SetLinearColors (CGpLinearGradientBrush)

Sets the starting color and ending color of this linear gradient brush.

```
FUNCTION SetLinearColors (BYVAL color1 AS ARGB, BYVAL color2 AS ARGB) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *color1* | The color at the starting boundary line of this linear gradient brush. |
| *color2* | The color that specifies the color at the ending boundary line of this linear gradient brush. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a linear gradient brush and uses it to fill a rectangle.
' Next, the code changes the linear colors and uses the modified brush to fill another rectangle.
' ========================================================================================
SUB Example_SetLinearColors (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a linear gradient brush.
   DIM rc AS GpRect = GDIP_RECT(0, 0, 100, 50)
   DIM linGrBrush AS CGpLinearGradientBrush = CGpLinearGradientBrush(@rc, _
      GDIP_ARGB(255, 0, 0, 0), GDIP_ARGB(255, 0, 0, 255), LinearGradientModeHorizontal)

   linGrBrush.SetLinearColors(GDIP_ARGB(255, 0, 0, 255), GDIP_ARGB(255, 0, 255, 0))
   graphics.FillRectangle(@linGrBrush, 0, 75, 100, 50)

END SUB
' ========================================================================================
```

# <a name="SetTransformLGBrush"></a>SetTransform (CGpLinearGradientBrush)

Sets the transformation matrix of this linear gradient brush.

```
FUNCTION SetTransform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMatrix* | Pointer to a **Matrix** object that specifies the transformation matrix to use. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A **LinearGradientBrush** object has a rectangle that specifies the starting and ending boundaries of the gradient and a mode or angle that affects the direction. If the brush's transformation matrix is set to represent any transformation other than the identity, then the boundaries and direction are transformed according to that matrix during rendering.

The transformation applies only during rendering. The boundaries stored by the **LinearGradientBrush** object are not altered by the **SetTransform** method.

#### Example

```
' ========================================================================================
' The following example creates a linear gradient brush and uses it to fill a rectangle.
' Next, the code modifies the brush's transformation matrix and fills a rectangle with the
' transformed brush.
' ========================================================================================
SUB Example_SetTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM rc AS GpRect = GDIP_RECT(0, 0, 80, 40)
   DIM linGrBrush AS CGpLinearGradientBrush = CGpLinearGradientBrush(@rc, _
       GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255), LinearGradientModeHorizontal)

   DIM matrix AS CGpMatrix = CGpMatrix(2.0, 0, 0, 1, 0, 0)   ' // horizontal doubling

   ' // Fill a large area with the gradient brush (no transformation).
   graphics.FillRectangle(@linGrBrush, 0, 0, 800, 50)

   linGrBrush.SetTransform(@matrix)

   ' // Fill a large area with the transformed linear gradient brush.
   graphics.FillRectangle(@linGrBrush, 0, 75, 800, 50)

END SUB
' ========================================================================================
```

# <a name="SetWrapModeLGBrush"></a>SetWrapMode (CGpLinearGradientBrush)

Sets the wrap mode of this linear gradient brush.

```
FUNCTION SetWrapMode (BYVAL wrapMode AS WrapMode) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *wrapMode* | Element of the **WrapMode** enumeration that specifies how areas painted with this linear gradient brush will be tiled. The value of this parameter must be one of the following elements: **WrapModeTile**, **WrapModeTileFlipX**, **WrapModeTileFlipY**, **WrapModeTileFlipXY**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

The boundary lines of a linear gradient brush form a tile. When you paint an area with a linear gradient brush, the tile repeats. A linear gradient brush may have alternate tiles flipped in a certain direction, as specified by the wrap mode. Flipping has the effect of reversing the order of the colors.

The wrap mode defaults to **WrapModeTile** when a **LinearGradientBrush** object is constructed.

#### Example

```
' ========================================================================================
' The following example creates a linear gradient brush and uses it to fill a rectangle.
' Next, the code modifies the brush's wrap mode and uses the modified brush to fill another
' rectangle.
' ========================================================================================
SUB Example_SetWrapMode (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a linear gradient brush.
   DIM rc AS GpRect = GDIP_RECT(0, 0, 100, 50)
   DIM linGrBrush AS CGpLinearGradientBrush = CGpLinearGradientBrush(@rc, _
      GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255), LinearGradientModeHorizontal)

   ' // Fill a large area using the gradient brush with the default wrap mode.
   graphics.FillRectangle(@linGrBrush, 0, 0, 800, 50)

   linGrBrush.SetWrapMode(WrapModeTileFlipX)

   ' // Fill a large area using the gradient brush with the new wrap mode.
   graphics.FillRectangle(@linGrBrush, 0, 75, 800, 50)

END SUB
' ========================================================================================
```

# <a name="TranslateTransformLGBrush"></a>TranslateTransform (CGpLinearGradientBrush)

Updates this brush's current transformation matrix with the product of itself and a translation matrix.

```
FUNCTION TranslateTransform (BYVAL dx AS SINGLE, BYVAL dy AS SINGLE, _
   BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *dx* | Single precision number that specifies the horizontal component of the translation. |
| *dy* | Single precision number that specifies the vertical component of the translation. |
| *order* | Optional. Element of the **MatrixOrder** enumeration that specifies the order of multiplication. **MatrixOrderPrepend** specifies that the passed matrix is on the left, and **MatrixOrderAppend** specifies that the passed matrix is on the right. The default value is **MatrixOrderPrepend**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A single 3×3 matrix can store any sequence of affine transformations. If you have several 3×3 matrices, each of which represents an affine transformation, the product of those matrices is a single 3×3 matrix that represents the entire sequence of transformations. The transformation represented by that product is called a composite transformation. For example, suppose matrix *S* represents a scaling, and matrix *T* represents a translation. If matrix M is the product *ST*, then matrix *M* represents a composite transformation: first scale, then translate.

The order of matrix multiplication is important. In general, the matrix product *RT* is not the same as the matrix product *TR*. In the example given in the previous paragraph, the composite transformation represented by *RT* (first rotate, then translate) is not the same as the composite transformation represented by *TR* (first translate, then rotate).

#### Example

```
' ========================================================================================
' The following example creates a linear gradient brush and uses it to fill a rectangle.
' Next, the code modifies the brush's transformation matrix, applying a composite transformation,
' and then fills a rectangle with the transformed brush.
' ========================================================================================
SUB Example_TranslateTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM rc AS GpRect = GDIP_RECT(0, 0, 80, 40)
   DIM linGrBrush AS CGpLinearGradientBrush = CGpLinearGradientBrush(@rc, _
       GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255), LinearGradientModeHorizontal)

   ' // Fill a large area with the gradient brush (no transformation).
   graphics.FillRectangle(@linGrBrush, 0, 0, 800, 150)

   ' // Apply a composite transformation: first scale, then translate.
   linGrBrush.ScaleTransform(2, 1)                           ' // horizontal doubling
   linGrBrush.TranslateTransform(30, 0, MatrixOrderAppend)   ' // translation

   ' // Fill a large area with the transformed linear gradient brush.
   graphics.FillRectangle(@linGrBrush, 0, 200, 800, 150)

END SUB
' ========================================================================================
```

# <a name="ConstructorsPGBrush"></a>Constructors (CGpPathGradientBrush)

Creates a **PathGradientBrush** object from another **PathGradientBrush** object.

```
CONSTRUCTOR CGpPathGradientBrush (BYVAL pPathGradientBrush AS CGpPathGradientBrush PTR)
```

Creates a **PathGradientBrush** object based on an array of points. Initializes the wrap mode of the path gradient brush

```
CONSTRUCTOR CGpPathGradientBrush (BYVAL pts AS GpPointF PTR, BYVAL nCount AS LONG, _
   BYVAL nWrapMode AS WrapMode = WrapModeClamp)
```

Creates a **PathGradientBrush** object based on an array of points. Initializes the wrap mode of the path gradient brush.

```
CONSTRUCTOR CGpPathGradientBrush (BYVAL pts AS GpPoint PTR, BYVAL nCount AS LONG, _
   BYVAL nWrapMode AS WrapMode = WrapModeClamp)
```

Creates a **PathGradientBrush** object based on a GraphicsPath object.

```
CONSTRUCTOR CGpPathGradientBrush (BYVAL pGraphPath AS CGpGraphicsPath PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pts* | Pointer to an array of points that specifies the boundary path of the path gradient brush. |
| *nCount* | Integer that specifies the number of elements in the *pts* array.  |
| *nWrapMode* | Optional. Element of the **WrapMode** enumeration that specifies how areas painted with the path gradient brush will be tiled. The default value is **WrapModeClamp**. |
| *pGraphPath* | Pointer to a **GraphicsPath** object that specifies the boundary path of the path gradient brush. |

# <a name="GetBlendPGBrush"></a>GetBlend (CGpPathGradientBrush)

Gets the blend factors and the corresponding blend positions currently set for this path gradient brush.

```
FUNCTION GetBlend (BYVAL blendFactors AS SINGLE PTR, BYVAL blendPositions AS SINGLE PTR, _
   BYVAL count AS LONG) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *blendFactors* | Pointer to an array that receives the blend factors. |
| *blendPositions* | Pointer to an array that receives the blend positions. |
| *count* | Integer that specifies the number of blend factors to retrieve. Before calling the **GetBlend** method of a **PathGradientBrush** object, call the **GetBlendCount** method of that same **PathGradientBrush** object to determine the current number of blend factors. The number of blend positions retrieved is the same as the number of blend factors retrieved. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A **PathGradientBrush** object has a boundary path and a center point. When you fill an area with a path gradient brush, the color changes gradually as you move from the boundary path to the center point. By default, the color is linearly related to the distance, but you can customize the relationship between color and distance by calling the SetBlend method.

#### Example

```
' ========================================================================================
' The following example demonstrates several methods of the PathGradientBrush class including
' PathGradientBrush.SetBlend, PathGradientBrush.GetBlendCount, and PathGradientBrush.GetBlend.
' The code creates a PathGradientBrush object and calls the PathGradientBrush.SetBlend method
' to establish a set of blend factors and blend positions for the brush. Then the code calls
' the PathGradientBrush.GetBlendCount method to retrieve the number of blend factors. After
' the number of blend factors is retrieved, the code allocates two buffers: one to receive
' the array of blend factors and one to receive the array of blend positions. Then the code
' calls the PathGradientBrush.GetBlend method to retrieve the blend factors and the blend
' positions.
' ========================================================================================
SUB Example_GetBlend (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a path that consists of a single ellipse.
   DIM path AS CGpGraphicsPath
   path.AddEllipse(0, 0, 200, 100)

   ' // Use the path to construct a brush.
   DIM pthGrBrush AS CGpPathGradientBrush = @path

   ' // Set the color at the center of the path to blue.
   pthGrBrush.SetCenterColor(GDIP_ARGB(255, 0, 0, 255))

   ' // Set the color along the entire boundary of the path to aqua.
   DIM colors(0) AS ARGB = {GDIP_ARGB(255, 0, 255, 255)}
   DIM count AS LONG = 1
   pthGrBrush.SetSurroundColors(@colors(0), @count)

   ' // Set blend factors and positions for the path gradient brush.
   DIM factors(0 TO 3) AS SINGLE = {0.0, 0.4, 0.8, 1.0}
   DIM positions(0 TO 3) AS SINGLE = {0.0, 0.3, 0.7, 1.0}

   pthGrBrush.SetBlend(@factors(0), @positions(0), 4)

   ' // Fill the ellipse with the path gradient brush.
   graphics.FillEllipse(@pthGrBrush, 0, 0, 200, 100)

   ' // Obtain information about the path gradient brush.
   DIM blendCount AS LONG = pthGrBrush.GetBlendCount
   DIM rgFactors(blendCount - 1) AS SINGLE
   DIM rgPositions(blendCount - 1) AS SINGLE

   pthGrBrush.GetBlend(@rgFactors(0), @rgPositions(0), blendCount)

   FOR j AS LONG = 0 TO blendCount - 1
'      // Inspect or use the value in rgFactors(j)
'      // Inspect or use the value in rgPositions(j)
      OutputDebugString STR(rgFactors(j)) & STR(rgPositions(j))
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetBlendCountPGBrush"></a>GetBlendCount (CGpPathGradientBrush)

Gets the number of blend factors currently set for this path gradient brush.

```
FUNCTION GetBlendCount () AS INT_
```

#### Return value

Before you call the **GetBlend** method of a **PathGradientBrush** object, you must allocate two buffers: one to receive an array of blend factors and one to receive an array of blend positions. To determine the size of the required buffers, call the **GetBlendCount** method of the **PathGradientBrush** object. The size (in bytes) of each buffer should be the return value of **GetBlendCount** multiplied by 4 (the size of a single precision number).

#### Example

```
' ========================================================================================
' The following example demonstrates several methods of the PathGradientBrush class including
' PathGradientBrush.SetBlend, PathGradientBrush.GetBlendCount, and PathGradientBrush.GetBlend.
' The code creates a PathGradientBrush object and calls the PathGradientBrush.SetBlend method
' to establish a set of blend factors and blend positions for the brush. Then the code calls
' the PathGradientBrush.GetBlendCount method to retrieve the number of blend factors. After
' the number of blend factors is retrieved, the code allocates two buffers: one to receive
' the array of blend factors and one to receive the array of blend positions. Then the code
' calls the PathGradientBrush.GetBlend method to retrieve the blend factors and the blend
' positions.
' ========================================================================================
SUB Example_GetBlend (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a path that consists of a single ellipse.
   DIM path AS CGpGraphicsPath
   path.AddEllipse(0, 0, 200, 100)

   ' // Use the path to construct a brush.
   DIM pthGrBrush AS CGpPathGradientBrush = @path

   ' // Set the color at the center of the path to blue.
   pthGrBrush.SetCenterColor(GDIP_ARGB(255, 0, 0, 255))

   ' // Set the color along the entire boundary of the path to aqua.
   DIM colors(0) AS ARGB = {GDIP_ARGB(255, 0, 255, 255)}
   DIM count AS LONG = 1
   pthGrBrush.SetSurroundColors(@colors(0), @count)

   ' // Set blend factors and positions for the path gradient brush.
   DIM factors(0 TO 3) AS SINGLE = {0.0, 0.4, 0.8, 1.0}
   DIM positions(0 TO 3) AS SINGLE = {0.0, 0.3, 0.7, 1.0}

   pthGrBrush.SetBlend(@factors(0), @positions(0), 4)

   ' // Fill the ellipse with the path gradient brush.
   graphics.FillEllipse(@pthGrBrush, 0, 0, 200, 100)

   ' // Obtain information about the path gradient brush.
   DIM blendCount AS LONG = pthGrBrush.GetBlendCount
   DIM rgFactors(blendCount - 1) AS SINGLE
   DIM rgPositions(blendCount - 1) AS SINGLE

   pthGrBrush.GetBlend(@rgFactors(0), @rgPositions(0), blendCount)

   FOR j AS LONG = 0 TO blendCount - 1
'      // Inspect or use the value in rgFactors(j)
'      // Inspect or use the value in rgPositions(j)
      OutputDebugString STR(rgFactors(j)) & STR(rgPositions(j))
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetCenterColor"></a>GetCenterColor (CGpPathGradientBrush)

Gets the center color of this path gradient brush.

```
FUNCTION GetCenterColor (BYVAL colour AS ARGB PTR) AS GpStatus
FUNCTION GetCenterColor () AS ARGB
```

| Parameter  | Description |
| ---------- | ----------- |
| *colour* | Pointer to a variable that receives the color of the center point. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

The second overloaded function returns the ARGB color as the result of the function.

#### Remarks

By default, the center point of a **PathGradientBrush** object is the centroid of the brush's boundary path, but you can set the center point to any location, inside or outside the path, by calling the **SetCenterPoint** method of the **PathGradientBrush** object.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object and uses it to fill an ellipse.
' Then the code calls the PathGradientBrush.GetCenterColor method of the PathGradientBrush
' object to obtain the center color.
' ========================================================================================
SUB Example_GetCenterColor (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a path that consists of a single ellipse.
   DIM path AS CGpGraphicsPath
   path.AddEllipse(0, 0, 200, 100)

   ' // Use the path to construct a brush.
   DIM pthGrBrush AS CGpPathGradientBrush = @path

   ' // Set the color at the center of the path to blue.
   pthGrBrush.SetCenterColor(GDIP_ARGB(255, 0, 0, 255))

   ' // Set the color along the entire boundary of the path to aqua.
   DIM colors(0) AS ARGB = {GDIP_ARGB(255, 0, 255, 255)}
   DIM count AS LONG = 1
   pthGrBrush.SetSurroundColors(@colors(0), @count)

   ' // Fill the ellipse with the path gradient brush.
   graphics.FillEllipse(@pthGrBrush, 0, 0, 200, 100)

   ' // Obtain information about the path gradient brush.
   DIM colour AS ARGB
   pthGrBrush.GetCenterColor(@colour)

   ' // Fill a rectangle with the retrieved color.
   DIM solidBrush AS CGpSolidBrush = colour
   graphics.FillRectangle(@solidBrush, 0, 120, 200, 30)

END SUB
' ========================================================================================
```

# <a name="GetCenterPoint"></a>GetCenterPoint (CGpPathGradientBrush)

Gets the center point of this path gradient brush.

```
FUNCTION GetCenterPoint (BYVAL pt AS PointF PTR) AS GpStatus
FUNCTION GetCenterPoint (BYVAL pt AS Point PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pt* | Pointer to a **PointF** or **Point** structure that receives the center point. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

By default, the center point of a **PathGradientBrush** object is at the centroid of the brush's boundary path, but you can set the center point to any location, inside or outside the path, by calling the **SetCenterPoint** method of the **PathGradientBrush** object.

#### Example

```
' ========================================================================================
' The following example demonstrates several methods of the PathGradientBrush class including
' PathGradientBrush.GetCenterPoint and PathGradientBrush.SetCenterColor. The code creates
' a PathGradientBrush object and then sets the brush's center color and boundary color.
' The code calls the PathGradientBrush.GetCenterPoint method to determine the center point
' of the path gradient and then draws a line from the origin to that center point.
' ========================================================================================
SUB Example_GetCenterPoint (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a path that consists of a single ellipse.
   DIM path AS CGpGraphicsPath
   path.AddEllipse(0, 0, 200, 100)

   ' // Use the path to construct a brush.
   DIM pthGrBrush AS CGpPathGradientBrush = @path

   ' // Set the color at the center of the path to blue.
   pthGrBrush.SetCenterColor(GDIP_ARGB(255, 0, 0, 255))

   ' // Set the color along the entire boundary of the path to aqua.
   DIM colors(0) AS ARGB = {GDIP_ARGB(255, 0, 255, 255)}
   DIM count AS LONG = 1
   pthGrBrush.SetSurroundColors(@colors(0), @count)

   ' // Fill the ellipse with the path gradient brush.
   graphics.FillEllipse(@pthGrBrush, 0, 0, 200, 100)

   ' // Obtain information about the path gradient brush.
   DIM centerPoint AS GpPointF
   pthGrBrush.GetCenterPoint(@centerPoint)

   ' // Draw a line from the origin to the center of the ellipse.
   DIM pen AS CGpPen = GDIP_ARGB(255, 0, 255, 0)
   DIM pt AS GpPointF = GDIP_POINTF(0, 0)
   graphics.DrawLine(@pen, @pt, @centerPoint)
   
END SUB
' ========================================================================================
```

# <a name="GetFocusScales"></a>GetFocusScales (CGpPathGradientBrush)

Gets the focus scales of this path gradient brush.

```
FUNCTION GetFocusScales (BYVAL xScale AS SINGLE PTR, BYVAL yScale AS SINGLE PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *xScale* | Pointer to a single precision that receives the x focus scale value. |
| *yScale* | Pointer to a single precision that receives the y focus scale value. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

By default, the center color of a path gradient is at the center point. By calling **SetFocusScales**, you can specify that the center color should appear along a path that surrounds the center point. For example, suppose the boundary path is a triangle and the center point is at the centroid of that triangle. Also assume that the boundary color is red and the center color is blue. If you set the focus scales to (0.2, 0.2), the color is blue along the boundary of a small triangle that surrounds the center point. That small triangle is the main boundary path scaled by a factor of 0.2 in the x direction and 0.2 in the y direction. When you paint with the path gradient brush, the color will change gradually from red to blue as you move from the boundary of the large triangle to the boundary of the small triangle. The area inside the small triangle will be filled with blue.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on a triangular path.
' The code sets the focus scales of the path gradient brush to (0.2, 0.2) and then uses
' the path gradient brush to fill an area that contains the triangular path. Finally, the
' code calls the PathGradientBrush.GetFocusScales method of the PathGradientBrush object
' to obtain the values of the x focus scale and the y focus scale.
' ========================================================================================
SUB Example_GetFocusScales (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM points(0 TO 2) AS GpPoint = {GDIP_POINT(100, 0), GDIP_POINT(200, 200), GDIP_POINT(0, 200)}

   ' // No GraphicsPath object is created. The PathGradientBrush
   ' // object is constructed directly from the array of points.
   DIM pthGrBrush AS CGpPathGradientBrush = CGpPathGradientBrush(@points(0), 3)

   DIM colors(0 TO 1) AS ARGB = {GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255)}

   ' // red at the boundary of the outer triangle
   ' // blue at the boundary of the inner triangle
   DIM relativePositions(0 TO 1) AS SINGLE = {0.0, 1.0}
   pthGrBrush.SetInterpolationColors(@colors(0), @relativePositions(0), 2)

   ' // The inner triangle is formed by scaling the outer triangle
   ' // about its centroid. The scaling factor is 0.2 in both the x and y directions.
   pthGrBrush.SetFocusScales(0.2, 0.2)

   ' // Fill a rectangle that is larger than the triangle
   ' // specified in the Point array. The portion of the
   ' // rectangle outside the triangle will not be painted.
   graphics.FillRectangle(@pthGrBrush, 0, 0, 200, 200)

   ' // Obtain information about the path gradient brush.
   DIM xScale AS SINGLE = 0.0
   DIM yScale AS SINGLE = 0.0
   pthGrBrush.GetFocusScales(@xScale, @yScale)

   ' // The value of xScale is now 0.2.
   ' // The value of yScale is now 0.2. 

END SUB
' ========================================================================================
```

# <a name="GetGammaCorrectionPGBrush"></a>GetGammaCorrection (CGpPathGradientBrush)

Determines whether gamma correction is enabled for this path gradient brush.

```
FUNCTION GetGammaCorrection () AS BOOL
```

#### Return value

If gamma correction is enabled, this method returns TRUE; otherwise, it returns FALSE.


# <a name="GetInterpolationColorCountPGBrush"></a>GetInterpolationColorCount (CGpPathGradientBrush)

Gets the number of preset colors currently specified for this path gradient brush.

```
FUNCTION GetInterpolationColorCount () AS INT_
```

#### Return value

This method returns the number of preset colors currently specified for this path gradient brush.

#### Remarks

Remarks

A simple path gradient brush has two colors: a boundary color and a center color. When you paint with such a brush, the color changes gradually from the boundary color to the center color as you move from the boundary path to the center point. You can create a more complex gradient by specifying an array of preset colors and an array of blend positions.

You can obtain the interpolation colors and interpolation positions currently set for a **PathGradientBrush** object by calling the **GetInterpolationColors** method of that **PathGradientBrush** object. Before you call the **GetInterpolationColors** method, you must allocate two buffers: one to hold the array of interpolation colors and one to hold the array of interpolation positions. You can call the **GetInterpolationColorCount** method of the **PathGradientBrush** object to determine the required size of those buffers. The size of the color buffer is the return value of **GetInterpolationColorCount** multiplied by 4. The size of the position buffer is the value of **GetInterpolationColorCount** multiplied by 4 (the size of a single precision number).


# <a name="GetInterpolationColorsPGBrush"></a>GetInterpolationColors (CGpPathGradientBrush)

Gets preset colors and blend positions currently specified for this path gradient brush.

```
FUNCTION GetInterpolationColors (BYVAL presetColors AS ARGB PTR, _
   BYVAL blendPositions AS SINGLE PTR, BYVAL count AS LONG) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *presetColors* | Pointer to an array that receives the colors. A color of a given index in the presetColors array corresponds to the blend position of that same index in the *blendPositions* array. |
| *blendPositions* | Pointer to an array that receives the blend positions. Each blend position is a number from 0 through 1, where 0 indicates the boundary of the gradient and 1 indicates the center point. A blend position between 0 and 1 indicates the set of all points that are a certain fraction of the distance from the boundary to the center point. For example, a blend position of 0.7 indicates the set of all points that are 70 percent of the way from the boundary to the center point. |
| *count* | Integer that specifies the number of elements in the *presetColorsarray*. This is the same as the number of elements in the *blendPositions* array. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A simple path gradient brush has two colors: a boundary color and a center color. When you paint with such a brush, the color changes gradually from the boundary color to the center color as you move from the boundary path to the center point. You can create a more complex gradient by specifying an array of preset colors and an array of blend positions.

Before you call the GetInterpolationColors method, you must allocate two buffers: one to hold the array of preset colors and one to hold the array of blend positions. You can call the **GetInterpolationColorCount** method of the **PathGradientBrush** object to determine the required size of those buffers. The size of the color buffer is the return value of **GetInterpolationColorCount** multiplied by 4. The size of the position buffer is the value of **GetInterpolationColorCount** multiplied by 4 (the size of a single precision number).

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object from a triangular path. The code
' sets the preset colors to red, blue, and aqua and sets the blend positions to 0, 0.6, and 1.
' The code calls the PathGradientBrush.GetInterpolationColorCount method of the PathGradientBrush
' object to obtain the number of preset colors currently set for the brush. Next, the code
' allocates two buffers: one to hold the array of preset colors, and one to hold the array
' of blend positions. The call to the PathGradientBrush.GetInterpolationColors method of
' the PathGradientBrush object fills the buffers with the preset colors and the blend positions.
' Finally the code fills a small square with each of the preset colors.
' ========================================================================================
SUB Example_GetInterpolationColors (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM points(0 TO 2) AS GpPoint = {GDIP_POINT(100, 0), GDIP_POINT(200, 200), GDIP_POINT(0, 200)}
   DIM pthGrBrush AS CGpPathGradientBrush = CGpPathGradientBrush(@points(0), 3)

   DIM colors(0 TO 2) AS ARGB = {GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255), GDIP_ARGB(255, 0, 255, 255)}
   DIM positions(0 TO 2) AS SINGLE = {0.0, 0.6, 1.0}

   pthGrBrush.SetInterpolationColors(@colors(0), @positions(0), 3)

   ' // Obtain information about the path gradient brush.
   DIM colorCount AS LONG = pthGrBrush.GetInterpolationColorCount
   DIM rgColors(colorCount - 1) AS ARGB
   DIM rgPositions(colorCount - 1) AS SINGLE
   pthGrBrush.GetInterpolationColors(@rgColors(0), @rgPositions(0), colorCount)

   ' // Fill a small square with each of the interpolation colors.
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 255, 255)

   FOR j AS LONG = 0 TO colorCount - 1
      solidBrush.SetColor(rgColors(j))
      graphics.FillRectangle(@solidBrush, 15 * j, 0, 10, 10)
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetPointCount"></a>GetPointCount (CGpPathGradientBrush)

Gets the number of points in the array of points that defines this brush's boundary path.

```
FUNCTION GetPointCount () AS INT_
```

#### Return value

This method returns the number of points in the array of points that defines this brush's boundary path.

# <a name="GetRectanglePGBrush"></a>GetRectangle (CGpPathGradientBrush)

Gets the smallest rectangle that encloses the boundary path of this path gradient brush.

```
FUNCTION GetRectangle (BYVAL rc AS GpRectF PTR) AS GpStatus
FUNCTION GetRectangle (BYVAL rc AS GpRect PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *rc* | Pointer to a **GpRectF** or **GpRect** structure that receives the bounding rectangle. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on a polygon that is defined
' by four points. The code calls the PathGradientBrush::GetRectangle method of the
' PathGradientBrush object to obtain the smallest rectangle that encloses the brush's
' boundary path. The code calls the Graphics.FillRectangle method of a Graphics object,
' passing the address of the PathGradientBrush object and a reference to the brush's bounding
' rectangle. That call fills only the portion of the bounding rectangle that is inside the
' brush's boundary path. Finally the code draws the outline of the bounding rectangle.
' ========================================================================================
SUB Example_GetRectangle (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM pen AS CGpPen = GDIP_ARGB(255, 0, 0, 0)

   ' // Create a path gradient brush based on an array of points.
   DIM points(0 TO 3) AS GpPoint = {GDIP_POINT(30, 20), GDIP_POINT(150, 40), GDIP_POINT(100, 100), GDIP_POINT(60, 200)}
   DIM pthGrBrush AS CGpPathGradientBrush = CGpPathGradientBrush(@points(0), 4)

   ' // Obtain information about the path gradient brush.
   DIM rc AS GpRectF
   pthGrBrush.GetRectangle(@rc)

   graphics.FillRectangle(@pthGrBrush, @rc)
   graphics.DrawRectangle(@pen, @rc)

END SUB
' ========================================================================================
```

# <a name="GetSurroundColorCount"></a>GetSurroundColorCount (CGpPathGradientBrush)

Gets the number of colors that have been specified for the boundary path of this path gradient brush.

```
FUNCTION GetSurroundColorCount () AS INT_
```

#### Return value

This method returns the number of colors that have been specified for the boundary path of this path gradient brush.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on a triangular path that
' is defined by an array of three points. The code calls the PathGradientBrush.SetSurroundColors
' method of the PathGradientBrush object to specify a color for each of the points that
' define the triangle. The PathGradientBrush.GetSurroundColorCount method determines the
' current number of surround colors (the colors specified for the brush's boundary path).
' Next, the code allocates a buffer large enough to receive the array of surround colors
' and calls PathGradientBrush.GetSurroundColors to fill that buffer. Finally the code fills
' a small square with each of the brush's surround colors.
' ========================================================================================
SUB Example_GetSurroundColors (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM points(0 TO 2) AS GpPoint = {GDIP_POINT(100, 0), GDIP_POINT(200, 200), GDIP_POINT(0, 200)}
   DIM pthGrBrush AS CGpPathGradientBrush = CGpPathGradientBrush(@points(0), 3)

   DIM nCount AS LONG = 3
   DIM colors(0 TO 2) AS ARGB = {GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255), GDIP_ARGB(255, 0, 255, 255)}
   pthGrBrush.SetSurroundColors(@colors(0), @nCount)

   ' // Obtain information about the path gradient brush.
   DIM colorCount AS LONG = pthGrBrush.GetSurroundColorCount
   DIM rgColors(colorCount - 1) AS ARGB
   pthGrBrush.GetSurroundColors(@rgColors(0), @colorCount)

   ' // Fill a small square with each of the surround colors.
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 255, 255)
   FOR j AS LONG = 0 TO colorCount - 1
      solidBrush.SetColor(rgColors(j))
      graphics.FillRectangle(@solidBrush, 15 * j, 0, 10, 10)
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetSurroundColors"></a>GetSurroundColors (CGpPathGradientBrush)

Gets the surround colors currently specified for this path gradient brush.

```
FUNCTION GetSurroundColors (BYVAL colors AS ARGB PTR, BYVAL count AS INT_ PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *colors* | Pointer to an array that receives the surround colors. |
| *count* | Out. Pointer to an integer that, on input, specifies the number of colors requested. If the method succeeds, this parameter, on output, receives the number of colors retrieved. If the method fails, this parameter does not receive a value. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A path gradient brush has a boundary path and a center point. The center point is set to a single color, but you can specify different colors for several points on the boundary. For example, suppose you specify red for the center color, and you specify blue, green, and yellow for distinct points on the boundary. Then as you move along the boundary, the color will change gradually from blue to green to yellow and back to blue. As you move along a straight line from any point on the boundary to the center point, the color will change from that boundary point's color to red.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on a triangular path that
' is defined by an array of three points. The code calls the PathGradientBrush.SetSurroundColors
' method of the PathGradientBrush object to specify a color for each of the points that
' define the triangle. The PathGradientBrush.GetSurroundColorCount method determines the
' current number of surround colors (the colors specified for the brush's boundary path).
' Next, the code allocates a buffer large enough to receive the array of surround colors
' and calls PathGradientBrush.GetSurroundColors to fill that buffer. Finally the code fills
' a small square with each of the brush's surround colors.
' ========================================================================================
SUB Example_GetSurroundColors (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM points(0 TO 2) AS GpPoint = {GDIP_POINT(100, 0), GDIP_POINT(200, 200), GDIP_POINT(0, 200)}
   DIM pthGrBrush AS CGpPathGradientBrush = CGpPathGradientBrush(@points(0), 3)

   DIM nCount AS LONG = 3
   DIM colors(0 TO 2) AS ARGB = {GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255), GDIP_ARGB(255, 0, 255, 255)}
   pthGrBrush.SetSurroundColors(@colors(0), @nCount)

   ' // Obtain information about the path gradient brush.
   DIM colorCount AS LONG = pthGrBrush.GetSurroundColorCount
   DIM rgColors(colorCount - 1) AS ARGB
   pthGrBrush.GetSurroundColors(@rgColors(0), @colorCount)

   ' // Fill a small square with each of the surround colors.
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 255, 255)
   FOR j AS LONG = 0 TO colorCount - 1
      solidBrush.SetColor(rgColors(j))
      graphics.FillRectangle(@solidBrush, 15 * j, 0, 10, 10)
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetTransformPGBrush"></a>GetTransform (CGpPathGradientBrush)

Gets the transformation matrix of this path gradient brush.

```
FUNCTION GetTransform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMatrix* | Pointer to a **Matrix** object that receives the transformation matrix. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A **PathGradientBrush** object maintains a transformation matrix that can store any affine transformation. When you use a path gradient brush to fill an area, GDI+ transforms the brush's boundary path according to the brush's transformation matrix and then fills the interior of the transformed path. The transformed path exists only during rendering; the boundary path stored in **PathGradientBrush** object is not transformed.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on an array of three points.
' The PathGradientBrush.ScaleTransform and PathGradientBrush.TranslateTransform methods set
' the elements of the brush's transformation matrix so that the matrix represents a composite
' transformation (first scale, then translate). That composite transformation applies to
' the brush's boundary path, so the call to FillRectangle fills the interior of a triangle
' that is the result of scaling and translating the boundary path. The code calls the
' PathGradientBrush.GetTransform method of the PathGradientBrush object to obtain the brush's
' transformation matrix and then calls the GetElements method of the retrieved Matrix object
' to fill an array with the matrix elements.
' ========================================================================================
SUB Example_GetTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM points(0 TO 2) AS GpPoint = {GDIP_POINT(0, 0), GDIP_POINT(50, 0), GDIP_POINT(50, 50)}
   DIM pthGrBrush AS CGpPathGradientBrush = CGpPathGradientBrush(@points(0), 3)

   pthGrBrush.ScaleTransform(3.0, 1.0)
   pthGrBrush.TranslateTransform(10.0, 30.0, MatrixOrderAppend)

   graphics.FillRectangle(@pthGrBrush, 0, 0, 200, 200)

   ' // Obtain information about the path gradient brush.
   DIM matrix AS CGpMatrix
   DIM elements(0 TO 5) AS SINGLE

   pthGrBrush.GetTransform(@matrix)
   matrix.GetElements(@elements(0))

   FOR j AS LONG = 0 TO 5
      ' // Inspect or use the value in elements(j)
   NEXT

END SUB
' ========================================================================================
```
# <a name="GetWrapModePGBrush"></a>GetWrapMode (CGpPathGradientBrush)

Gets the wrap mode currently set for this path gradient brush.

```
FUNCTION GetWrapMode () AS WrapMode
```

#### Return value

This method returns an element of the **WrapMode** enumeration that indicates the wrap mode currently set for this path gradient brush.

#### Remarks

The bounding rectangle of a path gradient brush is the smallest rectangle that encloses the brush's boundary path. When you paint the bounding rectangle with the path gradient brush, only the area inside the boundary path gets filled. The area inside the bounding rectangle but outside the boundary path does not get filled.

The default wrap mode for a path gradient brush is **WrapModeClamp**, which indicates that no painting occurs outside of the brush's bounding rectangle. All of the other wrap modes indicate that areas outside the brush's bounding rectangle will be tiled. Each tile is a copy (possibly flipped) of the filled path inside its bounding rectangle.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on a triangular path. The
' code calls the PathGradientBrush.SetWrapMode method of the PathGradientBrush object to
' set the wrap mode to WrapModeTileFlipX. Next, the code calls the PathGradientBrush.GetWrapMode
' method of the PathGradientBrush object to obtain the brush's wrap mode. If the obtained
' wrap mode is WrapModeTileFlipX, the code calls FillRectangle to tile a large area with
' the path gradient brush. 
' ========================================================================================
SUB Example_GetWrapMode (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM points(0 TO 2) AS GpPoint = {GDIP_POINT(0, 0), GDIP_POINT(100, 0), GDIP_POINT(100, 100)}
   DIM colors(0 TO 2) AS ARGB = {GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255), GDIP_ARGB(255, 0, 255, 0)}

   DIM nCount AS LONG = 3
   DIM pthGrBrush AS CGpPathGradientBrush = CGpPathGradientBrush(@points(0), 3)
   pthGrBrush.SetSurroundColors(@colors(0), @nCount)
   pthGrBrush.SetWrapMode(WrapModeTileFlipX)

   ' // Obtain information about the path gradient brush.
   DIM nWrapMode AS WrapMode = pthGrBrush.GetWrapMode

   IF nWrapMode = WrapModeTileFlipX THEN
      graphics.FillRectangle(@pthGrBrush, 0, 0, 800, 800)
   END IF

END SUB
' ========================================================================================
```

# <a name="MultiplyTransformPGBrush"></a>MultiplyTransform (CGpPathGradientBrush)

Updates the brush's transformation matrix with the product of itself and another matrix.

```
FUNCTION MultiplyTransform (BYVAL pMatrix AS CGpMatrix PTR, _
   BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMatrix* | Pointer to a matrix to be multiplied by the brush's current transformation matrix. |
| *order* | Optional. Element of the **MatrixOrder** enumeration that specifies the order of multiplication. **MatrixOrderPrepend** specifies that the passed matrix is on the left, and **MatrixOrderAppend** specifies that the passed matrix is on the right. The default value is **MatrixOrderPrepend**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A single 3 ×3 matrix can store any sequence of affine transformations. If you have several 3 ×3 matrices, each of which represents an affine transformation, the product of those matrices is a single 3 ×3 matrix that represents the entire sequence of transformations. The transformation represented by that product is called a composite transformation. For example, suppose matrix R represents a rotation and matrix T represents a translation. If matrix M is the product RT, then matrix M represents a composite transformation: first rotate, then translate.

The order of matrix multiplication is important. In general, the matrix product RT is not the same as the matrix product TR. In the example given in the previous paragraph, the composite transformation represented by RT (first rotate, then translate) is not the same as the composite transformation represented by TR (first translate, then rotate).

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on a triangular path. The
' code calls the PathGradientBrush.ScaleTransform method of the PathGradientBrush object
' to fill the brush's transformation matrix with the elements that represent a horizontal
' scaling by a factor of 3. Then the code calls the PathGradientBrush.MultiplyTransform
' method of that same PathGradientBrush object to multiply the brush's existing transformation
' matrix by a matrix that represents a translation (10 right, 30 down). The MatrixOrderAppend
' argument indicates that the multiplication is performed with the translation matrix on the right.
' After the multiplication, the brush's transformation matrix represents a composite
' transformation: first scale, then translate. That composite transformation is applied to
' the brush's boundary path during the call to FillRectangle, so it is the area inside the
' transformed path that gets painted.
' ========================================================================================
SUB Example_MultiplyTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM points(0 TO 2) AS GpPoint = {GDIP_POINT(0, 0), GDIP_POINT(50, 0), GDIP_POINT(50, 50)}

   ' // Translate 10 right, 30 down.
   DIM Matrix AS CGpMatrix = CGpMatrix(1.0, 0.0, 0.0, 1.0, 10.0, 30.0)

   DIM pthGrBrush AS CGpPathGradientBrush = CGpPathGradientBrush(@points(0), 3)
   pthGrBrush.ScaleTransform(3.0, 1.0)
   pthGrBrush.MultiplyTransform(@matrix, MatrixOrderAppend)

   graphics.FillRectangle(@pthGrBrush, 0, 0, 200, 200)

END SUB
' ========================================================================================
```
# <a name="ResetTransformPGBrush"></a>ResetTransform (CGpPathGradientBrush)

Resets the transformation matrix of this path gradient brush to the identity matrix. This means that no transformation will take place.

```
FUNCTION ResetTransform () AS GpStatus
```

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on a triangular path. The
' code calls the PathGradientBrush.ScaleTransform method of the PathGradientBrush object
' to fill the brush's transformation matrix with the elements that represent a horizontal
' scaling by a factor of 3. Then the code calls the PathGradientBrush.MultiplyTransform
' method of that same PathGradientBrush object to multiply the brush's existing transformation
' matrix by a matrix that represents a translation (10 right, 30 down). The MatrixOrderAppend
' argument indicates that the multiplication is performed with the translation matrix on the right.
' After the multiplication, the brush's transformation matrix represents a composite
' transformation: first scale, then translate. That composite transformation is applied to
' the brush's boundary path during the call to FillRectangle, so it is the area inside the
' transformed path that gets painted.
' ========================================================================================
SUB Example_ResetTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM points(0 TO 2) AS GpPoint = {GDIP_POINT(0, 0), GDIP_POINT(50, 0), GDIP_POINT(50, 50)}
   DIM pthGrBrush AS CGpPathGradientBrush = CGpPathGradientBrush(@points(0), 3)

   pthGrBrush.ScaleTransform(3.0, 1.0)
   pthGrBrush.TranslateTransform(100.0, 50.0, MatrixOrderAppend)

   ' // Fill an area with the transformed path gradient brush.
   graphics.FillRectangle(@pthGrBrush, 0, 0, 500, 500)

   pthGrBrush.ResetTransform

   ' // Fill the same area with the path gradient brush (no transformation).
   graphics.FillRectangle(@pthGrBrush, 0, 0, 500, 500)

END SUB
' ========================================================================================
```

# <a name="RotateTransformPGBrush"></a>RotateTransform (CGpPathGradientBrush)

Updates this brush's current transformation matrix with the product of itself and a rotation matrix.

```
FUNCTION RotateTransform (BYVAL angle AS SINGLE, _
   BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *angle* | Single precision number that specifies the angle of rotation in degrees. |
| *order* | Optional. Element of the MatrixOrder enumeration that specifies the order of multiplication. **MatrixOrderPrepend** specifies that the passed matrix is on the left, and **MatrixOrderAppend** specifies that the passed matrix is on the right. The default value is **MatrixOrderPrepend**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A single 3 ×3 matrix can store any sequence of affine transformations. If you have several 3 ×3 matrices, each of which represents an affine transformation, the product of those matrices is a single 3 ×3 matrix that represents the entire sequence of transformations. The transformation represented by that product is called a composite transformation. For example, suppose matrix T represents a translation and matrix R represents a rotation. If matrix M is the product TR, then matrix M represents a composite transformation: first translate, then rotate.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on a triangular path. The
' calls to the PathGradientBrush.ScaleTransform and PathGradientBrush.RotateTransform methods
' of the PathGradientBrush object set the elements of the brush's transformation matrix so
' that it represents a composite transformation: first scale, then rotate. The code uses
' the path gradient brush twice to paint a rectangle: once before the transformation is set
' and once after the transformation is set.
' ========================================================================================
SUB Example_RotateTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM points(0 TO 2) AS GpPoint = {GDIP_POINT(0, 0), GDIP_POINT(50, 0), GDIP_POINT(50, 50)}
   DIM pthGrBrush AS CGpPathGradientBrush = CGpPathGradientBrush(@points(0), 3)

   ' // Fill an area with the path gradient brush (no transformation).
   graphics.FillRectangle(@pthGrBrush, 0, 0, 500, 500)

   pthGrBrush.ScaleTransform(3.0, 1.0)                   ' // first scale
   pthGrBrush.RotateTransform(60.0, MatrixOrderAppend)   ' // then rotate

   ' // Fill the same area with the transformed path gradient brush.
   graphics.FillRectangle(@pthGrBrush, 0, 0, 500, 500)

END SUB
' ========================================================================================
```

# <a name="ScaleTransformPGBrush"></a>ScaleTransform (CGpPathGradientBrush)

Updates this brush's current transformation matrix with the product of itself and a scaling matrix.

```
FUNCTION ScaleTransform (BYVAL sx AS SINGLE, BYVAL sy AS SINGLE, _
   BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *sx* | Single precision number that specifies the horizontal scale factor. |
| *sy* | Single precision number that specifies the vertical scale factor. |
| *order* | Optional. Element of the MatrixOrder enumeration that specifies the order of multiplication. **MatrixOrderPrepend** specifies that the passed matrix is on the left, and **MatrixOrderAppend** specifies that the passed matrix is on the right. The default value is **MatrixOrderPrepend**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A single 3 ×3 matrix can store any sequence of affine transformations. If you have several 3 ×3 matrices, each of which represents an affine transformation, the product of those matrices is a single 3 ×3 matrix that represents the entire sequence of transformations. The transformation represented by that product is called a composite transformation. For example, suppose matrix T represents a translation and matrix S represents a scaling. If matrix M is the product TS, then matrix M represents a composite transformation: first translate, then scale.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on a triangular path. The
' calls to the PathGradientBrush::TranslateTransform and PathGradientBrush.ScaleTransform
' methods of the PathGradientBrush object set the elements of the brush's transformation
' matrix so that it represents a composite transformation: first translate, then scale. The
' code uses the path gradient brush twice to paint a rectangle: once before the transformation
' is set and once after the transformation is set.
' ========================================================================================
SUB Example_ScaleTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM points(0 TO 2) AS GpPoint = {GDIP_POINT(0, 0), GDIP_POINT(50, 0), GDIP_POINT(50, 50)}
   DIM pthGrBrush AS CGpPathGradientBrush = CGpPathGradientBrush(@points(0), 3)

   ' // Fill an area with the path gradient brush (no transformation).
   graphics.FillRectangle(@pthGrBrush, 0, 0, 500, 500)

   pthGrBrush.TranslateTransform(50.0, 40.0)                ' // translate
   pthGrBrush.ScaleTransform(3.0, 2.0, MatrixOrderAppend)   ' // then scale

   ' // Fill the same area with the transformed path gradient brush.
   graphics.FillRectangle(@pthGrBrush, 0, 0, 500, 500)

END SUB
' ========================================================================================
```

# <a name="SetBlendPGBrush"></a>SetBlend (CGpPathGradientBrush)

Sets the blend factors and the blend positions of this path gradient brush.

```
FUNCTION SetBlend (BYVAL blendFactors AS SINGLE PTR, BYVAL blendPositions AS SINGLE PTR, _
   BYVAL count AS LONG) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *blendFactors* | Pointer to an array of blend factors. Each number in the array should be in the range 0 through 1. |
| *blendPositions* | Pointer to an array of blend positions. Each number in the array should be in the range 0 through 1. |
| *count* | Optional. Integer that specifies the number of elements in the blendFactors array. This is the same as the number of elements in the *blendPositions* array. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A **PathGradientBrush** object has a boundary path and a center point. When you fill an area with a path gradient brush, the color changes gradually as you move from the boundary path to the center point. By default, the color is linearly related to the distance, but you can customize the relationship between color and distance by calling the **SetBlend** method.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on an ellipse. The code
' calls the PathGradientBrush::SetBlend method of the PathGradientBrush object to establish
' a set of blend factors and blend positions for the brush. Then the code uses the path
' gradient brush to fill the ellipse.
' ========================================================================================
SUB Example_SetBlend (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a path that consists of a single ellipse.
   DIM path AS CGpGraphicsPath
   path.AddEllipse(0, 0, 200, 100)

   ' // Use the path to construct a brush.
   DIM pthGrBrush AS CGpPathGradientBrush = @path

   ' // Set the color at the center of the path to blue.
   pthGrBrush.SetCenterColor(GDIP_ARGB(255, 0, 0, 255))

   ' // Set the color along the entire boundary of the path to aqua.
   DIM colors(0) AS ARGB = {GDIP_ARGB(255, 0, 255, 255)}
   DIM count AS LONG = 1
   pthGrBrush.SetSurroundColors(@colors(0), @count)

   ' // Set blend factors and positions for the path gradient brush.
   DIM factors(0 TO 3) AS SINGLE = {0.0, 0.4, 0.8, 1.0}
   DIM positions(0 TO 3) AS SINGLE = {0.0, 0.3, 0.7, 1.0}

   pthGrBrush.SetBlend(@factors(0), @positions(0), 4)

   ' // Fill the ellipse with the path gradient brush.
   graphics.FillEllipse(@pthGrBrush, 0, 0, 200, 100)

END SUB
' ========================================================================================
```

# <a name="SetBlendBellShapePGBrush"></a>SetBlendBellShape (CGpPathGradientBrush)

Sets the blend shape of this path gradient brush.

```
FUNCTION SetBlendBellShape (BYVAL focus AS SINGLE, BYVAL scale AS SINGLE = 1.0) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *focus* | Single precision number that specifies where the center color will be at its highest intensity. This number must be in the range 0 through 1. |
| *scale* | Single precision number that specifies the maximum intensity of center color that gets blended with the boundary color. This number must be in the range 0 through 1. The default value is 1. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

By default, as you move from the boundary of a path gradient to the center point, the color changes gradually from the boundary color to the center color. You can customize the positioning and blending of the boundary and center colors by calling the **SetBlendBellShape** method.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on an ellipse. The code
' calls the PathGradientBrush.SetBlendBellShape method of the PathGradientBrush object,
' passing a focus of 0.2 and a scale of 0.7. Then the code uses the path gradient brush to
' paint a rectangle that contains the ellipse.
' ========================================================================================
SUB Example_SetBlendBellShape (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a path that consists of a single ellipse.
   DIM path AS CGpGraphicsPath
   path.AddEllipse(0, 0, 200, 100)

   ' // Use the path to construct a brush.
   DIM pthGrBrush AS CGpPathGradientBrush = @path

   ' // Set the color at the center of the path to blue.
   pthGrBrush.SetCenterColor(GDIP_ARGB(255, 0, 0, 255))

   ' // Set the color along the entire boundary of the path to aqua.
   DIM colors(0) AS ARGB = {GDIP_ARGB(255, 0, 255, 255)}
   DIM count AS LONG = 1
   pthGrBrush.SetSurroundColors(@colors(0), @count)

   pthGrBrush.SetBlendBellShape(0.2, 0.7)

   ' // The color is blue on the boundary and at the center.
   ' // At points that are 20 percent of the way from the boundary to the
   ' // center, the color is 70 percent red and 30 percent blue.

   graphics.FillRectangle(@pthGrBrush, 0, 0, 300, 300)

END SUB
' ========================================================================================
```

# <a name="SetBlendTriangularShapePGBrush"></a>SetBlendTriangularShape (CGpPathGradientBrush)

Sets the blend shape of this path gradient brush.

```
FUNCTION SetBlendTriangularShape (BYVAL focus AS SINGLE, BYVAL scale AS SINGLE = 1.0) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *focus* | Single precision number that specifies where the center color will be at its highest intensity. This number must be in the range 0 through 1. |
| *scale* | Single precision number that specifies the maximum intensity of center color that gets blended with the boundary color. This number must be in the range 0 through 1. The default value is 1. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

By default, as you move from the boundary of a path gradient to the center point, the color changes gradually from the boundary color to the center color. You can customize the positioning and blending of the boundary and center colors by calling the **SetBlendTriangularShape** method.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on an ellipse. The code
' calls the PathGradientBrush.SetBlendTriangularShape method of the PathGradientBrush object,
' passing a focus of 0.2 and a scale of 0.7. Then the code uses the path gradient brush to
' paint a rectangle that contains the ellipse.
' ========================================================================================
SUB Example_SetBlendTriangularShape (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a path that consists of a single ellipse.
   DIM path AS CGpGraphicsPath
   path.AddEllipse(0, 0, 200, 100)

   ' // Use the path to construct a brush.
   DIM pthGrBrush AS CGpPathGradientBrush = @path

   ' // Set the color at the center of the path to blue.
   pthGrBrush.SetCenterColor(GDIP_ARGB(255, 0, 0, 255))

   ' // Set the color along the entire boundary of the path to aqua.
   DIM colors(0) AS ARGB = {GDIP_ARGB(255, 0, 255, 255)}
   DIM count AS LONG = 1
   pthGrBrush.SetSurroundColors(@colors(0), @count)

   pthGrBrush.SetBlendTriangularShape(0.2, 0.7)

   ' // The color is blue on the boundary and at the center.
   ' // At points that are 20 percent of the way from the boundary to the
   ' // center, the color is 70 percent red and 30 percent blue.

   graphics.FillRectangle(@pthGrBrush, 0, 0, 300, 300)

END SUB
' ========================================================================================
```

# <a name="SetCenterColor"></a>SetCenterColor (CGpPathGradientBrush)

Sets the center color of this path gradient brush. The center color is the color that appears at the brush's center point.

```
FUNCTION SetCenterColor (BYVAL colour AS ARGB) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *colour* | ARGB color that specifies the center color. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

By default the center point is the centroid of the brush's boundary path, but you can set the center point to any location inside or outside the path.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on an ellipse. The code
' calls the PathGradientBrush.SetCenterColor method of the PathGradientBrush object to set
' the center color to blue. The PathGradientBrush.SetSurroundColors method sets the color
' along the entire boundary to aqua. The FillRectangle Methods method uses the path gradient
' brush to paint a rectangle that contains the ellipse.
' ========================================================================================
SUB Example_SetCenterColor (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a path that consists of a single ellipse.
   DIM path AS CGpGraphicsPath
   path.AddEllipse(0, 0, 200, 100)

   ' // Use the path to construct a brush.
   DIM pthGrBrush AS CGpPathGradientBrush = @path

   ' // Set the color at the center of the path to blue.
   pthGrBrush.SetCenterColor(GDIP_ARGB(255, 0, 0, 255))

   ' // Set the color along the entire boundary of the path to aqua.
   DIM colors(0) AS ARGB = {GDIP_ARGB(255, 0, 255, 255)}
   DIM count AS LONG = 1
   pthGrBrush.SetSurroundColors(@colors(0), @count)

   ' // Fill the ellipse with the path gradient brush.
   graphics.FillEllipse(@pthGrBrush, 0, 0, 200, 100)

END SUB
' ========================================================================================
```

# <a name="SetCenterPoint"></a>SetCenterPoint (CGpPathGradientBrush)

Sets the center point of this path gradient brush. By default, the center point is at the centroid of the brush's boundary path, but you can set the center point to any location inside or outside the path.

```
FUNCTION SetCenterPoint (BYVAL pt AS PointF PTR) AS GpStatus
FUNCTION SetCenterPoint (BYVAL pt AS Point PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pt* | Reference to a **PointF** or **Point** structure that specifies the center point. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on an ellipse. The code
' sets the center color to blue and sets the color along the boundary to aqua. By default,
' the center point would be at the center of the ellipse (100, 50), but the call to the
' PathGradientBrush.SetCenterPoint method sets the center point to (180.5, 50.0).
' ========================================================================================
SUB Example_SetCenterPoint (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a path that consists of a single ellipse.
   DIM path AS CGpGraphicsPath
   path.AddEllipse(0, 0, 200, 100)

   ' // Use the path to construct a brush.
   DIM pthGrBrush AS CGpPathGradientBrush = @path

   ' // Set the color at the center of the path to blue.
   pthGrBrush.SetCenterColor(GDIP_ARGB(255, 0, 0, 255))

   ' // Set the center point.
   DIM pt AS GpPointF = GDIP_POINTF(180.5, 50.0)
   pthGrBrush.SetCenterPoint(@pt)

   ' // Set the color along the entire boundary of the path to aqua.
   DIM colors(0) AS ARGB = {GDIP_ARGB(255, 0, 255, 255)}
   DIM count AS LONG = 1
   pthGrBrush.SetSurroundColors(@colors(0), @count)

   graphics.FillRectangle(@pthGrBrush, 0, 0, 300, 300)

END SUB
' ========================================================================================
```

# <a name="SetFocusScales"></a>SetFocusScales (CGpPathGradientBrush)

Sets the focus scales of this path gradient brush.

```
FUNCTION SetFocusScales (BYVAL xScale AS SINGLE, BYVAL yScale AS SINGLE) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *xScale* | Single precision number that specifies the x focus scale. |
| *yScale* | Single precision number that specifies the y focus scale. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

By default, the center color of a path gradient is at the center point. By calling **SetFocusScales**, you can specify that the center color should appear along a path that surrounds the center point. That path is the boundary path scaled by a factor of *xScale* in the x direction and by a factor of *yScale* in the y direction. The area inside the scaled path is filled with the center color.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on a triangular path. The
' code calls the PathGradientBrush::SetFocusScales method of the PathGradientBrush object
' to set the brush's focus scales to (0.2, 0.2). Then the code uses the path gradient brush
' to paint a rectangle that includes the triangular path.
' ========================================================================================
SUB Example_SetFocusScales (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM points(0 TO 2) AS GpPoint = {GDIP_POINT(100, 0), GDIP_POINT(200, 200), GDIP_POINT(0, 200)}

   ' // No GraphicsPath object is created. The PathGradientBrush
   ' // object is constructed directly from the array of points.
   DIM pthGrBrush AS CGpPathGradientBrush = CGpPathGradientBrush(@points(0), 3)

   DIM colors(0 TO 1) AS ARGB = {GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255)}

   ' // red at the boundary of the outer triangle
   ' // blue at the boundary of the inner triangle
   DIM relativePositions(0 TO 1) AS SINGLE = {0.0, 1.0}
   pthGrBrush.SetInterpolationColors(@colors(0), @relativePositions(0), 2)

   ' // The inner triangle is formed by scaling the outer triangle
   ' // about its centroid. The scaling factor is 0.2 in both the x and y directions.
   pthGrBrush.SetFocusScales(0.2, 0.2)

   ' // Fill a rectangle that is larger than the triangle
   ' // specified in the Point array. The portion of the
   ' // rectangle outside the triangle will not be painted.
   graphics.FillRectangle(@pthGrBrush, 0, 0, 200, 200)

END SUB
' ========================================================================================
```

# <a name="SetInterpolationColorsPGBrush"></a>SetInterpolationColors (CGpPathGradientBrush)

Sets the preset colors and the blend positions of this path gradient brush.

```
FUNCTION SetInterpolationColors (BYVAL presetColors AS ARGB PTR, _
   BYVAL blendPositions AS SINGLE PTR, BYVAL count AS LONG) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *presetColors* | Pointer to an array of colors that specifies the interpolation colors for the gradient. A color of a given index in the presetColors array corresponds to the blend position of that same index in the *blendPositions* array. |
| *blendPositions* | Pointer to an array that specifies the blend positions. Each blend position is a number from 0 through 1, where 0 indicates the boundary of the gradient and 1 indicates the center point. A blend position between 0 and 1 specifies the set of all points that are a certain fraction of the distance from the boundary to the center point. For example, a blend position of 0.7 specifies the set of all points that are 70 percent of the way from the boundary to the center point. |
| *count* | Integer that specifies the number of elements in the *presetColors* array. This is the same as the number of elements in the blendPositions array. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

A simple path gradient brush has two colors: a boundary color and a center color. When you paint with such a brush, the color changes gradually from the boundary color to the center color as you move from the boundary path to the center point. You can create a more complex gradient by specifying an array of preset colors and an array of blend positions.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on a triangular path. The
' PathGradientBrush.SetInterpolationColors method sets the brush's preset colors to red,
' blue, and aqua and sets the blend positions to 0, 0, 4, and 1. The Graphics.FillRectangle
' method uses the path gradient brush to paint a rectangle that contains the triangular path.
' ========================================================================================
SUB Example_SetInterpolationColors (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM points(0 TO 2) AS GpPoint = {GDIP_POINT(100, 0), GDIP_POINT(200, 200), GDIP_POINT(0, 200)}
   DIM pthGrBrush AS CGpPathGradientBrush = CGpPathGradientBrush(@points(0), 3)

   DIM colors(0 TO 2) AS ARGB = {GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255), GDIP_ARGB(255, 0, 255, 255)}
   DIM positions(0 TO 2) AS SINGLE = {0.0, 0.4, 1.0}

   pthGrBrush.SetInterpolationColors(@colors(0), @positions(0), 3)
   graphics.FillRectangle(@pthGrBrush, 0, 0, 300, 300)

END SUB
' ========================================================================================
```

# <a name="SetSurroundColors"></a>SetSurroundColors (CGpPathGradientBrush)

Sets the surround colors of this path gradient brush. The surround colors are colors specified for discrete points on the brush's boundary path.

```
FUNCTION SetSurroundColors (BYVAL colors AS ARGB PTR, BYVAL count AS INT_ PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *colors* | Pointer to an array of colors that specify the surround colors.  |
| *count* | Out. Pointer to an integer that, on input, specifies the number of colors in the colors array. If the method succeeds, this parameter, on output, receives the number of surround colors set. If the method fails, this parameter does not receive a value. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A path gradient brush has a boundary path and a center point. The center point is set to a single color, but you can specify different colors for several points on the boundary. For example, suppose you specify red for the center color, and you specify blue, green, and yellow for distinct points on the boundary. Then as you move along the boundary, the color will change gradually from blue to green to yellow and back to blue. As you move along a straight line from any point on the boundary to the center point, the color will change from that boundary point's color to red.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on an array of three points
' that defines a triangular path. The code also initializes an array of three Color objects.
' The call to the PathGradientBrush::SetSurroundColors method associates each color in the
' color array with the corresponding (same index) point in the point array. After the surround
' colors of the path gradient brush have been set, the Graphics.FillRectangle method uses
' the path gradient brush to paint a rectangle that includes the triangular path.
' One edge of the rendered triangle changes gradually from red to green. The next edge
' changes gradually from green to black, and the third edge changes gradually from black
' to red. The code does not set the center color, so the center color has the default value
' of black. As you move along a straight line from any point on the boundary path (triangle)
' to the center point, the color changes gradually from that boundary point's color to black.
' ========================================================================================
SUB Example_SetSurroundColors (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM points(0 TO 2) AS GpPoint = {GDIP_POINT(20, 20), GDIP_POINT(100, 20), GDIP_POINT(100, 100)}
   DIM pthGrBrush AS CGpPathGradientBrush = CGpPathGradientBrush(@points(0), 3)

   DIM nCount AS LONG = 3
   DIM colors(0 TO 2) AS ARGB = {GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255), GDIP_ARGB(255, 0, 255, 255)}
   pthGrBrush.SetSurroundColors(@colors(0), @nCount)

   graphics.FillRectangle(@pthGrBrush, 0, 0, 200, 200)

END SUB
' ========================================================================================
```

# <a name="SetTransformPGBrush"></a>SetTransform (CGpPathGradientBrush)

Sets the transformation matrix of this path gradient brush.

```
FUNCTION SetTransform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMatrix* | Pointer to a **Matrix** object that specifies the transformation matrix to use. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A **PathGradientBrush** object has a **GraphicsPath** object that serves as the boundary path for the brush. When you paint with a path gradient brush, only the area inside the boundary path is filled. If the brush's transformation matrix is set to represent any transformation other than the identity, then the boundary path is transformed according to that matrix during rendering, and only the area inside the transformed path is filled.

The transformation applies only during rendering. The boundary path stored by the **PathGradientBrush** object is not altered by the **SetTransform** method.

Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on a triangular path. The
' Graphics.FillRectangle method uses the path gradient brush to paint a rectangle that
' contains the triangular path. Next, the code creates a Matrix object that represents a
' composite transformation (rotate, then translate) and passes the address of that Matrix
' object to the PathGradientBrush.SetTransform method of the PathGradientBrush object. The
' code calls FillRectangle a second time to paint the same rectangle using the transformed
' path gradient brush.
' ========================================================================================
SUB Example_SetTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM points(0 TO 2) AS GpPoint = {GDIP_POINT(0, 0), GDIP_POINT(100, 0), GDIP_POINT(100, 100)}
   DIM pthGrBrush AS CGpPathGradientBrush = CGpPathGradientBrush(@points(0), 3)

   DIM nCount AS LONG = 3
   DIM colors(0 TO 2) AS ARGB = {GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 255, 0), GDIP_ARGB(255, 0, 0, 0)}
   pthGrBrush.SetSurroundColors(@colors(0), @nCount)

   graphics.FillRectangle(@pthGrBrush, 0, 0, 200, 200)

   ' // Set the transformation for the brush (rotate, then translate).
   DIM matrix AS CGpMatrix = CGpMatrix(0.0, 1.0, -1.0, 0.0, 150.0, 60.0)
   pthGrBrush.SetTransform(@matrix)
   
   ' // Fill the same area with the transformed path gradient brush.
   graphics.FillRectangle(@pthGrBrush, 0, 0, 200, 200)

END SUB
' ========================================================================================
```

# <a name="SetWrapModePGBrush"></a>SetWrapMode (CGpPathGradientBrush)

Sets the wrap mode of this path gradient brush.

```
FUNCTION SetWrapMode (BYVAL wrapMode AS WrapMode) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *wrapMode* | Element of the **WrapMode** enumeration that specifies how areas painted with the path gradient brush will be tiled. The default value is **WrapModeClamp**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

The bounding rectangle of a path gradient brush is the smallest rectangle that encloses the brush's boundary path. When you paint the bounding rectangle with the path gradient brush, only the area inside the boundary path gets filled. The area inside the bounding rectangle but outside the boundary path does not get filled.

**WrapModeClamp** (the default wrap mode) indicates that no painting occurs outside of the brush's bounding rectangle. All of the other wrap modes indicate that areas outside the brush's bounding rectangle will be tiled. Each tile is a copy (possibly flipped) of the filled path inside its bounding rectangle.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on a triangular path. The
' code calls the PathGradientBrush::SetWrapMode method of the PathGradientBrush object to
' set the brush's wrap mode to WrapModeTileFlipX. The Graphics::FillRectangle method uses
' the path gradient brush to tile a large area.
' The output of the code is a grid of tiles. As you move from one tile to the next in a
' given row, the image (filled boundary path inside the bounding rectangle) is flipped
' horizontally.
' ========================================================================================
SUB Example_SetWrapMode (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM points(0 TO 2) AS GpPoint = {GDIP_POINT(0, 0), GDIP_POINT(100, 0), GDIP_POINT(100, 100)}
   DIM colors(0 TO 2) AS ARGB = {GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255), GDIP_ARGB(255, 0, 255, 0)}

   DIM nCount AS LONG = 3
   DIM pthGrBrush AS CGpPathGradientBrush = CGpPathGradientBrush(@points(0), 3)
   pthGrBrush.SetSurroundColors(@colors(0), @nCount)
   pthGrBrush.SetWrapMode(WrapModeTileFlipX)

   graphics.FillRectangle(@pthGrBrush, 0, 0, 800, 800)

END SUB
' ========================================================================================
```

# <a name="TranslateTransformPGBrush"></a>TranslateTransform (CGpPathGradientBrush)

Updates this brush's current transformation matrix with the product of itself and a translation matrix.

```
FUNCTION TranslateTransform (BYVAL dx AS SINGLE, BYVAL dy AS SINGLE, _
   BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *dx* | Single precision number that specifies the horizontal component of the translation. |
| *dy* | Single precision number that specifies the vertical component of the translation. |
| *order* | Optional. Element of the **MatrixOrder** enumeration that specifies the order of multiplication. **MatrixOrderPrepend** specifies that the passed matrix is on the left, and **MatrixOrderAppend** specifies that the passed matrix is on the right. The default value is **MatrixOrderPrepend**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A single 3×3 matrix can store any sequence of affine transformations. If you have several 3×3 matrices, each of which represents an affine transformation, the product of those matrices is a single 3×3 matrix that represents the entire sequence of transformations. The transformation represented by that product is called a composite transformation. For example, suppose matrix *S* represents a scaling and matrix *T* represents a translation. If matrix *M* is the product *ST*, then matrix *M* represents a composite transformation: first scale, then translate.

#### Example

````
' ========================================================================================
' The following example creates a PathGradientBrush object based on a triangular path. The
' calls to the PathGradientBrush.ScaleTransform and PathGradientBrush.TranslateTransform
' methods of the PathGradientBrush object set the elements of the brush's transformation
' matrix so that it represents a composite transformation: first scale, then translate.
' The code uses the path gradient brush twice to paint a rectangle: once before the
' transformation is set and once after the transformation is set.
' ========================================================================================
SUB Example_TranslateTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM points(0 TO 2) AS GpPoint = {GDIP_POINT(0, 0), GDIP_POINT(50, 0), GDIP_POINT(50, 50)}
   DIM pthGrBrush AS CGpPathGradientBrush = CGpPathGradientBrush(@points(0), 3)

   ' // Fill an area with the path gradient brush (no transformation).
   graphics.FillRectangle(@pthGrBrush, 0, 0, 500, 500)

   pthGrBrush.ScaleTransform(3.0, 1.0)
   pthGrBrush.TranslateTransform(100.0, 50.0, MatrixOrderAppend)

   ' // Fill the same area with the transformed path gradient brush.
   graphics.FillRectangle(@pthGrBrush, 0, 0, 500, 500)

END SUB
' ========================================================================================
````

# <a name="ConstructorsTBrush"></a>Constructors (CGpTextureBrush)

Creates a texture brush.

```
CONSTRUCTOR CGpTextureBrush (BYVAL pTextureBrush AS CGpTextureBrush PTR)
CONSTRUCTOR CGpTextureBrush (BYVAL pImage AS CGpImage PTR, BYVAL nWrapMode AS WrapMode = WrapModeTile)
CONSTRUCTOR CGpTextureBrush (BYVAL pImage AS CGpImage PTR, BYVAL dstRect AS GpRectF PTR, _
   BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL)
CONSTRUCTOR CGpTextureBrush (BYVAL pImage AS CGpImage PTR, BYVAL dstRect AS GpRect PTR, _
   BYVAL pImageAttributes AS CGpImageAttributes PTR = NULL)
CONSTRUCTOR CGpTextureBrush (BYVAL pImage AS CGpImage PTR, BYVAL nWrapMode AS WrapMode, _
   BYVAL dstRect AS GpRect PTR)
CONSTRUCTOR CGpTextureBrush (BYVAL pImage AS CGpImage PTR, BYVAL nWrapMode AS WrapMode, _
   BYVAL dstX AS SINGLE, BYVAL dstY AS SINGLE, BYVAL dstWidth AS SINGLE, BYVAL dstHeight AS SINGLE)
CONSTRUCTOR CGpTextureBrush (BYVAL pImage AS CGpImage PTR, BYVAL nWrapMode AS WrapMode, _
   BYVAL dstX AS INT_, BYVAL dstY AS INT_, BYVAL dstWidth AS INT_, BYVAL dstHeight AS INT_)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pImage* | Pointer to a variable that will receive a pointer to the Image object that is defined by this texture brush. |
| *nWrapMode* | Element of the **WrapMode** enumeration that specifies how repeated copies of an image are used to tile an area when it is painted with this texture brush.  |
| *dstX* | Leftmost coordinate of the image portion to be used by this brush. |
| *dstY* | Uppermost coordinate of the image portion to be used by this brush. |
| *dstWidth* | Width of the brush and width of the image portion to be used by the brush. |
| *dstHeight* | Height of the brush and height of the image portion to be used by the brush. |
| *dstHeight* | Height of the brush and height of the image portion to be used by the brush. |

#### Remarks

The *dstX*, *dstY*, *dstWidth*, and *dstHeight* parameters specify a rectangle. The size of the brush is defined by dstWidth and dstHeight. The *dstX* and *dstY* parameters have no affect on the brush's size or position — the brush is always oriented at (0, 0). The *dstX*, *dstY*, *dstWidth*, and *dstHeight* parameters define the portion of the image to be used by the brush.

For example, suppose you have an image that is stored in an **Image** object and is 256 ×512 (width ×height) pixels. Then you create a **TextureBrush** object based on this image as follows:

```
TextureBrush(@someImage, WrapModeTile, 12, 50, 100, 150)
```

The brush will have a width of 100 units and a height of 150 units. The brush will use a rectangular portion of the image. This portion begins at the pixel having coordinates (12, 50). The width and height of the portion are 100 and 150, respectively, measured from the starting pixel.

Now suppose you create another TextureBrush object based on the same image and specify a different rectangle:

```
TextureBrush(@someImage, WrapModeTile, 0, 0, 256, 512)
```

The brush will have width and height equal to 256 and 512, respectively. The brush will use the entire image instead of a portion of it because the rectangle specifies a starting pixel at coordinates (0, 0) and dimensions identical to those of the image.

# <a name="GetImage"></a>GetImage (CGpTextureBrush)

Gets a pointer to the Image object that is defined by this texture brush.

```
FUNCTION GetImage (BYVAL pImage AS CGpImage PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pImage* | Pointer to a variable that will receive a pointer to the Image object that is defined by this texture brush. |

#### Return value

This method returns the number of preset colors currently specified for this path gradient brush.

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a texture brush and uses it to fill an ellipse. The code
' then gets the brush's image and draws it.
' ========================================================================================
SUB Example_GetImage (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a texture brush, and use it to fill an ellipse.
   DIM pImage AS CGpImage = "HouseAndTree.gif"
   DIM textureBrush AS CGpTextureBrush = @pImage
   graphics.FillEllipse(@textureBrush, 0, 0, 200, 200)

   ' // Get the brush's image, and draw that image
   DIM pImage2 AS CGpImage
   textureBrush.GetImage(@pImage2)
   graphics.DrawImage(@pImage2, 10, 150)

END SUB
' ========================================================================================
```

# <a name="GetTransformTBrush"></a>GetTransform (CGpTextureBrush)

Gets the transformation matrix of this texture brush.

```
FUNCTION GetTransform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMatrix* | Pointer to a **Matrix** object that receives the transformation matrix. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A **TextureBrush** object maintains a transformation matrix that can store any affine transformation. When you use a texture brush to fill an area, GDI+ transforms the brush's image according to the brush's transformation matrix and then fills the area. The transformed image exists only during rendering; the image stored in the **TextureBrush** object is not transformed. For example, suppose you call *someTextureBrush.ScaleTransform(3)* and then paint an area with *someTextureBrush*. The width of the brush's image triples when the area is painted, but the image stored in *someTextureBrush* remains unchanged.

#### Example

```
' ========================================================================================
' The following example creates a texture brush and sets the transformation of the brush.
' The code then gets the brush's transformation matrix and proceeds to inspect or use the elements.
' ========================================================================================
SUB Example_GetTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a texture brush, and set its transformation.
   DIM pImage AS CGpImage = "HouseAndTree.gif"
   DIM textureBrush AS CGpTextureBrush = @pImage
   textureBrush.ScaleTransform(3, 2)

   DIM matrix AS CGpMatrix
   DIM elements(5) AS SINGLE

   textureBrush.GetTransform(@matrix)
   matrix.GetElements(@elements(0))

   FOR j AS LONG = 0 TO 5
      ' // Inspect or use the value in elements[j].
      PRINT elements(j)
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetWrapModeTBrush"></a>GetWrapMode (CGpTexturetBrush)

Gets the wrap mode currently set for this texture brush.

```
FUNCTION GetWrapMode () AS WrapMode
```

#### Return value

This method returns the wrap mode currently set for this texture brush. The value returned is one of the elements of the **WrapMode** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a texture brush, sets the wrap mode of the brush, and uses
' the brush to fill a rectangle. Next, the code gets the wrap mode of the brush and stores
' the value. The code creates a second texture brush and uses the stored wrap mode to set
' the wrap mode of the second texture brush. Then, the code uses the second texture brush
' to fill a rectangle.
' ========================================================================================
SUB Example_WrapMode (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a texture brush, and set its transformation.
   DIM pImage AS CGpImage = "HouseAndTree.gif"
   DIM textureBrush AS CGpTextureBrush = @pImage
   textureBrush.SetWrapMode(WrapModeTileFlipX)
   graphics.FillRectangle(@textureBrush, 0, 0, 200, 200)

   ' // Get the brush's wrap mode.
   DIM nWrapMode AS WrapMode = textureBrush.GetWrapMode

   ' // Create a second texture brush with the same wrap mode.
   DIM pImage2 AS CGpImage = "MyTexture.png"
   DIM textureBrush2 AS CGpTextureBrush = @pImage2
   textureBrush2.SetWrapMode(nWrapMode)
   graphics.FillRectangle(@textureBrush2, 210, 0, 200, 200)

END SUB
' ========================================================================================
```

# <a name="MultiplyTransformTBrush"></a>MultiplyTransform (CGpTextureBrush)

Updates the brush's transformation matrix with the product of itself and another matrix.

```
FUNCTION MultiplyTransform (BYVAL pMatrix AS CGpMatrix PTR, _
   BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMatrix* | Pointer to a matrix to be multiplied by the brush's current transformation matrix. |
| *order* | Optional. Element of the **MatrixOrder** enumeration that specifies the order of multiplication. **MatrixOrderPrepend** specifies that the passed matrix is on the left, and **MatrixOrderAppend** specifies that the passed matrix is on the right. The default value is **MatrixOrderPrepend**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a texture brush and sets the transformation of the brush.
' The code then uses the transformed brush to fill a rectangle.
' ========================================================================================
SUB Example_MultiplyTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Horizontal stretch
   DIM matrix AS CGpMatrix = CGpMatrix(3, 0, 0, 1, 0, 0)

   ' // Create a texture brush, and set its transformation.
   DIM pImage AS CGpImage = "HouseAndTree.gif"
   DIM textureBrush AS CGpTextureBrush = @pImage

   textureBrush.RotateTransform(30)   ' // rotate
   textureBrush.MultiplyTransform(@matrix, MatrixOrderAppend)   ' // stretch
   graphics.FillRectangle(@textureBrush, 0, 0, 400, 200)

END SUB
' ========================================================================================
```

# <a name="ResetTransformTBrush"></a>ResetTransform (CGpTextureBrush)

Resets the transformation matrix of this texture brush to the identity matrix. This means that no transformation takes place.

```
FUNCTION ResetTransform () AS GpStatus
```

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a texture brush and sets the transformation of the brush.
' Next, the code uses the transformed brush to fill a rectangle. Then, the code resets the
' transformation of the brush and uses the untransformed brush to fill a rectangle.
' ========================================================================================
SUB Example_ResetTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a texture brush, and set its transformation.
   DIM pImage AS CGpImage = "HouseAndTree.gif"
   DIM textureBrush AS CGpTextureBrush = @pImage
   textureBrush.RotateTransform(30)

   ' // Fill a rectangle with the transformed texture brush.
   graphics.FillEllipse(@textureBrush, 0, 0, 200, 100)

   ' // Reset the transformation
   textureBrush.ResetTransform

   ' // Fill a rectangle with the texture brush (no transformation).
   graphics.FillEllipse(@textureBrush, 210, 0, 200, 100)

END SUB
' ========================================================================================
```

# <a name="RotateTransformTBrush"></a>RotateTransform (CGpTextureBrush)

Updates this texture brush's current transformation matrix with the product of itself and a rotation matrix.

```
FUNCTION RotateTransform (BYVAL angle AS SINGLE, _
   BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *angle* | Single precision number that specifies the angle of rotation in degrees. |
| *order* | Optional. Element of the MatrixOrder enumeration that specifies the order of multiplication. **MatrixOrderPrepend** specifies that the passed matrix is on the left, and **MatrixOrderAppend** specifies that the passed matrix is on the right. The default value is **MatrixOrderPrepend**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A single 3×3 matrix can store any sequence of affine transformations. If you have several 3×3 matrices, each of which represents an affine transformation, the product of those matrices is a single 3×3 matrix that represents the entire sequence of transformations. The transformation represented by that product is called a composite transformation. For example, suppose matrix R represents a rotation, and matrix T represents a translation. If matrix M is the product RT, then matrix M represents a composite transformation: first rotate, then translate.

The order of matrix multiplication is important. In general, the matrix product RT is not the same as the matrix product TR. In the example given in the previous paragraph, the composite transformation represented by RT (first rotate, then translate) is not the same as the composite transformation represented by TR (first translate, then rotate).

#### Example

```
' ========================================================================================
' The following example creates a texture brush and sets the transformation of the brush.
' The code then uses the transformed brush to fill a rectangle.
' ========================================================================================
SUB Example_RotateTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a texture brush, and set its transformation.
   DIM pImage AS CGpImage = "HouseAndTree.gif"
   DIM textureBrush AS CGpTextureBrush = @pImage
   textureBrush.ScaleTransform(3, 1)
   textureBrush.RotateTransform(30, MatrixOrderAppend)
   graphics.FillEllipse(@textureBrush, 0, 0, 400, 200)

END SUB
' ========================================================================================
```

# <a name="ScaleTransformTBrush"></a>ScaleTransform (CGpTextureBrush)

Updates this texture brush's current transformation matrix with the product of itself and a scaling matrix.

```
FUNCTION ScaleTransform (BYVAL sx AS SINGLE, BYVAL sy AS SINGLE, _
   BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *sx* | Single precision number that specifies the amount to scale in the x direction. |
| *sy* | Single precision number that specifies the amount to scale in the y direction. |
| *order* | Optional. Element of the MatrixOrder enumeration that specifies the order of multiplication. **MatrixOrderPrepend** specifies that the passed matrix is on the left, and **MatrixOrderAppend** specifies that the passed matrix is on the right. The default value is **MatrixOrderPrepend**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A single 3×3 matrix can store any sequence of affine transformations. If you have several 3×3 matrices, each of which represents an affine transformation, the product of those matrices is a single 3×3 matrix that represents the entire sequence of transformations. The transformation represented by that product is called a composite transformation. For example, suppose matrix T represents a translation, and matrix S represents a scaling. If matrix M is the product TS, then matrix M represents a composite transformation: first translate, then scale.

The order of matrix multiplication is important. In general, the matrix product RT is not the same as the matrix product TR. In the example given in the previous paragraph, the composite transformation represented by RT (first rotate, then translate) is not the same as the composite transformation represented by TR (first translate, then rotate).

#### Example

```
' ========================================================================================
' The following example creates a texture brush and sets the transformation of the brush.
' The code then uses the transformed brush to fill a rectangle.
' ========================================================================================
SUB Example_ScaleTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a texture brush, and set its transformation.
   DIM pImage AS CGpImage = "HouseAndTree.gif"
   DIM textureBrush AS CGpTextureBrush = @pImage
   textureBrush.RotateTransform(30)
   textureBrush.ScaleTransform(3, 1, MatrixOrderAppend)
   graphics.FillEllipse(@textureBrush, 0, 0, 400, 200)

END SUB
' ========================================================================================
```

# <a name="SetTransformTBrush"></a>SetTransform (CGpTextureBrush)

Sets the transformation matrix of this texture brush.

```
FUNCTION SetTransform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMatrix* | Pointer to a **Matrix** object that specifies the transformation matrix to use. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A **TextureBrush** object maintains a transformation matrix that can store any affine transformation. When you use a texture brush to fill an area, Windows GDI+ transforms the brush's image according to the brush's transformation matrix and then fills the area. The transformed image exists only during rendering; the image stored in the TextureBrush object is not transformed. For example, suppose you call and then paint an area with *someTextureBrush.ScaleTransform(3)* and then paint an area with *someTextureBrush*. The width of the brush's image triples when the area is painted, but the image stored in someTextureBrush remains unchanged.

#### Example

```
' ========================================================================================
' The following example creates a texture brush and sets the transformation of the brush.
' The code then uses the transformed brush to fill an ellipse.
' ========================================================================================
SUB Example_SetTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Horizontal stretch
   DIM matrix AS CGpMatrix = CGpMatrix(2, 0, 0, 1, 0, 0)

   ' // Create a texture brush, and set its transformation.
   DIM pImage AS CGpImage = "HouseAndTree.gif"
   DIM textureBrush AS CGpTextureBrush = @pImage
   textureBrush.SetTransform(@matrix)
   graphics.FillEllipse(@textureBrush, 0, 0, 400, 200)

END SUB
' ========================================================================================
```

# <a name="SetWrapModeTBrush"></a>SetWrapMode (CGpTextureBrush)

Sets the wrap mode of this texture brush.

```
FUNCTION SetWrapMode (BYVAL wrapMode AS WrapMode) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *wrapMode* | Element of the **WrapMode** enumeration that specifies how repeated copies of an image are used to tile an area when it is painted with this texture brush. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

An area that extends beyond the boundaries of the brush is tiled with repeated copies of the brush. A texture brush may have alternate tiles flipped in a certain direction, as specified by the wrap mode. Flipping has the effect of reversing the brush's image. For example, if the wrap mode is specified as **WrapModeTileFlipX**, the brush is flipped about a line that is parallel to the y-axis.

The texture brush is always oriented at (0, 0). If the wrap mode is specified as **WrapModeClamp**, no area outside of the brush is tiled. For example, suppose you create a texture brush, specifying **WrapModeClamp** as the wrap mode:

```
TextureBrush(SomeImage, WrapModeClamp)
```

Then you paint an area with the brush. If the size of the brush has a height of 50 and the painted area is a rectangle with its upper-left corner at (0, 50), you will see no repeated copies of the brush (no tiling).

The default wrap mode for a texture brush is **WrapModeTile**, which specifies no flipping of the tile and no clamping.

#### Example

```
' ========================================================================================
' The following example creates a texture brush, sets the wrap mode of the brush, and uses
' the brush to fill a rectangle.
' ========================================================================================
SUB Example_SetWrapMode (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a texture brush, and set its transformation.
   DIM pImage AS CGpImage = "HouseAndTree.gif"
   DIM textureBrush AS CGpTextureBrush = @pImage
   textureBrush.SetWrapMode(WrapModeTileFlipX)
   graphics.FillRectangle(@textureBrush, 0, 0, 400, 200)

END SUB
' ========================================================================================
```

# <a name="TranslateTransformTBrush"></a>TranslateTransform (CGpTextureBrush)

Updates this brush's current transformation matrix with the product of itself and a translation matrix.

```
FUNCTION TranslateTransform (BYVAL dx AS SINGLE, BYVAL dy AS SINGLE, _
   BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *dx* | Single precision number that specifies the horizontal component of the translation. |
| *dy* | Single precision number that specifies the vertical component of the translation. |
| *order* | Optional. Element of the **MatrixOrder** enumeration that specifies the order of multiplication. **MatrixOrderPrepend** specifies that the passed matrix is on the left, and **MatrixOrderAppend** specifies that the passed matrix is on the right. The default value is **MatrixOrderPrepend**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a texture brush and sets the transformation of the brush.
' The code then uses the transformed brush to fill a rectangle.
' ========================================================================================
SUB Example_TranslateTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a texture brush, and set its transformation.
   DIM pImage AS CGpImage = "HouseAndTree.gif"
   DIM textureBrush AS CGpTextureBrush = @pImage
   textureBrush.TranslateTransform(30, 0, MatrixOrderAppend)
   graphics.FillEllipse(@textureBrush, 0, 0, 400, 200)

END SUB
' ========================================================================================
```
