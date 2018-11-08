# CGpPen Class

Encapsulates a **Pen** object. A **Pen** object is a Windows GDI+ object used to draw lines and curves.

**Inherits from**: CGpBase.<br>
**Include file**: CGpPen.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#Constructors) | Creates a **Pen** object. |
| [Clone](#Clone) | Copies the contents of the existing **Pen** object into a new **Pen** object. |
| [GetAlignment](#GetAlignment) | Gets the alignment currently set for this **Pen** object. |
| [GetBrush](#GetBrush) | Gets the the **Brush** object that is currently set for this **Pen** object. |
| [GetColor](#GetColor) | Gets the color currently set for this **Pen** object. |
| [GetCompoundArray](#GetCompoundArray) | Gets the the compound array currently set for this **Pen** object. |
| [GetCompoundArrayCount](#GetCompoundArrayCount) | Gets the number of elements in a compound array. |
| [GetCustomEndCap](#GetCustomEndCap) | Gets the custom end cap currently set for this **Pen** object. |
| [GetCustomStartCap](#GetCustomStartCap) | Gets the custom end cap currently set for this **Pen** object. |
| [GetDashCap](#GetDashCap) | Gets the dash cap style currently set for this **Pen** object. |
| [GetDashOffset](#GetDashOffset) | Gets the distance from the start of the line to the start of the first space in a dashed line. |
| [GetDashPattern](#GetDashPattern) | Gets an array of custom dashes and spaces currently set for this **Pen** object. |
| [GetDashPatternCount](#GetDashPatternCount) | Gets the number of elements in a dash pattern array. |
| [GetDashStyle](#GetDashStyle) | Gets the dash style currently set for this **Pen** object. |
| [GetEndCap](#GetEndCap) | Gets the end cap currently set for this **Pen** object. |
| [GetLineJoin](#GetLineJoin) | Gets the line join style currently set for this **Pen** object. |
| [GetMiterLimit](#GetMiterLimit) | Gets the miter length currently set for this **Pen** object. |
| [GetPenType](#GetPenType) | Gets the type currently set for this **Pen** object. |
| [GetStartCap](#GetStartCap) | Gets the start cap currently set for this **Pen** object. |
| [GetTransform](#GetTransform) | Gets the world transformation matrix currently set for this **Pen** object. |
| [GetWidth](#GetWidth) | Gets the width currently set for this **Pen** object. |
| [MultiplyTransform](#MultiplyTransform) | Updates the world transformation matrix of this **Pen** object with the product of itself and another matrix. |
| [ResetTransform](#ResetTransform) | Sets the world transformation matrix of this **Pen** object to the identity matrix. |
| [RotateTransform](#RotateTransform) | Updates the world transformation matrix of this **Pen** object with the product of itself and a rotation matrix. |
| [ScaleTransform](#ScaleTransform) | Sets the **Pen** object's world transformation matrix equal to the product of itself and a scaling matrix. |
| [SetAlignment](#SetAlignment) | Sets the alignment for this **Pen** object relative to the line. |
| [SetBrush](#SetBrush) | Sets the **Brush** object that a pen uses to fill a line. |
| [SetColor](#SetColor) | Sets the color for this **Pen** object. |
| [SetCompoundArray](#SetCompoundArray) | Sets the compound array for this **Pen** object. |
| [SetCustomEndCap](#SetCustomEndCap) | Sets the custom end cap for this **Pen** object. |
| [SetCustomStartCap](#SetCustomStartCap) | Sets the custom start cap for this **Pen** object. |
| [SetDashCap](#SetDashCap) | Sets the dash cap style for this **Pen** object. |
| [SetDashOffset](#SetDashOffset) | Sets the distance from the start of the line to the start of the first dash in a dashed line. |
| [SetDashPattern](#SetDashPattern) | Sets an array of custom dashes and spaces for this **Pen** object. |
| [SetDashStyle](#SetDashStyle) | Sets the dash style for this **Pen** object. |
| [SetEndCap](#SetEndCap) | Sets the end cap for this **Pen** object. |
| [SetLineCap](#SetLineCap) | Sets the cap styles for the start, end, and dashes in a line drawn with this pen. |
| [SetLineJoin](#SetLineJoin) | Sets the line join for this **Pen** object. |
| [SetMiterLimit](#SetMiterLimit) | Sets the miter limit of this **Pen** object. |
| [SetStartCap](#SetStartCap) | Sets the start cap for this **Pen** object. |
| [SetTransform](#SetTransform) | Sets the world transformation of this **Pen** object. |
| [SetWidth](#SetWidth) | Sets the width for this **Pen** object. |
| [TranslateTransform](#TranslateTransform) | Updates this brush's current transformation matrix with the product of itself and a translation matrix. |

# <a name="Constructors"></a>Constructors

Creates a **Pen** object that uses a specified color and width.

```
CONSTRUCTOR CGpPen
CONSTRUCTOR CGpPen (BYVAL pPen AS CGpPen PTR)
CONSTRUCTOR CGpPen (BYVAL pBrush AS CGpBrush PTR, BYVAL nWidth AS SINGLE = 1.0)
CONSTRUCTOR CGpPen (BYVAL colour AS ARGB, BYVAL nWidth AS SINGLE = 1.0)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pPen* | Pointer to a **Pen** object to be cloned. |
| *pBrush* | Pointer to a brush to base this pen on. |
| *colour* | An ARGB color that specifies the color for this **Pen** object. |
| *nWidth* | The width of this pen's stroke. The default value is 1.0. |

#### Remarks

If you pass the address of a pen to one of the draw methods of a **Graphics** object, the width of the pen's stroke is dependent on the unit of measure specified in the **Graphics** object. The default unit of measure is **UnitPixel**, which is an element of the **GpUnit** enumeration.

# <a name="Clone"></a>Clone

Copies the contents of the existing **Pen** object into a new **Pen** object.

```
FUNCTION Clone (BYVAL pClonePen AS CGpPen PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pClonePen* | Pointer to the **Pen** object where to copy the contents of the existing object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Pen object, creates a copy of the Pen object, and then
' draws an ellipse using the copied Pen object.
' ========================================================================================
SUB Example_ClonePen (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create and clone a Pen object.
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 4)
   DIM clonedPen AS CGpPen
   pen.Clone(@clonedPen)
   ' // You can also use:
   ' DIM clonedPen AS CGpPen = @pen

   ' // Draw a rectangle using the cloned Pen object.
   graphics.DrawRectangle(@clonedPen, 10, 10, 100, 50)

END SUB
' ========================================================================================
```

# <a name="GetAlignment"></a>GetAlignment

Gets the alignment currently set for this **Pen** object.

```
FUNCTION GetAlignment () AS PenAlignment
```

#### Return value

This method returns an element of the **PenAlignment** enumeration that indicates the current alignment setting for this **Pen** object. The default value of **PenAlignment** is **PenAlignmentCenter**.

#### Example

```
' ========================================================================================
' The following example creates a Pen object, sets the alignment, draws a line, and then
' gets the pen alignment settings.
' ========================================================================================
SUB Example_GetAlignment (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Pen object and set its alignment.
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 255, 0), 15)
   pen.SetAlignment(PenAlignmentCenter)
   
   ' // Draw a line.
   graphics.DrawLine(@pen, 0, 0, 100, 50)

   ' // Obtain information about the Pen object.
   DIM nPenAlignment AS PenAlignment = pen.GetAlignment

   IF nPenAlignment = PenAlignmentCenter THEN
      ' // The pixels will be centered on the theoretical line.
   ELSEIF nPenAlignment = PenAlignmentInset THEN
      '  // The pixels will lie inside the filled area  of the theoretical line.
   END IF

END SUB
' ========================================================================================
```

# <a name="GetBrush"></a>GetBrush

Gets the the **Brush** object that is currently set for this **Pen** object.

```
FUNCTION GetBrush (BYVAL pBrush AS CGpBrush PTR) AS GpStatus
```

#### Return value

This method returns a pointer to a **Brush** object that is currently used to fill a line.

#### Example

```
' ========================================================================================
' The following example creates a Brush object, creates a Pen object based on the Brush
' object, draws a line with the pen, gets the Brush from the pen, and then uses the Brush
' to fill a rectangle.
' ========================================================================================
SUB Example_GetBrush (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a HatchBrush object
   DIM hatchBrush AS CGpHatchBrush = CGpHatchBrush(HatchStyleVertical, GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255))

   ' // Create a pen, and set the brush for the pen
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 255, 0, 0), 10)
   pen.SetBrush(@hatchBrush)

   ' // Draw a line with the pen
   graphics.DrawLine(@pen, 0, 0, 200, 100)

   ' // Get the pen's brush, and use that brush to fill a rectangle.
   DIM pBrush AS CGpBrush
   pen.GetBrush(@pBrush)
   graphics.FillRectangle(@pBrush, 0, 100, 200, 100)

END SUB
' ========================================================================================
```

# <a name="GetColor"></a>GetColor

Gets the color currently set for this **Pen** object.

```
FUNCTION GetColor (BYVAL colour AS ARGB PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *colour* | Pointer to a variable that receives the color of this **Pen** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Pen object and draws a line. The code then gets the color
' of the pen and creates a Brush object based on that color. Finally, the code uses the
' Brush object to fill a rectangle.
' ========================================================================================
SUB Example_GetColor (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a pen, and use it to draw a line.
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 200, 150, 100), 5)
   graphics.DrawLine(@pen, 0, 0, 200, 100)

   ' // Get the pen's color, and use that color to create a brush.
   DIM colour AS ARGB
   pen.GetColor(@colour)
   DIM solidBrush AS CGpSolidBrush = colour

   ' // Use the brush to fill a rectangle.
   graphics.FillRectangle(@solidBrush, 0, 100, 200, 100)

END SUB
' ========================================================================================
```

# <a name="GetCompoundArray"></a>GetCompoundArray

Gets the the compound array currently set for this **Pen** object.

```
FUNCTION GetCompoundArray (BYVAL compoundArray AS SINGLE PTR, BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *compoundArray* | Pointer to an array that receives the compound array. |
| *count* | Integer that specifies the number of elements in the *compoundArray* array. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example declares an array, sets the compound array, draws a line, and gets
' the number of elements in the compound array.
' ========================================================================================
SUB Example_GetCompoundArray (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create an array of real numbers and a Pen object.
   DIM compVals(0 TO 5) AS SINGLE = {0.0, 0.2, 0.5, 0.7, 0.9, 1.0}
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 30)

   ' // Set the compound array of the pen.
   pen.SetCompoundArray(@compVals(0), 6)

   ' // Draw a line with the pen.
   graphics.DrawLine(@pen, 5, 20, 405, 200)

   ' // Obtain information about the pen
   DIM compValues(ANY) AS SINGLE
   DIM nCount AS LONG = pen.GetCompoundArrayCount
   REDIM compValues(nCount -1)
   pen.GetCompoundArray(@compValues(0), nCount)

   FOR j AS LONG = 0 TO nCount - 1
      ' // Inspect or use the value in compValues(j)
      PRINT compValues(j)
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetCompoundArrayCount"></a>GetCompoundArrayCount

Gets the number of elements in a compound array.

```
FUNCTION GetCompoundArrayCount () AS GpStatus
```

#### Example

```
' ========================================================================================
' The following example declares an array, sets the compound array, draws a line, and gets
' the number of elements in the compound array.
' ========================================================================================
SUB Example_GetCompoundArray (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create an array of real numbers and a Pen object.
   DIM compVals(0 TO 5) AS SINGLE = {0.0, 0.2, 0.5, 0.7, 0.9, 1.0}
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 30)

   ' // Set the compound array of the pen.
   pen.SetCompoundArray(@compVals(0), 6)

   ' // Draw a line with the pen.
   graphics.DrawLine(@pen, 5, 20, 405, 200)

   ' // Obtain information about the pen
   DIM compValues(ANY) AS SINGLE
   DIM nCount AS LONG = pen.GetCompoundArrayCount
   REDIM compValues(nCount -1)
   pen.GetCompoundArray(@compValues(0), nCount)

   FOR j AS LONG = 0 TO nCount - 1
      ' // Inspect or use the value in compValues(j)
      PRINT compValues(j)
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetCustomEndCap"></a>GetCustomEndCap

Gets the custom end cap currently set for this **Pen** object.

```
FUNCTION GetCustomEndCap (BYVAL pCustomLineCap AS CGpCustomLineCap PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pCustomLineCap* | Pointer to a **CustomLineCap** object that receives the custom end cap of this **Pen** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object and adds a rectangle to it. The code
' then creates a Pen object, sets the custom end cap using the GraphicsPath object, and
' draws a line. Finally, the code gets the custom end cap of the pen and creates another
' Pen object using the same custom end cap. It then draws a second line.
' ========================================================================================
SUB Example_GetCustomEndCap (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a GraphicsPath object, and add a rectangle to it.
   DIM pStrokePath AS CGpGraphicsPath
   pStrokePath.AddRectangle(-10, -5, 20, 10)

   ' // Create a pen, and set the custom end cap based on the GraphicsPath object.
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255))
   DIM custCap AS CGpCustomLineCap = CGpCustomLineCap(NULL, @pStrokePath)
   pen.SetCustomEndCap(@custCap)

   ' // Draw a line with the custom end cap.
   graphics.DrawLine(@pen, 0, 0, 200, 100)

   ' // Obtain the custom end cap for the pen.
   DIM customLineCap AS CGpCustomLineCap = CGpCustomLineCap(NULL, NULL)
   pen.GetCustomEndCap(@customLineCap)

   ' // Create another pen, and use the same custom end cap.
   DIM pen2 AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 255, 0), 3)
   pen2.SetCustomEndCap(@customLineCap)

   ' // Draw a line using the second pen.
   graphics.DrawLine(@pen2, 0, 100, 200, 150)

END SUB
' ========================================================================================
```

# <a name="GetCustomStartCap"></a>GetCustomStartCap

Gets the custom end cap currently set for this **Pen** object.

```
FUNCTION GetCustomStartCap (BYVAL pCustomLineCap AS CGpCustomLineCap PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pCustomLineCap* | Pointer to a **CustomLineCap** object that receives the custom start cap of this **Pen** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object and adds a rectangle to it. The code
' then creates a Pen object, sets the custom start cap using the GraphicsPath object, and
' draws a line. Finally, the code gets the custom start cap of the pen and creates another
' Pen object using the same custom end cap. It then draws a second line.
' ========================================================================================
SUB Example_GetCustomStartCap (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a GraphicsPath object, and add a rectangle to it.
   DIM pStrokePath AS CGpGraphicsPath
   pStrokePath.AddRectangle(-10, -5, 20, 10)

   ' // Create a pen, and set the custom start cap based on the GraphicsPath object.
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255))
   DIM custCap AS CGpCustomLineCap = CGpCustomLineCap(NULL, @pStrokePath)
   pen.SetCustomStartCap(@custCap)

   ' // Draw a line with the custom start cap.
   graphics.DrawLine(@pen, 20, 20, 200, 100)

   ' // Obtain the custom start cap for the pen.
   DIM customLineCap AS CGpCustomLineCap = CGpCustomLineCap(NULL, NULL)
   pen.GetCustomStartCap(@customLineCap)

   ' // Create another pen, and use the same custom end cap.
   DIM pen2 AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 255, 0), 3)
   pen2.SetCustomStartCap(@customLineCap)

   ' // Draw a line using the second pen.
   graphics.DrawLine(@pen2, 50, 100, 200, 150)

END SUB
' ========================================================================================
```

# <a name="GetDashCap"></a>GetDashCap

Gets the dash cap style currently set for this **Pen** object.

```
FUNCTION GetDashCap () AS DashCap
```

#### Return value

This method returns an element of the **DashCap** enumeration that indicates the dash cap style currently set for this Pen object.

#### Example

```
' ========================================================================================
' The following example creates a Pen object, sets the dash cap style, and draws a line.
' The code then gets the dash cap style of the pen and creates a second Pen object with
' the same dash cap style. Finally, the code draws a second line using the second pen.
' ========================================================================================
SUB Example_GetDashCap (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Pen object
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 20)

   ' // Set the dash style for the pen
   pen.SetDashStyle(DashStyleDash)

   ' // Set a rounded dash cap for the pen
   pen.SetDashCap(DashCapRound)

   ' // Draw a line using the pen
   graphics.DrawLine(@pen, 50, 50, 400, 200)

   ' // Obtain the dash cap for the pen
   DIM nDashCap AS DashCap = pen.GetDashCap

   ' // Create another pen, and use the same dash cap.
   DIM pen2 AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 255, 0), 15)
   pen2.SetDashStyle(DashStyleDash)
   pen2.SetDashCap(nDashCap)

   ' // Draw a second line.
   graphics.DrawLine(@pen2, 50, 100, 400, 300)

END SUB
' ========================================================================================
```

# <a name="GetDashOffset"></a>GetDashOffset

Gets the distance from the start of the line to the start of the first space in a dashed line.

```
FUNCTION GetDashOffset () AS SINGLE
```

#### Return value

This method returns a real number that indicates the distance from the start of the line to the start of the dashes.

# <a name="GetDashPattern"></a>GetDashPattern

Gets an array of custom dashes and spaces currently set for this **Pen** object.

```
FUNCTION GetDashPattern (BYVAL dashArray AS SINGLE PTR, BYVAL nCount AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *dashArray* | Pointer to an array that receives the length of the dashes and spaces in a custom dashed line. |
| *nCount* | Integer that specifies the number of elements in the *dashArray* array. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Return value

This method returns a real number that indicates the distance from the start of the line to the start of the dashes.

#### Example

```
' ========================================================================================
' The following example creates a Pen object, sets the dash pattern array, and draws a
' custom dashed line. The code then gets the number of elements in the custom dashed array.
' ========================================================================================
SUB Example_GetDashPattern (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create and set an array of real numbers.
   DIM dashVals(0 TO 3) AS SINGLE = {5.0, 2.0, 15.0, 4.0}

   ' // Create a Pen
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 0), 5)

   ' // Set the dash pattern for the custom dashed line.
   pen.SetDashPattern(@dashVals(0), 4)

   ' // Draw a line using the pen.
   graphics.DrawLine(@pen, 5, 20, 405, 200)

   ' // Obtain information about the pen.
   DIM nCount AS INT_
   nCount = pen.GetDashPatternCount
   DIM dashValues(nCount - 1) AS SINGLE
   pen.GetDashPattern(@dashValues(0), nCount)

   FOR j AS LONG = 0 TO nCount - 1
      OutputDebugStringW WSTR(dashValues(j))
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetDashPatternCount"></a>GetDashPatternCount

Gets the number of elements in a dash pattern array.

```
FUNCTION GetDashPatternCount () AS INT_
```

#### Example

```
' ========================================================================================
' The following example creates a Pen object, sets the dash pattern array, and draws a
' custom dashed line. The code then gets the number of elements in the custom dashed array.
' ========================================================================================
SUB Example_GetDashPattern (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create and set an array of real numbers.
   DIM dashVals(0 TO 3) AS SINGLE = {5.0, 2.0, 15.0, 4.0}

   ' // Create a Pen
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 0), 5)

   ' // Set the dash pattern for the custom dashed line.
   pen.SetDashPattern(@dashVals(0), 4)

   ' // Draw a line using the pen.
   graphics.DrawLine(@pen, 5, 20, 405, 200)

   ' // Obtain information about the pen.
   DIM nCount AS INT_
   nCount = pen.GetDashPatternCount
   DIM dashValues(nCount - 1) AS SINGLE
   pen.GetDashPattern(@dashValues(0), nCount)

   FOR j AS LONG = 0 TO nCount - 1
      OutputDebugStringW WSTR(dashValues(j))
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetDashStyle"></a>GetDashStyle

Gets the dash style currently set for this **Pen** object.

```
FUNCTION GetDashStyle () AS DashStyle
```

#### Return value

This method returns an element of the DashStyle enumeration that indicates the dash style for this **Pen** object.

#### Example

```
' ========================================================================================
' The following example creates a Pen object, sets the dash style, and draws a dashed line.
' The code then gets the dash style, creates a second pen with the dash style of the first
' pen, and draws a second dashed line.
' ========================================================================================
SUB Example_GetDashStyle (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Pen object
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 255, 0, 0), 9)

   ' // Set the dash style for the pen
   pen.SetDashStyle(DashStyleDashDot)
   graphics.DrawLine(@pen, 20, 20, 200, 100)

   ' // Obtain the dash style for the pen.
   DIM nDashStyle AS DashStyle
   nDashStyle = pen.GetDashStyle

   ' // Create another pen, and use the same dash style.
   DIM pen2 AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 255, 0), 9)
   pen2.SetDashStyle(nDashStyle)

   ' // Draw a second dashed line.
   graphics.DrawLine(@pen2, 20, 60, 200, 140)

END SUB
' ========================================================================================
```

# <a name="GetEndCap"></a>GetEndCap

Gets the end cap currently set for this **Pen** object.

```
FUNCTION GetEndCap () AS LineCap
```

#### Return value

This method returns an element of the **LineCap** enumeration that indicates the end cap for this **Pen** object.

#### Example

```
' ========================================================================================
' The following example creates a Pen object, sets the end cap, and draws a line. The code
' then resets the end cap, draws a second line, resets the end cap again, and draws a third line.
' ========================================================================================
SUB Example_GetEndCap (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Pen object
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 15)

   ' // Set the end cap of the pen, and draw a line.
   pen.SetEndCap(LineCapArrowAnchor)
   graphics.DrawLine(@pen, 20, 20, 200, 100)

   ' // Obtain the end cap for the pen.
   DIM nLineCap AS LineCap
   nLineCap = pen.GetEndCap

   ' // Create another pen, and use the same end cap.
   DIM pen2 AS CGpPen = CGpPen(GDIP_ARGB(255, 255, 0, 0), 9)
   pen2.SetEndCap(nLineCap)

   ' // Draw a second line.
   graphics.DrawLine(@pen2, 20, 60, 200, 140)

END SUB
' ========================================================================================
```

# <a name="GetLineJoin"></a>GetLineJoin

Gets the line join style currently set for this **Pen** object.

```
FUNCTION GetLineJoin () AS LineJoin
```

#### Return value

This method returns an element of the LineJoin enumeration that indicates the style used at the point where line segments join.

#### Example

```
' ========================================================================================
' The following example creates a Pen object, sets the end cap, and draws a line. The code
' then resets the end cap, draws a second line, resets the end cap again, and draws a third line.
' ========================================================================================
SUB Example_GetLineJoin (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Pen object
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 255, 0, 0), 15)

   ' // Set the join style, and draw a rectangle.
   pen.SetLineJoin(LineJoinRound)
   graphics.DrawRectangle(@pen, 20, 20, 200, 100)

   ' // Get the line join for the pen.
'   LineJoin lineJoin = pen.GetLineJoin();
   DIM nLineJoin AS LineJoin = pen.GetLineJoin

   ' // Create another pen, and use the same line join.
   DIM pen2 AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 255, 0), 15)
   pen2.SetLineJoin(nLineJoin)

   ' // Draw a second rectangle.
   graphics.DrawRectangle(@pen2, 250, 20, 200, 100)

END SUB
' ========================================================================================
```

# <a name="GetMiterLimit"></a>GetMiterLimit

Gets the miter length currently set for this **Pen** object.

```
FUNCTION GetMiterLimit () AS SINGLE
```

#### Remarks

The miter length is the distance from the intersection of the line walls on the inside of the join to the intersection of the line walls outside of the join. The miter length can be large when the angle between two lines is small. The miter limit is the maximum allowed ratio of miter length to line width. The default value is 10.

# <a name="GetPenType"></a>GetPenType

Gets the type currently set for this **Pen** object.

```
FUNCTION GetPenType () AS PenType
```

#### Return value

This method returns an element of the **PenType** enumeration that indicates the style of pen currently set for this **Pen** object.

#### Example

```
' ========================================================================================
' The following example creates a HatchBrush object and then passes the address of that
' HatchBrush object to a Pen constructor. The code uses the pen, which has a width of 15,
' to draw a line. The code calls the Pen::GetPenType method to determine the pen's type,
' and then checks to see whether the type is PenTypeHatchFill.
' ========================================================================================
SUB Example_GetPenType (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a HatchBrush object.
   DIM hatchBrush AS CGpHatchBrush = CGpHatchBrush(HatchStyleVertical, _
      GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255))

   ' // Create a pen based on a hatch brush, and use that pen to draw a line.
   DIM pen AS CgpPen = CgpPen(@hatchBrush, 15)
   graphics.DrawLine(@pen, 20, 20, 200, 100)
   
   ' // Obtain information about the pen.
   DIM nPenType AS PenType = pen.GetPenType

   IF nPenType = PenTypeHatchFill THEN
      ' // The pen will draw with a hatch pattern.
      PRINT "Pen type", nPenType
   END IF

END SUB
' ========================================================================================
```

# <a name="GetStartCap"></a>GetStartCap

Gets the start cap currently set for this **Pen** object.

```
FUNCTION GetStartCap () AS LineCap
```

#### Return value

This method returns an element of the **LineCap** enumeration that specifies the start cap for this **Pen** object.

#### Example

```
' ========================================================================================
' The following example creates a Pen object, sets the start cap, and draws a line. The
' code then gets the start cap, creates a second pen based on the first pen, and draws a
' second line.
' ========================================================================================
SUB Example_GetStartCap (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Pen object
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 15)

   ' // Set the start cap of the pen, and draw a line.
   pen.SetStartCap(LineCapArrowAnchor)
   graphics.DrawLine(@pen, 20, 20, 200, 100)

   ' // Obtain the end cap for the pen.
   DIM nLineCap AS LineCap
   nLineCap = pen.GetStartCap

   ' // Create another pen, and use the same end cap.
   DIM pen2 AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 15)
   pen2.SetStartCap(nLineCap)

   ' // Draw a second line.
   graphics.DrawLine(@pen2, 20, 60, 200, 140)

END SUB
' ========================================================================================
```

# <a name="GetTransform"></a>GetTransform

Gets the world transformation matrix currently set for this **Pen** object.

```
FUNCTION GetTransform (BYVAL pMatrix AS CGpMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMatrix* | Pointer to a **Matrix** object that receives the transformation matrix. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Pen object and sets its transformation. The code then
' gets the transformation of the pen.
' ========================================================================================
SUB Example_GetTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a pen and set its transformation.
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   pen.RotateTransform(45)

   ' // Obtain information about the pen.
   DIM matrix AS CGpMatrix
   DIM elements(0 TO 5) AS SINGLE

   pen.GetTransform(@matrix)
   matrix.GetElements(@elements(0))

   FOR j AS LONG = 0 TO 5
      ' // Inspect or use the value in elements(js)
      PRINT elements(j)
   NEXT
   
END SUB
' ========================================================================================
```

# <a name="GetWidth"></a>GetWidth

Gets the width currently set for this **Pen** object.

```
FUNCTION GetWidth () AS SINGLE
```

#### Remarks

If you pass the address of a pen to one of the draw methods of a **Graphics** object, the width of the pen's stroke is dependent on the unit of measure specified in the **Graphics** object. The default unit of measure is **UnitPixel**, which is an element of the **GpUnit** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Pen object with a specified width and draws a line. The
' code then gets the width of the pen, creates a second pen based on the width of the
' first pen, and draws a second line.
' ========================================================================================
SUB Example_GetWidth (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a pen and use it to draw a rectangle
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 15)
   graphics.DrawRectangle(@pen, 20, 20, 200, 100)

   ' // Get the width of the pen.
   DIM nWidth AS SINGLE = pen.GetWidth

   ' // Create another pen that has the same width
   DIM pen2 AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 255, 0), nWidth)

   ' // Draw a second line.
   graphics.DrawLine(@pen2, 20, 60, 200, 140)

END SUB
' ========================================================================================
```

# <a name="MultiplyTransform"></a>MultiplyTransform

Updates the world transformation matrix of this **Pen** object with the product of itself and another matrix.

```
FUNCTION MultiplyTransform (BYVAL pMatrix AS CGpMatrix PTR, _
   BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMatrix* | Pointer to the matrix to be multiplied by this brush's current transformation matrix. |
| *order* | Optional. Element of the **MatrixOrder** enumeration that specifies the order of multiplication. **MatrixOrderPrepend** specifies that the passed matrix is on the left, and **MatrixOrderAppend** specifies that the passed matrix is on the right. The default value is **MatrixOrderPrepend**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example sets a matrix and creates a Pen object. The code then sets the
' width of the pen, applies a rotation matrix and a stretch matrix to the pen, and then
' draws an ellipse.
' ========================================================================================
SUB Example_MultiplyTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Vertical stretch
   DIM matrix AS CGpMatrix = CGpMatrix(1, 0, 0, 4, 0, 0)

   ' // Create a pen, and use it to draw a rectangle.
   DIM pen AS CGpPen = GDIP_ARGB(255, 0, 0, 255)

   pen.SetWidth(5)
   pen.RotateTransform(30)   ' // first rorate
   pen.MultiplyTransform(@matrix, MatrixOrderPrepend)   ' // then strecth

   graphics.DrawEllipse(@pen, 90, 30, 200, 200)

END SUB
' ========================================================================================
```

# <a name="ResetTransform"></a>ResetTransform

Sets the world transformation matrix of this **Pen** object to the identity matrix.

```
FUNCTION ResetTransform () AS GpStatus
```

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

The identity matrix represents a transformation that does nothing. If the world transformation matrix of a **Pen** object is the identity matrix, then no world transformation is applied to items drawn using that Pen object.

#### Example

```
' ========================================================================================
' The following example sets a matrix and creates a Pen object. The code then sets the
' width of the pen, applies a rotation matrix and a stretch matrix to the pen, and then
' draws an ellipse.
' ========================================================================================
SUB Example_ResetTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a pen, and set its transformation
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 2)
   pen.ScaleTransform(8, 4)

   ' // Draw a rectangle with the transformed pen
   graphics.DrawRectangle(@pen, 50, 50, 150, 100)

   ' // Reset the transformation
   pen.ResetTransform

   ' // Draw a rectangle with no pen transformation
   graphics.DrawRectangle(@pen, 250, 50, 150, 100)

END SUB
' ========================================================================================
```

# <a name="RotateTransform"></a>RotateTransform

Updates the world transformation matrix of this **Pen** object with the product of itself and a rotation matrix.

```
FUNCTION RotateTransform (BYVAL angle AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *angle* | The angle, in degrees, of rotation. |
| *order* | Optional. Element of the **MatrixOrder** enumeration that specifies the order of multiplication. **MatrixOrderPrepend** specifies that the passed matrix is on the left, and **MatrixOrderAppend** specifies that the passed matrix is on the right. The default value is **MatrixOrderPrepend**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Pen object, sets a scaling matrix to the pen, and draws
' a rectangle. The code then resets the transformation of the pen and draws a second rectangle.
' ========================================================================================
SUB Example_RotateTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM pen AS CGpPen = CGpPen(&hFF0000FF, 5)   ' // Blue
   pen.ScaleTransform(1, 4)   ' // first stretch
   pen.RotateTransform(30, MatrixOrderAppend)   ' // then rotate
   graphics.DrawEllipse(@pen, 50, 30, 200, 200)

END SUB
' ========================================================================================
```

# <a name="ScaleTransform"></a>ScaleTransform

Sets the **Pen** object's world transformation matrix equal to the product of itself and a scaling matrix.

```
FUNCTION ScaleTransform (BYVAL sx AS SINGLE, BYVAL sy AS SINGLE, _
   BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *sx* | The amount to scale in the x direction. |
| *sy* | The amount to scale in the y direction. |
| *order* | Optional. Element of the **MatrixOrder** enumeration that specifies the order of multiplication. **MatrixOrderPrepend** specifies that the passed matrix is on the left, and **MatrixOrderAppend** specifies that the passed matrix is on the right. The default value is **MatrixOrderPrepend**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Pen object and draws a rectangle. The code then applies
' a scaling transformation to the pen and draws a second rectangle.
' ========================================================================================
SUB Example_ScaleTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a pen, and use it to draw a rectangle
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 2)
   graphics.DrawRectangle(@pen, 50, 50, 150, 100)

   ' // Apply a scaling transformation to the pen
   pen.ScaleTransform(8, 4)

   ' // Draw a rectangle with the transformed pen.
   graphics.DrawRectangle(@pen, 250, 50, 150, 100)

END SUB
' ========================================================================================
```

# <a name="SetAlignment"></a>SetAlignment

Sets the alignment for this **Pen** object relative to the line.

```
FUNCTION SetAlignment (BYVAL nPenAlignment AS PenAlignment) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nPenAlignment* | Element of the **PenAlignment** enumeration that specifies the alignment setting of the pen relative to the line that is drawn. The default value is **PenAlignmentCenter**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

If you set the alignment of a **Pen** object to **PenAlignmentInset**, you cannot use that pen to draw compound lines or triangular dash caps.

#### Example

```
' ========================================================================================
' The following example creates two Pen objects and sets the alignment for one of the pens.
' The code then draws two lines using each of the pens.
' ========================================================================================
SUB Example_SetAlignment (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a black and a green pen.
   DIM blackPen AS CgpPen = CGpPen(GDIP_ARGB(255, 0, 0, 0), 1)
   DIM greenPen AS CgpPen = CGpPen(GDIP_ARGB(255, 0, 255, 0), 15)

   ' // Set the alignment of the green pen.
   greenPen.SetAlignment(PenAlignmentInset)

   ' // Draw two lines using each pen.
   graphics.DrawEllipse(@greenPen, 0, 0, 100, 200)
   graphics.DrawEllipse(@blackPen, 0, 0, 100, 200)

END SUB
' ========================================================================================
```

# <a name="SetBrush"></a>SetBrush

Sets the **Brush** object that a pen uses to fill a line.

```
FUNCTION SetBrush (BYVAL pBrush AS CGpBrush PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pBrush* | Pointer to a **Brush** object for the pen to use to fill a line. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a HatchBrush object and then passes the address of that
' HatchBrush object to a Pen constructor. The code then sets the brush for the pen and
' draws a line.
' ========================================================================================
SUB Example_SetBrush (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a HatchBrush object
   DIM hatchBrush AS CGpHatchBrush = CGpHatchBrush(HatchStyleVertical, GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255))

   ' // Create a pen, and set the brush for the pen
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 255, 0, 0), 10)
   pen.SetBrush(@hatchBrush)

   ' // Draw a line with the pen
   graphics.DrawLine(@pen, 0, 0, 200, 100)

END SUB
' ========================================================================================
```

# <a name="SetColor"></a>SetColor

Sets the color for this **Pen** object.

```
FUNCTION SetColor (BYVAL colour AS ARGB) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *colour* | ARGB color that specifies the color for this **Pen** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a red pen and draws a line. The code then sets the pen's
' color to blue and draws a second line.
' ========================================================================================
SUB Example_SetColor (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a red pen, and use it to draw a line.
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 255, 0, 0), 5)
   graphics.DrawLine(@pen, 0, 0, 200, 100)

   ' // Change the pen's color to blue, and draw a second line.
   pen.SetColor(GDIP_ARGB(255, 0, 0, 255))
   graphics.DrawLine(@pen, 0, 40, 200, 140)

END SUB
' ========================================================================================
```

# <a name="SetCompoundArray"></a>SetCompoundArray

Sets the compound array for this **Pen** object.

```
FUNCTION SetCompoundArray (BYVAL compoundArray AS SINGLE PTR, BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *compoundArray* | Pointer to an array of real numbers that specifies the compound array. The elements in the array must be in increasing order, not less than 0, and not greater than 1. |
| *count* | Positive even integer that specifies the number of elements in the compoundArray array. The integer must not be greater than the number of elements in the compound array. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Pen object and sets the compound array for the pen. The
' code then draws a line using the Pen object.
' ========================================================================================
SUB Example_SetCompoundArray (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create an array of real numbers and a Pen object.
   DIM compVals(0 TO 5) AS SINGLE = {0.0, 0.2, 0.5, 0.7, 0.9, 1.0}
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 30)

   ' // Set the compound array of the pen.
   pen.SetCompoundArray(@compVals(0), 6)

   ' // Draw a line with the pen.
   graphics.DrawLine(@pen, 5, 20, 405, 200)

END SUB
' ========================================================================================
```

# <a name="SetCustomEndCap"></a>SetCustomEndCap

Sets the custom end cap for this **Pen** object.

```
FUNCTION SetCustomEndCap (BYVAL pCustomLineCap AS CGpCustomLineCap PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pCustomLineCap* | Pointer to a **CustomLineCap** object that specifies the custom end cap for this **Pen** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object and adds a rectangle to it. The code
' then creates a Pen object, sets the custom end cap based on the GraphicsPath object, and
' draws a line.
' ========================================================================================
SUB Example_SetCustomEndCap (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a GraphicsPath object, and add a rectangle to it.
   DIM pStrokePath AS CGpGraphicsPath
   pStrokePath.AddRectangle(-10, -5, 20, 10)

   ' // Create a pen, and set the custom end cap based on the GraphicsPath object.
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255))
   DIM custCap AS CGpCustomLineCap = CGpCustomLineCap(NULL, @pStrokePath)
   pen.SetCustomEndCap(@custCap)

   ' // Draw a line with the custom end cap.
   graphics.DrawLine(@pen, 0, 0, 200, 100)

END SUB
' ========================================================================================
```

# <a name="SetCustomStartCap"></a>SetCustomStartCap

Sets the custom start cap for this **Pen** object.

```
FUNCTION SetCustomStartCap (BYVAL pCustomLineCap AS CGpCustomLineCap PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pCustomLineCap* | Pointer to a **CustomLineCap** object that specifies the custom end cap for this **Pen** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object and adds a rectangle to it. The code
' then creates a Pen object, sets the custom end cap based on the GraphicsPath object, and
' draws a line.
' ========================================================================================
SUB Example_SetCustomStartCap (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a GraphicsPath object, and add a rectangle to it.
   DIM pStrokePath AS CGpGraphicsPath
   pStrokePath.AddRectangle(-10, -5, 20, 10)

   ' // Create a pen, and set the custom start cap based on the GraphicsPath object.
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255))
   DIM custCap AS CGpCustomLineCap = CGpCustomLineCap(NULL, @pStrokePath)
   pen.SetCustomStartCap(@custCap)

   ' // Draw a line with the custom end cap.
   graphics.DrawLine(@pen, 20, 20, 200, 100)

END SUB
' ========================================================================================
```

# <a name="SetDashCap"></a>SetDashCap

Sets the dash cap style for this **Pen** object.

```
FUNCTION SetDashCap (BYVAL nDashCap AS DashCap) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nDashCap* | Element of the **DashCap** enumeration that specifies the dash cap for this **Pen** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Pen object, sets the dash style and the dash cap, and
' draws a dashed line.
' ========================================================================================
SUB Example_SetDashCap (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Pen object
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 20)

   ' // Set the dash style for the pen
   pen.SetDashStyle(DashStyleDash)

   ' // Set a triangular dash cap for the pen
   pen.SetDashCap(DashCapTriangle)

   ' // Draw a line using the pen
   graphics.DrawLine(@pen, 20, 20, 200, 100)

END SUB
' ========================================================================================
```

# <a name="SetDashOffset"></a>SetDashOffset

Sets the distance from the start of the line to the start of the first dash in a dashed line.

```
FUNCTION SetDashOffset (BYVAL dashOffset AS SINGLE) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *dashOffset* | The number of times to shift the spaces in a dashed line. Each shift is equal to the length of a space in the dashed line. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Pen object, sets the dash style and the dash cap, and
' draws a dashed line.
' ========================================================================================
SUB Example_SetDashOffset (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Pen object
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 15)

   ' // Set the dash style for the pen
   pen.SetDashStyle(DashStyleDash)

   ' // Draw a line using the pen
   graphics.DrawLine(@pen, 0, 50, 400, 50)

   ' // Set the dash offset value for the pen, and draw a second line.
   pen.SetDashOffset(10)
   graphics.DrawLine(@pen, 0, 80, 400, 80)

END SUB
' ========================================================================================
```

# <a name="SetDashPattern"></a>SetDashPattern

Sets an array of custom dashes and spaces for this **Pen** object.

```
FUNCTION SetDashPattern (BYVAL dashArray AS SINGLE PTR, BYVAL nCount AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *dashArray* | Pointer to an array of real numbers that specifies the length of the custom dashes and spaces. All elements in the array must be positive real numbers. |
| *nCount* | Integer that specifies the number of elements in the dashArray array. The integer must be greater than 0 and not greater than the total number of elements in the array. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates an array of real numbers. The code then creates a Pen
' object, sets the dash pattern based on the array, and then draws the custom dashed line.
' ========================================================================================
SUB Example_SetDashPattern (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create and set an array of real numbers.
   DIM dashVals(0 TO 3) AS SINGLE = {5.0, 2.0, 15.0, 4.0}

   ' // Create a Pen
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 0), 5)

   ' // Set the dash pattern for the custom dashed line.
   pen.SetDashPattern(@dashVals(0), 4)

   ' // Draw a line using the pen.
   graphics.DrawLine(@pen, 5, 20, 405, 200)

END SUB
' ========================================================================================
```

# <a name="SetDashStyle"></a>SetDashStyle

Sets the dash style for this **Pen** object.

```
FUNCTION SetDashStyle (BYVAL nDashStyle AS DashStyle) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nDashStyle* | Element of the **DashStyle** enumeration that specifies the dash style for this Pen object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Pen object, sets the dash style, and draws a dashed line.
' The code then gets the dash style, creates a second pen with the dash style of the first
' pen, and draws a second dashed line.
' ========================================================================================
SUB Example_SetDashStyle (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Pen object
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 20)

   ' // Set the dash style for the pen
   pen.SetDashStyle(DashStyleDashDot)

   ' // Set a triangular dash cap for the pen
   pen.SetDashCap(DashCapTriangle)

   ' // Draw a line using the pen
   graphics.DrawLine(@pen, 20, 20, 200, 100)

   ' // Obtain the dash style for the pen.
   DIM nDashStyle AS DashStyle

   ' // Create another pen, and use the same dash style.
   DIM pen2 AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 255, 0), 9)

   ' // Draw a second dashed line.
   graphics.DrawLine(@pen2, 20, 60, 200, 140)

END SUB
' ========================================================================================
```

# <a name="SetEndCap"></a>SetEndCap

Sets the end cap for this **Pen** object.

```
FUNCTION SetEndCap (BYVAL endCap AS LineCap) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *endCap* | Element of the **LineCap** enumeration that specifies the end cap of a line. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Pen object, sets the end cap, and draws a line. The code
' then resets the end cap, draws a second line, resets the end cap again, and draws a third line.
' ========================================================================================
SUB Example_SetEndCap (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Pen object
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 15)

   ' // Set the end cap of the pen, and draw a line.
   pen.SetEndCap(LineCapArrowAnchor)
   graphics.DrawLine(@pen, 0, 50, 400, 150)

  ' // Reset the end cap, and draw a second line.
   pen.SetEndCap(LineCapTriangle)
   graphics.DrawLine(@pen, 0, 80, 400, 180)

   ' // Reset the end cap, and draw a third line.
   pen.SetEndCap(LineCapRound)
   graphics.DrawLine(@pen, 0, 110, 400, 210)

END SUB
' ========================================================================================
```

# <a name="SetLineCap"></a>SetLineCap

Sets the cap styles for the start, end, and dashes in a line drawn with this pen.

```
FUNCTION SetLineCap (BYVAL startCap AS LineCap, BYVAL endCap AS LineCap, _
   BYVAL nDashCap AS DashCap) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *startCap* | Element of the **LineCap** enumeration that specifies the start cap of a line. |
| *endCap* | Element of the **LineCap** enumeration that specifies the end cap of a line. |
| *nDashCap* | Element of the **DashCap** enumeration that specifies the start and end caps of the dashes in a dashed line. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Pen object and sets the line caps. The code then sets
' the dash style for the pen and draws a line.
' ========================================================================================
SUB Example_SetLineCap (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Pen object
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 15)

   ' // Set line caps for the pen.
   pen.SetLineCap(LineCapArrowAnchor, LineCapTriangle, DashCapRound)

   ' // Set the dash style for the pen.
   pen.SetDashStyle(DashStyleDash)

   ' // Draw a line.
   graphics.DrawLine(@pen, 50, 50, 420, 200)

END SUB
' ========================================================================================
```

# <a name="SetLineJoin"></a>SetLineJoin

Sets the line join for this **Pen** object.

```
FUNCTION SetLineJoin (BYVAL nLineJoin AS LineJoin) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nLineJoin* | Element of the **LineJoin** enumeration that specifies the join style used at the end of a line segment that meets another line segment. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Pen object, sets the line join style, and draws a rectangle.
' The code then resets the line join style and draws a second rectangle.
' ========================================================================================
SUB Example_SetLineJoin (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Pen object
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 15)

   ' // Set the join style, and draw a rectangle.
   pen.SetLineJoin(LineJoinRound)
   graphics.DrawRectangle(@pen, 20, 20, 150, 100)

   ' // Reset the join style, and draw a second rectangle.
   pen.SetLineJoin(LineJoinBevel)
   graphics.DrawRectangle(@pen, 200, 20, 150, 100)

END SUB
' ========================================================================================
```

# <a name="SetMiterLimit"></a>SetMiterLimit

Sets the miter limit of this **Pen** object.

```
FUNCTION SetMiterLimit (BYVAL miterLimit AS SINGLE) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *miterLimit* | Simple precision number that specifies the miter limit of this **Pen** object. A Simple precision number value that is less than 1.0! will be replaced with 1.0!. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.


# <a name="SetStartCap"></a>SetStartCap

Sets the start cap for this **Pen** object.

```
FUNCTION SetStartCap (BYVAL startCap AS LineCap) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *startCap* | Element of the **LineCap** enumeration that specifies the start cap of a line. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Pen object, sets the start cap, and draws a line. The
' code then resets the start cap, draws a second line, resets the start cap again, and
' draws a third line.
' ========================================================================================
SUB Example_SetStartCap (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Pen object
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 15)

   ' // Set the start cap of the pen, and draw a line.
   pen.SetStartCap(LineCapArrowAnchor)
   graphics.DrawLine(@pen, 50, 50, 400, 150)

  ' // Reset the start cap, and draw a second line.
   pen.SetStartCap(LineCapTriangle)
   graphics.DrawLine(@pen, 50, 80, 400, 180)

   ' // Reset the start cap, and draw a third line.
   pen.SetStartCap(LineCapRound)
   graphics.DrawLine(@pen, 50, 110, 400, 210)

END SUB
' ========================================================================================
```

# <a name="SetTransform"></a>SetTransform

Sets the world transformation of this **Pen** object.

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
' The following example creates a scale matrix and a Pen object, and then draws a rectangle.
' The code then scales the pen by the matrix and draws a second rectangle.
' ========================================================================================
SUB Example_SetTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a pen, and use it to draw a rectangle
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 0, 255), 2)
   graphics.DrawRectangle(@pen, 10, 50, 150, 100)

   ' // Scale the pen width by a factor of 20 in the horizontal
   ' // direction and a factor of 10 in the vertical direction.
   DIM matrix AS CGpMatrix = CGpMatrix(20, 0, 0, 10, 0, 0)
   pen.SetTransform(@matrix)

   ' // Draw a rectangle with the transformed pen.
   graphics.DrawRectangle(@pen, 200, 50, 150, 100)

END SUB
' ========================================================================================
```

# <a name="SetWidth"></a>SetWidth

Sets the width for this **Pen** object.

```
FUNCTION SetWidth ((BYVAL nWidth AS SINGLE) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nWidth* | Real number that specifies the width of the pen. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="TranslateTransform"></a>TranslateTransform

Updates this brush's current transformation matrix with the product of itself and a translation matrix.

```
FUNCTION TranslateTransform (BYVAL dx AS SINGLE, BYVAL dy AS SINGLE, _
   BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *dx* | Simple precision number that specifies the horizontal component of the translation. |
| *dy* | Simple precision number that specifies the vertical component of the translation. |
| *order* | Optional. Element of the **MatrixOrder** enumeration that specifies the order of multiplication. **MatrixOrderPrepend** specifies that the passed matrix is on the left, and **MatrixOrderAppend** specifies that the passed matrix is on the right. The default value is **MatrixOrderPrepend**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example sets the world transformation of a Graphics object to a rotation.
' The call to Graphics::TranslateTransform multiplies the Graphics object's existing world
' transformation matrix (rotation) by a translation matrix. The MatrixOrderAppend argument
' specifies that the multiplication is done with the translation matrix on the right. At
' that point, the world transformation matrix of the Graphics object represents a composite
' transformation: first rotate, then translate. The call to DrawEllipse draws a rotated and
' translated ellipse.
' ========================================================================================
SUB Example_TranslateTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM pen AS CGpPen = GDIP_ARGB(255, 0, 0, 255)
   pen.RotateTransform(30.0, MatrixOrderAppend)
   graphics.TranslateTransform(100.0, 50.0, MatrixOrderAppend)
   graphics.DrawEllipse(@pen, 0, 0, 200, 80)

END SUB
' ========================================================================================
```
