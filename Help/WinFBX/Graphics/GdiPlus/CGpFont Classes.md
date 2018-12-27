# CGpFont Class

The **CGpFont** class allows the creation of **Font** objects. The Font object encapsulates the characteristics, such as family, height, size, and style (or combination of styles), of a specific font. A **Font** object is used when drawing strings.

**Inherits from**: CGpBase.<br>
**Include file**: CGpFont.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#ConstructorsFont) | Creates a **Font** object. |
| [Clone](#Clone) | Copies the contents of the existing **Font** object into a new **Font** object. |
| [GetFamily](#GetFamily) | Gets the font family on which this font is based. |
| [GetHeight](#GetHeight) | Gets the line spacing of this font in the current unit of a specified Graphics object. |
| [GetHeight(DPI)](#GetHeightDPI) | Gets the line spacing of this font in pixels. |
| [GetLogFontA](#GetLogFont) | Uses a LOGFONTA structure to get the attributes of this **Font** object. |
| [GetLogFontW](#GetLogFont) | Uses a LOGFONTW structure to get the attributes of this **Font** object. |
| [GetSize](#GetSize) | Returns the font size (commonly called the em size) of this **Font** object. |
| [GetStyle](#GetStyle) | Returns the font size (commonly called the em size) of this **Font** object. |
| [GetUnit](#GetUnit) | Returns the unit of measure of this **Font** object. |
| [IsAvailable](#IsAvailable) | Determines whether this **Font** object was created successfully. |

# CGpFontCollection Class

The **CGpFontCollection** class contains methods for enumerating the font families in a collection of fonts. Objects built from this class include the **InstalledFontCollection** and **PrivateFontCollection**.

**Inherits from**: CGpBase.<br>
**Include file**: CGpFont.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#ConstructorFontCollection) | Creates a **FontCollection** object. |
| [GetFamilies](#GetFamilies) | Gets the font families contained in this font collection. |
| [GetFamilyCount](#GetFamilyCount) | Gets the number of font families contained in this font collection. |

# CGpInstalledFontCollection Class

Extends **CGpFontCollection**. Does not implement new additional methods. 

The **InstalledFontCollection** object represents the fonts installed on the system.

**Inherits from**: CGpFontCollection.<br>
**Include file**: CGpFont.inc.

# CGpPrivateFontCollection Class

Extends **CGpFontCollection**. The **PrivateFontCollection** object is a collection for fonts. This object keeps a collection of fonts specifically for an application. The fonts in the collection can include installed fonts as well as fonts that have not been installed on the system.

**Inherits from**: CGpFontCollection.<br>
**Include file**: CGpFont.inc.

| Name       | Description |
| ---------- | ----------- |
| [AddFontFile](#AddFontFile) | Adds a font file to this private font collection. |
| [AddMemoryFont](#AddMemoryFont) | Adds a font that is contained in system memory to a Windows GDI+ font collection. |

# CGpFontFamily Class

The **CGpFontFamily** class encapsulates a set of fonts that make up a font family. A font family is a group of fonts that have the same typeface but different styles.

**Inherits from**: CGpBase.
**Imclude file**: CGpFont.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#ConstructorFontFamily) | Creates a FontFamily object based on a specified font collection. |
| [Clone](#Clone) | Copies the contents of the existing FontFamily object into a new FontFamily object. |
| [GenericMonospace](#GenericMonospace) | Gets a FontFamily object that a specifies a generic monospace typeface. |
| [GenericSansSerif](#GenericSansSerif) | Gets a FontFamily object that a specifies a generic sans serif typeface. |
| [GenericSerif](#GenericSerif) | Gets a FontFamily object that a specifies a generic serif typeface. |
| [GetCellAscent](#GetCellAscent) | Gets the cell ascent, in design units, of this font family for the specified style or style combination. |
| [GetCellDescent](#GetCellDescent) | Gets the cell descent, in design units, of this font family for the specified style or style combination. |
| [GetEmHeight](#GetEmHeight) | Gets the size (commonly called em size or em height), in design units, of this font family. |
| [GetFamilyName](#GetFamilyName) | Gets the name of this font family. |
| [GetLineSpacing](#GetLineSpacing) | Gets the line spacing, in design units, of this font family for the specified style or style combination. |
| [IsAvailable](#IsAvailableFontFamily) | Determines whether this Font object was created successfully. |
| [IsStyleAvailable](#IsStyleAvailable) | Determines whether the specified style is available for this font family. |

# <a name="ConstructorsFont"></a>Constructors (CGpFont)

Creates a **Font** object.

```
CONSTRUCTOR CGpFont
```

Creates and initializes a **Font** object from another Font object.

```
CONSTRUCTOR CGpFont (BYVAL pFont AS CGpFont PTR)
```

Creates a **Font** object based on the Windows Graphics Device Interface (GDI) font object that is currently selected into a specified device context. This constructor is provided for compatibility with GDI.

```
CONSTRUCTOR CGpFont (BYVAL hdc AS HDC)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hdc* | Handle to a Windows device context that has a font selected. A handle is a number that Windows uses internally to reference an object. |

Creates a **Font** object indirectly from a Windows Graphics Device Interface (GDI) logical font by using a handle to a GDI LOGFONT structure.

```
CONSTRUCTOR CGpFont (BYVAL hdc AS HDC, BYVAL hFont AS HFONT)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hdc* | Handle to a Windows device context. A handle is a number that Windows uses internally to reference an object. |
| *hFont* | Handle to a logical font. |

Creates a **Font** object directly from a Windows Graphics Device Interface (GDI) logical font. This constructor is provided for compatibility with GDI.

```
CONSTRUCTOR CGpFont (BYVAL hdc AS HDC, BYVAL plogfont AS LOGFONTA PTR)
CONSTRUCTOR CGpFont (BYVAL hdc AS HDC, BYVAL plogfont AS LOGFONTW PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hdc* | Handle to a Windows device context. A handle is a number that Windows uses internally to reference an object. |
| *plogfont* | Pointer to a LOGFONT structure variable that contains attributes of the font. |

Creates a **Font** object based on a **FontFamily** object, a size, a font style, and a unit of measurement.

```
CONSTRUCTOR CGpFont (BYVAL pFamily AS CGpFontFamily PTR, BYVAL emSize AS SINGLE, _
   BYVAL nStyle AS INT_ = 0, BYVAL nUnit AS GpUnit = 0)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pFamily* | Pointer to a **FontFamily** object that specifies information such as the string that identifies the font family and the font family's text metrics measured in design units. |
| *emSize* | The *em* size of the font measured in the units specified in the unit parameter. |
| *nStyle* | Optional. Integer that specifies the style of the typeface. This value must be an element of the **FontStyle** enumeration or the result of a bitwise OR applied to two or more of these elements. For example, FontStyleBold OR FontStyleUnderline OR FontStyleStrikeout sets the style as a combination of the three styles. The default value is FontStyleRegular. |
| *nUnit* | Optional. Element of the **Unit** enumeration that specifies the unit of measurement for the font size. The default value is UnitPoint. |

Creates a **Font** object based on a font family, a size, a font style, and a unit of measurement.

```
CONSTRUCTOR CGpFont (BYVAL pwszFamilyName AS WSTRING PTR, BYVAL emSize AS SINGLE, _
   BYVAL nStyle AS INT_ = 0, BYVAL nUnit AS GpUnit = 0, BYVAL pFontCollection AS CGpFontCollection PTR = NULL)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszFamilyName* | Name of the font family. |
| *emSize* | The *em* size of the font measured in the units specified in the unit parameter. |
| *nStyle* | Optional. Integer that specifies the style of the typeface. This value must be an element of the **FontStyle** enumeration or the result of a bitwise OR applied to two or more of these elements. For example, FontStyleBold OR FontStyleUnderline OR FontStyleStrikeout sets the style as a combination of the three styles. The default value is FontStyleRegular. |
| *nUnit*  Optional.| Element of the **Unit** enumeration that specifies the unit of measurement for the font size. The default value is UnitPoint. |
| *pFontCollection* | Optional. Pointer to a **FontCollection** object that specifies a user-defined group of fonts. If the value of this parameter is NULL, the system font collection is used. The default value is NULL. |

# <a name="Clone"></a>Clone (CGpFont)

Copies the contents of the existing **Font** object into a new **Font** object.

```
FUNCTION Clone (BYVAL pCloneFont AS CGpFont PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pCloneFont* | Pointer to the **Font** object where to copy the contents of the existing object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Font object, clones it, and then uses the clone to draw text.
' ========================================================================================
SUB Example_CloneFont (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Font object.
   DIM font AS CGpFont = CGpFont("Arial", AfxPointsToPixelsX(16) / rxRatio)

   ' // Create a clone of the Font object
   DIM cloneFont AS CGpFont
   font.Clone(@cloneFont)

   ' // Draw Text with cloneFont.
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 0)
   DIM wszText AS WSTRING * 260 = "This is a cloned Font"
   graphics.DrawString(@wszText, -1, @cloneFont, 0, 0, @solidbrush)

END SUB
' ========================================================================================
```

# <a name="GetFamily"></a>GetFamily (CGpFont)

Gets the font family on which this font is based.

```
FUNCTION GetFamily (BYVAL pFamily AS CGpFontFamily PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pFamily* | Pointer to a **FontFamily** object that receives the font family. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Font object, retrieves the information about the font
' family on which it is based, and then uses the FontFamily object to create a second Font
' object. The example then uses the second Font object to draw text.
' ========================================================================================
SUB Example_GetFamily (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Font object.
   DIM font AS CGpFont = CGpFont("Arial", AfxPointsToPixelsX(16) / rxRatio)

   ' // Get the FontFamily of myFont.
   DIM fontFamily AS CGpFontFamily
   font.GetFamily(@fontFamily)

   ' // Create a new Font object from fontFamily.
   DIM familyFont AS CGpFont = CGpFont(@fontFamily, AfxPointsToPixelsX(16) / rxRatio)

   ' // Draw Text with familyFont
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 0)
   DIM wszText AS WSTRING * 260 = "This is a Font created from a FontFamily"
   graphics.DrawString(@wszText, -1, @familyFont, 0, 0, @solidbrush)

END SUB
' ========================================================================================
```

# <a name="GetHeight"></a>GetHeight (CGpFont)

Gets the line spacing of this font in the current unit of a specified **Graphics** object. The line spacing is the vertical distance between the base lines of two consecutive lines of text. Thus, the line spacing includes the blank space between lines along with the height of the character itself.

```
FUNCTION GetHeight (BYVAL pGraphics AS CGpGraphics PTR) AS SINGLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *pGraphics* | Pointer to a **Graphics** object whose unit and vertical resolution are used in the height calculation. |

#### Return value

This method returns the line spacing of this font.

#### Remarks

If the font unit is set to anything other than **UnitPixel**, the height, in pixels, is calculated using the vertical resolution of the specified **Graphics** object. For example, suppose the font unit is inches and the font size is 0.3. Also suppose that for the corresponding font family, the em height is 2048 and the line spacing is 2355. If the unit of the **Graphics** object is **UnitPixel** and the vertical resolution of the **Graphics** object is 96 dots per inch, the height is calculated as follows:

```
2355*(0.3/2048)*96 = 33.1171875
```

Continuing with the same example, suppose the unit of the **Graphics** object is something other than **UnitPixel**, say **UnitMillimeter**. Then (using 1 inch = 25.4 millimeters) the height, in millimeters, is calculated as follows:

```
2355*(0.3/2048)25.4 = 8.762256
```

# <a name="GetHeightDPI"></a>GetHeight (DPI) (CGpFont)

Gets the line spacing, in pixels, of this font. The line spacing is the vertical distance between the base lines of two consecutive lines of text. Thus, the line spacing includes the blank space between lines along with the height of the character itself.

```
FUNCTION GetHeight (BYVAL dpi AS SINGLE) AS SINGLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *dpi* | Simple precision number that specifies the vertical resolution, in dots per inch, of the device that displays the font. |

#### Return value

This method returns the line spacing of the font in pixels.

#### Remarks

If the font unit is set to anything other than **UnitPixel**, the height, in pixels, is calculated using the specified vertical resolution. For example, suppose the font unit is inches and the font size is 0.3. Also suppose that for the corresponding font family, the em height is 2048 and the line spacing is 2355. If the specified vertical resolution is 96 dots per inch, the height is calculated as follows:

```
2355*(0.3/2048)*96 = 33.1171875
```

# <a name="GetLogFont"></a>GetLogFontA (CGpFont)

Uses a **LOGFONT** structure to get the attributes of this Font object.

```
FUNCTION GetLogFontA (BYVAL pGraphics AS CGpGraphics PTR, BYVAL pLogFont AS LOGFONTA PTR) AS GpStatus
FUNCTION GetLogFontW (BYVAL pGraphics AS CGpGraphics PTR, BYVAL pLogFont AS LOGFONTW PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pGraphics* | Pointer to a Graphics object that contains attributes of the display device. |
| *pLogFont* | Pointer to a **LOGFONT** structure variable that receives the font attributes. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Font object, gets the font attributes from the Font object,
' and uses these attributes (contained in the LOGFONTA structure) to create a second Font
' object. The second Font object is then used to draw text.
' ========================================================================================
SUB Example_GetLogFontA (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Font object according to the DPI setting
   DIM font AS CGpFont = CGpFont("Arial", AfxPointsToPixelsX(16) / rxRatio)

   ' // Get attributes of font
   DIM logFont AS LOGFONTA
   font.GetLogFontA(@graphics, @logFont)
   ' // Adjust for DPI
   logFont.lfHeight /= rxRatio

   ' // Create a second Font object from logFont
   DIM logfontFont AS CGpFont = CGpFont(hdc, @logFont)

   ' // Draw text using logfontFont.
   DIM solidbrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 0)
   DIM wszText AS WSTRING * 260 = "Font from a LOGFONTA"
   graphics.DrawString(@wszText, -1, @logfontFont, 0, 0, @solidbrush)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' The following example creates a Font object, gets the font attributes from the Font object,
' and uses these attributes (contained in the LOGFONTW structure) to create a second Font
' object. The second Font object is then used to draw text.
' ========================================================================================
SUB Example_GetLogFontW (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Font object according to the DPI setting
   DIM font AS CGpFont = CGpFont("Arial", AfxPointsToPixelsX(16) / rxRatio)

   ' // Get attributes of font
   DIM logFont AS LOGFONTW
   font.GetLogFontW(@graphics, @logFont)
   ' // Adjust for DPI
   logFont.lfHeight /= rxRatio

   ' // Create a second Font object from logFont
   DIM logfontFont AS CGpFont = CGpFont(hdc, @logFont)

   ' // Draw text using logfontFont.
   DIM solidbrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 0)
   DIM wszText AS WSTRING * 260 = "Font from a LOGFONTW"
   graphics.DrawString(@wszText, -1, @logfontFont, 0, 0, @solidbrush)

END SUB
' ========================================================================================
```

# <a name="GetSize"></a>GetSize (CGpFont)

Returns the font size (commonly called the em size) of this Font object. The size is in the units of this **Font** object.

```
FUNCTION GetSize () AS SINGLE
```

#### Return value

The method returns the font size. The size is in the units of this **Font** object.

# <a name="GetStyle"></a>GetStyle (CGpFont)

Returns the font size (commonly called the em size) of this **Font** object. The size is in the units of this **Font** object.

```
FUNCTION GetStyle () AS INT_
```

#### Return value

The method returns the font size. The size is in the units of this **Font** object.

#### Example

```
' ========================================================================================
' The following example creates a Font object, gets the FontStyle associated with the font,
' and creates a second Font object by using the same style. The second Font object is then
' used to draw text.
' ========================================================================================
SUB Example_GetStyle (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Font object according to the DPI setting
   DIM font AS CGpFont = CGpFont("Arial", AfxPointsToPixelsX(16) / rxRatio, FontStyleItalic)

   ' // Get the style of font
   DIM style AS SINGLE = font.GetStyle

   ' // Create a second Font object with the same emSize as myFont.
   DIM styleFont AS CGpFont = CGpFont("Arial", AfxPointsToPixelsX(20) / rxRatio, style)

   ' // Draw text using sizeFont
   DIM solidbrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 0)
   DIM wszText AS WSTRING * 260 = "Font with an acquired style"
   graphics.DrawString(@wszText, -1, @styleFont, 0, 0, @solidbrush)

END SUB
' ========================================================================================
```

# <a name="GetUnit"></a>GetUnit (CGpFont)

Returns the unit of measure of this **Font** object.

```
FUNCTION GetUnit () AS GpUnit
```

#### Return value

This method returns one of the elements of the **Unit** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Font object, gets the unit of measure of the font, and
' then sets the page units of the Graphics object to the returned unit. It then uses the
' Font object to draw text.
' ========================================================================================
SUB Example_GetUnit (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Font object according to the DPI setting
   DIM font AS CGpFont = CGpFont("Arial", AfxPointsToPixelsX(12) / rxRatio)

   ' // Get the unit of measure for font
   DIM unit AS GpUnit = font.GetUnit
   ' // Set the Graphics units of graphics to the retrieved unit value.
   graphics.SetPageUnit(unit)

   ' // Draw text using font
   DIM solidbrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 0)
   DIM wszText AS WSTRING * 260 = "Here is some text"
   graphics.DrawString(@wszText, -1, @font, 0, 0, @solidbrush)

END SUB
' ========================================================================================
```

# <a name="IsAvailable"></a>IsAvailable (CGpFont)

Determines whether this **Font** object was created successfully.

```
FUNCTION IsAvailable () AS BOOLEAN
```

#### Return value

If the font was constructed successfully, this method returns TRUE; otherwise, it returns FALSE.

#### Example

```
' ========================================================================================
' The following example creates a Font object and then tests whether the Font object is
' available. If the Font object is available, it is used to draw text.
' ========================================================================================
SUB Example_IsAvailable (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Font object according to the DPI setting
   DIM font AS CGpFont = CGpFont("Arial", AfxPointsToPixelsX(18) / rxRatio)

   ' // Check whether font is available
   DIM available AS BOOLEAN = font.IsAvailable

   ' // Draw text using font, if it is availiable
   IF available THEN
      DIM solidbrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 0)
      DIM wszText AS WSTRING * 260 = "Here is some text"
      graphics.DrawString(@wszText, -1, @font, 0, 0, @solidbrush)
   END IF
   
END SUB
' ========================================================================================
```
# <a name="ConstructorFontCollection"></a>Constructor (CGpFontCollection)

Creates a **FontCollection** object.

```
CONSTRUCTOR CGpFontCollection
```

# <a name="GetFamilies"></a>GetFamilies (CGpFontCollection)

Gets the font families contained in this font collection.

```
FUNCTION GetFamilies (BYVAL numSought AS INT_, BYVAL gpfamilies AS CGpFontFamily PTR, _
   BYVAL numFound AS INT_ PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *numSought* | Integer that specifies the number of font families in this font collection. |
| *gpfamilies* | Out. Pointer to an array that receives the **FontFamily** objects. |
| *numFound* | Out. Pointer to an LONG that receives the number of font families found in this collection. This number should be the same as the *numSought* value. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

A font family consists of a single font type with related styles. An example of a single font type is Arial Regular. An example of a font family is a set of fonts containing Arial Regular, Arial Italic, and Arial Bold style fonts.

# <a name="GetFamilyCount"></a>GetFamilyCount (CGpFontCollection)

Gets the number of font families contained in this font collection.

```
FUNCTION GetFamilyCount () AS LONG
```

#### Return value

This method returns the number of font families contained in this font collection.

#### Remarks

A font family consists of a single font type with related styles. An example of a single font type is Arial Regular. An example of a font family is a set of fonts containing Arial Regular, Arial Italic, and Arial Bold style fonts.

# <a name="AddFontFile"></a>AddFontFile (CGpPrivateFontCollection)

Adds a font file to this private font collection.

```
FUNCTION AddFontFile (BYVAL pwszFileName AS WSTRING PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszFileName* | Pointer to a wide-character string that specifies the name of a font file. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

When you use the GDI+ API, you must never allow your application to download arbitrary fonts from untrusted sources. The operating system requires elevated privileges to assure that all installed fonts are trusted.

# <a name="AddMemoryFont"></a>AddMemoryFont (CGpPrivateFontCollection)

Adds a font that is contained in system memory to a Windows GDI+ font collection.

```
FUNCTION AddMemoryFont (BYVAL pMemory AS ANY PTR, BYVAL length AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pMemory* | Pointer to a font that is contained in memory. |
| *length* | Integer that specifies the number of bytes of data in the font. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

When you use the GDI+ API, you must never allow your application to download arbitrary fonts from untrusted sources. The operating system requires elevated privileges to assure that all installed fonts are trusted. 

# <a name="ConstructorFontFamily"></a>Constructor (CGpFontFamily)

Creates a **FontFamily** object based on a specified font collection.

```
CONSTRUCTOR CGpFontFamily
CONSTRUCTOR CGpFontFamily (BYVAL pFontFamily AS CGpFontFamily PTR)
CONSTRUCTOR CGpFontFamily (BYVAL pwszName AS WSTRING PTR, pFontCollection AS CGpFontCollection PTR = NULL)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszName* | Name of the font family. For example, Arial.ttf is the name of the Arial font family. |
| *pFontCollection* | Optional. Pointer to a **FontCollection** object that specifies the collection that the font family belongs to. If **FontCollection** is NULL, this font family is not part of a collection. The default value is NULL. |

# <a name="Clone"></a>Clone (CGpFontFamily)

Copies the contents of the existing **FontFamily** object into a new **FontFamily** object.

```
FUNCTION Clone (BYVAL pFontFamily AS CGpBrush PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pFontFamily* | Pointer to a variable that will receive a pointer to the cloned **FontFamily** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a FontFamily object, clones that object, and then creates
' a Font object from the clone. It then uses the Font object to draw text.
' ========================================================================================
SUB Example_CloneFontFamily (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a FontFamily object
   DIM arialFontFamily AS CGpFontFamily = "arial"

   ' // Clone the FontFamily object and use it to create a Font object
   DIM cloneFontFamily AS CGpFontFamily
   arialFontFamily.Clone(@cloneFontFamily)
   DIM arialFont AS CGpFont = CGpFont(@cloneFontFamily, AfxPointsToPixelsX(16) / rxRatio)

   ' // Draw text using the new font
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 0)
   graphics.DrawString("This is an Arial font", -1, @arialFont, 0, 0, @solidbrush)

END SUB
' ========================================================================================
```

# <a name="GenericMonospace"></a>GenericMonospace (CGpFontFamily)

Gets a **FontFamily** object that a specifies a generic monospace typeface.

```
FUNCTION GenericMonospace (BYVAL pFontFamily AS CGpFontFamily PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pFontFamily* | Pointer to a variable that will receive a pointer to the **FontFamily** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Font object by using a generic monospace FontFamily
' object and then uses that Font object to draw text.
' ========================================================================================
SUB Example_GenericMonospace (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Use a generic monospace FontFamily object to create a Font object.
   DIM fontFamily AS CGpFontFamily
   fontFamily.GenericMonospace(@fontFamily)
   DIM genericMonoFont AS CGpFont = CGpFont(@fontFamily, AfxPointsToPixelsX(16) / rxRatio)

   ' // Draw text using the new font
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 0)
   graphics.DrawString("This is a generic monospace font", -1, @genericMonoFont, 0, 0, @solidbrush)

END SUB
' ========================================================================================
```

# <a name="GenericSansSerif"></a>GenericSansSerif (CGpFontFamily)

Gets a **FontFamily** object that a specifies a generic sans serif typeface.

```
FUNCTION GenericSansSerif (BYVAL pFontFamily AS CGpFontFamily PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pFontFamily* | Pointer to a variable that will receive a pointer to the **FontFamily** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Font object by using a generic sans serif FontFamily
' object and then uses that Font object to draw text.
' ========================================================================================
SUB Example_GenericSansSerif (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Use a generic sans serif FontFamily object to create a Font object.
   DIM fontFamily AS CGpFontFamily
   fontFamily.GenericSansSerif(@fontFamily)
   DIM genericSansSerifFont AS CGpFont = CGpFont(@fontFamily, AfxPointsToPixelsX(16) / rxRatio)

   ' // Draw text using the new font
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 0)
   graphics.DrawString("This is a generic sans serif font", -1, @genericSansSerifFont, 0, 0, @solidbrush)

END SUB
' ========================================================================================
```

# <a name="GenericSerif"></a>GenericSerif (CGpFontFamily)

Gets a **FontFamily** object that a specifies a generic serif typeface.

```
FUNCTION GenericSansSerif (BYVAL pFontFamily AS CGpFontFamily PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pFontFamily* | Pointer to a variable that will receive a pointer to the **FontFamily** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Font object by using a generic serif FontFamily
' object and then uses that Font object to draw text.
' ========================================================================================
SUB Example_GenericSerif (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Use a generic serif FontFamily object to create a Font object.
   DIM fontFamily AS CGpFontFamily
   fontFamily.GenericSerif(@fontFamily)
   DIM genericSerifFont AS CGpFont = CGpFont(@fontFamily, AfxPointsToPixelsX(16) / rxRatio)

   ' // Draw text using the new font
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 0)
   graphics.DrawString("This is a generic serif font", -1, @genericSerifFont, 0, 0, @solidbrush)

END SUB
' ========================================================================================
```

# <a name="GetCellAscent"></a>GetCellAscent (CGpFontFamily)

Gets the cell ascent, in design units, of this font family for the specified style or style combination.

```
FUNCTION GetCellAscent (BYVAL nStyle AS INT_) AS UINT16
```

| Parameter  | Description |
| ---------- | ----------- |
| *nStyle* | Integer that specifies the style of the typeface. This value must be an element of the **FontStyle** enumeration or the result of a bitwise OR applied to two or more of these elements. For example, FontStyleBold OR FontStyleUnderline OR FontStyleStrikeout specifies a combination of the three styles. |

#### Return value

This method returns the cell ascent, in design units, of this font family for the specified style or style combination.

#### Example

```
' ========================================================================================
' The following example creates a FontFamily object, gets the cell ascent in design units,
' and outputs the value as text.
' ========================================================================================
SUB Example_GetCellAscent (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a FontFamily object
   DIM ascentFontFamily AS CGpFontFamily = "arial"

   ' // Get the cell ascent of the font family in design units
   DIM cellAscent AS LONG = ascentFontFamily.GetCellAscent(FontStyleRegular)

   ' // Copy the cell ascent into a string and draw the string
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 0)
   DIM font AS CGpFont = CGpFont(@ascentFontFamily, AfxPointsToPixelsX(16) / rxRatio)
   DIM wszText AS WSTRING * 260
   wszText = "ascentFontFamily.GetCellAscent() returns " & STR(cellAscent)
   graphics.DrawString(@wszText, -1, @font, 0, 0, @solidbrush)

END SUB
' ========================================================================================
```

# <a name="GetCellDescent"></a>GetCellDescent (CGpFontFamily)

Gets the cell descent, in design units, of this font family for the specified style or style combination.

```
FUNCTION GetCellDescent (BYVAL nStyle AS INT_) AS UINT16
```

| Parameter  | Description |
| ---------- | ----------- |
| *nStyle* | Integer that specifies the style of the typeface. This value must be an element of the **FontStyle** enumeration or the result of a bitwise OR applied to two or more of these elements. For example, FontStyleBold OR FontStyleUnderline OR FontStyleStrikeout specifies a combination of the three styles. |

#### Return value

This method returns the cell ascent, in design units, of this font family for the specified style or style combination.

#### Example

```
' ========================================================================================
' The following example creates a FontFamily object, gets the cell descent in design units,
' and outputs the value as text.
' ========================================================================================
SUB Example_GetCellDescent (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a FontFamily object
   DIM descentFontFamily AS CGpFontFamily = "arial"

   ' // Get the cell ascent of the font family in design units
   DIM cellDescent AS LONG = descentFontFamily.GetCellDescent(FontStyleRegular)

   ' // Copy the cell descent into a string and draw the string
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 0)
   DIM font AS CGpFont = CGpFont(@descentFontFamily, AfxPointsToPixelsX(16) / rxRatio)
   DIM wszText AS WSTRING * 260
   wszText = "ascentFontFamily.GetCellAscent() returns " & STR(cellDescent)
   graphics.DrawString(@wszText, -1, @font, 0, 0, @solidbrush)

END SUB
' ========================================================================================
```

# <a name="GetEmHeight"></a>GetEmHeight (CGpFontFamily)

Gets the size (commonly called em size or em height), in design units, of this font family.

```
FUNCTION GetEmHeight (BYVAL nStyle AS INT_) AS UINT16
```

| Parameter  | Description |
| ---------- | ----------- |
| *nStyle* | Integer that specifies the style of the typeface. This value must be an element of the **FontStyle** enumeration or the result of a bitwise OR applied to two or more of these elements. For example, FontStyleBold OR FontStyleUnderline OR FontStyleStrikeout specifies a combination of the three styles. |

#### Return value

This method returns the size, in design units, of this font family.

#### Example

```
' ========================================================================================
' The following example creates a FontFamily object, gets the em height in design units,
' and outputs the value as text.
' ========================================================================================
SUB Example_GetEmHeight (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a FontFamily object
   DIM emHeightFontFamily AS CGpFontFamily = "arial"

   ' // Get the cell ascent of the font family in design units.
   DIM emHeight AS LONG = emHeightFontFamily.GetEmHeight(FontStyleRegular)

   ' // Copy the height into a string and draw the string.
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 0)
   DIM font AS CGpFont = CGpFont(@emHeightFontFamily, AfxPointsToPixelsX(16) / rxRatio)
   DIM wszText AS WSTRING * 260
   wszText = "emHeightFontFamily.GetEmHeight() returns " & STR(emHeight)
   graphics.DrawString(@wszText, -1, @font, 0, 0, @solidbrush)

END SUB
' ========================================================================================
```

# <a name="GetFamilyName"></a>GetFamilyName (CGpFontFamily)

Gets the name of this font family.

```
FUNCTION GetFamilyName (BYVAL pwszName AS WSTRING PTR, BYVAL language AS LANGID = LANG_NEUTRAL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszName* | Out. Name of this font family. |
| *language* | In, optional. Sixteen-bit value that specifies the language to use. The default value is LANG_NEUTRAL, which is the user's default language. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a FontFamily object, gets the family name, and outputs the
' name as text.
' ========================================================================================
SUB Example_GetFamilyName (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a FontFamily object
   DIM nameFontFamily AS CGpFontFamily = "arial"

   ' // Get the cell ascent of the font family in design units.
   DIM familyName AS WSTRING * LF_FACESIZE
   nameFontFamily.GetFamilyName(@familyName)

   ' // Draw the family name
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 0)
   DIM font AS CGpFont = CGpFont(@nameFontFamily, AfxPointsToPixelsX(16) / rxRatio)
   graphics.DrawString(@familyName, -1, @font, 0, 0, @solidbrush)

END SUB
' ========================================================================================
```

# <a name="GetLineSpacing"></a>GetLineSpacing (CGpFontFamily)

Gets the line spacing, in design units, of this font family for the specified style or style combination. The line spacing is the vertical distance between the base lines of two consecutive lines of text.

```
FUNCTION GetLineSpacing (BYVAL nStyle AS INT_) AS UINT16
```

| Parameter  | Description |
| ---------- | ----------- |
| *nStyle* | Integer that specifies the style of the typeface. This value must be an element of the **FontStyle** enumeration or the result of a bitwise OR applied to two or more of these elements. For example, FontStyleBold OR FontStyleUnderline OR FontStyleStrikeout specifies a combination of the three styles. |

#### Return value

This method returns the line spacing of this font family.

##### Example

```
' ========================================================================================
' The following example creates a FontFamily object, gets the line spacing in design units,
' and outputs the value as text.
' ========================================================================================
SUB Example_GetLineSpacing (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a FontFamily object
   DIM lineSpacingFontFamily AS CGpFontFamily = "arial"

   ' // Get the cell ascent of the font family in design units.
   DIM lineSpacing AS LONG = lineSpacingFontFamily.GetLineSpacing(FontStyleRegular)

   ' // Copy the line spacing into a string and draw the string.
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 0)
   DIM font AS CGpFont = CGpFont(@lineSpacingFontFamily, AfxPointsToPixelsX(16) / rxRatio)
   DIM wszText AS WSTRING * 260 = "lineSpacingFontFamily.GetLineSpacing() returns " & STR(lineSpacing)
   graphics.DrawString(@wszText, -1, @font, 0, 0, @solidbrush)

END SUB
' ========================================================================================
```

# <a name="IsAvailableFontFamily"></a>IsAvailable (CGpFontFamily)

Determines whether this **Font** object was created successfully.

```
FUNCTION IsAvailable () AS BOOLEAN
```

#### Return value

If the font was constructed successfully, this method returns TRUE; otherwise, it returns FALSE.

#### Example

```
' ========================================================================================
' The following example creates a FontFamily object. If the FontFamily object is created
' successfully, the example draws text.
' ========================================================================================
SUB Example_IsAvailable (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a FontFamily object
   DIM myFontFamily AS CGpFontFamily = "arial"

   ' // Check to see if myFontFamily is available.
   DIM isAvailable AS BOOLEAN = myFontFamily.IsAvailable

   ' // If myFontFamily is available, draw text.
   IF isAvailable THEN
      DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 0)
      DIM font AS CGpFont = CGpFont(@myFontFamily, AfxPointsToPixelsX(16) / rxRatio)
      DIM wszText AS WSTRING * 260 = "myFontFamily is available"
      graphics.DrawString(@wszText, -1, @font, 0, 0, @solidbrush)
   END IF

END SUB
' ========================================================================================
```

# <a name="IsStyleAvailable"></a>IsStyleAvailable (CGpFontFamily)

Determines whether the specified style is available for this font family.

```
FUNCTION IsStyleAvailable (BYVAL nStyle AS INT_) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *nStyle* | Integer that specifies the style of the typeface. This value must be an element of the **FontStyle** enumeration or the result of a bitwise OR applied to two or more of these elements. For example, FontStyleBold OR FontStyleUnderline OR FontStyleStrikeout specifies a combination of the three styles. |

#### Return value

If the style or combination of styles is available, this method returns TRUE; otherwise, it returns FALSE.

#### Remarks

This method returns a misleading result on some third-party fonts. For example, IsStyleAvailable(FontStyleUnderline) may return FALSE because it is really testing for a regular style font that also is an underlined font: (FontStyleRegular OR FontStyleUnderline). If the font does not have a regular style, the IsStyleAvailable method returns FALSE.

#### Example

```
' ========================================================================================
' The following example creates a FontFamily object. If the font family has a regular style
' available, the example draws text.
' ========================================================================================
SUB Example_IsStyleAvailable (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a FontFamily object
   DIM myFontFamily AS CGpFontFamily = "arial"

   ' // Check to see if the regular style is available.
   DIM isStyleAvailable AS BOOLEAN = myFontFamily.IsStyleAvailable(FontStyleRegular)
   
   ' // If regular style is available, draw text.
   IF isStyleAvailable THEN
      DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 0)
      DIM font AS CGpFont = CGpFont(@myFontFamily, AfxPointsToPixelsX(16) / rxRatio)
      DIM wszText AS WSTRING * 260 = "myFontFamily is available in regular style"
      graphics.DrawString(@wszText, -1, @font, 0, 0, @solidbrush)
   END IF

END SUB
' ========================================================================================
```
