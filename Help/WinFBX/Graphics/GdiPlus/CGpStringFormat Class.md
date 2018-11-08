# CGpStringFormat Class

The **CGpStringFormat** class encapsulates text layout information (such as alignment, orientation, tab stops, and clipping) and display manipulations (such as trimming, font substitution for characters that are not supported by the requested font, and digit substitution for languages that do not use Western European digits). A **StringFormat** object can be passed to the **DrawString** methods to format a string.

**Inherits from**: CGpBase.<br>
**Include file**: CGpStringFormat.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#Constructors) | Creates a **StringFormat** object based on string format flags and a language. |
| [Clone](#Clone) | Copies the contents of the existing **StringFormat** object into a new **StringFormat** object. |
| [GenericDefault](#GenericDefault) | Creates a generic, default **StringFormat** object. |
| [GenericTypographic](#GenericTypographic) | Creates a generic, typographic **StringFormat** object. |
| [GetAlignment](#GetAlignment) | Gets an element of the **StringAlignment** enumeration that indicates the character alignment of this **StringFormat** object in relation to the origin of the layout rectangle. |
| [GetDigitSubstitutionLanguage](#GetDigitSubstitutionLanguage) | Gets the language that corresponds with the digits that are to be substituted for Western European digits. |
| [GetDigitSubstitutionMethod](#GetDigitSubstitutionMethod) | Gets an element of the **StringDigitSubstitute** enumeration that indicates the digit substitution method that is used by this **StringFormat** object. |
| [GetFormatFlags](#GetFormatFlags) | Gets the string format flags for this **StringFormat** object. |
| [GetHotkeyPrefix](#GetHotkeyPrefix) | Gets an element of the **HotkeyPrefix** enumeration that indicates the type of processing that is performed on a string when a hot key prefix, an ampersand (&), is encountered. |
| [GetLineAlignment](#GetLineAlignment) | Gets an element of the **StringAlignment** enumeration that indicates the line alignment of this **StringFormat** object in relation to the origin of the layout rectangle. |
| [GetMeasurableCharacterRangeCount](#GetMeasurableCharacterRangeCount) | Gets the number of measurable character ranges that are currently set. |
| [GetTabStopCount](#GetTabStopCount) | Gets the number of tab-stop offsets in this **StringFormat** object. |
| [GetTabStops](#GetTabStops) | Gets the offsets of the tab stops in this **StringFormat** object. |
| [GetTrimming](#GetTrimming) | Gets an element of the **StringTrimming** enumeration that indicates the trimming style of this **StringFormat** object. |
| [SetAlignment](#SetAlignment) | Sets the character alignment of this **StringFormat** object in relation to the origin of the layout rectangle. |
| [SetDigitSubstitution](#SetDigitSubstitution) | Sets the digit substitution method and the language that corresponds to the digit substitutes. |
| [SetFormatFlags](#SetFormatFlags) | Sets the format flags for this **StringFormat** object. |
| [SetHotkeyPrefix](#SetHotkeyPrefix) | Sets the type of processing that is performed on a string when the hot key prefix, an ampersand (&), is encountered. |
| [SetLineAlignment](#SetLineAlignment) | Sets the line alignment of this **StringFormat** object in relation to the origin of the layout rectangle. |
| [SetMeasurableCharacterRanges](#SetMeasurableCharacterRanges) | Sets a series of character ranges for this **StringFormat** object that, when in a string, can be measured by the **MeasureCharacterRanges** method. |
| [SetTabStops](#SetTabStops) | Sets the offsets for tab stops in this **StringFormat** object. |
| [SetTrimming](#SetTrimming) | Sets the trimming style for this **StringFormat** object. |

# <a name="Constructors"></a>Constructors

Creates **StringFormat** object from another **StringFormat** object.

```
CONSTRUCTOR CGpStringFormat (BYVAL pStringFormat AS CGpStringFormat PTR)
```

Creates a **StringFormat** object based on string format flags and a language.

```
CONSTRUCTOR CGpStringFormat (BYVAL formatFlags AS INT_ = 0, BYVAL language AS LANGID = LANG_NEUTRAL)
```

| Parameter  | Description |
| ---------- | ----------- |
| *formatFlags* | Value that specifies the format flags that control most of the characteristics of the **StringFormat** object. The flags are set by applying a bitwise OR to elements of the **StringFormatFlags** enumeration. The default value is 0 (no flags set). |
| *language* | Sixteen-bit value that specifies the language to use. The default value is LANG_NEUTRAL, which is the user's default language. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

# <a name="Clone"></a>Clone

Copies the contents of the existing **StringFormat** object into a new **StringFormat** object.

```
FUNCTION Clone (BYVAL pStringFormat AS CGpStringFormat PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pStringFormat* | Pointer to the **StringFormat** object where to copy the contents of the existing object. |

#### Example

```
' ========================================================================================
' The following example creates a StringFormat object, clones it, and then uses the clone
' to draw a formatted string.
' ========================================================================================
SUB Example_Clone (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a red solid brush
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   ' // Create a font family from name
   DIM fontFamily AS CGpFontFamily = "Times New Roman"
   ' // Create a font from the font family
   DIM pFont AS CGpFont = CGpFont(@fontFamily, 24, FontStyleRegular, UnitPixel)

   ' // Create a string format object and set the alignment
   DIM stringFormat AS CGpStringFormat
   stringFormat.SetAlignment(StringAlignmentCenter)

   ' // Clone the StringFormat object
   DIM pStringFormat AS CGpStringFormat
   stringFormat.Clone(@pStringFormat)
   ' // You can also use:
   ' DIM pStringFormat AS CGpStringFormat = @stringFormat

   ' // Use the cloned StringFormat object in a call to DrawString
   DIM wszText AS WSTRING * 260 = "This text was formatted by a StringFormat object."
   graphics.DrawString(@wszText, LEN(wszText), @pFont, 30, 30, 200, 200, @pStringFormat, @solidBrush)

   ' // Draw the rectangle that encloses the text
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawRectangle(@pen, 30, 30, 200, 200)

END SUB
' ========================================================================================
```

# <a name="GenericDefault"></a>GenericDefault

Creates a generic, default **StringFormat** object.

```
FUNCTION GenericDefault (BYVAL pStringFormat AS CGpStringFormat PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pStringFormat* | Pointer to the **StringFormat** object where to return the new **StringFormat** object. |

#### Remarks

A generic, default **StringFormat** object has the following characteristics:

* No string format flags are set.
* Character alignment and line alignment are set to **StringAlignmentNear**.
* Language ID is set to neutral language, which means that the current language associated with the calling thread is used.
* String digit substitution is set to **StringDigitSubstituteUser**.
* Hot key prefix is set to **HotkeyPrefixNone**.
* Number of tab stops is set to zero.
* String trimming is set to **StringTrimmingCharacter**.

#### Example

```
' ========================================================================================
' The following example creates a generic, default StringFormat object and then uses it to
' draw a formatted string. The code also draws the string's layout rectangle.
' ========================================================================================
SUB Example_GenericDefault (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a red solid brush
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   ' // Create a font family from name
   DIM fontFamily AS CGpFontFamily = "Times New Roman"
   ' // Create a font from the font family
   ' // Points must not be scaled for a generic font, so we are unscaling them dividing by rxRatio
   DIM pFont AS CGpFont = CGpFont(@fontFamily, AfxPointsToPixelsX(12) / rxRatio, FontStyleRegular, UnitPixel)

   ' // Create a generic default string format
   DIM stringFormat AS CGpStringFormat
   stringFormat.GenericDefault(@stringFormat)

   ' // Use the generic StringFormat object in a call to DrawString.
   DIM wszText AS WSTRING * 260 = "This text was formatted by a generic StringFormat object."
   graphics.DrawString(@wszText, LEN(wszText), @pFont, 30, 30, 100, 120, @stringFormat, @solidBrush)

   ' // Draw the rectangle that encloses the text.
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawRectangle(@pen, 30, 30, 100, 120)

END SUB
' ========================================================================================
```

# <a name="GenericTypographic"></a>GenericTypographic

Creates a generic, typographic **StringFormat** object.

```
FUNCTION GenericTypographic (BYVAL pStringFormat AS CGpStringFormat PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pStringFormat* | Pointer to the **StringFormat** object where to return the new **StringFormat** object. |

#### Remarks

A generic, typographic **StringFormat** object has the following characteristics:

* String format flags **StringFormatFlagsLineLimit**, **StringFormatFlagsNoClip**, and **StringFormatFlagsNoFitBlackBox** are set.
* Character alignment and line alignment are set to **StringAlignmentNear**.
* Language ID is set to neutral language, which means that the current language associated with the calling thread is used.
* String digit substitution is set to **StringDigitSubstituteUser**.
* Hot key prefix is set to **HotkeyPrefixNone**.
* Number of tab stops is set to zero.
* String trimming is set to **StringTrimmingNone**.

#### Example

```
' ========================================================================================
' The following example creates a generic, default StringFormat object and then uses it to
' draw a formatted string. The code also draws the string's layout rectangle.
' ========================================================================================
SUB Example_GenericTypographic (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a red solid brush
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   ' // Create a font family from name
   DIM fontFamily AS CGpFontFamily = "Times New Roman"
   ' // Create a font from the font family
   ' // Points must not be scaled for a typographic font, so we are unscaling them dividing by rxRatio
   DIM pFont AS CGpFont = CGpFont(@fontFamily, AfxPointsToPixelsX(12) / rxRatio, FontStyleRegular, UnitPixel)

   ' // Create a typographic string format object
   DIM stringFormat AS CGpStringFormat
   stringFormat.GenericTypographic(@stringFormat)

   ' // Use the generic StringFormat object in a call to DrawString.
   DIM wszText AS WSTRING * 260 = "Formatted by a generic typographic StringFormat object."
   graphics.DrawString(@wszText, LEN(wszText), @pFont, 30, 30, 100, 120, @stringFormat, @solidBrush)

   ' // Draw the rectangle that encloses the text.
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawRectangle(@pen, 30, 30, 100, 120)

END SUB
' ========================================================================================
```

# <a name="GetAlignment"></a>GetAlignment

Gets an element of the **StringAlignment** enumeration that indicates the character alignment of this **StringFormat** object in relation to the origin of the layout rectangle. A layout rectangle is used to position the displayed string.

```
FUNCTION GetAlignment () AS StringAlignment
```

#### Example

```
' ========================================================================================
' The following example creates a StringFormat object, sets the string's character alignment,
' gets the alignment value, and stores it in a variable. The code then creates a second
' StringFormat object and uses the stored alignment value to set the character alignment
' of the second StringFormat object. Next, the code uses the second StringFormat object to
' draw a formatted string. The code also draws the string's layout rectangle.
' ========================================================================================
SUB Example_GetAlignment (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a red solid brush
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   ' // Create a font family from name
   DIM fontFamily AS CGpFontFamily = "Times New Roman"
   ' // Create a font from the font family
   DIM pFont AS CGpFont = CGpFont(@fontFamily, 24, FontStyleRegular, UnitPixel)

   ' // Create a string format object and set the alignment
   DIM stringFormat AS CGpStringFormat
   stringFormat.SetAlignment(StringAlignmentFar)

   ' // Get the alignment setting from the StringFormat object.
   DIM nStringAlignment AS StringAlignment = stringFormat.GetAlignment

   ' // Create a second StringFormat object with the same alignment.
   DIM stringFormat2 AS CGpStringFormat
   stringFormat2.SetAlignment(nStringAlignment)

   ' // Use the second StringFormat object in a call to DrawString.
   DIM wszText AS WSTRING * 260 = "This text was formatted by a second StringFormat object."
   graphics.DrawString(@wszText, LEN(wszText), @pFont, 30, 30, 150, 200, @stringFormat2, @solidBrush)

   ' // Draw the rectangle that encloses the text
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawRectangle(@pen, 30, 30, 150, 200)

END SUB
' ========================================================================================
```

# <a name="GetDigitSubstitutionLanguage"></a>GetDigitSubstitutionLanguage

Gets the language that corresponds with the digits that are to be substituted for Western European digits.

```
FUNCTION GetDigitSubstitutionLanguage () AS LANGID
```

#### Return value

This method returns a 16-bit value that forms a National Language Support (NLS) language identifier. This identifier indicates the language that corresponds with the substitution digits. For example, if this **StringFormat** object uses Arabic substitution digits, then this method will return a value that indicates an Arabic language. An NLS language identifier is constructed by the MAKELANGID macro declared in Winnt.bi.

# <a name="GetDigitSubstitutionMethod"></a>GetDigitSubstitutionMethod

Gets an element of the **StringDigitSubstitute** enumeration that indicates the digit substitution method that is used by this **StringFormat** object.

```
FUNCTION GetDigitSubstitutionMethod () AS StringDigitSubstitute
```

#### Return value

The digit substitution method replaces, in a string, Western European digits with digits that correspond to a user's locale or language.

# <a name="GetFormatFlags"></a>GetFormatFlags

Gets the string format flags for this **StringFormat** object.

```
FUNCTION GetFormatFlags () AS INT_
```

#### Return value

This method returns a value that indicates which string format flags are set for this **StringFormat** object. This value can be any combination (the result of a bitwise OR applied to two or more elements) of elements of the **StringFormatFlag**s enumeration.

#### Example

```
' ========================================================================================
' The following example creates a StringFormat object, sets the string's format flags, and
' then gets the 32-bit value that contains the format flags and stores it in a variable.
' The code then creates another StringFormat object and uses the stored format flags value
' to set the format flags of the second StringFormat object. Next, the code uses the second
' StringFormat object to draw a formatted string . The code also draws the string's layout
' rectangle.
' ========================================================================================
SUB Example_GetFormatFlags (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a red solid brush
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   ' // Create a font family from name
   DIM fontFamily AS CGpFontFamily = "Times New Roman"
   ' // Create a font from the font family
   DIM pFont AS CGpFont = CGpFont(@fontFamily, 24, FontStyleRegular, UnitPixel)

   ' // Create a string format object and set its format flags
   DIM stringFormat AS CGpStringFormat
   stringFormat.SetFormatFlags(StringFormatFlagsDirectionVertical OR StringFormatFlagsNoFitBlackBox)

   ' // Get the format flags from the StringFormat object.
   DIM flags AS LONG = stringFormat.GetFormatFlags

   ' // Create a second StringFormat object with the same flags.
   DIM stringFormat2 AS CGpStringFormat
   stringFormat2.SetFormatFlags(flags)

      ' // Use the second StringFormat object in a call to DrawString
   DIM wszText AS WSTRING * 260 = "This text is vertical because of a format flag."
   graphics.DrawString(@wszText, LEN(wszText), @pFont, 30, 30, 150, 200, @stringFormat2, @solidBrush)

   ' // Draw the rectangle that encloses the text
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawRectangle(@pen, 30, 30, 150, 200)

END SUB
' ========================================================================================
```

# <a name="GetHotkeyPrefix"></a>GetHotkeyPrefix

Gets an element of the **HotkeyPrefix** enumeration that indicates the type of processing that is performed on a string when a hot key prefix, an ampersand (&), is encountered.

```
FUNCTION GetHotkeyPrefix () AS HotkeyPrefix
```

#### Remarks

Hot keys, also called access keys, are keys that are programmed to provide an end user with keyboard shortcuts to functionality and are activated by pressing the ALT key. The keys are application dependent and are identified by an underscored letter, typically in a menu name or menu item; for example, when you press ALT, the letter F of the File menu is underscored. The F key is a shortcut to display the **File** menu.

A client programmer designates a hot key in an application by using the hot key prefix, an ampersand (&), in a string that typically is displayed as the name of a menu or an item in a menu and by using the **SetHotkeyPrefix** method to set the appropriate type of processing. When a character in a string is preceded with an ampersand, the key that corresponds to the character becomes a hot key during the processing that occurs when the string is drawn on the display device. The ampersand is called a hot key prefix because it precedes the character to be activated. If **HotkeyPrefixNone** is passed to **SetHotkeyPrefix**, no processing of hot key prefixes occurs.

**Note**: The term hot key is used synonymously here with the term access key. The term hot key may have a different meaning in other Windows APIs.

#### Example

```
' ========================================================================================
' The following example creates a StringFormat object, sets the type of hot key prefix
' processing to be performed on the string, and then gets the type of processing and stores
' it in a variable. The code then creates a second StringFormat object and uses the stored
' value to set the type of hot key prefix processing for the second StringFormat object.
' The code uses the second StringFormat object to draw a string that contains the hot key
' prefix character. The code also draws the string's layout rectangle.
' ========================================================================================
SUB Example_GetHotKeyPrefix (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a red solid brush
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   ' // Create a font family from name
   DIM fontFamily AS CGpFontFamily = "Times New Roman"
   ' // Create a font from the font family
   DIM pFont AS CGpFont = CGpFont(@fontFamily, 24, FontStyleRegular, UnitPixel)

   ' // Create a string format object and set its hot key prefix
   DIM stringFormat AS CGpStringFormat
   stringFormat.SetHotkeyPrefix(HotkeyPrefixShow)

   ' // Get the hot key prefix from the StringFormat object.
   DIM nHotkeyPrefix AS HotkeyPrefix = stringFormat.GetHotkeyPrefix

   ' // Create a second StringFormat object with the same hot key prefix.
   DIM stringFormat2 AS CGpStringFormat
   stringFormat2.SetHotkeyPrefix(nHotkeyPrefix)

   ' // Use the seconds StringFormat object in a call to DrawString
   DIM wszText AS WSTRING * 260 = "This &text has some &underlined characters."
   graphics.DrawString(@wszText, LEN(wszText), @pFont, 30, 30, 160, 200, @stringFormat2, @solidBrush)

   ' // Draw the rectangle that encloses the text
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawRectangle(@pen, 30, 30, 160, 200)

END SUB
' ========================================================================================
```

# <a name="GetLineAlignment"></a>GetLineAlignment

Gets an element of the **StringAlignment** enumeration that indicates the line alignment of this **StringFormat** object in relation to the origin of the layout rectangle. The line alignment setting specifies how to align the string vertically in the layout rectangle. The layout rectangle is used to position the displayed string.

```
FUNCTION GetLineAlignment () AS StringAlignment
```

#### Example

```
' ========================================================================================
' The following example creates a StringFormat object, sets the string's line alignment,
' and then gets the line alignment setting and stores it in a variable. The code then
' creates a second StringFormat object and uses the stored alignment value to set the line
' alignment of the second StringFormat object. Next, the code uses the second StringFormat
' object to draw a formatted string. The code also draws the string's layout rectangle.
' ========================================================================================
SUB Example_GetLineAlignment (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a red solid brush
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   ' // Create a font family from name
   DIM fontFamily AS CGpFontFamily = "Times New Roman"
   ' // Create a font from the font family
   DIM pFont AS CGpFont = CGpFont(@fontFamily, 24, FontStyleRegular, UnitPixel)

   ' // Create a string format object and set the alignment
   DIM stringFormat AS CGpStringFormat
   stringFormat.SetLineAlignment(StringAlignmentCenter)

   ' // Get the line alignment setting from the StringFormat object.
   DIM nStringAlignment AS StringAlignment = stringFormat.GetLineAlignment

   ' // Create a second StringFormat object with the same line alignment.
   DIM stringFormat2 AS CGpStringFormat
   stringFormat2.SetLineAlignment(nStringAlignment)

   ' // Use the second StringFormat object in a call to DrawString
   DIM wszText AS WSTRING * 260 = "This text was formatted by a second StringFormat object."
   graphics.DrawString(@wszText, LEN(wszText), @pFont, 30, 30, 150, 200, @stringFormat2, @solidBrush)

   ' // Draw the rectangle that encloses the text
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawRectangle(@pen, 30, 30, 150, 200)

END SUB
' ========================================================================================
```

# <a name="GetMeasurableCharacterRangeCount"></a>GetMeasurableCharacterRangeCount

Gets the number of measurable character ranges that are currently set. The character ranges that are set can be measured in a string by using the **MeasureCharacterRanges** method.

```
FUNCTION GetMeasurableCharacterRangeCount () AS INT_
```

#### Return value

This method returns an integer that indicates the number of character ranges that can be measured by **MeasureCharacterRanges**.

#### Example

```
' ========================================================================================
' The following example defines three ranges of character positions within a string and sets
' those ranges in a StringFormat object. Next, the StringFormat::GetMeasurableCharacterRangeCount
' method is used to get the number of character ranges that are currently set in the StringFormat
' object. This number is then used to allocate a buffer large enough to store the regions
' that correspond with the ranges. Then, the MeasureCharacterRanges method is used to get
' the three regions of the display that are occupied by the characters that are specified
' by the ranges.
' Remarks: It doesn't work with the 64-bit headers because they lack a declare for the
' GdipMeasureCharacterRanges function.
' ========================================================================================
SUB Example_GetMeasurableCharacterRanges (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Brushes and pens used for drawing and painting
   DIM blueBrush AS CGpSOlidBrush = GDIP_ARGB(255, 0, 0, 255)
   DIM redBrush AS CGpSOlidBrush = GDIP_ARGB(255, 255, 0, 0)
   DIM blackPen AS CGpPen = GDIP_ARGB(255, 0, 0, 0)

   ' // Layout rectangles used for drawing strings
   DIM layoutRect AS GpRectF = TYPE<GpRectF>(20.0, 20.0, 130.0, 130.0)

   ' // Three ranges of character positions within the string
   DIM charRanges(2) AS CharacterRange
   charRanges(0).First = 3  : charRanges(0).Length = 5
   charRanges(1).First = 15 : charRanges(1).Length = 2
   charRanges(2).First = 30 : charRanges(2).Length = 15

   ' // Font and string format used to apply to string when drawing
   DIM myFont AS CGpFont = CGpFont("Times New Roman", AfxPointsToPixelsX(16) / rxRatio, FontStyleRegular, UnitPixel)
   DIM strFormat AS CGpStringFormat

   DIM wszText AS WSTRING * 260
   wszText = "The quick, brown fox easily jumps over the lazy dog."

   ' // Set three ranges of character positions.
   strFormat.SetMeasurableCharacterRanges(3, @charRanges(0))

   ' // Get the number of ranges that have been set, and allocate memory to
   ' // store the regions that correspond to the ranges.
   DIM nCount AS LONG = strFormat.GetMeasurableCharacterRangeCount
   DIM rgCharRangeRegions(nCount - 1) AS CGpRegion

   ' // Get the regions that correspond to the ranges within the string
   graphics.MeasureCharacterRanges(@wszText, -1, @myFont, @layoutRect, @strFormat, nCount, @rgCharRangeRegions(0))
   graphics.DrawString(@wszText, -1, @myFont, @layoutRect, @strFormat, @blueBrush)
   graphics.DrawRectangle(@blackPen, @layoutRect)
   FOR i AS LONG = 0 TO nCount - 1
      graphics.FillRegion(@redBrush, @rgCharRangeRegions(i))
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetTabStopCount"></a>GetTabStopCount

Gets the number of tab-stop offsets in this **StringFormat** object.

```
FUNCTION GetTabStopCount () AS INT_
```

#### Example

```
' ========================================================================================
' The following example creates a StringFormat object, sets tab stops, and uses the StringFormat
' object to draw a string that contains tab characters (\t). The code also draws the string's
' layout rectangle. Then, the code gets the tab stops and proceeds to use or inspect the
' tab stops in some way.
' ========================================================================================
SUB Example_GetTabStopCount (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   ' // Create a red solid brush
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 255)
   ' // Create a font family from name
   DIM fontFamily AS CGpFontFamily = "Times New Roman"
   ' // Create a font from the font family
   DIM pFont AS CGpFont = CGpFont(@fontFamily, 16, FontStyleRegular, UnitPixel)

   ' // Create a string format object and set the tab stops
   DIM tabs(0 TO 2) AS SINGLE = {150, 100, 100}
   DIM stringFormat AS CGpStringFormat
   stringFormat.SetTabStops(0, 3, @tabs(0))

   ' // Use the StringFormat object in a call to DrawString
   DIM wszText AS WSTRING * 260 = "Name" & CHR(9) &"Test 1" & CHR(9) & "Test 2" & CHR(9) & "Test 3"
   graphics.DrawString(@wszText, LEN(wszText), @pFont, 20, 20, 500, 100, @stringFormat, @solidBrush)

   ' // Get the tab stops
   DIM tabStopCount AS LONG, firstTabOffset AS SINGLE
   DIM tabStops(ANY) AS SINGLE
   tabStopCount = stringFormat.GetTabStopCount
   REDIM tabStops(tabStopCount - 1)
   stringFormat.GetTabStops(tabStopCount, @firstTabOffset, @tabStops(0))

   FOR j AS LONG = 0 TO tabStopCount - 1
      ' // Inspect or use the value in tabStops(j)
      PRINT tabStops(j)
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetTabStops"></a>GetTabStops

Gets the offsets of the tab stops in this **StringFormat** object.

```
FUNCTION GetTabStops (BYVAL count AS INT_, BYVAL firstTabOffset AS SINGLE PTR, _
   BYVAL tabStops AS SINGLE PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *count* | Integer that specifies the number of tab-stop offsets in the *tabStops* array. |
| *firstTabOffset* | Out. Pointer to a simple precision variable that receives the initial offset position. This initial offset position is relative to the string's origin and the offset of the first tab stop is relative to the initial offset position. |
| *tabStops* | Out. Pointer to an array of simple precision variables that receives the tab-stop offsets. The offset of the first tab stop is the first value in the array, the offset of the second tab stop, the second value in the array, and so on. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a StringFormat object, sets tab stops, and uses the
' StringFormat object to draw a string that contains tab characters. The code also draws
' the string's layout rectangle. Then, the code gets the tab stops and proceeds to use or
' inspect the tab stops in some way.
' ========================================================================================
SUB Example_GetTabStops (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   ' // Create a red solid brush
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 255)
   ' // Create a font family from name
   DIM fontFamily AS CGpFontFamily = "Times New Roman"
   ' // Create a font from the font family
   DIM pFont AS CGpFont = CGpFont(@fontFamily, 16, FontStyleRegular, UnitPixel)

   ' // Create a string format object and set the tab stops
   DIM tabs(0 TO 2) AS SINGLE = {150, 100, 100}
   DIM stringFormat AS CGpStringFormat
   stringFormat.SetTabStops(0, 3, @tabs(0))

   ' // Use the StringFormat object in a call to DrawString
   DIM wszText AS WSTRING * 260 = "Name" & CHR(9) &"Test 1" & CHR(9) & "Test 2" & CHR(9) & "Test 3"
   graphics.DrawString(@wszText, LEN(wszText), @pFont, 20, 20, 500, 100, @stringFormat, @solidBrush)

   ' // Draw the rectangle that encloses the text.
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 255, 0, 0))
   graphics.DrawRectangle(@pen, 20, 20, 500, 100)

   ' // Get the tab stops
   DIM tabStopCount AS LONG, firstTabOffset AS SINGLE
   DIM tabStops(ANY) AS SINGLE
   tabStopCount = stringFormat.GetTabStopCount
   REDIM tabStops(tabStopCount - 1)
   stringFormat.GetTabStops(tabStopCount, @firstTabOffset, @tabStops(0))

   FOR j AS LONG = 0 TO tabStopCount - 1
      ' // Inspect or use the value in tabStops(j)
      PRINT tabStops(j)
   NEXT

END SUB
' ========================================================================================
```

# <a name="GetTrimming"></a>GetTrimming

Gets an element of the **StringTrimming** enumeration that indicates the trimming style of this **StringFormat** object. The trimming style determines how to trim characters from a string that is too large to fit in the layout rectangle. 

```
FUNCTION GetTrimming () AS StringTrimming
```

#### Example

```
' ========================================================================================
' The following example creates a StringFormat object, sets the string's trimming style,
' and then gets the trimming style and stores it in a variable. The code then creates a
' second StringFormat object and uses the stored trimming style to set the trimming style
' of the second StringFormat object. Next, the code uses the second StringFormat object to
' draw a formatted string. The code also draws the string's layout rectangle.
' ========================================================================================
SUB Example_GetTrimming (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a red solid brush
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   ' // Create a font family from name
   DIM fontFamily AS CGpFontFamily = "Times New Roman"
   ' // Create a font from the font family
   DIM pFont AS CGpFont = CGpFont(@fontFamily, 24, FontStyleRegular, UnitPixel)

   ' // Create a string format object and set its trimming style
   DIM stringFormat AS CGpStringFormat
   stringFormat.SetTrimming(StringTrimmingEllipsisWord)

   ' // Get the trimming style from the StringFormat object.
   DIM nStringTrimming AS StringTrimming = stringFormat.GetTrimming

   ' // Create a second StringFormat object with the same trimming style.
   DIM stringFormat2 AS CGpStringFormat
   stringFormat2.SetTrimming(nStringTrimming)

   ' // Use the second StringFormat object in a call to DrawString.
   DIM wszText AS WSTRING * 260 = "One two three four five six seven eight nine ten"
   graphics.DrawString(@wszText, LEN(wszText), @pFont, 30, 30, 160, 60, @stringFormat, @solidBrush)

   ' // Draw the rectangle that encloses the text
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawRectangle(@pen, 30, 30, 160, 60)

END SUB
' ========================================================================================
```

# <a name="SetAlignment"></a>SetAlignment

Sets the character alignment of this **StringFormat** object in relation to the origin of the layout rectangle. A layout rectangle is used to position the displayed string.

```
FUNCTION SetAlignment (BYVAL nAlign AS StringAlignment) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nAlign* | Element of the **StringAlignment** enumeration that specifies how a string is aligned in reference to the bounding rectangle. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a StringFormat object, sets the character alignment, and
' draws the string using that alignment. The code also draws the string's layout rectangle.
' ========================================================================================
SUB Example_SetAlignment (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a red solid brush
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   ' // Create a font family from name
   DIM fontFamily AS CGpFontFamily = "Times New Roman"
   ' // Create a font from the font family
   DIM pFont AS CGpFont = CGpFont(@fontFamily, 24, FontStyleRegular, UnitPixel)

   ' // Create a string format object and set the alignment
   DIM stringFormat AS CGpStringFormat
   stringFormat.SetAlignment(StringAlignmentFar)

   ' // Use the StringFormat object in a call to DrawString
   DIM wszText AS WSTRING * 260 = "This text was formatted by a StringFormat object."
   graphics.DrawString(@wszText, LEN(wszText), @pFont, 30, 30, 150, 200, @stringFormat, @solidBrush)

   ' // Draw the rectangle that encloses the text
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawRectangle(@pen, 30, 30, 150, 200)

END SUB
' ========================================================================================
```

# <a name="SetDigitSubstitution"></a>SetDigitSubstitution

Sets the digit substitution method and the language that corresponds to the digit substitutes.

```
FUNCTION SetDigitSubstitution (BYVAL language AS LANGID, BYVAL substitute AS StringDigitSubstitute) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *language* | Sixteen-bit value that forms a NLS language identifier. The identifier specifies the language associated with the substitute digits. For example, if this **StringFormat** object uses Arabic substitution digits, then this method will return a value that indicates an Arabic language. An NLS language identifier is constructed by the **MAKELANGID** macro, declared in Winnt.inc. |
| *substitute* | Element of the **StringDigitSubstitute** enumeration that specifies the digit substitution method to be used. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

The digit substitution method, specified by an element of the **StringDigitSubstitute** enumeration, replaces, in a string, Western European digits with digits that correspond to a user's locale or language.

When specifying LANG_NEUTRAL as the language ID, it is common practice to pass just LANG_NEUTRAL as in the following example:

stat = FontFamily.GetFamilyName(name, LANG_NEUTRAL)

If you are specifying a language other than LANG_NEUTRAL, use MAKELANGID to create the language and sublanguage combination as in the following example:

language = MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_TRADITIONAL)

For a list of the available languages and sublanguages, see Winnt.inc.

# <a name="SetFormatFlags"></a>SetFormatFlags

Sets the format flags for this **StringFormat** object. The format flags determine most of the characteristics of a **StringFormat** object.

```
FUNCTION SetFormatFlags (BYVAL flags AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *flags* | Thirty-two bit value that specifies the format flags that control most of the characteristics of the **StringFormat** object. The flags are set by applying a bitwise OR to elements of the **StringFormatFlags** enumeration. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a StringFormat object, sets the format flags, and draws
' the formatted string. The code also draws the string's layout rectangle.
' ========================================================================================
SUB Example_SetFormatFlags (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a red solid brush
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   ' // Create a font family from name
   DIM fontFamily AS CGpFontFamily = "Times New Roman"
   ' // Create a font from the font family
   DIM pFont AS CGpFont = CGpFont(@fontFamily, 24, FontStyleRegular, UnitPixel)

   ' // Create a string format object and set its format flags
   DIM stringFormat AS CGpStringFormat
   stringFormat.SetFormatFlags(StringFormatFlagsDirectionVertical OR StringFormatFlagsNoFitBlackBox)

   ' // Use the StringFormat object in a call to DrawString
   DIM wszText AS WSTRING * 260 = "This text is vertical because of a format flag."
   graphics.DrawString(@wszText, LEN(wszText), @pFont, 30, 30, 150, 200, @stringFormat, @solidBrush)

   ' // Draw the rectangle that encloses the text
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawRectangle(@pen, 30, 30, 150, 200)

END SUB
' ========================================================================================
```

# <a name="SetHotkeyPrefix"></a>SetHotkeyPrefix

Sets the type of processing that is performed on a string when the hot key prefix, an ampersand (&), is encountered. The ampersand is called the hot key prefix and can be used to designate certain keys as hot keys.

```
FUNCTION SetHotkeyPrefix (BYVAL nHotkeyPrefix AS HotkeyPrefix) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *flags* | Element of the **HotkeyPrefix** enumeration that specifies how to process the hot key prefix. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a StringFormat object and sets the type of hot key prefix
' processing to be performed on the string. The code then uses the StringFormat object to
' draw a string that contains the hot key prefix character. The code also draws the
' string's layout rectangle.
' ========================================================================================
SUB Example_SetHotKeyPrefix (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a red solid brush
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   ' // Create a font family from name
   DIM fontFamily AS CGpFontFamily = "Times New Roman"
   ' // Create a font from the font family
   DIM pFont AS CGpFont = CGpFont(@fontFamily, 24, FontStyleRegular, UnitPixel)

   ' // Create a string format object and set its hot key prefix
   DIM stringFormat AS CGpStringFormat
   stringFormat.SetHotkeyPrefix(HotkeyPrefixShow)

   ' // Use the StringFormat object in a call to DrawString
   DIM wszText AS WSTRING * 260 = "This &text has some &underlined characters."
   graphics.DrawString(@wszText, LEN(wszText), @pFont, 30, 30, 150, 200, @stringFormat, @solidBrush)

   ' // Draw the rectangle that encloses the text
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawRectangle(@pen, 30, 30, 150, 200)

END SUB
' ========================================================================================
```

# <a name="SetLineAlignment"></a>SetLineAlignment

Sets the line alignment of this **StringFormat** object in relation to the origin of the layout rectangle. The line alignment setting specifies how to align the string vertically in the layout rectangle. The layout rectangle is used to position the displayed string.

```
FUNCTION SetLineAlignment (BYVAL nAlign AS StringAlignment) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nAlign* | Element of the **StringAlignment** enumeration that specifies how to align a string in reference to the layout rectangle. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a StringFormat object, sets the trimming style, and uses
' the StringFormat object to draw a string. The code also draws the string's layout rectangle.
' ========================================================================================
SUB Example_SetLineAlignment (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a red solid brush
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   ' // Create a font family from name
   DIM fontFamily AS CGpFontFamily = "Times New Roman"
   ' // Create a font from the font family
   DIM pFont AS CGpFont = CGpFont(@fontFamily, 24, FontStyleRegular, UnitPixel)

   ' // Create a string format object and set the alignment
   DIM stringFormat AS CGpStringFormat
   stringFormat.SetLineAlignment(StringAlignmentCenter)

   ' // Use the StringFormat object in a call to DrawString
   DIM wszText AS WSTRING * 260 = "This text was formatted by a StringFormat object."
   graphics.DrawString(@wszText, LEN(wszText), @pFont, 30, 30, 150, 200, @stringFormat, @solidBrush)

   ' // Draw the rectangle that encloses the text
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawRectangle(@pen, 30, 30, 150, 200)

END SUB
' ========================================================================================
```

# <a name="SetMeasurableCharacterRanges"></a>SetMeasurableCharacterRanges

Sets a series of character ranges for this **StringFormat** object that, when in a string, can be measured by the **MeasureCharacterRanges** method.

```
FUNCTION SetMeasurableCharacterRanges (BYVAL rangeCount AS INT_, _
   BYVAL ranges AS CharacterRange PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *rangeCount* | Integer that specifies the number of character ranges in the *ranges* array. |
| *ranges* | Pointer to an array of **CharacterRange** objects that specify the character ranges to be measured. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a StringFormat object, sets tab stops, and uses the
' StringFormat object to draw a string that contains tab characters (\t). The code also
' draws the string's layout rectangle.
' Remarks: It doesn't work with the 64-bit headers because they lack a declare for the
' GdipMeasureCharacterRanges function.
' ========================================================================================
SUB Example_SetMeasurableCharacterRanges (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Brushes and pens used for drawing and painting
   DIM blueBrush AS CGpSOlidBrush = GDIP_ARGB(255, 0, 0, 255)
   DIM redBrush AS CGpSOlidBrush = GDIP_ARGB(255, 255, 0, 0)
   DIM blackPen AS CGpPen = GDIP_ARGB(255, 0, 0, 0)

   ' // Layout rectangles used for drawing strings
   DIM layoutRect_A AS GpRectF = TYPE<GpRectF>(20.0, 20.0, 130.0, 130.0)
   DIM layoutRect_B AS GpRectF = TYPE<GpRectF>(160.0, 20.0, 165.0, 130.0)
   DIM layoutRect_C AS GpRectF = TYPE<GpRectF>(335.0, 20.0, 165.0, 130.0)

   ' // Three ranges of character positions within the string
   DIM charRanges(2) AS CharacterRange
   charRanges(0).First = 3  : charRanges(0).Length = 5
   charRanges(1).First = 15 : charRanges(1).Length = 2
   charRanges(2).First = 30 : charRanges(2).Length = 15

   ' // Font and string format used to apply to string when drawing
   DIM myFont AS CGpFont = CGpFont("Times New Roman", AfxPointsToPixelsX(16) / rxRatio, FontStyleRegular, UnitPixel)
   DIM strFormat AS CGpStringFormat

   DIM wszText AS WSTRING * 260
   wszText = "The quick, brown fox easily jumps over the lazy dog."

   ' // Set three ranges of character positions.
   strFormat.SetMeasurableCharacterRanges(3, @charRanges(0))

   ' // Get the number of ranges that have been set, and allocate memory to
   ' // store the regions that correspond to the ranges.
   DIM nCount AS LONG = strFormat.GetMeasurableCharacterRangeCount
   DIM rgCharRangeRegions(nCount - 1) AS CGpRegion

   ' // Get the regions that correspond to the ranges within the string when
   ' // layout rectangle A is used. Then draw the string, and show the regions.
   graphics.MeasureCharacterRanges(@wszText, -1, @myFont, @layoutRect_A, @strFormat, nCount, @rgCharRangeRegions(0))
   graphics.DrawString(@wszText, -1, @myFont, @layoutRect_A, @strFormat, @blueBrush)
   graphics.DrawRectangle(@blackPen, @layoutRect_A)
   FOR i AS LONG = 0 TO nCount - 1
      graphics.FillRegion(@redBrush, @rgCharRangeRegions(i))
   NEXT

   ' // Get the regions that correspond to the ranges within the string when
   ' // layout rectangle B is used. Then draw the string, and show the regions.
   graphics.MeasureCharacterRanges(@wszText, -1, @myFont, @layoutRect_B, @strFormat, nCount, @rgCharRangeRegions(0))
   graphics.DrawString(@wszText, -1, @myFont, @layoutRect_B, @strFormat, @blueBrush)
   graphics.DrawRectangle(@blackPen, @layoutRect_B)
   FOR i AS LONG = 0 TO nCount - 1
      graphics.FillRegion(@redBrush, @rgCharRangeRegions(i))
   NEXT

   ' // Get the regions that correspond to the ranges within the string when
   ' // layout rectangle C is used. Set trailing spaces to be included in the
   ' // regions. Then draw the string, and show the regions.
   strFormat.SetFormatFlags(StringFormatFlagsMeasureTrailingSpaces)
   graphics.MeasureCharacterRanges(@wszText, -1, @myFont, @layoutRect_C, @strFormat, nCount, @rgCharRangeRegions(0))
   graphics.DrawString(@wszText, -1, @myFont, @layoutRect_C, @strFormat, @blueBrush)
   graphics.DrawRectangle(@blackPen, @layoutRect_C)
   FOR i AS LONG = 0 TO nCount - 1
      graphics.FillRegion(@redBrush, @rgCharRangeRegions(i))
   NEXT

END SUB
' ========================================================================================
```

# <a name="SetTabStops"></a>SetTabStops

Sets the offsets for tab stops in this **StringFormat** object.

```
FUNCTION SetTabStops (BYVAL firstTabOffset AS SINGLE, BYVAL nCount AS INT_, _
   BYVAL tabStops AS SINGLE PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *firstTabOffset* | Simple precision number that specifies the initial offset position. This initial offset position is relative to the string's origin and the offset of the first tab stop is relative to the initial offset position.  |
| *nCount* | Integer that specifies the number of tab-stop offsets in the *tabStops* array. |
| *tabStops* | Pointer to an array of real numbers that specify the tab-stop offsets. The offset of the first tab stop is the first value in the array, the offset of the second tab stop, the second value in the array, and so on. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks

Each tab-stop offset in the tabStops array, except the first one, is relative to the previous one. The first tab-stop offset is relative to the initial offset position specified by *firstTabOffset*. For example, if the initial offset position is 8 and the first tab-stop offset is 50, then the first tab stop is at position 58. If the initial offset position is zero, then the first tab-stop offset is relative to position 0, the string origin.

#### Example

```
' ========================================================================================
' The following example creates a StringFormat object, sets tab stops, and uses the
' StringFormat object to draw a string that contains tab characters (\t). The code also
' draws the string's layout rectangle.
' ========================================================================================
SUB Example_SetTabStops (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   ' // Create a red solid brush
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 0, 0, 255)
   ' // Create a font family from name
   DIM fontFamily AS CGpFontFamily = "Times New Roman"
   ' // Create a font from the font family
   DIM pFont AS CGpFont = CGpFont(@fontFamily, 16, FontStyleRegular, UnitPixel)

   ' // Create a string format object and set the tab stops
   DIM tabs(0 TO 2) AS SINGLE = {150, 100, 100}
   DIM stringFormat AS CGpStringFormat
   stringFormat.SetTabStops(0, 3, @tabs(0))

   ' // Use the StringFormat object in a call to DrawString
   DIM wszText AS WSTRING * 260 = "Name" & CHR(9) &"Test 1" & CHR(9) & "Test 2" & CHR(9) & "Test 3"
   graphics.DrawString(@wszText, LEN(wszText), @pFont, 20, 20, 500, 100, @stringFormat, @solidBrush)

   ' // Draw the rectangle that encloses the text
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawRectangle(@pen, 20, 20, 500, 100)

END SUB
' ========================================================================================
```

# <a name="SetTrimming"></a>SetTrimming

Sets the trimming style for this **StringFormat** object. The trimming style determines how to trim a string so that it fits into the layout rectangle.

```
FUNCTION SetTrimming (BYVAL trimming AS StringTrimming) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *trimming* | Element of the **StringTrimming** enumeration that specifies the style of trimming to be performed on the string. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a StringFormat object, sets the trimming style, and uses
' the StringFormat object to draw a string. The code also draws the string's layout rectangle.
' ========================================================================================
SUB Example_SetTrimming (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, rxRatio)

   ' // Create a red solid brush
   DIM solidBrush AS CGpSolidBrush = GDIP_ARGB(255, 255, 0, 0)
   ' // Create a font family from name
   DIM fontFamily AS CGpFontFamily = "Times New Roman"
   ' // Create a font from the font family
   DIM pFont AS CGpFont = CGpFont(@fontFamily, 24, FontStyleRegular, UnitPixel)

   ' // Create a string format object and set its trimming style
   DIM stringFormat AS CGpStringFormat
   stringFormat.SetTrimming(StringTrimmingEllipsisWord)

   ' // Use the StringFormat object in a call to DrawString
   DIM wszText AS WSTRING * 260 = "One two three four five six seven eight nine ten"
   graphics.DrawString(@wszText, LEN(wszText), @pFont, 30, 30, 160, 60, @stringFormat, @solidBrush)

   ' // Draw the rectangle that encloses the text
   DIM pen AS CGpPen = GDIP_ARGB(255, 255, 0, 0)
   graphics.DrawRectangle(@pen, 30, 30, 160, 60)

END SUB
' ========================================================================================
```
