# Fonts

Internally, the visual designer handles creating and assigning fonts tot he various controls being used in your application. However, there may come a time when you need to create and assign fonts via your code. WinFBE creates fonts that are high dpi aware meaning that they scale correctly even on monitors that have very high resolutions.

#### Creating and assigning a new font

```
Form1.CheckBox1.Font = New wfxFont("Segoe UI", 9, FontStyles.Bold, FontCharset.Ansi)
```

#### Parameters  
**FontName**

**FontSize**

**FontStyle** (use **OR** to combine multiple styles)
​	FontStyles.Normal
​	FontStyles.Bold
​	FontStyles.Italic
​	FontStyles.Strikeout
​	FontStyles.Underline

**CharacterSet**
​	FontCharset.Default
​	FontCharset.Ansi
​	FontCharset.Arabic
​	FontCharset.Baltic
​	FontCharset.ChineseBig5
​	FontCharset.EastEurope
​	FontCharset.GB2312
​	FontCharset.Greek
​	FontCharset.Hangul
​	FontCharset.Hebrew
​	FontCharset.Johab
​	FontCharset.Mac
​	FontCharset.OEM
​	FontCharset.Russian
​	FontCharset.Shiftjis
​	FontCharset.Symbol
​	FontCharset.Thai
​	FontCharset.Turkish
​	FontCharset.Vietnamese

