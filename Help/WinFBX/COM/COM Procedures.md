# Helper COM Procedures

Assorted COM procedures.

**Include file**: AfxCOM.inc

| Name       | Description |
| ---------- | ----------- |
| [AfxAcode](#AfxAcode) | Translates unicode bytes to ansi bytes. |
| [AfxUcode](#AfxUcode) | Translates ansi bytes to unicode bytes. |
| [AfxIsBstr](#AfxIsBstr) | Checks if the passed pointer is a BSTR. |
| [AfxNewCOM(PROGID](#AfxNewCOM1) | Creates a single uninitialized object of the class associated with a specified ProgID or CLSID. |
| [AfxNewCOM(CLSID)](#AfxNewCOM2) | Creates a single uninitialized object of the class associated with a specified CLSID. |
| [AfxNewCOM(CLSID,IID)](#AfxNewCOM3) | Creates a single uninitialized object of the class associated with the specified CLSID and IID. |
| [AfxNewCOM(LibName)](#AfxNewCOM4) | Loads the specified library from file and creates an instance of an object. |
| [AfxGetCOM](#AfxGetCOM) | Returns a dispatch pointer to a running out-of-process server that is registered in the ROT. |
| [AfxAnyCOM](#AfxAnyCOM) | Tries to use an existing, running application if available, or creates a new instance if not. |
| [AfxGuid](#AfxGuid) | Converts a string into a 16-byte (128-bit) Globally Unique Identifier (GUID). |
| [AfxGuidText](#AfxGuidText) | Returns a 38-byte human-readable guid string from a 16-byte GUID. |
| [AfxSafeAddRef](#AfxSafeAddRef) | Increments the reference count for an interface on an object. |
| [AfxSafeRelease](#AfxSafeRelease) | Decrements the reference count for an interface on an object. |
| [AfxAdvise](#AfxAdvise) | Establishes a connection between the connection point object and the client's sink. |
| [AfxUnadvise](#AfxUnadvise) | Releases the events connection identified with the cookie returned by the Advise method. |
| [AfxWstrAlloc](#AfxWstrAlloc) | Takes a null terminated wide string as input, and returns a pointer to a new wide string allocated with CoTaskMemAlloc. |
| [AfxGetOleErrorInfo](#AfxGetOleErrorInfo) | Returns the description of the most recent OLE error in the current logical thread and clears the error state for the thread. |
| [AfxOleCreateFont](#AfxOleCreateFont) | Creates a standard IFont object. |
| [AfxOleCreateFontDisp](#AfxOleCreateFontDisp) | Creates a standard IFontDisp object. |
| [AfxGetVarType](#AfxGetVarType) | Returns the type of a variant. |
| [AfxIsVarTypeFloat](#AfxIsVarTypeFloat) | Checks if a variant contains a float value. |
| [AfxIsVarTypeSignedInteger](#AfxIsVarTypeSignedInteger) | Checks if a variant contains a signed integer. |
| [AfxIsVarTypeUnsignedInteger](#AfxIsVarTypeUnsignedInteger) | Checks if a variant contains an unsigned integer. |
| [AfxIsVarTypeInteger](#AfxIsVarTypeInteger) | Checks if a variant contains an integer. |
| [AfxIsVarTypeNumber](#AfxIsVarTypeNumber) | Checks if a variant contains a number. |
| [AfxIsVariantString](#AfxIsVariantString) | Checks if a variant contains an string. |
| [AfxIsVariantArray](#AfxIsVariantArray) | Checks if a variant contains an array. |
| [AfxVarToStr](#AfxVarToStr) | Extracts the contents of a VARIANT and returns them as a CWSTR. |
| [AfxVariantGetElementCount](#AfxVariantGetElementCount) | Retrieves the element count of a variant structure. |
| [AfxVariantToBuffer](#AfxVariantToBuffer) | Extracts the contents of a variant that contains an array of bytes. |
| [AfxVariantToString](#AfxVariantToString) | Extracts the variant value of a variant structure to a string. |
| [AfxVariantToStringAlloc](#AfxVariantToStringAlloc) | Extracts the variant value of a variant structure to a string. |

# <a name="AfxAcode"></a>AfxAcode

Translates unicode bytes to ansi bytes.

```
FUNCTION AfxAcode (BYVAL pwszStr AS WSTRING PTR, BYVAL nCodePage AS LONG = 0) AS STRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszStr* | The WSTRING or CWSTR to convert. |
| *nCodePage* | Optional. The code page used in the conversion, e.g. 1251 for Russian. If you specify CP_UTF8, the returned string will be UTF8 encoded. If you don't pass an unicode page, the function will use CP_ACP (0), which is the system default Windows ANSI code page. |

#### Return value

An ansi or UTF8 encoded string.

#### Usage example
(Russian bytes to unicode string:

```
DIM cws AS CWSTR
cws = AfxUcode(CHR(209, 229, 236, 229, 237), 1251)
MessageBoxW 0, cws, "", MB_OK
DIM s AS STRING
s = AfxAcode(cws, 1251)
MessageBoxW 0, s, "", MB_OK
```


# <a name="AfxUcode"></a>AfxUcode

Translates ansi bytes to unicode bytes. This function works like the intrinsic Free Basic WSTR function, but accepting an optional code page.

```
FUNCTION AfxUcode (BYREF ansiStr AS CONST STRING, BYVAL nCodePage AS LONG = 0) AS CWSTR
```
| Parameter  | Description |
| ---------- | ----------- |
| *ansiStr* | Ansi or UTF8 string to convert. |
| *nCodePage* | Optional. The code page used in the conversion, e.g. 1251 for Russian. If you specify CP_UTF8, it is assumed that ansiStr contains an UTF8 encoded string. If you don't pass an unicode page, the function will use CP_ACP (0), which is the system default Windows ANSI code page. |

#### Return value

The ansi or utf-8 string converted to utf-16.

#### Usage example
(Russian bytes to unicode string and then to ansi):

```
DIM cws AS CWSTR
cws = AfxUcode(CHR(209, 229, 236, 229, 237), 1251)
MessageBoxW 0, cws, "", MB_OK
DIM s AS STRING
s = AfxAcode(cws, 1251)
MessageBoxW 0, s, "", MB_OK```
```

# <a name="AfxIsBstr"></a>AfxIsBstr

Checks if the passed pointer is a BSTR.

```
FUNCTION AfxIsBstr (BYVAL pv AS ANY PTR) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *pv* | Pointer to the string. |

#### Remarks

Will return FALSE if it is a null pointer.
If it is an OLE string it must have a descriptor; otherwise, don't.
Gets the length in bytes looking at the descriptor and divides by 2 to get the number of unicode characters, that is the value returned by the FreeBASIC LEN operator. If the retrieved length if the same that the returned by LEN, then it must be an OLE string.

# <a name="AfxNewCOM1"></a>AfxNewCOM (PROGID)

Creates a single uninitialized object of the class associated with a specified ProgID or CLSID.

```
FUNCTION AfxNewCom (BYREF wszProgID AS CONST WSTRING, BYREF wszLicKey AS WSTRING = "") AS ANY PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszProgID* | The ProgID or the CLSID of the object to create.<br>A ProgID such as "MSCAL.Calendar.7"<br>A CLSID such as "{8E27C92B-1264-101C-8A2F-040224009C02}" |
| *wszLicKey* | Optional. The license key as a unicode string. |

#### Return value

An interface pointer or NULL.

#### Usage examples

```
DIM pDic AS IDictionary PTR
pDic = AfxNewCom("Scripting.Dictionary")
-or-
pDic = AfxNewCom(CLSID_Dictionary)
```
where CLSID_Dictionary has been declared as CONST CLSID_Dictionary = "{EE09B103-97E0-11CF-978F-00A02463E06F}"

# <a name="AfxNewCOM2"></a>AfxNewCOM (CLSID)

Creates a single uninitialized object of the class associated with a specified CLSID.

```
FUNCTION AfxNewCom (BYREF classID AS CONST CLSID) AS ANY PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *classID* | The CLSID (class identifier) associated with the data and code that will be used to create the object. |

#### Return value

An interface pointer or NULL.

#### Usage example

```
DIM pDic AS IDictionary PTR
pDic = AfxNewCom(CLSID_Dictionary)
```

where CLSID_Dictionary has been declared as<br>
DIM CLSID_Dictionary AS CLSID = (&hEE09B103, &h97E0, &h11CF, {&h97, &h8F, &h00, &hA0, &h24, &h63, &hE0, &h6F})

# <a name="AfxNewCOM3"></a>AfxNewCOM (CLSID, IID)

Creates a single uninitialized object of the class associated with the specified CLSID and IID.

```
FUNCTION AfxNewCom (BYREF classID AS CONST CLSID, BYREF riid AS CONST IID) AS ANY PTR
FUNCTION AfxNewCom (BYREF wszClsID AS CONST WSTRING, BYREF wszIID AS CONST WSTRING) AS ANY PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *classID* | The CLSID (class identifier) associated with the data and code that will be used to create the object. |
| *riid* | A reference to the identifier of the interface to be used to communicate with the object. |
| *wszClsID* | The CLSID in string format. |
| *wszIID* | The IID in string format. |

#### Return value

An interface pointer or NULL.

#### Usage examples

```
DIM pDic AS IDictionary PTR
pDic = AfxNewCom(CLSID_Dictionary, IID_IDictionary)
```

where CLSID_Dictionary has been declared asCONST CLSID_Dictionary = "{EE09B103-97E0-11CF-978F-00A02463E06F}"<br>
and IID_IDictionary as CONST IID_IDictionary = "{42C642C1-97E1-11CF-978F-00A02463E06F}"

# <a name="AfxNewCOM4"></a>AfxNewCOM (LibName)

Loads the specified library from file and creates an instance of an object.

```
FUNCTION AfxNewCom (BYREF wszLibName AS CONST WSTRING, BYREF rclsid AS CONST CLSID, _
   BYREF riid AS CONST IID, BYREF wszLicKey AS WSTRING) AS ANY PTR
FUNCTION AfxNewCom (BYREF wszLibName AS CONST WSTRING, BYREF wszClsid AS CONST WSTRING, _
   BYREF riid AS CONST IID, BYREF wszLicKey AS WSTRING = "") AS ANY PTR
FUNCTION AfxNewCom (BYREF wszLibName AS CONST WSTRING, BYREF wszClsid AS CONST WSTRING, _
   BYREF wszIid AS CONST WSTRING, BYREF wszLicKey AS WSTRING = "") AS ANY PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszLibName* | Full path where the library is located. |
| *classID* | The CLSID (class identifier) associated with the data and code that will be used to create the object. |
| *riid* | A reference to the identifier of the interface to be used to communicate with the object. |
| *wszClsID* | The CLSID in string format. |
| *wszIID* | The IID in string format. |
| *wszLicKey* | Optional. The license key. |

#### Return value

An interface pointer or NULL.

#### Remarks

* Not every component is a suitable candidate for use under this overloaded **AfxNewCom** function.
* Only in-process servers (DLLs) are supported.
** Components that are system components or part of the operating system, such as XML, Data Access, Internet Explorer, or DirectX, aren't supported.
* Components that are part of an application, such Microsoft Office, aren't supported.
* Components intended for use as an add-in or a snap-in, such as an Office add-in or a control in a Web browser, aren't supported.
* Components that manage a shared physical or virtual system resource aren't supported.
* Visual ActiveX controls aren't supported because they need to be initilized and activated by the OLE container.

# <a name="AfxGetCOM"></a>AfxGetCOM

If the requested object is in an EXE (out-of-process server), such Office applications, and it is running and registered in the Running Object Table (ROT), **AfxGetCom** will return a pointer to its interface. Internally, **AfxGetCOM** calls the API function **GetActiveObject**.

```
FUNCTION AfxGetCom (BYREF wszProgID AS CONST WSTRING) AS IDispatch PTR
FUNCTION AfxGetCom (BYREF classID AS CONST CLSID) AS IDispatch PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszProgID* | The ProgID of the object to retrieve. |
| *classID* | The ClsID of the object to retrieve. |

Be aware that **AfxGetCom** can fail under if Office is running but not registered in the ROT.

When an Office application starts, it does not immediately register its running objects. This optimizes the application's startup process. Instead of registering at startup, an Office application registers its running objects in the ROT once it loses focus. Therefore, if you attempt to use **AfxGetCOM** to attach to a running instance of an Office application before the application has lost focus, you might receive an error.

See: [GetObject or GetActiveObject cannot find a running Office application](https://support.microsoft.com/en-us/help/238610/getobject-or-getactiveobject-cannot-find-a-running-office-application)

# <a name="AfxAnyCOM"></a>AfxAnyCOM

Tries to use an existing, running application if available, or creates a new instance if not.

```
FUNCTION AfxAnyCOM (BYREF wszProgID AS CONST WSTRING) AS IDispatch PTR
FUNCTION AfxAnyCOM (BYREF classID AS CONST CLSID) AS IDispatch PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszProgID* | The ProgID of the object to retrieve or create. |
| *classID* | The ClsID of the object to retrieve or create. |

# <a name="AfxGuid"></a>AfxGuid

Converts a string into a 16-byte (128-bit) Globally Unique Identifier (GUID). To be valid, the string must contain exactly 32 hexadecimal digits, delimited by hyphens and enclosed by curly braces. For example: {B09DE715-87C1-11D1-8BE3-0000F8754DA1}

If *pwszGuidText* is omited, **AfxGuid** generates a new unique guid.

```
FUNCTION AfxGuid (BYVAL pwszGuidText AS WSTRING PTR = NULL) AS GUID
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszGuidText* | A GUID in string format. |

#### Return value

A GUID.

# <a name="AfxGuidText"></a>AfxGuidText

Returns a 38-byte human-readable guid string from a 16-byte GUID.

```
FUNCTION AfxGuidText (BYVAL classID AS CLSID PTR) AS STRING
FUNCTION AfxGuidText (BYVAL classID AS CLSID) AS STRING
FUNCTION AfxGuidText (BYVAL riid AS REFIID) AS STRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *classID, riid* | The GUID to convert. |

#### Return value

A 38-byte human-readable guid string.

# <a name="AfxSafeAddRef"></a>AfxSafeAddRef

Increments the reference count for an interface on an object.

```
FUNCTION AfxSafeAddRef (BYVAL pv AS ANY PTR) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pv* | The COM interface pointer. |

#### Return value

The function returns the new reference count. This value is intended to be used only for test purposes.

#### Remarks

This method should be called for every new copy of a pointer to an interface on an object. For example, if you are passing a copy of a pointer back from a method, you must call **AfxSafeAddRef** on that pointer. You must also call **AfxSafeAddRef** on a pointer before passing it as an in-out parameter to a method; the method will call **IUnknown_Release** before copying the out-value on top of it.

Objects use a reference counting mechanism to ensure that the lifetime of the object includes the lifetime of references to it. You use AddRef to stabilize a copy of an interface pointer. It can also be called when the life of a cloned pointer must extend beyond the lifetime of the original pointer. The cloned pointer must be released by calling **AfxSafeRelease**.

# <a name="AfxSafeRelease"></a>AfxSafeRelease

Decrements the reference count for an interface on an object and sets the value of the passed pointer to NULL.

```
FUNCTION AfxSafeRelease (BYVAL pv AS ANY PTR) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pv* | The COM interface pointer to release. |

#### Return value

The function returns the new reference count. This value is intended to be used only for test purposes.

#### Remarks

When the reference count on an object reaches zero, **AfxSafeRelease** must cause the interface pointer to free itself. When the released pointer is the only existing reference to an object (whether the object supports single or multiple interfaces), the implementation must free the object.

# <a name="AfxAdvise"></a>AfxAdvise

Establishes a connection between the connection point object and the client's sink.

```
FUNCTION AfxAdvise (BYVAL pUnk AS ANY PTR, BYVAL pEvtObj AS ANY PTR, _
   BYVAL riid AS IID PTR, BYVAL pdwCookie AS DWORD PTR) AS HRESULT
FUNCTION AfxAdvise (BYVAL pUnk AS ANY PTR, BYVAL pEvtObj AS ANY PTR, _
   BYREF riid AS CONST IID, BYVAL pdwCookie AS DWORD PTR) AS HRESULT
FUNCTION AfxAdvise (BYVAL pUnk AS ANY PTR, BYVAL pEvtObj AS ANY PTR, _
   BYREF riid AS IID, BYVAL pdwCookie AS DWORD PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pUnk* | A pointer to the **IUnknown** of the object the client wants to connect with. |
| *pEvtObj* | A pointer to the client's **IUnknown**. |
| *riid* | The GUID of the connection point. Typically, this is the same as the outgoing interface managed by the connection point. |
| *pdwCookie* | Pointer to a DWORD variable that will receive the cookie that uniquely identifies the connection. |

#### Return value

S_OK (0) on success, or an HRESULT error code on failure.

# <a name="AfxUnadvise"></a>AfxUnadvise

Releases the events connection identified with the cookie returned by the **AfxAdvise** method.

```
FUNCTION AfxUnadvise (BYVAL pUnk AS ANY PTR, BYVAL riid AS IID PTR, BYVAL dwCookie AS DWORD) AS HRESULT
FUNCTION AfxUnadvise (BYVAL pUnk AS ANY PTR, BYREF riid AS CONST IID, BYVAL dwCookie AS DWORD) AS HRESULT
FUNCTION AfxUnadvise (BYVAL pUnk AS ANY PTR, BYREF riid AS IID, BYVAL dwCookie AS DWORD PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pUnk* | A pointer to the **IUnknown** of the object the client wants to disconnect with. |
| *riid* | The GUID of the connection point. Typically, this is the same as the outgoing interface managed by the connection point. |
| *dwCookie* | The cookie returned by **AfxAdvise** that uniquely identifies the connection. |

#### Return value

S_OK (0) on success, or an HRESULT error code on failure.

# <a name="AfxWstrAlloc"></a>AfxWstrAlloc

Takes a null terminated wide string as input, and returns a pointer to a new wide string allocated with **CoTaskMemAlloc**.

```
FUNCTION AfxWstrAlloc () AS WSTRING PTR
```

#### Remarks

The new string must be freed with **CoTaskMemFree**.

Useful when we need to pass a pointer to a null terminated wide string to a function or method that will release it. If we pass a WSTRING it will GPF. If the length of the input string is 0, **CoTaskMemAlloc** allocates a zero-length item and returns a valid pointer to that item. If there is insufficient memory available, **CoTaskMemAlloc** returns NULL.

# <a name="AfxGetOleErrorInfo"></a>AfxGetOleErrorInfo

Returns the description of the most recent OLE error in the current logical thread and clears the error state for the thread. It should be called as soon as possible after calling a method of an Automation interface and will only succeed if the object supports the IErrorInfo interface.

```
FUNCTION AfxGetOleErrorInfo () AS CWSTR
```

#### Return value

The description of the error on success or an empty string on failure.

# <a name="AfxOleCreateFont"></a>AfxOleCreateFont

Creates a standard **IFont** object.

```
FUNCTION AfxOleCreateFont (BYREF wszFaceName AS WSTRING, BYVAL cySize AS LONG, _
   BYVAL fWeight AS LONG = FW_NORMAL, BYVAL fItalic AS UBYTE = FALSE, _
   BYVAL fUnderline AS UBYTE = FALSE, BYVAL fStrikeOut AS UBYTE = FALSE, _
   BYVAL fCharSet AS UBYTE = DEFAULT_CHARSET) AS IFont PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFaceName* | A string that specifies the typeface name of the font. The length of this string must not exceed 31 characters. The **EnumFontFamilies** function can be used to enumerate the typeface names of all currently available fonts. If *wszFaceName* is an empty string, GDI uses the first font that matches the other specified attributes. |
| *cySize* | Specifies the height, in logical units, of the font's character cell or character. |
| *fWeight* | Initial weight of the font. If the weight is below 550 (the average of FW_NORMAL, 400, and FW_BOLD, 700), then the Bold property is also initialized to FALSE. If the weight is above 550, the Bold property is set to TRUE.<br>The following values are defined for convenience: FW_DONTCARE (0), FW_THIN (100), FW_EXTRALIGHT (200), FW_ULTRALIGHT (200), FW_LIGHT (300), FW_NORMAL (400), FW_REGULAR (400), FW_MEDIUM (500), FW_SEMIBOLD (600), FW_DEMIBOLD (600), FW_BOLD (700), FW_EXTRABOLD (800), FW_ULTRABOLD (800), FW_HEAVY (900), FW_BLACK (900) |
| *fItalic* | Specifies an italic font if set to TRUE. |
| *fUnderline* | Specifies an underlined font if set to TRUE. |
| *fStrikeOut* | Specifies a strikeout font if set to TRUE. |
| *fCharSet* | Specifies the character set. The following values are predefined:<br>ANSI_CHARSET, BALTIC_CHARSET, CHINESEBIG5_CHARSET, DEFAULT_CHARSET, EASTEUROPE_CHARSET, GB2312_CHARSET, GREEK_CHARSET, HANGUL_CHARSET, MAC_CHARSET, OEM_CHARSET, RUSSIAN_CHARSET, SHIFTJIS_CHARSET, SYMBOL_CHARSET, TURKISH_CHARSET.<br>Korean Windows: JOHAB_CHARSET<br>Middle-Eastern Windows: HEBREW_CHARSET, ARABIC_CHARSET.<br>Thai Windows: THAI_CHARSET.|

#### Remarks

The OEM_CHARSET value specifies a character set that is operating-system dependent.

DEFAULT_CHARSET is set to a value based on the current system locale. For example, when the system locale is English (United States), the value is ANSI_CHARSET.

Fonts with other character sets may exist in the operating system. If an application uses a font with an unknown character set, it should not attempt to translate or interpret strings that are rendered with that font.

This member is important in the font mapping process. To ensure consistent results, specify a specific character set.

#### Return value

A pointer to the object indicates success. NULL indicates failure.

#### Usage examples

```
DIM pFont AS IFont PTR = AfxOleCreateFont("MS Sans Serif", 8, FW_NORMAL, , , , DEFAULT_CHARSET)
DIM pFont AS IFont PTR = AfxOleCreateFont("Courier New", 10, FW_BOLD, , , , DEFAULT_CHARSET)
DIM pFont AS IFont PTR = AfxOleCreateFont("Marlett", 8, FW_NORMAL, , , , SYMBOL_CHARSET)
```

# <a name="AfxOleCreateFontDisp"></a>AfxOleCreateFontDisp

Creates a standard **IFontDisp** object.

```
FUNCTION AfxOleCreateFontDisp (BYREF wszFaceName AS WSTRING, BYVAL cySize AS LONG, _
   BYVAL fWeight AS LONG = FW_NORMAL, BYVAL fItalic AS UBYTE = FALSE, _
   BYVAL fUnderline AS UBYTE = FALSE, BYVAL fStrikeOut AS UBYTE = FALSE, _
   BYVAL fCharSet AS UBYTE = DEFAULT_CHARSET) AS IFontDisp PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFaceName* | A string that specifies the typeface name of the font. The length of this string must not exceed 31 characters. The **EnumFontFamilies** function can be used to enumerate the typeface names of all currently available fonts. If *wszFaceName* is an empty string, GDI uses the first font that matches the other specified attributes. |
| *cySize* | Specifies the height, in logical units, of the font's character cell or character. |
| *fWeight* | Initial weight of the font. If the weight is below 550 (the average of FW_NORMAL, 400, and FW_BOLD, 700), then the Bold property is also initialized to FALSE. If the weight is above 550, the Bold property is set to TRUE.<br>The following values are defined for convenience: FW_DONTCARE (0), FW_THIN (100), FW_EXTRALIGHT (200), FW_ULTRALIGHT (200), FW_LIGHT (300), FW_NORMAL (400), FW_REGULAR (400), FW_MEDIUM (500), FW_SEMIBOLD (600), FW_DEMIBOLD (600), FW_BOLD (700), FW_EXTRABOLD (800), FW_ULTRABOLD (800), FW_HEAVY (900), FW_BLACK (900) |
| *fItalic* | Specifies an italic font if set to TRUE. |
| *fUnderline* | Specifies an underlined font if set to TRUE. |
| *fStrikeOut* | Specifies a strikeout font if set to TRUE. |
| *fCharSet* | Specifies the character set. The following values are predefined:<br>ANSI_CHARSET, BALTIC_CHARSET, CHINESEBIG5_CHARSET, DEFAULT_CHARSET, EASTEUROPE_CHARSET, GB2312_CHARSET, GREEK_CHARSET, HANGUL_CHARSET, MAC_CHARSET, OEM_CHARSET, RUSSIAN_CHARSET, SHIFTJIS_CHARSET, SYMBOL_CHARSET, TURKISH_CHARSET.<br>Korean Windows: JOHAB_CHARSET<br>Middle-Eastern Windows: HEBREW_CHARSET, ARABIC_CHARSET.<br>Thai Windows: THAI_CHARSET.|

#### Remarks

The OEM_CHARSET value specifies a character set that is operating-system dependent.

DEFAULT_CHARSET is set to a value based on the current system locale. For example, when the system locale is English (United States), the value is ANSI_CHARSET.

Fonts with other character sets may exist in the operating system. If an application uses a font with an unknown character set, it should not attempt to translate or interpret strings that are rendered with that font.

This member is important in the font mapping process. To ensure consistent results, specify a specific character set.

#### Return value

A pointer to the object indicates success. NULL indicates failure.

#### Usage examples

```
DIM pFontDisp AS IFontDisp PTR = AfxOleCreateFontDisp("MS Sans Serif", 8, FW_NORMAL, , , , DEFAULT_CHARSET)
DIM pFontDisp AS IFontDisp PTR = AfxOleCreateFontDisp("Courier New", 10, FW_BOLD, , , , DEFAULT_CHARSET)
DIM pFontDisp AS IFontDisp PTR = AfxOleCreateFontDisp("Marlett", 8, FW_NORMAL, , , , SYMBOL_CHARSET)
```

# <a name="AfxGetVarType"></a>AfxGetVarType

Return the variant type.

```
FUNCTION AfxGetVarType (BYVAL pvar AS VARIANT PTR) AS VARTYPE
```

# <a name="AfxIsVarTypeFloat"></a>AfxIsVarTypeFloat

Returns true if the vartype is of type float (VT_R4 or VT_R8); false, otherwise.

```
FUNCTION AfxIsVarTypeFloat (BYVAL vt AS VARTYPE) AS BOOLEAN
```

# <a name="AfxIsVarTypeSignedInteger"></a>AfxIsVarTypeSignedInteger

Returns true if the vartype is of type signed integer (VT_I1, VT_I2, VT_I4, VT_I8); false, otherwise.

```
FUNCTION AfxIsVarTypeSignedInteger (BYVAL vt AS VARTYPE) AS BOOLEAN
```

# <a name="AfxIsVarTypeUnsignedInteger"></a>AfxIsVarTypeUnsignedInteger

Returns true if the vartype is of type unsigned integer (VT_UI1, VT_UI2, VT_UI4, VT_UI8); false, otherwise.

```
FUNCTION AfxIsVarTypeUnsignedInteger (BYVAL vt AS VARTYPE) AS BOOLEAN
```

# <a name="AfxIsVarTypeInteger"></a>AfxIsVarTypeInteger

Returns true if the vartype is of type integer (signed or unsigned); false, otherwise.

```
FUNCTION AfxIsVarTypeInteger (BYVAL vt AS VARTYPE) AS BOOLEAN
```

# <a name="AfxIsVarTypeNumber"></a>AfxIsVarTypeNumber

Returns true if the vartype is of type numeric (integer or float); false, otherwise.

```
FUNCTION AfxIsVarTypeNumber (BYVAL vt AS VARTYPE) AS BOOLEAN
```

# <a name="AfxIsVariantString"></a>AfxIsVariantString

Returns true if the variant contains an string.

```
PRIVATE FUNCTION AfxIsVariantString (BYVAL pvar AS VARIANT PTR) AS BOOLEAN
```

# <a name="AfxIsVariantArray"></a>AfxIsVariantArray

Returns true if the variant is of type VT_ARRAY; false, otherwise.

```
FUNCTION AfxIsVariantArray (BYVAL pvar AS VARIANT PTR) AS BOOLEAN
```

# <a name="AfxVarToStr"></a>AfxVarToStr

Extracts the contents of a VARIANT and returns them as a CWSTR.

```
FUNCTION AfxVarToStr (BYVAL pvarIn AS VARIANT PTR, BYVAL bClear AS BOOLEAN = FALSE) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *pvarIn* | Pointer to the variant. |
| *bClear* | Optional. Clear the contents of the variant (TRUE or FALSE). |

#### Return value

If the function succeeds, it returns the contents of the variant in string form; if it fails, it returns an empty string and the contents of the variant aren't cleared.

#### Remarks

When *pvarIn* contains an array, each element of the array is appended to the resulting string separated with a semicolon and a space. For variants that contains an array of bytes, use **AfxVariantToBuffer**.

# <a name="AfxVariantGetElementCount"></a>AfxVariantGetElementCount

Retrieves the element count of a variant structure.

```
FUNCTION AfxVariantGetElementCount (BYVAL pvarIn AS VARIANT PTR) AS ULONG
```

#### Return value

Returns the element count for values of type VT_ARRAY; otherwise, returns 1.

# <a name="AfxVariantToBuffer"></a>AfxVariantToBuffer

Extracts the contents of a variant that contains an array of bytes.

```
FUNCTION AfxVariantToBuffer (BYVAL pvarIn AS VARIANT PTR, BYVAL pv AS LPVOID, BYVAL cb AS ULONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pvarIn* | Pointer to the variant. |
| *pv* | Pointer to a buffer of length *cb* bytes. When this function returns, contains the first cb bytes of the extracted buffer value. |
| *cb* | The size of the *pv* buffer, in bytes. The buffer should be the same size as the data to be extracted. |

#### Return value

Returns one of the following values:

| HRESULT    | Description |
| ---------- | ----------- |
| S_OK | Data successfully extracted. |
| E_INVALIDARG | The VARIANT was not of type VT_ARRRAY OR VT_UI1. |
| E_FAIL | The VARIANT buffer value had fewer than *cb* bytes. |

# <a name="AfxVariantToString"></a>AfxVariantToString

Extracts the variant value of a variant structure to a string.

```
FUNCTION AfxVariantToString (BYVAL pvarIn AS VARIANT PTR, BYVAL pwszBuf AS WSTRING PTR, _
   BYVAL cchBuf AS UINT) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pvarIn* | Pointer to the variant. |
| *pwszBuf* | Pointer to a buffer of length *cchBuf* bytes. When this function returns, contains the first *cchBuf* bytes of the extracted buffer value. |
| *cchBuf* | The size of the *pwszBuf* buffer, in bytes. The buffer should be the same size as the data to be extracted. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

# <a name="AfxVariantToStringAlloc"></a>AfxVariantToStringAlloc

Extracts the variant value of a variant structure to a string.

```
FUNCTION AfxVariantToStringAlloc (BYVAL pvarIn AS VARIANT PTR, BYVAL ppwszBuf AS WSTRING PTR PTR) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pvarIn* | Pointer to the variant. |
| *ppwszBuf* | Pointer to a buffer that contains the extracted string exists. |

#### Return value

If this function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.
