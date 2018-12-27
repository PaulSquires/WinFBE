# CAxHost Class

**CAxHost** implements an OLE Container using a custom control to host ActiveX controls.

**Include file**: CaxHost.inc.

### Constructors

| Name       | Description |
| ---------- | ----------- |
| [CAXHOST_AMBIENTDISP structure](#CAXHOST_AMBIENTDISP) | Contains information the ambient properties of the CAxHost control. |
| [Constructor(ProgID)](#Constructor1) | Creates an instance of the OLE container using a ProgId. |
| [Constructor(ClsId)](#Constructor2) | Creates an instance of the OLE container using a Clsid. |
| [Constructor(LibName)](#Constructor3) | Creates an instance of the OLE container using an unregistered OCX. |

### Methods

| Name       | Description |
| ---------- | ----------- |
| [Advise](#Advise) | Establishes a connection between a connection point object and the client's sink. |
| [CtrlID](#CtrlID) | Returns the identifier of the container's window. |
| [GuidFromStr](#GuidFromStr) | Converts a string into a 16-byte (128-bit) Globally Unique Identifier (GUID). |
| [GuidText](#GuidText) | Returns a 38-byte human-readable guid string from a 16-byte GUID. |
| [hWindow](#hWindow) | Returns the handle of the hosting window. |
| [OcxDispObj](#OcxDispObj) | Returns a reference to the control's default interface given the handle of the window that hosts it. |
| [OcxDispPtr](#OcxDispPtr) | Returns a reference to the control's default interface given the handle of the window that hosts it. |
| [OleCreateFont](#OleCreateFont) | Creates a standard IFont object. |
| [OleCreateFontDisp](#OleCreateFontDisp) | Creates a standard IFont object. |
| [Unadvise](#Unadvise) | Terminates an advisory connection previously established by the **Advise** method between a connection point object and a client's sink. |

### Helper Procedures

| Name       | Description |
| ---------- | ----------- |
| [AfxCAxHostDispObj](#AfxCAxHostDispObj) | Returns a reference to the control's default interface given the handle of the window that hosts it. |
| [AfxCAxHostDispPtr](#AfxCAxHostDispPtr) | Returns a reference to the control's default interface given the handle of the window that hosts it. |
| [AfxCAxHostPtr](#AfxCAxHostPtr) | Returns a reference to the OLE container class given the handle of its associated window. |
| [AfxCAxHostWindow](#AfxCAxHostWindow) | Returns the OLE container window handle given the handle of the form, or any control in the form, and its control identifier. |
| [AfxCAxHostForwardMessage](#AfxCAxHostForwardMessage) | Forwards the Windows messages to the OLE Container window for processing. |

# <a name="CAXHOST_AMBIENTDISP"></a>CAXHOST_AMBIENTDISP Structure

Contains information the ambient properties of the **CAxHost** control.

```
TYPE CAXHOST_AMBIENTDISP
   Font AS IFontDisp PTR
   BackColor AS LONG = &hFFFFFF
   ForeColor AS LONG = 0
   LocaleID AS LONG = LOCALE_USER_DEFAULT
   Silent AS VARIANT_BOOL = -1
   UIDead AS VARIANT_BOOL = 0
   DisplayAsDefault AS VARIANT_BOOL = 0
   SupportMnemonics AS VARIANT_BOOL = -1
   OffLineIfNoConnected AS VARIANT_BOOL = -1
   DlCtFlags AS LONG = 0
END TYPE
```

| Member     | Description |
| ---------- | ----------- |
| **Font** | Pointer to a **IFontDisp** interface. To create the font you can use the **OleCreateFontDisp** method of the CAxHost class or the **AfxOleCreateFontDisp** function. The default font is SegoeUI, 9 points. |
| **BackColor** | Specifies the back color of the container. The default color is white. You can use the FreeBasic BGR macro to set the color. |
| **ForeColor** | Specifies the fore color of the container. The default color is black. You can use the FreeBasic BGR macro to set the color. |
| **LocaleID** | Specifies the ambient locale ID of the container. |
| **Silent** | Indicates the bind operation will be completed silently. No user interface or user notification will occur. |
| **UIDead** | Indicates whether the container wants the control to respond to user-interface actions. If TRUE, the control should not respond. Default value: FALSE. |
| **DisplayAsDefault** | A flag that is TRUE if the container has marked the control in this site to be a default button, and therefore a button control should draw itself with a thicker frame. Default value: FALSE. |
| **SupportMnemonics** | Indicates whether the container supports keyboard mnemonics. Default value: TRUE. |
| **OffLineIfNotConnected** | WebBroser control. The control support offline browsing. |
| **DlCtFlags** | A combination of following flags, using the bitwise OR operator, to indicate your preferences. Most of the flag values have negative effects, that is, they prevent behavior that normally happens. For instance, scripts are normally executed by the WebBrowser Control if you don't customize its behavior. But if you set the DLCTL_NOSCRIPTS flag, no scripts will execute in that instance of the control. However, three flags—DLCTL_DLIMAGES, DLCTL_VIDEOS, and DLCTL_BGSOUNDS—work exactly opposite. If you set flags at all, you must set these three for the WebBrowser Control to behave in its default manner vis-a-vis images, videos and sounds. |

**Flags**

* DLCTL_DLIMAGES, DLCTL_VIDEOS, and DLCTL_BGSOUNDS: Images, videos, and background sounds will be downloaded from the server and displayed or played if these flags are set. They will not be downloaded and displayed if the flags are not set.
* DLCTL_NO_SCRIPTS and DLCTL_NO_JAVA: Scripts and Java applets will not be executed.
* DLCTL_NO_DLACTIVEXCTLS and DLCTL_NO_RUNACTIVEXCTLS : ActiveX controls will not be downloaded or will not be executed.
* DLCTL_DOWNLOADONLY: The page will only be downloaded, not displayed.
* DLCTL_NO_FRAMEDOWNLOAD: The WebBrowser Control will download and parse a frameset, but not the individual frame objects within the frameset.
* DLCTL_RESYNCHRONIZE and DLCTL_PRAGMA_NO_CACHE: These flags cause cache refreshes. With DLCTL_RESYNCHRONIZE, the server will be asked for update status. Cached files will be used if the server indicates that the cached information is up-to-date. With DLCTL_PRAGMA_NO_CACHE, files will be re-downloaded from the server regardless of the update status of the files.
* DLCTL_NO_BEHAVIORS: Behaviors are not downloaded and are disabled in the document.
* DLCTL_NO_METACHARSET_HTML: Character sets specified in meta elements are suppressed.
* DLCTL_URL_ENCODING_DISABLE_UTF8 and DLCTL_URL_ENCODING_ENABLE_UTF8: These flags function similarly to the DOCHOSTUIFLAG_URL_ENCODING_DISABLE_UTF8 and DOCHOSTUIFLAG_URL_ENCODING_ENABLE_UTF8 flags used with IDocHostUIHandler.GetHostInfo. The difference is that the DOCHOSTUIFLAG flags are checked only when the WebBrowser Control is first instantiated. The download flags here for the ambient property change are checked whenever the WebBrowser Control needs to perform a download.
* DLCTL_NO_CLIENTPULL: No client pull operations will be performed.
* DLCTL_SILENT: No user interface will be displayed during downloads.
* DLCTL_FORCEOFFLINE: The WebBrowser Control always operates in offline mode.
* DLCTL_OFFLINEIFNOTCONNECTED and DLCTL_OFFLINE: These flags are the same. The WebBrowser Control will operate in offline mode if not connected to the Internet.

# <a name="Constructor1"></a>Constructor(ProgId)

Creates an instance of the OLE container using a ProgId.

```
CONSTRUCTOR CAxHost (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, _
   BYREF wszProgID AS WSTRING, BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, _
   BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
   BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, _
   BYVAL pAmbientDisp AS CAXHOST_AMBIENTDISP PTR = NULL)
```
```
CONSTRUCTOR CAxHost (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, _
   BYREF wszProgID AS WSTRING, BYREF wszLicKey AS WSTRING, BYVAL x AS LONG = 0, _
   BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
   BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, _
   BYVAL pAmbientDisp AS CAXHOST_AMBIENTDISP PTR = NULL)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWindow* | Pointer to the instance of the **CWindow** class used to create the form. |
| *cID* | The identifier of the control. It must be unique. |
| *wszProgID* | The **ProgID** of the object to create, such "MSCAL.Calendar". |
| *wszLicKey* | The license key of the control, if the control is licensed. |
| *x, y* | The coordinates of the upper-left corner of the window relative to the upper-left corner of the parent window's client area. |
| *nWidth* | The width of the window. |
| *nHeight* | The height of the window. |
| *dwStyle* | The style of the window being created. |
| *dwExStyle* | The extended style of the window being created. |
| *pAmbientDisp* | Pointer to a **CAXHOST_AMBIENTDISP** structure. |

The following example embeds an instance of the WebBrowser control and navigates to a site:

| Name       | Description |
| ---------- | ----------- |
| [Example1](#Example1) | Embedding the WebBroser control. |
| [Example2](#Example2) | Google Map. |


# <a name="Constructor2"></a>Constructor(ClsId)

Creates an instance of the OLE container using a ClsId.

```
CONSTRUCTOR CAxHost (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, _
   BYREF classID AS CONST CLSID, BYREF riid AS CONST IID, BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, _
   BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, BYVAL dwStyle AS DWORD = 0, _
   BYVAL dwExStyle AS DWORD = 0, BYVAL pAmbientDisp AS CAXHOST_AMBIENTDISP PTR = NULL)
```
```
CONSTRUCTOR CAxHost (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, _
   BYREF classID AS CONST CLSID, BYREF riid AS CONST IID, BYREF wszLicKey AS WSTRING, _
   BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
   BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, _
   BYVAL pAmbientDisp AS CAXHOST_AMBIENTDISP PTR = NULL)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWindow* | Pointer to the instance of the **CWindow** class used to create the form. |
| *cID* | The identifier of the control. It must be unique. |
| *classID* | The **CLSID** of the object to create, such "{8E27C92B-1264-101C-8A2F-040224009C02}" |
| *riid* | The **IID** of the object to create, such "{8E27C92C-1264-101C-8A2F-040224009C02}". |
| *wszLicKey* | The license key of the control, if the control is licensed. |
| *x, y* | The coordinates of the upper-left corner of the window relative to the upper-left corner of the parent window's client area. |
| *nWidth* | The width of the window. |
| *nHeight* | The height of the window. |
| *dwStyle* | The style of the window being created. |
| *dwExStyle* | The extended style of the window being created. |
| *pAmbientDisp* | Pointer to a **CAXHOST_AMBIENTDISP** structure. |

# <a name="Constructor3"></a>Constructor(LibName)

Creates an instance of the OLE container using an unregistered OCX.

```
CONSTRUCTOR CAxHost (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, _
   BYREF wszLibName AS WSTRING, BYREF classID AS CONST CLSID, BYREF riid AS CONST IID, _
   BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, _
   BYVAL nHeight AS LONG = 0, BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, _
   BYVAL pAmbientDisp AS CAXHOST_AMBIENTDISP = NULL)
```
```
CONSTRUCTOR CAxHost (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, _
   BYREF wszLibName AS WSTRING, BYREF classID AS CONST CLSID, BYREF riid AS CONST IID, _
   BYREF wszLicKey AS WSTRING, BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, _
   BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, BYVAL dwStyle AS DWORD = 0, _
   BYVAL dwExStyle AS DWORD = 0, BYVAL pAmbientDisp AS CAXHOST_AMBIENTDISP PTR = NULL)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWindow* | Pointer to the instance of the **CWindow** class used to create the form. |
| *cID* | The identifier of the control. It must be unique. |
| *wszLibName* | The full qualified path to the OCX file. |
| *classID* | The **CLSID** of the object to create, such "{8E27C92B-1264-101C-8A2F-040224009C02}" |
| *riid* | The **IID** of the object to create, such "{8E27C92C-1264-101C-8A2F-040224009C02}". |
| *wszLicKey* | The license key of the control, if the control is licensed. |
| *x, y* | The coordinates of the upper-left corner of the window relative to the upper-left corner of the parent window's client area. |
| *nWidth* | The width of the window. |
| *nHeight* | The height of the window. |
| *dwStyle* | The style of the window being created. |
| *dwExStyle* | The extended style of the window being created. |
| *pAmbientDisp* | Pointer to a **CAXHOST_AMBIENTDISP** structure. |

```
DIM wszLibName AS WSTRING * 260 = ExePath & "\MSCOMCT2.OCX"
DIM CLSID_MSComCtl2_MonthView AS CLSID = (&h232E456A, &h87C3, &h11D1, {&h8B, &hE3,&h00, &h00, &hF8, &h75, &h4D, &hA1})
DIM IID_MSComCtl2_MonthView AS CLSID = (&h232E4565, &h87C3, &h11D1, {&h8B, &hE3,&h00, &h00, &hF8, &h75, &h4D, &hA1})
DIM RTLKEY_MSCOMCT2 AS WSTRING * 260 = "651A8940-87C5-11d1-8BE3-0000F8754DA1"
DIM pAxHost AS CAxHost PTR = CAxHost(@pWindow, IDC_WEBBROWSER, wszLibName, CLSID_MSComCtl2_MonthView, _
    IID_MSComCtl2_MonthView, RTLKEY_MSCOMCT2, 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
```

# <a name="Advise"></a>Advise

Establishes a connection between a connection point object and the client's sink.

```
FUNCTION Advise (BYVAL pEvtObj AS IDispatch PTR, BYVAL riid AS IID PTR) AS HRESULT
FUNCTION Advise (BYVAL pEvtObj AS IDispatch PTR, BYVAL riid AS CONST IID) AS HRESULT
FUNCTION Advise (BYVAL pEvtObj AS IDispatch PTR, BYREF riid AS IID) AS HRESULT
FUNCTION Advise (BYVAL pEvtObj AS IDispatch PTR, BYVAL riid AS IID PTR) AS HRESULT
FUNCTION Advise (BYVAL pEvtObj AS IDispatch PTR, BYREF riid AS CONST IID) AS HRESULT
FUNCTION Advise (BYVAL pEvtObj AS IDispatch PTR, BYREF riid AS IID) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pEvtObj* | Pointer to the sink event's class. The client's sink receives outgoing calls from the connection point. |
| *riid* | IID of the events interface of the hosted OCX. |

Return value

S_OK (0) or an error code.

# <a name="CtrlID"></a>CtrlID

Returns the identifier of the container's window.

```
FUNCTION CtrlID () AS LONG
```

# <a name="GuidFromStr"></a>GuidFromStr

Converts a string into a 16-byte (128-bit) Globally Unique Identifier (GUID)

```
FUNCTION GuidFromStr (BYVAL pwszGuidText AS WSTRING PTR = NULL) AS GUID
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszGuidText* | To be valid, the string must contain exactly 32 hexadecimal digits, delimited by hyphens and enclosed by curly braces. For example: {B09DE715-87C1-11D1-8BE3-0000F8754DA1}. If pwszGuidText is omited, **GuidFromStr** generates a new unique guid. |

# <a name="GuidText"></a>GuidText

Returns a 38-byte human-readable guid string from a 16-byte GUID.

```
FUNCTION GuidText (BYVAL classID AS CLSID PTR) AS STRING
FUNCTION GuidText (BYVAL riid AS REFIID) AS STRING
FUNCTION GuidText (BYVAL classID AS CLSID) AS STRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *classID, riid* | The 16-byte GUID. |

#### Return value

The GUID as a human readable string.

# <a name="hWindow"></a>hWindow

Returns the handle of the hosting window.

```
FUNCTION hWindow () AS HWND
```

# <a name="OcxDispObj"></a>OcxDispObj

Returns a reference to the control's default interface given the handle of the window that hosts it. This is a reference counted pointer, so you must release it calling the **Release** method of the interface, the **IUnknown_Release** macro or the **AfxSaRelease** function.

```
FUNCTION OcxDispObj () AS IDispatch PTR
```

#### Return value

IDispatch. A reference to the ActiveX control's default interface.

# <a name="OcxDispPtr"></a>OcxDispPtr

Returns a reference to the control's default interface given the handle of the window that hosts it. This is not a reference counted pointer, so you must not release it.

```
FUNCTION OcxDispPtr () AS IDispatch PTR
```

#### Return value

IDispatch. A reference to the ActiveX control's default interface.

# <a name="OleCreateFont"></a>OleCreateFont

Creates a standard **IFont** object.

```
FUNCTION OleCreateFont (BYREF wszFaceName AS WSTRING, BYVAL cySize AS LONG, _
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

#### Return value

A pointer to the object indicates success. NULL indicates failure.

#### Usage examples

```
DIM pFont AS IFont = CAxHost.CreateFont("MS Sans Serif", 8, FW_NORMAL, , , , DEFAULT_CHARSET)
DIM pFont AS IFont = CAxHost.CreateFont("Courier New", 10, FW_BOLD, , , , DEFAULT_CHARSET)
DIM pFont AS IFont = CAxHost.CreateFont("Marlett", 8, FW_NORMAL, , , , SYMBOL_CHARSET)
```

# <a name="OleCreateFontDisp"></a>OleCreateFontDisp

Creates a standard **IFontDisp** object.

```
FUNCTION OleCreateFontDisp (BYREF wszFaceName AS WSTRING, BYVAL cySize AS LONG, _
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

#### Return value

A pointer to the object indicates success. NULL indicates failure.

#### Usage examples

```
DIM pFontDisp AS IFontDisp = CAxHost.CreateFontDisp("MS Sans Serif", 8, FW_NORMAL, , , , DEFAULT_CHARSET)
DIM pFontDisp AS IFontDisp = CAxHost.CreateFontDisp("Courier New", 10, FW_BOLD, , , , DEFAULT_CHARSET)
DIM pFontDisp AS IFontDisp = CAxHost.CreateFontDisp("Marlett", 8, FW_NORMAL, , , , SYMBOL_CHARSET)
```

# <a name="Unadvise"></a>Unadvise

Terminates an advisory connection previously established by the **Advise** method between a connection point object and a client's sink.

```
FUNCTION Unadvise () AS HRESULT
```

#### Retgurn value

S_OK (0) or an error code.

# <a name="AfxCAxHostDispObj"></a>AfxCAxHostDispObj

Returns a reference to the control's default interface given the handle of the window that hosts it. This is a reference counted pointer, so you must release it calling the **Release** method of the interface, the **IUnknown_Release** macro or the **AfxSafeRelease** function.

```
FUNCTION AfxCAxHostDispObj (BYVAL hwnd AS HWND) AS IDispatch PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle of the window that hosts he ActiveX control. If you don't know the value of this handle, you can get it calling the **AfxCAxHostWindow** function. |

#### Return value

IDispatch. A reference to the ActiveX control's default interface.

# <a name="AfxCAxHostDispPtr"></a>AfxCAxHostDispPtr

Returns a reference to the OLE container class given the handle of its associated window. This is not and AddRefed pointer, so don't release it.

```
FUNCTION AfxCAxHostDispPtr (BYVAL hwnd AS HWND) AS IDispatch PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle of the window that hosts he ActiveX control. If you don't know the value of this handle, you can get it calling the **AfxCAxHostWindow** function. |

#### Return value

IDispatch. A reference to the ActiveX control's default interface.

# <a name="AfxCAxHostPtr"></a>AfxCAxHostPtr

Returns a reference to the OLE container class given the handle of its associated window.

```
FUNCTION AfxCAxHostPtr (BYVAL hwnd AS HWND) AS CAxHost PTR
FUNCTION AfxCAxHostPtr (BYVAL hwnd AS HWND, BYVAL cID AS WORD) AS CAxHost PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle of the window that hosts he ActiveX control. If you don't know the value of this handle, you can get it calling the **AfxCAxHostWindow** function. |
| *cID* | Identifier of the control, e.g. IDC_WEBBROWSER. |

#### Return value

A pointer to the **CAxHost** class or NULL.

# <a name="AfxCAxHostWindow"></a>AfxCAxHostWindow

Returns the OLE container window handle given the handle of the form, or any control in the form, and its control identifier.

```
FUNCTION AfxCAxHostWindow (BYVAL hwnd AS HWND, BYVAL cID AS WORD) AS HWND
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle of the window that hosts he ActiveX control. If you don't know the value of this handle, you can get it calling the **AfxCAxHostWindow** function. |
| *cID* | Identifier of the control, e.g. IDC_WEBBROWSER. |

#### Return value

The handle of the window or NULL.

# <a name="AfxCAxHostForwardMessage"></a>AfxCAxHostForwardMessage

Forwards the Windows messages to the OLE Container window for processing.

```
FUNCTION AfxCAxHostForwardMessage (BYVAL hctl AS HWND, BYVAL pMsg AS tagMSG PTR) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hctl* | Handle of the control. |
| *pMsg* | A pointer to a **MSG** structure that contains information about a message |

#### Return value

TRUE if the message was processed, FALSE if not.

#### Remarks

Active in-place objects must always be given the first chance at translating accelerator keystrokes. You can provide this opportunity by calling **IOleInPlaceActiveObject.TranslateAccelerator** from your container's message loop before doing any other translation. You should apply your own translation only when this method returns FALSE.

#### Usage example

```
' // Dispatch Windows messages
DIM uMsg AS MSG
WHILE (GetMessageW(@uMsg, NULL, 0, 0) <> FALSE)
   IF AfxForwardMessage(GetFocus, @uMsg) = FALSE THEN
      IF IsDialogMessageW(hWndMain, @uMsg) = 0 THEN
         TranslateMessage(@uMsg)
         DispatchMessageW(@uMsg)
      END IF
   END IF
WEND
```


# <a name="Example1"></a>Embedding the WebBrowser control

The following example embeds an instance of the WebBrowser control and navigates to a site:

```
' ########################################################################################
' Microsoft Windows
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#include once "Afx/CAxHost/CAxHost.inc"
USING Afx

CONST IDC_WEBBROWSER = 1000

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG

   END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

' ========================================================================================
' Main
' ========================================================================================
FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                  BYVAL hPrevInstance AS HINSTANCE, _
                  BYVAL szCmdLine AS ZSTRING PTR, _
                  BYVAL nCmdShow AS LONG) AS LONG

   ' // Set process DPI aware
   ' // The recommended way is to use a manifest file
   AfxSetProcessDPIAware

   ' // Creates the main window
   DIM pWindow AS CWindow
   ' -or- DIM pWindow AS CWindow = "MyClassName" (use the name that you wish)
   DIM hwndMain AS HWND = pWindow.Create(NULL, "CAxHost test", @WndProc)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(900, 450)
   ' // Centers the window
   pWindow.Center

   ' // Create an instance of the WebBrowser control
   DIM pHost AS CAxHost = CAxHost(@pWindow, IDC_WEBBROWSER, "Shell.Explorer", 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   ' // Get a direct pointer to the Afx_IWebBrowser2 interface
   DIM hWb AS HWND = GetDlgItem(pWindow.hWindow, IDC_WEBBROWSER)
   DIM pWb AS Afx_IWebBrowser2 PTR = CAST(ANY PTR, AfxCAxHostDispPtr(hWb))
   IF pWb THEN
      ' // Navigate to a web page
      DIM wszUrl AS WSTRING * 260 = "http://www.planetsquires.com/protect/forum/index.php"
      DIM vUrl AS VARIANT : vUrl.vt = VT_BSTR : vUrl.bstrVal = SysAllocString(wszUrl)
      DIM hr AS HRESULT = pWb->Navigate2(@vUrl, NULL, NULL, NULL, NULL)
      VariantClear @vurl
   END IF

   ' // Display the window
   ShowWindow(hWndMain, nCmdShow)
   UpdateWindow(hWndMain)

   ' // Dispatch Windows messages
   DIM uMsg AS MSG
   WHILE (GetMessageW(@uMsg, NULL, 0, 0) <> FALSE)
      IF AfxCAxHostForwardMessage(GetFocus, @uMsg) = FALSE THEN
         IF IsDialogMessageW(hWndMain, @uMsg) = 0 THEN
            TranslateMessage(@uMsg)
            DispatchMessageW(@uMsg)
         END IF
      END IF
   WEND
   FUNCTION = uMsg.wParam

END FUNCTION
' ========================================================================================

' ========================================================================================
' Main window procedure
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
         ' // Optional resizing code
         IF wParam <> SIZE_MINIMIZED THEN
            ' // Retrieve a pointer to the CWindow class
            DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
            ' // Move the position of the control
            IF pWindow THEN pWindow->MoveWindow GetDlgItem(hwnd, IDC_WEBBROWSER), 0, 0, pWindow->ClientWidth, pWindow->ClientHeight, CTRUE
         END IF

    	CASE WM_DESTROY
         ' // Ends the application by sending a WM_QUIT message
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
```

# <a name="Example2"></a>Google Map

The following example embeds an instance of the WebBrowser control and uses an script to gost Google Map:

```
' ########################################################################################
' Microsoft Windows
' Contents: WebBrowser - Google Map
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#define _CAXH_DEBUG_ 1
#include once "Afx/CAxHost/CAxHost.inc"
USING Afx

CONST IDC_WEBBROWSER = 1001

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG

   END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

' ========================================================================================
' Build the script
' Added ".infomsg {display:none;}" to the style to avoid this message:
' "You are using a browser that is not supported by the Google Maps JavaScript API.
' Consider changing your browser".
' ========================================================================================
FUNCTION BuildScript () AS STRING
   DIM s AS STRING
   s  = "<!DOCTYPE html>"
   s += "<html>"
   s += "   <head>"
   s += "      <meta http-equiv='X-UA-Compatible' content='IE=edge'>"
   s += "      <meta http-equiv='MSThemeCompatible' content='yes'>"
   s += "      <style>"
   s += "         .infomsg {display:none;}"
   s += "         html, body, #map-canvas"
   s += "         {"
   s += "            height: 100%;"
   s += "            margin: 0px;"
   s += "            padding: 0px"
   s += "         }"
   s += "      </style>"
   s += "      <script src='https://maps.googleapis.com/maps/api/js?v=3'></script>"
   s += "      <script>"
   s += "         var globalLatLng = '';"
   s += "         function initialize()"
   s += "         {"
   s += "            var mapOptions ="
   s += "            {"
   s += "               zoom: 8,"
   s += "               center: new google.maps.LatLng(51.5, -0.2),"
   s += "               mapTypeId: google.maps.MapTypeId.ROADMAP"
   s += "            };"
   s += "            var map = new google.maps.Map(document.getElementById('map-canvas'), mapOptions);"
   s += "         }"
   s += "         google.maps.event.addDomListener(window, 'load', initialize);"
   s += "      </script>"
   s += "   </head>"
   s += "   <body>"
   s += "      <div id='map-canvas'></div>"
   s += "   </body>"
   s += "</html>"
   FUNCTION = s
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
   ' // The recommended way is to use a manifest file
   AfxSetProcessDPIAware

   ' // Creates the main window
   DIM pWindow AS CWindow
   ' -or- DIM pWindow AS CWindow = "MyClassName" (use the name that you wish)
   DIM hwndMain AS HWND = pWindow.Create(NULL, "WebBrowser - Google Map", @WndProc)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(750, 450)
   ' // Centers the window
   pWindow.Center

   ' // Create an instance of the WebBrowser control
   DIM pHost AS CAxHost = CAxHost(@pWindow, IDC_WEBBROWSER, "Shell.Explorer", 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   ' // Get a direct pointer to the Afx_IWebBrowser2 interface
   DIM hWb AS HWND = GetDlgItem(pWindow.hWindow, IDC_WEBBROWSER)
   DIM pWb AS Afx_IWebBrowser2 PTR = CAST(ANY PTR, AfxCAxHostDispPtr(hWb))
   IF pWb THEN
      ' // Build the script
      DIM s AS STRING = BuildScript
      ' // Save the script as a temporary file
      DIM wszPath AS WSTRING * MAX_PATH = AfxSaveTempFile(s, "html")
      ' // Navigate to a web page
      DIM vUrl AS VARIANT : vUrl.vt = VT_BSTR : vUrl.bstrVal = SysAllocString(wszPath)
      DIM hr AS HRESULT = pWb->Navigate2(@vUrl, NULL, NULL, NULL, NULL)
      VariantClear @vurl
      ' // Processes pending Windows messages to allow the page to load
      ' // Needed if the message pump isn't running
      AfxPumpMessages
      ' // Kill the temporary file
      KILL wszPath
   END IF

   ' // Displays the window and dispatches the Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Main window procedure
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
         ' // Optional resizing code
         IF wParam <> SIZE_MINIMIZED THEN
            ' // Retrieve a pointer to the CWindow class
            DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
            ' // Move the position of the control
            IF pWindow THEN pWindow->MoveWindow GetDlgItem(hwnd, IDC_WEBBROWSER), _
               0, 0, pWindow->ClientWidth, pWindow->ClientHeight, CTRUE
         END IF

    	CASE WM_DESTROY
         ' // Ends the application by sending a WM_QUIT message
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================

```
