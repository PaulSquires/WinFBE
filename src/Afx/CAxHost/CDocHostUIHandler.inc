' ########################################################################################
' CDocHostUIHandler2 class
' IID = {3050F6D0-98B5-11CF-BB82-00AA00BDCE0B}
' ########################################################################################
TYPE CDocHostUIHandler2 EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION ShowContextMenu (BYVAL dwID AS DWORD, BYVAL ppt AS POINT PTR, BYVAL pcmdtReserved AS IUnknown PTR, BYVAL pdispReserved AS IDispatch PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetHostInfo (BYVAL pInfo AS DOCHOSTUIINFO PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION ShowUI (BYVAL dwID AS DWORD, BYVAL pActiveObject AS IOleInPlaceActiveObject PTR, BYVAL pCommandTarget AS IOleCommandTarget PTR, BYVAL pFrame AS IOleInPlaceFrame PTR, BYVAL pDoc AS IOleInPlaceUIWindow PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION HideUI () AS HRESULT
   DECLARE VIRTUAL FUNCTION UpdateUI () AS HRESULT
   DECLARE VIRTUAL FUNCTION EnableModeless (BYVAL fEnable AS WINBOOL) AS HRESULT
   DECLARE VIRTUAL FUNCTION OnDocWindowActivate (BYVAL fActivate AS WINBOOL) AS HRESULT
   DECLARE VIRTUAL FUNCTION OnFrameWindowActivate (BYVAL fActivate AS WINBOOL) AS HRESULT
   DECLARE VIRTUAL FUNCTION ResizeBorder (BYVAL prcBorder AS LPCRECT, BYVAL pUIWindow AS IOleInPlaceUIWindow PTR, BYVAL fRameWindow AS WINBOOL) AS HRESULT
   DECLARE VIRTUAL FUNCTION TranslateAccelerator (BYVAL lpMsg AS LPMSG, BYVAL pguidCmdGroup AS const GUID PTR, BYVAL nCmdID AS DWORD) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetOptionKeyPath (BYVAL pchKey AS LPOLESTR PTR, BYVAL dw AS DWORD) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetDropTarget (BYVAL pDropTarget AS IDropTarget PTR, BYVAL ppDropTarget AS IDropTarget PTR PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetExternal (BYVAL ppDispatch AS IDispatch PTR PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION TranslateUrl (BYVAL dwTranslate AS DWORD, BYVAL pchURLIn AS LPWSTR, BYVAL ppchURLOut AS LPWSTR PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION FilterDataObject (BYVAL pDO AS IDataObject PTR, BYVAL ppDORet AS IDataObject PTR PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetOverrideKeyPath (BYVAL pchKey AS LPOLESTR PTR, BYVAL dw AS DWORD) AS HRESULT

   DECLARE CONSTRUCTOR (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR)
   DECLARE DESTRUCTOR

   m_pWebCtx AS CWebCtx PTR                ' // Pointer to the CWbCtx class
   m_pWebBrowser AS Afx_IWebBrowser2 PTR   ' // Pointer to the IWebBrowser interface
   cRef AS ULONG                           ' // Reference count

   ' // Pointers to user implemented event procedures
   m_ShowContextMenuProc AS FUNCTION (BYVAL pWebCtx AS CWebCtx PTR, BYVAL dwID AS DWORD, BYVAL ppt AS POINT PTR, BYVAL pcmdtReserved AS IUnknown PTR, BYVAL pdispReserved AS IDispatch PTR) AS HRESULT
   m_GetHostInfoProc AS FUNCTION (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pInfo AS DOCHOSTUIINFO PTR) AS HRESULT
   m_ShowUIProc AS FUNCTION (BYVAL pWebCtx AS CWebCtx PTR, BYVAL dwID AS DWORD, BYVAL pActiveObject AS IOleInPlaceActiveObject PTR, BYVAL pCommandTarget AS IOleCommandTarget PTR, BYVAL pFrame AS IOleInPlaceFrame PTR, BYVAL pDoc AS IOleInPlaceUIWindow PTR) AS HRESULT
   m_HideUIProc AS FUNCTION (BYVAL pWebCtx AS CWebCtx PTR) AS HRESULT
   m_UpdateUIProc AS FUNCTION (BYVAL pWebCtx AS CWebCtx PTR) AS HRESULT
   m_EnableModelessProc AS FUNCTION (BYVAL pWebCtx AS CWebCtx PTR, BYVAL fEnable AS WINBOOL) AS HRESULT
   m_OnDocWindowActivateProc AS FUNCTION (BYVAL pWebCtx AS CWebCtx PTR, BYVAL fActivate AS WINBOOL) AS HRESULT
   m_OnFrameWindowActivateProc AS FUNCTION (BYVAL pWebCtx AS CWebCtx PTR, BYVAL fActivate AS WINBOOL) AS HRESULT
   m_ResizeBorderProc AS FUNCTION (BYVAL pWebCtx AS CWebCtx PTR, BYVAL prcBorder AS LPCRECT, BYVAL pUIWindow AS IOleInPlaceUIWindow PTR, BYVAL frameWindow AS WINBOOL) AS HRESULT
   m_TranslateAcceleratorProc AS FUNCTION (BYVAL pWebCtx AS CWebCtx PTR, BYVAL lpMsg AS LPMSG, BYVAL pguidCmdGroup AS const GUID PTR, BYVAL nCmdID AS DWORD) AS HRESULT
   m_GetOptionKeyPathProc AS FUNCTION (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pchKey AS LPOLESTR PTR, BYVAL dw AS DWORD) AS HRESULT
   m_GetDropTargetProc AS FUNCTION (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pDropTarget AS IDropTarget PTR, BYVAL ppDropTarget AS IDropTarget PTR PTR) AS HRESULT
   m_GetExternalProc AS FUNCTION (BYVAL pWebCtx AS CWebCtx PTR, BYVAL ppDispatch AS IDispatch PTR PTR) AS HRESULT
   m_TranslateUrlProc AS FUNCTION (BYVAL pWebCtx AS CWebCtx PTR, BYVAL dwTranslate AS DWORD, BYVAL pchURLIn AS LPWSTR, BYVAL ppchURLOut AS LPWSTR PTR) AS HRESULT
   m_FilterDataObjectProc AS FUNCTION (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pDO AS IDataObject PTR, BYVAL ppDORet AS IDataObject PTR PTR) AS HRESULT
   m_GetOverrideKeyPathProc AS FUNCTION (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pchKey AS LPOLESTR PTR, BYVAL dw AS DWORD) AS HRESULT

END TYPE
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
CONSTRUCTOR CDocHostUIHandler2 (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR)
   CWBX_DP("CDocHostUIHandler2 CONSTRUCTOR")
   m_pWebCtx = pWebCtx
   m_pWebBrowser = pWebBrowser
END CONSTRUCTOR
' =====================================================================================

' =====================================================================================
' Destructor
' =====================================================================================
DESTRUCTOR CDocHostUIHandler2
   CWBX_DP("CDocHostUIHandler2 DESTRUCTOR")
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
FUNCTION CDocHostUIHandler2.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWBX_DP("CDocHostUIHandler2 QueryInterface")
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @IID_IDocHostUIHandler) OR _
      IsEqualIID(riid, @IID_IUnknown) OR _
      IsEqualIID(riid, @IID_IDispatch) THEN
      *ppvObj = @this
      ' // Not really needed, since this is not a COM object
      CAST(Afx_IUnknown PTR, *ppvObj)->AddRef
      RETURN S_OK
   END IF
   RETURN E_NOINTERFACE
END FUNCTION
' ========================================================================================

' ========================================================================================
FUNCTION CDocHostUIHandler2.AddRef () AS ULONG
   this.cRef += 1
   CWBX_DP("CDocHostUIHandler2 AddRef = " & WSTR(this.cRef))
   RETURN this.cRef
END FUNCTION
' ========================================================================================

' ========================================================================================
FUNCTION CDocHostUIHandler2.Release () AS ULONG
   this.cRef -= 1
   CWBX_DP("CDocHostUIHandler2 Release = " & WSTR(this.cRef))
   FUNCTION = this.cRef
   ' // Don't delete it because this is not a COM class
   'IF this.cRef = 0 THEN DELETE @this
END FUNCTION
' ========================================================================================

' =====================================================================================
' Enables MSHTML to display a shortcut menu.
' =====================================================================================
FUNCTION CDocHostUIHandler2.ShowContextMenu (BYVAL dwID AS DWORD, BYVAL ppt AS POINT PTR, BYVAL pcmdtReserved AS IUnknown PTR, BYVAL pdispReserved AS IDispatch PTR) AS HRESULT
   CWBX_DP("CDocHostUIHandler2.ShowContextMenu")
   IF m_ShowContextMenuProc THEN RETURN m_ShowContextMenuProc(m_pWebCtx, dwID, ppt, pcmdtReserved, pdispReserved)
   RETURN S_FALSE
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the UI capabilities of the application that is hosting MSHTML.
' DOCHOSTUIFLAG_DIALOG
'     MSHTML does not enable selection of the text in the form.
' DOCHOSTUIFLAG_DISABLE_HELP_MENU
'     MSHTML does not add the Help menu item to the container's menu.
' DOCHOSTUIFLAG_NO3DBORDER
'     MSHTML does not use 3-D borders on any frames or framesets. To turn the border off
'     on only the outer frameset use DOCHOSTUIFLAG_NO3DOUTERBORDER
' DOCHOSTUIFLAG_SCROLL_NO
'     MSHTML does not have scroll bars.
' DOCHOSTUIFLAG_DISABLE_SCRIPT_INACTIVE
'     MSHTML does not execute any script until fully activated. This flag is used to
'     postpone script execution until the host is active and, therefore, ready for script
'     to be executed.
' DOCHOSTUIFLAG_OPENNEWWIN
'     MSHTML opens a site in a new window when a link is clicked rather than browse to
'     the new site using the same browser window.
' DOCHOSTUIFLAG_DISABLE_OFFSCREEN
'     Not implemented.
' DOCHOSTUIFLAG_FLAT_SCROLLBAR
'     MSHTML uses flat scroll bars for any UI it displays.
' DOCHOSTUIFLAG_DIV_BLOCKDEFAULT
'     MSHTML inserts the div tag if a return is entered in edit mode. Without this flag,
'     MSHTML will use the p tag.
' DOCHOSTUIFLAG_ACTIVATE_CLIENTHIT_ONLY
'     MSHTML only becomes UI active if the mouse is clicked in the client area of the
'     window. It does not become UI active if the mouse is clicked on a non-client area,
'     such as a scroll bar.
' DOCHOSTUIFLAG_OVERRIDEBEHAVIORFACTORY
'     MSHTML consults the host before retrieving a behavior from the URL specified on
'     the page. If the host does not support the behavior, MSHTML does not proceed to
'     query other hosts or instantiate the behavior itself, even for behaviors developed
'     in script (HTML Components (HTCs)).
' DOCHOSTUIFLAG_CODEPAGELINKEDFONTS
'     Microsoft Internet Explorer 5 and later. Provides font selection compatibility
'     for Microsoft Outlook Express. If the flag is enabled, the displayed characters
'     are inspected to determine whether the current font supports the code page. If
'     disabled, the current font is used, even if it does not contain a glyph for the
'     character. This flag assumes that the user is using Internet Explorer 5 and
'     Outlook Express 4.0.
' DOCHOSTUIFLAG_URL_ENCODING_DISABLE_UTF8
'     Internet Explorer 5 and later. Controls how nonnative URLs are transmitted over
'     the Internet. Nonnative refers to characters outside the multibyte encoding of
'     the URL. If this flag is set, the URL is not submitted to the server in UTF-8 encoding.
' DOCHOSTUIFLAG_URL_ENCODING_ENABLE_UTF8
'     Internet Explorer 5 and later. Controls how nonnative URLs are transmitted over
'     the Internet. Nonnative refers to characters outside the multibyte encoding of
'     the URL. If this flag is set, the URL is submitted to the server in UTF-8 encoding.
' DOCHOSTUIFLAG_ENABLE_FORMS_AUTOCOMPLETE
'     Internet Explorer 5 and later. Enables the AutoComplete feature for forms in the
'     hosted browser. The Intelliforms feature is only turned on if the user has
'     previously enabled it. If the user has turned the AutoComplete feature off for
'     forms, it is off whether this flag is specified or not.
' DOCHOSTUIFLAG_ENABLE_INPLACE_NAVIGATION
'     Internet Explorer 5 and later. Enables the host to specify that navigation should
'     happen in place. This means that applications hosting MSHTML directly can specify
'     that navigation happen in the application's window. For instance, if this flag is
'     set, you can click a link in HTML mail and navigate in the mail instead of opening
'     a new Windows Internet Explorer window.
' DOCHOSTUIFLAG_IME_ENABLE_RECONVERSION
'     Internet Explorer 5 and later. During initialization, the host can set this flag
'     to enable Input Method Editor (IME) reconversion, allowing computer users to employ
'     IME reconversion while browsing Web pages. An input method editor is a program that
'     allows users to enter complex characters and symbols, such as Japanese Kanji
'     characters, using a standard keyboard. For more information, see the International
'     Features reference in the Base Services section of the Windows Software Development
'     Kit (SDK).
' DOCHOSTUIFLAG_THEME
'     Internet Explorer 6 and later. Specifies that the hosted browser should use themes
'     for pages it displays.
' DOCHOSTUIFLAG_NOTHEME
'     Internet Explorer 6 and later. Specifies that the hosted browser should not use
'     themes for pages it displays.
' DOCHOSTUIFLAG_NOPICS
'     Internet Explorer 6 and later. Disables PICS ratings for the hosted browser.
' DOCHOSTUIFLAG_NO3DOUTERBORDER
'     Internet Explorer 6 and later. Turns off any 3-D border on the outermost frame or
'     frameset only. To turn borders off on all frame sets, use DOCHOSTUIFLAG_NO3DBORDER
' DOCHOSTUIFLAG_DISABLE_EDIT_NS_FIXUP
'     Internet Explorer 6 and later. Disables the automatic correction of namespaces when
'     editing HTML elements.
' DOCHOSTUIFLAG_LOCAL_MACHINE_ACCESS_CHECK
'     Internet Explorer 6 and later. Prevents Web sites in the Internet zone from accessing
'     files in the Local Machine zone.
' DOCHOSTUIFLAG_DISABLE_UNTRUSTEDPROTOCOL
'     Internet Explorer 6 and later. Turns off untrusted protocols. Untrusted protocols
'     include ms-its, ms-itss, its, and mk:@msitstore.
' DOCHOSTUIFLAG_HOST_NAVIGATES
'     Internet Explorer 7. Indicates that navigation is delegated to the host; otherwise,
'     MSHTML will perform navigation. This flag is used primarily for non-HTML document types.
' DOCHOSTUIFLAG_ENABLE_REDIRECT_NOTIFICATION
'     Internet Explorer 7. Causes MSHTML to fire an additional DWebBrowserEvents2::BeforeNavigate2
'     event when redirect navigations occur. Applications hosting the WebBrowser Control
'     can choose to cancel or continue the redirect by returning an appropriate value in
'     the Cancel parameter of the event.
' DOCHOSTUIFLAG_USE_WINDOWLESS_SELECTCONTROL
'     Internet Explorer 7. Causes MSHTML to use the Document Object Model (DOM) to create
'     native "windowless" select controls that can be visually layered under other elements.
' DOCHOSTUIFLAG_USE_WINDOWED_SELECTCONTROL
'     Internet Explorer 7. Causes MSHTML to create standard Microsoft Win32 "windowed"
'     select and drop-down controls.
' DOCHOSTUIFLAG_ENABLE_ACTIVEX_INACTIVATE_MODE
'     Internet Explorer 6 for Windows XP Service Pack 2 (SP2) and later. Requires user
'     activation for Microsoft ActiveX controls and Java Applets embedded within a web page.
'     This flag enables interactive control blocking, which provisionally disallows direct
'     interaction with ActiveX controls loaded by the APPLET, EMBED, or OBJECT elements.
'     When a control is inactive, it does not respond to user input; however, it can perform
'     operations that do not involve interaction.
' DOCHOSTUIFLAG_DPI_AWARE
'     Internet Explorer 8. Causes layout engine to calculate document pixels as 96 dots
'     per inch (dpi). Normally, a document pixel is the same size as a screen pixel. This
'     flag is equivalent to setting the FEATURE_96DPI_PIXEL feature control key on a
'     per-host basis.
' Remarks
'     The DOCHOSTUIFLAG_BROWSER flag, a supplementary defined constant (not technically
'     a part of this enumeration), combines the values of DOCHOSTUIFLAG_DISABLE_HELP_MENU
'     and DOCHOSTUIFLAG_DISABLE_SCRIPT_INACTIVE.
' =====================================================================================
FUNCTION CDocHostUIHandler2.GetHostInfo (BYVAL pInfo AS DOCHOSTUIINFO PTR) AS HRESULT
   CWBX_DP("*** CDocHostUIHandler2.GetHostInfo pInfo = " & WSTR(pInfo))
   IF m_GetHostInfoProc THEN RETURN m_GetHostInfoProc(m_pWebCtx, pInfo)
   ' // Default behavior if the user does not processes this event...
   IF pInfo THEN
      pInfo->cbSize = SIZEOF(DOCHOSTUIINFO)
      pInfo->dwFlags = DOCHOSTUIFLAG_NO3DBORDER OR DOCHOSTUIFLAG_THEME OR DOCHOSTUIFLAG_DPI_AWARE
      pInfo->dwDoubleClick = DOCHOSTUIDBLCLK_DEFAULT
      pInfo->pchHostCss = NULL
      pInfo->pchHostNS = NULL
   END IF
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' Enables the host to replace MSHTML menus and toolbars.
' =====================================================================================
FUNCTION CDocHostUIHandler2.ShowUI (BYVAL dwID AS DWORD, BYVAL pActiveObject AS IOleInPlaceActiveObject PTR, _
BYVAL pCommandTarget AS IOleCommandTarget PTR, BYVAL pFrame AS IOleInPlaceFrame PTR, BYVAL pDoc AS IOleInPlaceUIWindow PTR) AS HRESULT
   CWBX_DP("CDocHostUIHandler2.ShowUI")
   IF m_ShowUIProc THEN RETURN m_ShowUIProc(m_pWebCtx, dwID,pActiveObject, pCommandTarget, pFrame, pDoc)
   RETURN S_FALSE
END FUNCTION
' =====================================================================================

' =====================================================================================
' Enables the host to remove its menus and toolbars.
' =====================================================================================
FUNCTION CDocHostUIHandler2.HideUI () AS HRESULT
   CWBX_DP("CDocHostUIHandler2.HideUI")
   IF m_HideUIProc THEN RETURN m_HideUIProc(m_pWebCtx)
   RETURN E_NOTIMPL
END FUNCTION
' =====================================================================================

' =====================================================================================
' Notifies the host that the command state has changed.
' =====================================================================================
FUNCTION CDocHostUIHandler2.UpdateUI () AS HRESULT
   CWBX_DP("CDocHostUIHandler2.UpdateUI")
   IF m_UpdateUIProc THEN RETURN m_UpdateUIProc(m_pWebCtx)
   RETURN E_NOTIMPL
END FUNCTION
' =====================================================================================

' =====================================================================================
' Called by the MSHTML implementation of IOleInPlaceActiveObject::EnableModeless.
' Also called when MSHTML displays a modal UI.
' =====================================================================================
FUNCTION CDocHostUIHandler2.EnableModeless (BYVAL fEnable AS WINBOOL) AS HRESULT
   CWBX_DP("CDocHostUIHandler2.EnableModeless")
   IF m_EnableModelessProc THEN RETURN m_EnableModelessProc(m_pWebCtx, fEnable)
   RETURN E_NOTIMPL
END FUNCTION
' =====================================================================================

' =====================================================================================
' Called by the MSHTML implementation of IOleInPlaceActiveObject::OnDocWindowActivate.
' =====================================================================================
FUNCTION CDocHostUIHandler2.OnDocWindowActivate (BYVAL fActivate AS WINBOOL) AS HRESULT
   CWBX_DP("CDocHostUIHandler2.OnDocWindowActivate")
   IF m_OnDocWindowActivateProc THEN RETURN m_OnDocWindowActivateProc(m_pWebCtx, fActivate)
   RETURN E_NOTIMPL
END FUNCTION
' =====================================================================================

' =====================================================================================
' Called by the MSHTML implementation of IOleInPlaceActiveObject::OnFrameWindowActivate.
' =====================================================================================
FUNCTION CDocHostUIHandler2.OnFrameWindowActivate (BYVAL fActivate AS WINBOOL) AS HRESULT
   CWBX_DP("CDocHostUIHandler2.OnFrameWindowActivate")
   IF m_OnFrameWindowActivateProc THEN RETURN m_OnFrameWindowActivateProc(m_pWebCtx, fActivate)
   RETURN E_NOTIMPL
END FUNCTION
' =====================================================================================

' =====================================================================================
' Called by the MSHTML implementation of IOleInPlaceActiveObject::ResizeBorder.
' =====================================================================================
FUNCTION CDocHostUIHandler2.ResizeBorder (BYVAL prcBorder AS LPCRECT, BYVAL pUIWindow AS IOleInPlaceUIWindow PTR, BYVAL frameWindow AS WINBOOL) AS HRESULT
   CWBX_DP("CDocHostUIHandler2.ResizeBorder")
   IF m_ResizeBorderProc THEN RETURN m_ResizeBorderProc(m_pWebCtx, prcBorder, pUIWindow, frameWindow)
   RETURN E_NOTIMPL
END FUNCTION
' =====================================================================================

' =====================================================================================
' Called by MSHTML when IOleInPlaceActiveObject.TranslateAccelerator or
' IOleControlSite.TranslateAccelerator is called.
' =====================================================================================
FUNCTION CDocHostUIHandler2.TranslateAccelerator (BYVAL lpMsg AS LPMSG, BYVAL pguidCmdGroup AS const GUID PTR, BYVAL nCmdID AS DWORD) AS HRESULT
   CWBX_DP("CDocHostUIHandler2.TranslateAccelerator")
   IF m_TranslateAcceleratorProc THEN RETURN m_TranslateAcceleratorProc(m_pWebCtx, lpMsg, pguidCmdGroup, nCmdID)
   RETURN S_FALSE
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets a registry subkey path that overrides the default Windows Internet Explorer
' registry settings.
' =====================================================================================
FUNCTION CDocHostUIHandler2.GetOptionKeyPath (BYVAL pchKey AS LPOLESTR PTR, BYVAL dw AS DWORD) AS HRESULT
   CWBX_DP("CDocHostUIHandler2.GetOptionKeyPath")
   IF m_GetOptionKeyPathProc THEN RETURN m_GetOptionKeyPathProc(m_pWebCtx, pchKey, dw)
   RETURN E_NOTIMPL
END FUNCTION
' =====================================================================================

' =====================================================================================
' Enables the host to supply an alternative IDropTarget interface.
' =====================================================================================
FUNCTION CDocHostUIHandler2.GetDropTarget (BYVAL pDropTarget AS IDropTarget PTR, BYVAL ppDropTarget AS IDropTarget PTR PTR) AS HRESULT
   CWBX_DP("CDocHostUIHandler2.GetDropTarget")
   IF m_GetDropTargetProc THEN RETURN m_GetDropTargetProc(m_pWebCtx, pDropTarget, ppDropTarget)
   RETURN E_NOTIMPL
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets the host's IDispatch interface.
' =====================================================================================
FUNCTION CDocHostUIHandler2.GetExternal (BYVAL ppDispatch AS IDispatch PTR PTR) AS HRESULT
   CWBX_DP("CDocHostUIHandler2.GetExternal")
   IF m_GetExternalProc THEN RETURN m_GetExternalProc(m_pWebCtx, ppDispatch)
   *ppDispatch = NULL
   RETURN S_FALSE
END FUNCTION
' =====================================================================================

' =====================================================================================
' Enables the host to modify the URL to be loaded.
' =====================================================================================
FUNCTION CDocHostUIHandler2.TranslateUrl (BYVAL dwTranslate AS DWORD, BYVAL pchURLIn AS LPWSTR, BYVAL ppchURLOut AS LPWSTR PTR) AS HRESULT
   CWBX_DP("CDocHostUIHandler2.TranslateUrl")
   IF m_TranslateUrlProc THEN RETURN m_TranslateUrlProc(m_pWebCtx, dwTranslate, pchURLIn, ppchURLOut)
   RETURN S_FALSE
END FUNCTION
' =====================================================================================

' =====================================================================================
' Enables the host to replace the MSHTML data object.
' =====================================================================================
FUNCTION CDocHostUIHandler2.FilterDataObject (BYVAL pDO AS IDataObject PTR, BYVAL ppDORet AS IDataObject PTR PTR) AS HRESULT
   CWBX_DP("CDocHostUIHandler2.FilterDataObject")
   IF m_FilterDataObjectProc THEN RETURN m_FilterDataObjectProc(m_pWebCtx, pDO, ppDORet)
   *ppDORet = NULL
   RETURN S_FALSE
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets a registry subkey path that modifies Windows Internet Explorer user preferences.
' =====================================================================================
FUNCTION CDocHostUIHandler2.GetOverrideKeyPath (BYVAL pchKey AS LPOLESTR PTR, BYVAL dw AS DWORD) AS HRESULT
   CWBX_DP("CDocHostUIHandler2.GetOverrideKeyPath")
   IF m_GetOverrideKeyPathProc THEN RETURN m_GetOverrideKeyPathProc(m_pWebCtx, pchKey, dw)
   *pchKey = NULL
   RETURN E_NOTIMPL
END FUNCTION
' =====================================================================================
