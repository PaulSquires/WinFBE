# CWebCtx Class

Wrapper class to embed instances of the **WebBrowser** control in our application using an OLE Container (**CAxHost** class) to host it.

It also provides methods to connect and disconnect to the events fired by the **WebBrowser** control, set event handlers (pointers to callback procedures) for both the **DWebBrowser2** and **IDocHostUIHandler2** interfaces, a method to navigate to a URL, and function to get references to the OLE Container class and the **IWebBrowser2** interface.

The file **AfxExDisp.bi** provides declarations to call the methods of the **WebBrowser** interfaces using abstract methods.

The **WebBrowser** events sink class is provided in the file **CWebBrowserEvents.inc**, and the **DocHostUIHandler** events sink class is provided in the file **CDocHostUIHandler2.inc**.

**Include file**: CWebCtx.inc.

### Constructor

| Name       | Description |
| ---------- | ----------- |
| [CAXHOST_AMBIENTDISP structure](#CAXHOST_AMBIENTDISP) | Contains information the ambient properties of the **CAxHost** control. |
| [Constructor](#Constructor) | Creates an instance of the OLE container using a ProgId. |

### Methods

| Name       | Description |
| ---------- | ----------- |
| [Advise](#Advise) | Establishes a connection between a connection point object and the client's sink. |
| [AxHostPtr](#AxHostPtr) | Returns a pointer to the **CAxHost** class. |
| [BrowserPtr](#BrowserPtr) | Returns a direct pointer to the **IWebBrowser2** interface of the hosted **WebBrowser** control. |
| [Busy](#Busy) | Returns a boolean value indicating whether the object is engaged in a navigation or downloading operation. |
| [CtrlID](#CtrlID) | Returns the identifier of the container's window. |
| [Document](#Document) | Retrieves the automation object of the active document, if any. |
| [ExecWB](#ExecWB) | Executes a command on an OLE object and returns the status of the command execution using the **IOleCommandTarget** interface. |
| [Find](#Find) | Activates the find dialog. |
| [GetActiveElementId](#GetActiveElementId) | Retrieves the ID of the active element (the object that has the focus when the parent document has focus). |
| [GetBodyInnerHtml](#GetBodyInnerHtml) | Returns a string that represents the text and html elements between the start and end body tags. |
| [GetBodyInnerText](#GetBodyInnerText) | Returns a string that represents the text between the start and end body tags without any associated HTML. |
| [GoBack](#GoBack) | Navigates backward one item in the history list. |
| [GetElementInnerHtmlById](#GetElementInnerHtmlById) | Retrieves the HTML between the start and end tags of the object. |
| [GetElementValueById](#GetElementValueById) | Retrieves the value attribute of the specified attribute. |
| [GoForward](#GoForward) | Navigates forward one item in the history list. |
| [GoHome](#GoHome) | Navigates to the current home or start page. |
| [GoSearch](#GoSearch) | Navigates to the current search page. |
| [hWindow](#hWindow) | Returns the handle of the OLE Container hosting window. |
| [InternetOptions](#InternetOptions) | Activates the Internet options dialog. |
| [LocationName](#LocationName) | Retrieves the name of the resource that Microsoft Internet Explorer is currently displaying. |
| [LocationURL](#LocationURL) | Retrieves the URL of the resource that Microsoft Internet Explorer is currently displaying. |
| [Navigate](#Navigate) | Returns the handle of the OLE Container hosting window. |
| [PageProperties](#PageProperties) | Activates the properties dialog. |
| [PageSetup](#PageSetup) | Activates the page setup dialog. |
| [PrintPage](#PrintPage) | Activates the print dialog. |
| [PrintPreview](#PrintPreview) | Activates the print preview dialog. |
| [QueryStatusWB](#QueryStatusWB) | Queries the OLE object for the status of commands using the **QueryStatus** method of the **IOleCommandTarget** interface.|
| [ReadyState](#ReadyState) | Retrieves the ready state of the object. |
| [Refresh](#Refresh) | Reloads the file that is currently displayed in the object. |
| [Refresh2](#Refresh2) | Reloads the file that is currently displayed in the object. Unlike Refresh, this method contains a parameter that specifies the refresh level. |
| [RegisterAsBrowser](#RegisterAsBrowser) | Sets or retrieves a value that indicates whether the object is registered as a top-level browser for target name resolution. |
| [RegisterAsDropTarget](#RegisterAsDropTarget) | Sets or retrieves a value that indicates whether the object is registered as a drop target for navigation. |
| [SaveAs](#SaveAs) | Activates the save file dialog. |
| [SetElementFocusById](#SetElementFocusById) | Sets the HTML between the start and end tags of the object. |
| [SetElementInnerHtmlById](#SetElementInnerHtmlById) | Sets the HTML between the start and end tags of the object. |
| [SetElementValueById](#SetElementValueById) | Sets the value attribute of the specified identifier. |
| [SetEventProc](#SetEventProc) | Sets pointers to user implemented callback procedures to receive events of the hosted WebBrowser control. |
| [SetFocus](#SetFocus) | Sets the focus in the hosted document (usually an html page). |
| [SetUIEventProc](#SetUIEventProc) | Sets pointers to user implemented callback procedures to receive events of the **IDocHostUIHandler2** interface. |
| [SetUIHandler](#SetUIHandler) | Sets our implementation of the **IDocHostUIHandler** interface to customize the WebBrowser. |
| [ShowSource](#ShowSource) | Displays the source code of the page in an instance of **NotePad**. |
| [Silent](#Silent) | Sets or gets a value that indicates whether the object can display dialog boxes. |
| [Stop](#Stop) | Cancels any pending navigation or download operation and stops any dynamic page elements, such as background sounds and animations. |
| [Unadvise](#Unadvise) | Releases the events connection. |
| [WaitForPageLoad](#WaitForPageLoad) | Waits until the page had been fully downloaded or te timeout has expired. |
| [WriteHtml](#WriteHtml) | Writes one or more HTML expressions to a document. |

### Helper Procedures

| Name       | Description |
| ---------- | ----------- |
| [AfxGetBrowserPtr](#AfxGetBrowserPtr) | Returns a pointer to the hosted WebBrowser control given the handle of the form, or any control in the form, and the control identifier. |
| [AfxGetActiveElementId](#AfxGetActiveElementId) | Retrieves the ID of the active element (the object that has the focus when the parent document has focus). |
| [AfxGetBodyInnerHtml](#AfxGetBodyInnerHtml) | Returns a string that represents the text and html elements between the start and end body tags. |
| [AfxGetBodyInnerText](#AfxGetBodyInnerText) | Returns a string that represents the text between the start and end body tags without any associated HTML. |
| [AfxGetElementInnerHtmlById](#AfxGetElementInnerHtmlById) | Retrieves the HTML between the start and end tags of the object. |
| [AfxGetElementValueById](#AfxGetElementValueById) | Retrieves the value attribute of the specified attribute. |
| [AfxSetElementFocusById](#AfxSetElementFocusById) | Sets the focus in the specified element. |
| [AfxSetElementInnerHtmlById](#AfxSetElementInnerHtmlById) | Sets the HTML between the start and end tags of the object. |
| [AfxSetElementValueById](#AfxSetElementValueById) | Sets the value attribute of the specified identifier. |
| [AfxWriteHtml](#AfxWriteHtml) | Writes one or more HTML expressions to a document. |

# CWebBrowserEvents Class

The **CWebBrowserEvents** class implements an event sink interface to receive event notifications from the **WebBrowser** control.

The constructor of the **CWebCtx** class creates automatically an instance of this class and establishes a connection with the hosted control calling the **Advise** method.

To set pointers to the callback functions of the events in which you are interested use the SetEventProc of the **CWebCtx** class.

All these callback functions will receive a pointer to the **CwebCtx** class as the first parameter. With this pointer, we can get:

* Access to all the methods of the CWebCtx class.
* The handle of the container window calling pCWbCtx->hWindow.
* A pointer to the hosted WebBrowser control calling pWebCtx->BrowserPtr.
* A pointer to the CAxHost class calling pWebCtx->HostPtr.

| Name       | Description |
| ---------- | ----------- |
| [BeforeNavigate2](#BeforeNavigate2) | Fires before navigation occurs in the given object (on either a window or frameset element). |
| [ClientToHostWindow](#ClientToHostWindow) | Fires to request that the client window size be converted to the host window size. |
| [CommandStateChange](#CommandStateChange) | Fires when the enabled state of a command changes. |
| [DocumentComplete](#DocumentComplete) | Fires when a document has been completely loaded and initialized. |
| [DownloadBegin](#DownloadBegin) | Fires when a navigation operation is beginning. |
| [DownloadComplete](#DownloadComplete) | Fires when a navigation operation finishes, is halted, or fails. |
| [FileDownload](#FileDownload) | Fires to indicate that a file download is about to occur. If a file download dialog is to be displayed, this event is fired prior to the display of the dialog. |
| [HtmlDocumentEventsProc](#HtmlDocumentEventsProc) | Provides access to properties and methods exposed by an object. |
| [NavigateComplete2](#NavigateComplete2) | Fires after a navigation to a link is completed on either a window or frameSet element. |
| [NavigateError](#NavigateError) | Fires when an error occurs during navigation. |
| [NewWindow2](#NewWindow2) | Raised when a new window is to be created. |
| [NewWindow3](#NewWindow3) | Raised when a new window is to be created. Extends **NewWindow2** with additional information about the new window. |
| [OnVisible](#OnVisible) | Fires when the **Visible** property of the object is changed. |
| [PrintTemplateInstantiation](#PrintTemplateInstantiation) | Fires when a print template has been instantiated. |
| [PrintTemplateTeardown](#PrintTemplateTeardown) | Fires when a print template has been destroyed. |
| [PrivacyImpactedStateChange](#PrivacyImpactedStateChange) | Fired when an event occurs that impacts privacy or when a user navigates away from a URL that has impacted privacy. |
| [ProgressChange](#ProgressChange) | Fires when the progress of a download operation is updated on the object. |
| [PropertyChange](#PropertyChange) | Fires when the **PutProperty** method of the object changes the value of a property. |
| [SetSecureLockIcon](#SetSecureLockIcon) | Fires when there is a change in encryption level. |
| [StatusTextChange](#StatusTextChange) | Fires when the status bar text of the object has changed. |
| [TitleChange](#TitleChange) | Fires when the title of a document in the object becomes available or changes. |
| [WindowClosing](#WindowClosing) | Fires when the window of the object is about to be closed by script. |
| [WindowSetHeight](#WindowSetHeight) | Fires when the object changes its height. |
| [WindowSetLeft](#WindowSetLeft) | Fires when the object changes its left position. |
| [WindowSetResizable](#WindowSetResizable) | Fires to indicate whether the host window should allow or disallow resizing of the object. |
| [WindowSetTop](#WindowSetTop) | Fires when the object changes its top position. |
| [WindowSetWidth](#WindowSetWidth) | Fires when the object changes its width. |
| [WindowStateChanged](#WindowStateChanged) | Fires when the progress of a download operation is updated on the object. |

# CDocHostUIHandler Class

The **CDocHostUIHandler** class implements the **IDocHostUIHandler** interface, that enables an application that is hosting the WebBrowser control or automating Windows Internet Explorer to replace the menus, toolbars, and context menus used by MSHTML.

The constructor of the **CWebCtx** class creates automatically an instance of this class, and its destructor deletes it.

To set pointers to the callback functions of the events in which you are interested use the **SetUIEventProc** of the **CWebCtx** class.

All these callback functions will receive a pointer to the **CWebCtx** class as the first parameter. With this pointer, we can get:

* Access to all the methods of the **CWebCtx** class.
* The handle of the container window calling **CWebCtx->hWindow**.
* A pointer to the hosted WebBrowser control calling **CWebCtx->BrowserPtr**.
* A pointer to the CAxHost class calling **CWebCtx->HostPtr**.

| Name       | Description |
| ---------- | ----------- |
| [EnableModeless](#EnableModeless) | Called by the MSHTML implementation of **IOleInPlaceActiveObject.EnableModeless**. Also called when MSHTML displays a modal UI. |
| [FilterDataObject](#FilterDataObject) | Called by MSHTML to allow the host to replace the MSHTML data object. |
| [GetDropTarget](#GetDropTarget) | Called by MSHTML when it is used as a drop target. This method enables the host to supply an alternative **IDropTarget** interface. |
| [GetExternal](#GetExternal) | Called by MSHTML to obtain the host's **IDispatch** interface. |
| [GetHostInfo](#GetHostInfo) | Called by MSHTML to retrieve the user interface (UI) capabilities of the application that is hosting MSHTML. |
| [GetOptionKeyPath](#GetOptionKeyPath) | Called by the WebBrowser Control to retrieve a registry subkey path that overrides the default Microsoft Internet Explorer registry settings. |
| [GetOverrideKeyPath](#GetOverrideKeyPath) | Called by the WebBrowser Control to retrieve a registry subkey path that modifies Microsoft Internet Explorer user preferences. |
| [HideUI](#HideUI) | Called when MSHTML removes its menus and toolbars. |
| [OnDocWindowActivate](#OnDocWindowActivate) | Called by the MSHTML implementation of **IOleInPlaceActiveObject.OnDocWindowActivate**. |
| [OnFrameWindowActivate](#OnFrameWindowActivate) | Called by the MSHTML implementation of **IOleInPlaceActiveObject.OnFrameWindowActivate**. |
| [ResizeBorder](#ResizeBorder) | Called by the MSHTML implementation of **IOleInPlaceActiveObject.ResizeBorder**. |
| [ShowContextMenu](#ShowContextMenu) | Called by MSHTML to display a shortcut menu. |
| [ShowUI](#ShowUI) | Called by MSHTML to enable the host to replace MSHTML menus and toolbars. |
| [TranslateAccelerator](#TranslateAccelerator) | Called by MSHTML when **IOleInPlaceActiveObject.TranslateAccelerator** or **IOleControlSite.TranslateAccelerator** is called. |
| [TranslateUrl](#TranslateUrl) | Called by MSHTML to give the host an opportunity to modify the URL to be loaded. |
| [UpdateUI](#UpdateUI) | Called by MSHTML to notify the host that the command state has changed. |

# CHTMLDocumentEvents2 Class

The **CHTMLDocumentEvents2** class implements the IHTMLDocumentEvents2 interface, that enables an application that is hosting the WebBrowser control to intercept events fired by a document object,

The **DocumentComplete** event of the **CWebBrowserEvents** class connects automatically with the **HTMLDocumentEvents2** interface of the loaded page if the user has set a pointer to a callback procedure using the SetEventProc method of the **CWebCtx** class.

The callback procedure will receive three parameters, *pWebCtx*, *dispid* and *pEvetObj*.

With *pWebCtx*, which is a pointer to the **CWebCtx** clas, we can get:

* Access to all the methods of the CWebCtx class.
* The handle of the container window calling pCWbCtx->hWindow.
* A pointer to the hosted WebBrowser control calling pWebCtx->BrowserPtr.
* A pointer to the CAxHost class calling pWebCtx->HostPtr.

With *dispid* we can identify which event we are being notified:

```
' // DispIds of the events
CONST DISPID_HTMLELEMENTEVENTS2_ONHELP             = -2147418102
CONST DISPID_HTMLELEMENTEVENTS2_ONHELP             = -2147418102
CONST DISPID_HTMLELEMENTEVENTS2_ONCLICK            = -600
CONST DISPID_HTMLELEMENTEVENTS2_ONDBLCLICK         = -601
CONST DISPID_HTMLELEMENTEVENTS2_ONKEYPRESS         = -603
CONST DISPID_HTMLELEMENTEVENTS2_ONKEYDOWN          = -602
CONST DISPID_HTMLELEMENTEVENTS2_ONKEYUP            = -604
CONST DISPID_HTMLELEMENTEVENTS2_ONMOUSEOUT         = -2147418103
CONST DISPID_HTMLELEMENTEVENTS2_ONMOUSEOVER        = -2147418104
CONST DISPID_HTMLELEMENTEVENTS2_ONMOUSEMOVE        = -606
CONST DISPID_HTMLELEMENTEVENTS2_ONMOUSEDOWN        = -605
CONST DISPID_HTMLELEMENTEVENTS2_ONMOUSEUP          = -607
CONST DISPID_HTMLELEMENTEVENTS2_ONSELECTSTART      = -2147418100
CONST DISPID_HTMLELEMENTEVENTS2_ONFILTERCHANGE     = -2147418095
CONST DISPID_HTMLELEMENTEVENTS2_ONDRAGSTART        = -2147418101
CONST DISPID_HTMLELEMENTEVENTS2_ONBEFOREUPDATE     = -2147418108
CONST DISPID_HTMLELEMENTEVENTS2_ONAFTERUPDATE      = -2147418107
CONST DISPID_HTMLELEMENTEVENTS2_ONERRORUPDATE      = -2147418099
CONST DISPID_HTMLELEMENTEVENTS2_ONROWEXIT          = -2147418106
CONST DISPID_HTMLELEMENTEVENTS2_ONROWENTER         = -2147418105
CONST DISPID_HTMLELEMENTEVENTS2_ONDATASETCHANGED   = -2147418098
CONST DISPID_HTMLELEMENTEVENTS2_ONDATAAVAILABLE    = -2147418097
CONST DISPID_HTMLELEMENTEVENTS2_ONDATASETCOMPLETE  = -2147418096
CONST DISPID_HTMLELEMENTEVENTS2_ONLOSECAPTURE      = -2147418094
CONST DISPID_HTMLELEMENTEVENTS2_ONPROPERTYCHANGE   = -2147418093
CONST DISPID_HTMLELEMENTEVENTS2_ONSCROLL           = 1014
CONST DISPID_HTMLELEMENTEVENTS2_ONFOCUS            = -2147418111
CONST DISPID_HTMLELEMENTEVENTS2_ONBLUR             = -2147418112
CONST DISPID_HTMLELEMENTEVENTS2_ONRESIZE           = 1016
CONST DISPID_HTMLELEMENTEVENTS2_ONDRAG             = -2147418092
CONST DISPID_HTMLELEMENTEVENTS2_ONDRAGEND          = -2147418091
CONST DISPID_HTMLELEMENTEVENTS2_ONDRAGENTER        = -2147418090
CONST DISPID_HTMLELEMENTEVENTS2_ONDRAGOVER         = -2147418089
CONST DISPID_HTMLELEMENTEVENTS2_ONDRAGLEAVE        = -2147418088
CONST DISPID_HTMLELEMENTEVENTS2_ONDROP             = -2147418087
CONST DISPID_HTMLELEMENTEVENTS2_ONBEFORECUT        = -2147418083
CONST DISPID_HTMLELEMENTEVENTS2_ONCUT              = -2147418086
CONST DISPID_HTMLELEMENTEVENTS2_ONBEFORECOPY       = -2147418082
CONST DISPID_HTMLELEMENTEVENTS2_ONCOPY             = -2147418085
CONST DISPID_HTMLELEMENTEVENTS2_ONBEFOREPASTE      = -2147418081
CONST DISPID_HTMLELEMENTEVENTS2_ONPASTE            = -2147418084
CONST DISPID_HTMLELEMENTEVENTS2_ONCONTEXTMENU      = 1023
CONST DISPID_HTMLELEMENTEVENTS2_ONROWSDELETE       = -2147418080
CONST DISPID_HTMLELEMENTEVENTS2_ONROWSINSERTED     = -2147418079
CONST DISPID_HTMLELEMENTEVENTS2_ONCELLCHANGE       = -2147418078
CONST DISPID_HTMLELEMENTEVENTS2_ONREADYSTATECHANGE = -609
CONST DISPID_HTMLELEMENTEVENTS2_ONLAYOUTCOMPLETE   = 1030
CONST DISPID_HTMLELEMENTEVENTS2_ONPAGE             = 1031
CONST DISPID_HTMLELEMENTEVENTS2_ONMOUSEENTER       = 1042
CONST DISPID_HTMLELEMENTEVENTS2_ONMOUSELEAVE       = 1043
CONST DISPID_HTMLELEMENTEVENTS2_ONACTIVATE         = 1044
CONST DISPID_HTMLELEMENTEVENTS2_ONDEACTIVATE       = 1045
CONST DISPID_HTMLELEMENTEVENTS2_ONBEFOREDEACTIVATE = 1034
CONST DISPID_HTMLELEMENTEVENTS2_ONBEFOREACTIVATE   = 1047
CONST DISPID_HTMLELEMENTEVENTS2_ONFOCUSIN          = 1048
CONST DISPID_HTMLELEMENTEVENTS2_ONFOCUSOUT         = 1049
CONST DISPID_HTMLELEMENTEVENTS2_ONMOVE             = 1035
CONST DISPID_HTMLELEMENTEVENTS2_ONCONTROLSELECT    = 1036
CONST DISPID_HTMLELEMENTEVENTS2_ONMOVESTART        = 1038
CONST DISPID_HTMLELEMENTEVENTS2_ONMOVEEND          = 1039
CONST DISPID_HTMLELEMENTEVENTS2_ONRESIZESTART      = 1040
CONST DISPID_HTMLELEMENTEVENTS2_ONRESIZEEND        = 1041
CONST DISPID_HTMLELEMENTEVENTS2_ONMOUSEWHEEL       = 1033
```

See: [HTMLDocumentEvents2 interface](https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa769764(v=vs.85))

*pEvtObj* is a pointer to the **IHTMLEventObj** interface. This interface provides access to the event processes, such as the element in which the event occurred, the state of the keyboard keys, the location of the mouse, and the state of the mouse buttons.

See: [IHTMLEventObj interface](https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa703876(v=vs.85))

### Example

```
' ########################################################################################
' Microsoft Windows
' Contents: WebBrowser - Events
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "Afx/CVAR.inc"
#INCLUDE ONCE "Afx/CAxHost/CWebCtx.inc"
USING Afx

CONST IDC_WEBBROWSER = 1001
CONST IDC_SATUSBAR = 1002

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG

   END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT
DECLARE SUB WebBrowser_StatusTextChangeProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pwszText AS WSTRING PTR)
DECLARE SUB WebBrowser_DocumentCompleteProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR, BYVAL vUrl AS VARIANT PTR)
DECLARE SUB WebBrowser_BeforeNavigate2Proc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR, _
    BYVAL vUrl AS VARIANT PTR, BYVAL Flags AS VARIANT PTR, BYVAL TargetFrameName AS VARIANT PTR, _
    BYVAL PostData AS VARIANT PTR, BYVAL Headers AS VARIANT PTR, BYVAL pbCancel AS VARIANT_BOOL PTR)
DECLARE FUNCTION WebBrowser_HtmlDocumentEventsProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL dispId AS LONG, BYVAL pEvtObj AS IHTMLEventObj PTR) AS BOOLEAN
DECLARE FUNCTION DocHostUI_ShowContextMenuProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL dwID AS DWORD, BYVAL ppt AS POINT PTR, BYVAL pcmdtReserved AS IUnknown PTR, BYVAL pdispReserved AS IDispatch PTR) AS HRESULT
DECLARE FUNCTION DocHostUI_GetHostInfo (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pInfo AS DOCHOSTUIINFO PTR) AS HRESULT
DECLARE FUNCTION DocHostUI_TranslateAccelerator (BYVAL pWebCtx AS CWebCtx PTR, BYVAL lpMsg AS LPMSG, BYVAL pguidCmdGroup AS const GUID PTR, BYVAL nCmdID AS DWORD) AS HRESULT

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
   DIM hwndMain AS HWND = pWindow.Create(NULL, "Embedded WebBrowser control", @WndProc)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(750, 450)
   ' // Centers the window
   pWindow.Center

   ' // Add a status bar
   DIM hStatusbar AS HWND = pWindow.AddControl("Statusbar", , IDC_SATUSBAR)

   ' // Add a WebBrowser control
   DIM pwb AS CWebCtx = CWebCtx(@pWindow, IDC_WEBBROWSER, 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   ' // Set web browser event callback procedures
   pwb.SetEventProc("StatusTextChange", @WebBrowser_StatusTextChangeProc)
   pwb.SetEventProc("DocumentComplete", @WebBrowser_DocumentCompleteProc)
   pwb.SetEventProc("BeforeNavigate2", @WebBrowser_BeforeNavigate2Proc)
   pwb.SetEventProc("HtmlDocumentEvents", @WebBrowser_HtmlDocumentEventsProc)
   ' // Set DocHostUI event callback procedures
   pwb.SetUIEventProc("ShowContextMenu", @DocHostUI_ShowContextMenuProc)
   pwb.SetUIEventProc("GetHostInfo", @DocHostUI_GetHostInfo)
   pwb.SetUIEventProc("TranslateAccelerator", @DocHostUI_TranslateAccelerator)

   ' // Navigate to a URL
   pwb.Navigate "http://com.it-berater.org/"
   ' // Set the focus in the page (the page must be fully loaded)
   pwb.SetFocus

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
            ' // Resize the status bar
            DIM hStatusBar AS HWND = GetDlgItem(hwnd, IDC_SATUSBAR)
            SendMessage hStatusBar, uMsg, wParam, lParam
            ' // Calculate the size of the status bar
            DIM StatusBarHeight AS DWORD, rc AS RECT
            GetWindowRect hStatusBar, @rc
            StatusBarHeight = rc.Bottom - rc.Top
            ' // Retrieve a pointer to the CWindow class
            DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
            ' // Move the position of the control
            IF pWindow THEN pWindow->MoveWindow GetDlgItem(hwnd, IDC_WEBBROWSER), _
               0, 0, pWindow->ClientWidth, pWindow->ClientHeight - StatusBarHeight / pWindow->ryRatio, CTRUE
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

' ========================================================================================
' Process the WebBrowser StatusTextChange event.
' ========================================================================================
SUB WebBrowser_StatusTextChangeProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pwszText AS WSTRING PTR)
   IF pwszText THEN StatusBar_SetText(GetDlgItem(GetParent(pWebCtx->hWindow), IDC_SATUSBAR), 0, pwszText)
END SUB
' ========================================================================================

' ========================================================================================
' Process the WebBrowser DocumentComplete event.
' ========================================================================================
SUB WebBrowser_DocumentCompleteProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR, BYVAL vUrl AS VARIANT PTR)
   ' // The vUrl parameter is a VT_BYREF OR VT_BSTR variant
   ' // It can be a VT_BSTR variant or a VT_ARRAY OR VT_UI1 with a pidl
   DIM varUrl AS VARIANT
   VariantCopyInd(@varUrl, vUrl)
   StatusBar_SetText(GetDlgItem(GetParent(pWebCtx->hWindow), IDC_SATUSBAR), 0, "Document complete: " & AfxVarToStr(@varUrl))
   VariantClear(@varUrl)
END SUB
' ========================================================================================

' ========================================================================================
' Process the IDocHostUIHandler ShowContextMenu event.
' ========================================================================================
FUNCTION DocHostUI_ShowContextMenuProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL dwID AS DWORD, BYVAL ppt AS POINT PTR, BYVAL pcmdtReserved AS IUnknown PTR, BYVAL pdispReserved AS IDispatch PTR) AS HRESULT
   ' // This event notifies that the user has clicked the right mouse button to show the
   ' // context menu. We can anulate it returning %S_OK and show our context menu.
   ' // Do not allow to show the context menu
'   AfxMsg "Sorry! Context menu disabled"
'   RETURN S_OK
   ' // Host did not display its UI. MSHTML will display its UI.
   RETURN S_FALSE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Process the IDocHostUIHandler GetHostInfo event.
' ========================================================================================
PRIVATE FUNCTION DocHostUI_GetHostInfo (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pInfo AS DOCHOSTUIINFO PTR) AS HRESULT
   IF pInfo THEN
      pInfo->cbSize = SIZEOF(DOCHOSTUIINFO)
      pInfo->dwFlags = DOCHOSTUIFLAG_NO3DBORDER OR DOCHOSTUIFLAG_THEME OR DOCHOSTUIFLAG_DPI_AWARE
      pInfo->dwDoubleClick = DOCHOSTUIDBLCLK_DEFAULT
      pInfo->pchHostCss = NULL
      pInfo->pchHostNS = NULL
   END IF
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Process the IDocHostUIHandler TranslateAccelerator event.
' ========================================================================================
PRIVATE FUNCTION DocHostUI_TranslateAccelerator (BYVAL pWebCtx AS CWebCtx PTR, BYVAL lpMsg AS LPMSG, BYVAL pguidCmdGroup AS const GUID PTR, BYVAL nCmdID AS DWORD) AS HRESULT

   IF lpMsg->message = WM_KEYDOWN THEN
      pWebCtx->SetElementInnerHtmlById "output", "ID: " & pWebCtx->GetActiveElementId & " KeyDown - Key: " & WSTR(lpMsg->wParam)
   END IF

   ' // When you use accelerator keys such as TAB, you may need to override the
   ' // default host behavior. The example shows how to do this.
    IF lpMsg->message = WM_KEYDOWN AND lpMsg->wParam = VK_TAB THEN
       RETURN S_FALSE   ' S_OK to disable tab navigation
    END IF
   ' // Return S_FALSE if you don't process the message
   RETURN S_FALSE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Fires before navigation occurs in the given object (on either a window or frameset element).
' ========================================================================================
SUB WebBrowser_BeforeNavigate2Proc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR, _
    BYVAL vUrl AS VARIANT PTR, BYVAL Flags AS VARIANT PTR, BYVAL TargetFrameName AS VARIANT PTR, _
    BYVAL PostData AS VARIANT PTR, BYVAL Headers AS VARIANT PTR, BYVAL pbCancel AS VARIANT_BOOL PTR)

   OutputDebugStringW AfxVarToStr(vUrl)

   ' // Sample code to redirect navigation to another url
   IF AfxVarToStr(vUrl) = "http://com.it-berater.org/" THEN
      ' // Stop loading the page
      pWebCtx->BrowserPtr->Stop
      ' // Cancel the navigation operation
      *pbCancel = VARIANT_TRUE
      ' // Navigate to another new url
      DIM cvNewUrl AS CVAR = "http://www.planetsquires.com/protect/forum/index.php"
      pWebCtx->BrowserPtr->Navigate2(cvNewUrl)
   END IF

END SUB
' ========================================================================================

' ========================================================================================
' For cancelable document events return TRUE to indicate that Internet Explorer should
' perform its own event processing or FALSE to cancel the event.
' ========================================================================================
PRIVATE FUNCTION WebBrowser_HtmlDocumentEventsProc (BYVAL pWebCtx AS CWebCtx PTR, _
   BYVAL dispid AS LONG, BYVAL pEvtObj AS IHTMLEventObj PTR) AS BOOLEAN

'   SELECT CASE dispid
'
'      CASE DISPID_HTMLELEMENTEVENTS2_ONCLICK   ' // click event
'         ' // Get a reference to the element that has fired the event
'         DIM pElement AS IHTMLElement PTR
'         IF pEvtObj THEN pEvtObj->lpvtbl->get_srcElement(pEvtObj, @pElement)
'         IF pElement = NULL THEN EXIT FUNCTION
'         DIM bstrHtml AS AFX_BSTR   ' // Outer html
'         pElement->lpvtbl->get_outerHtml(pElement, @bstrHtml)
''         DIM bstrId AS AFX_BSTR   ' // identifier
''         pElement->lpvtbl->get_id(pElement, @bstrId)
'         pElement->lpvtbl->Release(pElement)
'         AfxMsg *bstrHtml
'         SysFreeString bstrHtml
'         RETURN TRUE
'
'   END SELECT

   RETURN FALSE

END FUNCTION
' ========================================================================================
```

#### Example

```
' ########################################################################################
' Microsoft Windows
' Contents: HTML GUI
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#define _CAXH_DEBUG_ 1
'#define _CWBX_DEBUG_ 1
#INCLUDE ONCE "Afx/CAxHost/CWebCtx.inc"
USING Afx

CONST IDC_WEBBROWSER = 1001
CONST IDC_SATUSBAR = 1002

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG

   END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT
DECLARE SUB WebBrowser_StatusTextChangeProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pwszText AS WSTRING PTR)
DECLARE SUB WebBrowser_DocumentCompleteProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR, BYVAL vUrl AS VARIANT PTR)
DECLARE SUB WebBrowser_BeforeNavigate2Proc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR, _
    BYVAL vUrl AS VARIANT PTR, BYVAL Flags AS VARIANT PTR, BYVAL TargetFrameName AS VARIANT PTR, _
    BYVAL PostData AS VARIANT PTR, BYVAL Headers AS VARIANT PTR, BYVAL pbCancel AS VARIANT_BOOL PTR)
DECLARE FUNCTION WebBrowser_HtmlDocumentEventsProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL dispId AS LONG, BYVAL pEvtObj AS IHTMLEventObj PTR) AS BOOLEAN
DECLARE FUNCTION DocHostUI_ShowContextMenuProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL dwID AS DWORD, BYVAL ppt AS POINT PTR, BYVAL pcmdtReserved AS IUnknown PTR, BYVAL pdispReserved AS IDispatch PTR) AS HRESULT
DECLARE FUNCTION DocHostUI_GetHostInfo (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pInfo AS DOCHOSTUIINFO PTR) AS HRESULT
DECLARE FUNCTION DocHostUI_TranslateAccelerator (BYVAL pWebCtx AS CWebCtx PTR, BYVAL lpMsg AS LPMSG, BYVAL pguidCmdGroup AS const GUID PTR, BYVAL nCmdID AS DWORD) AS HRESULT

' ========================================================================================
' Build the script
' ========================================================================================
FUNCTION BuildScript () AS STRING

   DIM s AS STRING
   s = "<!DOCTYPE html>"
   s += "<head>"
   s += "<meta http-equiv=""MSThemeCompatible"" content=""Yes"">"
   s += "  <title>WebGui</title>"
   s += ""
   s += "<style type=""text/css"">"
   s += "<!--"
   s += ""
   s += "#output"
   s += "{"
   s += "background: #FFFFCC;"
   s += "border: thin solid black;"
   s += "text-align: center;"
   s += "width: 300px;"
   s += "}"
   s += "-->"
   s += "</style>"
   s += ""
   s += "</head>"
   s += "<body scroll='auto'>"
   s += "<input type =""Button"" id=""Button_1"" name=""Button_1"" value=""Button 1""><br />"
   s += "<input type =""Button"" id=""Button_2"" name=""Button_2"" value=""Button 2""><br />"
   s += "<input type =""Button"" id=""Button_3"" name=""Button_3"" value=""Button 3""><br />"
   s += "<input type =""Button"" id=""Button_4"" name=""Button_4"" value=""Button 4""><br />"
   s += "<br />"
   s += "<div id=""output"">"
   s += "Click a button"
   s += "</div>"
   s += "<br />"
   s += "<br />"
   s += "<input type=""Text"" id=""Input_Text"" name=""Input_Text"" value="""" size=40><br />"
   s += "<br />"
   s += "<input type =""Button"" id=""Button_GetText"" name=""Button_GetTex"" value=""Get text""><br />"
   s += "</body>"
   s += "</html>"

   RETURN s

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
   DIM hwndMain AS HWND = pWindow.Create(NULL, "Web GUI", @WndProc)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(450, 300)
   ' // Centers the window
   pWindow.Center

   ' // Add a status bar
   DIM hStatusbar AS HWND = pWindow.AddControl("Statusbar", , IDC_SATUSBAR)

   ' // Add a WebBrowser control
   DIM pwb AS CWebCtx = CWebCtx(@pWindow, IDC_WEBBROWSER, 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   ' // Set web browser event callback procedures
   pwb.SetEventProc("StatusTextChange", @WebBrowser_StatusTextChangeProc)
   pwb.SetEventProc("DocumentComplete", @WebBrowser_DocumentCompleteProc)
   pwb.SetEventProc("BeforeNavigate2", @WebBrowser_BeforeNavigate2Proc)
   pwb.SetEventProc("HtmlDocumentEvents", @WebBrowser_HtmlDocumentEventsProc)
   ' // Set DocHostUI event callback procedures
   pwb.SetUIEventProc("ShowContextMenu", @DocHostUI_ShowContextMenuProc)
   pwb.SetUIEventProc("GetHostInfo", @DocHostUI_GetHostInfo)
   pwb.SetUIEventProc("TranslateAccelerator", @DocHostUI_TranslateAccelerator)

   ' // Build the script
   DIM s AS STRING = BuildScript
   ' // Save the script as a temporary file
   DIM wszPath AS WSTRING * MAX_PATH = AfxSaveTempFile(s, "html")
   ' // Navigate to the path
   pwb.Navigate(wszPath)
   ' // Wait for page load with a timeout of 5 seconds
   DIM lReadyState AS READYSTATE = pwb.WaitForPageLoad(5)
   ' // Kill the temporary file
   KILL wszPath
   ' // Set focus
   pwb.SetFocus

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
            ' // Resize the status bar
            DIM hStatusBar AS HWND = GetDlgItem(hwnd, IDC_SATUSBAR)
            SendMessage hStatusBar, uMsg, wParam, lParam
            ' // Calculate the size of the status bar
            DIM StatusBarHeight AS DWORD, rc AS RECT
            GetWindowRect hStatusBar, @rc
            StatusBarHeight = rc.Bottom - rc.Top
            ' // Retrieve a pointer to the CWindow class
            DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
            ' // Move the position of the control
            IF pWindow THEN pWindow->MoveWindow GetDlgItem(hwnd, IDC_WEBBROWSER), _
               0, 0, pWindow->ClientWidth, pWindow->ClientHeight - StatusBarHeight / pWindow->ryRatio, CTRUE
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

' ========================================================================================
' Process the WebBrowser StatusTextChange event.
' ========================================================================================
SUB WebBrowser_StatusTextChangeProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pwszText AS WSTRING PTR)
   IF pwszText THEN StatusBar_SetText(GetDlgItem(GetParent(pWebCtx->hWindow), IDC_SATUSBAR), 0, pwszText)
END SUB
' ========================================================================================

' ========================================================================================
' Process the WebBrowser DocumentComplete event.
' ========================================================================================
SUB WebBrowser_DocumentCompleteProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR, BYVAL vUrl AS VARIANT PTR)
   ' // The vUrl parameter is a VT_BYREF OR VT_BSTR variant
   ' // It can be a VT_BSTR variant or a VT_ARRAY OR VT_UI1 with a pidl
   DIM varUrl AS VARIANT
   VariantCopyInd(@varUrl, vUrl)
   StatusBar_SetText(GetDlgItem(GetParent(pWebCtx->hWindow), IDC_SATUSBAR), 0, "Document complete: " & AfxVarToStr(@varUrl))
   VariantClear(@varUrl)
END SUB
' ========================================================================================

' ========================================================================================
' Process the IDocHostUIHandler ShowContextMenu event.
' ========================================================================================
FUNCTION DocHostUI_ShowContextMenuProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL dwID AS DWORD, BYVAL ppt AS POINT PTR, BYVAL pcmdtReserved AS IUnknown PTR, BYVAL pdispReserved AS IDispatch PTR) AS HRESULT
   ' // This event notifies that the user has clicked the right mouse button to show the
   ' // context menu. We can anulate it returning %S_OK and show our context menu.
   ' // Do not allow to show the context menu
'   AfxMsg "Sorry! Context menu disabled"
'   RETURN S_OK
   ' // Host did not display its UI. MSHTML will display its UI.
   RETURN S_FALSE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Process the IDocHostUIHandler GetHostInfo event.
' ========================================================================================
PRIVATE FUNCTION DocHostUI_GetHostInfo (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pInfo AS DOCHOSTUIINFO PTR) AS HRESULT
   IF pInfo THEN
      pInfo->cbSize = SIZEOF(DOCHOSTUIINFO)
      pInfo->dwFlags = DOCHOSTUIFLAG_NO3DBORDER OR DOCHOSTUIFLAG_THEME OR DOCHOSTUIFLAG_DPI_AWARE
      pInfo->dwDoubleClick = DOCHOSTUIDBLCLK_DEFAULT
      pInfo->pchHostCss = NULL
      pInfo->pchHostNS = NULL
   END IF
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Process the IDocHostUIHandler TranslateAccelerator event.
' ========================================================================================
PRIVATE FUNCTION DocHostUI_TranslateAccelerator (BYVAL pWebCtx AS CWebCtx PTR, BYVAL lpMsg AS LPMSG, BYVAL pguidCmdGroup AS const GUID PTR, BYVAL nCmdID AS DWORD) AS HRESULT

   IF lpMsg->message = WM_KEYDOWN THEN
      pWebCtx->SetElementInnerHtmlById "output", "ID: " & pWebCtx->GetActiveElementId & " KeyDown - Key: " & WSTR(lpMsg->wParam)
   END IF

   ' // When you use accelerator keys such as TAB, you may need to override the
   ' // default host behavior. The example shows how to do this.
    IF lpMsg->message = WM_KEYDOWN AND lpMsg->wParam = VK_TAB THEN
       RETURN S_FALSE   ' S_OK to disable tab navigation
    END IF
   ' // Return S_FALSE if you don't process the message
   RETURN S_FALSE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Fires before navigation occurs in the given object (on either a window or frameset element).
' ========================================================================================
SUB WebBrowser_BeforeNavigate2Proc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR, _
    BYVAL vUrl AS VARIANT PTR, BYVAL Flags AS VARIANT PTR, BYVAL TargetFrameName AS VARIANT PTR, _
    BYVAL PostData AS VARIANT PTR, BYVAL Headers AS VARIANT PTR, BYVAL pbCancel AS VARIANT_BOOL PTR)

   OutputDebugStringW AfxVarToStr(vUrl)

'   ' // Sample code to redirect navigation to another url
'   IF AfxVarToStr(vUrl) = "http://com.it-berater.org/" THEN
'      ' // Stop loading the page
'      pWebCtx->BrowserPtr->Stop
'      ' // Cancel the navigation operation
'      *pbCancel = VARIANT_TRUE
'      ' // Navigate to another new url
'      DIM cvNewUrl AS CVAR = "http://www.planetsquires.com/protect/forum/index.php"
'      pWebCtx->BrowserPtr->Navigate2(cvNewUrl)
'   END IF

END SUB
' ========================================================================================

' ========================================================================================
' For cancelable document events return TRUE to indicate that Internet Explorer should
' perform its own event processing or FALSE to cancel the event.
' ========================================================================================
PRIVATE FUNCTION WebBrowser_HtmlDocumentEventsProc (BYVAL pWebCtx AS CWebCtx PTR, _
   BYVAL dispid AS LONG, BYVAL pEvtObj AS IHTMLEventObj PTR) AS BOOLEAN

   SELECT CASE dispid

      CASE DISPID_HTMLELEMENTEVENTS2_ONCLICK   ' // click event

         IF pWebCtx->BrowserPtr = NULL THEN EXIT FUNCTION   ' // Wait for page load with a timeout of 5 seconds
         ' // Get a reference to the element that has fired the event
         DIM pElement AS IHTMLElement PTR
         IF pEvtObj THEN pEvtObj->lpvtbl->get_srcElement(pEvtObj, @pElement)
         IF pElement = NULL THEN EXIT FUNCTION
'         DIM bstrHtml AS AFX_BSTR   ' // Outer html
'         pElement->lpvtbl->get_outerHtml(pElement, @bstrHtml)
'         DIM bstrId AS AFX_BSTR   ' // identifier
'         pElement->lpvtbl->get_id(pElement, @bstrId)
'         pElement->lpvtbl->Release(pElement)
'         AfxMsg *bstrHtml
'         SysFreeString bstrHtml
'         RETURN TRUE

         DIM bstrId AS AFX_BSTR   ' // identifier
         pElement->lpvtbl->get_id(pElement, @bstrId)
         DIM cwsId AS CWSTR = *bstrId
         SysFreeString bstrId

         SELECT CASE **cwsId
            CASE "Button_1", "Button_2", "Button_3", "Button_4"
               pWebCtx->SetElementInnerHtmlById("output", "You have clicked " & cwsId)
            CASE "Button_GetText"
               DIM vValue AS VARIANT = pWebCtx->GetElementValueById("Input_Text")
               AfxMsg AfxVarToStr(@vValue, TRUE)
         END SELECT

   END SELECT

   RETURN FALSE

END FUNCTION
' ========================================================================================
```

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

# <a name="Constructor"></a>Constructor

Creates an instance of the **CWebCtx** class.

```
CONSTRUCTOR CWebCtx (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, BYVAL x AS LONG = 0, _
   BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, BYVAL dwStyle AS DWORD = 0, _
   BYVAL dwExStyle AS DWORD = 0, BYVAL pAmbientFlags AS CAXHOST_AMBIENTDISP PTR = NULL)
```

| Member     | Description |
| ---------- | ----------- |
| *pWindow* | Pointer to the instance of the **CWindow** class used to create the form. |
| *cID* |The identifier of the control. It must be unique. |
| *x* | The x-coordinate of the upper-left corner of the window relative to the upper-left corner of the parent window's client area. |
| *y* | The initial y-coordinate of the upper-left corner of the window relative to the upper-left corner of the parent window's client area. |
| *nWidth* | The width of the window. |
| *nHeight* | The height of the window. |
| *dwStyle* | The style of the window being created. |
| *dwExStyle* | The extended style of the window being created. |
| *pAmbientDisp* | Pointer to a **CAXHOST_AMBIENTDISP** structure. |

#### Example

````
Example

#define UNICODE
#INCLUDE ONCE "Afx/CVAR.inc"
#INCLUDE ONCE "Afx/CAxHost/CWebCtx.inc"
USING Afx

CONST IDC_WEBBROWSER = 1001
CONST IDC_SATUSBAR = 1002

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG

   END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT
DECLARE SUB WebBrowser_StatusTextChangeProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pwszText AS WSTRING PTR)
DECLARE SUB WebBrowser_DocumentCompleteProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR, BYVAL vUrl AS VARIANT PTR)
DECLARE SUB WebBrowser_BeforeNavigate2Proc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR, _
    BYVAL vUrl AS VARIANT PTR, BYVAL Flags AS VARIANT PTR, BYVAL TargetFrameName AS VARIANT PTR, _
    BYVAL PostData AS VARIANT PTR, BYVAL Headers AS VARIANT PTR, BYVAL pbCancel AS VARIANT_BOOL PTR)
DECLARE FUNCTION WebBrowser_HtmlDocumentEventsProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL dispId AS LONG, BYVAL pEvtObj AS IHTMLEventObj PTR) AS BOOLEAN
DECLARE FUNCTION DocHostUI_ShowContextMenuProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL dwID AS DWORD, BYVAL ppt AS POINT PTR, BYVAL pcmdtReserved AS IUnknown PTR, BYVAL pdispReserved AS IDispatch PTR) AS HRESULT
DECLARE FUNCTION DocHostUI_GetHostInfo (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pInfo AS DOCHOSTUIINFO PTR) AS HRESULT
DECLARE FUNCTION DocHostUI_TranslateAccelerator (BYVAL pWebCtx AS CWebCtx PTR, BYVAL lpMsg AS LPMSG, BYVAL pguidCmdGroup AS const GUID PTR, BYVAL nCmdID AS DWORD) AS HRESULT

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
   DIM hwndMain AS HWND = pWindow.Create(NULL, "Embedded WebBrowser control", @WndProc)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(750, 450)
   ' // Centers the window
   pWindow.Center

   ' // Add a status bar
   DIM hStatusbar AS HWND = pWindow.AddControl("Statusbar", , IDC_SATUSBAR)

   ' // Add a WebBrowser control
   DIM pwb AS CWebCtx = CWebCtx(@pWindow, IDC_WEBBROWSER, 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   ' // Set web browser event callback procedures
   pwb.SetEventProc("StatusTextChange", @WebBrowser_StatusTextChangeProc)
   pwb.SetEventProc("DocumentComplete", @WebBrowser_DocumentCompleteProc)
   pwb.SetEventProc("BeforeNavigate2", @WebBrowser_BeforeNavigate2Proc)
   pwb.SetEventProc("HtmlDocumentEvents", @WebBrowser_HtmlDocumentEventsProc)
   ' // Set DocHostUI event callback procedures
   pwb.SetUIEventProc("ShowContextMenu", @DocHostUI_ShowContextMenuProc)
   pwb.SetUIEventProc("GetHostInfo", @DocHostUI_GetHostInfo)
   pwb.SetUIEventProc("TranslateAccelerator", @DocHostUI_TranslateAccelerator)

   ' // Navigate to a URL
   pwb.Navigate "http://com.it-berater.org/"
   ' // Set the focus in the page (the page must be fully loaded)
   pwb.SetFocus

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
            ' // Resize the status bar
            DIM hStatusBar AS HWND = GetDlgItem(hwnd, IDC_SATUSBAR)
            SendMessage hStatusBar, uMsg, wParam, lParam
            ' // Calculate the size of the status bar
            DIM StatusBarHeight AS DWORD, rc AS RECT
            GetWindowRect hStatusBar, @rc
            StatusBarHeight = rc.Bottom - rc.Top
            ' // Retrieve a pointer to the CWindow class
            DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
            ' // Move the position of the control
            IF pWindow THEN pWindow->MoveWindow GetDlgItem(hwnd, IDC_WEBBROWSER), _
               0, 0, pWindow->ClientWidth, pWindow->ClientHeight - StatusBarHeight / pWindow->ryRatio, CTRUE
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

' ========================================================================================
' Process the WebBrowser StatusTextChange event.
' ========================================================================================
SUB WebBrowser_StatusTextChangeProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pwszText AS WSTRING PTR)
   IF pwszText THEN StatusBar_SetText(GetDlgItem(GetParent(pWebCtx->hWindow), IDC_SATUSBAR), 0, pwszText)
END SUB
' ========================================================================================

' ========================================================================================
' Process the WebBrowser DocumentComplete event.
' ========================================================================================
SUB WebBrowser_DocumentCompleteProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR, BYVAL vUrl AS VARIANT PTR)
   ' // The vUrl parameter is a VT_BYREF OR VT_BSTR variant
   ' // It can be a VT_BSTR variant or a VT_ARRAY OR VT_UI1 with a pidl
   DIM varUrl AS VARIANT
   VariantCopyInd(@varUrl, vUrl)
   StatusBar_SetText(GetDlgItem(GetParent(pWebCtx->hWindow), IDC_SATUSBAR), 0, "Document complete: " & AfxVarToStr(@varUrl))
   VariantClear(@varUrl)
END SUB
' ========================================================================================

' ========================================================================================
' Process the IDocHostUIHandler ShowContextMenu event.
' ========================================================================================
FUNCTION DocHostUI_ShowContextMenuProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL dwID AS DWORD, BYVAL ppt AS POINT PTR, BYVAL pcmdtReserved AS IUnknown PTR, BYVAL pdispReserved AS IDispatch PTR) AS HRESULT
   ' // This event notifies that the user has clicked the right mouse button to show the
   ' // context menu. We can anulate it returning %S_OK and show our context menu.
   ' // Do not allow to show the context menu
'   AfxMsg "Sorry! Context menu disabled"
'   RETURN S_OK
   ' // Host did not display its UI. MSHTML will display its UI.
   RETURN S_FALSE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Process the IDocHostUIHandler GetHostInfo event.
' ========================================================================================
PRIVATE FUNCTION DocHostUI_GetHostInfo (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pInfo AS DOCHOSTUIINFO PTR) AS HRESULT
   IF pInfo THEN
      pInfo->cbSize = SIZEOF(DOCHOSTUIINFO)
      pInfo->dwFlags = DOCHOSTUIFLAG_NO3DBORDER OR DOCHOSTUIFLAG_THEME OR DOCHOSTUIFLAG_DPI_AWARE
      pInfo->dwDoubleClick = DOCHOSTUIDBLCLK_DEFAULT
      pInfo->pchHostCss = NULL
      pInfo->pchHostNS = NULL
   END IF
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Process the IDocHostUIHandler TranslateAccelerator event.
' ========================================================================================
PRIVATE FUNCTION DocHostUI_TranslateAccelerator (BYVAL pWebCtx AS CWebCtx PTR, BYVAL lpMsg AS LPMSG, BYVAL pguidCmdGroup AS const GUID PTR, BYVAL nCmdID AS DWORD) AS HRESULT

   IF lpMsg->message = WM_KEYDOWN THEN
      pWebCtx->SetElementInnerHtmlById "output", "ID: " & pWebCtx->GetActiveElementId & " KeyDown - Key: " & WSTR(lpMsg->wParam)
   END IF

   ' // When you use accelerator keys such as TAB, you may need to override the
   ' // default host behavior. The example shows how to do this.
    IF lpMsg->message = WM_KEYDOWN AND lpMsg->wParam = VK_TAB THEN
       RETURN S_FALSE   ' S_OK to disable tab navigation
    END IF
   ' // Return S_FALSE if you don't process the message
   RETURN S_FALSE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Fires before navigation occurs in the given object (on either a window or frameset element).
' ========================================================================================
SUB WebBrowser_BeforeNavigate2Proc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR, _
    BYVAL vUrl AS VARIANT PTR, BYVAL Flags AS VARIANT PTR, BYVAL TargetFrameName AS VARIANT PTR, _
    BYVAL PostData AS VARIANT PTR, BYVAL Headers AS VARIANT PTR, BYVAL pbCancel AS VARIANT_BOOL PTR)

   OutputDebugStringW AfxVarToStr(vUrl)

   ' // Sample code to redirect navigation to another url
   IF AfxVarToStr(vUrl) = "http://com.it-berater.org/" THEN
      ' // Stop loading the page
      pWebCtx->BrowserPtr->Stop
      ' // Cancel the navigation operation
      *pbCancel = VARIANT_TRUE
      ' // Navigate to another new url
      DIM cvNewUrl AS CVAR = "http://www.planetsquires.com/protect/forum/index.php"
      pWebCtx->BrowserPtr->Navigate2(cvNewUrl)
   END IF

END SUB
' ========================================================================================

' ========================================================================================
' For cancelable document events return TRUE to indicate that Internet Explorer should
' perform its own event processing or FALSE to cancel the event.
' ========================================================================================
PRIVATE FUNCTION WebBrowser_HtmlDocumentEventsProc (BYVAL pWebCtx AS CWebCtx PTR, _
   BYVAL dispid AS LONG, BYVAL pEvtObj AS IHTMLEventObj PTR) AS BOOLEAN

'   SELECT CASE dispid
'
'      CASE DISPID_HTMLELEMENTEVENTS2_ONCLICK   ' // click event
'         ' // Get a reference to the element that has fired the event
'         DIM pElement AS IHTMLElement PTR
'         IF pEvtObj THEN pEvtObj->lpvtbl->get_srcElement(pEvtObj, @pElement)
'         IF pElement = NULL THEN EXIT FUNCTION
'         DIM bstrHtml AS AFX_BSTR   ' // Outer html
'         pElement->lpvtbl->get_outerHtml(pElement, @bstrHtml)
''         DIM bstrId AS AFX_BSTR   ' // identifier
''         pElement->lpvtbl->get_id(pElement, @bstrId)
'         pElement->lpvtbl->Release(pElement)
'         AfxMsg *bstrHtml
'         SysFreeString bstrHtml
'         RETURN TRUE
'
'   END SELECT

   RETURN FALSE

END FUNCTION
' ========================================================================================
````

#### Example

```
' ########################################################################################
' Microsoft Windows
' Contents: WebBrowser - Events
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "Afx/CVAR.inc"
#INCLUDE ONCE "Afx/CAxHost/CWebCtx.inc"
USING Afx

CONST IDC_WEBBROWSER = 1001
CONST IDC_SATUSBAR = 1002

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG

   END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT
DECLARE SUB WebBrowser_StatusTextChangeProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pwszText AS WSTRING PTR)
DECLARE SUB WebBrowser_DocumentCompleteProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR, BYVAL vUrl AS VARIANT PTR)
DECLARE SUB WebBrowser_BeforeNavigate2Proc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR, _
    BYVAL vUrl AS VARIANT PTR, BYVAL Flags AS VARIANT PTR, BYVAL TargetFrameName AS VARIANT PTR, _
    BYVAL PostData AS VARIANT PTR, BYVAL Headers AS VARIANT PTR, BYVAL pbCancel AS VARIANT_BOOL PTR)
DECLARE FUNCTION WebBrowser_HtmlDocumentEventsProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL dispId AS LONG, BYVAL pEvtObj AS IHTMLEventObj PTR) AS BOOLEAN
DECLARE FUNCTION DocHostUI_ShowContextMenuProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL dwID AS DWORD, BYVAL ppt AS POINT PTR, BYVAL pcmdtReserved AS IUnknown PTR, BYVAL pdispReserved AS IDispatch PTR) AS HRESULT
DECLARE FUNCTION DocHostUI_GetHostInfo (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pInfo AS DOCHOSTUIINFO PTR) AS HRESULT
DECLARE FUNCTION DocHostUI_TranslateAccelerator (BYVAL pWebCtx AS CWebCtx PTR, BYVAL lpMsg AS LPMSG, BYVAL pguidCmdGroup AS const GUID PTR, BYVAL nCmdID AS DWORD) AS HRESULT

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
   DIM hwndMain AS HWND = pWindow.Create(NULL, "Embedded WebBrowser control", @WndProc)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(750, 450)
   ' // Centers the window
   pWindow.Center

   ' // Add a status bar
   DIM hStatusbar AS HWND = pWindow.AddControl("Statusbar", , IDC_SATUSBAR)

   ' // Add a WebBrowser control
   DIM pwb AS CWebCtx = CWebCtx(@pWindow, IDC_WEBBROWSER, 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   ' // Set web browser event callback procedures
   pwb.SetEventProc("StatusTextChange", @WebBrowser_StatusTextChangeProc)
   pwb.SetEventProc("DocumentComplete", @WebBrowser_DocumentCompleteProc)
   pwb.SetEventProc("BeforeNavigate2", @WebBrowser_BeforeNavigate2Proc)
   pwb.SetEventProc("HtmlDocumentEvents", @WebBrowser_HtmlDocumentEventsProc)
   ' // Set DocHostUI event callback procedures
   pwb.SetUIEventProc("ShowContextMenu", @DocHostUI_ShowContextMenuProc)
   pwb.SetUIEventProc("GetHostInfo", @DocHostUI_GetHostInfo)
   pwb.SetUIEventProc("TranslateAccelerator", @DocHostUI_TranslateAccelerator)

   ' // Navigate to a URL
   pwb.Navigate "http://com.it-berater.org/"
   ' // Set the focus in the page (the page must be fully loaded)
   pwb.SetFocus

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
            ' // Resize the status bar
            DIM hStatusBar AS HWND = GetDlgItem(hwnd, IDC_SATUSBAR)
            SendMessage hStatusBar, uMsg, wParam, lParam
            ' // Calculate the size of the status bar
            DIM StatusBarHeight AS DWORD, rc AS RECT
            GetWindowRect hStatusBar, @rc
            StatusBarHeight = rc.Bottom - rc.Top
            ' // Retrieve a pointer to the CWindow class
            DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
            ' // Move the position of the control
            IF pWindow THEN pWindow->MoveWindow GetDlgItem(hwnd, IDC_WEBBROWSER), _
               0, 0, pWindow->ClientWidth, pWindow->ClientHeight - StatusBarHeight / pWindow->ryRatio, CTRUE
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

' ========================================================================================
' Process the WebBrowser StatusTextChange event.
' ========================================================================================
SUB WebBrowser_StatusTextChangeProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pwszText AS WSTRING PTR)
   IF pwszText THEN StatusBar_SetText(GetDlgItem(GetParent(pWebCtx->hWindow), IDC_SATUSBAR), 0, pwszText)
END SUB
' ========================================================================================

' ========================================================================================
' Process the WebBrowser DocumentComplete event.
' ========================================================================================
SUB WebBrowser_DocumentCompleteProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR, BYVAL vUrl AS VARIANT PTR)
   ' // The vUrl parameter is a VT_BYREF OR VT_BSTR variant
   ' // It can be a VT_BSTR variant or a VT_ARRAY OR VT_UI1 with a pidl
   DIM varUrl AS VARIANT
   VariantCopyInd(@varUrl, vUrl)
   StatusBar_SetText(GetDlgItem(GetParent(pWebCtx->hWindow), IDC_SATUSBAR), 0, "Document complete: " & AfxVarToStr(@varUrl))
   VariantClear(@varUrl)
END SUB
' ========================================================================================

' ========================================================================================
' Process the IDocHostUIHandler ShowContextMenu event.
' ========================================================================================
FUNCTION DocHostUI_ShowContextMenuProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL dwID AS DWORD, BYVAL ppt AS POINT PTR, BYVAL pcmdtReserved AS IUnknown PTR, BYVAL pdispReserved AS IDispatch PTR) AS HRESULT
   ' // This event notifies that the user has clicked the right mouse button to show the
   ' // context menu. We can anulate it returning %S_OK and show our context menu.
   ' // Do not allow to show the context menu
'   AfxMsg "Sorry! Context menu disabled"
'   RETURN S_OK
   ' // Host did not display its UI. MSHTML will display its UI.
   RETURN S_FALSE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Process the IDocHostUIHandler GetHostInfo event.
' ========================================================================================
PRIVATE FUNCTION DocHostUI_GetHostInfo (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pInfo AS DOCHOSTUIINFO PTR) AS HRESULT
   IF pInfo THEN
      pInfo->cbSize = SIZEOF(DOCHOSTUIINFO)
      pInfo->dwFlags = DOCHOSTUIFLAG_NO3DBORDER OR DOCHOSTUIFLAG_THEME OR DOCHOSTUIFLAG_DPI_AWARE
      pInfo->dwDoubleClick = DOCHOSTUIDBLCLK_DEFAULT
      pInfo->pchHostCss = NULL
      pInfo->pchHostNS = NULL
   END IF
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Process the IDocHostUIHandler TranslateAccelerator event.
' ========================================================================================
PRIVATE FUNCTION DocHostUI_TranslateAccelerator (BYVAL pWebCtx AS CWebCtx PTR, BYVAL lpMsg AS LPMSG, BYVAL pguidCmdGroup AS const GUID PTR, BYVAL nCmdID AS DWORD) AS HRESULT

   IF lpMsg->message = WM_KEYDOWN THEN
      pWebCtx->SetElementInnerHtmlById "output", "ID: " & pWebCtx->GetActiveElementId & " KeyDown - Key: " & WSTR(lpMsg->wParam)
   END IF

   ' // When you use accelerator keys such as TAB, you may need to override the
   ' // default host behavior. The example shows how to do this.
    IF lpMsg->message = WM_KEYDOWN AND lpMsg->wParam = VK_TAB THEN
       RETURN S_FALSE   ' S_OK to disable tab navigation
    END IF
   ' // Return S_FALSE if you don't process the message
   RETURN S_FALSE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Fires before navigation occurs in the given object (on either a window or frameset element).
' ========================================================================================
SUB WebBrowser_BeforeNavigate2Proc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR, _
    BYVAL vUrl AS VARIANT PTR, BYVAL Flags AS VARIANT PTR, BYVAL TargetFrameName AS VARIANT PTR, _
    BYVAL PostData AS VARIANT PTR, BYVAL Headers AS VARIANT PTR, BYVAL pbCancel AS VARIANT_BOOL PTR)

   OutputDebugStringW AfxVarToStr(vUrl)

   ' // Sample code to redirect navigation to another url
   IF AfxVarToStr(vUrl) = "http://com.it-berater.org/" THEN
      ' // Stop loading the page
      pWebCtx->BrowserPtr->Stop
      ' // Cancel the navigation operation
      *pbCancel = VARIANT_TRUE
      ' // Navigate to another new url
      DIM cvNewUrl AS CVAR = "http://www.planetsquires.com/protect/forum/index.php"
      pWebCtx->BrowserPtr->Navigate2(cvNewUrl)
   END IF

END SUB
' ========================================================================================

' ========================================================================================
' For cancelable document events return TRUE to indicate that Internet Explorer should
' perform its own event processing or FALSE to cancel the event.
' ========================================================================================
PRIVATE FUNCTION WebBrowser_HtmlDocumentEventsProc (BYVAL pWebCtx AS CWebCtx PTR, _
   BYVAL dispid AS LONG, BYVAL pEvtObj AS IHTMLEventObj PTR) AS BOOLEAN

'   SELECT CASE dispid
'
'      CASE DISPID_HTMLELEMENTEVENTS2_ONCLICK   ' // click event
'         ' // Get a reference to the element that has fired the event
'         DIM pElement AS IHTMLElement PTR
'         IF pEvtObj THEN pEvtObj->lpvtbl->get_srcElement(pEvtObj, @pElement)
'         IF pElement = NULL THEN EXIT FUNCTION
'         DIM bstrHtml AS AFX_BSTR   ' // Outer html
'         pElement->lpvtbl->get_outerHtml(pElement, @bstrHtml)
''         DIM bstrId AS AFX_BSTR   ' // identifier
''         pElement->lpvtbl->get_id(pElement, @bstrId)
'         pElement->lpvtbl->Release(pElement)
'         AfxMsg *bstrHtml
'         SysFreeString bstrHtml
'         RETURN TRUE
'
'   END SELECT

   RETURN FALSE

END FUNCTION
' ========================================================================================
```

# <a name="Advise"></a>Advise

Establishes a connection between a connection point object and the client's sink.

```
FUNCTION Advise () AS HRESULT
```

#### Return value

S_OK (0) or an error code.

#### Remarks

You don't have to call this method. The **CWebCtx** constructor calls it.

# <a name="AxHostPtr"></a>AxHostPtr

Returns a pointer to the **CAxHost** class.

```
FUNCTION AxHostPtr () AS CAxHost PTR
```

#### Return value

A pointer to the **CAxHost** class or NULL.

# <a name="BrowserPtr"></a>BrowserPtr

Returns a direct pointer to the **IWebBrowser2** interface of the hosted **WebBrowser** control.

```
FUNCTION BrowserPtr () AS Afx_IWebBrowser2 PTR
```

#### Return value

A pointer to the **IWebBrowser2** interface of the hosted **WebBrowser** control or NULL.

#### Remarks

Since it is a direct pointer, you don't have to release it calling the **Release** method.

# <a name="Busy"></a>Busy

Returns a boolean value indicating whether the object is engaged in a navigation or downloading operation.

```
PROPERTY Busy () AS BOOLEAN
```

#### Return value

A pointer to the **IWebBrowser2** interface of the hosted **WebBrowser** control or NULL.

#### Remarks

It returns TRUE if the control is busy; FALSE otherwise.

# <a name="CtrlID"></a>CtrlID

Returns the identifier of the container's window.

```
FUNCTION CtrlID () AS LONG
```

# <a name="Document"></a>Document

Retrieves the automation object of the active document, if any.

```
PROPERTY Document () AS IHtmlDocument2 PTR
```

### Return value

Returns one of the following values:

| Error code | Description |
| ---------- | ----------- |
| S_OK | The operation was successful. |
| E_FAIL | The operation failed. |
| E_INVALIDARG | One or more parameters are invalid. |
| E_NOINTERFACE | The interface is not supported. |

#### Remarks

When the active document is an HTML page, this property provides access to the contents of the HTML Document Object Model (DOM). Specifically, it returns an **IDispatch** interface pointer to the **HTMLDocument** component object class (coclass). The **HTMLDocument** coclass is functionally equivalent to the Dynamic HTML (DHTML) **document** object used in HTML script. It supports all the properties and methods necessary to access the entire contents of the active HTML document.

FreeBASIC programs can retrieve the Component Object Model (COM) interfaces **IHTMLDocument**, **IHTMLDocument2**, and **IHTMLDocument3** by calling **QueryInterface** on the **IDispatch** received from this property.

When other document types are active, such as a Microsoft Word document, this property returns the default **IDispatch** dispatch interface (dispinterface) pointer for the hosted document object. For Word documents, this would be functionally equivalent to the **Document** object in the Word object model.

# <a name="ExecWB"></a>ExecWB

Executes a command on an OLE object and returns the status of the command execution using the **IOleCommandTarget** interface.

```
FUNCTION ExecWB (BYVAL cmdID AS OLECMDID, BYVAL cmdexecopt AS OLECMDEXECOPT, _
   BYVAL pvaIn AS VARIANT PTR = NULL, BYVAL pvaOut AS VARIANT PTR = NULL) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cmdID* | **OLECMDID** value that specifies the command to execute. |
| *cmdexecopt* | **OLECMDEXECOPT** value that specifies the command options. |
| *pvaIn* | inter to a **VARIANT** that contains command input arguments. The type of this **VARIANT** depends on the type of the command identifier. This argument can be NULL.  |
| *pvaOut* | In, Out. Pointer to a **VARIANT** that receives and specifies command output. The type of this **VARIANT** depends on the type of the command identifier. This argument can be NULL.  |

#### Return value

Returns S_OK (0) if successful, or an error value otherwise. 

# <a name="Find"></a>Find

Activates the find dialog.

```
FUNCTION Find () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

# <a name="GetActiveElementId"></a>GetActiveElementId

Retrieves the ID of the active element (the object that has the focus when the parent document has focus).

```
FUNCTION GetActiveElementId () AS CWSTR
```

#### Return value

The ID of the active element, if any. If the element has not an identifier, it returns an empty string.

# <a name="GetBodyInnerHtml"></a>GetBodyInnerHtml

Returns a string that represents the text and html elements between the start and end body tags.

```
FUNCTION GetBodyInnerHtml () AS CWSTR
```

#### Return value

A string containing the html text.

# <a name="GetBodyInnerText"></a>GetBodyInnerText

Returns a string that represents the text between the start and end body tags without any associated HTML.

```
FUNCTION GetBodyInnerText () AS CWSTR
```

#### Return value

A string containing the text.

# <a name="GoBack"></a>GoBack

Navigates backward one item in the history list.

```
FUNCTION GoBack () AS HRESULT
```

#### Return value

Returns S_OK (0) if successful, or an error value otherwise.

#### Remarks

During a browsing session, the WebBrowser control and Microsoft Internet Explorer maintain a history list of all Web sites visited during a session (unless you specify the NAVNOHISTORY flag when using the Navigate method). Use the **CommandStateChange** event to check the enabled state of backward navigation. If the event's CSC_NAVIGATEBACK command is disabled, the beginning of the history list has been reached and the **GoBack** method should not be used.

# <a name="GetElementInnerHtmlById"></a>GetElementInnerHtmlById

Retrieves the HTML between the start and end tags of the object.

```
FUNCTION GetElementInnerHtmlById (BYREF cwsId AS CWSTR) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *cwsId* | The identifier. |

#### Return value

A string containing the HTML text.

#### Remarks

This method performs a case insensitive property search. If two or more attributes have the same name (differing only in uppercase and lowercase letters) this function retrieves values only for the last attribute created with this name, and ignores all other attributes with the same name.

# <a name="GetElementValueById"></a>GetElementValueById

Retrieves the value attribute of the specified attribute.

```
FUNCTION GetElementValueById (BYREF cwsId AS CWSTR) AS VARIANT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cwsId* | The identifier. |

#### Return value

A variant containing the value as defined by the attribute.

#### Remarks

This method performs a case insensitive property search. If two or more attributes have the same name (differing only in uppercase and lowercase letters) this function retrieves values only for the last attribute created with this name, and ignores all other attributes with the same name.

# <a name="GoForward"></a>GoForward

Navigates forward one item in the history list.

```
FUNCTION GoForward () AS HRESULT
```

#### Return value

Returns S_OK (0) if successful, or an error value otherwise.

#### Remarks

During a browsing session, the WebBrowser control and Microsoft Internet Explorer application maintain a history list of all Web sites visited during a session (unless you specify the NAVNOHISTORY flag when using the Navigate method). Use the **CommandStateChange** event to check the enabled state of forward navigation. If the event's CSC_NAVIGATEFORWARD command is disabled, the end of the history list has been reached and the **GoForward** method should not be used.

# <a name="GoHome"></a>GoHome

Navigates to the current home or start page.

```
FUNCTION GoHome () AS HRESULT
```

#### Return value

Returns S_OK (0) if successful, or an error value otherwise.

#### Remarks

The user can indicate the URL to use for the home or start page either from Internet Options in Microsoft Internet Explorer or from Internet Properties in Control Panel.

# <a name="GoSearch"></a>GoSearch

Navigates to the current search page.

```
FUNCTION GoSearch () AS HRESULT
```

#### Return value

Returns S_OK (0) if successful, or an error value otherwise.

#### Remarks

The user can indicate the URL to use for the search page either from Internet Options in Microsoft Internet Explorer or from Internet Properties in Control Panel.

# <a name="hWindow"></a>hWindow

Returns the handle of the OLE Container hosting window.

```
FUNCTION hWindow () AS HWND
```

# <a name="InternetOptions"></a>InternetOptions

Activates the Internet options dialog.

```
FUNCTION InternetOptions () AS HRESULT
```

#### Return value

Returns S_OK (0) if successful, or an error value otherwise.

# <a name="LocationName"></a>LocationName

Retrieves the name of the resource that Microsoft Internet Explorer is currently displaying.

```
PROPERTY LocationName () AS CWSTR
```

#### Return value

The name of the resource that Microsoft Internet Explorer is currently displaying.

#### Remarks

If the resource is an HTML page on the World Wide Web, the name is the title of that page. If the resource is a folder or file on the network or local computer, the name is the full path of the folder or file in Universal Naming Convention (UNC) format.

# <a name="LocationURL"></a>LocationURL

Retrieves the URL of the resource that Microsoft Internet Explorer is currently displaying.

```
PROPERTY LocationURL () AS CWSTR
```

#### Return value

The URL of the resource that Microsoft Internet Explorer is currently displaying.

#### Remarks

If the resource is a folder or file on the network or local computer, the name is the full path of the folder or file in the Universal Naming Convention (UNC) format.

# <a name="Navigate"></a>Navigate

Returns the handle of the OLE Container hosting window.

```
FUNCTION Navigate (BYVAL pwszUrl AS WSTRING PTR, BYVAL Flags AS VARIANT PTR = NULL, _
   BYVAL TargetFrameName AS VARIANT PTR = NULL, BYVAL PostData AS VARIANT PTR = NULL, _
   BYVAL Headers AS VARIANT PTR = NULL) AS HRESULT
```
```
FUNCTION Navigate (BYVAL vUrl AS VARIANT PTR, BYVAL Flags AS VARIANT PTR = NULL, _
   BYVAL TargetFrameName AS VARIANT PTR = NULL, BYVAL PostData AS VARIANT PTR = NULL, _
   BYVAL Headers AS VARIANT PTR = NULL) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszUrl* | A variable or expression that evaluates to the URL of the resource to display or the full path to the file location. |
| *vUrl* | A variable or expression that evaluates to the URL of the resource to display, the full path to the file location, or a PIDL that represents a folder in the Shell namespace. |
| *Flags* | Optional. A constant or value that specifies a combination of the values defined by the **BrowserNavConstants** enumeration. |
| *TargetFrameName* | Optional. Case-sensitive string expression that evaluates to the name of the frame in which to display the resource. The possible values for this parameter are:<br>**\_blank**: Load the link into a new unnamed window.<br>**\_parent**: Load the link into the immediate parent of the document the link is in.<br>**\_self**: Load the link into the same window the link was clicked in.<br>**\_top**: Load the link into the full body of the current window.<br>**WindowName**: A named HTML frame. If no frame or window exists that matches the specified target name, a new window is opened for the specified link. |
| *PostData* | Optional. Data that is sent to the server as part of a HTTP POST transaction. A POST transaction typically is used to send data collected by an HTML form. If this parameter does not specify any POST data, this method issues an HTTP GET transaction. This parameter is ignored if URL is not an HTTP (or HTTPS) URL. |
| *Headers* | Optional. A String that contains additional HTTP headers to send to the server. These headers are added to the default Internet Explorer headers. For example, headers can specify the action required of the server, the type of data being passed to the server, or a status code. This parameter is ignored if the URL is not an HTTP (or HTTPS)  URL. |

### BrowserNavConstants Enumeration

Contains values used by the **Navigate2** method.

| Flag       | Description |
| ---------- | ----------- |
| **navOpenInNewWindow** | Open the resource or file in a new window. |
| **navNoHistory** | Do not add the resource or file to the history list. The new page replaces the current page in the list. |
| **navNoReadFromCache** | Not implemented. |
| **navNoWriteToCache** | Not implemented. |
| **navAllowAutosearch** | If the navigation fails, the autosearch functionality attempts to navigate common root domains (.com, .edu, and so on). If this also fails, the URL is passed to a search engine. |
| **navBrowserBar** | Causes the current Explorer Bar to navigate to the given item, if possible. |
| **navHyperlink** | Microsoft Internet Explorer 6 for Windows XP Service Pack 2 (SP2) and later. If the navigation fails when a hyperlink is being followed, this constant specifies that the resource should then be bound to the moniker using the BINDF_HYPERLINK flag. |
| **navEnforceRestricted** | Internet Explorer 6 for Windows XP SP2 and later. Force the URL into the restricted zone. |
| **navNewWindowsManaged** | Internet Explorer 6 for Windows XP SP2 and later. Use the default Popup Manager to block pop-up windows. |
| **navUntrustedForDownload** | Internet Explorer 6 for Windows XP SP2 and later. Block files that normally trigger a file download dialog box.|
| **navTrustedForActiveX** | Internet Explorer 6 for Windows XP SP2 and later. Prompt for the installation of Microsoft ActiveX controls. |
| **navOpenInNewTab** | Windows Internet Explorer 7. Open the resource or file in a new tab. Allow the destination window to come to the foreground, if necessary. |
| **navOpenInBackgroundTab** | Internet Explorer 7. Open the resource or file in a new background tab; the currently active window and/or tab remains open on top. |
| **navKeepWordWheelText** | Internet Explorer 7. Maintain state for dynamic navigation based on the filter string entered in the search band text box (wordwheel). Restore the wordwheel text when the navigation completes. |
| **navVirtualTab** | Internet Explorer 8. Open the resource as a replacement for the current or target tab. The existing tab is closed while the new tab takes its place in the tab bar and replaces it in the tab group, if any. Browser history is copied forward to the new tab. On Windows Vista, this flag is implied if the navigation would cross integrity levels and **navOpenInNewTab**, **navOpenInBackgroundTab**, or **navOpenInNewWindow** is not specified. |
| **navBlockRedirectsXDomain** | Internet Explorer 8. Block cross-domain redirect requests. The navigation triggers the RedirectXDomainBlocked event if blocked. |
| **navOpenNewForegroundTab** | Internet Explorer 8 and later. Open the resource in a new tab that becomes the foreground tab. |

#### Return value

Returns one of the following values:

| Error code | Description |
| ---------- | ----------- |
| S_OK | The operation was successful. |
| E_FAIL | The operation failed. |
| E_INVALIDARG | One or more parameters are invalid. |
| E_OUTOFMEMORY | Out of memory. |

# <a name="PageProperties"></a>PageProperties

Activates the properties dialog.

```
FUNCTION PageProperties () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

# <a name="PageSetup"></a>PageSetup

Activates the page setup dialog.

```
FUNCTION PageSetup () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

# <a name="PrintPage"></a>PrintPage

Activates the print dialog.

```
FUNCTION PrintPage () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

# <a name="PrintPreview"></a>PrintPreview

Activates the print preview dialog.

```
FUNCTION PrintPreview () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

# <a name="QueryStatusWB"></a>QueryStatusWB

Queries the OLE object for the status of commands using the **QueryStatus** method of the **IOleCommandTarget** interface.

```
FUNCTION QueryStatusWB (BYVAL cmdID AS OLECMDID) AS OLECMDF
```

| Parameter  | Description |
| ---------- | ----------- |
| *cmdID* | **OLECMDID** value of the command for which the caller needs status information. |

#### Return value

The status of the command.

# <a name="ReadyState"></a>ReadyState

Retrieves the ready state of the object.

```
PROPERTY ReadyState () AS tagREADYSTATE
```

#### Return value

The ready state of the object.

# <a name="Refresh"></a>Refresh

Reloads the file that is currently displayed in the object.

```
FUNCTION Refresh () AS HRESULT
```

#### Return value

Returns S_OK (0) if successful, or E_FAIL otherwise.

#### Remarks

The WebBrowser control and InternetExplorer application store Web pages from recently visited sites in cached memory on the user's hard disk. This saves time when revisiting a site by reloading the page from the local disk rather than downloading it again across the network from the remote HTTP server. You can force the WebBrowser control and WebBrowser application to redownload a page by using the **Refresh** or **Refresh2** methods of the **IWebBrowser2** interface to ensure that you are viewing the most current version of the page. Also, you can disable the cache from being used by specifying the **navNoReadFromCache** and **navNoWriteToCache** flags when calling the **Navigate2** method of the **IWebBrowser2** interface. 

# <a name="Refresh2"></a>Refresh2

Reloads the file that is currently displayed in the object. Unlike **Refresh**, this method contains a parameter that specifies the refresh level.

```
FUNCTION Refresh2 (BYVAL nLevel AS RefreshConstants) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *nLevel* | One of the **RefreshConstants** enumeration values. |

### RefreshConstants enumeration

| Constant   | Description |
| ---------- | ----------- |
| REFRESH_NORMAL (0) | Perform a lightweight refresh that does not include the pragma:nocache header. The pragma:nocache header tells the server not to return a cached copy. This can cause problems with some servers. |
| REFRESH_IFEXPIRED (1) | Only refresh if the page has expired. Do not include the pragma:nocache header. |
| REFRESH_COMPLETELY (3) | Perform a full refresh, including the pragma:nocache header. Using this option is the same as calling the Refresh method. |

#### Return value

Returns S_OK (0) if successful, or E_FAIL otherwise.

#### Remarks

The "pragma:nocache" header tells the server not to return a cached copy. This ensures that the information is as fresh as possible. Browsers typically send this header when the user selects refresh, but the header can cause problems for some servers.

# <a name="RegisterAsBrowser"></a>RegisterAsBrowser

Sets or retrieves a value that indicates whether the object is registered as a top-level browser for target name resolution.

```
PROPERTY RegisterAsBrowser () AS BOOLEAN
PROPERTY RegisterAsBrowser (BYVAL bRegister AS BOOLEAN)
```

| Parameter  | Description |
| ---------- | ----------- |
| *bRegister* | True of False. |

# <a name="RegisterAsDropTarget"></a>RegisterAsDropTarget

Sets or retrieves a value that indicates whether the object is registered as a drop target for navigation.

```
PROPERTY RegisterAsDropTarget () AS BOOLEAN
PROPERTY RegisterAsDropTarget (BYVAL bRegister AS BOOLEAN)
```

| Parameter  | Description |
| ---------- | ----------- |
| *bRegister* | True of False. |

# <a name="SaveAs"></a>SaveAs

Activates the save file dialog.

```
FUNCTION SaveAs () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

# <a name="SetElementFocusById"></a>SetElementFocusById

Sets the HTML between the start and end tags of the object.

```
FUNCTION SetElementFocusById (BYREF cwsId AS CWSTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cwsId* | The identifier. |

#### Return value

S_OK (0) or an HRESULT code.

# <a name="SetElementInnerHtmlById"></a>SetElementInnerHtmlById

Sets the HTML between the start and end tags of the object.

```
FUNCTION SetElementInnerHtmlById (BYREF cwsId AS CWSTR, BYREF cwsHtml AS CWSTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cwsId* | The identifier. |
| *cwsHtml* | The html text to set. |

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

This method performs a case insensitive property search. If two or more attributes have the same name (differing only in uppercase and lowercase letters) this function sets values only for the last attribute created with this name, and ignores all other attributes with the same name.

# <a name="SetElementValueById"></a>SetElementValueById

Sets the value attribute of the specified identifier.

```
FUNCTION SetElementValueById (BYREF cwsId AS CWSTR, BYVAL vValue AS VARIANT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cwsId* | The identifier. |
| *vValue* | Variant that specifies the string, number, or Boolean to assign to the attribute. |

#### Return value

S_OK (0) or an HRESULT code.

#### Remarks

This method performs a case insensitive property search. If two or more attributes have the same name (differing only in uppercase and lowercase letters) this function sets values only for the last attribute created with this name, and ignores all other attributes with the same name.

# <a name="SetEventProc"></a>SetEventProc

Sets pointers to user implemented callback procedures to receive events of the hosted WebBrowser control.

```
FUNCTION SetEventProc (BYVAL pwszEventName AS WSTRING PTR, BYVAL pProc AS ANY PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszEventName* | The name of the event. |
| *pProc* | The address of the callback procedure. |

### Event Names

| Parameter  |
| ---------- |
| StatusTextChange |
| DownloadComplete |
| CommandStateChange |
| DownloadBegin |
| ProgressChange |
| PropertyChange |
| TitleChange |
| PrintTemplateInstantiation |
| PrintTemplateTeardown |
| BeforeNavigate2 |
| NewWindow2 |
| NavigateComplete2 |
| OnVisible |
| OnToolBar |
| OnMenuBar |
| OnStatusBar |
| OnFullScreen |
| DocumentComplete |
| OnTheaterMode |
| WindowSetResizable |
| WindowClosing |
| WindowSetLeft |
| WindowSetTop |
| WindowSetWidth |
| WindowSetHeight |
| ClientToHostWindow |
| SetSecureLockIcon |
| FileDownload |
| NavigateError |
| PrivacyImpactedStateChange |
| NewWindow3 |
| SetPhishingFilterStatus |
| WindowStateChanged |
| HtmlDocumentEvents |

#### Return value

S_OK (0) or an HRESULT code.

### Function callback prototypes

```
SUB BeforeNavigate2Proc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR, _
    BYVAL vUrl AS VARIANT PTR, BYVAL Flags AS VARIANT PTR, BYVAL TargetFrameName AS VARIANT PTR, _
    BYVAL PostData AS VARIANT PTR, BYVAL Headers AS VARIANT PTR, BYVAL pbCancel AS VARIANT_BOOL PTR)
SUB ClientToHostWindowProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL CX AS LONG PTR, BYVAL CY AS LONG PTR)
SUB CommandStateChangeProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL nCommand AS LONG, BYVAL fEnable AS VARIANT_BOOL)
SUB DocumentCompleteProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR, BYVAL vUrl AS VARIANT PTR)
SUB DownloadBeginProc (BYVAL pWebCtx AS CWebCtx PTR)
SUB DownloadCompleteProc (BYVAL pWebCtx AS CWebCtx PTR)
SUB FileDownloadProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pbCancel AS VARIANT_BOOL PTR)
SUB NavigateComplete2Proc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR, BYVAL vUrl AS VARIANT PTR)
SUB NavigateErrorProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR, _
    BYVAL vUrl AS VARIANT PTR, BYVAL vFrame AS VARIANT PTR, BYVAL StatusCode AS VARIANT PTR, BYVAL pbCancel AS VARIANT_BOOL PTR)
SUB NewWindow3Proc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL ppdispVal AS IDispatch PTR PTR, _
    BYVAL pbCancel AS VARIANT_BOOL PTR, BYVAL dwFlags AS ULONG, BYVAL pwszUrlContext AS WSTRING PTR, BYVAL pwszUrl AS WSTRING PTR)
SUB OnVisibleProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL bVisible AS VARIANT_BOOL)
SUB PrintTemplateInstantiationProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR)
SUB PrintTemplateTeardownProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR)
SUB PrivacyImpactedStateChangeProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL bImpacted AS VARIANT_BOOL)
SUB ProgressChangeProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL Progress AS LONG, BYVAL ProgressMax AS LONG)
SUB PropertyChangeProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pwszProperty AS WSTRING PTR)
SUB SetSecureLockIconProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL SecureLockIcon AS LONG)
SUB StatusTextChangeProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pwszText AS WSTRING PTR)
SUB TitleChangeProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pwszText AS WSTRING PTR)
SUB WindowClosingProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL IsChildWindow AS VARIANT_BOOL, BYVAL pbCancel AS VARIANT_BOOL PTR)
SUB WindowSetHeightProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL nHeight AS LONG)
SUB WindowSetLeftProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL nLenft AS LONG)
SUB WindowSetResizableProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL bResizable AS VARIANT_BOOL)
SUB WindowSetTopProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL nTop AS LONG)
SUB WindowSetWidthProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL nWidth AS LONG)
SUB WindowStateChangedProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL dwFlags AS LONG, BYVAL dwValidFlagsMask AS LONG)
FUNCTION HtmlDocumentEventsProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL dispid AS LONG, BYVAL pEvtObj AS IHTMLEventObj) AS BOOLEAN
```
#### Usage example

```
' // Add a WebBrowser control
DIM pwb AS CWebCtx = CWebCtx(@pWindow, IDC_WEBBROWSER, 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
' // Set web browser event callback procedures
pwb.SetEventProc("StatusTextChange", @WebBrowser_StatusTextChangeProc)
pwb.SetEventProc("DocumentComplete", @WebBrowser_DocumentCompleteProc)
pwb.SetEventProc("BeforeNavigate2", @WebBrowser_BeforeNavigate2Proc)
pwb.SetEventProc("HtmlDocumentEvents", @WebBrowser_HtmlDocumentEventsProc)
' // Set DocHostUI event callback procedures
pwb.SetUIEventProc("ShowContextMenu", @DocHostUI_ShowContextMenuProc)
pwb.SetUIEventProc("GetHostInfo", @DocHostUI_GetHostInfo)
pwb.SetUIEventProc("TranslateAccelerator", @DocHostUI_TranslateAccelerator)

' ========================================================================================
' Process the WebBrowser StatusTextChange event.
' ========================================================================================
SUB WebBrowser_StatusTextChangeProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pwszText AS WSTRING PTR)
   IF pwszText THEN StatusBar_SetText(GetDlgItem(GetParent(pWebCtx->hWindow), IDC_SATUSBAR), 0, pwszText)
END SUB
' ========================================================================================

' ========================================================================================
' Process the WebBrowser DocumentComplete event.
' ========================================================================================
SUB WebBrowser_DocumentCompleteProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR, BYVAL vUrl AS VARIANT PTR)
   ' // The vUrl parameter is a VT_BYREF OR VT_BSTR variant
   ' // It can be a VT_BSTR variant or a VT_ARRAY OR VT_UI1 with a pidl
   DIM varUrl AS VARIANT
   VariantCopyInd(@varUrl, vUrl)
   StatusBar_SetText(GetDlgItem(GetParent(pWebCtx->hWindow), IDC_SATUSBAR), 0, "Document complete: " & AfxVarToStr(@varUrl))
   VariantClear(@varUrl)
END SUB
' ========================================================================================

' ========================================================================================
' Process the IDocHostUIHandler ShowContextMenu event.
' ========================================================================================
FUNCTION DocHostUI_ShowContextMenuProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL dwID AS DWORD, BYVAL ppt AS POINT PTR, BYVAL pcmdtReserved AS IUnknown PTR, BYVAL pdispReserved AS IDispatch PTR) AS HRESULT
   ' // This event notifies that the user has clicked the right mouse button to show the
   ' // context menu. We can anulate it returning %S_OK and show our context menu.
   ' // Do not allow to show the context menu
'   AfxMsg "Sorry! Context menu disabled"
'   RETURN S_OK
   ' // Host did not display its UI. MSHTML will display its UI.
   RETURN S_FALSE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Process the IDocHostUIHandler GetHostInfo event.
' ========================================================================================
PRIVATE FUNCTION DocHostUI_GetHostInfo (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pInfo AS DOCHOSTUIINFO PTR) AS HRESULT
   IF pInfo THEN
      pInfo->cbSize = SIZEOF(DOCHOSTUIINFO)
      pInfo->dwFlags = DOCHOSTUIFLAG_NO3DBORDER OR DOCHOSTUIFLAG_THEME OR DOCHOSTUIFLAG_DPI_AWARE
      pInfo->dwDoubleClick = DOCHOSTUIDBLCLK_DEFAULT
      pInfo->pchHostCss = NULL
      pInfo->pchHostNS = NULL
   END IF
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Process the IDocHostUIHandler TranslateAccelerator event.
' ========================================================================================
PRIVATE FUNCTION DocHostUI_TranslateAccelerator (BYVAL pWebCtx AS CWebCtx PTR, BYVAL lpMsg AS LPMSG, BYVAL pguidCmdGroup AS const GUID PTR, BYVAL nCmdID AS DWORD) AS HRESULT

   IF lpMsg->message = WM_KEYDOWN THEN
      pWebCtx->SetElementInnerHtmlById "output", "ID: " & pWebCtx->GetActiveElementId & " KeyDown - Key: " & WSTR(lpMsg->wParam)
   END IF

   ' // When you use accelerator keys such as TAB, you may need to override the
   ' // default host behavior. The example shows how to do this.
    IF lpMsg->message = WM_KEYDOWN AND lpMsg->wParam = VK_TAB THEN
       RETURN S_FALSE   ' S_OK to disable tab navigation
    END IF
   ' // Return S_FALSE if you don't process the message
   RETURN S_FALSE
END FUNCTION
' ========================================================================================
```

# <a name="SetFocus"></a>SetFocus

Sets the focus in the hosted document (usually an html page).

```
FUNCTION SetFocus () AS HRESULT
```

#### Return value

Returns S_OK (0) to indicate that the operation was successful or an error code (E_NOINTERFACE).

# <a name="SetUIEventProc"></a>SetUIEventProc

Sets pointers to user implemented callback procedures to receive events of the **IDocHostUIHandler2** interface.

```
FUNCTION SetUIEventProc (BYVAL pwszEventName AS WSTRING PTR, BYVAL pProc AS ANY PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszEventName* | The name of the event. |
| *pProc* | The address of the callback procedure. |

| Parameter  |
| ---------- |
| ShowContextMenu |
| GetHostInfo |
| ShowUI |
| HideUI |
| UpdateUI |
| EnableModeless |
| OnDocWindowActivate |
| OnFrameWindowActivate |
| ResizeBorder |
| TranslateAccelerator |
| GetOptionKeyPath |
| GetDropTarget |
| GetExternal |
| TranslateUrl |
| FilterDataObject |
| GetOverrideKeyPath |

### Function callback prototypes

```
FUNCTION EnableModelessProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL fEnable AS WINBOOL) AS HRESULT
FUNCTION FilterDataObjectProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pDO AS IDataObject PTR, BYVAL ppDORet AS IDataObject PTR PTR) AS HRESULT
FUNCTION GetDropTargetProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pDropTarget AS IDropTarget PTR, BYVAL ppDropTarget AS IDropTarget PTR PTR) AS HRESULT
FUNCTION GetExternalProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL ppDispatch AS IDispatch PTR PTR) AS HRESULT
FUNCTION GetHostInfoProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pInfo AS DOCHOSTUIINFO PTR) AS HRESULT
FUNCTION GetOptionKeyPathProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pchKey AS LPOLESTR PTR, BYVAL dw AS DWORD) AS HRESULT
FUNCTION GetOverrideKeyPathProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pchKey AS LPOLESTR PTR, BYVAL dw AS DWORD) AS HRESULT
FUNCTION HideUIProc (BYVAL pWebCtx AS CWebCtx PTR) AS HRESULT
FUNCTION OnDocWindowActivateProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL fActivate AS WINBOOL) AS HRESULT
FUNCTION OnFrameWindowActivateProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL fActivate AS WINBOOL) AS HRESULT
FUNCTION ResizeBorderProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL prcBorder AS LPCRECT, BYVAL pUIWindow AS IOleInPlaceUIWindow PTR, BYVAL frameWindow AS WINBOOL) AS HRESULT
FUNCTION ShowContextMenuProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL dwID AS DWORD, BYVAL ppt AS POINT PTR, BYVAL pcmdtReserved AS IUnknown PTR, BYVAL pdispReserved AS IDispatch PTR) AS HRESULT
FUNCTION ShowUIProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL dwID AS DWORD, BYVAL pActiveObject AS IOleInPlaceActiveObject PTR, BYVAL pCommandTarget AS IOleCommandTarget PTR, BYVAL pFrame AS IOleInPlaceFrame PTR, BYVAL pDoc AS IOleInPlaceUIWindow PTR) AS HRESULT
FUNCTION TranslateAcceleratorProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL lpMsg AS LPMSG, BYVAL pguidCmdGroup AS const GUID PTR, BYVAL nCmdID AS DWORD) AS HRESULT
FUNCTION TranslateUrlProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL dwTranslate AS DWORD, BYVAL pchURLIn AS LPWSTR, BYVAL ppchURLOut AS LPWSTR PTR) AS HRESULT
FUNCTION UpdateUIProc (BYVAL pWebCtx AS CWebCtx PTR) AS HRESULT
```

#### Usage example

```
' // Add a WebBrowser control
DIM pwb AS CWebCtx = CWebCtx(@pWindow, IDC_WEBBROWSER, 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
' // Set web browser event callback procedures
pwb.SetEventProc("StatusTextChange", @WebBrowser_StatusTextChangeProc)
pwb.SetEventProc("DocumentComplete", @WebBrowser_DocumentCompleteProc)
pwb.SetEventProc("BeforeNavigate2", @WebBrowser_BeforeNavigate2Proc)
pwb.SetEventProc("HtmlDocumentEvents", @WebBrowser_HtmlDocumentEventsProc)
' // Set DocHostUI event callback procedures
pwb.SetUIEventProc("ShowContextMenu", @DocHostUI_ShowContextMenuProc)
pwb.SetUIEventProc("GetHostInfo", @DocHostUI_GetHostInfo)
pwb.SetUIEventProc("TranslateAccelerator", @DocHostUI_TranslateAccelerator)

' ========================================================================================
' Process the WebBrowser StatusTextChange event.
' ========================================================================================
SUB WebBrowser_StatusTextChangeProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pwszText AS WSTRING PTR)
   IF pwszText THEN StatusBar_SetText(GetDlgItem(GetParent(pWebCtx->hWindow), IDC_SATUSBAR), 0, pwszText)
END SUB
' ========================================================================================

' ========================================================================================
' Process the WebBrowser DocumentComplete event.
' ========================================================================================
SUB WebBrowser_DocumentCompleteProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR, BYVAL vUrl AS VARIANT PTR)
   ' // The vUrl parameter is a VT_BYREF OR VT_BSTR variant
   ' // It can be a VT_BSTR variant or a VT_ARRAY OR VT_UI1 with a pidl
   DIM varUrl AS VARIANT
   VariantCopyInd(@varUrl, vUrl)
   StatusBar_SetText(GetDlgItem(GetParent(pWebCtx->hWindow), IDC_SATUSBAR), 0, "Document complete: " & AfxVarToStr(@varUrl))
   VariantClear(@varUrl)
END SUB
' ========================================================================================

' ========================================================================================
' Process the IDocHostUIHandler ShowContextMenu event.
' ========================================================================================
FUNCTION DocHostUI_ShowContextMenuProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL dwID AS DWORD, BYVAL ppt AS POINT PTR, BYVAL pcmdtReserved AS IUnknown PTR, BYVAL pdispReserved AS IDispatch PTR) AS HRESULT
   ' // This event notifies that the user has clicked the right mouse button to show the
   ' // context menu. We can anulate it returning %S_OK and show our context menu.
   ' // Do not allow to show the context menu
'   AfxMsg "Sorry! Context menu disabled"
'   RETURN S_OK
   ' // Host did not display its UI. MSHTML will display its UI.
   RETURN S_FALSE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Process the IDocHostUIHandler GetHostInfo event.
' ========================================================================================
PRIVATE FUNCTION DocHostUI_GetHostInfo (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pInfo AS DOCHOSTUIINFO PTR) AS HRESULT
   IF pInfo THEN
      pInfo->cbSize = SIZEOF(DOCHOSTUIINFO)
      pInfo->dwFlags = DOCHOSTUIFLAG_NO3DBORDER OR DOCHOSTUIFLAG_THEME OR DOCHOSTUIFLAG_DPI_AWARE
      pInfo->dwDoubleClick = DOCHOSTUIDBLCLK_DEFAULT
      pInfo->pchHostCss = NULL
      pInfo->pchHostNS = NULL
   END IF
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Process the IDocHostUIHandler TranslateAccelerator event.
' ========================================================================================
PRIVATE FUNCTION DocHostUI_TranslateAccelerator (BYVAL pWebCtx AS CWebCtx PTR, BYVAL lpMsg AS LPMSG, BYVAL pguidCmdGroup AS const GUID PTR, BYVAL nCmdID AS DWORD) AS HRESULT

   IF lpMsg->message = WM_KEYDOWN THEN
      pWebCtx->SetElementInnerHtmlById "output", "ID: " & pWebCtx->GetActiveElementId & " KeyDown - Key: " & WSTR(lpMsg->wParam)
   END IF

   ' // When you use accelerator keys such as TAB, you may need to override the
   ' // default host behavior. The example shows how to do this.
    IF lpMsg->message = WM_KEYDOWN AND lpMsg->wParam = VK_TAB THEN
       RETURN S_FALSE   ' S_OK to disable tab navigation
    END IF
   ' // Return S_FALSE if you don't process the message
   RETURN S_FALSE
END FUNCTION
' ========================================================================================
```

# <a name="SetUIHandler"></a>SetUIHandler

Sets our implementation of the **IDocHostUIHandler** interface to customize the WebBrowser.

```
FUNCTION SetUIHandler () AS HRESULT
```

#### Remarks

You don't have to call this method. The **CWebCtx** constructor calls it.

# <a name="ShowSource"></a>ShowSource

Displays the source code of the page in an instance of NotePad.

```
FUNCTION ShowSource () AS HRESULT
```

#### Return value

S_OK (0) or an HRESULT code.

# <a name="Silent"></a>Silent

Sets or gets a value that indicates whether the object can display dialog boxes.

```
PROPERTY Silent (BYVAL bSilent AS VARIANT_BOOL)
PROPERTY Silent () AS VARIANT_BOOL
```

| Parameter  | Description |
| ---------- | ----------- |
| *bSilent* | VARIANT_FALSE Dialog boxes and messages can be displayed.<br>VARIANT_TRUE  Dialog boxes are not displayed. |

#### Return value

VARIANT_TRUE or VARIANT_FALSE.

# <a name="Stop"></a>Stop

Cancels any pending navigation or download operation and stops any dynamic page elements, such as background sounds and animations.

```
FUNCTION Stop () AS HRESULT
```

#### Return value

Returns S_OK (0) to indicate that the operation was successful.

# <a name="Unadvise"></a>Unadvise

Releases the events connection.

```
FUNCTION Unadvise () AS HRESULT
```

#### Return value

Returns S_OK (0) or an HRESULT code.

#### Remarks

You don't have to call this method. The **CWebCtx** destructor calls it.

# <a name="WaitForPageLoad"></a>WaitForPageLoad

Waits until the page had been fully downloaded or te timeout has expired.

```
FUNCTION FUNCTION WaitForPageLoad (BYVAL dbTimeout AS DOUBLE = 10) AS tagREADYSTATE
```

| Parameter  | Description |
| ---------- | ----------- |
| *dbTimeout* | Optional. Timeout in seconds. Default value: 10 seconds. |

#### Return value

| Value      | Description |
| ---------- | ----------- |
| READYSTATE_UNINITIALIZED = 0 | Default initialization state. |
| READYSTATE_LOADING = 1 | Object is currently loading its properties. |
| READYSTATE_LOADED = 2 | Object has been initialized. |
| READYSTATE_INTERACTIVE = 3 | Object is interactive, but not all of its data is available. |
| READYSTATE_COMPLETE = 4 | Object has received all of its data. |

# <a name="WriteHtml"></a>WriteHtml

Writes one or more HTML expressions to a document.

```
FUNCTION WriteHtml (BYREF cwsHtml AS CWSTR, BYVAL cr AS BOOLEAN = FALSE) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cwsHtml* | Text and HTML tags to write. |
| *cr* | Optional. Write the HTML text followed by a carriage return. |

#### Return value

S_OK if successful, or an error value otherwise.

#### Example

```
Example

' ########################################################################################
' Microsoft Windows
' Contents: WebBrowser with an animated GIF
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "Afx/CWindow.inc"
#INCLUDE ONCE "Afx/CWebBrowser/CWebBrowser.inc"
USING Afx

CONST IDC_WEBBROWSER = 1001

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG

   END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

' ========================================================================================
' Displays an animated gif file
' ========================================================================================
PRIVATE FUNCTION LoadGif (BYREF wszName AS WSTRING, BYVAL nWidth AS LONG, BYVAL nHeight AS LONG, BYREF wszAltName AS WSTRING = "") AS STRING

   DIM s AS STRING
   s  = "<!DOCTYPE html>"
   s += "<html>"
   s += "<head>"
   s += "<meta http-equiv='MSThemeCompatible' content='yes'>"
   s += "<title>Animated gif</title>"
   s += "</head>"
   s += "<body scroll='no' style='MARGIN: 0px 0px 0px 0px'>"
   s += "<img src='" & wszName & "' alt=' " & wszAltName & _
         "' style='width:" & STR(nWidth) & "px;height:" & STR(nHeight) & "px;'>"
   s += "</body>"
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
   DIM hwndMain AS HWND = pWindow.Create(NULL, "WebBrowser control - Animated gif", @WndProc)
   ' // Change the window style
   pWindow.WindowStyle = WS_OVERLAPPED OR WS_CAPTION OR WS_SYSMENU
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(400, 300)
   ' // Centers the window
   pWindow.Center

   ' // Add a WebBrowser control
   DIM pwb AS CWebBrowser = CWebBrowser(@pWindow, IDC_WEBBROWSER, 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   ' // Set the IDocHostUIHandler interface
   pwb.SetUIHandler

   ' // Load the gif file --> change the path to your file
   ' // The file can be located anywhere and you can use an url if it is located in the web
   DIM s AS STRING = LoadGif(ExePath & "\Circles-3.gif", _
       pWindow.ClientWidth, pWindow.ClientHeight, "Circles")
   ' // Write the script to the web page
   pwb.WriteHtml(s)

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
         SELECT CASE LOWORD(wParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF HIWORD(wParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF
         END SELECT

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

# <a name="AfxGetBrowserPtr"></a>AfxGetBrowserPtr

Returns a pointer to the hosted WebBrowser control given the handle of the form, or any control in the form, and the control identifier.

```
FUNCTION AfxGetBrowserPtr (BYVAL hwnd AS HWND, BYVAL cID AS WORD) AS Afx_IWebBrowser2 PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle of the window of the form or any of its child controls. |
| *cID* | Identifier of the control, e.g. IDC_WEBBROWSER. |

#### Return value

A pointer to the **IWeBbrowser2** interface or NULL.

# <a name="AfxGetActiveElementId"></a>AfxGetActiveElementId

Retrieves the ID of the active element (the object that has the focus when the parent document has focus).

```
FUNCTION AfxGetActiveElementId (BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebBrowser* | Pointer to the **IWebBrowser2** interface |

#### Return value

The ID of the active element, if any. If the element has not an identifier, it returns an empty string.

# <a name="AfxGetBodyInnerHtml"></a>AfxGetBodyInnerHtml

Returns a string that represents the text and html elements between the start and end body tags.

```
FUNCTION AfxGetBodyInnerHtml (BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebBrowser* | Pointer to the **IWebBrowser2** interface |

#### Return value

A string containing the html text.

# <a name="AfxGetBodyInnerText"></a>AfxGetBodyInnerText

Returns a string that represents the text between the start and end body tags without any associated HTML.

```
FUNCTION AfxGetBodyInnerText (BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebBrowser* | Pointer to the **IWebBrowser2** interface |

#### Return value

A string containing the html text.

# <a name="AfxGetElementInnerHtmlById"></a>AfxGetElementInnerHtmlById

Retrieves the HTML between the start and end tags of the object.

```
FUNCTION AfxGetElementInnerHtmlById (BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR, _
   BYREF cwsId AS CWSTR) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebBrowser* | Pointer to the **IWebBrowser2** interface. |
| *cwsId* | The identifier. |

#### Return value

A string containing the HTML text.

#### Remarks

This method performs a case insensitive property search. If two or more attributes have the same name (differing only in uppercase and lowercase letters) this function retrieves values only for the last attribute created with this name, and ignores all other attributes with the same name.

# <a name="AfxGetElementValueById"></a>AfxGetElementValueById

Retrieves the value attribute of the specified attribute.

```
FUNCTION AfxGetElementValueById (BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR, _
   BYREF cwsId AS CWSTR) AS VARIANT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebBrowser* | Pointer to the **IWebBrowser2** interface. |
| *cwsId* | The identifier. |

#### Return value

A variant containing the value as defined by the attribute.

#### Remarks

This method performs a case insensitive property search. If two or more attributes have the same name (differing only in uppercase and lowercase letters) this function retrieves values only for the last attribute created with this name, and ignores all other attributes with the same name.

# <a name="AfxSetElementFocusById"></a>AfxSetElementFocusById

Sets the focus in the specified element.

```
FUNCTION AfxSetElementFocusById (BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR, _
   BYREF cwsId AS CWSTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebBrowser* | Pointer to the **IWebBrowser2** interface. |
| *cwsId* | The identifier. |

#### Return value

S_OK if successful, or an error value otherwise.

#### Remarks

This method performs a case insensitive property search. If two or more attributes have the same name (differing only in uppercase and lowercase letters) this function sets values only for the last attribute created with this name, and ignores all other attributes with the same name.

# <a name="AfxSetElementInnerHtmlById"></a>AfxSetElementInnerHtmlById

Sets the HTML between the start and end tags of the object.

```
FUNCTION AfxSetElementInnerHtmlById (BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR, _
   BYREF cwsId AS CWSTR, BYREF cwsHtml AS CWSTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebBrowser* | Pointer to the **IWebBrowser2** interface. |
| *cwsId* | The identifier. |
| *cwsHtml* | The html text to set. |

#### Return value

S_OK if successful, or an error value otherwise.

#### Remarks

This method performs a case insensitive property search. If two or more attributes have the same name (differing only in uppercase and lowercase letters) this function sets values only for the last attribute created with this name, and ignores all other attributes with the same name.

# <a name="AfxSetElementValueById"></a>AfxSetElementValueById

Sets the value attribute of the specified identifier.

```
FUNCTION AfxSetElementValueById (BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR, _
   BYREF cwsId AS CWSTR, BYVAL vValue AS VARIANT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebBrowser* | Pointer to the **IWebBrowser2** interface. |
| *cwsId* | The identifier. |
| *vValue* | Variant that specifies the string, number, or Boolean to assign to the attribute. |

#### Return value

S_OK if successful, or an error value otherwise.

#### Remarks

This method performs a case insensitive property search. If two or more attributes have the same name (differing only in uppercase and lowercase letters) this function sets values only for the last attribute created with this name, and ignores all other attributes with the same name.

# <a name="AfxWriteHtml"></a>AfxWriteHtml

Writes one or more HTML expressions to a document.

```
FUNCTION AfxWriteHtml (BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR, _
   BYREF cwsHtml AS CWSTR, BYVAL cr AS BOOLEAN = FALSE) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebBrowser* | Pointer to the **IWebBrowser2** interface. |
| *cwsHtml* | Text and HTML tags to write. |
| *cr* | Write the HTML text followed by a carriage return. |

#### Return value

S_OK if successful, or an error value otherwise.

# <a name="BeforeNavigate2"></a>BeforeNavigate2 Event

Fires before navigation occurs in the given object (on either a window or frameset element).

```
SUB BeforeNavigate2 (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pDisp AS IDispatch PTR, _
   BYVAL vUrl AS VARIANT PTR, BYVAL vFlags AS VARIANT PTR, BYVAL vTargetFrameName AS VARIANT PTR, _
   BYVAL vPostData AS VARIANT PTR, BYVAL vHeaders AS VARIANT PTR, BYVAL pbCancel AS VARIANT_BOOL PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *pDisp* | Pointer to the **IDispatch** interface for the **WebBrowser** object that represents the window or frame. This interface can be queried for the **IWebBrowser2** interface.  |
| *vUrl* | Pointer to a VARIANT of type VT_BSTR that contains the URL to be navigated to |
| *vFlags* | Pointer to a VARIANT of type VT_I4 that contains the following flag, or zero.<br>**beforeNavigateExternalFrameTarget**: Windows Internet Explorer 7 or later. This navigate was the result of an external window or tab targeting this browser. |
| *vTargetFrameName* | Pointer to a VARIANT of type VT_BSTR that contains the name of the frame in which to display the resource, or NULL if no named frame is targeted for the resource. |
| *vPostData* | Pointer to a VARIANT of type VT_BYREF|VT_VARIANT that contains the data to send to the server if the HTTP POST transaction is being used. |
| *vHeaders* | Pointer to a VARIANT of type VT_BSTR that contains additional HTTP headers to send to the server (HTTPURLs only). The headers can specify things such as the action required of the server, the type of data being passed to the server, or a status code.  |
| *pbCancel* | In, Out. Pointer to a variable of type VARIANT_BOOL that contains the cancel flag. An application can set this parameter to VARIANT_TRUE to cancel the navigation operation, or to VARIANT_FALSE to allow it to proceed.  |

### Remarks

The post data specified by *vPostData* is passed as a SAFEARRAY structure. The variant should be of type VT_BYREF OR VT_VARIANT, which points to a SAFEARRAY. The SAFEARRAY should be of element type VT_UI1, dimension one, and have an element count equal to the number of bytes of post data.

The *vUrl* parameter can be a pointer to an item identifier list (PIDL) in the case of a shell namespace entity for which there is no URL representation.

The *pDisp* parameter specifies the WebBrowser object of the top-level frame corresponding to the navigation. Navigating to a different URL could happen as a result of external automation, internal automation from a script, or the user clicking a link or typing in the address bar. The processing of this navigation can be modified by setting the *pbCancel* parameter to VARIANT_TRUE and either ignoring or reissuing a modified navigation method to the WebBrowser object.

When reissuing a navigation for the WebBrowser object, the **Stop** method must first be executed for *pDisp*. This prevents the display of a web page that declares a cancelled navigation from appearing while the new navigation is being processed.

#### Example

The following example demonstrates how to redirect navigation to another url.

```
SUB WebBrowser_BeforeNavigate2Proc (BYVAL pAxHost AS CAxHost PTR, BYVAL pdisp AS IDispatch PTR, _
    BYVAL vUrl AS VARIANT PTR, BYVAL Flags AS VARIANT PTR, BYVAL TargetFrameName AS VARIANT PTR, _
    BYVAL PostData AS VARIANT PTR, BYVAL Headers AS VARIANT PTR, BYVAL pbCancel AS VARIANT_BOOL PTR)

   ' // Sample code to redirect navigation to another url
   IF AfxVarToStr(vUrl) = "http://com.it-berater.org/" THEN
      ' // Get a reference to the Afx_IWebBrowser2 interface
      DIM pwb AS Afx_IWebBrowser2 PTR = cast(Afx_IWebBrowser2 PTR, cast(ULONG_PTR, pdisp))
      IF pwb THEN
         ' // Stop loading the page
         pwb->Stop
         ' // Cancel the navigation operation
         *pbCancel = VARIANT_TRUE
         DIM vNewUrl AS VARIANT
         vNewUrl.vt = VT_BSTR
         vNewUrl.bstrVal = SysAllocString("http://www.planetsquires.com/protect/forum/index.php")
         pwb->Navigate2(@vNewUrl)
         VariantClear @vNewUrl
      END IF
   END IF

END SUB
' ========================================================================================
```

# <a name="ClientToHostWindow"></a>ClientToHostWindow Event

Fires to request that the client window size be converted to the host window size.

```
SUB ClientToHostWindow (BYVAL pWebCtx AS CWebCtx PTR, BYVAL CX AS LONG PTR, BYVAL CY AS LONG PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *cx* | In, Out. LONG variable that receives the width of the client window and can be set by the host application. |
| *cy* | In, Out. LONG that receives the height of the client window and can be set by the host application. |

#### Remarks

The **ClientToHostWindow** event gives the host application an opportunity to adjust the size of the WebBrowser control window, to account for any user interface items such as a toolbar, menu bar, or address bar.

This event is fired when a new window is opened through scripting, using the open method.

This event is available only to an application that is hosting the WebBrowser control installed by Microsoft Internet Explorer 5.5 and later.

# <a name="CommandStateChange"></a>CommandStateChange Event

Fires when the enabled state of a command changes.

```
SUB CommandStateChange (BYVAL pWebCtx AS CWebCtx PTR, BYVAL nCommand AS LONG, BYVAL fEnable AS VARIANT_BOOL)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *nCommand* | A **CommandStateChangeConstants** enumeration value that specifies the command that changed. |
| *fEnable* | Boolean value that specifies the enabled state.<br>VARIANT_FALSE: Command is disabled.<br>VARIANT_TRUE: Command is enabled. |

### CommandStateChangeConstants Enumeration

| Constant   | Description |
| ---------- | ----------- |
| CSC_UPDATECOMMANDS | The enabled state of a toolbar button might have changed. |
| CSC_NAVIGATEFORWARD | The enabled state of the Forward button has changed. |
| CSC_NAVIGATEBACK | The enabled state of the Back button has changed. |

# <a name="DocumentComplete"></a>DocumentComplete Event

Fires when a document has been completely loaded and initialized.

```
SUB DocumentComplete (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pDisp AS IDispatch PTR, BYVAL vURL AS VARIANT PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *pDisp* | Pointer to the **IDispatch** interface of the window or frame in which the document has loaded. This **IDispatch** interface can be queried for the **IWebBrowser2** interface. |
| *vURL* | Pointer to a variant of type VT_BSTR that specifies the URL, Universal Naming Convention (UNC) file name, or pointer to an item identifier list (PIDL) of the loaded document. |

#### Remarks

The value of the *vURL* parameter might not match the URL that was originally given to the WebBrowser Control. One possible reason for this is that the URL might be converted to a qualified form. For example, if an application specified a URL of www.microsoft.com in a call to the Navigate2 method of the **IWebBrowser2** interface, then the URL passed into **DocumentComplete** is http://www.microsoft.com/. In addition, if the server has redirected the browser to a different URL, the redirected URL is passed into the URL parameter.

The WebBrowser Control fires the **DocumentComplete** event when the document has completely loaded and the **READYSTATE** property has changed to **READYSTATE_COMPLETE**. Here are some important points regarding the firing of this event.

* In pages with no frames, this event fires once after loading is complete.
* In pages where multiple frames are loaded, this event fires for each frame where the **DownloadBegin** event has fired.
* This event's *pDisp* parameter is the same as the **IDispatch** interface pointer of the frame in which this event fires.
* In the loading process, the highest level frame (which is not necessarily the top-level frame) fires the final **DocumentComplete** event. At this time, the *pDisp* parameter will be the same as the **IDispatch** interface pointer of the highest level frame. 

Currently, the **DocumentComplete** does not fire when the **Visible** property of the WebBrowser Control is set to false.

# <a name="DownloadBegin"></a>DownloadBegin Event

Fires when a navigation operation is beginning.

```
SUB DownloadBegin (BYVAL pWebCtx AS CWebCtx PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |

#### Remarks

This event is fired shortly after the **BeforeNavigate2** event, unless the navigation is canceled. Any animation or "busy" indication that the container needs to display should be connected to this event.

Each **DownloadBegin** event will have a corresponding **DownloadComplete** event.

# <a name="DownloadComplete"></a>DownloadComplete Event

Fires when a navigation operation finishes, is halted, or fails.

```
SUB DownloadComplete (BYVAL pWebCtx AS CWebCtx PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |

#### Remarks

Unlike **NavigateComplete2**, which are fired only when a URL is successfully navigated to, this event is always fired after a navigation starts. Any animation or "busy" indication that the container needs to display should be connected to this event.

Each **DownloadBegin** event will have a corresponding **DownloadComplete** event.

# <a name="FileDownload"></a>FileDownload Event

Fires to indicate that a file download is about to occur. If a file download dialog is to be displayed, this event is fired prior to the display of the dialog.

```
SUB FileDownload (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pbCancel AS VARIANT_BOOL PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *pbCancel* | In, Out. Boolean value that specifies whether to continue the download process and display the download dialog.<br>VARIANT_FALSE: Default. Continue with the download process and display download dialog.<br>VARIANT_TRUE: Cancel the download process. |

#### Remarks

This event allows alternative action to be taken during a file download.

# <a name="HtmlDocumentEventsProc"></a>HtmlDocumentEventsProc Event

Provides access to properties and methods exposed by an object.

```
FUNCTION HtmlDocumentEventsProc (BYVAL pWebCtx AS CWebCtx PTR, BYVAL dispid AS LONG, _
   BYVAL pEvtObj AS IHTMLEventObj PTR) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *dispid* | The dispatch identifier member. |
| *pEvtObj* | Pointer to the **IHTMLEventObj** interface. |

#### Remarks

For cancelable document events return TRUE to indicate that Internet Explorer should perform its own event processing or FALSE to cancel the event.

#### Event DISPIDs

DISPID_HTMLELEMENTEVENTS2_ONHELP (-2147418102)<br>
Fires when the user presses the F1 key while the client is the active window. 

DISPID_HTMLELEMENTEVENTS2_ONCLICK (-600)<br>
Fires when the user clicks the left mouse button on the object.

DISPID_HTMLELEMENTEVENTS2_ONDBLCLICK (-601)<br>
Fires when the user double-clicks the object. 

DISPID_HTMLELEMENTEVENTS2_ONKEYPRESS (-603)<br>
Fires when the user presses an alphanumeric key.

DISPID_HTMLELEMENTEVENTS2_ONKEYDOWN (-602)<br>
Fires when the user presses a key.

DISPID_HTMLELEMENTEVENTS2_ONKEYUP (-604)<br>
Fires when the user releases a key.

DISPID_HTMLELEMENTEVENTS2_ONMOUSEOUT (-2147418103)<br>
Fires when the user moves the mouse pointer outside the boundaries of the object. 

DISPID_HTMLELEMENTEVENTS2_ONMOUSEOVER (-2147418104)<br>
Fires when the user moves the mouse pointer into the object. 

DISPID_HTMLELEMENTEVENTS2_ONMOUSEMOVE (-606)<br>
Fires when the user moves the mouse over the object. 

DISPID_HTMLELEMENTEVENTS2_ONMOUSEDOWN (-605)<br>
Fires when the user clicks the object with either mouse button. 

DISPID_HTMLELEMENTEVENTS2_ONMOUSEUP (-607)<br>
Fires when the user releases a mouse button while the mouse is over the object. 

DISPID_HTMLELEMENTEVENTS2_ONSELECTSTART (-2147418100)<br>
Fires when the object is being selected. 

DISPID_HTMLELEMENTEVENTS2_ONFILTERCHANGE (-2147418095)<br>
Fires when a visual filter changes state or completes a transition. 

DISPID_HTMLELEMENTEVENTS2_ONDRAGSTART (-2147418101)<br>
Fires on the source object when the user starts to drag a text selection or selected object. 

DISPID_HTMLELEMENTEVENTS2_ONBEFOREUPDATE (-2147418108)<br>
Fires on a databound object before updating the associated data in the data source object. 

DISPID_HTMLELEMENTEVENTS2_ONAFTERUPDATE (-2147418107)<br>
Fires on a databound object after successfully updating the associated data in the data source object. 

DISPID_HTMLELEMENTEVENTS2_ONERRORUPDATE (-2147418099)<br>
Fires on a databound object when an error occurs while updating the associated data in the data source object. 

DISPID_HTMLELEMENTEVENTS2_ONROWEXIT (-2147418106)<br>
Fires just before the data source control changes the current row in the object. 

DISPID_HTMLELEMENTEVENTS2_ONROWENTER (-2147418105)<br>
Fires to indicate that the current row has changed in the data source and new data values are available on the object. 

DISPID_HTMLELEMENTEVENTS2_ONDATASETCHANGED (-2147418098)<br>
Fires when the data set exposed by a data source object changes.

DISPID_HTMLELEMENTEVENTS2_ONDATAAVAILABLE (-2147418097)<br>
Fires periodically as data arrives from data source objects that asynchronously transmit their data. 

DISPID_HTMLELEMENTEVENTS2_ONDATASETCOMPLETE (-2147418096)<br>
Fires to indicate that all data is available from the data source object. 

DISPID_HTMLELEMENTEVENTS2_ONLOSECAPTURE (-2147418094)<br>
Fires when the object loses the mouse capture. 

DISPID_HTMLELEMENTEVENTS2_ONPROPERTYCHANGE (-2147418093)<br>
Fires when a property changes on the object.

DISPID_HTMLELEMENTEVENTS2_ONSCROLL (1014)<br>
Fires when the user repositions the scroll box in the scroll bar on the object. 

DISPID_HTMLELEMENTEVENTS2_ONFOCUS (-2147418111)<br>
Fires when the object receives focus. 

DISPID_HTMLELEMENTEVENTS2_ONBLUR (-2147418112)<br>
Fires when the object loses the input focus. 

DISPID_HTMLELEMENTEVENTS2_ONRESIZE (1016)<br>
Fires when the size of the object is about to change. 

DISPID_HTMLELEMENTEVENTS2_ONDRAG (-2147418092)<br>
Fires on the source object continuously during a drag operation.

DISPID_HTMLELEMENTEVENTS2_ONDRAGEND (-2147418091)<br>
Fires on the source object when the user releases the mouse at the close of a drag operation.

DISPID_HTMLELEMENTEVENTS2_ONDRAGENTER (-2147418090)<br>
Fires on the target element when the user drags the object to a valid drop target.

DISPID_HTMLELEMENTEVENTS2_ONDRAGOVER (-2147418089)<br>
Fires on the target element continuously while the user drags the object over a valid drop target.

DISPID_HTMLELEMENTEVENTS2_ONDRAGLEAVE (-2147418088)<br>
Fires on the target object when the user moves the mouse out of a valid drop target during a drag operation.

DISPID_HTMLELEMENTEVENTS2_ONDROP (-2147418087)<br>
Fires on the target object when the mouse button is released during a drag-and-drop operation.

DISPID_HTMLELEMENTEVENTS2_ONBEFORECUT (-2147418083)<br>
Fires on the source object before the selection is deleted from the document.

DISPID_HTMLELEMENTEVENTS2_ONCUT (-2147418086)<br>
Fires on the source element when the object or selection is removed from the document and added to the system clipboard.

DISPID_HTMLELEMENTEVENTS2_ONBEFORECOPY (-2147418082)<br>
Fires on the source object before the selection is copied to the system clipboard.

DISPID_HTMLELEMENTEVENTS2_ONCOPY (-2147418085)<br>
Fires on the source element when the user copies the object or selection, adding it to the system clipboard.

DISPID_HTMLELEMENTEVENTS2_ONBEFOREPASTE (-2147418081)<br>
Fires on the target object before the selection is pasted from the system clipboard to the document.

DISPID_HTMLELEMENTEVENTS2_ONPASTE (-2147418084)<br>
Fires on the target object when the user pastes data, transferring the data from the system clipboard to the document.

DISPID_HTMLELEMENTEVENTS2_ONCONTEXTMENU (1023)<br>
Fires when the user clicks the right mouse button in the client area, opening the context menu. 

DISPID_HTMLELEMENTEVENTS2_ONROWSDELETE (-2147418080)<br>
Fires when rows are about to be deleted from the recordset. 

DISPID_HTMLELEMENTEVENTS2_ONROWSINSERTED (-2147418079)<br>
Fires just after new rows are inserted in the current recordset.

DISPID_HTMLELEMENTEVENTS2_ONCELLCHANGE (-2147418078)<br>
Fires when data changes in the data provider.

DISPID_HTMLELEMENTEVENTS2_ONREADYSTATECHANGE (-609)<br>
Sets or retrieves the event handler for asynchronous requests.

DISPID_HTMLELEMENTEVENTS2_ONLAYOUTCOMPLETE (1030)<br>
Fires when the print or print preview layout process finishes filling the current LayoutRect object with content from the source document.

DISPID_HTMLELEMENTEVENTS2_ONPAGE (1031)<br>
This event is not implemented.

DISPID_HTMLELEMENTEVENTS2_ONMOUSEENTER (1042)<br>
Fires when the user moves the mouse pointer into the object.

DISPID_HTMLELEMENTEVENTS2_ONMOUSELEAVE (1043)<br>
Fires when the user moves the mouse pointer outside the boundaries of the object.

DISPID_HTMLELEMENTEVENTS2_ONACTIVATE (1044)<br>
Fires when the object is set as the active element.

DISPID_HTMLELEMENTEVENTS2_ONDEACTIVATE (1045)<br>
Fires when the activeElement is changed from the current object to another object in the parent document.

DISPID_HTMLELEMENTEVENTS2_ONBEFOREDEACTIVATE (1034)<br>
Fires immediately before the activeElement is changed from the current object to another object in the parent document.

DISPID_HTMLELEMENTEVENTS2_ONBEFOREACTIVATE (1047)<br>
Fires immediately before the object is set as the IHTMLDocument2::activeElement.

DISPID_HTMLELEMENTEVENTS2_ONFOCUSIN (1048)<br>
Fires for an element just prior to setting focus on that element.

DISPID_HTMLELEMENTEVENTS2_ONFOCUSOUT (1049)<br>
Fires for the current element with focus immediately after moving focus to another element. 

DISPID_HTMLELEMENTEVENTS2_ONMOVE (1035)<br>
Fires when the object moves.

DISPID_HTMLELEMENTEVENTS2_ONCONTROLSELECT (1036)<br>
Fires when the user is about to make a control selection of the object.

DISPID_HTMLELEMENTEVENTS2_ONMOVESTART (1038)<br>
Fires when the object starts to move.

DISPID_HTMLELEMENTEVENTS2_ONMOVEEND (1039)<br>
Fires when the object stops moving.

DISPID_HTMLELEMENTEVENTS2_ONRESIZESTART (1040)<br>
Fires when the user begins to change the dimensions of the object in a control selection.

DISPID_HTMLELEMENTEVENTS2_ONRESIZEEND (1041)<br>
Fires when the user finishes changing the dimensions of the object in a control selection.

DISPID_HTMLELEMENTEVENTS2_ONMOUSEWHEEL (1033)<br>
Fires when the wheel button is rotated. 

#### Example

```
pwb.SetEventProc("HtmlDocumentEvents", @WebBrowser_HtmlDocumentEventsProc)
```

```
PRIVATE FUNCTION WebBrowser_HtmlDocumentEventsProc (BYVAL pAxHost AS CAxHost PTR, BYVAL dispid AS LONG, BYVAL pEvtObj AS IHTMLEventObj PTR) AS BOOLEAN

   SELECT CASE dispid
      
      CASE DISPID_HTMLELEMENTEVENTS2_ONCLICK   ' // click event
         ' // Get a reference to the element that has fired the event
         DIM pElement AS IHTMLElement PTR
         IF pEvtObj THEN pEvtObj->lpvtbl->get_srcElement(pEvtObj, @pElement)
         IF pElement = NULL THEN EXIT FUNCTION
         DIM bstrHtml AS AFX_BSTR   ' // Outer html
         pElement->lpvtbl->get_outerHtml(pElement, @bstrHtml)
'         DIM bstrId AS AFX_BSTR   ' // identifier
'         pElement->lpvtbl->get_id(pElement, @bstrId)
         pElement->lpvtbl->Release(pElement)
         AfxMsg *bstrHtml
         SysFreeString bstrHtml
         RETURN TRUE
     
   END SELECT

   RETURN FALSE

END FUNCTION
```

Each **DownloadBegin** event will have a corresponding **DownloadComplete** event.

# <a name="NavigateComplete2"></a>NavigateComplete2 Event

Fires after a navigation to a link is completed on either a window or frameSet element.

```
SUB NavigateComplete2 (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pDisp AS IDispatch PTR, BYVAL vURL AS VARIANT PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *pDisp* | Pointer to the **IDispatch** interface for the WebBrowser object that represents the window or frame. This interface can be queried for the **IWebBrowser2** interface.  |
| *vURL* | Pointer to a variant that contains the URL, UNC file name, or PIDL that was navigated to. Note that this URL can be different from the URL that the browser was told to navigate to. One reason is that this URL is the canonicalized and qualified URL. For example, if an application specified a URL of "www.microsoft.com" in a call to the **Navigate2** method, the URL passed by Navigate2 would be "http://www.microsoft.com/". Also, if the server has redirected the browser to a different URL, the redirected URL will be reflected here. |

#### Remarks

The *vURL* parameter can be a PIDL in the case of a shell namespace entity for which there is no URL representation.

The document might still be downloading (and in the case of HTML, images might still be downloading), but at least part of the document has been received from the server, and the viewer for the document has been created.

In Internet Explorer 6 or later, the **Navigate2++ event fires only after the first navigation made in code. It does not fire when a user clicks a link on a Web page. 

# <a name="NavigateError"></a>NavigateError Event

Fires when an error occurs during navigation.

```
SUB NavigateError (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pDisp AS IDispatch PTR, _
   BYVAL vURL AS VARIANT PTR, BYVAL vTargetFrameName AS VARIANT PTR, _
   BYVAL vStatusCode AS VARIANT PTR, BYVAL pbCancel AS VARIANT_BOOL PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *pDisp* | Pointer to an **IDispatch** interface for the WebBrowser object that represents the window or frame in which the navigation error occurred. This interface can be queried for the **IWebBrowser2** interface. |
| *vURL* | Pointer to a VARIANT of type VT_BSTR that contains the URL for which navigation failed. |
| *vTargetFrameName* | Pointer to a VARIANT of type VT_BSTR that contains the name of the frame in which the resource is to be displayed, or NULL if no named frame was targeted for the resource.  |
| *vStatusCode* | Pointer to a VT_I4 containing an error status code, if available. For a list of the possible HRESULT and HTTP status codes, see [NavigateError Event Status Codes](https://msdn.microsoft.com/en-us/library/bb268233(v=vs.85).aspx). |
| *pbCancel* | In, Out- Pointer to a VARIANT of type VARIANT_BOOL that specifies whether to cancel the navigation to an error page and/or any further autosearch.<br>VARIANT_FALSE: Default. Continue with navigation to an error page and/or autosearch.<br>VARIANT_TRUE: Cancel navigation to an error page and/or autosearch. |

#### Remarks

This event fires before Internet Explorer displays an error page due to an error in navigation. An application has a chance to stop the display of the error page by setting the bCancel parameter to VARIANT_TRUE. However, if the server contacted in the original navigation supplies its own substitute page navigation, setting *pbCancel* to VARIANT_TRUE will have no effect and the navigation to the server's alternate page will proceed. For example, assume that a navigation to http://www.www.wingtiptoys.com/BigSale.htm causes this event to fire because the page does not exist. However, the server is set to redirect you to http://www.www.wingtiptoys.com/home.htm instead. In this case, setting *pbCancel* to VARIANT_TRUE has no effect and navigation will proceed to http://www.www.wingtiptoys.com/home.htm.

The *pDisp* parameter should be used to match this event with its corresponding **Navigate2** request. For example, multiple **NavigateError** events can fire for a single **Navigate2** request. Reasons for this include navigation to a URL with multiple frames or multiple attempts by an autosearch engine to resolve an invalid URL. In each of these cases, the URL passed into these events might not match the URL that was originally requested. However, each of these events will have the same *pDisp*.

As with other events, you can trap the **NavigateError** event by implementing **IDispatch.Invoke** as your event sink, and connecting it to the WebBrowser control's **DIID_DWebBrowserEvents2** connection point. You can then extract information such as the status code from the method's *pDispParams* parameter.

A URL passed into **Navigate2** might not match the URL passed into this event because the URL goes through a normalization process. For example, the URL string "www.wingtiptoys.co" could be passed into **Navigate2**, but because of the normalization process, the URL parameter is set to "http://www.wingtiptoys.co/".

In Internet Explorer 6 or later, the **NavigateError** event fires only after the first navigation made in code. It does not fire when a user clicks a link on a Web page. 

# <a name="NewWindow2"></a>NewWindow2 Event

Raised when a new window is to be created.

```
SUB NewWindow2 (BYVAL pWebCtx AS CWebCtx PTR, BYVAL ppDisp AS IDispatch PTR PTR, _
   BYVAL pbCancel AS VARIANT_BOOL PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *ppDisp* | In, Out. Address of an interface pointer that, optionally, receives the **IDispatch** interface pointer of a new WebBrowser or InternetExplorer object. |
| *pbCancel* | In, Out. Boolean value to determine whether the current navigation should be canceled.<br>VARIANT_TRUE: Cancel the navigation.<br>VARIANT_FALSE: Do not cancel the navigation. |

#### Remarks

Microsoft Internet Explorer 6 for Windows XP Service Pack 2 (SP2) or later. **NewWindow3** is raised instead of this event.

# <a name="NewWindow3"></a>NewWindow3 Event

Raised when a new window is to be created. Extends **NewWindow2** with additional information about the new window.

```
SUB NewWindow3 (BYVAL pWebCtx AS CWebCtx PTR, BYVAL ppDisp AS IDispatch PTR PTR, _
   BYVAL pbCancel AS VARIANT_BOOL PTR, BYVAL dwFlags AS DWORD, _
   BYVAL pwszUrlContext AS WSTRING PTR, BYVAL pwszUrl AS WSTRING PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *ppDisp* | In, Out. Address of an interface pointer that, optionally, receives the **IDispatch** interface pointer of a new WebBrowser or InternetExplorer object. |
| *pbCancel* | In, Out. Boolean value to determine whether the current navigation should be canceled.<br>VARIANT_TRUE: Cancel the navigation.<br>VARIANT_FALSE: Do not cancel the navigation. |
| *dwFlags* | Flags from the **NWMF** enumeration that pertain to the new window. |
| *pwszUrlContext* | URL of the page opening the new window. |
| *pwszUrl* | URL being opened in the new window. |

#### Remarks

**NewWindow3** is raised before **NewWindow2**.

In Microsoft Internet Explorer, the **NewWindow3** event is not raised when the user selects Window from the New command on the File menu. This event precedes the creation of a new window from within the WebBrowser. For example, NewWindow2 is raised in response to a navigation targeted to a new window, or from script using the **IHTMLWindow2.open** method.

The **NewWindow3** event is raised when a window is about to be created, such as during the following actions.

* The user clicks a link while pressing the SHIFT key.
* The user right-clicks a link and selects Open In New Window.
* The user selects New Window from the File menu.
* There is a targeted navigation to a frame name that does not yet exist.

Your browser application can also trigger this event by calling the **Navigate2** method with the **navOpenInNewWindow** flag. The WebBrowser control has an opportunity to handle the new window creation itself. If it does not, a top-level Internet Explorer window is created as a separate (nonhosted) process.

The application that processes this notification can respond in one of three ways.

* Create a new, hidden, nonnavigated WebBrowser or InternetExplorer object that is returned in *ppDisp*. Upon return from this event, the object that raised this event will then configure and navigate (including a **BeforeNavigate2** event) the new object to the target location.
* Cancel the navigation by setting *pbCancel* to VARIANT_TRUE.
* Do nothing and do not set *ppDisp* to any value. This will cause the object that raised the event to create a new InternetExplorer object to handle the navigation.

#### Remarks

Microsoft Internet Explorer 6 for Windows XP Service Pack 2 (SP2) or later. **NewWindow3** is raised instead of this event.

# <a name="OnVisible"></a>OnVisible Event

Fires when the **Visible** property of the object is changed.

```
METHOD OnVisible (BYVAL pWebCtx AS CWebCtx PTR, BYVAL bVisible AS VARIANT_BOOL)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *bVisible* | Pointer to the **CWebCtx** class.<br>Boolean value  that specifies whether the object is visible.<br>VARIANT_FALSE: Object is not visible.<br>VARIANT_TRUE: Object is visible. |

# <a name="PrintTemplateInstantiation"></a>PrintTemplateInstantiation Event

Fires when a print template has been instantiated.

```
SUB PrintTemplateInstantiation (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pDisp As IDispatch PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *pDisp* | Pointer to the **IDispatch** interface for the WebBrowser object that represents the window or frame. This interface can be queried for the **IWebBrowser2** interface.  |

# <a name="PrintTemplateTeardown"></a>PrintTemplateTeardown Event

Fires when a print template has been destroyed.

```
SUB PrintTemplateTeardown (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pDisp AS IDispatch PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *pDisp* | Pointer to the **IDispatch** interface for the WebBrowser object that represents the window or frame. This interface can be queried for the **IWebBrowser2** interface. |

# <a name="PrivacyImpactedStateChange"></a>PrivacyImpactedStateChange Event

Fired when an event occurs that impacts privacy or when a user navigates away from a URL that has impacted privacy.

```
SUB PrivacyImpactedStateChange (BYVAL pWebCtx AS CWebCtx PTR, BYVAL bImpacted AS VARIANT_BOOL)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *bImpacted* | Boolean value that specifies whether the current top-level URL has violated the browser's privacy settings.<br>VARIANT_TRUE: Privacy has been impacted.<br>VARIANT_FALSE: A user-initiated navigation from a URL with privacy violations has occurred. |

#### Remarks

The firing of this event corresponds to a change in the privacy state from impacted to unimpacted and vice-versa. A change in the privacy state coincides with the displaying or clearing of the privacy-impacted icon from the status bar. Although a URL's privacy policy might not agree with the browser's privacy settings, privacy is considered impacted only if violating cookie operations are attempted. This event only fires the first time a violating cookie action is attempted for a URL.

This event also fires when there is a user-initiated navigation away from a URL that has privacy violations. If the new URL has no record of privacy violations, the icon no longer displays and the privacy state remains as unimpacted. However, when navigating from a URL with privacy violations to revisit a URL that has a history of privacy violations (for example, the page has been retrieved from the cache), this event fires twice: once to signal that it is navigating away from the first violating URL and a second time to signal that it is navigating to a URL which has impacted privacy.

User-initiated navigations include:

* Typing a URL in the address bar
* Navigating using "Go To" option off of the "View" menu
* Choosing a URL from the Favorites List
* Clicking on a hyperlink that does not contain a script href
* Performing a manual top-level refresh
* Performing a top-level navigating using the Back and Forward buttons.

Privacy is considered impacted when any of the following occur:

* A cookie is suppressed when sending and HTTP request.
* A cookie is blocked from being written.
* A cached file is retrieved that has a history of privacy violations.

# <a name="ProgressChange"></a>ProgressChange Event

Fires when the progress of a download operation is updated on the object.

```
SUB ProgressChange (BYVAL pWebCtx AS CWebCtx PTR, BYVAL Progress AS LONG, BYVAL ProgressMax AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *Progress* | LONG that specifies the amount of total progress to show, or -1 when progress is complete. |
| *ProgressMax* | LONG that specifies the maximum progress value. |

#### Remarks

The container can use the information provided by this event to display the number of bytes downloaded so far or to update a progress indicator.

To calculate the percentage of progress to show in a progress indicator, multiply the value of *Progress* by 100 and divide by the value of *nProgressMax* (unless Progress is -1, in which case the container can indicate that the operation is finished or hide the progress indicator).

# <a name="PropertyChange"></a>PropertyChange Event

Fires when the **PutProperty** method of the object changes the value of a property.

```
SUB PropertyChange (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pwszProperty AS WSTRING PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *pwszProperty* | WSTRING that specifies the name of the property whose value has changed. |

#### Remarks

The WebBrowser object ignores this event.

# <a name="SetSecureLockIcon"></a>SetSecureLockIcon Event

Fires when there is a change in encryption level.

```
SUB SetSecureLockIcon (BYVAL pWebCtx AS CWebCtx PTR, BYVAL SecureLockIcon AS VARIANT)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *SecureLockIcon* | One of the **SecureLockIconConstants**. |

### SecureLockIconConstants Enumeration

| Constant   | Description |
| ---------- | ----------- |
| **secureLockIconUnsecure** | There is no security encryption present. |
| **secureLockIconMixed** | There are multiple security encryption methods present. |
| **secureLockIconSecureUnknownBits** | The security encryption level is not known. |
| **secureLockIconSecure40Bit** | There is 40-bit security encryption present. |
| **secureLockIconSecure56Bit** | There is 56-bit security encryption present. |
| **secureLockIconSecureFortezza** | There is Fortezza security encryption present. |
| **secureLockIconSecure128Bit** | There is 128-bit security encryption present. |

#### Remarks

This event fires when moving to and from a site that uses encryption. When moving between two sites that both use encryption, this event fires twice: once to denote a movement away from a site using encryption and once to denote a movement to a site using encryption.

The browser's status bar displays a lock icon when encryption is present. Mousing over this icon invokes a ToolTip indicating the level of encryption.

An application hosting the WebBrowser Control can use this event to display a user interface (UI) that shows the current encryption level.

Multiple levels of encryption can coexist in a single window in situations where there are multiple frames. If there are multiple levels of encryption present on a page, the SecureLockIcon has the value of **secureLockIconMixed**.

The encryption level on a page can be higher than the default level set in the browser if the site is using a special type of certificate (for example, a Server Gated Cryptography certificate).

# <a name="StatusTextChange"></a>StatusTextChange Event

Fires when the status bar text of the object has changed.

```
SUB StatusTextChange (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pwszText AS WSTRING PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *pwszText* | The container can use the information provided by this event to update the text of a status bar. |

# <a name="TitleChange"></a>TitleChange Event

Fires when the title of a document in the object becomes available or changes.

```
SUB TitleChange (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pwszText AS WSTRING PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *pwszText* | WSTRING that specifies the new status bar text. |

#### Remarks

Because the title might change while an HTML page is downloading, the URL of the document is set as the title. After the title specified in the HTML page, if there is one, is parsed, the title is changed to reflect the actual title.

# <a name="WindowClosing"></a>WindowClosing Event

Fires when the window of the object is about to be closed by script.

```
METHOD WindowClosing (BYVAL pWebCtx AS CWebCtx PTR, BYVAL IsChildWindow AS VARIANT_BOOL, _
   BYVAL pbCancel AS VARIANT_BOOL PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *IsChildWindow* | Boolean value that specifies whether the window was created from script. Can be one of the following values.<br>VARIANT_FALSE: Window was not created from script.<br>VARIANT_TRUE: Window was created from script. |
| *pbCancel* | In, Out. Pointer to a cancel flag. Boolean value that specifies whether the window is prevented from closing. Can be one of the following values.<br>VARIANT_FALSE: Window is allowed to close.<br>VARIANT_TRUE: Window is prevented from closing. |

#### Remarks

This event is fired when a window is closed from script, using the **window.close** method.

The default behavior of Microsoft Internet Explorer is to close windows that were created by script without asking the user. If an attempt is made to close the main InternetExplorer or WebBrowser control window through script, the user is prompted.

This event is available only to an application that is hosting the WebBrowser control installed by Internet Explorer 5.5 and later.

# <a name="WindowSetHeight"></a>WindowSetHeight Event

Fires when the object changes its height.

```
SUB WindowSetHeight (BYVAL pWebCtx AS CWebCtx PTR, BYVAL nHeight AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *nHeight* | LONG that specifies the new height of the WebBrowser control. |

#### Remarks

The **WindowSetHeight** event is fired when the height of a window is changed, using the Height property of the **IWebBrowser2** interface.

This event is also fired when a new window is opened through scripting, using the window.open method. The value of the nHeight parameter indicates the height requested in the call to **window.open**.

This event is available only to an application that is hosting the WebBrowser control installed by Microsoft Internet Explorer 5.5 and later.

# <a name="WindowSetLeft"></a>WindowSetLeft Event

Fires when the object changes its left position.

```
SUB WindowSetLeft (BYVAL pWebCtx AS CWebCtx PTR, BYVAL nLeft AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *nLeft* | LONG that specifies the new left position of the WebBrowser window. |

#### Remarks

The **WindowSetLeft** event is fired when the left position of a window is changed, using the **Left** property of the **IWebBrowser2** interface.

This event is also fired when a new window is opened through scripting, using the window.open method. The value of the *nLeft* parameter indicates the left position requested in the call to **window.open**.

This event is available only to an application that is hosting the WebBrowser control installed by Microsoft Internet Explorer 5.5 and later.

# <a name="WindowSetResizable"></a>WindowSetResizable Event

Fires to indicate whether the host window should allow or disallow resizing of the object.

```
SUB WindowSetResizable (BYVAL pWebCtx AS CWebCtx PTR, BYVAL bResizable AS VARIANT_BOOL)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *bResizable* | Boolean value that specifies whether the WebBrowser control is resizable. Can be one of the following values.<br>VARIANT_FALSE: Control is not resizable.<br>VARIANT_TRUE: Control is resizable. |

#### Remarks

The **WindowSetResizable** event is fired when a new window is opened through scripting, using the window.open method. The value of the *bResizable* parameter indicates whether the resizable feature was requested in the call to **window.open**.

The host window can use this event to decide whether to display resizing user interface items, such as a size box and maximize box.

This event is available only to an application that is hosting the WebBrowser control installed by Microsoft Internet Explorer 5.5 and later.

# <a name="WindowSetTop"></a>WindowSetTop Event

Fires when the object changes its top position.

```
SUB WindowSetTop (BYVAL pWebCtx AS CWebCtx PTR, BYVAL nTop AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *nTop* | LONG that specifies the top position of the WebBrowser control. |

#### Remarks

The **WindowSetTop** event is fired when the top position of a window is changed, using the **Top** property of the **IWebBrowser2** interface.

This event is also fired when a new window is opened through scripting, using the window.open method. The value of the nTop parameter indicates the top position requested in the call to **window.open**.

This event is available only to an application that is hosting the WebBrowser control installed by Microsoft Internet Explorer 5.5 and later.

# <a name="WindowSetWidth"></a>WindowSetWidth Event

Fires when the object changes its width.

```
SUB WindowSetWidth (BYVAL pWebCtx AS CWebCtx PTR, BYVAL nWidth AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *nWidth* | LONG that specifies the width of the WebBrowser control. |

#### Remarks

The **WindowSetWidth** event is fired when the width of a window is changed, using the **Width** property of the **IWebBrowser2** interface.

This event is also fired when a new window is opened through scripting, using the window.open method. The value of the **Width** parameter indicates the width requested in the call to **window.open**.

This event is available only to an application that is hosting the WebBrowser control installed by Microsoft Internet Explorer 5.5 and later.

# <a name="WindowStateChanged"></a>WindowStateChanged Event

Fires when the progress of a download operation is updated on the object.

```
SUB WindowStateChanged (BYVAL pWebCtx AS CWebCtx PTR, BYVAL dwFlags AS LONG, BYVAL dwValidFlagsMask AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *dwFlags* | The flags indicating the current window state.<br>**1**: The content window is visible to the user.<br>**2**: The content window is enabled. |
| *dwValidFlagsMask* | The flags indicating which flags in the dwFlags parameter value have been initialized.<br>**1**: The OLECMDIDF_WINDOWSTATE_USERVISIBLE flag can be checked.<br>**2**: The OLECMDIDF_WINDOWSTATE_ENABLED flag can be checked. |

#### Remarks

**WindowStateChanged** is available only in Windows XP Service Pack 2 (SP2) or later.

The **WindowStateChanged** event is raised when the state of a content window, such as the browser window or a tab, might have changed. The following actions raise this event.

* The browser window is minimized or restored.
* An active tab becomes inactive.
* An inactive tab becomes active.
* The browser window is enabled or disabled due to a modal dialog box.

A content window is visible to the user when it is displayed to the user and can be interacted with. If tabbed browsing is enabled, the active tab (the one with focus) contains the content window. Background tabs are inactive. When tabbed browsing is disabled, the browser window displays the content window. When the browser window is minimized, the content window is not visible. This event can be used to minimize CPU usage and prolong battery life by reducing unnecessary updates to inactive windows.

**Note**: This event can be raised even if the state of the parameter flag values have not changed.

# <a name="EnableModeless"></a>EnableModeless Event

Called by the MSHTML implementation of **IOleInPlaceActiveObject.EnableModeless**. Also called when MSHTML displays a modal UI.

```
FUNCTION EnableModeless (BYVAL pWebCtx AS CWebCtx PTR, BYVAL fEnable AS LONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *fEnable* | Boolean value that indicates if the host's modeless dialog boxes are enabled or disabled.<br>**TRUE**: Modeless dialog boxes are enabled.<br>**FALSE**: Modeless dialog boxes are disabled. |

#### Return value

Returns S_OK (0) if successful, or an error value otherwise. 

# <a name="FilterDataObject"></a>FilterDataObject Event

Called by MSHTML to allow the host to replace the MSHTML data object.

```
FUNCTION FilterDataObject (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pDO AS IDataObject PTR, _
   BYVAL ppDORet AS IDataObject PTR PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *pDO* | Pointer to an **IDataObject** interface supplied by MSHTML. |
| *ppDORet* | Out. Address of a pointer variable that receives an **IDataObject** interface pointer supplied by the host. |

#### Return value

Returns S_OK (0) if the data object is replaced, or S_FALSE (1) if it's not replaced.

#### Remarks

This method enables the host to block certain clipboard formats or support additional clipboard formats.

If the implementation of this method does not supply its own **IDataObject**, *ppDORet* should be set to NULL, even if the method fails or returns S_FALSE.

# <a name="GetDropTarget"></a>GetDropTarget Event

Called by MSHTML when it is used as a drop target. This method enables the host to supply an alternative **IDropTarget** interface.

```
FUNCTION GetDropTarget (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pDropTarget AS IDropTarget PTR, _
   BYVAL ppDropTarget AS IDropTarget PTR PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *pDropTarget* | Pointer to an **IDropTarget** interface for the current drop target object supplied by MSHTML. |
| *ppDropTarget* | Out. Address of a pointer variable that receives an **IDropTarget** interface pointer for the alternative drop target object supplied by the host. |

#### Return value

Returns S_OK (0) if successful, or an error value otherwise.

#### Remarks

If the host does not supply an alternative drop target, this method should return a failure code, such as E_NOTIMPL or E_FAIL.

# <a name="GetExternal"></a>GetExternal Event

Called by MSHTML to obtain the host's **IDispatch** interface.

```
FUNCTION GetExternal (BYVAL pWebCtx AS CWebCtx PTR, BYVAL ppDispatch AS IDispatch PTR PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *ppDispatch* | Out. Address of a pointer to a variable that receives an **IDispatch** interface pointer for the host application. |

#### Return value

Returns S_OK (0) if successful, or an error value otherwise.

#### Remarks

If the host exposes an automation interface, it can provide a reference to MSHTML through ppDispatch.

If the method implementation does not supply an **IDispatch**, *ppDispatch* should be set to NULL, even if the method fails or returns S_FALSE.

# <a name="GetHostInfo"></a>GetHostInfo Event

Called by MSHTML to retrieve the user interface (UI) capabilities of the application that is hosting MSHTML.

```
FUNCTION GetHostInfo (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pInfo AS DOCHOSTUIINFO PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *ppDispatch* | Out. Pointer to a **DOCHOSTUIINFO** structure that receives the host's UI capabilities.  |

#### Return value

Returns S_OK (0) if successful, or an error value otherwise.

### DOCHOSTUIINFO structure

```
TYPE DOCHOSTUIINFO
   cbSize AS ULONG
   dwFlags AS DWORD
   dwDoubleClick AS DWORD
   pchHostCss AS OLECHAR PTR
   pchHostNS AS OLECHAR PTR
END TYPE
```

#### Members

| Member     | Description |
| ---------- | ----------- |
| **cbSize** | DWORD containing the size of this structure, in bytes. |
| **dwFlags** | One or more of the DOCHOSTUIFLAG values that specify the UI capabilities of the host. |
| **dwDoubleClick** | One of the **DOCHOSTUIDBLCLK** values that specify the operation that should take place in response to a double-click. |
| **pchHostCss** | Pointer to a set of Cascading Style Sheets (CSS) rules sent down by the host. These CSS rules affect the page containing them. |
| **pchHostNS** | Pointer to a semicolon-delimited namespace list. This list allows the host to supply a namespace declaration for custom tags on the page. |

### DOCHOSTUIFLAG values

| Flag       | Description |
| ---------- | ----------- |
| DOCHOSTUIFLAG_DIALOG | MSHTML does not enable selection of the text in the form. |
| DOCHOSTUIFLAG_DISABLE_HELP_MENU | MSHTML does not add the Help menu item to the container's menu. |
| DOCHOSTUIFLAG_NO3DBORDER | MSHTML does not use 3-D borders on any frames or framesets. To turn the border off on only the outer frameset use DOCHOSTUIFLAG_NO3DOUTERBORDER. |
| DOCHOSTUIFLAG_SCROLL_NO | MSHTML does not have scroll bars. |
| DOCHOSTUIFLAG_DISABLE_SCRIPT_INACTIVE | MSHTML does not execute any script until fully activated. This flag is used to postpone script execution until the host is active and, therefore, ready for script to be executed. |
| DOCHOSTUIFLAG_OPENNEWWIN | MSHTML opens a site in a new window when a link is clicked rather than browse to the new site using the same browser window. |
| DOCHOSTUIFLAG_DISABLE_OFFSCREEN | Not implemented. |
| DOCHOSTUIFLAG_FLAT_SCROLLBAR | MSHTML uses flat scroll bars for any UI it displays. |
| DOCHOSTUIFLAG_DIV_BLOCKDEFAULT | MSHTML inserts the div tag if a return is entered in edit mode. Without this flag, MSHTML will use the p tag. |
| DOCHOSTUIFLAG_ACTIVATE_CLIENTHIT_ONLY | MSHTML only becomes UI active if the mouse is clicked in the client area of the window. It does not become UI active if the mouse is clicked on a non-client area, such as a scroll bar. |
| DOCHOSTUIFLAG_OVERRIDEBEHAVIORFACTORY | MSHTML consults the host before retrieving a behavior from the URL specified on the page. If the host does not support the behavior, MSHTML does not proceed to query other hosts or instantiate the behavior itself, even for behaviors developed in script (HTML Components (HTCs)). |
| DOCHOSTUIFLAG_CODEPAGELINKEDFONTS | Microsoft Internet Explorer 5 and later. Provides font selection compatibility for Microsoft Outlook Express. If the flag is enabled, the displayed characters are inspected to determine whether the current font supports the code page. If disabled, the current font is used, even if it does not contain a glyph for the character. This flag assumes that the user is using Internet Explorer 5 and Outlook Express 4.0. |
| DOCHOSTUIFLAG_URL_ENCODING_DISABLE_UTF8 | Internet Explorer 5 and later. Controls how nonnative URLs are transmitted over the Internet. Nonnative refers to characters outside the multibyte encoding of the URL. If this flag is set, the URL is not submitted to the server in UTF-8 encoding. |
| DOCHOSTUIFLAG_URL_ENCODING_ENABLE_UTF8 | Internet Explorer 5 and later. Controls how nonnative URLs are transmitted over the Internet. Nonnative refers to characters outside the multibyte encoding of the URL. If this flag is set, the URL is submitted to the server in UTF-8 encoding. |
| DOCHOSTUIFLAG_ENABLE_FORMS_AUTOCOMPLETE | Internet Explorer 5 and later. Enables the AutoComplete feature for forms in the hosted browser. The Intelliforms feature is only turned on if the user has previously enabled it. If the user has turned the AutoComplete feature off for forms, it is off whether this flag is specified or not. |
| DOCHOSTUIFLAG_ENABLE_INPLACE_NAVIGATION | Internet Explorer 5 and later. Enables the host to specify that navigation should happen in place. This means that applications hosting MSHTML directly can specify that navigation happen in the application's window. For instance, if this flag is set, you can click a link in HTML mail and navigate in the mail instead of opening a new Windows Internet Explorer window. |
| DOCHOSTUIFLAG_IME_ENABLE_RECONVERSION | Internet Explorer 5 and later. During initialization, the host can set this flag to enable Input Method Editor (IME) reconversion, allowing computer users to employ IME reconversion while browsing Web pages. An input method editor is a program that allows users to enter complex characters and symbols, such as Japanese Kanji characters, using a standard keyboard. For more information, see the International Features reference in the Base Services section of the Windows Software Development Kit (SDK). |
| DOCHOSTUIFLAG_THEME | Internet Explorer 6 and later. Specifies that the hosted browser should use themes for pages it displays. |
| DOCHOSTUIFLAG_NOTHEME | Internet Explorer 6 and later. Specifies that the hosted browser should not use themes for pages it displays. |
| DOCHOSTUIFLAG_NOPICS | Internet Explorer 6 and later. Disables PICS ratings for the hosted browser. |
| DOCHOSTUIFLAG_NO3DOUTERBORDER | Internet Explorer 6 and later. Turns off any 3-D border on the outermost frame or frameset only. To turn borders off on all frame sets, use DOCHOSTUIFLAG_NO3DBORDER. |
| DOCHOSTUIFLAG_DISABLE_EDIT_NS_FIXUP | Internet Explorer 6 and later. Disables the automatic correction of namespaces when editing HTML elements. |
| DOCHOSTUIFLAG_LOCAL_MACHINE_ACCESS_CHECK | Internet Explorer 6 and later. Prevents Web sites in the Internet zone from accessing files in the Local Machine zone. |
| DOCHOSTUIFLAG_DISABLE_UNTRUSTEDPROTOCOL | Internet Explorer 6 and later. Turns off untrusted protocols. Untrusted protocols include ms-its, ms-itss, its, and mk:@msitstore. |
| DOCHOSTUIFLAG_HOST_NAVIGATES | Internet Explorer 7. Indicates that navigation is delegated to the host; otherwise, MSHTML will perform navigation. This flag is used primarily for non-HTML document types. |
| DOCHOSTUIFLAG_ENABLE_REDIRECT_NOTIFICATION | Internet Explorer 7. Causes MSHTML to fire an additional DWebBrowserEvents2.BeforeNavigate2 event when redirect navigations occur. Applications hosting the WebBrowser Control can choose to cancel or continue the redirect by returning an appropriate value in the Cancel parameter of the event. |
| DOCHOSTUIFLAG_USE_WINDOWLESS_SELECTCONTROL | Internet Explorer 7. Causes MSHTML to use the Document Object Model (DOM) to create native "windowless" select controls that can be visually layered under other elements. |
| DOCHOSTUIFLAG_USE_WINDOWED_SELECTCONTROL | Internet Explorer 7. Causes MSHTML to create standard Microsoft Win32 "windowed" select and drop-down controls. |
| DOCHOSTUIFLAG_ENABLE_ACTIVEX_INACTIVATE_MODE | Internet Explorer 6 for Windows XP Service Pack 2 (SP2) and later. Requires user activation for Microsoft ActiveX controls and Java Applets embedded within a web page. This flag enables interactive control blocking, which provisionally disallows direct interaction with ActiveX controls loaded by the APPLET, EMBED, or OBJECT elements. When a control is inactive, it does not respond to user input; however, it can perform operations that do not involve interaction. |
| DOCHOSTUIFLAG_DPI_AWARE | Internet Explorer 8. Causes layout engine to calculate document pixels as 96 dots per inch (dpi). Normally, a document pixel is the same size as a screen pixel. This flag is equivalent to setting the FEATURE_96DPI_PIXEL feature control key on a per-host basis. |

#### Remarks

The DOCHOSTUIFLAG_BROWSER flag, a supplementary defined constant (not technically a part of this enumeration), combines the values of DOCHOSTUIFLAG_DISABLE_HELP_MENU and DOCHOSTUIFLAG_DISABLE_SCRIPT_INACTIVE.

# <a name="GetOptionKeyPath"></a>GetOptionKeyPath Event

Called by the WebBrowser Control to retrieve a registry subkey path that overrides the default Microsoft Internet Explorer registry settings.

```
FUNCTION GetOptionKeyPath (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pchKey AS LPOLESTR PTR, _
   BYVAL dw AS DWORD) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *pchKey* | Out. Pointer to an LPOLESTR that receives the registry subkey string where the host stores its registry settings. |
| *dw* | Reserved. Must be set to NULL. |

#### Return value

Returns S_OK (0) if successful, or an error value otherwise.

#### Remarks

A WebBrowser Control instance calls **GetOptionKeyPath** on the host at initialization so that the host can specify a registry location containing settings that override the default Internet Explorer registry settings. If the host returns S_FALSE for this method, or if the registry key path returned to the WebBrowser Control in pchKey is NULL or empty, the WebBrowser Control reverts to the default Internet Explorer registry settings.

**GetOptionKeyPath** and **GetOverrideKeyPath** provide two alternate mechanisms for a WebBrowser Control host to make changes in the registry settings for the WebBrowser Control. With **GetOptionKeyPath**, a WebBrowser Control instance defaults to its original settings before registry changes are applied from the registry path specified by the method. With **GetOptionKeyPath**, a WebBrowser Control instance preserves registry settings for the current user. Any registry changes located at the registry path specified by this method override those located in HKEY_CURRENT_USER/Software/Microsoft/Internet Explorer.

For example, assume that the user has changed the Internet Explorer default text size to the largest font. Implementing **GetOverrideKeyPath** preserves that change—unless the size has been specifically overridden in the registry settings located at the registry path specified by the implementation of **GetOptionKeyPath**. Implementing **GetOptionKeyPath** would not preserve the user's text size change. Instead, the WebBrowser Control defaults back to its original medium-size font before applying registry settings from the registry path specified by the **GetOptionKeyPath** implementation.

An implementation of **GetOptionKeyPath** must allocate memory for *pchKey* using **CoTaskMemAlloc**. (The WebBrowser Control is responsible for freeing this memory using **CoTaskMemFree**). Even if this method is unimplemented, the parameter should be set to NULL.

The key specified by this method must be a subkey of the HKEY_CURRENT_USER key.

#### C++ Example

This example points the WebBrowser Control to a registry key located at HKEY_CURRENT_USER/Software/YourCompany/YourApp for Internet Explorer registry overrides. Of course, you need to set registry keys at this location in the registry for the WebBrowser Control to use them.

```
HRESULT CBrowserHost::GetOptionKeyPath(LPOLESTR *pchKey, DWORD dwReserved)
{
    HRESULT hr;
    WCHAR* szKey = L"Software\\MyCompany\\MyApp";
	
    //  cbLength is the length of szKey in bytes.
    size_t cbLength;
    hr = StringCbLengthW(szKey, 1280, &cbLength);
    //  TODO: Add error handling code here.
    
    if (pchKey)
    {
        *pchKey = (LPOLESTR)CoTaskMemAlloc(cbLength + sizeof(WCHAR));
        if (*pchKey)
            hr = StringCbCopyW(*pchKey, cbLength + sizeof(WCHAR), szKey);
    }
    else
        hr = E_INVALIDARG;

    return hr;
}
```

# <a name="GetOverrideKeyPath"></a>GetOverrideKeyPath Event

Called by the WebBrowser Control to retrieve a registry subkey path that modifies Microsoft Internet Explorer user preferences.

```
FUNCTION GetOverrideKeyPath (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pchKey AS LPOLESTR PTR, _
   BYVAL dw AS DWORD) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *pchKey* | Out. Pointer to an LPOLESTR that receives the registry subkey string where the host stores its registry settings. |
| *dw* | Reserved. Must be set to NULL. |

#### Return value

Returns S_OK (0) if successful, or an error value otherwise.

#### Remarks

A WebBrowser Control instance calls **GetOverrideKeyPath** on the host at initialization so that the host can specify a registry location containing settings that modify the Internet Explorer registry settings for the current user. If the host returns S_FALSE for this method, or if the registry key path returned to the WebBrowser Control in pchKey is NULL or empty, the WebBrowser Control reverts to the registry settings for the current user.

**GetOverrideKeyPath** provides an alternate mechanism to **GetOptionKeyPath** for a WebBrowser Control host to make changes in the registry settings for the WebBrowser Control. With **GetOverrideKeyPath**, your WebBrowser Control instance preserves registry settings for the current user. Any registry changes located at the registry path specified by this method override those located in HKEY_CURRENT_USER/Software/Microsoft/Internet Explorer. Compare this to **GetOptionKeyPath**, which causes a WebBrowser Control instance to default to its original settings before registry changes are applied from the registry path specified by the method.

For example, assume the user has changed the Internet Explorer default text size to the largest font. Implementing **GetOverrideKeyPath** will preserve that change—unless the size has been specifically overridden in the registry settings located at the registry path specified by the implementation of **GetOverrideKeyPath**. Implementing **GetOptionKeyPath** would not preserve the user's text size change. Instead, the WebBrowser Control defaults back to its original medium-size font before applying registry settings from the registry path specified by the **GetOptionKeyPath** implementation.

An implementation of **GetOverrideKeyPath** must allocate memory for pchKey using **CoTaskMemAlloc**. (The WebBrowser Control will be responsible for freeing this memory using **CoTaskMemFree**). Even if this method is unimplemented, this parameter should be set to NULL.

The key specified by this method must be a subkey of the HKEY_CURRENT_USER key.

#### C++ Example

This example points the WebBrowser Control to a registry key located at HKEY_CURRENT_USER/Software/YourCompany/YourApp for user preference overrides. Of course, you need to set registry keys at this location in the registry for the WebBrowser Control to use them.

```
HRESULT CBrowserHost::GetOverrideKeyPath(LPOLESTR *pchKey, DWORD dwReserved)
{
    HRESULT hr;
    WCHAR* szKey = L"Software\\MyCompany\\MyApp";
	
    //  cbLength is the length of szKey in bytes.
    size_t cbLength;
    hr = StringCbLengthW(szKey, 1280, &cbLength);
    //  TODO: Add error handling code here.
	
    if (pchKey)
    {
        *pchKey = (LPOLESTR)CoTaskMemAlloc(cbLength + sizeof(WCHAR));
        if (*pchKey)
            hr = StringCbCopyW(*pchKey, cbLength + sizeof(WCHAR), szKey);
    }
    else
        hr = E_INVALIDARG;

    return hr;
}
```

# <a name="HideUI"></a>HideUI Event

Called when MSHTML removes its menus and toolbars.

```
FUNCTION HideUI (BYVAL pWebCtx AS CWebCtx PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |

#### Return value

Returns S_OK (0) if successful, or an error value otherwise.

#### Remarks

If a host displayed menus and toolbars during the call to **ShowUI**, the host should remove them when this method is called. This method is called regardless of the return value from the **ShowUI** method.

# <a name="OnDocWindowActivate"></a>OnDocWindowActivate Event

Called by the MSHTML implementation of **IOleInPlaceActiveObject.OnDocWindowActivate**.

```
FUNCTION OnDocWindowActivate (BYVAL pWebCtx AS CWebCtx PTR, BYVAL fActivate AS WINBOOL) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *fActivate* | Boolean value that indicates the state of the document window.<br>**TRUE**: The window is being activated.<br>**FALSE**: The window is being deactivated. |

#### Return value

Returns S_OK (0) if successful, or an error value otherwise.

# <a name="OnFrameWindowActivate"></a>OnFrameWindowActivate Event

Called by the MSHTML implementation of **IOleInPlaceActiveObject.OnFrameWindowActivate**.

```
FUNCTION OnFrameWindowActivate (BYVAL pWebCtx AS CWebCtx PTR, BYVAL fActivate AS WINBOOL) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *fActivate* | Boolean value that indicates the state of the container's top-level frame window.<br>**TRUE**: The frame window is being activated.<br>**FALSE**: The frame window is being deactivated. |

#### Return value

Returns S_OK (0) if successful, or an error value otherwise.

# <a name="ResizeBorder"></a>ResizeBorder Event

Called by the MSHTML implementation of **IOleInPlaceActiveObject.ResizeBorder**.

```
FUNCTION ResizeBorder (BYVAL pWebCtx AS CWebCtx PTR, BYVAL prcBorder AS LPCRECT, _
   BYVAL pUIWindow AS IOleInPlaceUIWindow PTR, BYVAL fFrameWindow AS WINBOOL) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *prcBorder* | Constant pointer to a **RECT** structure for the new outer rectangle of the border. |
| *pUIWindow* | Pointer to an **IOleInPlaceUIWindow** interface for the frame or document window whose border is to be changed. |
| *fFrameWindow* | Boolean value that is TRUE if the frame window is calling **ResizeBorder**, or FALSE otherwise.  |

#### Return value

Returns S_OK (0) if successful, or an error value otherwise.

# <a name="ShowContextMenu"></a>ShowContextMenu Event

Called by MSHTML to display a shortcut menu.

```
FUNCTION ShowContextMenu (BYVAL pWebCtx AS CWebCtx PTR, BYVAL dwID AS DWORD, _
   BYVAL ppt AS POINT PTR, BYVAL pcmdtReserved AS IUnknown PTR, _
   BYVAL pdispReserved AS IDispatch PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *dwID* | DWORD that specifies the identifier of the shortcut menu to be displayed. This identifier is a bitwise shift of the value 0x1 by the shortcut menu values (e.g., CONTEXT_MENU_DEFAULT) defined in Mshtmhst.inc.<br>**&h2**: value of (&h1 << CONTEXT_MENU_DEFAULT)<br>**&h4**: value of (&h1 << CONTEXT_MENU_CONTROL)<br>**h8**: value of (&h1 << CONTEXT_MENU_TABLE)<br>**&h10**: value of (&h1 << CONTEXT_MENU_TEXTSELECT)<br>**&h30**: value of (&h1 << CONTEXT_MENU_ANCHOR)<br>**&h20**: value of (&h1 << CONTEXT_MENU_UNKNOWN) |
| *ppt* | Pointer to a **POINT** structure containing the screen coordinates for the menu. |
| *pcmdtReserved* | Pointer to the **IUnknown** of an **IOleCommandTarget** interface used to query command status and execute commands on this object. |
| *pdispReserved* | Pointer to an **IDispatch** interface of the object at the screen coordinates specified in ppt. This allows a host to differentiate particular objects to provide more specific context. |

#### Return value

Returns one of the following values:

| Value      | Meaning     |
| ---------- | ----------- |
| S_OK | Host displayed its own user interface (UI). MSHTML will not attempt to display its UI. |
| S_FALSE | Host did not display any UI. MSHTML will display its UI. |
| DOCHOST_E_UNKNOWN | Menu identifier is unknown. MSHTML may attempt an alternative identifier from a previous version. |

# <a name="ShowUI"></a>ShowUI Event

Called by MSHTML to enable the host to replace MSHTML menus and toolbars.

```
FUNCTION ShowUI (BYVAL pWebCtx AS CWebCtx PTR, BYVAL dwID AS DWORD, _
   BYVAL pActiveObject AS IOleInPlaceActiveObject PTR, BYVAL pCommandTarget AS IOleCommandTarget PTR, _
   BYVAL pFrame AS IOleInPlaceFrame PTR, BYVAL pDoc AS IOleInPlaceUIWindow PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *dwID* | DWORD that receives a **DOCHOSTUITYPE** value indicating the type of user interface (UI). |
| *pActiveObject* | Pointer to an **IOleInPlaceActiveObject** interface for the active object. |
| *pCommandTarget* | Pointer to an **IOleCommandTarget** interface for the object. |
| *pFrame* | Pointer to an **IOleInPlaceFrame** interface for the object. Menus and toolbars must use this parameter. |
| *pDoc* | Pointer to an **IOleInPlaceUIWindow** interface for the object. Toolbars must use this parameter. |

# <a name="TranslateAccelerator"></a>TranslateAccelerator Event

Called by MSHTML when **IOleInPlaceActiveObject.TranslateAccelerator** or **IOleControlSite.TranslateAccelerator** is called.

```
FUNCTION TranslateAccelerator (BYVAL pWebCtx AS CWebCtx PTR, BYVAL lpMsg AS LPMSG, _
   BYVAL pguidCmdGroup AS const GUID PTR, BYVAL nCmdID AS DWORD) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *lpMsg* | Pointer to a **MSG** structure that specifies the message to be translated. |
| *pguidCmdGroup* | Pointer to a **GUID** for the command group identifier. |
| *nCmdID* | DWORD that specifies a command identifier. |

#### Return value

Returns S_OK (0) if successful, or an error value otherwise.

#### Remarks

When you use accelerator keys such as TAB, you may need to override the default host behavior. The example shows how to do this.

#### Example

This example shows how to override the default host behavior that occurs when a user tabs out of the first or last element.

````
FUNCTION DocHostUI_TranslateAccelerator (BYVAL pWebCtx AS CWebCtx PTR, BYVAL lpMsg AS LPMSG, _
BYVAL pguidCmdGroup AS const GUID PTR, BYVAL nCmdID AS DWORD) AS HRESULT
    IF lpMsg->message = WM_KEYDOWN AND lpMsg->wParam = VK_TAB THEN RETURN S_FALSE
END FUNCTION
````

# <a name="TranslateUrl"></a>TranslateUrl Event

Called by MSHTML to give the host an opportunity to modify the URL to be loaded.

```
FUNCTION TranslateUrl (BYVAL pWebCtx AS CWebCtx PTR, BYVAL dwTranslate AS DWORD, _
   BYVAL pchURLIn AS LPWSTR, BYVAL ppchURLOut AS LPWSTR PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |
| *dwTranslate* | Reserved. Must be set to NULL. |
| *pchURLIn* | Pointer to an OLECHAR that specifies the current URL for navigation. |
| *ppchURLOut* | Put. Address of a pointer variable that receives an OLECHAR pointer containing the new URL. |

#### Return value

Returns S_OK (0) if successful, or an error value otherwise.

#### Remarks

The host allocates the buffer *ppchURLOut* using **CoTaskMemAlloc**.

If the implementation of this method does not supply a URL, ppchURLOut should be set to NULL, even if the method fails or returns S_FALSE.

# <a name="UpdateUI"></a>UpdateUI Event

Called by MSHTML to notify the host that the command state has changed.

```
FUNCTION UpdateUI (BYVAL pWebCtx AS CWebCtx PTR) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWebCtx* | Pointer to the **CWebCtx** class. |

#### Return value

Returns S_OK (0) if successful, or an error value otherwise.

#### Remarks

The host should update the state of toolbar buttons in an implementation of this method.

This method is called regardless of the return value from the **ShowUI** method.
